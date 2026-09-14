//go:build windows && (amd64 || arm64)

package native

import (
	"crypto/sha256"
	"embed"
	"encoding/xml"
	"fmt"
	"runtime"
	"sync"
	"syscall"
	"unsafe"
)

// Ribbon validation is a pure function of the XML bytes, but constructing the
// MSXML schema cache is relatively expensive. Keep a small process-local cache
// so repeated source-only checks do not pay that cost again. The bound matters
// because an agent can inspect many unrelated workspaces in one resident
// process. Callers receive a copy so one response cannot mutate cached data.
const ribbonValidationCacheLimit = 64

var ribbonValidationCache = struct {
	sync.Mutex
	values map[[32]byte]map[string]any
	order  [][32]byte
}{values: map[[32]byte]map[string]any{}}

func copyRibbonValidation(value map[string]any) map[string]any {
	copy := make(map[string]any, len(value))
	for key, item := range value {
		copy[key] = item
	}
	return copy
}

func cachedRibbonValidation(data []byte) (map[string]any, bool) {
	key := sha256.Sum256(data)
	ribbonValidationCache.Lock()
	defer ribbonValidationCache.Unlock()
	value, ok := ribbonValidationCache.values[key]
	if !ok {
		return nil, false
	}
	return copyRibbonValidation(value), true
}

func rememberRibbonValidation(data []byte, value map[string]any) {
	key := sha256.Sum256(data)
	ribbonValidationCache.Lock()
	defer ribbonValidationCache.Unlock()
	if _, exists := ribbonValidationCache.values[key]; exists {
		ribbonValidationCache.values[key] = copyRibbonValidation(value)
		return
	}
	if len(ribbonValidationCache.order) >= ribbonValidationCacheLimit {
		oldest := ribbonValidationCache.order[0]
		delete(ribbonValidationCache.values, oldest)
		ribbonValidationCache.order = ribbonValidationCache.order[1:]
	}
	ribbonValidationCache.values[key] = copyRibbonValidation(value)
	ribbonValidationCache.order = append(ribbonValidationCache.order, key)
}

//go:embed schemas/*.xsd
var ribbonSchemas embed.FS

func automationObject(progID string) (dispatch, error) {
	var cls syscall.GUID
	p, e := utf(progID)
	if e != nil {
		return dispatch{}, e
	}
	hr, _, _ := ole32.NewProc("CLSIDFromProgID").Call(uintptr(unsafe.Pointer(p)), uintptr(unsafe.Pointer(&cls)))
	if failed(hr) {
		return dispatch{}, fmt.Errorf("CLSIDFromProgID %s: %#x", progID, hr)
	}
	var out dispatch
	hr, _, _ = ole32.NewProc("CoCreateInstance").Call(uintptr(unsafe.Pointer(&cls)), 0, 1, uintptr(unsafe.Pointer(&iidDispatch)), uintptr(unsafe.Pointer(&out.ptr)))
	if failed(hr) {
		return out, fmt.Errorf("CoCreateInstance %s: %#x", progID, hr)
	}
	return out, nil
}

// ValidateRibbon uses the Windows XML engine and bundled Office schemas; it
// neither launches Word nor downloads schemas or external document entities.
func ValidateRibbon(data []byte) (map[string]any, error) {
	if cached, ok := cachedRibbonValidation(data); ok {
		return cached, nil
	}
	var root struct{ XMLName xml.Name }
	if e := xml.Unmarshal(data, &root); e != nil {
		return nil, e
	}
	file := "customUI.xsd"
	switch root.XMLName.Space {
	case "http://schemas.microsoft.com/office/2006/01/customui":
	case "http://schemas.microsoft.com/office/2009/07/customui":
		file = "customui14.xsd"
	default:
		return nil, fmt.Errorf("unsupported RibbonX namespace %q", root.XMLName.Space)
	}
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	hr, _, _ := coInit.Call(0, 2)
	if failed(hr) {
		return nil, fmt.Errorf("initialize XML validation COM: %#x", hr)
	}
	defer coUninit.Call()
	cache, e := automationObject("Msxml2.XMLSchemaCache.6.0")
	if e != nil {
		return nil, e
	}
	defer cache.release()
	load := func(text string) (dispatch, error) {
		d, e := automationObject("Msxml2.DOMDocument.6.0")
		if e != nil {
			return d, e
		}
		for _, p := range []string{"async", "validateOnParse", "resolveExternals"} {
			if e = d.put(p, false); e != nil {
				d.release()
				return dispatch{}, e
			}
		}
		v, e := d.call("setProperty", "ProhibitDTD", true)
		v.clear()
		if e == nil {
			v, e = d.call("loadXML", text)
		}
		if e == nil {
			ok, er := v.value(0)
			e = er
			if ok != true && e == nil {
				e = fmt.Errorf("XML parsing failed")
			}
		}
		v.clear()
		if e != nil {
			d.release()
			return dispatch{}, e
		}
		return d, nil
	}
	schema, _ := ribbonSchemas.ReadFile("schemas/" + file)
	s, e := load(string(schema))
	if e != nil {
		return nil, e
	}
	defer s.release()
	v, e := cache.invoke("add", 1, []any{root.XMLName.Space, map[string]any{"object": "schema"}}, nil, map[string]dispatch{"schema": s})
	v.clear()
	if e != nil {
		return nil, e
	}
	d, e := load(string(data))
	if e != nil {
		return nil, e
	}
	defer d.release()
	v, e = d.invoke("schemas", 8, []any{map[string]any{"object": "cache"}}, nil, map[string]dispatch{"cache": cache})
	v.clear()
	if e != nil {
		return nil, e
	}
	v, e = d.call("validate")
	if e != nil {
		return nil, e
	}
	validation, e := v.object()
	v.clear()
	if e != nil {
		return nil, e
	}
	defer validation.release()
	result := map[string]any{"engine": "MSXML 6.0", "schema": file, "native_word_verified": false}
	for _, key := range []string{"errorCode", "reason", "line", "linepos", "srcText"} {
		v, e = validation.get(key)
		if e != nil {
			return nil, e
		}
		result[key], e = v.value(0)
		v.clear()
		if e != nil {
			return nil, e
		}
	}
	result["valid"] = fmt.Sprint(result["errorCode"]) == "0"
	rememberRibbonValidation(data, result)
	return result, nil
}
