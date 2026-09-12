//go:build windows && (amd64 || arm64)

package native

import (
	"embed"
	"encoding/xml"
	"fmt"
	"runtime"
	"syscall"
	"unsafe"
)

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
	return result, nil
}
