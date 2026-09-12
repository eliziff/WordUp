//go:build darwin && (amd64 || arm64)

package native

import (
	"encoding/base64"
	"encoding/binary"
	"fmt"
	"math"
	"runtime"
	"unicode/utf16"
	"unsafe"

	a "wordwright.local/internal/darwinapi"
)

// Carbon AEDesc is 16 bytes on 64-bit macOS. Apple owns the descriptor storage;
// it is disposed through AEDisposeDesc, never through the Go allocator.
type aeDesc struct {
	Type uint32
	_    uint32
	Data uintptr
}

func code(s string) uint32 {
	if len(s) != 4 {
		return 0
	}
	return binary.BigEndian.Uint32([]byte(s))
}
func codeText(n uint32) string {
	b := make([]byte, 4)
	binary.BigEndian.PutUint32(b, n)
	return string(b)
}
func status(name string, r uintptr) error {
	if int32(r) == 0 {
		return nil
	}
	return Fail("apple_event_error", fmt.Sprintf("%s returned OSStatus %d", name, int32(r)), map[string]any{"osstatus": int32(r), "automation_permission_may_be_required": int32(r) == -1743})
}
func (d *aeDesc) close() {
	if d.Data != 0 || d.Type != 0 {
		a.Call(a.AEDisposeDesc(), a.Pointer(d))
		*d = aeDesc{}
	}
}
func newDesc(kind string, b []byte) (aeDesc, error) {
	var d aeDesc
	var p uintptr
	if len(b) > 0 {
		p = uintptr(unsafe.Pointer(&b[0]))
	}
	r := a.Call(a.AECreateDesc(), uintptr(code(kind)), p, uintptr(len(b)), a.Pointer(&d))
	runtime.KeepAlive(b)
	return d, status("AECreateDesc", r)
}
func (d *aeDesc) clone() (aeDesc, error) {
	var v aeDesc
	r := a.Call(a.AEDuplicateDesc(), a.Pointer(d), a.Pointer(&v))
	return v, status("AEDuplicateDesc", r)
}
func (d *aeDesc) bytes() ([]byte, error) {
	n := int64(int32(a.Call(a.AEGetDescDataSize(), a.Pointer(d))))
	if n < 0 || n > 64<<20 {
		return nil, fmt.Errorf("Apple-event descriptor size budget")
	}
	if n == 0 {
		return []byte{}, nil
	}
	b := make([]byte, n)
	r := a.Call(a.AEGetDescData(), a.Pointer(d), uintptr(unsafe.Pointer(&b[0])), uintptr(n))
	runtime.KeepAlive(b)
	return b, status("AEGetDescData", r)
}
func descValue(v any, objects map[string]aeDesc) (aeDesc, error) {
	switch x := v.(type) {
	case nil:
		return newDesc("null", nil)
	case string:
		return newDesc("utf8", []byte(x))
	case bool:
		b := byte(0)
		if x {
			b = 1
		}
		return newDesc("bool", []byte{b})
	case float64:
		b := make([]byte, 8)
		if x == math.Trunc(x) && x >= math.MinInt32 && x <= math.MaxInt32 {
			binary.LittleEndian.PutUint32(b, uint32(int32(x)))
			return newDesc("long", b[:4])
		}
		binary.LittleEndian.PutUint64(b, math.Float64bits(x))
		return newDesc("doub", b)
	case int:
		return descValue(float64(x), objects)
	case []any:
		var list aeDesc
		r := a.Call(a.AECreateList(), 0, 0, 0, a.Pointer(&list))
		if e := status("AECreateList", r); e != nil {
			return list, e
		}
		for _, item := range x {
			d, e := descValue(item, objects)
			if e != nil {
				list.close()
				return aeDesc{}, e
			}
			r = a.Call(a.AEPutDesc(), a.Pointer(&list), 0, a.Pointer(&d))
			d.close()
			if e = status("AEPutDesc", r); e != nil {
				list.close()
				return aeDesc{}, e
			}
		}
		return list, nil
	case map[string]any:
		if id, ok := x["object"].(string); ok {
			d, ok := objects[id]
			if !ok {
				return aeDesc{}, fmt.Errorf("unknown Apple-event object %q", id)
			}
			return d.clone()
		}
		if t, ok := x["ae_type"].(string); ok {
			if len(t) != 4 {
				return aeDesc{}, fmt.Errorf("ae_type is a four-character code")
			}
			if text, ok := x["base64"].(string); ok {
				b, e := base64.StdEncoding.DecodeString(text)
				if e != nil {
					return aeDesc{}, e
				}
				return newDesc(t, b)
			}
			if value, ok := x["code"].(string); ok && len(value) == 4 {
				b := make([]byte, 4)
				binary.LittleEndian.PutUint32(b, code(value))
				return newDesc(t, b)
			}
			return aeDesc{}, fmt.Errorf("raw descriptor requires base64 bytes or a four-character code")
		}
		var record aeDesc
		r := a.Call(a.AECreateList(), 0, 0, 1, a.Pointer(&record))
		if e := status("AECreateList(record)", r); e != nil {
			return record, e
		}
		for k, value := range x {
			if len(k) != 4 {
				record.close()
				return aeDesc{}, fmt.Errorf("raw Apple-event record keys are four-character codes")
			}
			d, e := descValue(value, objects)
			if e != nil {
				record.close()
				return aeDesc{}, e
			}
			r = a.Call(a.AEPutParamDesc(), a.Pointer(&record), uintptr(code(k)), a.Pointer(&d))
			d.close()
			if e = status("AEPutParamDesc", r); e != nil {
				record.close()
				return aeDesc{}, e
			}
		}
		return record, nil
	default:
		return aeDesc{}, fmt.Errorf("unsupported Apple-event value %T", v)
	}
}
func (d *aeDesc) value(depth int) (any, error) {
	if depth > 24 {
		return nil, fmt.Errorf("Apple-event nesting budget")
	}
	if d.Type == code("list") || d.Type == code("reco") {
		var count int64
		r := a.Call(a.AECountItems(), a.Pointer(d), a.Pointer(&count))
		if e := status("AECountItems", r); e != nil {
			return nil, e
		}
		if count < 0 || count > 100000 {
			return nil, fmt.Errorf("Apple-event collection budget")
		}
		list := []any{}
		record := map[string]any{}
		for i := int64(1); i <= count; i++ {
			var child aeDesc
			var key uint32
			r = a.Call(a.AEGetNthDesc(), a.Pointer(d), uintptr(i), uintptr(code("****")), a.Pointer(&key), a.Pointer(&child))
			if e := status("AEGetNthDesc", r); e != nil {
				return nil, e
			}
			v, e := child.value(depth + 1)
			child.close()
			if e != nil {
				return nil, e
			}
			list = append(list, v)
			record[codeText(key)] = v
		}
		if d.Type == code("reco") {
			return record, nil
		}
		return list, nil
	}
	b, e := d.bytes()
	if e != nil {
		return nil, e
	}
	switch codeText(d.Type) {
	case "null", "msng":
		return nil, nil
	case "utf8", "TEXT":
		return string(b), nil
	case "utxt":
		if len(b)%2 != 0 {
			return nil, fmt.Errorf("odd Unicode Apple-event bytes")
		}
		r := make([]uint16, len(b)/2)
		for i := range r {
			r[i] = binary.LittleEndian.Uint16(b[2*i:])
		}
		return string(utf16.Decode(r)), nil
	case "bool":
		return len(b) > 0 && b[0] != 0, nil
	case "true":
		return true, nil
	case "fals":
		return false, nil
	case "long":
		if len(b) == 4 {
			return int32(binary.LittleEndian.Uint32(b)), nil
		}
	case "shor":
		if len(b) == 2 {
			return int16(binary.LittleEndian.Uint16(b)), nil
		}
	case "comp":
		if len(b) == 8 {
			return map[string]any{"int64": fmt.Sprint(int64(binary.LittleEndian.Uint64(b)))}, nil
		}
	case "doub":
		if len(b) == 8 {
			return math.Float64frombits(binary.LittleEndian.Uint64(b)), nil
		}
	case "type", "enum":
		if len(b) == 4 {
			return map[string]any{"ae_type": codeText(d.Type), "code": codeText(binary.LittleEndian.Uint32(b))}, nil
		}
	}
	return map[string]any{"ae_type": codeText(d.Type), "base64": base64.StdEncoding.EncodeToString(b)}, nil
}
func propertySpecifier(container *aeDesc, property string) (aeDesc, error) {
	key, e := descValue(map[string]any{"ae_type": "type", "code": property}, nil)
	if e != nil {
		return aeDesc{}, e
	}
	defer key.close()
	var spec aeDesc
	r := a.Call(a.CreateObjSpecifier(), uintptr(code("prop")), a.Pointer(container), uintptr(code("prop")), a.Pointer(&key), 0, a.Pointer(&spec))
	return spec, status("CreateObjSpecifier", r)
}
func sendAppleEvent(pid int, eventClass, eventID string, params map[string]aeDesc, timeoutMS int) (aeDesc, error) {
	b := make([]byte, 4)
	binary.LittleEndian.PutUint32(b, uint32(pid))
	target, e := newDesc("kpid", b)
	if e != nil {
		return aeDesc{}, e
	}
	defer target.close()
	var event aeDesc
	r := a.Call(a.AECreateAppleEvent(), uintptr(code(eventClass)), uintptr(code(eventID)), a.Pointer(&target), 0xffff, 0, a.Pointer(&event))
	if e = status("AECreateAppleEvent", r); e != nil {
		return aeDesc{}, e
	}
	defer event.close()
	for key, d := range params {
		if len(key) != 4 {
			return aeDesc{}, fmt.Errorf("event parameter code must be four characters")
		}
		r = a.Call(a.AEPutParamDesc(), a.Pointer(&event), uintptr(code(key)), a.Pointer(&d))
		if e = status("AEPutParamDesc", r); e != nil {
			return aeDesc{}, e
		}
	}
	ticks := int64(timeoutMS) * 60 / 1000
	if ticks < 1 {
		ticks = 7200
	}
	var reply aeDesc
	// Wait for the real application, without bringing its windows to the front.
	// The OS may present its own Automation consent; this is never bypassed.
	r = a.Call(a.AESendMessage(), a.Pointer(&event), a.Pointer(&reply), 3|0x10, uintptr(ticks))
	defer reply.close()
	if e = status("AESendMessage", r); e != nil {
		return aeDesc{}, e
	}
	var faultDesc aeDesc
	if int32(a.Call(a.AEGetParamDesc(), a.Pointer(&reply), uintptr(code("errn")), uintptr(code("long")), a.Pointer(&faultDesc))) == 0 {
		n, _ := faultDesc.value(0)
		faultDesc.close()
		if n != int32(0) && n != nil {
			var text aeDesc
			_ = a.Call(a.AEGetParamDesc(), a.Pointer(&reply), uintptr(code("errs")), uintptr(code("utf8")), a.Pointer(&text))
			s, _ := text.value(0)
			text.close()
			return aeDesc{}, Fail("word_apple_event_error", fmt.Sprintf("Word returned %v: %v", n, s), nil)
		}
	}
	var result aeDesc
	r = a.Call(a.AEGetParamDesc(), a.Pointer(&reply), uintptr(code("----")), uintptr(code("****")), a.Pointer(&result))
	if int32(r) == -1701 {
		return newDesc("null", nil)
	}
	return result, status("AEGetParamDesc(result)", r)
}
