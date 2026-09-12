//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"syscall"
	"unsafe"
)

type uiaObject struct{ dispatch }

func (d uiaObject) method(index int) uintptr {
	return (*[64]uintptr)(unsafe.Pointer(*(*uintptr)(unsafe.Pointer(d.ptr))))[index]
}

// UIA evaluates a name/role condition in its provider instead of making remote
// calls for every property of every control. No search results are cached.
func uiaNamedInvoke(pid uint32, op Operation) (any, error, bool) {
	if op.Op != "ui.invoke" {
		return nil, nil, false
	}
	scope, _ := op.Named["scope"].(string)
	// UserForm selectors refer to the MSAA client area. UIA also exposes the
	// title-bar Close button, so it is not the same search scope.
	if scope != "ribbon" && scope != "menu" {
		return nil, nil, false
	}
	name, _ := op.Named["name"].(string)
	if name == "" {
		return nil, nil, false
	}
	cls := syscall.GUID{Data1: 0xff48dba4, Data2: 0x60ef, Data3: 0x4201, Data4: [8]byte{0xaa, 0x87, 0x54, 0x10, 0x3e, 0xef, 0x59, 0x4e}}
	iid := syscall.GUID{Data1: 0x30cbe57d, Data2: 0xd9d0, Data3: 0x452a, Data4: [8]byte{0xab, 0x13, 0x7a, 0xc5, 0xac, 0x48, 0x25, 0xee}}
	var automation uiaObject
	hr, _, _ := ole32.NewProc("CoCreateInstance").Call(uintptr(unsafe.Pointer(&cls)), 0, 1, uintptr(unsafe.Pointer(&iid)), uintptr(unsafe.Pointer(&automation.ptr)))
	if failed(hr) || automation.ptr == 0 {
		return nil, nil, false
	}
	defer automation.release()
	condition := func(property int, value any) (uiaObject, error) {
		v, e := makeVariant(value, nil)
		if e != nil {
			return uiaObject{}, e
		}
		defer v.clear()
		var out uiaObject
		hr, _, _ := syscall.SyscallN(automation.method(23), automation.ptr, uintptr(property), uintptr(unsafe.Pointer(&v)), uintptr(unsafe.Pointer(&out.ptr)))
		if failed(hr) || out.ptr == 0 {
			return uiaObject{}, fmt.Errorf("UIA property condition unavailable")
		}
		return out, nil
	}
	filter, e := condition(30005, name)
	if e != nil {
		return nil, nil, false
	}
	if role, ok := op.Named["role"]; ok {
		roleFilter, e := condition(30095, role)
		if e != nil {
			filter.release()
			return nil, nil, false
		}
		var combined uiaObject
		hr, _, _ = syscall.SyscallN(automation.method(25), automation.ptr, filter.ptr, roleFilter.ptr, uintptr(unsafe.Pointer(&combined.ptr)))
		filter.release()
		roleFilter.release()
		if failed(hr) || combined.ptr == 0 {
			return nil, nil, false
		}
		filter = combined
	}
	defer filter.release()
	roots := []uint64{}
	for _, item := range windowInventory(pid) {
		w := item.(map[string]any)
		if w["visible_on_private_desktop"] != true || (op.HWND != 0 && op.HWND != w["hwnd"]) {
			continue
		}
		if title, ok := op.Named["window"].(string); ok && title != w["title"] {
			continue
		}
		switch scope {
		case "ribbon":
			if w["class"] != "OpusApp" {
				continue
			}
			if children, ok := w["children"].([]any); ok {
				for _, child := range children {
					c := child.(map[string]any)
					if c["class"] == "NetUIHWND" {
						roots = append(roots, c["hwnd"].(uint64))
					}
				}
			}
		case "menu":
			if w["class"] == "NetUIHWND" || w["class"] == "MsoCommandBarPopup" || w["class"] == "Net UI Tool Window" {
				roots = append(roots, w["hwnd"].(uint64))
			}
		}
	}
	var selected uiaObject
	defer func() {
		if selected.ptr != 0 {
			selected.release()
		}
	}()
	var selectedWindow uint64
	for _, hwnd := range roots {
		var root uiaObject
		hr, _, _ = syscall.SyscallN(automation.method(6), automation.ptr, uintptr(hwnd), uintptr(unsafe.Pointer(&root.ptr)))
		if failed(hr) || root.ptr == 0 {
			return nil, nil, false
		}
		var elements uiaObject
		hr, _, _ = syscall.SyscallN(root.method(6), root.ptr, 7, filter.ptr, uintptr(unsafe.Pointer(&elements.ptr)))
		root.release()
		if failed(hr) || elements.ptr == 0 {
			return nil, nil, false
		}
		var count int32
		hr, _, _ = syscall.SyscallN(elements.method(3), elements.ptr, uintptr(unsafe.Pointer(&count)))
		if failed(hr) || count < 0 {
			elements.release()
			return nil, nil, false
		}
		if count > 1 || (count == 1 && selected.ptr != 0) {
			elements.release()
			// UIA can expose extra presentation peers. Let the existing client
			// selector establish uniqueness rather than arbitrarily choosing one.
			return nil, nil, false
		}
		if count == 1 {
			hr, _, _ = syscall.SyscallN(elements.method(4), elements.ptr, 0, uintptr(unsafe.Pointer(&selected.ptr)))
			selectedWindow = hwnd
		}
		elements.release()
		if failed(hr) {
			return nil, nil, false
		}
	}
	if selected.ptr == 0 {
		return nil, nil, false
	}
	var elementPID int32
	hr, _, _ = syscall.SyscallN(selected.method(20), selected.ptr, uintptr(unsafe.Pointer(&elementPID)))
	if failed(hr) || uint32(elementPID) != pid {
		return nil, nil, false
	}
	var pattern uiaObject
	hr, _, _ = syscall.SyscallN(selected.method(16), selected.ptr, 10018, uintptr(unsafe.Pointer(&pattern.ptr)))
	if failed(hr) || pattern.ptr == 0 {
		return nil, nil, false
	}
	defer pattern.release()
	var actual uintptr
	hr, _, _ = syscall.SyscallN(pattern.method(7), pattern.ptr, uintptr(unsafe.Pointer(&actual)))
	text, e := bstrText(actual)
	if actual != 0 {
		sysFree.Call(actual)
	}
	if failed(hr) || e != nil || text != name {
		return nil, Fail("stale_ui_selector", "UIA control changed before invocation", nil), true
	}
	// Once an action is attempted, never fall back and risk invoking it twice.
	hr, _, _ = syscall.SyscallN(pattern.method(4), pattern.ptr)
	if failed(hr) {
		return nil, Fail("ui_action_failed", fmt.Sprintf("UIA default action failed: 0x%08x", uint32(hr)), nil), true
	}
	return map[string]any{"executed": true, "backend": "native UIAutomation", "hwnd": selectedWindow}, nil, true
}
