//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"syscall"
	"time"
	"unsafe"
)

type uiaObject struct{ dispatch }

func (d uiaObject) method(index int) uintptr {
	return (*[64]uintptr)(unsafe.Pointer(*(*uintptr)(unsafe.Pointer(d.ptr))))[index]
}

// UIA evaluates a name/role condition in its provider instead of making remote
// calls for every property of every control. No search results are cached.
func uiaNamedInvoke(pid uint32, op Operation) (any, error, bool) {
	started := time.Now()
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
	cls := syscall.GUID{Data1: 0xe22ad333, Data2: 0xb25f, Data3: 0x460c, Data4: [8]byte{0x83, 0xd0, 0x05, 0x81, 0x10, 0x73, 0x95, 0xc9}}
	iid := syscall.GUID{Data1: 0x34723aff, Data2: 0x0c9d, Data3: 0x49d0, Data4: [8]byte{0x98, 0x96, 0x7a, 0xb5, 0x2d, 0xf8, 0xcd, 0x8a}}
	var automation uiaObject
	hr, _, _ := ole32.NewProc("CoCreateInstance").Call(uintptr(unsafe.Pointer(&cls)), 0, 1, uintptr(unsafe.Pointer(&iid)), uintptr(unsafe.Pointer(&automation.ptr)))
	if failed(hr) || automation.ptr == 0 {
		return nil, nil, false
	}
	defer automation.release()
	// IUIAutomation2 uses native provider deadlines (Windows 8+).
	for _, method := range []int{61, 63} {
		hr, _, _ = syscall.SyscallN(automation.method(method), automation.ptr, 2000)
		if failed(hr) {
			return nil, nil, false
		}
	}
	providerMS := float64(time.Since(started).Microseconds()) / 1000
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
	windowCount := 0
	for _, item := range windowInventory(pid) {
		w := item.(map[string]any)
		if w["visible_on_private_desktop"] != true || (op.HWND != 0 && op.HWND != w["hwnd"]) {
			continue
		}
		if title, ok := op.Named["window"].(string); ok && title != w["title"] {
			continue
		}
		before := len(roots)
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
		if scope == "ribbon" {
			roots = append(roots[:before], ribbonWindowRoots(roots[before:])...)
		}
		if len(roots) > before {
			windowCount++
		}
	}
	var selected uiaObject
	rootsMS := float64(time.Since(started).Microseconds())/1000 - providerMS
	searchStarted := time.Now()
	defer func() {
		if selected.ptr != 0 {
			selected.release()
		}
	}()
	var selectedWindow uint64
	rootSearches := []map[string]any{}
	for _, hwnd := range roots {
		rootStarted := time.Now()
		var root uiaObject
		hr, _, _ = syscall.SyscallN(automation.method(6), automation.ptr, uintptr(hwnd), uintptr(unsafe.Pointer(&root.ptr)))
		if uint32(hr) == 0x80131505 {
			return nil, uiaProviderTimeout("ElementFromHandle", hwnd, rootSearches), true
		}
		if failed(hr) || root.ptr == 0 {
			return nil, nil, false
		}
		var elements uiaObject
		hr, _, _ = syscall.SyscallN(root.method(6), root.ptr, 7, filter.ptr, uintptr(unsafe.Pointer(&elements.ptr)))
		root.release()
		if uint32(hr) == 0x80131505 {
			return nil, uiaProviderTimeout("FindAll", hwnd, rootSearches), true
		}
		if failed(hr) || elements.ptr == 0 {
			return nil, nil, false
		}
		var count int32
		hr, _, _ = syscall.SyscallN(elements.method(3), elements.ptr, uintptr(unsafe.Pointer(&count)))
		if failed(hr) || count < 0 {
			elements.release()
			return nil, nil, false
		}
		rootSearches = append(rootSearches, map[string]any{"hwnd": hwnd, "matches": count, "duration_ms": float64(time.Since(rootStarted).Microseconds()) / 1000})
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
	searchMS := float64(time.Since(started).Microseconds()) / 1000
	actionStarted := time.Now()
	hr, _, _ = syscall.SyscallN(pattern.method(4), pattern.ptr)
	if uint32(hr) == 0x80131505 {
		return nil, uiaProviderTimeout("DoDefaultAction", selectedWindow, rootSearches), true
	}
	if failed(hr) {
		return nil, Fail("ui_action_failed", fmt.Sprintf("UIA default action failed: 0x%08x", uint32(hr)), nil), true
	}
	result := map[string]any{"executed": true, "backend": "native UIAutomation", "hwnd": selectedWindow, "search_ms": searchMS, "provider_start_ms": providerMS, "selector_setup_ms": rootsMS, "provider_search_ms": float64(actionStarted.Sub(searchStarted).Microseconds()) / 1000, "action_ms": float64(time.Since(actionStarted).Microseconds()) / 1000, "roots_searched": len(roots), "windows_searched": windowCount, "root_searches": rootSearches}
	if windowCount > 1 {
		result["selector_hint"] = "Search crossed multiple owned windows. Set target to the intended document handle to avoid searching unrelated windows."
	}
	return result, nil, true
}

// NetUI also hosts status bars and task panes. Prefer native Ribbon descendants
// when Office exposes that host; preserve the existing search on other layouts.
func ribbonWindowRoots(roots []uint64) []uint64 {
	selected := []uint64{}
	parent := user32.NewProc("GetParent")
	for _, root := range roots {
		for hwnd := uintptr(root); hwnd != 0; {
			if classOf(hwnd) == "MsoCommandBar" && textOf(hwnd) == "Ribbon" {
				selected = append(selected, root)
				break
			}
			hwnd, _, _ = parent.Call(hwnd)
		}
	}
	if len(selected) == 0 {
		return roots
	}
	return selected
}

func uiaProviderTimeout(operation string, hwnd uint64, searches []map[string]any) error {
	return Fail("ui_provider_timeout", "UI Automation provider did not respond within its native deadline", map[string]any{"operation": operation, "hwnd": hwnd, "provider_timeout_ms": 2000, "root_searches": searches, "action_attempted": operation == "DoDefaultAction"})
}
