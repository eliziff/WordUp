//go:build windows && (amd64 || arm64)

package native

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"syscall"
	"time"
	"unsafe"
)

var iidAccessible = syscall.GUID{Data1: 0x618736e0, Data2: 0x3c3d, Data3: 0x11cf, Data4: [8]byte{0x81, 0x0c, 0, 0xaa, 0, 0x38, 0x9b, 0x71}}
var accessibleChildren = oleacc.NewProc("AccessibleChildren")

func (d dispatch) query(iid *syscall.GUID) (dispatch, error) {
	var p uintptr
	hr, _, _ := syscall.SyscallN(d.method(0), d.ptr, uintptr(unsafe.Pointer(iid)), uintptr(unsafe.Pointer(&p)))
	if failed(hr) || p == 0 {
		return dispatch{}, fmt.Errorf("QueryInterface failed: 0x%08X", uint32(hr))
	}
	return dispatch{p}, nil
}
func accessibleRoot(hwnd uintptr) (dispatch, error) {
	var p uintptr
	hr, _, _ := accessibleWindow.Call(hwnd, 0xfffffffc, uintptr(unsafe.Pointer(&iidAccessible)), uintptr(unsafe.Pointer(&p)))
	if failed(hr) || p == 0 {
		return dispatch{}, Fail("accessibility_unavailable", fmt.Sprintf("Owned window did not expose IAccessible (HRESULT 0x%08X)", uint32(hr)), nil)
	}
	return dispatch{p}, nil
}
func windowInventory(pid uint32) []any {
	out := []any{}
	for _, hwnd := range windowsFor(pid) {
		v, _, _ := windowVisible.Call(hwnd)
		node := map[string]any{"hwnd": uint64(hwnd), "pid": pid, "class": classOf(hwnd), "title": textOf(hwnd), "visible_on_private_desktop": v != 0}
		if classOf(hwnd) == "OpusApp" {
			q := windowEnum{pid: pid}
			enumChildren.Call(hwnd, collectWindowsCallback, uintptr(unsafe.Pointer(&q)))
			children := []any{}
			for _, child := range q.windows {
				children = append(children, map[string]any{"hwnd": uint64(child), "class": classOf(child), "title": textOf(child)})
			}
			node["children"] = children
		}
		out = append(out, node)
	}
	return out
}
func childrenOf(d dispatch) ([]variant, error) {
	v, e := d.invokeIDs("accChildCount", 2, nil, nil, nil, nil, []int32{-5001})
	if e != nil {
		return nil, e
	}
	n := int(int32(v.Value))
	v.clear()
	if n < 0 || n > 10000 {
		return nil, fmt.Errorf("accessibility child budget")
	}
	if n == 0 {
		return nil, nil
	}
	children := make([]variant, n)
	var got int32
	hr, _, _ := accessibleChildren.Call(d.ptr, 0, uintptr(n), uintptr(unsafe.Pointer(&children[0])), uintptr(unsafe.Pointer(&got)))
	if failed(hr) || got < 0 || int(got) > n {
		for i := range children {
			children[i].clear()
		}
		return nil, fmt.Errorf("AccessibleChildren failed: 0x%08X", uint32(hr))
	}
	return children[:got], nil
}
func resolveAccessible(hwnd uintptr, path []any) (dispatch, error) {
	a, e := accessibleRoot(hwnd)
	if e != nil {
		return dispatch{}, e
	}
	for _, entry := range path {
		n, ok := entry.(float64)
		if !ok || n < 0 || n > 10000 || float64(int(n)) != n {
			a.release()
			return dispatch{}, fmt.Errorf("accessibility path entries are zero-based enumeration indexes")
		}
		children, e := childrenOf(a)
		if e != nil {
			a.release()
			return dispatch{}, e
		}
		if int(n) >= len(children) {
			for i := range children {
				children[i].clear()
			}
			a.release()
			return dispatch{}, fmt.Errorf("accessibility path is stale")
		}
		chosen := &children[int(n)]
		var next dispatch
		if chosen.VT == 9 {
			d, e2 := chosen.object()
			if e2 == nil {
				next, e = d.query(&iidAccessible)
				d.release()
			} else {
				e = e2
			}
		} else {
			e = fmt.Errorf("path ends at a simple child; specify its child ID on the parent object")
		}
		for i := range children {
			children[i].clear()
		}
		a.release()
		if e != nil {
			return dispatch{}, e
		}
		a = next
	}
	return a, nil
}

var accessibilityIDs = map[string]int32{"accName": -5003, "accValue": -5004, "accDescription": -5005, "accRole": -5006, "accState": -5007, "accKeyboardShortcut": -5010, "accDefaultAction": -5013}

func accProperty(a dispatch, name string, child int) (any, error) {
	id, known := accessibilityIDs[name]
	var v variant
	var e error
	if known {
		v, e = a.invokeIDs(name, 2, []any{child}, nil, nil, nil, []int32{id})
	} else {
		v, e = a.get(name, child)
	}
	if e != nil {
		return nil, e
	}
	defer v.clear()
	return v.value(0)
}
func accNode(a dispatch, child int, path []any, hwnd uintptr, depth int, budget *int) map[string]any {
	*budget--
	out := map[string]any{"selector": map[string]any{"hwnd": uint64(hwnd), "args": path, "child": child}}
	for _, p := range []struct{ member, key string }{{"accName", "name"}, {"accValue", "value"}, {"accRole", "role"}, {"accState", "state"}, {"accDescription", "description"}, {"accDefaultAction", "default_action"}, {"accKeyboardShortcut", "shortcut"}} {
		v, e := accProperty(a, p.member, child)
		if e == nil && v != nil {
			out[p.key] = v
		}
	}
	if child != 0 || depth <= 0 || *budget <= 0 {
		return out
	}
	entries, e := childrenOf(a)
	if e != nil {
		out["children_error"] = e.Error()
		return out
	}
	defer func() {
		for i := range entries {
			entries[i].clear()
		}
	}()
	items := []any{}
	for i := range entries {
		if *budget <= 0 {
			out["truncated"] = true
			break
		}
		v := &entries[i]
		if v.VT == 3 {
			items = append(items, accNode(a, int(int32(v.Value)), path, hwnd, 0, budget))
		} else if v.VT == 9 {
			d, e := v.object()
			if e != nil {
				continue
			}
			c, e := d.query(&iidAccessible)
			d.release()
			if e != nil {
				continue
			}
			p := append(append([]any(nil), path...), float64(i))
			items = append(items, accNode(c, 0, p, hwnd, depth-1, budget))
			c.release()
		}
	}
	if len(items) > 0 {
		out["children"] = items
	}
	return out
}
func findNamed(a dispatch, child int, path []any, hwnd uintptr, name string, role any, depth int, budget *int, matches *[]map[string]any) {
	if *budget <= 0 || len(*matches) > 1 {
		return
	}
	*budget--
	actual, _ := accProperty(a, "accName", child)
	if actual == name {
		actualRole, _ := accProperty(a, "accRole", child)
		if role == nil || fmt.Sprint(actualRole) == fmt.Sprint(role) {
			n := 1
			*matches = append(*matches, accNode(a, child, path, hwnd, 0, &n))
		}
	}
	if child != 0 || depth == 0 {
		return
	}
	entries, e := childrenOf(a)
	if e != nil {
		return
	}
	defer func() {
		for i := range entries {
			entries[i].clear()
		}
	}()
	for i := range entries {
		v := &entries[i]
		if v.VT == 3 {
			findNamed(a, int(int32(v.Value)), path, hwnd, name, role, 0, budget, matches)
		} else if v.VT == 9 {
			d, e := v.object()
			if e != nil {
				continue
			}
			c, e := d.query(&iidAccessible)
			d.release()
			if e != nil {
				continue
			}
			p := append(append([]any(nil), path...), float64(i))
			findNamed(c, 0, p, hwnd, name, role, depth-1, budget, matches)
			c.release()
		}
		if *budget <= 0 || len(*matches) > 1 {
			return
		}
	}
}

func uiOperation(pid uint32, directory string, execute bool, op Operation) (any, error) {
	if wait, ok := op.Named["wait_ms"].(float64); ok {
		if wait < 0 || wait > 10000 {
			return nil, fmt.Errorf("UI wait_ms must be 0..10000")
		}
		named := map[string]any{}
		for k, v := range op.Named {
			if k != "wait_ms" {
				named[k] = v
			}
		}
		op.Named = named
		deadline := time.Now().Add(time.Duration(wait) * time.Millisecond)
		for {
			r, e := uiOperation(pid, directory, execute, op)
			f, ok := e.(*Fault)
			if e == nil || !ok || f.Code != "ui_selector_not_unique" || time.Now().After(deadline) {
				return r, e
			}
			if matches, ok := f.Details.([]map[string]any); !ok || len(matches) != 0 {
				return r, e
			}
			pump()
			time.Sleep(50 * time.Millisecond)
		}
	}
	if op.Op == "ui.windows" {
		return windowInventory(pid), nil
	}
	if op.Op == "ui.diagnostics" {
		out := []any{}
		for _, item := range windowInventory(pid) {
			w := item.(map[string]any)
			if w["visible_on_private_desktop"] != true {
				continue
			}
			hwnd := w["hwnd"].(uint64)
			if op.Named["trees"] != false {
				tree, err := uiOperation(pid, directory, execute, Operation{Op: "ui.tree", HWND: hwnd, Depth: 4})
				w["tree"] = tree
				if err != nil {
					w["tree_error"] = err.Error()
				}
			}
			if op.File != "" {
				if err := os.MkdirAll(op.File, 0700); err != nil {
					return nil, err
				}
				path := filepath.Join(op.File, fmt.Sprintf("window-%d.png", hwnd))
				if err := screenshot(uintptr(hwnd), path); err != nil {
					w["capture_error"] = err.Error()
				} else {
					w["screenshot"] = path
				}
			}
			out = append(out, w)
		}
		return map[string]any{"pid": pid, "windows": out}, nil
	}
	if name, ok := op.Named["name"].(string); ok && name != "" {
		if execute {
			if result, err, handled := uiaNamedInvoke(pid, op); handled {
				return result, err
			}
		}
		matches := []map[string]any{}
		budget := 4000
		search := func(hwnd uint64) {
			a, e := accessibleRoot(uintptr(hwnd))
			if e != nil {
				return
			}
			defer a.release()
			findNamed(a, 0, nil, uintptr(hwnd), name, op.Named["role"], 16, &budget, &matches)
		}
		scope, _ := op.Named["scope"].(string)
		for _, item := range windowInventory(pid) {
			w := item.(map[string]any)
			if w["visible_on_private_desktop"] != true || (op.HWND != 0 && op.HWND != w["hwnd"]) {
				continue
			}
			if title, ok := op.Named["window"].(string); ok && title != w["title"] {
				continue
			}
			if scope == "ribbon" && w["class"] != "OpusApp" {
				continue
			}
			if scope == "menu" && w["class"] != "NetUIHWND" && w["class"] != "MsoCommandBarPopup" && w["class"] != "Net UI Tool Window" {
				continue
			}
			if scope == "form" && w["class"] != "ThunderDFrame" && w["class"] != "ThunderXFrame" {
				continue
			}
			// Word exposes its Ribbon through child NetUI windows, separately
			// from the main document's client accessibility tree.
			before := len(matches)
			if children, ok := w["children"].([]any); ok {
				for _, child := range children {
					c := child.(map[string]any)
					if c["class"] != "NetUIHWND" {
						continue
					}
					search(c["hwnd"].(uint64))
				}
			}
			if len(matches) == before && scope != "ribbon" {
				search(w["hwnd"].(uint64))
			}
		}
		if budget <= 0 {
			return nil, Fail("ui_search_budget", "Narrow the selector with target, window or scope; search did not establish uniqueness", nil)
		}
		if len(matches) != 1 {
			return nil, Fail("ui_selector_not_unique", fmt.Sprintf("Expected one control named %q; found %d", name, len(matches)), matches)
		}
		if op.Op == "ui.find" {
			return matches[0], nil
		}
		b, _ := json.Marshal(matches[0]["selector"])
		var selector Operation
		if err := json.Unmarshal(b, &selector); err != nil {
			return nil, err
		}
		op.HWND, op.Args, op.Child = selector.HWND, selector.Args, selector.Child
		op.Named = map[string]any{"expected_name": name}
	}
	hwnd := uintptr(op.HWND)
	if hwnd == 0 || pidOf(hwnd) != pid {
		return nil, Fail("window_not_owned", "UI operations require an exact window handle belonging to the app-owned Word process", nil)
	}
	if op.Op == "ui.context_menu" {
		if !execute {
			return nil, Fail("execution_not_authorized", "Context-menu interaction requires execute authority", nil)
		}
		q := windowEnum{pid: pid}
		enumChildren.Call(hwnd, findDocumentCallback, uintptr(unsafe.Pointer(&q)))
		if q.found != 0 {
			hwnd = q.found
		}
		ok, _, err := user32.NewProc("PostMessageW").Call(hwnd, 0x7b, hwnd, ^uintptr(0))
		if ok == 0 {
			return nil, winError("PostMessageW(WM_CONTEXTMENU)", err)
		}
		return map[string]any{"posted": true, "hwnd": uint64(hwnd)}, nil
	}
	if op.Op == "ui.capture" {
		file := op.File
		if file == "" {
			file = filepath.Join(directory, "window.png")
		}
		if e := screenshot(hwnd, file); e != nil {
			return nil, e
		}
		return map[string]any{"file": file, "capture": "PrintWindow on the app-owned desktop", "hwnd": op.HWND}, nil
	}
	a, e := resolveAccessible(hwnd, op.Args)
	if e != nil {
		return nil, e
	}
	defer a.release()
	if expected, ok := op.Named["expected_name"].(string); ok {
		actual, e := accProperty(a, "accName", op.Child)
		if e != nil || actual != expected {
			return nil, Fail("stale_ui_selector", "Accessible name no longer matches the supplied precondition", map[string]any{"expected": expected, "actual": actual})
		}
	}
	if op.Op == "ui.tree" {
		depth := op.Depth
		if depth == 0 {
			depth = 4
		}
		if depth < 0 || depth > 16 {
			return nil, fmt.Errorf("accessibility depth must be 0..16")
		}
		budget := 4000
		return accNode(a, op.Child, op.Args, hwnd, depth, &budget), nil
	}
	if !execute {
		return nil, Fail("execution_not_authorized", "Native UI interaction requires explicit execute authorization", nil)
	}
	var v variant
	switch op.Op {
	case "ui.invoke":
		v, e = a.call("accDoDefaultAction", op.Child)
	case "ui.set_value":
		value, ok := op.Value.(string)
		if !ok {
			return nil, fmt.Errorf("ui.set_value requires a string")
		}
		v, e = a.invoke("accValue", 4, []any{op.Child, value}, nil, nil)
	case "ui.select":
		flags := float64(1)
		if n, ok := op.Value.(float64); ok {
			flags = n
		}
		v, e = a.call("accSelect", flags, op.Child)
	default:
		return nil, Fail("unknown_ui_operation", "Unsupported native UI action "+strings.TrimPrefix(op.Op, "ui."), nil)
	}
	v.clear()
	if e != nil {
		return nil, e
	}
	return map[string]any{"executed": true, "backend": "native IAccessible", "hwnd": op.HWND}, nil
}
