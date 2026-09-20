//go:build windows && (amd64 || arm64)

package native

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
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
	return scopedWindowInventory(pid, "")
}

func scopedWindowInventory(pid uint32, windowClass string) []any {
	out := []any{}
	for _, hwnd := range windowsFor(pid) {
		class := classOf(hwnd)
		if windowClass != "" && class != windowClass {
			continue
		}
		v, _, _ := windowVisible.Call(hwnd)
		node := map[string]any{"hwnd": uint64(hwnd), "pid": pid, "class": class, "title": textOf(hwnd), "visible_on_private_desktop": v != 0}
		if class == "OpusApp" || class == "ThunderDFrame" || class == "ThunderXFrame" || class == "#32770" {
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

// Standard dialog controls remain usable when MSAA exposes an incomplete tree.
func nativeButtonName(hwnd uintptr) (string, bool) {
	if classOf(hwnd) != "Button" {
		return "", false
	}
	style, _, _ := user32.NewProc("GetWindowLongW").Call(hwnd, ^uintptr(15)) // GWL_STYLE
	switch style & 15 {
	case 0, 1, 14, 15:
	default:
		return "", false
	}
	// A doubled ampersand is literal; a single one introduces a mnemonic.
	caption := textOf(hwnd)
	var name strings.Builder
	for i := 0; i < len(caption); i++ {
		if caption[i] == '&' {
			if i+1 >= len(caption) || caption[i+1] != '&' {
				continue
			}
			i++
		}
		name.WriteByte(caption[i])
	}
	return name.String(), true
}

func nativeDialogSummary(window map[string]any) map[string]any {
	messages, buttons := []string{}, []string{}
	controls := []any{}
	children, _ := window["children"].([]any)
	for _, item := range children {
		child := item.(map[string]any)
		if child["class"] == "Static" && child["title"] != "" {
			messages = append(messages, child["title"].(string))
		}
		hwnd, _ := child["hwnd"].(uint64)
		if name, ok := nativeButtonName(uintptr(hwnd)); ok {
			buttons = append(buttons, name)
			controls = append(controls, map[string]any{"name": name, "role": 43, "selector": Operation{HWND: hwnd, Named: map[string]any{"expected_name": name}}})
		}
	}
	return map[string]any{"messages": messages, "buttons": buttons, "controls": controls, "hwnd": window["hwnd"], "title": window["title"]}
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
		} else if e != nil {
			if out["property_errors"] == nil {
				out["property_errors"] = map[string]string{}
			}
			out["property_errors"].(map[string]string)[p.key] = e.Error()
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
func findNamed(a dispatch, child int, path []any, hwnd uintptr, name string, role any, ancestor string, depth int, budget *int, matches *[]map[string]any) {
	if *budget <= 0 || len(*matches) > 1 {
		return
	}
	*budget--
	actual, _ := accProperty(a, "accName", child)
	inside := ancestor == ""
	if actual == ancestor {
		ancestor = ""
	}
	if inside && actual == name {
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
			findNamed(a, int(int32(v.Value)), path, hwnd, name, role, ancestor, 0, budget, matches)
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
			findNamed(c, 0, p, hwnd, name, role, ancestor, depth-1, budget, matches)
			c.release()
		}
		if *budget <= 0 || len(*matches) > 1 {
			return
		}
	}
}

func uiOperation(pid uint32, directory string, execute bool, op Operation) (any, error) {
	searchStarted := time.Now()
	var searchMS float64
	if op.Op == "ui.patterns" {
		if op.HWND == 0 || pidOf(uintptr(op.HWND)) != pid {
			return nil, Fail("window_not_owned", "Supply an owned window handle", nil)
		}
		action := op.Member
		if action == "" {
			action = "inspect"
		}
		if action != "inspect" && !execute {
			return nil, Fail("execution_not_authorized", "Changing a control pattern requires execute authority", nil)
		}
		name, _ := op.Named["name"].(string)
		controlType, _ := op.Named["control_type"].(string)
		selectorStarted := time.Now()
		rootHandles := []string{}
		if op.Named["scope"] == "ribbon" && name != "" && op.Named["role"] == nil {
			// Find window roots cheaply; let UIA establish control uniqueness once.
			q := windowEnum{pid: pid}
			enumChildren.Call(uintptr(op.HWND), collectWindowsCallback, uintptr(unsafe.Pointer(&q)))
			roots := []uint64{}
			for _, child := range q.windows {
				if classOf(child) == "NetUIHWND" {
					roots = append(roots, uint64(child))
				}
			}
			for _, root := range ribbonWindowRoots(roots) {
				rootHandles = append(rootHandles, fmt.Sprint(root))
			}
			if len(rootHandles) == 0 {
				return nil, Fail("ui_selector_not_found", "No owned Ribbon hosts found", map[string]any{"selector": op.Named, "hwnd": op.HWND})
			}
		} else if op.Named["scope"] == "ribbon" && name != "" {
			found, err := uiOperation(pid, directory, execute, Operation{Op: "ui.find", HWND: op.HWND, Named: map[string]any{"name": name, "scope": "ribbon", "role": op.Named["role"]}})
			if err != nil {
				return nil, err
			}
			selected := found.(map[string]any)
			b, err := json.Marshal(selected["selector"])
			if err != nil {
				return nil, err
			}
			var locator Operation
			if err = json.Unmarshal(b, &locator); err != nil {
				return nil, err
			}
			op.HWND = locator.HWND
		}
		selectorMS := float64(time.Since(selectorStarted).Microseconds()) / 1000
		if len(rootHandles) == 0 {
			rootHandles = append(rootHandles, fmt.Sprint(op.HWND))
		}
		result, err := OfficeTools(context.Background(), "patterns", fmt.Sprint(pid), strings.Join(rootHandles, ","), name, action, controlType)
		if details, ok := result.(map[string]any); ok {
			details["selector_ms"] = selectorMS
			details["pattern_root_hwnd"] = op.HWND
			details["pattern_root_handles"] = rootHandles
		}
		return result, err
	}
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
		attempts := 0
		for {
			attempts++
			r, e := uiOperation(pid, directory, execute, op)
			f, ok := e.(*Fault)
			if e == nil || !ok || f.Code != "ui_selector_not_found" || time.Now().After(deadline) {
				if result, ok := r.(map[string]any); ok {
					result["attempts"] = attempts
					result["wait_and_action_ms"] = float64(time.Since(searchStarted).Microseconds()) / 1000
				}
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
		dialogs := []any{}
		title, _ := op.Named["window"].(string)
		windowClass, _ := op.Named["window_class"].(string)
		matched := false
		inventoryClass := ""
		if windowClass == "#32770" {
			inventoryClass = windowClass
		}
		for _, item := range scopedWindowInventory(pid, inventoryClass) {
			w := item.(map[string]any)
			if w["visible_on_private_desktop"] != true && w["class"] != "#32770" {
				continue
			}
			hwnd := w["hwnd"].(uint64)
			selected := (op.HWND == 0 || op.HWND == hwnd) && (title == "" || title == w["title"]) && (windowClass == "" || windowClass == w["class"])
			matched = matched || selected
			// Targeted visual inspection avoids traversing every open document.
			// Retain owned modal dialogs even when another window was selected.
			if !selected && w["class"] != "#32770" {
				continue
			}
			var summary map[string]any
			if w["class"] == "#32770" {
				summary = nativeDialogSummary(w)
			}
			needTree := op.Named["trees"] != false || (summary != nil && (len(summary["messages"].([]string)) == 0 || len(summary["buttons"].([]string)) == 0))
			if needTree {
				tree, err := uiOperation(pid, directory, execute, Operation{Op: "ui.tree", HWND: hwnd, Depth: 4})
				if err != nil {
					w["tree_error"] = err.Error()
				}
				if op.Named["trees"] != false {
					w["tree"] = tree
				}
				if summary != nil {
					accessible := dialogSummary(tree)
					if len(summary["messages"].([]string)) == 0 {
						summary["messages"] = accessible["messages"]
					}
					if len(summary["buttons"].([]string)) == 0 {
						summary["buttons"], summary["controls"] = accessible["buttons"], accessible["controls"]
					}
				}
			}
			if summary != nil {
				dialogs = append(dialogs, summary)
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
		if !matched && (op.HWND != 0 || title != "" || windowClass != "") {
			return nil, Fail("diagnostic_window_not_found", "Requested visible owned window was not found", map[string]any{"hwnd": op.HWND, "window": title, "window_class": windowClass, "dialogs": dialogs})
		}
		return map[string]any{"pid": pid, "windows": out, "dialogs": dialogs, "session_logs": sessionLogs(directory)}, nil
	}
	if name, ok := op.Named["name"].(string); ok && name != "" {
		var messagePattern *regexp.Regexp
		if pattern, exists := op.Named["message_pattern"]; exists {
			text, ok := pattern.(string)
			if !ok || text == "" || op.Named["scope"] != "dialog" {
				return nil, fmt.Errorf("message_pattern requires a nonempty regular expression and scope=dialog")
			}
			var err error
			messagePattern, err = regexp.Compile(text)
			if err != nil {
				return nil, fmt.Errorf("invalid dialog message_pattern: %w", err)
			}
		}
		ancestor, _ := op.Named["ancestor"].(string)
		if execute && ancestor == "" {
			if result, err, handled := uiaNamedInvoke(pid, op); handled {
				return result, err
			}
		}
		matches := []map[string]any{}
		unmatchedDialogs := []map[string]any{}
		searches := []map[string]any{}
		budget := 4000
		search := func(hwnd uint64) {
			started, before := time.Now(), budget
			searched := map[string]any{"hwnd": hwnd}
			defer func() {
				searched["nodes_visited"] = before - budget
				searched["duration_ms"] = float64(time.Since(started).Microseconds()) / 1000
				searches = append(searches, searched)
			}()
			a, e := accessibleRoot(uintptr(hwnd))
			if e != nil {
				searched["error"] = e.Error()
				return
			}
			defer a.release()
			findNamed(a, 0, nil, uintptr(hwnd), name, op.Named["role"], ancestor, 16, &budget, &matches)
		}
		scope, _ := op.Named["scope"].(string)
		for _, item := range windowInventory(pid) {
			w := item.(map[string]any)
			if (w["visible_on_private_desktop"] != true && op.HWND != w["hwnd"]) || (op.HWND != 0 && op.HWND != w["hwnd"]) {
				continue
			}
			if title, ok := op.Named["window"].(string); ok && title != w["title"] {
				continue
			}
			var matchedDialog map[string]any
			if scope == "dialog" {
				if w["class"] != "#32770" {
					continue
				}
				matchedDialog = nativeDialogSummary(w)
				if messagePattern != nil {
					messages := matchedDialog["messages"].([]string)
					if len(messages) == 0 {
						tree, err := uiOperation(pid, directory, execute, Operation{Op: "ui.tree", HWND: w["hwnd"].(uint64), Depth: 4})
						if err != nil {
							return nil, err
						}
						messages = dialogSummary(tree)["messages"].([]string)
						matchedDialog["messages"] = messages
					}
					if !messagePattern.MatchString(strings.Join(messages, "\n")) {
						unmatchedDialogs = append(unmatchedDialogs, matchedDialog)
						continue
					}
				}
				if ancestor == "" && (op.Named["role"] == nil || fmt.Sprint(op.Named["role"]) == "43") {
					before := len(matches)
					for _, item := range matchedDialog["controls"].([]any) {
						control := item.(map[string]any)
						if control["name"] == name {
							matches = append(matches, map[string]any{"name": control["name"], "role": control["role"], "selector": control["selector"], "dialog": matchedDialog})
						}
					}
					if len(matches) != before {
						continue
					}
				}
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
				roots := []uint64{}
				for _, child := range children {
					c := child.(map[string]any)
					if c["class"] != "NetUIHWND" {
						continue
					}
					roots = append(roots, c["hwnd"].(uint64))
				}
				if scope == "ribbon" {
					roots = ribbonWindowRoots(roots)
				}
				for _, root := range roots {
					search(root)
				}
			}
			if len(matches) == before && scope != "ribbon" {
				search(w["hwnd"].(uint64))
			}
			if matchedDialog != nil {
				for _, match := range matches[before:] {
					match["dialog"] = matchedDialog
				}
			}
		}
		if budget <= 0 {
			return nil, Fail("ui_search_budget", "Narrow the selector with target, window or scope; search did not establish uniqueness", map[string]any{"selector": op.Named, "hwnd": op.HWND, "root_searches": searches})
		}
		if len(matches) == 0 {
			details := map[string]any{"selector": op.Named, "hwnd": op.HWND, "matches": matches, "root_searches": searches}
			if len(unmatchedDialogs) > 0 {
				details["unmatched_dialogs"] = unmatchedDialogs
			}
			return nil, Fail("ui_selector_not_found", fmt.Sprintf("No control matched %q; inspect ui.tree or ui.diagnostics to check the name, role and scope", name), details)
		}
		if len(matches) != 1 {
			return nil, Fail("ui_selector_not_unique", fmt.Sprintf("Multiple controls matched %q; narrow the selector using window, scope, role or ancestor", name), matches)
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
		op.Named = map[string]any{"expected_name": name, "matched_dialog": matches[0]["dialog"]}
		searchMS = float64(time.Since(searchStarted).Microseconds()) / 1000
	}
	hwnd := uintptr(op.HWND)
	if hwnd == 0 || pidOf(hwnd) != pid {
		return nil, Fail("window_not_owned", "UI operations require an exact window handle belonging to the app-owned Word process", nil)
	}
	if op.Op == "ui.close" {
		if !execute {
			return nil, Fail("execution_not_authorized", "Closing an owned window requires execute authority", nil)
		}
		ok, _, err := user32.NewProc("PostMessageW").Call(hwnd, 0x10, 0, 0)
		if ok == 0 {
			return nil, winError("PostMessageW(WM_CLOSE)", err)
		}
		return map[string]any{"posted": true, "hwnd": op.HWND}, nil
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
	if op.Op == "ui.invoke" && len(op.Args) == 0 && op.Child == 0 {
		if name, ok := nativeButtonName(hwnd); ok {
			if expected, exists := op.Named["expected_name"]; exists && expected != name {
				return nil, Fail("stale_ui_selector", "Button name no longer matches the supplied precondition", map[string]any{"expected": expected, "actual": name})
			}
			if !execute {
				return nil, Fail("execution_not_authorized", "Native UI interaction requires explicit execute authorization", nil)
			}
			enabled, _, _ := user32.NewProc("IsWindowEnabled").Call(hwnd)
			if enabled == 0 {
				return nil, Fail("ui_control_disabled", "The selected button is disabled", nil)
			}
			posted, _, err := user32.NewProc("PostMessageW").Call(hwnd, 0x00F5, 0, 0) // BM_CLICK
			if posted == 0 {
				return nil, winError("PostMessageW(BM_CLICK)", err)
			}
			return map[string]any{"posted": true, "backend": "native Button", "hwnd": op.HWND, "dialog": op.Named["matched_dialog"]}, nil
		}
	}
	resolveStarted := time.Now()
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
	resolveMS := float64(time.Since(resolveStarted).Microseconds()) / 1000
	actionStarted := time.Now()
	switch op.Op {
	case "ui.click":
		return physicalClick(pid, hwnd, a, op.Child)
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
	result := map[string]any{"executed": true, "backend": "native IAccessible", "hwnd": op.HWND,
		"search_ms": searchMS, "resolve_ms": resolveMS, "action_ms": float64(time.Since(actionStarted).Microseconds()) / 1000}
	if dialog := op.Named["matched_dialog"]; dialog != nil {
		result["dialog"] = dialog
	}
	return result, nil
}
