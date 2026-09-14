//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"strings"
	"time"
	"unsafe"
)

const (
	wmKeyDown = 0x0100
	wmKeyUp   = 0x0101
	vkReturn  = 0x0D
	wmClose   = 0x0010
)

type immediateEditor struct {
	window uint64
	path   []any
	child  int
	value  string
}

// vbaImmediate evaluates against the currently paused VBA frame. It uses the
// VBE Immediate window only on the app-owned private desktop and returns the
// text Word exposes there; it never substitutes a scratch project for a
// paused frame.
func vbaImmediate(app dispatch, pid uint32, execute bool, op Operation) (any, error) {
	if !execute {
		return nil, Fail("execution_not_authorized", "Paused VBA inspection requires execute capability", nil)
	}
	state, err := inspectVBA(app, execute, false)
	if err != nil {
		return nil, err
	}
	stateMap := state.(map[string]any)
	if stateMap["project_mode"] != int32(1) {
		return nil, Fail("vba_not_paused", "Paused-frame Immediate evaluation requires a VBA project in break mode", state)
	}

	expression, _ := op.Value.(string)
	expression = strings.TrimSpace(expression)
	if len(expression) > 16384 || strings.ContainsAny(expression, "\r\n") {
		return nil, fmt.Errorf("Immediate expression must be one bounded line")
	}
	symbols := debugSymbols(op.Named["symbols"])
	if op.Op == "ui.vba.locals" && expression == "" && len(symbols) == 0 {
		return nil, fmt.Errorf("ui.vba.locals requires symbols or an expression")
	}
	if expression == "" && len(symbols) == 0 {
		return nil, fmt.Errorf("Immediate expression required")
	}

	hwnd, opened, err := openImmediateWindow(app, pid)
	if err != nil {
		return nil, err
	}
	result := map[string]any{
		"paused":        true,
		"source":        "native VBE Immediate Window",
		"selection":     state,
		"window":        hwnd,
		"window_opened": opened,
	}
	defer func() {
		if opened {
			if closeErr := closeOwnedWindow(hwnd); closeErr != nil {
				result["cleanup_error"] = closeErr.Error()
			}
		}
	}()

	editor, err := locateImmediateEditor(hwnd)
	if err != nil {
		return nil, err
	}
	result["editor"] = map[string]any{"path": editor.path, "child": editor.child}
	if expression != "" {
		row := evaluateImmediate(editor, expression)
		result["expression"] = row
	}
	if len(symbols) > 0 {
		rows := make([]map[string]any, 0, len(symbols))
		unavailable := []string{}
		for _, symbol := range symbols {
			row := evaluateImmediate(editor, symbol)
			row["symbol"] = symbol
			if row["available"] != true {
				unavailable = append(unavailable, symbol)
			}
			rows = append(rows, row)
		}
		result["symbols"] = rows
		if len(unavailable) > 0 {
			result["unavailable_symbols"] = unavailable
		}
	}
	return result, nil
}

func debugSymbols(raw any) []string {
	values, ok := raw.([]any)
	if !ok {
		if typed, typedOK := raw.([]string); typedOK {
			values = make([]any, len(typed))
			for i, value := range typed {
				values[i] = value
			}
		} else {
			return nil
		}
	}
	seen := map[string]bool{}
	result := []string{}
	for _, value := range values {
		text, ok := value.(string)
		text = strings.TrimSpace(text)
		if !ok || text == "" || len(text) > 512 || strings.ContainsAny(text, "\r\n") || seen[text] {
			continue
		}
		seen[text] = true
		result = append(result, text)
	}
	return result
}

func evaluateImmediate(editor immediateEditor, expression string) map[string]any {
	query := strings.TrimSpace(expression)
	if !strings.HasPrefix(query, "?") {
		query = "? " + query
	}
	row := map[string]any{"expression": expression, "query": query, "available": false}
	dispatch, err := editorDispatch(editor)
	if err != nil {
		row["error"] = err.Error()
		return row
	}
	before, err := accProperty(dispatch, "accValue", editor.child)
	dispatch.release()
	if err == nil {
		row["buffer_before"] = before
	}
	dispatch, err = editorDispatch(editor)
	if err != nil {
		row["error"] = err.Error()
		return row
	}
	value, err := dispatch.invoke("accValue", 4, []any{editor.child, query}, nil, nil)
	dispatch.release()
	value.clear()
	if err != nil {
		row["error"] = fault(err)
		return row
	}
	if err = postImmediateEnter(editor.window); err != nil {
		row["error"] = err.Error()
		return row
	}
	deadline := time.Now().Add(1200 * time.Millisecond)
	after := ""
	var readErr error
	for time.Now().Before(deadline) {
		pump()
		dispatch, dispatchErr := editorDispatch(editor)
		if dispatchErr != nil {
			readErr = dispatchErr
			break
		}
		var current any
		current, readErr = accProperty(dispatch, "accValue", editor.child)
		dispatch.release()
		if readErr == nil {
			after, _ = current.(string)
			if after != "" && after != fmt.Sprint(before) {
				break
			}
		}
		time.Sleep(25 * time.Millisecond)
	}
	if readErr != nil {
		row["error"] = readErr.Error()
		return row
	}
	row["buffer_after"] = after
	row["changed"] = after != fmt.Sprint(before)
	if after != "" && after != query {
		row["available"] = true
		row["result_text"] = after
	} else {
		row["error"] = "Immediate window did not expose a result buffer"
	}
	return row
}

// editorDispatch resolves the part-qualified accessibility path on each call.
// The VBE can recreate its child controls after Enter, so retaining a dispatch
// pointer across evaluations would be stale by construction.
func editorDispatch(editor immediateEditor) (dispatch, error) {
	return resolveAccessible(uintptr(editor.window), editor.path)
}

func openImmediateWindow(app dispatch, pid uint32) (uint64, bool, error) {
	if candidates := immediateWindows(pid); len(candidates) == 1 {
		return uint64(candidates[0]), false, nil
	} else if len(candidates) > 1 {
		return 0, false, Fail("vba_immediate_ambiguous", "More than one owned Immediate window is visible", windowDetails(candidates))
	}
	vbe, err := objectProperty(app, "VBE")
	if err != nil {
		return 0, false, err
	}
	defer vbe.release()
	bars, err := objectProperty(vbe, "CommandBars")
	if err != nil {
		return 0, false, err
	}
	defer bars.release()
	commandFound := false
	for _, id := range []int{2554, 2557, 2558, 2559} {
		value, findErr := bars.invoke("FindControl", 1, nil, map[string]any{"ID": id}, nil)
		if findErr != nil {
			continue
		}
		control, objectErr := value.object()
		value.clear()
		if objectErr != nil {
			continue
		}
		caption := ""
		if captionValue, captionErr := control.get("Caption"); captionErr == nil {
			if text, textErr := captionValue.value(0); textErr == nil {
				caption, _ = text.(string)
			}
			captionValue.clear()
		}
		if !strings.Contains(strings.ToLower(caption), "immediate") && id != 2554 {
			control.release()
			continue
		}
		_, executeErr := control.call("Execute")
		control.release()
		if executeErr != nil {
			return 0, false, executeErr
		}
		commandFound = true
		break
	}
	if !commandFound {
		return 0, false, Fail("vba_immediate_unavailable", "The VBE did not expose its Immediate Window command", map[string]any{"command_ids_tried": []int{2554, 2557, 2558, 2559}})
	}
	deadline := time.Now().Add(2 * time.Second)
	for time.Now().Before(deadline) {
		pump()
		candidates := immediateWindows(pid)
		if len(candidates) == 1 {
			return uint64(candidates[0]), true, nil
		}
		time.Sleep(20 * time.Millisecond)
	}
	return 0, false, Fail("vba_immediate_unavailable", "The owned Immediate Window did not appear after its VBE command", map[string]any{"windows": windowInventory(pid)})
}

func immediateWindows(pid uint32) []uintptr {
	all := map[uintptr]bool{}
	var visit func(uintptr, int)
	visit = func(hwnd uintptr, depth int) {
		if hwnd == 0 || all[hwnd] || depth > 8 || pidOf(hwnd) != pid {
			return
		}
		all[hwnd] = true
		children := windowEnum{pid: pid}
		enumChildren.Call(hwnd, collectWindowsCallback, uintptr(unsafe.Pointer(&children)))
		for _, child := range children.windows {
			visit(child, depth+1)
		}
	}
	for _, root := range windowsFor(pid) {
		visit(root, 0)
	}
	result := []uintptr{}
	for hwnd := range all {
		if windowVisibleValue(hwnd) == 0 {
			continue
		}
		title := strings.ToLower(strings.TrimSpace(textOf(hwnd)))
		class := strings.ToLower(classOf(hwnd))
		if title == "immediate" || strings.Contains(title, "immediate window") || strings.Contains(class, "immediate") {
			result = append(result, hwnd)
		}
	}
	return result
}

func windowVisibleValue(hwnd uintptr) uintptr {
	visible, _, _ := windowVisible.Call(hwnd)
	return visible
}

func windowDetails(windows []uintptr) []map[string]any {
	result := make([]map[string]any, 0, len(windows))
	for _, hwnd := range windows {
		result = append(result, map[string]any{"hwnd": uint64(hwnd), "class": classOf(hwnd), "title": textOf(hwnd)})
	}
	return result
}

func locateImmediateEditor(hwnd uint64) (immediateEditor, error) {
	root, err := accessibleRoot(uintptr(hwnd))
	if err != nil {
		return immediateEditor{}, err
	}
	defer root.release()
	budget := 3000
	path, value, ok := findImmediateEditor(root, 0, nil, 8, &budget)
	if !ok {
		return immediateEditor{}, Fail("vba_immediate_editor_unavailable", "The owned Immediate Window did not expose an editable accessibility value", map[string]any{"hwnd": hwnd, "budget_remaining": budget})
	}
	child := 0
	if len(path) > 0 {
		last := path[len(path)-1]
		if n, ok := last.(float64); ok {
			// The final path entry identifies the child object; its value is the
			// accessible child ID on that object, recovered during the search.
			_ = n
		}
	}
	// findImmediateEditor returns the child ID in the last synthetic path
	// marker. Keep the marker out of the public selector path.
	if marker, ok := pathValueChild(path); ok {
		path = marker.path
		child = marker.child
	}
	return immediateEditor{window: hwnd, path: path, child: child, value: value}, nil
}

type editorPathChild struct {
	path  []any
	child int
}

func pathValueChild(path []any) (editorPathChild, bool) {
	if len(path) == 0 {
		return editorPathChild{}, false
	}
	if marker, ok := path[len(path)-1].(editorPathChild); ok {
		return marker, true
	}
	return editorPathChild{}, false
}

func findImmediateEditor(a dispatch, child int, path []any, depth int, budget *int) ([]any, string, bool) {
	if *budget <= 0 {
		return nil, "", false
	}
	*budget--
	role, _ := accProperty(a, "accRole", child)
	value, err := accProperty(a, "accValue", child)
	if err == nil {
		if text, ok := value.(string); ok && fmt.Sprint(role) == "42" {
			return append(append([]any(nil), path...), editorPathChild{path: append([]any(nil), path...), child: child}), text, true
		}
	}
	if child != 0 || depth == 0 {
		return nil, "", false
	}
	entries, err := childrenOf(a)
	if err != nil {
		return nil, "", false
	}
	defer func() {
		for i := range entries {
			entries[i].clear()
		}
	}()
	for i := range entries {
		if *budget <= 0 {
			break
		}
		entry := &entries[i]
		if entry.VT == 3 {
			if foundPath, foundValue, ok := findImmediateEditor(a, int(int32(entry.Value)), path, 0, budget); ok {
				return foundPath, foundValue, true
			}
			continue
		}
		if entry.VT != 9 {
			continue
		}
		object, err := entry.object()
		if err != nil {
			continue
		}
		next, err := object.query(&iidAccessible)
		object.release()
		if err != nil {
			continue
		}
		nextPath := append(append([]any(nil), path...), float64(i))
		foundPath, foundValue, ok := findImmediateEditor(next, 0, nextPath, depth-1, budget)
		next.release()
		if ok {
			return foundPath, foundValue, true
		}
	}
	return nil, "", false
}

func postImmediateEnter(hwnd uint64) error {
	input := uintptr(hwnd)
	for _, child := range descendantWindows(uintptr(hwnd), 4) {
		class := strings.ToLower(classOf(child))
		if strings.Contains(class, "edit") || strings.Contains(class, "richedit") {
			input = child
			break
		}
	}
	setFocus := user32.NewProc("SetFocus")
	setFocus.Call(input)
	sendMessage := user32.NewProc("SendMessageW")
	sendMessage.Call(input, wmKeyDown, vkReturn, 0)
	sendMessage.Call(input, wmKeyUp, vkReturn, 0)
	pump()
	return nil
}

func descendantWindows(hwnd uintptr, depth int) []uintptr {
	if hwnd == 0 || depth <= 0 {
		return nil
	}
	q := windowEnum{pid: pidOf(hwnd)}
	enumChildren.Call(hwnd, collectWindowsCallback, uintptr(unsafe.Pointer(&q)))
	result := append([]uintptr(nil), q.windows...)
	for _, child := range q.windows {
		result = append(result, descendantWindows(child, depth-1)...)
	}
	return result
}

func closeOwnedWindow(hwnd uint64) error {
	posted, _, err := user32.NewProc("PostMessageW").Call(uintptr(hwnd), wmClose, 0, 0)
	if posted == 0 {
		return fmt.Errorf("PostMessageW(WM_CLOSE) failed: %v", err)
	}
	return nil
}
