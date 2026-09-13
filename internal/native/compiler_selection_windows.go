//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"runtime"
	"strings"
	"time"
)

func compilerSelection(vbe dispatch) map[string]any {
	details := map[string]any{"location_kind": "native_vbe_selection", "location_source": "VBE.ActiveCodePane.GetSelection"}
	pane, err := objectProperty(vbe, "ActiveCodePane")
	if err != nil {
		details["selection_error"] = fault(err)
		return details
	}
	defer pane.release()
	var startLine, startColumn, endLine, endColumn int32
	v, err := pane.call("GetSelection", &startLine, &startColumn, &endLine, &endColumn)
	v.clear()
	if err != nil {
		details["selection_error"] = fault(err)
		return details
	}
	details["line"], details["column"] = startLine, startColumn
	details["end_line"], details["end_column"] = endLine, endColumn
	module, err := objectProperty(pane, "CodeModule")
	if err != nil {
		details["source_error"] = fault(err)
		return details
	}
	defer module.release()
	for _, property := range []string{"Name", "Lines"} {
		var args []any
		if property == "Lines" {
			args = []any{int(startLine), 1}
		}
		v, err := module.get(property, args...)
		if err != nil {
			details["source_error"] = fault(err)
			continue
		}
		value, err := v.value(0)
		v.clear()
		if err != nil {
			details["source_error"] = fault(err)
			continue
		}
		if property == "Name" {
			details["module"] = value
		} else {
			details["line_text"] = value
		}
	}
	return details
}

func inspectVBA(app dispatch, execute, reset bool) (any, error) {
	if reset && !execute {
		return nil, Fail("execution_not_authorized", "Resetting VBA requires explicit execute capability", nil)
	}
	vbe, err := objectProperty(app, "VBE")
	if err != nil {
		return nil, err
	}
	defer vbe.release()
	project, err := objectProperty(vbe, "ActiveVBProject")
	if err != nil {
		return nil, err
	}
	defer project.release()
	state := compilerSelection(vbe)
	for _, property := range []string{"Name", "Mode"} {
		v, err := project.get(property)
		if err != nil {
			return nil, err
		}
		value, err := v.value(0)
		v.clear()
		if err != nil {
			return nil, err
		}
		state["project_"+strings.ToLower(property)] = value
	}
	if reset {
		if state["project_mode"] != int32(1) {
			return nil, Fail("vba_not_paused", "Reset requires a paused VBA project", state)
		}
		document, documentErr := objectProperty(app, "ActiveDocument")
		if documentErr == nil {
			defer document.release()
		} else {
			state["document_error"] = fault(documentErr)
		}
		bars, err := objectProperty(vbe, "CommandBars")
		if err != nil {
			return nil, err
		}
		defer bars.release()
		v, err := bars.invoke("FindControl", 1, nil, map[string]any{"ID": 228}, nil)
		if err != nil {
			return nil, err
		}
		control, err := v.object()
		v.clear()
		if err != nil {
			return nil, err
		}
		defer control.release()
		v, err = control.call("Execute")
		v.clear()
		if err != nil {
			return nil, err
		}
		v, err = project.get("Mode")
		if err != nil {
			return nil, err
		}
		mode, err := v.value(0)
		v.clear()
		if err != nil {
			return nil, err
		}
		state["project_mode_after"] = mode
		if mode != int32(2) {
			return nil, Fail("vba_reset_not_confirmed", "VBA did not return to design mode", state)
		}
		state["reset"] = true
		// Reset can reactivate Word's seed document while Run unwinds.
		// Restore the paused document so automatic XML records the actual edit.
		if documentErr == nil {
			v, err = document.call("Activate")
			v.clear()
			if err != nil {
				state["document_restore_error"] = fault(err)
			}
		}
	}
	return state, nil
}

func vbaStack(app dispatch, pid uint32, directory string, execute bool) (any, error) {
	if !execute {
		return nil, Fail("execution_not_authorized", "Opening the VBA stack dialog requires execute capability", nil)
	}
	state, err := inspectVBA(app, execute, false)
	if err != nil {
		return nil, err
	}
	if state.(map[string]any)["project_mode"] != int32(1) {
		return nil, Fail("vba_not_paused", "Call Stack requires a paused VBA project", state)
	}
	find := func() uint64 {
		for _, item := range windowInventory(pid) {
			window := item.(map[string]any)
			if window["class"] == "#32770" && window["title"] == "Call Stack" && window["visible_on_private_desktop"] == true {
				return window["hwnd"].(uint64)
			}
		}
		return 0
	}
	hwnd := find()
	opened := hwnd == 0
	var invocation <-chan error
	if opened {
		vbe, err := objectProperty(app, "VBE")
		if err != nil {
			return nil, err
		}
		defer vbe.release()
		bars, err := objectProperty(vbe, "CommandBars")
		if err != nil {
			return nil, err
		}
		defer bars.release()
		value, err := bars.invoke("FindControl", 1, nil, map[string]any{"ID": 620}, nil)
		if err != nil {
			return nil, err
		}
		control, err := value.object()
		value.clear()
		if err != nil {
			return nil, err
		}
		commandInfo := map[string]any{}
		for _, property := range []string{"Caption", "Enabled", "Id"} {
			v, e := control.get(property)
			if e == nil {
				commandInfo[property], _ = v.value(0)
			}
			v.clear()
		}
		stream, err := marshalDispatch(control)
		control.release()
		if err != nil {
			return nil, err
		}
		done := make(chan error, 1)
		invocation = done
		go func() {
			runtime.LockOSThread()
			defer runtime.UnlockOSThread()
			hr, _, _ := coInit.Call(0, 2)
			if failed(hr) {
				ole32.NewProc("CoReleaseMarshalData").Call(stream)
				(dispatch{stream}).release()
				done <- fmt.Errorf("Call Stack CoInitializeEx HRESULT 0x%08X", uint32(hr))
				return
			}
			defer coUninit.Call()
			command, err := unmarshalDispatch(stream)
			if err != nil {
				done <- err
				return
			}
			defer command.release()
			value, err := command.call("Execute")
			value.clear()
			done <- err
		}()
		deadline := time.Now().Add(2 * time.Second)
		for hwnd == 0 && time.Now().Before(deadline) {
			select {
			case err := <-invocation:
				if err != nil {
					return nil, err
				}
				invocation = nil
			default:
			}

			pump()
			hwnd = find()
			if hwnd == 0 {
				time.Sleep(10 * time.Millisecond)
			}
		}
		if hwnd == 0 {
			return nil, Fail("vba_stack_unavailable", "The owned Call Stack dialog did not appear", commandInfo)
		}
	}
	tree, captureErr := uiOperation(pid, directory, execute, Operation{Op: "ui.tree", HWND: hwnd, Depth: 6})
	frames := []string{}
	truncated := false
	var visit func(map[string]any)
	visit = func(node map[string]any) {
		if node["role"] == int32(34) {
			name, _ := node["name"].(string)
			if name == "" {
				truncated = true
			} else {
				frames = append(frames, name)
			}
		}
		truncated = truncated || node["truncated"] == true || node["children_error"] != nil
		children, _ := node["children"].([]any)
		for _, child := range children {
			visit(child.(map[string]any))
		}
	}
	if captureErr == nil {
		visit(tree.(map[string]any))
	}
	result := map[string]any{"frames": frames, "truncated": truncated, "source": "native VBE Call Stack dialog"}
	if truncated || len(frames) == 0 {
		result["tree"] = tree
	}
	if opened {
		_, closeErr := uiOperation(pid, directory, execute, Operation{Op: "ui.close", HWND: hwnd})
		if closeErr != nil {
			result["close_error"] = fault(closeErr)
		} else {
			deadline := time.Now().Add(time.Second)
			for find() == hwnd && time.Now().Before(deadline) {
				pump()
				time.Sleep(10 * time.Millisecond)
			}
			if find() == hwnd {
				result["close_error"] = "Call Stack dialog remained open"
			}
		}
	}
	if invocation != nil {
		select {
		case err := <-invocation:
			if err != nil {
				result["command_error"] = fault(err)
			}
		case <-time.After(time.Second):
			result["command_error"] = "Call Stack command has not returned"
		}
	}
	if captureErr != nil {
		return result, captureErr
	}
	if len(frames) == 0 {
		return result, Fail("vba_stack_unavailable", "Call Stack contained no readable frames", result)
	}
	return result, nil
}
