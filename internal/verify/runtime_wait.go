package verify

import (
	"context"
	"fmt"
	"regexp"
	"strconv"
	"time"

	"github.com/eliziff/WordUp/internal/native"
)

var runtimeMessage = regexp.MustCompile(`(?s)^Run-time error '(-?\d+)(?: \(([0-9A-Fa-f]+)\))?':\s*(.+)$`)
var hiddenCompileMessage = regexp.MustCompile(`(?s)^Compile error in hidden module:\s*(.+)$`)

func runtimeDialog(dialog map[string]any) map[string]any {
	if dialog["title"] != "Microsoft Visual Basic" && dialog["title"] != "Microsoft Visual Basic for Applications" {
		return nil
	}
	buttons, _ := dialog["buttons"].([]any)
	end, debug := false, false
	for _, button := range buttons {
		end = end || button == "End"
		debug = debug || button == "Debug"
	}
	if !end || !debug {
		return nil
	}
	messages, _ := dialog["messages"].([]any)
	for _, message := range messages {
		match := runtimeMessage.FindStringSubmatch(fmt.Sprint(message))
		if match == nil {
			continue
		}
		number, err := strconv.ParseInt(match[1], 10, 32)
		if err != nil {
			continue
		}
		return map[string]any{"number": number, "hex_code": match[2], "description": match[3], "runtime_dialog": dialog}
	}
	return nil
}

// A recognized unhandled-error dialog is a failure, even if End makes Run
// return normally. Preserve its evidence before allowing the macro to unwind.
func waitRuntime(ctx context.Context, h native.Host, wait func(context.Context, string) (any, error), task string, op native.Operation) (any, error) {
	ctx, cancel := context.WithCancel(ctx)
	defer cancel()
	type completion struct {
		result any
		err    error
	}
	done := make(chan completion, 1)
	go func() { v, e := wait(ctx, task); done <- completion{v, e} }()
	timer := time.NewTicker(100 * time.Millisecond)
	defer timer.Stop()
	var failure error
	for {
		select {
		case c := <-done:
			if failure != nil {
				return c.result, failure
			}
			return c.result, c.err
		case <-ctx.Done():
			if failure != nil {
				return nil, failure
			}
			return nil, ctx.Err()
		case <-timer.C:
			if failure != nil {
				continue
			}
			failure = acknowledgeRuntimeError(ctx, h, task, op)
		}
	}
}

func acknowledgeRuntimeError(ctx context.Context, h native.Host, task string, op native.Operation) error {
	value, err := h.Call(ctx, native.Operation{Op: "ui.diagnostics", Named: map[string]any{"trees": false, "window_class": "#32770"}})
	if err != nil {
		return nil
	}
	m, _ := canonical(value).(map[string]any)
	dialogs, _ := m["dialogs"].([]any)
	for _, item := range dialogs {
		dialog, _ := item.(map[string]any)
		if dialog["title"] == "Microsoft Visual Basic" || dialog["title"] == "Microsoft Visual Basic for Applications" {
			messages, _ := dialog["messages"].([]any)
			for _, message := range messages {
				if !hiddenCompileMessage.MatchString(fmt.Sprint(message)) {
					continue
				}
				details := map[string]any{"operation": op, "task": task, "observed_dialogs": []any{dialog}}
				if hwnd, ok := dialog["hwnd"].(float64); ok {
					_, ackErr := h.Call(ctx, native.Operation{Op: "ui.invoke", HWND: uint64(hwnd), Named: map[string]any{"scope": "dialog", "name": "OK", "role": 43, "message_pattern": hiddenCompileMessage.String()}})
					if ackErr != nil {
						details["acknowledgement_error"] = ErrorValue(ackErr)
					}
				}
				return native.Fail("vba_compile_error", fmt.Sprint(message), details)
			}
		}
		details := runtimeDialog(dialog)
		if details == nil {
			continue
		}
		details["operation"], details["task"] = op, task
		failure := native.Fail("vba_runtime_error", fmt.Sprint(details["description"]), details)
		hwnd, ok := dialog["hwnd"].(float64)
		if !ok {
			return failure
		}
		if op.Op == "ui.invoke" || op.Op == "run" {
			_, debugErr := h.Call(ctx, native.Operation{Op: "ui.invoke", HWND: uint64(hwnd), Named: map[string]any{"scope": "dialog", "name": "Debug", "role": 43, "message_pattern": runtimeMessage.String()}})
			if debugErr == nil {
				stack, stackErr := h.Call(ctx, native.Operation{Op: "ui.vba.stack"})
				for attempt := 0; stackErr != nil && attempt < 2 && dismissBreakModeNotice(ctx, h, details); attempt++ {
					recoveryErrors, _ := details["stack_recovery_errors"].([]any)
					details["stack_recovery_errors"] = append(recoveryErrors, ErrorValue(stackErr))
					stack, stackErr = h.Call(ctx, native.Operation{Op: "ui.vba.stack"})
				}
				details["runtime_stack"] = stack
				if stackErr != nil {
					details["stack_error"] = ErrorValue(stackErr)
				}
				location, resetErr := h.Call(ctx, native.Operation{Op: "ui.vba.reset"})
				if resetErr != nil && dismissBreakModeNotice(ctx, h, details) {
					details["initial_reset_error"] = ErrorValue(resetErr)
					location, resetErr = h.Call(ctx, native.Operation{Op: "ui.vba.reset"})
				}
				details["runtime_location"] = location
				if resetErr != nil {
					details["reset_error"] = ErrorValue(resetErr)
				}
				return failure
			}
			details["debug_error"] = ErrorValue(debugErr)
		}
		_, err = h.Call(ctx, native.Operation{Op: "ui.invoke", HWND: uint64(hwnd), Named: map[string]any{"scope": "dialog", "name": "End", "role": 43, "message_pattern": runtimeMessage.String()}})
		if err != nil {
			details["end_error"] = ErrorValue(err)
		}
		return failure
	}
	return nil
}

// A queued Office callback can raise this secondary notice after Debug pauses
// the original failure. Dismiss only this recognized notice, retaining it as
// evidence; arbitrary modal prompts remain untouched. Callers bound retries.
func dismissBreakModeNotice(ctx context.Context, h native.Host, details map[string]any) bool {
	value, err := h.Call(ctx, native.Operation{Op: "ui.diagnostics", Named: map[string]any{"trees": false, "window_class": "#32770"}})
	if err != nil {
		return false
	}
	m, _ := canonical(value).(map[string]any)
	dialogs, _ := m["dialogs"].([]any)
	for _, item := range dialogs {
		d, _ := item.(map[string]any)
		if d["title"] != "Microsoft Visual Basic for Applications" && d["title"] != "Microsoft Visual Basic" {
			continue
		}
		messages, _ := d["messages"].([]any)
		if len(messages) != 1 || messages[0] != "Can't execute code in break mode" {
			continue
		}
		hwnd, ok := d["hwnd"].(float64)
		if !ok {
			continue
		}
		notices, _ := details["break_mode_notices"].([]any)
		details["break_mode_notices"] = append(notices, d)
		_, err := h.Call(ctx, native.Operation{Op: "ui.invoke", HWND: uint64(hwnd), Named: map[string]any{"scope": "dialog", "name": "OK", "role": 43, "message_pattern": "^Can't execute code in break mode$"}})
		if err != nil {
			details["break_mode_notice_error"] = ErrorValue(err)
		}
		return err == nil
	}
	return false
}
