package verify

import (
	"context"
	"fmt"
	"github.com/eliziff/WordUp/internal/native"
	"strings"
	"time"
)

// Only acknowledge a recognized compiler error while an explicit compile is pending.
// Other dialogs retain the normal timeout and captured diagnostics.
func waitCompile(ctx context.Context, h native.Host, wait func(context.Context, string) (any, error), task string) (any, error) {
	type completion struct {
		result any
		err    error
	}
	done := make(chan completion, 1)
	go func() { v, e := wait(ctx, task); done <- completion{v, e} }()
	timer := time.NewTicker(100 * time.Millisecond)
	defer timer.Stop()
	var captured any
	for {
		select {
		case c := <-done:
			if result, ok := c.result.(map[string]any); ok && result["error"] != nil && captured == nil {
				captured = ackCompilerDialog(ctx, h)
			}
			if captured != nil {
				if result, ok := c.result.(map[string]any); ok {
					if failure, ok := canonical(result["error"]).(map[string]any); ok {
						details, ok := failure["details"].(map[string]any)
						if !ok {
							details = map[string]any{"cause_details": failure["details"]}
						}
						details["compiler_dialog"] = captured
						failure["details"] = details
						result["error"] = failure
					}
				}
			}
			return c.result, c.err
		case <-ctx.Done():
			return nil, ctx.Err()
		case <-timer.C:
			if captured != nil {
				continue
			}
			captured = ackCompilerDialog(ctx, h)
		}
	}
}

func compilerErrorDialog(dialog map[string]any) bool {
	if dialog["title"] != "Microsoft Visual Basic for Applications" {
		return false
	}
	messages, _ := dialog["messages"].([]any)
	for _, message := range messages {
		if strings.HasPrefix(strings.TrimSpace(fmt.Sprint(message)), "Compile error:") {
			return true
		}
	}
	return false
}

func ackCompilerDialog(ctx context.Context, h native.Host) any {
	result, err := h.Call(ctx, native.Operation{Op: "ui.diagnostics", Named: map[string]any{"trees": false, "window": "Microsoft Visual Basic for Applications"}})
	if err != nil {
		return nil
	}
	m, _ := canonical(result).(map[string]any)
	dialogs, _ := m["dialogs"].([]any)
	for _, item := range dialogs {
		dialog, ok := item.(map[string]any)
		if !ok || !compilerErrorDialog(dialog) {
			continue
		}
		hwnd, ok := dialog["hwnd"].(float64)
		if !ok {
			continue
		}
		_, err = h.Call(ctx, native.Operation{Op: "ui.invoke", HWND: uint64(hwnd), Named: map[string]any{"name": "OK", "role": 43}})
		if err == nil {
			return dialog
		}
	}
	return nil
}
