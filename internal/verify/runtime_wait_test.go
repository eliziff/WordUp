package verify

import (
	"context"
	"errors"
	"runtime"
	"testing"

	"github.com/eliziff/WordUp/internal/native"
)

type runtimeFailureHost struct {
	mockHost
	cancel    context.CancelFunc
	completed bool
	finished  chan struct{}
	forgotten int
}

func (h *runtimeFailureHost) Call(ctx context.Context, op native.Operation) (any, error) {
	switch op.Op {
	case "forget":
		h.forgotten++
		return nil, nil
	case "poll":
		status := "running"
		if h.completed {
			status = "completed"
		}
		return map[string]any{"status": status}, nil
	case "begin":
		return map[string]any{"task": "runtime-test"}, nil
	case "ui.diagnostics":
		return map[string]any{"dialogs": []any{map[string]any{"hwnd": 123, "title": "Microsoft Visual Basic", "buttons": []any{"End", "Debug"}, "messages": []any{"Run-time error '5':\nInvalid procedure call"}}}}, nil
	case "ui.invoke":
		if op.HWND != 123 || op.Named["name"] != "End" || op.Named["message_pattern"] == nil {
			return nil, errors.New("unscoped acknowledgement")
		}
		if h.finished != nil {
			close(h.finished)
		} else {
			h.cancel() // The deadline expires while Word is unwinding.
		}
		return nil, nil
	}
	return h.mockHost.Call(ctx, op)
}

func TestPollClassifiesOnlyPendingRuntimeFailure(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows observer")
	}
	for _, completed := range []bool{false, true} {
		ctx, cancel := context.WithCancel(context.Background())
		h := &runtimeFailureHost{cancel: cancel, completed: completed}
		_, err := observedCall(ctx, h, native.Operation{Op: "poll", Value: "callback"})
		cancel()
		var fault *native.Fault
		if completed && err != nil {
			t.Fatal(err)
		}
		if !completed && (!errors.As(err, &fault) || fault.Code != "vba_runtime_error") {
			t.Fatal(err)
		}
	}
}

func (h *runtimeFailureHost) WaitTask(ctx context.Context, task string) (any, error) {
	if h.finished != nil {
		<-h.finished
		return map[string]any{"status": "completed", "duration_ms": 123, "result": map[string]any{"action_ms": 100}}, nil
	}
	<-ctx.Done()
	return nil, ctx.Err()
}

func TestCompletedRuntimeFailureRetainsTimingAndReleasesTask(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows observer")
	}
	h := &runtimeFailureHost{finished: make(chan struct{})}
	_, err := observedCall(context.Background(), h, native.Operation{Op: "ui.invoke"})
	var fault *native.Fault
	if !errors.As(err, &fault) || fault.Code != "vba_runtime_error" || h.forgotten != 1 {
		t.Fatal(err, h.forgotten)
	}
	completion := fault.Details.(map[string]any)["task_completion"].(map[string]any)
	if completion["duration_ms"] != 123 {
		t.Fatal(completion)
	}
}

func TestRuntimeFailureSurvivesDeadlineDuringUnwind(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("asynchronous observer is Windows-only")
	}
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	h := &runtimeFailureHost{cancel: cancel}
	_, err := observedCall(ctx, h, native.Operation{Op: "run", Macro: "FailureProof.Fail"})
	var fault *native.Fault
	if !errors.As(err, &fault) || fault.Code != "vba_runtime_error" {
		t.Fatal("lost runtime evidence at deadline", err)
	}
	if h.forgotten != 0 {
		t.Fatal("forgot a task without completion evidence")
	}
}

func TestRuntimeDialogRequiresVBEErrorAndActions(t *testing.T) {
	dialog := map[string]any{"title": "Microsoft Visual Basic", "buttons": []any{"End", "Debug", "Help"}, "messages": []any{"Run-time error '-2147220991 (80040201)':\n\nDeliberate failure"}}
	details := runtimeDialog(dialog)
	if details == nil || details["number"] != int64(-2147220991) || details["description"] != "Deliberate failure" {
		t.Fatal(details)
	}
	for key, value := range map[string]any{"title": "Other application", "buttons": []any{"OK"}, "messages": []any{"Expected operation completed"}} {
		changed := map[string]any{}
		for k, v := range dialog {
			changed[k] = v
		}
		changed[key] = value
		if runtimeDialog(changed) != nil {
			t.Fatal("misclassified ordinary prompt", changed)
		}
	}
}

type breakNoticeHost struct {
	mockHost
	dialog  map[string]any
	invoked bool
}

func (h *breakNoticeHost) Call(_ context.Context, op native.Operation) (any, error) {
	if op.Op == "ui.diagnostics" {
		return map[string]any{"dialogs": []any{h.dialog}}, nil
	}
	if op.Op != "ui.invoke" || op.HWND != 123 || op.Named["name"] != "OK" || op.Named["message_pattern"] != "^Can't execute code in break mode$" {
		return nil, errors.New("unexpected recovery operation")
	}
	h.invoked = true
	return nil, nil
}

func TestBreakNoticeRecoveryLeavesOtherPromptsAlone(t *testing.T) {
	for _, tc := range []struct {
		title, message string
		dismiss        bool
	}{
		{"Microsoft Visual Basic for Applications", "Can't execute code in break mode", true},
		{"Microsoft Visual Basic for Applications", "Save changes?", false},
		{"Other application", "Can't execute code in break mode", false},
	} {
		h := &breakNoticeHost{dialog: map[string]any{"hwnd": 123, "title": tc.title, "messages": []any{tc.message}}}
		details := map[string]any{}
		if got := dismissBreakModeNotice(context.Background(), h, details); got != tc.dismiss || h.invoked != tc.dismiss {
			t.Fatalf("wrong recovery for %+v: %v", tc, details)
		}
		if tc.dismiss && details["break_mode_notices"] == nil {
			t.Fatal("discarded the secondary failure evidence")
		}
	}
}

// A dead UI lane cannot provide post-failure diagnostics. Retain the last
// inventory taken before the action, explicitly labelled as earlier evidence.
type timeoutUIHost struct {
	mockHost
	cancel context.CancelFunc
}

func (h *timeoutUIHost) Call(ctx context.Context, op native.Operation) (any, error) {
	switch op.Op {
	case "ui.windows":
		return []any{map[string]any{"hwnd": 321, "title": "Owned document"}}, nil
	case "begin":
		return map[string]any{"task": "timeout-ui"}, nil
	}
	return nil, native.Fail("session_closed", "worker closed", nil)
}
func (h *timeoutUIHost) WaitTask(ctx context.Context, task string) (any, error) {
	h.cancel()
	return nil, native.Fail("native_deadline", "worker deadline", map[string]any{"worker_log_tail": "last worker signal"})
}
func TestUITimeoutRetainsEarlierInventoryAndCause(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows observer")
	}
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	h := &timeoutUIHost{cancel: cancel}
	_, err := observedCall(ctx, h, native.Operation{Op: "ui.invoke"})
	var fault *native.Fault
	if !errors.As(err, &fault) || fault.Code != "native_step_timeout" {
		t.Fatal(err)
	}
	details := fault.Details.(map[string]any)
	before := details["before_operation_ui"].(map[string]any)
	if before["phase"] != "before_operation" || len(before["windows"].([]any)) != 1 {
		t.Fatal(before)
	}
	if details["cause"] == nil {
		t.Fatal("missing timeout cause", details)
	}
	// The direct waiter preserves its structured worker failure as well.
	ctx, cancel = context.WithCancel(context.Background())
	defer cancel()
	h.cancel = cancel
	_, err = observedCall(ctx, h, native.Operation{Op: "new"})
	if !errors.As(err, &fault) {
		t.Fatal(err)
	}
	details = fault.Details.(map[string]any)
	cause := details["cause"].(*native.Fault)
	if cause.Code != "native_deadline" || cause.Details.(map[string]any)["worker_log_tail"] != "last worker signal" {
		t.Fatal(cause)
	}
}
