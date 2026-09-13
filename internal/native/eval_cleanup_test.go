package native

import "testing"

func TestEvaluationCleanupRetainsBothFailures(t *testing.T) {
	primary := Fail("scratch_vba_failed", "Bad range", map[string]any{"number": 5, "body_line": 12})
	cleanup := Fail("word_automation_error", "Object deleted", map[string]any{"member": "Installed"})
	if evaluationCleanupError(primary, nil, "scratch.dotm") != primary {
		t.Fatal("successful cleanup changed primary failure")
	}
	for _, cause := range []error{nil, primary} {
		got := evaluationCleanupError(cause, cleanup, "scratch.dotm").(*Fault)
		details := got.Details.(map[string]any)
		if details["cleanup_error"] != cleanup || details["scratch_path"] != "scratch.dotm" {
			t.Fatalf("lost cleanup evidence: %#v", got)
		}
		if cause == nil && got.Code != "scratch_cleanup_failed" {
			t.Fatal("failed cleanup reported success")
		}
		if cause != nil && (got.Code != "scratch_vba_failed" || details["body_line"] != 12 || details["number"] != 5) {
			t.Fatalf("lost original runtime location: %#v", got)
		}
	}
}
