package native

import (
	"fmt"
	"testing"
)

func TestFaultRetainsWrappedNativeDetails(t *testing.T) {
	original := &Fault{Code: "word_automation_error", Message: "Edit failed", Details: map[string]any{"hresult": "0x800A17EC", "member": "Text"}}
	wrapped := fmt.Errorf("cleanup of owned operation: %w", original)
	got := fault(wrapped)
	if got.Code != original.Code || got.Message != wrapped.Error() || got.Details.(map[string]any)["member"] != "Text" {
		t.Fatalf("lost wrapped failure: %#v", got)
	}
	if original.Message != "Edit failed" || fault(original) != original || fault(nil) != nil {
		t.Fatal("mutated direct fault")
	}
	if got = fault(fmt.Errorf("ordinary failure")); got.Code != "native_error" || got.Message != "ordinary failure" {
		t.Fatal(got)
	}
}
