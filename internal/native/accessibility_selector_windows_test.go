//go:build windows && (amd64 || arm64)

package native

import (
	"os"
	"testing"
)

func TestMissingUISelector(t *testing.T) {
	for _, wait := range []float64{0, 1} {
		_, err := uiOperation(uint32(os.Getpid()), t.TempDir(), false, Operation{
			Op: "ui.find", Named: map[string]any{"name": "Missing control", "scope": "form", "wait_ms": wait},
		})
		fault, ok := err.(*Fault)
		if !ok || fault.Code != "ui_selector_not_found" {
			t.Fatalf("missing selector: %v", err)
		}
		details := fault.Details.(map[string]any)
		if details["selector"].(map[string]any)["name"] != "Missing control" {
			t.Fatalf("missing attempted selector: %v", details)
		}
		if roots, ok := details["root_searches"].([]map[string]any); !ok || len(roots) != 0 {
			t.Fatalf("missing evidence that no form roots were available: %v", details)
		}
	}
}

func TestDialogSelectorRejectsInvalidMessageConstraint(t *testing.T) {
	for _, named := range []map[string]any{
		{"name": "OK", "scope": "dialog", "message_pattern": "["},
		{"name": "OK", "scope": "dialog", "message_pattern": ""},
		{"name": "OK", "scope": "form", "message_pattern": "expected"},
	} {
		_, err := uiOperation(uint32(os.Getpid()), t.TempDir(), false, Operation{Op: "ui.find", Named: named})
		if err == nil {
			t.Fatal("accepted invalid dialog constraint", named)
		}
		if _, searching := err.(*Fault); searching {
			t.Fatal("invalid constraint reached control search", err)
		}
	}
}
