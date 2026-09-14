//go:build windows && (amd64 || arm64)

package native

import "testing"

func TestDebugSymbolsAreBoundedAndStable(t *testing.T) {
	got := debugSymbols([]any{"value", " value ", "", "value", "line\nbreak", 42})
	if len(got) != 1 || got[0] != "value" {
		t.Fatalf("unexpected symbol request normalization: %#v", got)
	}
}
