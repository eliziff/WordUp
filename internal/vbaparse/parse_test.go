package vbaparse

import (
	"os"
	"strings"
	"testing"
)

func BenchmarkRealModule(b *testing.B) {
	path := os.Getenv("WORDUP_PROFILE_VBA")
	if path == "" {
		b.Skip("set WORDUP_PROFILE_VBA to a module path")
	}
	source, err := os.ReadFile(path)
	if err != nil {
		b.Fatal(err)
	}
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		r, e := Parse(string(source))
		if e != nil || r["syntax_valid"] != true {
			b.Fatal(r, e)
		}
	}
}

func TestRubberduckSyntaxDiagnostics(t *testing.T) {
	for _, tc := range []struct {
		source string
		valid  bool
	}{
		{"Option Explicit\nPublic Sub Example()\nDim x As Long\nx = 1 + 2\nEnd Sub", true},
		{"Public Sub Example()\nIf True Then\nDebug.Print 1\nEnd If\nEnd Sub\n", true},
		{"Public Sub Example()\nDim x As\nEnd Sub\n", false},
		{"Public Sub Example()\nIf True Then\nEnd Sub\n", false},
	} {
		r, e := Parse(tc.source)
		if e != nil || r["syntax_valid"] != tc.valid {
			t.Fatalf("%q: %v %v", tc.source, r, e)
		}
	}
}

func TestLargeModuleStillChecksTrailingSyntax(t *testing.T) {
	prefix := strings.Repeat("' ordinary module comment\n", 6000)
	for _, tc := range []struct {
		body  string
		valid bool
	}{
		{"Public Sub Example()\nEnd Sub\n", true},
		{"Public Sub Example()\nDim x As\nEnd Sub\n", false},
	} {
		result, err := Parse(prefix + tc.body)
		if err != nil || result["syntax_valid"] != tc.valid {
			t.Fatal(result, err)
		}
	}
	if _, err := Parse(strings.Repeat("x", (1<<20)+1)); err == nil {
		t.Fatal("module size guard missing")
	}
}
