package verify

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
	"wordwright.local/internal/native"
	"wordwright.local/internal/office"
)

func TestAssertions(t *testing.T) {
	value := map[string]any{"a/b": []any{1, "Citation"}, "x": 1.01}
	for _, a := range []Assertion{{Path: "/a~1b/0", Kind: "equals", Expected: 1}, {Path: "/a~1b/1", Kind: "contains", Expected: "itat"}, {Path: "/x", Kind: "near", Expected: 1, Tolerance: .02}} {
		if e := Check(value, a); e != nil {
			t.Fatal(e)
		}
	}
	for _, a := range []Assertion{{Path: "/missing", Kind: "equals", Expected: nil}, {Kind: "simulated_pass"}, {Path: "/x", Kind: "equals", Expected: 1}, {Path: "/a~1b/5", Kind: "equals", Expected: nil}} {
		if e := Check(value, a); e == nil {
			t.Fatal("false pass", a)
		}
	}
}

type mockHost struct{ n int }

func (h *mockHost) Call(context.Context, native.Operation) (any, error) { h.n++; return 42, nil }
func (h *mockHost) Close() error                                        { return nil }
func (h *mockHost) Info() map[string]any                                { return map[string]any{"mock_only": true} }
func TestSuiteGuardAndEvidence(t *testing.T) {
	if e := (Suite{Schema: 1, Name: "empty"}).Validate(); e == nil {
		t.Fatal("empty suite passed")
	}
	p := office.BlankPackage()
	b, _ := p.Bytes()
	file := filepath.Join(t.TempDir(), "test.docx")
	os.WriteFile(file, b, 0600)
	s := Suite{Schema: 1, Name: "contract mock", Steps: []Step{{Name: "observation", Operation: native.Operation{Op: "get", Member: "Version"}, Assert: []Assertion{{Kind: "equals", Expected: 42}}}}}
	h := &mockHost{}
	r, e := Run(context.Background(), file, s, h, false)
	if e == nil || r.WordExecuted || h.n != 0 {
		t.Fatal("unauthorized runtime")
	}
	r, e = Run(context.Background(), file, s, h, true)
	if e != nil || r.Status != "passed" || r.FreshProcess || r.VBACompiled || r.SHA256 != office.Hash(b) {
		t.Fatal(r, e)
	}
	s.RequireCompile = true
	r, e = Run(context.Background(), file, s, &mockHost{}, true)
	if e == nil || r.Status == "passed" {
		t.Fatal("fake compilation pass")
	}
	raw, _ := json.Marshal(r)
	if len(raw) == 0 {
		t.Fatal("no report")
	}
}
