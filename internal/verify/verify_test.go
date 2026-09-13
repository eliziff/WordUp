package verify

import (
	"context"
	"encoding/json"
	"fmt"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
)

func TestWrappedFaultPreservesDiagnostics(t *testing.T) {
	fault := native.Fail("scratch_vba_failed", "Unsupported method", map[string]any{"number": 438, "line": 12})
	err := fmt.Errorf("step conversion: %w", fault)
	got, ok := ErrorValue(err).(*native.Fault)
	if !ok || got.Code != "scratch_vba_failed" || got.Message != err.Error() {
		t.Fatalf("lost contextual fault: %#v", got)
	}
	if got.Details.(map[string]any)["line"] != 12 || ErrorValue(fault) != fault || ErrorValue(nil) != nil {
		t.Fatal("lost diagnostic details or changed direct faults")
	}
}

func TestCompileDialogSummaryRetainsLocation(t *testing.T) {
	err := native.Fail("compile_not_confirmed", "Compiler did not finish", map[string]any{"module": "Broken", "line": 2})
	diagnostics := map[string]any{"dialogs": []any{map[string]any{"title": "VBE", "messages": []any{"Expected: identifier"}}}, "windows": "large raw evidence"}
	f := failureDialogs(err, diagnostics).(*native.Fault)
	d := f.Details.(map[string]any)
	if d["module"] != "Broken" || d["line"] != 2 || len(d["observed_dialogs"].([]any)) != 1 || d["windows"] != nil {
		t.Fatalf("lost compact failure evidence: %+v", f)
	}
	if err.(*native.Fault).Details.(map[string]any)["observed_dialogs"] != nil {
		t.Fatal("mutated original fault")
	}
	if failureDialogs(err, map[string]any{}) != err {
		t.Fatal("invented missing dialog")
	}
}

func TestAssertionFailurePreservesMissingVersusNull(t *testing.T) {
	for _, present := range []bool{false, true} {
		value := map[string]any{}
		if present {
			value["result"] = nil
		}
		a := Assertion{Path: "/result", Kind: "equals", Expected: "formatted"}
		failure := assertionFailure(value, a, 2, Check(value, a))
		fault := ErrorValue(fmt.Errorf("step formatting: %w", failure)).(*native.Fault)
		details := fault.Details.(map[string]any)
		if fault.Code != "assertion_failed" || details["actual_available"] != present || details["assertion_index"] != 2 || details["expected"] != "formatted" {
			t.Fatalf("lost actionable assertion evidence: %#v", fault)
		}
		_, hasPointerError := details["pointer_error"]
		if hasPointerError == present {
			t.Fatalf("missing and null conflated: %#v", details)
		}
	}
}

func TestAssertions(t *testing.T) {
	if err := Check(map[string]any{"status": "completed"}, Assertion{Path: "/error", Kind: "absent"}); err != nil {
		t.Fatal(err)
	}
	if err := Check(map[string]any{"error": nil}, Assertion{Path: "/error", Kind: "absent"}); err == nil {
		t.Fatal("present null treated as absent")
	}
	for _, path := range []string{"/missing/error", "/error~2", "/error~", "error", ""} {
		if err := Check(map[string]any{}, Assertion{Path: path, Kind: "absent"}); err == nil {
			t.Fatalf("invalid absence evidence accepted for %q", path)
		}
	}
	if err := Check(map[string]any{"result": map[string]any{}}, Assertion{Path: "/result/error", Kind: "absent"}); err != nil {
		t.Fatal(err)
	}
	for _, op := range []string{"poll", "run"} {
		s := Suite{Schema: 1, Name: "wait", Steps: []Step{{Name: "completion", Operation: native.Operation{Op: op}, EventuallyMS: 100, Assert: []Assertion{{Path: "/error", Kind: "absent"}}}}}
		if err := s.Validate(); (err == nil) != (op == "poll") {
			t.Fatalf("retry safety for %s: %v", op, err)
		}
	}
	if err := Check(2, Assertion{Kind: "less_than", Expected: 3}); err != nil {
		t.Fatal(err)
	}
	if err := Check(3, Assertion{Kind: "less_than", Expected: 3}); err == nil {
		t.Fatal("performance bound accepted its limit")
	}
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

func TestSuiteXMLComparison(t *testing.T) {
	dir := t.TempDir()
	a, b := filepath.Join(dir, "before.xml"), filepath.Join(dir, "after.xml")
	for path, text := range map[string]string{a: `<p id="1">keep</p>`, b: `<p id="2">keep</p>`} {
		if err := os.WriteFile(path, []byte(text), 0600); err != nil {
			t.Fatal(err)
		}
	}
	op := native.Operation{Op: "xml.compare", Target: a, File: b}
	result, err := observedCall(context.Background(), nil, op)
	if err != nil || result.(map[string]any)["equal"] != false {
		t.Fatal(result, err)
	}
	op.Named = map[string]any{"xml_policy": map[string]any{"attributes": []any{map[string]any{"Space": "", "Local": "id"}}}}
	result, err = observedCall(context.Background(), nil, op)
	if err != nil || result.(map[string]any)["equal"] != true {
		t.Fatal(result, err)
	}
	op.Named = map[string]any{"xml_policy": map[string]any{"typo": true}}
	if _, err = observedCall(context.Background(), nil, op); err == nil {
		t.Fatal("invalid policy accepted")
	}
}

type mockHost struct{ n int }

type failingCleanupHost struct{ mockHost }

func (h *failingCleanupHost) Call(ctx context.Context, op native.Operation) (any, error) {
	switch op.Op {
	case "eval.prepare":
		return map[string]any{"program": "test-program"}, nil
	case "begin":
		return nil, native.Fail("primary_failure", "original operation failed", nil)
	case "eval.release":
		return nil, native.Fail("release_failure", "scratch release failed", nil)
	}
	return h.mockHost.Call(ctx, op)
}

func TestCleanupFailureDoesNotReplacePrimaryFailure(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows prepared scratch path")
	}
	b, _ := office.BlankPackage().Bytes()
	file := filepath.Join(t.TempDir(), "test.docx")
	if err := os.WriteFile(file, b, 0600); err != nil {
		t.Fatal(err)
	}
	s := Suite{Schema: 1, Name: "failure retention", Steps: []Step{
		{Name: "first", Operation: native.Operation{Op: "eval", Value: "Evaluate=1"}, Assert: []Assertion{{Kind: "equals", Expected: 1}}},
		{Name: "second", Operation: native.Operation{Op: "eval", Value: "Evaluate=2"}},
	}}
	r, err := Run(context.Background(), file, s, &failingCleanupHost{}, true)
	cleanup, ok := r.CleanupErrors["scratch_release"].(*native.Fault)
	if err == nil || !strings.Contains(err.Error(), "primary_failure") || !ok || cleanup.Code != "release_failure" {
		t.Fatalf("lost primary or cleanup failure: %v %#v", err, r)
	}
}

type delayedSelectorHost struct {
	mockHost
	searches int
	code     string
	op       string
}

func (h *delayedSelectorHost) Call(ctx context.Context, op native.Operation) (any, error) {
	if op.Op != h.op {
		return h.mockHost.Call(ctx, op)
	}
	h.searches++
	if h.searches == 1 {
		return nil, native.Fail(h.code, "test selector", nil)
	}
	return map[string]any{"name": "Ready"}, nil
}

func TestEventuallyRetriesOnlyTransientReadFailures(t *testing.T) {
	p := office.BlankPackage()
	b, _ := p.Bytes()
	file := filepath.Join(t.TempDir(), "test.docx")
	if err := os.WriteFile(file, b, 0600); err != nil {
		t.Fatal(err)
	}
	s := Suite{Schema: 1, Name: "selector wait", Steps: []Step{{Name: "ready", Operation: native.Operation{Op: "ui.find"}, EventuallyMS: 100, Assert: []Assertion{{Kind: "equals", Path: "/name", Expected: "Ready"}}}}}
	for _, tc := range []struct {
		op, code string
		retry    bool
	}{
		{"ui.find", "ui_selector_not_found", true},
		{"ui.find", "ui_selector_not_unique", false},
		{"get", "word_busy", true},
		{"get", "unknown_member", false},
		{"get", "word_automation_error", false},
	} {
		s.Steps[0].Operation.Op = tc.op
		h := &delayedSelectorHost{code: tc.code, op: tc.op}
		r, err := Run(context.Background(), file, s, h, true)
		if tc.retry {
			if err != nil || h.searches != 2 || r.Observations[0].Attempts != 2 {
				t.Fatal(r, err, h.searches)
			}
		} else if err == nil || h.searches != 1 {
			t.Fatal("retried non-transient failure", tc, err, h.searches)
		}
	}
}

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
	snapshot, snapshotErr := os.ReadFile(r.ArtifactSnapshot)
	if snapshotErr != nil || office.Hash(snapshot) != r.SHA256 || r.ArtifactSnapshot == file {
		t.Fatal("tested bytes not independently retained", snapshotErr)
	}
	progress, err := os.Open(filepath.Join(r.EvidenceDirectory, "progress.jsonl"))
	if err != nil {
		t.Fatal(err)
	}
	defer progress.Close()
	decoder := json.NewDecoder(progress)
	var started, finished map[string]any
	if err := decoder.Decode(&started); err != nil {
		t.Fatal(err)
	}
	if err := decoder.Decode(&finished); err != nil {
		t.Fatal(err)
	}
	if started["event"] != "step_started" || started["artifact_sha256"] != r.SHA256 || finished["event"] != "step_finished" || finished["observation"].(map[string]any)["passed"] != true {
		t.Fatal(started, finished)
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

func TestEvidenceRunsDoNotOverwrite(t *testing.T) {
	artifact := filepath.Join(t.TempDir(), "Example.dotm")
	first, err := createEvidenceDirectory(artifact, "0123456789abcdef")
	if err != nil {
		t.Fatal(err)
	}
	if err = os.WriteFile(filepath.Join(first, "output.xml"), []byte("first"), 0600); err != nil {
		t.Fatal(err)
	}
	second, err := createEvidenceDirectory(artifact, "0123456789abcdef")
	if err != nil {
		t.Fatal(err)
	}
	if first == second || filepath.Dir(first) != filepath.Dir(second) {
		t.Fatal(first, second)
	}
	if err = os.WriteFile(filepath.Join(second, "output.xml"), []byte("second"), 0600); err != nil {
		t.Fatal(err)
	}
	data, err := os.ReadFile(filepath.Join(first, "output.xml"))
	if err != nil || string(data) != "first" {
		t.Fatal("earlier evidence overwritten", string(data), err)
	}
}

func TestSuiteRejectsMalformedAssertions(t *testing.T) {
	for _, a := range []Assertion{
		{Kind: "typo"}, {Kind: "equals", Path: "result"}, {Kind: "equals", Path: "/bad~2"},
		{Kind: "matches", Expected: "["}, {Kind: "contains", Expected: 1}, {Kind: "near", Expected: 1, Tolerance: -1},
		{Kind: "greater_than", Expected: "5"}, {Kind: "absent"},
	} {
		s := Suite{Schema: 1, Name: "preflight", Steps: []Step{{Name: "editing operation", Operation: native.Operation{Op: "eval", Value: "ActiveDocument.Content.Text = 1"}, Assert: []Assertion{a}}}}
		if err := s.Validate(); err == nil {
			t.Fatalf("invalid assertion accepted: %+v", a)
		}
	}
	for _, path := range []string{"", "/literal~0", "/slash~1", "/literal~01"} {
		if err := (Assertion{Kind: "equals", Path: path}).validate(); err != nil {
			t.Fatal(path, err)
		}
	}
}
