//go:build windows && (amd64 || arm64)

package verify_test

import (
	"context"
	"encoding/binary"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"syscall"
	"testing"
	"time"

	"github.com/eliziff/WordUp/internal/example"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
)

func TestMain(m *testing.M) {
	if len(os.Args) > 1 && os.Args[1] == "__host" {
		if native.HostMain(os.Args[2:]) != nil {
			os.Exit(1)
		}
		return
	}
	os.Exit(m.Run())
}

func TestNativeFailureXML(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("FailureXML", root); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	b, err := w.Build("")
	if err != nil {
		t.Fatal(err)
	}
	s := verify.Suite{Schema: 1, Name: "Automatic partial XML", Steps: []verify.Step{
		{Name: "Open", Operation: native.Operation{Op: "open", File: "$artifact", As: "doc"}},
		{Name: "Fail after edit", Operation: native.Operation{Op: "eval", Value: "ActiveDocument.Content.Text = \"Retained partial edit\"\nErr.Raise 713, , \"Deliberate failure\""}, Assert: []verify.Assertion{{Path: "/result", Kind: "equals", Expected: true}}},
	}}
	r, err := verify.Run(context.Background(), b.Artifact, s, nil, true)
	if err == nil || r.Status != "failed" || len(r.Observations) != 2 {
		t.Fatalf("failure lost: %v", err)
	}
	failure, _ := json.Marshal(r.Error)
	if !strings.Contains(string(failure), "Deliberate failure") {
		t.Fatalf("primary error replaced: %s", failure)
	}
	x, ok := r.Observations[1].FailureXML.(map[string]any)
	if !ok {
		t.Fatalf("missing automatic XML: %#v", r.Observations[1])
	}
	file, _ := x["file"].(string)
	raw, err := os.ReadFile(file)
	if err != nil || !strings.Contains(string(raw), "Retained partial edit") || office.Hash(raw) != x["sha256"] {
		t.Fatalf("invalid partial XML: %v, %#v", err, x)
	}
	diagnosticJSON, _ := json.Marshal(r.Observations[1].Diagnostics)
	var diagnostics struct {
		Windows []map[string]any `json:"windows"`
	}
	if err := json.Unmarshal(diagnosticJSON, &diagnostics); err != nil || len(diagnostics.Windows) == 0 {
		t.Fatalf("missing visual failure evidence: %s", diagnosticJSON)
	}
	for _, window := range diagnostics.Windows {
		if window["tree"] != nil {
			t.Fatal("located scratch failure traversed unrelated UI trees")
		}
		path, _ := window["screenshot"].(string)
		if _, err := os.Stat(path); err != nil {
			t.Fatalf("missing failure screenshot: %v", err)
		}
	}
	t.Logf("located failure capture: %.3f ms", r.TimingMS["failure_diagnostics"])
	s.Steps[1].Operation.Value = "Dim tab As Table\nEvaluate = True"
	s.Steps[1].TimeoutMS = 5000
	r, err = verify.Run(context.Background(), b.Artifact, s, nil, true)
	failure, _ = json.Marshal(r.Error)
	if err == nil || !strings.Contains(string(failure), `"code":"vba_compile_error"`) || !strings.Contains(string(failure), "observed_dialogs") || !strings.Contains(string(failure), "Compile error") {
		t.Fatalf("execution error omitted compiler dialog: %s", failure)
	}
}

func TestNativeSecondaryBreakNoticeRecovery(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	root := filepath.Join(t.TempDir(), "studio")
	_, err := example.Studio(root)
	if err != nil {
		t.Fatal(err)
	}
	// Inject the fault only into this disposable test artifact. Starter users
	// should not inherit a broken helper just to exercise debugger recovery.
	source, err := os.ReadFile(filepath.Join(root, "vba", "Studio.bas"))
	if err != nil {
		t.Fatal(err)
	}
	source = append(source, []byte("\nPublic Sub DeliberateRibbonFailure()\n    Dim ribbon As IRibbonUI\n    ribbon.ActivateTab \"wwStudio\"\nEnd Sub\n")...)
	if err = project.Write(root, "vba/Studio.bas", source, ""); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	b, err := w.Build("")
	if err != nil {
		t.Fatal(err)
	}
	suite := example.WindowsSuite()
	for i, step := range suite.Steps {
		if step.Name == "Click the actual custom Ribbon tab" {
			// Debug exposes queued Ribbon callbacks' secondary break-mode notices.
			suite.Steps = suite.Steps[:i+1]
			suite.Steps[i].Operation = native.Operation{Op: "run", Macro: "Studio.DeliberateRibbonFailure"}
			break
		}
	}
	r, err := verify.Run(context.Background(), b.Artifact, suite, nil, true)
	if err == nil || r.Status != "failed" {
		t.Fatalf("deliberate failure was not retained: %v", err)
	}
	raw, _ := json.Marshal(r.Error)
	for _, expected := range []string{`"code":"vba_runtime_error"`, `"number":91`, `"break_mode_notices"`, `"reset":true`, `"frames":["WordUpStudio.Studio.DeliberateRibbonFailure"]`, `"source_file":"vba/Studio.bas"`, `"status":"completed"`} {
		if !strings.Contains(string(raw), expected) {
			t.Fatalf("missing %s: %s", expected, raw)
		}
	}
	last := r.Observations[len(r.Observations)-1]
	xml, ok := last.FailureXML.(map[string]any)
	if !ok {
		t.Fatal("missing native failure XML")
	}
	content, err := os.ReadFile(xml["file"].(string))
	if err != nil || office.Hash(content) != xml["sha256"] {
		t.Fatal("invalid retained XML", err)
	}
	t.Logf("secondary notice recovery %.3f ms", last.DurationMS)
}

func TestNativeRibbonSearchUsesRibbonHost(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	b, err := example.Studio(filepath.Join(t.TempDir(), "studio"))
	if err != nil {
		t.Fatal(err)
	}
	r, err := verify.Run(context.Background(), b.Artifact, example.WindowsSuite(), nil, true)
	if err != nil || r.Status != "passed" {
		t.Fatalf("native starter failed: %v", err)
	}
	for _, o := range r.Observations {
		if o.Name != "Click the real Ribbon button" {
			continue
		}
		result, ok := o.Result.(map[string]any)
		if !ok || fmt.Sprint(result["roots_searched"]) != "1" || fmt.Sprint(result["windows_searched"]) != "1" {
			t.Fatalf("Ribbon selector searched unrelated panes: %#v", o.Result)
		}
		t.Logf("Ribbon button %.3f ms, one native host", o.DurationMS)
		return
	}
	t.Fatal("suite did not exercise the Ribbon button")
}

func TestNativeMissingRibbonControlReportsSearch(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	b, err := example.Studio(filepath.Join(t.TempDir(), "studio"))
	if err != nil {
		t.Fatal(err)
	}
	suite := example.WindowsSuite()
	for i, step := range suite.Steps {
		if step.Name == "Click the real Ribbon button" {
			suite.Steps = suite.Steps[:i+1]
			suite.Steps[i].Operation = native.Operation{Op: "ui.find", Target: "formDocument", Named: map[string]any{"name": "Deliberately nonexistent command", "scope": "ribbon", "role": 43}}
			break
		}
	}
	r, err := verify.Run(context.Background(), b.Artifact, suite, nil, true)
	raw, _ := json.Marshal(r.Error)
	if err == nil || !strings.Contains(string(raw), `"ui_selector_not_found"`) || !strings.Contains(string(raw), `"root_searches":[{"duration_ms":`) || strings.Contains(string(raw), `"nodes_visited":0`) || !strings.Contains(string(raw), `"nodes_visited":`) {
		t.Fatalf("missing actual scoped search evidence: %v %s", err, raw)
	}
	t.Logf("missing-control search evidence: %s", raw)
}

func TestNativeDialogPollingCost(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	ctx := context.Background()
	h, err := native.Start(ctx, native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	if _, err := h.Call(ctx, native.Operation{Op: "new", As: "document"}); err != nil {
		t.Fatal(err)
	}
	full, filtered := []float64{}, []float64{}
	for i := 0; i < 11; i++ {
		started := time.Now()
		windows, err := h.Call(ctx, native.Operation{Op: "ui.windows"})
		fullMS := float64(time.Since(started).Nanoseconds()) / 1e6
		if err != nil || len(windows.([]any)) == 0 {
			t.Fatalf("missing owned window inventory: %v %v", windows, err)
		}
		started = time.Now()
		_, err = h.Call(ctx, native.Operation{Op: "ui.diagnostics", Named: map[string]any{"trees": false, "window_class": "#32770"}})
		filteredMS := float64(time.Since(started).Nanoseconds()) / 1e6
		if fault, ok := err.(*native.Fault); !ok || fault.Code != "diagnostic_window_not_found" {
			t.Fatalf("unexpected dialog in idle Word: %v", err)
		}
		if i > 0 {
			full, filtered = append(full, fullMS), append(filtered, filteredMS)
		}
	}
	sort.Float64s(full)
	sort.Float64s(filtered)
	t.Logf("same-session idle polling, 10 samples excluding warmup: full inventory median %.3f ms; dialog observer median %.3f ms", (full[4]+full[5])/2, (filtered[4]+filtered[5])/2)
}

func TestNativeTOCInspectionDetectsStalePages(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	root := filepath.Join(t.TempDir(), "toc")
	if _, err := project.New("TOCProof", root); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	b, err := w.Build("")
	if err != nil {
		t.Fatal(err)
	}
	suite := verify.Suite{Schema: 1, Name: "Live TOC pagination diagnostics", Steps: []verify.Step{
		{Name: "Open", Operation: native.Operation{Op: "open", File: "$artifact", As: "doc"}},
		{Name: "Create live TOC", Operation: native.Operation{Op: "eval", Value: `Dim d As Document
Set d=ActiveDocument
d.Content.Text=vbCr & "First heading" & vbCr & "Body" & vbCr & "Second heading" & vbCr & "Body" & vbCr
d.Paragraphs(2).Style=wdStyleHeading1
d.Paragraphs(4).Style=wdStyleHeading1
d.TablesOfContents.Add Range:=d.Range(0,0), UseHeadingStyles:=True, UpperHeadingLevel:=1, LowerHeadingLevel:=1, UseHyperlinks:=True
d.Repaginate
d.TablesOfContents(1).UpdatePageNumbers
Evaluate=True`}},
		{Name: "Current pages", Operation: native.Operation{Op: "toc.inspect", Target: "doc", Named: map[string]any{"repaginate": true}}, Assert: []verify.Assertion{{Path: "/checked", Kind: "equals", Expected: 2}, {Path: "/mismatches", Kind: "equals", Expected: 0}, {Path: "/complete", Kind: "equals", Expected: true}}},
		{Name: "Move destination without updating TOC", Operation: native.Operation{Op: "eval", Value: `Dim d As Document, destination As Range
Set d=ActiveDocument
Set destination=d.Bookmarks(d.TablesOfContents(1).Range.Hyperlinks(1).SubAddress).Range
destination.Collapse wdCollapseStart
destination.InsertBreak wdPageBreak
Evaluate=d.TablesOfContents(1).Range.Text`}},
		{Name: "Detect stale pages", Operation: native.Operation{Op: "toc.inspect", Target: "doc", Named: map[string]any{"repaginate": true}}, Assert: []verify.Assertion{{Path: "/mismatches", Kind: "greater_than", Expected: 0}, {Path: "/fields_updated", Kind: "equals", Expected: false}}},
		{Name: "Observe unchanged TOC text", Operation: native.Operation{Op: "eval", Value: `Evaluate=ActiveDocument.TablesOfContents(1).Range.Text`}},
		{Name: "Update pages", Operation: native.Operation{Op: "eval", Value: `ActiveDocument.TablesOfContents(1).UpdatePageNumbers
Evaluate=True`}},
		{Name: "Check repaired pages", Operation: native.Operation{Op: "toc.inspect", Target: "doc"}, Assert: []verify.Assertion{{Path: "/mismatches", Kind: "equals", Expected: 0}, {Path: "/complete", Kind: "equals", Expected: true}}},
		{Name: "Break one destination", Operation: native.Operation{Op: "eval", Value: `ActiveDocument.Bookmarks(ActiveDocument.TablesOfContents(1).Range.Hyperlinks(1).SubAddress).Delete
Evaluate=True`}},
		{Name: "Missing destination is incomplete", Operation: native.Operation{Op: "toc.inspect", Target: "doc"}, Assert: []verify.Assertion{{Path: "/unresolved", Kind: "greater_than", Expected: 0}, {Path: "/complete", Kind: "equals", Expected: false}}},
		{Name: "Remove broken hyperlink but retain entry", Operation: native.Operation{Op: "eval", Value: `ActiveDocument.TablesOfContents(1).Range.Hyperlinks(1).Delete
Evaluate=True`}},
		{Name: "Partially linked TOC is incomplete", Operation: native.Operation{Op: "toc.inspect", Target: "doc"}, Assert: []verify.Assertion{{Path: "/checked", Kind: "equals", Expected: 1}, {Path: "/complete", Kind: "equals", Expected: false}, {Path: "/unresolved", Kind: "greater_than", Expected: 0}}},
	}}
	r, err := verify.Run(context.Background(), b.Artifact, suite, nil, true)
	if err != nil {
		t.Fatal(err)
	}
	before, _ := json.Marshal(r.Observations[3].Result)
	after, _ := json.Marshal(r.Observations[5].Result)
	var a, z map[string]any
	json.Unmarshal(before, &a)
	json.Unmarshal(after, &z)
	if a["result"] != z["result"] {
		t.Fatal("inspection rewrote stale TOC results")
	}
	t.Logf("TOC inspection %.3f ms", r.Observations[4].DurationMS)
}

func TestNativeScratchLocations(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	ctx := context.Background()
	h, err := native.Start(ctx, native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	if _, err = h.Call(ctx, native.Operation{Op: "new", As: "scratch"}); err != nil {
		t.Fatal(err)
	}
	body := "Dim x As String\nx = \"a\" & _\n    \"b\"\nRem comment\n#If VBA7 Then\nIf x = \"ab\" Then\nErr.Raise 5, , \"located failure\"\nElse\nEvaluate = False\nEnd If\n#End If"
	check := func(op native.Operation) {
		t.Helper()
		_, err := h.Call(ctx, op)
		f, ok := err.(*native.Fault)
		if !ok || f.Code != "scratch_vba_failed" {
			t.Fatalf("missing native runtime failure: %v", err)
		}
		b, _ := json.Marshal(f.Details)
		if !strings.Contains(string(b), `"body_line":7`) || !strings.Contains(string(b), `"number":5`) || !strings.Contains(string(b), `"body_line_text":"Err.Raise 5`) {
			t.Fatalf("missing original source location: %s", b)
		}
	}
	check(native.Operation{Op: "eval", Value: body})
	selected, selectErr := h.Call(ctx, native.Operation{Op: "eval", Value: "Select Case 2\nCase 1, 2\nEvaluate = 42\nCase Else\nEvaluate = 0\nEnd Select"})
	if selectErr != nil || fmt.Sprint(selected.(map[string]any)["result"]) != "42" {
		t.Fatalf("instrumented Select Case failed: %v %v", selected, selectErr)
	}
	xmlPath := filepath.Join(t.TempDir(), "failure.xml")
	if _, err = h.Call(ctx, native.Operation{Op: "xml.snapshot", File: xmlPath, Named: map[string]any{"failure_capture": true}}); err != nil {
		t.Fatalf("capture after runtime failure: %v", err)
	}
	if raw, err := os.ReadFile(xmlPath); err != nil || !strings.Contains(string(raw), "word/document.xml") {
		t.Fatalf("missing live document XML: %v", err)
	}
	if _, err = h.Call(ctx, native.Operation{Op: "eval", Value: "Application.UndoRecord.StartCustomRecord \"Preserve pending undo\"\nEvaluate = True"}); err != nil {
		t.Fatal(err)
	}
	_, err = h.Call(ctx, native.Operation{Op: "xml.snapshot", Named: map[string]any{"failure_capture": true}})
	if f, ok := err.(*native.Fault); !ok || f.Code != "xml_capture_undo_active" {
		t.Fatalf("exported inside undo record: %v", err)
	}
	if _, err = h.Call(ctx, native.Operation{Op: "eval", Value: "Application.UndoRecord.EndCustomRecord\nEvaluate = True"}); err != nil {
		t.Fatal(err)
	}
	prepared, err := h.Call(ctx, native.Operation{Op: "eval.prepare", Steps: []native.Operation{{Op: "eval", As: "failure", Value: body}}})
	if err != nil {
		t.Fatal(err)
	}
	program := prepared.(map[string]any)["program"].(string)
	// Modal workflows can load a direct scratch add-in while a prepared
	// program remains installed. Its retained COM AddIn may become stale.
	check(native.Operation{Op: "eval", Value: body})
	check(native.Operation{Op: "eval.execute", Target: program, Member: "failure"})
	if _, err = h.Call(ctx, native.Operation{Op: "eval.release", Target: program}); err != nil {
		t.Fatal(err)
	}
	remaining, err := h.Call(ctx, native.Operation{Op: "eval", Value: "Dim a As AddIn\nEvaluate = False\nFor Each a In Application.AddIns\nIf a.Name = \"" + program + ".dotm\" Then Evaluate = a.Installed\nNext"})
	if err != nil || remaining.(map[string]any)["result"] != false {
		t.Fatalf("released scratch program remains installed: %v, %v", remaining, err)
	}
	_, err = h.Call(ctx, native.Operation{Op: "ui.diagnostics", Named: map[string]any{"window": "WordUp deliberately missing window"}})
	if f, ok := err.(*native.Fault); !ok || f.Code != "diagnostic_window_not_found" {
		t.Fatalf("missing diagnostic target silently accepted: %v", err)
	}
}

func TestNativeCompilerDiagnostics(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	t.Run("syntax", func(t *testing.T) {
		nativeCompilerDiagnosticCase(t, "Public Sub DeliberatelyBroken(\nEnd Sub\n", 2, "Public Sub DeliberatelyBroken(", "Expected: identifier")
	})
	t.Run("undefined_type", func(t *testing.T) {
		nativeCompilerDiagnosticCase(t, "Public Sub BrokenType()\nDim value As WordUpMissingType\nEnd Sub\n", 3, "Dim value As WordUpMissingType", "User-defined type not defined")
	})
}

func nativeCompilerDiagnosticCase(t *testing.T, body string, line int, statement, message string) {
	t.Helper()
	root := filepath.Join(t.TempDir(), "workspace")
	_, err := project.New("BrokenCompile", root)
	if err != nil {
		t.Fatal(err)
	}
	if err = project.Write(root, "vba/Broken.bas", []byte("Attribute VB_Name = \"Broken\"\nOption Explicit\n"+body), ""); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	build, err := w.Build("")
	if err != nil {
		t.Fatal(err)
	}
	path := build.Artifact
	s := verify.Suite{Schema: 1, Name: "Compiler diagnostic proof", Steps: []verify.Step{
		{Name: "Open", Operation: native.Operation{Op: "open", File: "$artifact", As: "doc"}},
		{Name: "Compile", Operation: native.Operation{Op: "compile", Target: "doc", Member: "$project"}, TimeoutMS: 2000, Assert: []verify.Assertion{{Path: "/vba_compiled", Kind: "equals", Expected: true}}},
	}}
	r, err := verify.Run(context.Background(), path, s, nil, true)
	if directory := os.Getenv("WORDUP_EVIDENCE_DIR"); directory != "" {
		if writeErr := os.WriteFile(filepath.Join(directory, "compiler-"+filepath.Base(t.Name())+".json"), project.JSON(r), 0600); writeErr != nil {
			t.Fatal(writeErr)
		}
	}
	if err == nil || r.VBACompiled || len(r.Observations) != 2 {
		t.Fatalf("invalid compiler evidence: steps=%d compiled=%v error=%v", len(r.Observations), r.VBACompiled, err)
	}
	failure, _ := json.Marshal(r.Error)
	if !strings.Contains(string(failure), `"module":"Broken"`) || !strings.Contains(string(failure), fmt.Sprintf(`"line":%d`, line)) || !strings.Contains(string(failure), statement) {
		t.Fatalf("missing native compiler selection: %s", failure)
	}
	if !strings.Contains(string(failure), message) {
		t.Fatalf("compiler message missing from report-level error: %s", failure)
	}
	diagnostics, _ := json.Marshal(r.Observations[1].Diagnostics)
	if !strings.Contains(strings.ToLower(string(diagnostics)), "compile error") {
		t.Fatalf("compiler failure lacks compiler dialog text: %v", r.Error)
	}
}

func TestNativeCompletion(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	ctx := context.Background()
	h, err := native.Start(ctx, native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	w := h.(interface {
		WaitTask(context.Context, string) (any, error)
	})
	if _, err = h.Call(ctx, native.Operation{Op: "new", As: "scratch"}); err != nil {
		t.Fatal(err)
	}
	begin := func(body string, ms int) string {
		t.Helper()
		v, e := h.Call(ctx, native.Operation{Op: "begin", TimeoutMS: ms, Steps: []native.Operation{{Op: "eval", Value: body}}})
		if e != nil {
			t.Fatal(e)
		}
		return v.(map[string]any)["task"].(string)
	}
	for _, body := range []string{"Evaluate = 6 * 7", "100 Err.Raise 5"} {
		key := begin(body, 5000)
		v, e := w.WaitTask(ctx, key)
		if e != nil {
			t.Fatal(e)
		}
		b, _ := json.Marshal(v)
		if body == "Evaluate = 6 * 7" && !strings.Contains(string(b), `"result":42`) {
			t.Fatalf("missing result: %s", b)
		}
		if body != "Evaluate = 6 * 7" && (!strings.Contains(string(b), `"number":5`) || !strings.Contains(string(b), `"line":100`)) {
			t.Fatalf("missing runtime fault: %s", b)
		}
		if _, e = h.Call(ctx, native.Operation{Op: "forget", Value: key}); e != nil {
			t.Fatal(e)
		}
	}
	key := begin("MsgBox \"WordUp modal lane test\"\nEvaluate = 42", 10000)
	started := time.Now()
	_, err = h.Call(ctx, native.Operation{Op: "get", Target: "app", Member: "ScreenUpdating"})
	if f, ok := err.(*native.Fault); !ok || f.Code != "word_task_pending" || time.Since(started) > time.Second {
		t.Fatalf("blocked object-model call did not fail promptly: %v", err)
	}
	if _, err = h.Call(ctx, native.Operation{Op: "ui.invoke", Named: map[string]any{"window": "Microsoft Word", "name": "OK", "role": 43, "wait_ms": 5000}}); err != nil {
		t.Fatal(err)
	}
	if _, err = w.WaitTask(ctx, key); err != nil {
		t.Fatal(err)
	}
	if _, err = h.Call(ctx, native.Operation{Op: "get", Target: "app", Member: "ScreenUpdating"}); err != nil {
		t.Fatalf("completion did not reopen object-model lane: %v", err)
	}
	key = begin("Do\nLoop", 1000)
	short, cancel := context.WithTimeout(ctx, 20*time.Millisecond)
	_, err = w.WaitTask(short, key)
	cancel()
	if err != context.DeadlineExceeded || h.Info()["closed"] == true {
		t.Fatalf("caller deadline killed diagnostic lane: %v", err)
	}
	deadline, cancel := context.WithTimeout(ctx, 5*time.Second)
	defer cancel()
	_, err = w.WaitTask(deadline, key)
	if err == nil || !strings.Contains(err.Error(), "wall-clock limit") {
		t.Fatalf("independent watchdog: %v", err)
	}
}

// Opt-in: these checks run real VBA in owned Microsoft Word processes.
func TestNativeFeedback(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1 to authorize native Word tests")
	}
	root := t.TempDir()
	if output := os.Getenv("WORDUP_EVIDENCE_DIR"); output != "" {
		var e error
		root, e = os.MkdirTemp(output, "feedback-")
		if e != nil {
			t.Fatal(e)
		}
	}
	write := func(name string, v any) {
		b, _ := json.MarshalIndent(v, "", "  ")
		if e := os.WriteFile(filepath.Join(root, name), b, 0600); e != nil {
			t.Fatal(e)
		}
	}
	build, e := example.Studio(filepath.Join(root, "studio"))
	if e != nil {
		t.Fatal(e)
	}
	ctx := context.Background()
	r, e := verify.Run(ctx, build.Artifact, example.WindowsSuite(), nil, true)
	write("acceptance.json", r)
	if e != nil {
		t.Fatal(e)
	}
	before, e := os.ReadFile(build.Artifact)
	if e != nil {
		t.Fatal(e)
	}
	h, e := native.Start(ctx, native.Options{Execute: true})
	if e != nil {
		t.Fatal(e)
	}
	defer h.Close()
	call := func(op native.Operation) (any, error) {
		c, cancel := context.WithTimeout(ctx, 10*time.Second)
		defer cancel()
		return h.Call(c, op)
	}
	if _, e = call(native.Operation{Op: "new", As: "sentinel"}); e != nil {
		t.Fatal(e)
	}
	latencies := []float64{}
	for i := 0; i < 5; i++ {
		start := time.Now()
		v, e := call(native.Operation{Op: "eval", Value: "Evaluate = 21 * 2"})
		if e != nil || v.(map[string]any)["result"] != float64(42) {
			t.Fatalf("warm eval: %v %v", v, e)
		}
		latencies = append(latencies, float64(time.Since(start).Microseconds())/1000)
	}
	write("warm-latencies-ms.json", latencies)
	_, e = call(native.Operation{Op: "eval", Value: `100 Err.Raise 5, "WordUp fixture", "Expected runtime failure"`})
	f, ok := e.(*native.Fault)
	if !ok || f.Code != "scratch_vba_failed" {
		t.Fatalf("runtime error was not observable: %v", e)
	}
	write("runtime-error.json", f)
	details := f.Details.(map[string]any)
	if details["line"] != float64(100) || details["number"] != float64(5) {
		t.Fatalf("lost VBA diagnostic: %v", details)
	}
	w, e := project.Open(filepath.Join(root, "studio"))
	if e != nil {
		t.Fatal(e)
	}
	path := filepath.Join(root, "studio", "vba", "Studio.bas")
	source, e := os.ReadFile(path)
	if e != nil {
		t.Fatal(e)
	}
	iterations := []map[string]any{}
	for i := 0; i < 5; i++ {
		start := time.Now()
		revision := append(append([]byte(nil), source...), []byte(fmt.Sprintf("\n' Warm revision %d\n", i))...)
		if e = os.WriteFile(path, revision, 0600); e != nil {
			t.Fatal(e)
		}
		candidate, e := w.Build(filepath.Join(root, "warm.dotm"))
		if e != nil {
			t.Fatal(e)
		}
		compileStart := time.Now()
		if _, e = call(native.Operation{Op: "open", File: candidate.Artifact, As: "warm"}); e != nil {
			t.Fatal(e)
		}
		v, e := call(native.Operation{Op: "compile", Target: "warm", Member: "WordUpStudio"})
		if e != nil || v.(map[string]any)["vba_compiled"] != true {
			t.Fatalf("warm compile: %v %v", v, e)
		}
		v, e = call(native.Operation{Op: "run", Macro: "Studio.NativeAcceptance"})
		if e != nil || v != "WORDUP_NATIVE_OK" {
			t.Fatalf("warm VBA: %v %v", v, e)
		}
		if _, e = call(native.Operation{Op: "unload", Target: "warm", File: candidate.Artifact}); e != nil {
			t.Fatal(e)
		}
		iterations = append(iterations, map[string]any{"iteration": i + 1, "sha256": candidate.SHA256, "build_ms": candidate.DurationMS, "reload_compile_test_unload_ms": float64(time.Since(compileStart).Microseconds()) / 1000, "total_ms": float64(time.Since(start).Microseconds()) / 1000})
	}
	write("warm-edit-cycles.json", iterations)
	if e = os.WriteFile(path, []byte(strings.Replace(string(source), "Option Explicit", "Option Explicit\nPublic Sub DeliberatelyBroken(\n", 1)), 0600); e != nil {
		t.Fatal(e)
	}
	bad, e := w.Build(filepath.Join(root, "broken.dotm"))
	if e != nil {
		t.Fatal(e)
	}
	suite := verify.Suite{Schema: 1, Name: "Deliberate compiler failure", Steps: []verify.Step{
		{Name: "Open broken source", Operation: native.Operation{Op: "open", File: "$artifact", As: "doc"}},
		{Name: "Reject syntax error", Operation: native.Operation{Op: "compile", Target: "doc", Member: "$project"}, TimeoutMS: 3000, Assert: []verify.Assertion{{Path: "/vba_compiled", Kind: "equals", Expected: true}}},
	}}
	failed, e := verify.Run(ctx, bad.Artifact, suite, nil, true)
	write("compiler-error.json", failed)
	if e == nil || failed.Status != "failed" || failed.VBACompiled {
		t.Fatalf("syntax error falsely passed: %v", failed)
	}
	if len(failed.Observations) != 2 || failed.Observations[1].Diagnostics == nil {
		t.Fatal("missing automatic failure diagnostics")
	}
	// The independent Word process must remain usable after the failing host is
	// terminated. No attach-to-user-Word or global process termination is used.
	if _, e = call(native.Operation{Op: "get", Target: "sentinel", Member: "Name"}); e != nil {
		t.Fatalf("other Word instance affected: %v", e)
	}
	runaway, e := native.Start(ctx, native.Options{Execute: true})
	if e != nil {
		t.Fatal(e)
	}
	defer runaway.Close()
	start := time.Now()
	_, e = runaway.Call(ctx, native.Operation{Op: "begin", As: "runaway", TimeoutMS: 1000, Steps: []native.Operation{{Op: "eval", Value: "Do\nLoop"}}})
	if e != nil {
		t.Fatal(e)
	}
	// No polling: the parent watchdog must enforce the limit independently.
	for time.Since(start) < 6*time.Second && runaway.Info()["closed"] != true {
		time.Sleep(50 * time.Millisecond)
	}
	if runaway.Info()["closed"] != true {
		t.Fatal("unpolled infinite loop outlived its watchdog")
	}
	if e = runaway.Close(); e != nil {
		t.Fatal(e)
	}
	_, e = runaway.Call(ctx, native.Operation{Op: "poll", Value: "runaway"})
	f, ok = e.(*native.Fault)
	if !ok || f.Code != "macro_deadline" {
		t.Fatalf("lost runaway termination reason: %v", e)
	}
	write("runaway.json", map[string]any{"elapsed_ms": float64(time.Since(start).Microseconds()) / 1000, "fault": f, "limits": runaway.Info()["limits"]})
	if _, e = call(native.Operation{Op: "get", Target: "sentinel", Member: "Name"}); e != nil {
		t.Fatalf("runaway affected other Word: %v", e)
	}
	after, e := os.ReadFile(build.Artifact)
	if e != nil || office.Hash(before) != office.Hash(after) {
		t.Fatal("original artifact changed")
	}
	write("isolation.json", map[string]any{"other_owned_word_survived": true, "original_sha256": office.Hash(after), "original_unchanged": true})
	t.Log("native evidence:", root)
}

// Run beside the bundled OfficeTools helper. All UI belongs to hidden Word.
func TestNativeRibbonPatterns(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	ctx := context.Background()
	h, err := native.Start(ctx, native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	if _, err = h.Call(ctx, native.Operation{Op: "new", As: "document"}); err != nil {
		t.Fatal(err)
	}
	call := func(action, name, controlType string, role any) (any, error) {
		named := map[string]any{"scope": "ribbon", "name": name, "control_type": controlType}
		if role != nil {
			named["role"] = role
		}
		return h.Call(ctx, native.Operation{Op: "ui.patterns", Target: "document", Member: action, Named: named})
	}
	for _, name := range []string{"Font Size", "WordUp missing control 78219"} {
		if _, err = call("default_action", name, "", nil); err == nil || !strings.Contains(err.Error(), "Expected exactly one control") {
			t.Fatalf("ambiguous/missing selector did not reject before action: %v", err)
		}
	}
	for _, role := range []any{nil, float64(46)} {
		for _, state := range []string{"Expanded", "Collapsed"} {
			value, err := call("default_action", "Font Size", "ComboBox", role)
			if err != nil {
				t.Fatal(err)
			}
			result := value.(map[string]any)
			if result["expand_state"] != state {
				t.Fatalf("role %v: %v", role, result)
			}
		}
	}
	for _, action := range []string{"expand", "expand", "collapse", "collapse"} {
		value, err := call(action, "Font Size", "ComboBox", nil)
		if err != nil {
			t.Fatal(err)
		}
		wanted := "Collapsed"
		if action == "expand" {
			wanted = "Expanded"
		}
		result := value.(map[string]any)
		if result["expand_state"] != wanted {
			t.Fatalf("%s: %v", action, result)
		}
		t.Logf("%s: action=%v ms total=%v ms", action, result["action_ms"], result["duration_ms"])
	}
}

func nativeCallbackArtifact(t *testing.T) string {
	t.Helper()
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("CallbackProof", root); err != nil {
		t.Fatal(err)
	}
	files := map[string]string{
		"forms/FailureForm.json": `{"name":"FailureForm","mode":"replace","properties":{"Caption":"Callback diagnostic proof","Width":260,"Height":110},"controls":[{"name":"cmdFail","type":"CommandButton","properties":{"Caption":"Run deliberate failure","Left":20,"Top":20,"Width":210,"Height":30}}]}`,
		"vba/FailureForm.vba": `Attribute VB_Name = "FailureForm"
Option Explicit
Private Sub cmdFail_Click()
 ActiveDocument.Content.Text = "Retained callback edit"
 Err.Raise 713, , "Deliberate callback failure"
End Sub
`,
		"vba/Launch.bas": `Attribute VB_Name = "Launch"
Option Explicit
Public Sub ShowForm()
 FailureForm.Show vbModeless
End Sub
Public Sub FailDirect()
 Dim counter As Long: counter = 42: ActiveDocument.Content.Text = "Retained callback edit"
 Err.Raise 713, , "Deliberate direct failure"
End Sub
Public Sub NestedRun()
 Recurse 32
End Sub
Private Sub Recurse(ByVal depth As Long)
 If depth > 0 Then
  Recurse depth - 1
 Else
  FailDirect
 End If
End Sub
`,
	}
	for name, source := range files {
		if err := os.WriteFile(filepath.Join(root, name), []byte(source), 0600); err != nil {
			t.Fatal(err)
		}
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	built, err := w.Build("")
	if err != nil {
		t.Fatal(err)
	}
	return built.Artifact
}

func TestNativeUserFormRuntimeFailure(t *testing.T) { nativeRuntimeFailure(t, false) }

func TestNativeDirectRuntimeFailure(t *testing.T) { nativeRuntimeFailure(t, true) }

func nativeRuntimeFailure(t *testing.T, direct bool) {
	artifact := nativeCallbackArtifact(t)
	suite := verify.Suite{Schema: 1, Name: "Real modeless callback failure", Steps: []verify.Step{
		{Name: "Open", Operation: native.Operation{Op: "open", File: "$artifact", As: "template"}},
		{Name: "Compile", Operation: native.Operation{Op: "compile", Target: "template", Member: "$project"}, Assert: []verify.Assertion{{Kind: "equals", Path: "/vba_compiled", Expected: true}}},
		{Name: "Show modeless form", Operation: native.Operation{Op: "run", Macro: "Launch.ShowForm"}},
		{Name: "Invoke failing callback", Operation: native.Operation{Op: "ui.invoke", Named: map[string]any{"window": "Callback diagnostic proof", "name": "Run deliberate failure", "role": 43, "wait_ms": 3000}}},
	}}
	module, line := "FailureForm", "4"
	if direct {
		module, line = "Launch", "7"
		suite.Steps = suite.Steps[:2]
		suite.Steps = append(suite.Steps, verify.Step{Name: "Invoke direct failure", Operation: native.Operation{Op: "run", Macro: "Launch.FailDirect"}})
	}
	report, err := verify.Run(context.Background(), artifact, suite, nil, true)
	if err == nil || len(report.Observations) != len(suite.Steps) {
		t.Fatalf("callback failure not observed: %v", err)
	}
	failure, _ := json.Marshal(report.Error)
	if !strings.Contains(string(failure), `"code":"vba_runtime_error"`) || !strings.Contains(string(failure), "Deliberate") {
		t.Fatalf("lost callback diagnosis: %s", failure)
	}
	if !strings.Contains(string(failure), `"runtime_location"`) || !strings.Contains(string(failure), `"reset":true`) || !strings.Contains(string(failure), `"module":"`+module+`"`) || !strings.Contains(string(failure), `"line":`+line) {
		t.Fatalf("missing automatic paused callback location/reset: %s", failure)
	}
	if !strings.Contains(string(failure), `"runtime_stack"`) || strings.Contains(string(failure), `"stack_error"`) || !strings.Contains(string(failure), `"truncated":false`) {
		t.Fatalf("missing automatic native stack: %s", failure)
	}
	extension := ".vba"
	if direct {
		extension = ".bas"
	}
	if !strings.Contains(string(failure), `"source_file":"vba/`+module+extension+`"`) || !strings.Contains(string(failure), `"artifact_module_source_sha256"`) {
		t.Fatalf("runtime location did not map to tested artifact source: %s", failure)
	}
	t.Logf("callback diagnosis %.3f ms; failure evidence %.3f ms", report.Observations[len(suite.Steps)-1].DurationMS, report.TimingMS["failure_diagnostics"])
	xml, ok := report.Observations[len(suite.Steps)-1].FailureXML.(map[string]any)
	if !ok {
		t.Fatal("missing callback XML")
	}
	raw, err := os.ReadFile(xml["file"].(string))
	if err != nil || office.Hash(raw) != xml["sha256"] || !strings.Contains(string(raw), "Retained callback edit") {
		t.Fatal("missing exact partial edit", err)
	}
}

func TestNativePausedCallbackInspection(t *testing.T) {
	t.Run("UserForm", func(t *testing.T) { nativePausedInspection(t, false) })
	t.Run("DirectRun", func(t *testing.T) { nativePausedInspection(t, true) })
}

func nativePausedInspection(t *testing.T, direct bool) {
	artifact := nativeCallbackArtifact(t)
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	h, err := native.Start(ctx, native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	call := func(op native.Operation) any {
		t.Helper()
		value, err := h.Call(ctx, op)
		if err != nil {
			t.Fatal(op.Op, op.Member, err)
		}
		return value
	}
	call(native.Operation{Op: "open", File: artifact, As: "template"})
	module := "FailureForm"
	if direct {
		module = "Launch"
		call(native.Operation{Op: "begin", As: "callback", Steps: []native.Operation{{Op: "run", Macro: "Launch.NestedRun"}}})
	} else {
		call(native.Operation{Op: "run", Macro: "Launch.ShowForm"})
		call(native.Operation{Op: "begin", As: "callback", Steps: []native.Operation{{Op: "ui.invoke", Named: map[string]any{"window": "Callback diagnostic proof", "name": "Run deliberate failure", "role": 43}}}})
	}
	call(native.Operation{Op: "ui.find", Named: map[string]any{"window": "Microsoft Visual Basic", "name": "Debug", "role": 43, "wait_ms": 3000}})
	call(native.Operation{Op: "ui.invoke", Named: map[string]any{"window": "Microsoft Visual Basic", "name": "Debug", "role": 43}})
	state := call(native.Operation{Op: "ui.vba.inspect"}).(map[string]any)
	t.Logf("paused state: %#v", state)
	if fmt.Sprint(state["project_mode"]) != "1" || state["module"] != module || !strings.Contains(fmt.Sprint(state["line_text"]), "Err.Raise 713") {
		t.Fatalf("wrong paused location: %v", state)
	}
	stackStarted := time.Now()
	stack := call(native.Operation{Op: "ui.vba.stack"}).(map[string]any)
	t.Logf("stack capture: %v; %v", time.Since(stackStarted), stack)
	frames := fmt.Sprint(stack["frames"])
	if stack["truncated"] != false || stack["close_error"] != nil || (!direct && !strings.Contains(frames, "cmdFail_Click")) || (direct && (!strings.Contains(frames, "Launch.FailDirect") || !strings.Contains(frames, "Launch.NestedRun") || strings.Count(frames, "Launch.Recurse") != 33)) {
		t.Fatalf("missing paused stack or failed cleanup: %v", stack)
	}

	started := time.Now()
	reset := call(native.Operation{Op: "ui.vba.reset"}).(map[string]any)
	if reset["reset"] != true || reset["line_text"] != state["line_text"] {
		t.Fatalf("reset omitted original location: %v", reset)
	}
	t.Logf("inspect and native reset: %v", time.Since(started))
	state = call(native.Operation{Op: "ui.vba.inspect"}).(map[string]any)
	if fmt.Sprint(state["project_mode"]) != "2" {
		t.Fatalf("reset left project paused: %v", state)
	}
	call(native.Operation{Op: "get", Target: "app", Member: "ActiveDocument", As: "active"})
	call(native.Operation{Op: "get", Target: "active", Member: "Content", As: "content"})
	value := call(native.Operation{Op: "get", Target: "content", Member: "Text"})
	if !strings.Contains(fmt.Sprint(value), "Retained callback edit") {
		t.Fatalf("reset lost partial document edit: %v", value)
	}

}

func TestNativeOwnedWordExitDiagnostics(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
	defer cancel()
	h, err := native.Start(ctx, native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	info := h.Info()
	if info["word_exit_monitor"] != true {
		t.Fatal("owned process monitor unavailable", info)
	}
	pid := uint32(info["pid"].(float64))
	if pid == 0 {
		t.Fatal("missing owned PID")
	}
	process, err := syscall.OpenProcess(syscall.PROCESS_TERMINATE, false, pid)
	if err != nil {
		t.Fatal(err)
	}
	defer syscall.CloseHandle(process)
	if _, err = h.Call(ctx, native.Operation{Op: "new", As: "scratch"}); err != nil {
		t.Fatal(err)
	}
	started, err := h.Call(ctx, native.Operation{Op: "begin", TimeoutMS: 10000, Steps: []native.Operation{{Op: "eval", Value: "Do\nLoop"}}})
	if err != nil {
		t.Fatal(err)
	}
	const exitCode = uint32(0xE0000713)
	if err = syscall.TerminateProcess(process, exitCode); err != nil {
		t.Fatal(err)
	}
	waiter := h.(interface {
		WaitTask(context.Context, string) (any, error)
	})
	_, err = waiter.WaitTask(ctx, started.(map[string]any)["task"].(string))
	failure, _ := json.Marshal(verify.ErrorValue(err))
	if !strings.Contains(string(failure), `"code":"word_process_exited"`) || !strings.Contains(string(failure), fmt.Sprintf(`"exit_code":%d`, exitCode)) {
		t.Fatalf("process exit lost: %s", failure)
	}
	if h.Info()["word_alive"] != false || h.Info()["word_exit_code"] != exitCode {
		t.Fatal("process state not retained", h.Info())
	}
}

func TestNativeOwnedWorkerExitDiagnostics(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
	defer cancel()
	h, err := native.Start(ctx, native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	pid := h.Info()["worker_pid"].(uint32)
	if pid == 0 {
		t.Fatal("missing owned worker PID")
	}
	process, err := syscall.OpenProcess(syscall.PROCESS_TERMINATE, false, pid)
	if err != nil {
		t.Fatal(err)
	}
	defer syscall.CloseHandle(process)
	started, err := h.Call(ctx, native.Operation{Op: "begin", TimeoutMS: 10000, Steps: []native.Operation{{Op: "eval", Value: "Do\nLoop"}}})
	if err != nil {
		t.Fatal(err)
	}
	const exitCode = uint32(0xE0000714)
	if err = syscall.TerminateProcess(process, exitCode); err != nil {
		t.Fatal(err)
	}
	waiter := h.(interface {
		WaitTask(context.Context, string) (any, error)
	})
	_, err = waiter.WaitTask(ctx, started.(map[string]any)["task"].(string))
	failure, _ := json.Marshal(verify.ErrorValue(err))
	if !strings.Contains(string(failure), `"code":"worker_exited"`) || !strings.Contains(string(failure), fmt.Sprintf(`"worker_exit_code":%d`, exitCode)) || !strings.Contains(string(failure), fmt.Sprintf(`"worker_pid":%d`, pid)) {
		t.Fatalf("worker exit lost: %s", failure)
	}
}

func TestNativeProcessDump(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	h, err := native.Start(context.Background(), native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	path := filepath.Join(t.TempDir(), "owned.dmp")
	value, err := h.Call(context.Background(), native.Operation{Op: "process.dump", File: path})
	if err != nil {
		t.Fatal(err)
	}
	result := value.(map[string]any)
	if result["owned_job_verified"] != true || result["full_memory"] != false {
		t.Fatal(result)
	}
	if fmt.Sprint(result["pid"]) != fmt.Sprint(h.Info()["pid"]) {
		t.Fatal("wrong process", result)
	}
	b, err := os.ReadFile(path)
	if err != nil || len(b) < 32 || string(b[:4]) != "MDMP" {
		t.Fatal("invalid minidump", err)
	}
	if office.Hash(b) != result["sha256"] {
		t.Fatal("dump hash mismatch")
	}
	count, offset := binary.LittleEndian.Uint32(b[8:]), binary.LittleEndian.Uint32(b[12:])
	streams := map[uint32]bool{}
	for i := uint32(0); i < count; i++ {
		at := uint64(offset) + uint64(i)*12
		if at+12 > uint64(len(b)) {
			t.Fatal("invalid stream directory")
		}
		kind := binary.LittleEndian.Uint32(b[at:])
		size, rva := binary.LittleEndian.Uint32(b[at+4:]), binary.LittleEndian.Uint32(b[at+8:])
		if uint64(rva)+uint64(size) > uint64(len(b)) {
			t.Fatal("invalid stream bounds")
		}
		if kind == 3 || kind == 4 {
			if size < 4 || binary.LittleEndian.Uint32(b[rva:]) == 0 {
				t.Fatal("empty threads/modules")
			}
		}
		streams[kind] = true
	}
	if !streams[3] || !streams[4] || !streams[17] {
		t.Fatal("missing thread/module evidence", streams)
	}
	if _, err = h.Call(context.Background(), native.Operation{Op: "new", As: "afterDump"}); err != nil {
		t.Fatal("dump disrupted Word", err)
	}
	t.Logf("dump bytes=%d duration_ms=%v", len(b), result["duration_ms"])
}
