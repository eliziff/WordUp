//go:build windows && (amd64 || arm64)

package verify_test

import (
	"context"
	"encoding/binary"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
	"unicode/utf16"

	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
)

func TestNativeSerialTrace(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	root := filepath.Join(t.TempDir(), "trace")
	if _, err := project.New("TraceDemo", root); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "vba/WordUp_Trace.bas", []byte(verify.TraceRecorderSource), ""); err != nil {
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
	ctx := context.Background()
	h, err := native.Start(ctx, native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	if _, err := h.Call(ctx, native.Operation{Op: "open", File: b.Artifact, As: "template"}); err != nil {
		t.Fatal(err)
	}
	compiled, err := h.Call(ctx, native.Operation{Op: "compile", Target: "template", Member: "TraceDemo"})
	if err != nil {
		t.Fatal(err)
	}
	if err := verify.Check(compiled, verify.Assertion{Path: "/vba_compiled", Kind: "equals", Expected: true}); err != nil {
		t.Fatal(err)
	}
	action := `Dim d As Document, r As Range, key As String
Set d = ActiveDocument
d.Content.Text = "Before tail."
Application.Run "WordUp_TraceStart", "[""demo replacement and downstream formatting""]"
Application.UndoRecord.StartCustomRecord "Serial demo"
d.TrackRevisions = True
Set r = d.Range(0, 6)
key = Application.Run("WordUp_TraceWrite", "replace", r, "Text", """After""")
r.Text = "After"
Application.Run "WordUp_TraceResult", key, r, CLng(0)
Set r = d.Content
Application.Run "WordUp_TraceEvent", "read", "format", r, """purpose"":""read actual pending revisions"""
key = Application.Run("WordUp_TraceWrite", "format", r, "Italic", "true")
r.Italic = True
Application.Run "WordUp_TraceResult", key, r, CLng(0)
d.TrackRevisions = False
Application.UndoRecord.EndCustomRecord
Application.Run "WordUp_TraceFinish", "$output/trace.json"
Evaluate = (Application.UndoRecord.CustomRecordLevel = 0)`
	s := verify.Suite{Schema: 1, Name: "Native non-legal serial trace", SerialTrace: "trace.json", Steps: []verify.Step{
		{Name: "New", Operation: native.Operation{Op: "new", File: "$artifact", As: "doc"}},
		{Name: "Serial edit and capture", Operation: native.Operation{Op: "eval", Value: action}, Assert: []verify.Assertion{{Path: "/result", Kind: "equals", Expected: true}}},
		{Name: "One-step undo", Operation: native.Operation{Op: "eval", Value: "ActiveDocument.Undo\nEvaluate = ActiveDocument.Content.Text"}, Assert: []verify.Assertion{{Path: "/result", Kind: "equals", Expected: "Before tail.\r"}}},
		{Name: "Close", Operation: native.Operation{Op: "unload", Target: "doc"}},
	}}
	a, err := verify.Run(ctx, b.Artifact, s, h, true)
	if err != nil {
		t.Fatal(err)
	}
	bRun, err := verify.Run(ctx, b.Artifact, s, h, true)
	if err != nil {
		t.Fatal(err)
	}
	if a.SerialTrace == nil || a.SerialTrace.Events != 5 {
		t.Fatal("missing serial events", a.SerialTrace)
	}
	raw, err := os.ReadFile(a.SerialTrace.File)
	if err != nil {
		t.Fatal(err)
	}
	units := make([]uint16, (len(raw)-2)/2)
	for i := range units {
		units[i] = binary.LittleEndian.Uint16(raw[2+i*2:])
	}
	var trace verify.SerialTrace
	if err := json.Unmarshal([]byte(string(utf16.Decode(units))), &trace); err != nil {
		t.Fatal(err)
	}
	touches, ok := trace.Events[2].Data["touches"].([]any)
	if !ok || len(touches) != 1 || touches[0].(map[string]any)["operation"] != "w1" {
		t.Fatal("native downstream read lost earlier edit anchor", trace.Events[2])
	}
	if _, err := verify.Compare(a, bRun); err != nil {
		t.Fatal(err)
	}
	t.Logf("Native trace: %d events, exact ordered repeat, one-step undo; macro step %.1f/%.1f ms", a.SerialTrace.Events, a.Observations[1].DurationMS, bRun.Observations[1].DurationMS)
	identity := `Dim d As Document, other As Document
Set d = ActiveDocument
d.Content.Text = "Section one"
d.Sections.Add
d.Sections(2).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
d.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = "First header"
d.Sections(2).Headers(wdHeaderFooterPrimary).Range.Text = "Second header"
Application.Run "WordUp_TraceStart", "[""distinct story instances and documents""]"
Application.Run "WordUp_TraceEvent", "read", "identity", d.Sections(1).Headers(wdHeaderFooterPrimary).Range, ""
Application.Run "WordUp_TraceEvent", "read", "identity", d.Sections(2).Headers(wdHeaderFooterPrimary).Range, ""
Set other = Documents.Add
other.Content.Text = "Other document"
d.Activate
Application.Run "WordUp_TraceEvent", "read", "identity", other.Content, ""
Application.Run "WordUp_TraceFinish", "$output/trace.json"
other.Close wdDoNotSaveChanges
Evaluate = True`
	identities := verify.Suite{Schema: 1, Name: "Story identities", SerialTrace: "trace.json", Steps: []verify.Step{
		{Name: "New", Operation: native.Operation{Op: "new", File: "$artifact", As: "doc"}},
		{Name: "Observe identities", Operation: native.Operation{Op: "eval", Value: identity}, Assert: []verify.Assertion{{Path: "/result", Kind: "equals", Expected: true}}},
		{Name: "Close", Operation: native.Operation{Op: "unload", Target: "doc"}},
	}}
	iReport, err := verify.Run(ctx, b.Artifact, identities, h, true)
	if err != nil {
		t.Fatal(err)
	}
	iRaw, err := os.ReadFile(iReport.SerialTrace.File)
	if err != nil {
		t.Fatal(err)
	}
	iUnits := make([]uint16, (len(iRaw)-2)/2)
	for i := range iUnits {
		iUnits[i] = binary.LittleEndian.Uint16(iRaw[2+i*2:])
	}
	var distinct verify.SerialTrace
	if err := json.Unmarshal([]byte(string(utf16.Decode(iUnits))), &distinct); err != nil {
		t.Fatal(err)
	}
	if len(distinct.Events) != 3 || distinct.Events[0].Story == distinct.Events[1].Story || distinct.Events[0].Data["story_type"] != distinct.Events[1].Data["story_type"] || distinct.Events[0].Document == distinct.Events[2].Document {
		t.Fatal("native story/document identity collision", distinct)
	}
}
