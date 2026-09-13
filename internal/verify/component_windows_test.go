//go:build windows && (amd64 || arm64)

package verify_test

import (
	"context"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/component"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
)

func TestNativeComponentCleanup(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	root := filepath.Join(t.TempDir(), "components")
	if _, err := project.New("ComponentProof", root); err != nil {
		t.Fatal(err)
	}
	for _, id := range []string{"operation.safe-edit", "document.style-converter", "structure.detect", "ui.progress-cancel", "ui.ribbon-command", "command.hotkey", "command.context-menu"} {
		if _, err := component.Add(root, id); err != nil {
			t.Fatal(err)
		}
	}
	if err := project.Write(root, "vba/Proof.bas", []byte(componentProof), ""); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "vba/StructureProof.bas", []byte(structureProof), ""); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "vba/ProofControl.cls", []byte(proofControl), ""); err != nil {
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
	suite := verify.Suite{Schema: 1, Name: "Component cleanup and error propagation", RequireCompile: true, Steps: []verify.Step{
		{Name: "Open", Operation: native.Operation{Op: "open", File: "$artifact", As: "doc"}},
		{Name: "Compile", Operation: native.Operation{Op: "compile", Target: "doc", Member: "$project"}, Assert: []verify.Assertion{{Path: "/vba_compiled", Kind: "equals", Expected: true}}},
		{Name: "Failure, undo, cached styles and caller state", Operation: native.Operation{Op: "run", Macro: "Proof.Check"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "PASS"}}},
		{Name: "Hotkey caller context", Operation: native.Operation{Op: "run", Macro: "Proof.CheckHotkeys"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "PASS"}}},
		{Name: "Context menu ownership and repetition", Operation: native.Operation{Op: "run", Macro: "Proof.CheckMenus"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "PASS"}}},
		{Name: "Ribbon dispatch and cooperative cancellation", Operation: native.Operation{Op: "run", Macro: "Proof.CheckRibbonAndProgress"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "PASS"}}},
		{Name: "Structure counterexamples and cached style votes", Operation: native.Operation{Op: "run", Macro: "StructureProof.Check"}, Assert: []verify.Assertion{{Path: "/array/0", Kind: "equals", Expected: "PASS"}}},
	}}
	report, err := verify.Run(context.Background(), b.Artifact, suite, nil, true)
	if err != nil {
		t.Fatalf("native component acceptance failed: %v", err)
	}
	t.Logf("native component acceptance: %.3f ms", report.DurationMS)
	t.Logf("structure result: %v", report.Observations[len(report.Observations)-1].Result)
}

const structureProof = `Attribute VB_Name = "StructureProof"
Option Explicit
Public Function Check() As Variant
    Dim d As Document, r As Variant, before As String, mixed As Range
    Dim i As Long, text As String, started As Single, elapsed As Single
    Set d = Documents.Add
    d.Content.Text = "First heading" & vbCr & "Second heading" & vbCr & "Ordinary body sentence." & vbCr & "Mixed emphasis text" & vbCr & "THIRD HEADING"
    d.Content.Style = wdStyleNormal
    d.Content.Font.Bold = False
    d.Content.Font.SmallCaps = False
    d.Content.Font.AllCaps = False
    d.Content.ParagraphFormat.KeepWithNext = False
    d.Paragraphs(1).OutlineLevel = 1
    d.Paragraphs(2).OutlineLevel = 1
    Set mixed = d.Paragraphs(4).Range.Duplicate
    mixed.End = mixed.Start + 5
    mixed.Bold = True
    d.Paragraphs(5).Range.Bold = True
    d.Paragraphs(5).KeepWithNext = True
    before = d.Content.Text
    d.Saved = True
    r = WU_DetectStructure(d)
    If r(2, WU_ROLE) <> "body" Or r(2, WU_LEVEL) <> 0 Then Err.Raise 5, , "ordinary body inherited outline from direct formatting"
    If InStr(r(3, WU_EVIDENCE), "aggregate-bold") > 0 Then Err.Raise 5, , "mixed bold treated as fully bold"
    If r(4, WU_LEVEL) <> 1 Or InStr(r(4, WU_EVIDENCE), "coherent-style-family") = 0 Then Err.Raise 5, , "coherent candidate not resolved"
    If r(0, WU_PARENT) <> 0 Or r(2, WU_AMBIGUOUS) <> False Then Err.Raise 5, , "uninitialized contract values"
    If d.Content.Text <> before Or Not d.Saved Then Err.Raise 5, , "detector changed document"
    d.Paragraphs(2).OutlineLevel = 2
    r = WU_DetectStructure(d)
    If InStr(r(4, WU_EVIDENCE), "coherent-style-family") > 0 Then Err.Raise 5, , "conflicting style votes accepted"
    For i = 1 To 1400
        text = text & "Ordinary manuscript text with no heading evidence." & vbCr
    Next i
    d.Content.Text = text
    d.Content.Style = wdStyleNormal
    d.Content.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
    started = Timer
    r = WU_DetectStructure(d)
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400
    If UBound(r, 1) < 1399 Then Err.Raise 5, , "missing manuscript paragraphs"
    d.Close SaveChanges:=wdDoNotSaveChanges
    Check = Array("PASS", CDbl(elapsed) * 1000)
End Function
`

const componentProof = `Attribute VB_Name = "Proof"
Option Explicit
Private WU_Dispatched As Boolean
Public Sub WU_Command_proof_button()
    WU_Dispatched = True
End Sub
Public Function CheckRibbonAndProgress() As String
    Dim control As New ProofControl
    WU_ResetProgress
    If WU_CancelRequested Then Err.Raise 5, , "reset left cancellation requested"
    WU_RequestCancel
    If Not WU_CancelRequested Then Err.Raise 5, , "cancellation request was lost"
    control.Id = "proof-button"
    WU_Dispatched = False
    WU_RibbonCommand control
    If Not WU_Dispatched Then Err.Raise 5, , "Ribbon command was not dispatched"
    CheckRibbonAndProgress = "PASS"
End Function
Public Function CheckMenus() As String
    Dim prior As Object, failure As String, stage As String
    On Error GoTo Failed
    Set prior = Application.CustomizationContext
    stage = "register"
    WU_RegisterContextMenu "First", "Proof.HotkeyTarget"
    WU_RegisterContextMenu "Second", "Proof.HotkeyTarget"
    If Not WU_ContextMenuRegistered("Second", "Proof.HotkeyTarget") Then Err.Raise 5, , "replacement menu wiring incorrect"
    If Not (Application.CustomizationContext Is prior) Then Err.Raise 5, , "registration changed caller context"
    stage = "remove"
    WU_RemoveContextMenu
    WU_RemoveContextMenu
    If WU_ContextMenuRegistered Then Err.Raise 5, , "component retained removed menu"
    If Not (Application.CustomizationContext Is prior) Then Err.Raise 5, , "removal changed caller context"
    CheckMenus = "PASS"
CleanUp:
    On Error Resume Next
    WU_RemoveContextMenu
    If failure <> "" Then CheckMenus = failure
    Exit Function
Failed:
    failure = stage & ": " & CStr(Err.Number) & ": " & Err.Description
    Resume CleanUp
End Function
Public Function CheckHotkeys() As String
    Dim prior As Object, key As Long, failure As Long
    Dim stage As String
    On Error GoTo Failed
    Set prior = Application.CustomizationContext
    key = BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF12)
    stage = "register"
    WU_RegisterHotkey key, "Proof.HotkeyTarget"
    If Not (Application.CustomizationContext Is prior) Then Err.Raise 5, , "registration changed caller context"
    stage = "remove"
    WU_RemoveHotkey key
    If Not (Application.CustomizationContext Is prior) Then Err.Raise 5, , "removal changed caller context"
    On Error Resume Next
    WU_RegisterHotkey -1, "Proof.HotkeyTarget"
    failure = Err.Number
    On Error GoTo 0
    If failure = 0 Then Err.Raise 5, , "invalid hotkey did not fail"
    If Not (Application.CustomizationContext Is prior) Then Err.Raise 5, , "failure changed caller context"
    CheckHotkeys = "PASS"
    Exit Function
Failed:
    CheckHotkeys = stage & ": " & CStr(Err.Number) & ": " & Err.Description
End Function
Public Sub HotkeyTarget()
End Sub
Public Sub FailAfterEdit()
    Dim updating As Boolean, opened As Boolean, number As Long, source As String, description As String
    On Error GoTo Failed
    WU_BeginSafeEdit updating, opened
    ActiveDocument.Content.InsertBefore "partial"
    Err.Raise 713, "ComponentProof", "deliberate"
CleanUp:
    WU_EndSafeEdit updating, opened
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, source, description
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    Resume CleanUp
End Sub
Public Function Check() As String
    Dim number As Long, source As String, description As String, changed As Long
    ActiveDocument.Content.Text = "original"
    Application.ScreenUpdating = False
    On Error Resume Next
    FailAfterEdit
    number = Err.Number: source = Err.Source: description = Err.Description
    On Error GoTo 0
    If number <> 713 Or source <> "ComponentProof" Or description <> "deliberate" Then Err.Raise 5, , "lost callback error"
    If Application.ScreenUpdating Or Application.UndoRecord.CustomRecordLevel <> 0 Then Err.Raise 5, , "lost caller state"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing undo"
    If ActiveDocument.Content.Text <> "original" & vbCr Then Err.Raise 5, , "partial edit not undone"
    On Error Resume Next
    changed = WU_ConvertStyle(ActiveDocument, "Normal", "MissingProofStyle")
    number = Err.Number
    On Error GoTo 0
    If number = 0 Then Err.Raise 5, , "missing style silently accepted"
    If Application.ScreenUpdating Or Application.UndoRecord.CustomRecordLevel <> 0 Then Err.Raise 5, , "failed conversion lost state"
    ActiveDocument.Styles.Add "ProofTarget", wdStyleTypeParagraph
    ActiveDocument.Content.Style = ActiveDocument.Styles(wdStyleNormal)
    changed = WU_ConvertStyle(ActiveDocument, ActiveDocument.Styles(wdStyleNormal).NameLocal, "ProofTarget")
    If changed <> 1 Then Err.Raise 5, , "conversion count"
    If WU_ConvertStyle(ActiveDocument, "ProofTarget", "ProofTarget") <> 0 Then Err.Raise 5, , "same style not a no-op"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing conversion undo"
    If ActiveDocument.Paragraphs(1).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "conversion not undone"
    Application.ScreenUpdating = True
    Check = "PASS"
End Function
`

const proofControl = `Option Explicit
Private WU_Id As String
Public Property Get Id() As String
    Id = WU_Id
End Property
Public Property Let Id(ByVal value As String)
    WU_Id = value
End Property
`

func TestNativeComponentCorpus(t *testing.T) {
	corpus := os.Getenv("WORDUP_CORPUS")
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" || corpus == "" {
		t.Skip("set WORDUP_NATIVE_TEST=1 and WORDUP_CORPUS")
	}
	files, err := filepath.Glob(filepath.Join(corpus, "*.docx"))
	if err != nil || len(files) != 109 {
		t.Fatalf("corpus: %d documents, %v", len(files), err)
	}
	root := filepath.Join(t.TempDir(), "corpus-proof")
	if _, err := project.New("CorpusProof", root); err != nil {
		t.Fatal(err)
	}
	for _, id := range []string{"structure.detect", "document.style-converter"} {
		if _, err := component.Add(root, id); err != nil {
			t.Fatal(err)
		}
	}
	if err := project.Write(root, "vba/ManuscriptProof.bas", []byte(manuscriptProof), ""); err != nil {
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
	steps := []verify.Step{
		{Name: "Load components", Operation: native.Operation{Op: "open", File: "$artifact", As: "components"}},
		{Name: "Compile components", Operation: native.Operation{Op: "compile", Target: "components", Member: "$project"}, Assert: []verify.Assertion{{Path: "/vba_compiled", Kind: "equals", Expected: true}}},
	}
	for _, file := range files {
		if strings.HasSuffix(file, "--3a668c4999c4.docx") {
			t.Log("excluded known malformed ZIP separators; original unchanged")
			continue
		}
		steps = append(steps,
			verify.Step{Name: "Open " + filepath.Base(file), Operation: native.Operation{Op: "open", File: file, As: "manuscript"}},
			verify.Step{Name: "Attach " + filepath.Base(file), Operation: native.Operation{Op: "put", Target: "manuscript", Member: "AttachedTemplate", Value: "$artifact"}},
			verify.Step{Name: "Preservation " + filepath.Base(file), Operation: native.Operation{Op: "run", Macro: "ManuscriptProof.CheckDocument"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "PASS"}}},
			verify.Step{Name: "Discard " + filepath.Base(file), Operation: native.Operation{Op: "unload", Target: "manuscript"}},
		)
	}
	report, err := verify.Run(ctx, b.Artifact, verify.Suite{Schema: 1, Name: "Retained manuscript component preservation", RequireCompile: true, Steps: steps}, h, true)
	if output := os.Getenv("WORDUP_COMPONENT_REPORT"); output != "" {
		if saveErr := os.WriteFile(output, project.JSON(report), 0600); saveErr != nil {
			t.Fatal(saveErr)
		}
	}
	if err != nil {
		t.Fatalf("%v", err)
	}
	t.Logf("108 native manuscripts passed; %.3f ms", report.DurationMS)
}

const manuscriptProof = `Attribute VB_Name = "ManuscriptProof"
Option Explicit
Public Function CheckDocument() As Variant
    Dim doc As Document, baseline As String, detected As Variant
    Dim originalStyle As String, changed As Long
    Set doc = ActiveDocument
    baseline = doc.WordOpenXML
    baseline = doc.WordOpenXML
    If baseline <> doc.WordOpenXML Then CheckDocument = Array("no-op XML mismatch", baseline, doc.WordOpenXML): Exit Function
    detected = WU_DetectStructure(doc)
    If baseline <> doc.WordOpenXML Then CheckDocument = Array("detector XML mismatch", baseline, doc.WordOpenXML): Exit Function
    doc.Styles.Add "WordUpPreservationTarget", wdStyleTypeParagraph
    originalStyle = doc.Paragraphs(1).Style.NameLocal
    baseline = doc.WordOpenXML
    changed = WU_ConvertStyle(doc, originalStyle, "WordUpPreservationTarget")
    If changed < 1 Then Err.Raise 5, , "no paragraph converted"
    If Not doc.Undo Then Err.Raise 5, , "missing conversion undo"
    If baseline <> doc.WordOpenXML Then CheckDocument = Array("undo XML mismatch", baseline, doc.WordOpenXML): Exit Function
    CheckDocument = "PASS"
End Function
`
