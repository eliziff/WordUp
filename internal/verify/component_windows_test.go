//go:build windows && (amd64 || arm64)

package verify_test

import (
	"context"
	"os"
	"path/filepath"
	"testing"

	"github.com/eliziff/WordUp/internal/component"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
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
	for _, id := range []string{"operation.safe-edit", "document.style-converter", "document.text-operations", "structure.detect", "ui.form-shell", "ui.progress-cancel", "ui.ribbon-command", "command.hotkey", "command.context-menu"} {
		var err error
		if id == "ui.progress-cancel" {
			_, err = component.AddWith(root, id, map[string]string{"module_prefix": "ProofProgress"})
		} else {
			_, err = component.Add(root, id)
		}
		if err != nil {
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
	if err := project.Write(root, "vba/ProofForm.vba", []byte(proofForm), ""); err != nil {
		t.Fatal(err)
	}
	design := office.Design{Name: "ProofForm", Mode: "replace", Properties: map[string]any{"Caption": "Component proof", "Width": 180.0, "Height": 90.0}}
	if err := project.Write(root, "forms/ProofForm.json", project.JSON(design), ""); err != nil {
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
		{Name: "Bounded literal replacement and formatting preservation", Operation: native.Operation{Op: "run", Macro: "Proof.CheckTextOperations"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "PASS"}}},
		{Name: "Form shell native lifecycle", Operation: native.Operation{Op: "run", Macro: "Proof.CheckFormShell"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "PASS"}}},
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
    Dim d As Document, markerDoc As Document, r As Variant, markerResult As Variant, before As String, mixed As Range
    Dim i As Long, manuscript(1 To 400) As String, stress(1 To 1400) As String
    Dim started As Single, manuscriptElapsed As Single, stressElapsed As Single
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
    Set markerDoc = Documents.Add
    markerDoc.Content.Text = "I) Introduction" & vbCr & "Part IV – Scope" & vbCr
    markerResult = WU_DetectStructure(markerDoc)
    If markerResult(0, WU_AMBIGUOUS) <> True Or InStr(CStr(markerResult(0, WU_ALTERNATIVES)), "roman:") = 0 Or InStr(CStr(markerResult(0, WU_ALTERNATIVES)), "upper_alpha:") = 0 Then Err.Raise 5, , "ambiguous Roman/alpha marker was collapsed"
    markerDoc.Content.Text = "Article 1. Definitions" & vbCr
    markerResult = WU_DetectStructure(markerDoc)
    If CStr(markerResult(0, WU_MARKER)) <> "named:article:1" Then Err.Raise 5, , "period-delimited named marker was not retained"
    markerDoc.Close SaveChanges:=wdDoNotSaveChanges
    d.Paragraphs(2).OutlineLevel = 2
    r = WU_DetectStructure(d)
    If InStr(r(4, WU_EVIDENCE), "coherent-style-family") > 0 Then Err.Raise 5, , "conflicting style votes accepted"
    For i = 1 To 400: manuscript(i) = String$(149, "x") & ".": Next i
    d.Content.Text = Join(manuscript, vbCr)
    d.Content.Style = wdStyleNormal
    d.Content.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
    started = Timer
    r = WU_DetectStructure(d)
    manuscriptElapsed = Timer - started
    If manuscriptElapsed < 0 Then manuscriptElapsed = manuscriptElapsed + 86400
    If manuscriptElapsed * 1000 > 500 Then Err.Raise 5, , "60000-character detection exceeded 500 ms"
    For i = 1 To 1400: stress(i) = "Ordinary manuscript text with no heading evidence.": Next i
    d.Content.Text = Join(stress, vbCr)
    d.Content.Style = wdStyleNormal
    d.Content.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
    started = Timer
    r = WU_DetectStructure(d)
    stressElapsed = Timer - started
    If stressElapsed < 0 Then stressElapsed = stressElapsed + 86400
    If UBound(r, 1) < 1399 Then Err.Raise 5, , "missing manuscript paragraphs"
    If stressElapsed * 1000 > 1000 Then Err.Raise 5, , "1400-paragraph stress detection exceeded one second"
    d.Close SaveChanges:=wdDoNotSaveChanges
    Check = Array("PASS", CDbl(manuscriptElapsed) * 1000, CDbl(stressElapsed) * 1000)
End Function
`

const componentProof = `Attribute VB_Name = "Proof"
Option Explicit
Private WU_Dispatched As Long
Public WU_FormShown As Boolean
Public Sub WU_Command_proof_x002D_button()
    WU_Dispatched = WU_Dispatched Or 1
End Sub
Public Sub WU_Command_proof_x005F_button()
    WU_Dispatched = WU_Dispatched Or 2
End Sub
Public Sub WU_Command__x00E9_()
    WU_Dispatched = WU_Dispatched Or 4
End Sub
Public Function CheckRibbonAndProgress() As String
    Dim control As New ProofControl, i As Long, started As Single, elapsed As Single
    ProofProgress_ResetProgress
    If ProofProgress_CancelRequested Then Err.Raise 5, , "reset left cancellation requested"
    started = Timer
    For i = 1 To 10000
        If ProofProgress_CancelRequested Then Err.Raise 5, , "checkpoint invented cancellation"
    Next i
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400
    If elapsed * 1000 > 250 Then Err.Raise 5, , "10000 cancellation checkpoints exceeded 250 ms: " & CStr(elapsed * 1000)
    ProofProgress_RequestCancel
    If Not ProofProgress_CancelRequested Then Err.Raise 5, , "cancellation request was lost"
    control.Id = "proof-button"
    WU_Dispatched = 0
    WU_RibbonCommand control
    If WU_Dispatched <> 1 Then Err.Raise 5, , "hyphenated Ribbon control ID was not dispatched independently"
    control.Id = "proof_button"
    WU_RibbonCommand control
    If WU_Dispatched <> 3 Then Err.Raise 5, , "underscore Ribbon control ID collided with another control"
    control.Id = ChrW(&HE9)
    WU_RibbonCommand control
    If WU_Dispatched <> 7 Then Err.Raise 5, , "non-ASCII Ribbon control ID was not escaped deterministically"
    CheckRibbonAndProgress = "PASS"
End Function
Public Function CheckTextOperations() As String
    Dim d As Document, result As Boolean, before As String, formatted As Range, rejected As Boolean
    Dim i As Long, lines(1 To 400) As String, started As Single, elapsed As Single
    Set d = Documents.Add
    d.Content.Text = "Alpha alpha alphabet" & vbCr
    Set formatted = d.Paragraphs(1).Range.Duplicate
    formatted.End = formatted.Start + 5
    formatted.Italic = True
    before = d.Content.Text
    result = WU_ReplaceLiteral(d, "alpha", "omega", "main", False, True)
    If Not result Then Err.Raise 5, , "literal replacement did not report a change"
    If d.Content.Text <> "omega omega alphabet" & vbCr Then Err.Raise 5, , "whole-word replacement changed the wrong text"
    If Not d.Paragraphs(1).Range.Characters(1).Italic Then Err.Raise 5, , "replacement lost direct italic formatting"
    If Not d.Undo Then Err.Raise 5, , "replacement did not create one undo record"
    If d.Content.Text <> before Then Err.Raise 5, , "replacement undo did not restore text"
    On Error Resume Next
    result = WU_ReplaceLiteral(d, "alpha", "omega", "invalid", False, False)
    rejected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    If Not rejected Then Err.Raise 5, , "invalid story scope was accepted"
    For i = 1 To 400: lines(i) = "Alpha text.": Next i
    d.Content.Text = Join(lines, vbCr)
    started = Timer
    result = WU_ReplaceLiteral(d, "Alpha", "Omega", "main", True, True)
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400
    If Not result Then Err.Raise 5, , "bulk literal replacement did not report a change"
    If elapsed * 1000 > 250 Then Err.Raise 5, , "400 literal replacements exceeded 250 ms: " & CStr(elapsed * 1000)
    d.Close SaveChanges:=wdDoNotSaveChanges
    CheckTextOperations = "PASS"
End Function
Public Function CheckFormShell() As String
    WU_FormShown = False
    WU_ShowForm "ProofForm"
    If Not WU_FormShown Then Err.Raise 5, , "form activation did not run"
    CheckFormShell = "PASS"
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
    Dim prior As Object, key As Long, failure As Long, beforeCount As Long, afterCount As Long, errorNumber As Long, errorDescription As String
    Dim stage As String
    On Error GoTo Failed
    Set prior = Application.CustomizationContext
    key = BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF12)
    WU_RemoveHotkey key
    Application.CustomizationContext = ThisDocument
    beforeCount = KeyBindings.Count
    Application.CustomizationContext = prior
    stage = "register"
    WU_RegisterHotkey key, "Proof.HotkeyTarget"
    If Not (Application.CustomizationContext Is prior) Then Err.Raise 5, , "registration changed caller context"
    stage = "status"
    If Not WU_HotkeyRegistered(key, "Proof.HotkeyTarget") Then Err.Raise 5, , "status did not find template binding from caller context"
    If Not (Application.CustomizationContext Is prior) Then Err.Raise 5, , "status changed caller context"
    stage = "reregister"
    WU_RegisterHotkey key, "Proof.HotkeyTarget"
    Application.CustomizationContext = ThisDocument
    afterCount = KeyBindings.Count
    Application.CustomizationContext = prior
    If afterCount <> beforeCount + 1 Then Err.Raise 5, , "re-registration duplicated template binding"
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
    errorNumber = Err.Number: errorDescription = Err.Description
    On Error Resume Next
    WU_RemoveHotkey key
    Application.CustomizationContext = prior
    On Error GoTo 0
    CheckHotkeys = stage & ": " & CStr(errorNumber) & ": " & errorDescription
End Function
Public Sub HotkeyTarget()
End Sub
Public Sub FailAfterEdit()
    Dim updating As Boolean, opened As Boolean, captured As Boolean, number As Long, source As String, description As String
    On Error GoTo Failed
    WU_BeginSafeEdit updating, opened, captured
    ActiveDocument.Content.InsertBefore "partial"
    Err.Raise 713, "ComponentProof", "deliberate"
CleanUp:
    WU_EndSafeEdit updating, opened, captured
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, source, description
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    Resume CleanUp
End Sub
Public Function Check() As String
    Dim number As Long, source As String, description As String, changed As Boolean, i As Long
    Dim lines(1 To 400) As String, started As Single, elapsed As Single
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
    If Not changed Then Err.Raise 5, , "conversion did not report a change"
    If WU_ConvertStyle(ActiveDocument, "ProofTarget", "ProofTarget") Then Err.Raise 5, , "same style not a no-op"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing conversion undo"
    If ActiveDocument.Paragraphs(1).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "conversion not undone"
    For i = 1 To 400: lines(i) = "Ordinary paragraph.": Next i
    ActiveDocument.Content.Text = Join(lines, vbCr)
    ActiveDocument.Content.Style = ActiveDocument.Styles(wdStyleNormal)
    started = Timer
    changed = WU_ConvertStyle(ActiveDocument, ActiveDocument.Styles(wdStyleNormal).NameLocal, "ProofTarget")
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400
    If Not changed Then Err.Raise 5, , "bulk conversion did not report a change"
    If elapsed * 1000 > 250 Then Err.Raise 5, , "bulk conversion exceeded 250 ms: " & CStr(elapsed * 1000)
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing bulk conversion undo"
    If ActiveDocument.Paragraphs(400).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "bulk conversion not undone"
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

const proofForm = `Attribute VB_Name = "ProofForm"
Option Explicit
Private Sub UserForm_Activate()
    Proof.WU_FormShown = True
    Unload Me
End Sub
`
