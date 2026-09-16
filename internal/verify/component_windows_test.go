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
	for _, id := range []string{"operation.safe-edit", "document.stories", "document.style-converter", "document.field-refresh", "document.text-operations", "document.paragraph-index", "structure.detect", "ui.form-shell", "batch.progress", "ui.ribbon-command", "command.hotkey", "command.context-menu"} {
		var err error
		if id == "batch.progress" {
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
		{Name: "Field refresh scopes, return codes and undo", Operation: native.Operation{Op: "run", Macro: "Proof.CheckFieldRefresh"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "PASS"}}},
		{Name: "Paragraph index matches the object model", Operation: native.Operation{Op: "run", Macro: "Proof.CheckParagraphIndex"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "PASS"}}},
		{Name: "Batch progress, stage timing and notification capture", Operation: native.Operation{Op: "run", Macro: "Proof.CheckBatch"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "PASS"}}},
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
    Dim control As New ProofControl, i As Long, started As Single, elapsed As Single, stage As String
    On Error GoTo Failed
    stage = "reset"
    ProofProgress_ResetProgress
    If ProofProgress_CancelRequested Then Err.Raise 5, , "reset left cancellation requested"
    stage = "checkpoints"
    started = Timer
    For i = 1 To 10000
        If ProofProgress_CancelRequested Then Err.Raise 5, , "checkpoint invented cancellation"
    Next i
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400
    If elapsed * 1000 > 250 Then Err.Raise 5, , "10000 cancellation checkpoints exceeded 250 ms: " & CStr(elapsed * 1000)
    stage = "cancel"
    ProofProgress_RequestCancel
    If Not ProofProgress_CancelRequested Then Err.Raise 5, , "cancellation request was lost"
    stage = "ribbon hyphen"
    control.Id = "proof-button"
    WU_Dispatched = 0
    WU_RibbonCommand control
    If WU_Dispatched <> 1 Then Err.Raise 5, , "hyphenated Ribbon control ID was not dispatched independently"
    stage = "ribbon underscore"
    control.Id = "proof_button"
    WU_RibbonCommand control
    If WU_Dispatched <> 3 Then Err.Raise 5, , "underscore Ribbon control ID collided with another control"
    stage = "ribbon unicode"
    control.Id = ChrW(&HE9)
    WU_RibbonCommand control
    If WU_Dispatched <> 7 Then Err.Raise 5, , "non-ASCII Ribbon control ID was not escaped deterministically"
    CheckRibbonAndProgress = "PASS"
    Exit Function
Failed:
    CheckRibbonAndProgress = stage & ": " & CStr(Err.Number) & ": " & Err.Description
End Function
Public Function CheckTextOperations() As String
    Dim d As Document, result As Boolean, before As String, bodyAfterNote As String, unicodeText As String, formatted As Range, scoped As Range, header As Range, footer As Range, rejected As Boolean, note As Footnote, priorUpdating As Boolean, table As Table
    Dim rules(0 To 2, 0 To 1) As Variant, characterRuns(0 To 1, 0 To 2) As Variant, batchChanged As Long, rangeBatchChanged As Long, literalCount As Long, citationStyle As Style, styled As Long, styleRunChanged As Long
    Dim i As Long, lines(1 To 400) As String, started As Single, elapsed As Single, wildcardCount As Long, wildcardRangeCount As Long
    Set d = Documents.Add
    d.Content.Text = "Alpha alpha alphabet"
    Set formatted = d.Paragraphs(1).Range.Duplicate
    formatted.End = formatted.Start + 5
    formatted.Italic = True
    before = d.Content.Text
    priorUpdating = Application.ScreenUpdating
    d.Saved = True
    result = WU_ReplaceText(d, "Alpha", "Alpha", "main", True, False)
    If result Or Not d.Saved Then Err.Raise 5, , "exact no-op opened an edit or changed the document"
    d.Saved = True
    result = WU_ReplaceText(d, "Missing", "Absent", "main", True, False)
    If result Or Not d.Saved Then Err.Raise 5, , "no-match replacement opened an edit or changed the document"
    result = WU_ReplaceText(d, "alpha", "omega", "main", False, True)
    If Not result Then Err.Raise 5, , "literal replacement did not report a change"
    If Application.ScreenUpdating <> priorUpdating Then Err.Raise 5, , "replacement did not restore ScreenUpdating"
    If d.Content.Text <> "Omega omega alphabet" & vbCr Then Err.Raise 5, , "whole-word replacement changed the wrong text"
    If Not d.Paragraphs(1).Range.Characters(1).Italic Then Err.Raise 5, , "replacement lost direct italic formatting"
    If Not d.Undo Then Err.Raise 5, , "replacement did not create one undo record"
    If d.Content.Text <> before Then Err.Raise 5, , "replacement undo did not restore text"
    d.Content.Text = "a*b aab"
    d.Saved = True
    result = WU_ReplaceText(d, "a*b", "literal-hit", "main", True, False)
    If Not result Or d.Content.Text <> "literal-hit aab" & vbCr Then Err.Raise 5, , "literal replacement interpreted wildcard syntax"
    If Not d.Undo Then Err.Raise 5, , "literal wildcard-safety replacement did not create one undo record"
    If d.Content.Text <> "a*b aab" & vbCr Then Err.Raise 5, , "literal wildcard-safety undo did not restore text"
    d.Content.Text = Left$(before, Len(before) - 1)
    Set scoped = d.Paragraphs(1).Range.Duplicate
    scoped.Start = scoped.Start + 6
    scoped.End = scoped.Start + 5
    result = WU_ReplaceTextInRange(scoped, "alpha", "beta", True, True)
    If Not result Or d.Content.Text <> "Alpha beta alphabet" & vbCr Then Err.Raise 5, , "range replacement escaped its requested boundary: result=" & CStr(result) & " text=[" & Replace(d.Content.Text, vbCr, "|") & "] scoped=" & scoped.Start & "-" & scoped.End & " paragraphs=" & d.Paragraphs.Count
    If Application.ScreenUpdating <> priorUpdating Then Err.Raise 5, , "range replacement did not restore ScreenUpdating"
    If Not d.Undo Then Err.Raise 5, , "range replacement did not create one undo record"
    If d.Content.Text <> before Then Err.Raise 5, , "range replacement undo did not restore text"
    d.Content.Text = "Case 12; Case 34"
    d.Saved = True
    wildcardCount = WU_CountText(d, "Case [0-9]{2}", "main", True, True, True)
    If wildcardCount <> 2 Or Not d.Saved Then Err.Raise 5, , "wildcard count missed a native pattern match or changed document state"
    Set scoped = d.Paragraphs(1).Range.Duplicate
    wildcardRangeCount = WU_CountTextInRange(scoped, "Case [0-9]{2}", True, True, True)
    If wildcardRangeCount <> 2 Then Err.Raise 5, , "range wildcard count missed a bounded pattern match"
    scoped.End = scoped.Start + 8
    If WU_CountTextInRange(scoped, "Case [0-9]{2}", True, True, True) <> 1 Then Err.Raise 5, , "range wildcard count escaped its boundary"
    d.Saved = True
    result = WU_ReplaceText(d, "Case ([0-9]{2})", "Matter \1", "main", True, True, True)
    If Not result Or d.Content.Text <> "Matter 12; Matter 34" & vbCr Then Err.Raise 5, , "wildcard replacement did not preserve its capture"
    If Application.ScreenUpdating <> priorUpdating Then Err.Raise 5, , "wildcard replacement did not restore ScreenUpdating"
    If Not d.Undo Then Err.Raise 5, , "wildcard replacement did not create one undo record"
    If d.Content.Text <> "Case 12; Case 34" & vbCr Then Err.Raise 5, , "wildcard replacement undo did not restore text"
    d.Content.Text = "Case 12; Case 34"
    rules(0, 0) = "Case ([0-9]{2})": rules(0, 1) = "Matter \1"
    rules(1, 0) = "Unused ([0-9]{2})": rules(1, 1) = "Still unused"
    rules(2, 0) = "Never ([0-9]{2})": rules(2, 1) = "Still never"
    d.Saved = True
    batchChanged = WU_ReplaceBatch(d, rules, "main", True, True, True)
    If batchChanged <> 1 Or d.Content.Text <> "Matter 12; Matter 34" & vbCr Then Err.Raise 5, , "wildcard replacement batch did not apply exactly its matching rule"
    If Application.ScreenUpdating <> priorUpdating Then Err.Raise 5, , "wildcard replacement batch did not restore ScreenUpdating"
    If Not d.Undo Then Err.Raise 5, , "wildcard replacement batch did not create one undo record"
    If d.Content.Text <> "Case 12; Case 34" & vbCr Then Err.Raise 5, , "wildcard replacement batch undo did not restore text"
    d.Content.Text = "Alpha alpha Gamma"
    rules(0, 0) = "Alpha": rules(0, 1) = "Omega"
    rules(1, 0) = "Gamma": rules(1, 1) = "Delta"
    rules(2, 0) = "Unused": rules(2, 1) = "Still unused"
    d.Saved = True
    batchChanged = WU_ReplaceBatch(d, rules, "main", True, True)
    If batchChanged <> 2 Or d.Content.Text <> "Omega alpha Delta" & vbCr Then Err.Raise 5, , "batch replacement did not apply exactly its matching pairs"
    If Application.ScreenUpdating <> priorUpdating Then Err.Raise 5, , "batch replacement did not restore ScreenUpdating"
    If Not d.Undo Then Err.Raise 5, , "batch replacement did not create one undo record"
    If d.Content.Text <> "Alpha alpha Gamma" & vbCr Then Err.Raise 5, , "batch replacement undo did not restore text"
    d.Content.Text = "First Alpha" & vbCr & "Second Gamma"
    Set scoped = d.Paragraphs(1).Range.Duplicate
    rangeBatchChanged = WU_ReplaceBatchInRange(scoped, rules, True, True)
    If rangeBatchChanged <> 1 Or d.Content.Text <> "First Omega" & vbCr & "Second Gamma" & vbCr Then Err.Raise 5, , "range batch replacement escaped its requested boundary"
    d.Saved = True
    literalCount = WU_CountText(d, "Gamma", "main", True, True)
    If literalCount <> 1 Or Not d.Saved Then Err.Raise 5, , "literal counter returned the wrong main-story count or changed document state"
    literalCount = WU_CountTextInRange(scoped, "Omega", True, True)
    If literalCount <> 1 Then Err.Raise 5, , "range literal counter missed the bounded match"
    If Application.ScreenUpdating <> priorUpdating Then Err.Raise 5, , "range batch replacement did not restore ScreenUpdating"
    If Not d.Undo Then Err.Raise 5, , "range batch replacement did not create one undo record"
    If d.Content.Text <> "First Alpha" & vbCr & "Second Gamma" & vbCr Then Err.Raise 5, , "range batch replacement undo did not restore text"
    d.Content.Text = "Alpha Alpha Gamma"
    ' Assigned text inherits the direct formatting of the first character it
    ' replaced (the italic "Alpha" above); reset so italics below come only
    ' from the character style under test.
    d.Content.Font.Reset
    Set citationStyle = d.Styles.Add(Name:="Proof Citation", Type:=wdStyleTypeCharacter)
    citationStyle.Font.Italic = True
    d.Saved = True
    styled = WU_StyleMatches(d, "Alpha", citationStyle.NameLocal, "main", True, True)
    If styled <> 2 Then Err.Raise 5, , "character style matcher missed a literal occurrence"
    If d.Content.Text <> "Alpha Alpha Gamma" & vbCr Then Err.Raise 5, , "character style matcher changed text"
    If Not d.Paragraphs(1).Range.Characters(1).Italic Or Not d.Paragraphs(1).Range.Characters(7).Italic Or d.Paragraphs(1).Range.Characters(13).Italic Then Err.Raise 5, , "character style matcher did not apply the requested style to exactly its matches: italic1=" & CStr(d.Paragraphs(1).Range.Characters(1).Italic) & " italic7=" & CStr(d.Paragraphs(1).Range.Characters(7).Italic) & " italic13=" & CStr(d.Paragraphs(1).Range.Characters(13).Italic) & " style1=" & d.Paragraphs(1).Range.Characters(1).Style.NameLocal & " style7=" & d.Paragraphs(1).Range.Characters(7).Style.NameLocal & " style13=" & d.Paragraphs(1).Range.Characters(13).Style.NameLocal & " text=[" & Replace(d.Content.Text, vbCr, "|") & "]"
    If Not d.Undo Then Err.Raise 5, , "character style matcher did not create one undo record"
    If d.Paragraphs(1).Range.Characters(1).Italic Or d.Paragraphs(1).Range.Characters(7).Italic Then Err.Raise 5, , "character style matcher undo did not restore formatting"
    d.Saved = True
    styled = WU_StyleMatches(d, "Missing", citationStyle.NameLocal, "main", True, True)
    If styled <> 0 Or Not d.Saved Then Err.Raise 5, , "no-match style application opened an edit"
    d.Content.Text = "Case 12; Case 34"
    d.Content.Font.Reset
    styled = WU_StyleMatches(d, "Case [0-9]{2}", citationStyle.NameLocal, "main", True, True, True)
    If styled <> 2 Or d.Content.Text <> "Case 12; Case 34" & vbCr Then Err.Raise 5, , "wildcard character style matcher changed text or missed a match"
    If Not d.Paragraphs(1).Range.Characters(1).Italic Or Not d.Paragraphs(1).Range.Characters(10).Italic Then Err.Raise 5, , "wildcard character style matcher missed a citation span"
    If Not d.Undo Then Err.Raise 5, , "wildcard character style matcher did not create one undo record"
    If d.Paragraphs(1).Range.Characters(1).Italic Or d.Paragraphs(1).Range.Characters(10).Italic Then Err.Raise 5, , "wildcard character style matcher undo did not restore formatting"
    d.Content.Text = "First Alpha Gamma" & vbCr & "Second Alpha Gamma"
    d.Content.Font.Reset
    Set scoped = d.Paragraphs(1).Range.Duplicate
    styled = WU_StyleMatchesInRange(scoped, "Alpha", citationStyle.NameLocal, True, True)
    If styled <> 1 Or Not d.Paragraphs(1).Range.Characters(7).Italic Then Err.Raise 5, , "range character style matcher missed its bounded match"
    If d.Paragraphs(2).Range.Characters(8).Italic Then Err.Raise 5, , "range character style matcher escaped its boundary"
    If Not d.Undo Then Err.Raise 5, , "range character style matcher did not create one undo record"
    If d.Paragraphs(1).Range.Characters(7).Italic Then Err.Raise 5, , "range character style matcher undo did not restore formatting"
    Set formatted = d.Paragraphs(1).Range.Duplicate
    formatted.End = formatted.Start + 5
    formatted.Bold = True
    characterRuns(0, 0) = scoped.Start + 6: characterRuns(0, 1) = scoped.Start + 11: characterRuns(0, 2) = citationStyle.NameLocal
    characterRuns(1, 0) = scoped.Start + 12: characterRuns(1, 1) = scoped.Start + 17: characterRuns(1, 2) = citationStyle.NameLocal
    d.Saved = True
    styleRunChanged = WU_ApplyCharacterStyleRuns(scoped, characterRuns)
    If styleRunChanged <> 2 Then Err.Raise 5, , "character style offset runs did not style each exact span"
    If d.Content.Text <> "First Alpha Gamma" & vbCr & "Second Alpha Gamma" & vbCr Then Err.Raise 5, , "character style offset runs changed text"
    If Not d.Paragraphs(1).Range.Characters(7).Italic Or Not d.Paragraphs(1).Range.Characters(13).Italic Then Err.Raise 5, , "character style offset runs missed a span"
    If Not d.Paragraphs(1).Range.Characters(1).Bold Then Err.Raise 5, , "character style offset runs lost unrelated direct formatting"
    If d.Paragraphs(2).Range.Characters(8).Italic Then Err.Raise 5, , "character style offset runs escaped its exact range"
    If WU_ApplyCharacterStyleRuns(scoped, characterRuns) <> 0 Then Err.Raise 5, , "re-applying identical runs is not a no-op"
    If Not d.Undo Then Err.Raise 5, , "character style offset runs did not create one undo record"
    If d.Paragraphs(1).Range.Characters(7).Italic Or d.Paragraphs(1).Range.Characters(13).Italic Then Err.Raise 5, , "character style offset runs undo did not restore formatting"
    If Not d.Paragraphs(1).Range.Characters(1).Bold Then Err.Raise 5, , "character style offset runs undo lost unrelated direct formatting"
    d.Saved = True
    characterRuns(1, 2) = "Missing proof style"
    On Error Resume Next
    styleRunChanged = WU_ApplyCharacterStyleRuns(scoped, characterRuns)
    rejected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    If Not rejected Or Not d.Saved Then Err.Raise 5, , "invalid character style offset run was not rejected before mutation"
    characterRuns(1, 2) = citationStyle.NameLocal
    characterRuns(1, 0) = scoped.Start + 8
    On Error Resume Next
    styleRunChanged = WU_ApplyCharacterStyleRuns(scoped, characterRuns)
    rejected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    If Not rejected Or Not d.Saved Then Err.Raise 5, , "overlapping offset runs were not rejected before mutation"
    d.Content.Text = Left$(before, Len(before) - 1)
    Set scoped = d.Paragraphs(1).Range.Duplicate
    scoped.Collapse wdCollapseStart
    d.Saved = True
    result = WU_ReplaceTextInRange(scoped, "Alpha", "Omega", True, True)
    If result Or Not d.Saved Then Err.Raise 5, , "empty range opened an edit or changed the document"
    On Error Resume Next
    result = WU_ReplaceText(d, "alpha", "omega", "invalid", False, False)
    rejected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    If Not rejected Then Err.Raise 5, , "invalid story scope was accepted"
    On Error Resume Next
    result = WU_ReplaceText(d, String$(200, "^"), "omega", "main", True, False)
    rejected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    If Not rejected Then Err.Raise 5, , "escaped find input limit was not enforced"
    For i = 1 To 400: lines(i) = "Alpha text.": Next i
    d.Content.Text = Join(lines, vbCr)
    started = Timer
    result = WU_ReplaceText(d, "Alpha", "Omega", "main", True, True)
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400
    If Not result Then Err.Raise 5, , "bulk literal replacement did not report a change"
    If elapsed * 1000 > 500 Then Err.Raise 5, , "400 literal replacements exceeded 500 ms: " & CStr(elapsed * 1000)
    started = Timer
    literalCount = WU_CountText(d, "Omega", "main", True, True)
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400
    If literalCount <> 400 Then Err.Raise 5, , "bulk count returned " & CStr(literalCount)
    If elapsed * 1000 > 500 Then Err.Raise 5, , "400 counted matches exceeded 500 ms: " & CStr(elapsed * 1000)
    d.Content.Text = "literal ^p marker"
    result = WU_ReplaceText(d, "^p", "^t", "main", True, False)
    If Not result Or d.Content.Text <> "literal ^t marker" & vbCr Then Err.Raise 5, , "literal caret text was interpreted as a Word control token"
    unicodeText = "R" & ChrW(&HE9) & "sum" & ChrW(&HE9)
    d.Content.Text = unicodeText & " " & unicodeText
    Set formatted = d.Paragraphs(1).Range.Duplicate
    formatted.End = formatted.Start + Len(unicodeText)
    formatted.Italic = True
    result = WU_ReplaceText(d, unicodeText, "Case" & ChrW(&H2013) & "Name", "main", True, True)
    If Not result Or d.Content.Text <> "Case" & ChrW(&H2013) & "Name " & "Case" & ChrW(&H2013) & "Name" & vbCr Then Err.Raise 5, , "Unicode literal replacement changed the wrong text"
    If Not d.Paragraphs(1).Range.Characters(1).Italic Then Err.Raise 5, , "Unicode replacement lost direct formatting"
    If Not d.Undo Then Err.Raise 5, , "Unicode replacement did not create one undo record"
    d.Content.Text = "table host"
    Set table = d.Tables.Add(Range:=d.Paragraphs(1).Range, NumRows:=1, NumColumns:=2)
    table.Cell(1, 1).Range.Text = "Alpha"
    table.Cell(1, 2).Range.Text = "Alpha"
    Set scoped = table.Cell(1, 1).Range.Duplicate
    scoped.End = scoped.End - 1
    result = WU_ReplaceTextInRange(scoped, "Alpha", "Beta", True, True)
    If Not result Or InStr(1, table.Cell(1, 1).Range.Text, "Beta", vbBinaryCompare) = 0 Or InStr(1, table.Cell(1, 2).Range.Text, "Alpha", vbBinaryCompare) = 0 Then Err.Raise 5, , "table-cell range replacement escaped its requested cell"
    If Not d.Undo Then Err.Raise 5, , "table-cell range replacement did not create one undo record"
    If InStr(1, table.Cell(1, 1).Range.Text, "Alpha", vbBinaryCompare) = 0 Then Err.Raise 5, , "table-cell range replacement undo did not restore text"
    d.Content.Text = "Alpha body"
    Set note = d.Footnotes.Add(Range:=d.Paragraphs(1).Range, Text:="Alpha note")
    bodyAfterNote = d.Content.Text
    result = WU_ReplaceText(d, "Alpha", "Omega", "notes", True, True)
    If Not result Or d.Content.Text <> bodyAfterNote Then Err.Raise 5, , "notes-scoped replacement changed the main story"
    If InStr(1, note.Range.Text, "Omega note", vbBinaryCompare) = 0 Then Err.Raise 5, , "notes-scoped replacement did not update the note story"
    If WU_CountText(d, "Omega", "notes", True, True) <> 1 Or WU_CountText(d, "Omega", "main", True, True) <> 0 Then Err.Raise 5, , "scoped counts disagree with the scoped replacement"
    d.Content.Text = "Alpha body"
    Set header = d.Sections(1).Headers(wdHeaderFooterPrimary).Range.Duplicate
    header.Text = "Alpha header" & vbCr
    Set footer = d.Sections(1).Footers(wdHeaderFooterPrimary).Range.Duplicate
    footer.Text = "Alpha footer" & vbCr
    result = WU_ReplaceText(d, "Alpha", "Omega", "headers", True, True)
    If Not result Or d.Content.Text <> "Alpha body" & vbCr Or InStr(1, header.Text, "Omega header", vbBinaryCompare) = 0 Or InStr(1, footer.Text, "Alpha footer", vbBinaryCompare) = 0 Then Err.Raise 5, , "header-scoped replacement escaped its story boundary"
    If Not d.Undo Then Err.Raise 5, , "header-scoped replacement did not create one undo record"
    If InStr(1, header.Text, "Alpha header", vbBinaryCompare) = 0 Then Err.Raise 5, , "header-scoped replacement undo did not restore text"
    result = WU_ReplaceText(d, "Alpha", "Omega", "footers", True, True)
    If Not result Or d.Content.Text <> "Alpha body" & vbCr Or InStr(1, footer.Text, "Omega footer", vbBinaryCompare) = 0 Or InStr(1, header.Text, "Alpha header", vbBinaryCompare) = 0 Then Err.Raise 5, , "footer-scoped replacement escaped its story boundary"
    If Not d.Undo Then Err.Raise 5, , "footer-scoped replacement did not create one undo record"
    If InStr(1, footer.Text, "Alpha footer", vbBinaryCompare) = 0 Then Err.Raise 5, , "footer-scoped replacement undo did not restore text"
    result = WU_ReplaceText(d, "Alpha", "Omega", "all", True, True)
    If Not result Or d.Content.Text <> "Omega body" & vbCr Or InStr(1, header.Text, "Omega header", vbBinaryCompare) = 0 Or InStr(1, footer.Text, "Omega footer", vbBinaryCompare) = 0 Then Err.Raise 5, , "all-story replacement missed a story"
    If Not d.Undo Then Err.Raise 5, , "all-story replacement did not create one undo record"
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
    Dim number As Long, source As String, description As String, changed As Boolean, i As Long, bounded As Range, headerRange As Range
    Dim sourceCharacterStyle As Style, targetCharacterStyle As Style
    Dim styleMap(0 To 1, 0 To 1) As Variant, styleRuns(0 To 1, 0 To 2) As Variant, styleBatchChanged As Long, runChanged As Long
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
    changed = WU_ConvertParagraphStyle(ActiveDocument, "Normal", "MissingProofStyle")
    number = Err.Number
    On Error GoTo 0
    If number = 0 Then Err.Raise 5, , "missing style silently accepted"
    If Application.ScreenUpdating Or Application.UndoRecord.CustomRecordLevel <> 0 Then Err.Raise 5, , "failed conversion lost state"
    ActiveDocument.Styles.Add "ProofTarget", wdStyleTypeParagraph
    ActiveDocument.Content.Style = ActiveDocument.Styles(wdStyleNormal)
    changed = WU_ConvertParagraphStyle(ActiveDocument, ActiveDocument.Styles(wdStyleNormal).NameLocal, "ProofTarget")
    If Not changed Then Err.Raise 5, , "conversion did not report a change"
    If WU_ConvertParagraphStyle(ActiveDocument, "ProofTarget", "ProofTarget") Then Err.Raise 5, , "same style not a no-op"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing conversion undo"
    If ActiveDocument.Paragraphs(1).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "conversion not undone"
    ActiveDocument.Content.Text = "first" & vbCr & "second"
    ActiveDocument.Content.Style = ActiveDocument.Styles(wdStyleNormal)
    Set bounded = ActiveDocument.Paragraphs(2).Range.Duplicate
    changed = WU_ConvertStyleInRange(bounded, ActiveDocument.Styles(wdStyleNormal).NameLocal, "ProofTarget")
    If Not changed Or ActiveDocument.Paragraphs(1).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Or ActiveDocument.Paragraphs(2).Style.NameLocal <> "ProofTarget" Then Err.Raise 5, , "bounded conversion widened its range"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing bounded conversion undo"
    If ActiveDocument.Paragraphs(2).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "bounded conversion not undone"
    ' Character styles: Word resets direct character formatting inside a
    ' restyled run, exactly as its user interface does. The proof asserts that
    ' documented behavior and that one undo restores the original run.
    ActiveDocument.Content.Text = "first cite second"
    Set sourceCharacterStyle = ActiveDocument.Styles.Add("Proof Citation Source", wdStyleTypeCharacter)
    sourceCharacterStyle.Font.Italic = True
    Set targetCharacterStyle = ActiveDocument.Styles.Add("Proof Citation Target", wdStyleTypeCharacter)
    targetCharacterStyle.Font.Bold = True
    Set bounded = ActiveDocument.Paragraphs(1).Range.Duplicate
    bounded.Start = bounded.Start + 6: bounded.End = bounded.Start + 4
    bounded.Style = sourceCharacterStyle
    bounded.Italic = True
    changed = WU_ConvertCharacterStyle(ActiveDocument, "Proof Citation Source", "Proof Citation Target", "main")
    If Not changed Then Err.Raise 5, , "character style conversion did not report a change"
    Set bounded = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(1).Range.Start + 6, End:=ActiveDocument.Paragraphs(1).Range.Start + 10)
    If bounded.Style.NameLocal <> "Proof Citation Target" Or Not bounded.Bold Then Err.Raise 5, , "character style conversion lost the target style"
    If bounded.Italic Then Err.Raise 5, , "Word kept direct italics through a character style change; the documented reset no longer holds"
    If ActiveDocument.Content.Text <> "first cite second" & vbCr Then Err.Raise 5, , "character style conversion changed text: [" & ActiveDocument.Content.Text & "] run style=" & bounded.Style.NameLocal
    If ActiveDocument.Range(0, 5).Bold Then Err.Raise 5, , "character style conversion escaped the styled run"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing character style conversion undo"
    If bounded.Style.NameLocal <> "Proof Citation Source" Or Not bounded.Italic Or bounded.Bold Then Err.Raise 5, , "character style conversion undo did not restore inline formatting"
    ActiveDocument.Content.Text = "direct cite untouched" & vbCr & "outside"
    ActiveDocument.Content.Style = ActiveDocument.Styles(wdStyleNormal)
    Set bounded = ActiveDocument.Paragraphs(1).Range.Duplicate
    bounded.Start = bounded.Start + 7: bounded.End = bounded.Start + 4
    bounded.Style = sourceCharacterStyle
    changed = WU_ApplyCharacterStyleInRange(bounded, "Proof Citation Target")
    If Not changed Or bounded.Style.NameLocal <> "Proof Citation Target" Or Not bounded.Bold Then Err.Raise 5, , "exact character style application did not apply the style"
    If ActiveDocument.Content.Text <> "direct cite untouched" & vbCr & "outside" & vbCr Then Err.Raise 5, , "exact character style application changed text"
    If ActiveDocument.Paragraphs(2).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "exact character style application escaped its range"
    If ActiveDocument.Range(ActiveDocument.Paragraphs(1).Range.Start, ActiveDocument.Paragraphs(1).Range.Start + 6).Bold Then Err.Raise 5, , "exact character style application widened its range"
    If WU_ApplyCharacterStyleInRange(bounded, "Proof Citation Target") Then Err.Raise 5, , "re-applying the current character style is not a no-op"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing exact character style undo"
    If bounded.Style.NameLocal <> "Proof Citation Source" Or bounded.Bold Then Err.Raise 5, , "exact character style undo did not restore source formatting"
    ActiveDocument.Content.Text = "first cite second" & vbCr & "third cite fourth"
    Set bounded = ActiveDocument.Paragraphs(1).Range.Duplicate
    bounded.Start = bounded.Start + 6: bounded.End = bounded.Start + 4
    bounded.Style = sourceCharacterStyle
    Set bounded = ActiveDocument.Paragraphs(2).Range.Duplicate
    bounded.Start = bounded.Start + 6: bounded.End = bounded.Start + 4
    bounded.Style = sourceCharacterStyle
    Set bounded = ActiveDocument.Paragraphs(1).Range.Duplicate
    changed = WU_ConvertStyleInRange(bounded, "Proof Citation Source", "Proof Citation Target")
    If Not changed Then Err.Raise 5, , "bounded character style conversion did not report a change"
    If ActiveDocument.Paragraphs(1).Range.Characters(7).Style.NameLocal <> "Proof Citation Target" Or ActiveDocument.Paragraphs(2).Range.Characters(7).Style.NameLocal <> "Proof Citation Source" Then Err.Raise 5, , "bounded character style conversion escaped its requested range"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing bounded character style conversion undo"
    If ActiveDocument.Paragraphs(1).Range.Characters(7).Style.NameLocal <> "Proof Citation Source" Then Err.Raise 5, , "bounded character style conversion undo did not restore source style"
    styleMap(0, 0) = "Proof Citation Source": styleMap(0, 1) = "Proof Citation Target"
    styleMap(1, 0) = "Proof Citation Target": styleMap(1, 1) = "Proof Citation Target"
    styleBatchChanged = WU_ConvertStyleBatch(ActiveDocument, styleMap, "main")
    If styleBatchChanged <> 1 Or ActiveDocument.Paragraphs(1).Range.Characters(7).Style.NameLocal <> "Proof Citation Target" Or ActiveDocument.Paragraphs(2).Range.Characters(7).Style.NameLocal <> "Proof Citation Target" Then Err.Raise 5, , "character style batch did not restyle both runs in one native pass"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing character style batch undo"
    If ActiveDocument.Paragraphs(1).Range.Characters(7).Style.NameLocal <> "Proof Citation Source" Or ActiveDocument.Paragraphs(2).Range.Characters(7).Style.NameLocal <> "Proof Citation Source" Then Err.Raise 5, , "character style batch undo did not restore source styles"
    ActiveDocument.Content.Text = "first" & vbCr & "second" & vbCr & "third"
    ActiveDocument.Content.Style = ActiveDocument.Styles(wdStyleNormal)
    ActiveDocument.Paragraphs(2).Range.Italic = True
    ActiveDocument.Styles.Add "ProofSecond", wdStyleTypeParagraph
    styleMap(0, 0) = ActiveDocument.Styles(wdStyleNormal).NameLocal: styleMap(0, 1) = "ProofTarget"
    styleMap(1, 0) = "ProofTarget": styleMap(1, 1) = "ProofSecond"
    styleBatchChanged = WU_ConvertStyleBatch(ActiveDocument, styleMap)
    If styleBatchChanged <> 2 Then Err.Raise 5, , "style batch did not report both ordered mappings"
    If ActiveDocument.Paragraphs(1).Style.NameLocal <> "ProofSecond" Or ActiveDocument.Paragraphs(3).Style.NameLocal <> "ProofSecond" Then Err.Raise 5, , "ordered style batch did not reach its final style"
    If Not ActiveDocument.Paragraphs(2).Range.Italic Then Err.Raise 5, , "style batch lost direct italic formatting"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing style batch undo"
    If ActiveDocument.Paragraphs(2).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "style batch undo did not restore the source style"
    ActiveDocument.Content.Text = "first" & vbCr & "second" & vbCr & "third"
    ActiveDocument.Content.Style = ActiveDocument.Styles(wdStyleNormal)
    Set bounded = ActiveDocument.Paragraphs(2).Range.Duplicate
    styleMap(0, 0) = ActiveDocument.Styles(wdStyleNormal).NameLocal: styleMap(0, 1) = "ProofTarget"
    styleMap(1, 0) = "ProofTarget": styleMap(1, 1) = "ProofSecond"
    styleBatchChanged = WU_ConvertStyleBatchInRange(bounded, styleMap)
    If styleBatchChanged <> 2 Or ActiveDocument.Paragraphs(1).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Or ActiveDocument.Paragraphs(2).Style.NameLocal <> "ProofSecond" Or ActiveDocument.Paragraphs(3).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "bounded style batch escaped its requested range"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing bounded style batch undo"
    If ActiveDocument.Paragraphs(2).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "bounded style batch undo did not restore the source style"
    ActiveDocument.Content.Text = "first" & vbCr & "second" & vbCr & "third"
    ActiveDocument.Content.Style = ActiveDocument.Styles(wdStyleNormal)
    ActiveDocument.Paragraphs(2).Range.Italic = True
    styleRuns(0, 0) = ActiveDocument.Paragraphs(1).Range.Start
    styleRuns(0, 1) = ActiveDocument.Paragraphs(2).Range.End
    styleRuns(0, 2) = "ProofTarget"
    styleRuns(1, 0) = ActiveDocument.Paragraphs(3).Range.Start
    styleRuns(1, 1) = ActiveDocument.Paragraphs(3).Range.End
    styleRuns(1, 2) = "ProofSecond"
    runChanged = WU_ApplyParagraphStyleRuns(ActiveDocument.Content, styleRuns)
    If runChanged <> 2 Or ActiveDocument.Paragraphs(1).Style.NameLocal <> "ProofTarget" Or ActiveDocument.Paragraphs(2).Style.NameLocal <> "ProofTarget" Or ActiveDocument.Paragraphs(3).Style.NameLocal <> "ProofSecond" Then Err.Raise 5, , "offset style runs did not apply exact ordered ranges"
    If Not ActiveDocument.Paragraphs(2).Range.Italic Then Err.Raise 5, , "offset style runs lost direct italic formatting"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing offset style runs undo"
    If ActiveDocument.Paragraphs(1).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Or ActiveDocument.Paragraphs(2).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Or ActiveDocument.Paragraphs(3).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "offset style runs undo did not restore source styles"
    ActiveDocument.Saved = True
    styleRuns(1, 2) = "MissingProofStyle"
    On Error Resume Next
    runChanged = WU_ApplyParagraphStyleRuns(ActiveDocument.Content, styleRuns)
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    If number = 0 Or Not ActiveDocument.Saved Then Err.Raise 5, , "invalid offset style run was not rejected before mutation"
    styleRuns(1, 2) = "ProofSecond"
    Set bounded = ActiveDocument.Paragraphs(2).Range.Duplicate
    changed = WU_ApplyParagraphStyleInRange(bounded, "ProofSecond")
    If Not changed Or ActiveDocument.Paragraphs(1).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Or ActiveDocument.Paragraphs(2).Style.NameLocal <> "ProofSecond" Then Err.Raise 5, , "bounded paragraph style application widened its range"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing bounded paragraph style undo"
    If ActiveDocument.Paragraphs(2).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "bounded paragraph style undo did not restore the source style"
    ActiveDocument.Content.Text = "first" & vbCr & "second"
    ActiveDocument.Content.Style = ActiveDocument.Styles(wdStyleNormal)
    ActiveDocument.Saved = True
    styleMap(0, 0) = ActiveDocument.Styles(wdStyleNormal).NameLocal: styleMap(0, 1) = "MissingProofStyle"
    styleMap(1, 0) = "ProofTarget": styleMap(1, 1) = "ProofSecond"
    On Error Resume Next
    styleBatchChanged = WU_ConvertStyleBatch(ActiveDocument, styleMap)
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    If number = 0 Or Not ActiveDocument.Saved Then Err.Raise 5, , "invalid style batch was not rejected before mutation"
    styleMap(0, 0) = ActiveDocument.Styles(wdStyleNormal).NameLocal: styleMap(0, 1) = "Proof Citation Target"
    On Error Resume Next
    styleBatchChanged = WU_ConvertStyleBatch(ActiveDocument, styleMap)
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    If number = 0 Or Not ActiveDocument.Saved Then Err.Raise 5, , "paragraph-to-character mapping was not rejected before mutation"
    Set headerRange = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Duplicate
    headerRange.Text = "header" & vbCr
    headerRange.Style = ActiveDocument.Styles(wdStyleNormal)
    ActiveDocument.Content.Text = "body"
    ActiveDocument.Content.Style = ActiveDocument.Styles(wdStyleNormal)
    changed = WU_ConvertParagraphStyle(ActiveDocument, ActiveDocument.Styles(wdStyleNormal).NameLocal, "ProofTarget", "main")
    If Not changed Or ActiveDocument.Paragraphs(1).Style.NameLocal <> "ProofTarget" Or headerRange.Paragraphs(1).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "main style scope touched a header or missed the body"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing scoped style undo"
    If ActiveDocument.Paragraphs(1).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Or headerRange.Paragraphs(1).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "scoped style undo did not restore both stories"
    changed = WU_ConvertParagraphStyle(ActiveDocument, ActiveDocument.Styles(wdStyleNormal).NameLocal, "ProofTarget", "headers")
    If Not changed Or ActiveDocument.Paragraphs(1).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Or headerRange.Paragraphs(1).Style.NameLocal <> "ProofTarget" Then Err.Raise 5, , "header style scope touched the body or missed the header"
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing header style undo"
    If headerRange.Paragraphs(1).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "header style undo did not restore the source style"
    For i = 1 To 400: lines(i) = "Ordinary paragraph.": Next i
    ActiveDocument.Content.Text = Join(lines, vbCr)
    ActiveDocument.Content.Style = ActiveDocument.Styles(wdStyleNormal)
    started = Timer
    changed = WU_ConvertParagraphStyle(ActiveDocument, ActiveDocument.Styles(wdStyleNormal).NameLocal, "ProofTarget")
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400
    If Not changed Then Err.Raise 5, , "bulk conversion did not report a change"
    ' One native pass over 400 paragraphs runs in roughly 300 ms on a quiet
    ' machine; the guard catches a regression to per-paragraph COM loops
    ' (seconds), not scheduler noise.
    If elapsed * 1000 > 1000 Then Err.Raise 5, , "bulk conversion exceeded 1000 ms: " & CStr(elapsed * 1000)
    If Not ActiveDocument.Undo Then Err.Raise 5, , "missing bulk conversion undo"
    If ActiveDocument.Paragraphs(400).Style.NameLocal <> ActiveDocument.Styles(wdStyleNormal).NameLocal Then Err.Raise 5, , "bulk conversion not undone"
    Application.ScreenUpdating = True
    Check = "PASS"
End Function
Public Function CheckFieldRefresh() As String
    Dim d As Document, note As Footnote, bodyField As Field, noteField As Field, failures As Long, priorUpdating As Boolean
    Set d = Documents.Add
    d.Content.Text = "Body text."
    d.Variables.Add "ProofVar", "one"
    Set bodyField = d.Fields.Add(Range:=d.Range(0, 0), Type:=wdFieldDocVariable, Text:="ProofVar")
    Set note = d.Footnotes.Add(Range:=d.Paragraphs(1).Range)
    Set noteField = d.Fields.Add(Range:=note.Range, Type:=wdFieldDocVariable, Text:="ProofVar")
    If bodyField.Result.Text <> "one" Or noteField.Result.Text <> "one" Then Err.Raise 5, , "fields did not evaluate: " & bodyField.Result.Text & "/" & noteField.Result.Text
    d.Variables("ProofVar").Value = "two"
    priorUpdating = Application.ScreenUpdating
    failures = WU_RefreshFields(d, "main", False)
    If failures <> 0 Or bodyField.Result.Text <> "two" Or noteField.Result.Text <> "one" Then Err.Raise 5, , "main-scope refresh missed the body field or touched the note: " & bodyField.Result.Text & "/" & noteField.Result.Text
    If Application.ScreenUpdating <> priorUpdating Then Err.Raise 5, , "field refresh did not restore ScreenUpdating"
    failures = WU_RefreshFields(d, "notes", False)
    If failures <> 0 Or noteField.Result.Text <> "two" Then Err.Raise 5, , "notes-scope refresh missed the note field"
    If Not d.Undo Then Err.Raise 5, , "field refresh did not create one undo record"
    If noteField.Result.Text <> "one" Then Err.Raise 5, , "field refresh undo did not restore the note field"
    d.Variables("ProofVar").Value = "three"
    failures = WU_RefreshFieldsInRange(d.Paragraphs(1).Range)
    If failures <> 0 Or bodyField.Result.Text <> "three" Then Err.Raise 5, , "range refresh missed the body field"
    d.Saved = True
    If WU_RefreshFields(d, "headers", False) <> 0 Or Not d.Saved Then Err.Raise 5, , "field-free scope opened an edit"
    d.Close SaveChanges:=wdDoNotSaveChanges
    CheckFieldRefresh = "PASS"
End Function
Public Function CheckParagraphIndex() As String
    Dim d As Document, count As Long, starts() As Long, ends() As Long, styles() As String, inTable() As Boolean, texts() As String, tbl As Table, i As Long
    Set d = Documents.Add
    d.Content.Text = "First" & vbCr & "Second" & vbCr & "Table host" & vbCr & "Last"
    d.Content.Style = d.Styles(wdStyleNormal)
    d.Paragraphs(2).Style = d.Styles(wdStyleHeading1)
    Set tbl = d.Tables.Add(Range:=d.Paragraphs(3).Range, NumRows:=1, NumColumns:=2)
    d.Saved = True
    count = WU_IndexParagraphs(d, starts, ends, styles, inTable, texts)
    If count <> d.Paragraphs.Count Or Not d.Saved Then Err.Raise 5, , "index count or document state changed"
    For i = 1 To count
        If starts(i) <> d.Paragraphs(i).Range.Start Or ends(i) <> d.Paragraphs(i).Range.End Then Err.Raise 5, , "offset mismatch at paragraph " & CStr(i)
        If texts(i) <> d.Paragraphs(i).Range.Text Then Err.Raise 5, , "text mismatch at paragraph " & CStr(i)
        If styles(i) <> d.Paragraphs(i).Style.NameLocal Then Err.Raise 5, , "style mismatch at paragraph " & CStr(i)
        If inTable(i) <> (d.Paragraphs(i).Range.Tables.Count > 0) Then Err.Raise 5, , "table membership mismatch at paragraph " & CStr(i)
    Next i
    If styles(2) <> d.Styles(wdStyleHeading1).NameLocal Then Err.Raise 5, , "heading style not indexed"
    d.Close SaveChanges:=wdDoNotSaveChanges
    CheckParagraphIndex = "PASS"
End Function
Public Function CheckBatch() As String
    Dim report As Variant, started As Single
    ProofProgress_BatchStart 2
    If Not ProofProgress_BatchActive Then Err.Raise 5, , "batch did not start"
    ProofProgress_Notify "quiet", vbInformation
    ProofProgress_BatchStage "first stage"
    started = Timer
    Do While Timer - started < 0.02
    Loop
    ProofProgress_BatchStage "second stage"
    ProofProgress_Notify "problem", vbExclamation
    ProofProgress_BatchEnd
    If ProofProgress_BatchActive Then Err.Raise 5, , "batch did not end"
    report = ProofProgress_BatchReport
    If report(0) <> True Then Err.Raise 5, , "exclamation notification did not mark failure"
    If InStr(report(1), "stage_ms" & vbTab & "first stage" & vbTab) = 0 Or InStr(report(1), "stage_ms" & vbTab & "second stage" & vbTab) = 0 Then Err.Raise 5, , "stage timings missing: " & report(1)
    If InStr(report(1), "message" & vbTab & "quiet") = 0 Then Err.Raise 5, , "notification was not captured"
    If report(2) <> 2 Or report(3) <> 2 Then Err.Raise 5, , "step counters wrong"
    CheckBatch = "PASS"
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
