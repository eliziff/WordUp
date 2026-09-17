//go:build windows

package verify_test

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/eliziff/WordUp/internal/journal"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
)

// The typeset proof runs every profile-driven stage of a generated journal
// template on a synthetic manuscript and checks the document against the
// template's own constants, so the same proof holds for all 30 journals.
const journalTypesetProof = `Attribute VB_Name = "JournalTypesetProof"
Option Explicit

Public Function Check() As String
    Dim d As Document, before As String, report As Variant, note As Footnote, rng As Range
    Dim updating As Boolean, opened As Boolean, captured As Boolean
    Dim widthBefore As Single, highlighted As Long, headerText As String, i As Long
    Dim stage As String
    On Error GoTo Failed
    stage = "fixture"
    Set d = Documents.Add
    d.Content.Text = "A Synthetic Article Title" & vbCr & "By Jane Doe" & vbCr & "Introduction" & vbCr & _
        "Body text with ""straight quotes"" and 'single quotes' in it." & vbCr & "Second paragraph of body text." & vbCr
    d.Content.Style = wdStyleNormal
    d.Paragraphs(3).OutlineLevel = wdOutlineLevel1
    Set rng = d.Paragraphs(4).Range.Duplicate
    rng.Collapse wdCollapseEnd
    rng.Move wdCharacter, -1
    d.Footnotes.Add rng, , "See Smith et al., ibid at para. 12; supra, note 3; (emphasis added)."
    Set rng = d.Paragraphs(5).Range.Duplicate
    rng.Collapse wdCollapseEnd
    rng.Move wdCharacter, -1
    d.Footnotes.Add rng, , "Ibid at 12-15, s. 7."
    d.BuiltInDocumentProperties("Title").Value = "A Synthetic Article Title"
    widthBefore = d.Sections(1).PageSetup.PageWidth
    before = d.Content.Text
    WU_BatchStart 7
    WU_BeginSafeEdit updating, opened, captured, "Typeset proof"
    ' Same order as the form and WU_Typeset: footnote tabs before the style
    ' pass and running heads last, the two sequences that keep Word's custom
    ' undo record whole (see typesetStages).
    WU_BatchStage "Page size and margins": stage = "page": WU_TypesetPageSetup
    WU_BatchStage "Footnote numbers": stage = "notes": WU_TypesetFootnoteNumbers
    WU_BatchStage "House styles": stage = "styles": WU_TypesetStyles
    WU_BatchStage "Curly quotes": stage = "quotes": WU_TypesetQuotes
    WU_BatchStage "Citation audit": stage = "audit": WU_TypesetCitationAudit
    WU_BatchStage "Citation fixes": stage = "fix": WU_TypesetCitationFix
    WU_BatchStage "Running heads": stage = "heads": WU_TypesetRunningHeads
    WU_EndSafeEdit updating, opened, captured
    WU_BatchEnd
    report = WU_BatchReport()
    stage = "assertions"
    If report(0) Then Err.Raise 5, , "a typeset stage reported failure: " & report(1)
    ' Page setup follows the profile when it carries evidence.
    If WU_TS_PAGE_WIDTH_PT > 0 Then
        If Abs(d.Sections(1).PageSetup.PageWidth - WU_TS_PAGE_WIDTH_PT) > 0.5 Or Abs(d.Sections(1).PageSetup.PageHeight - WU_TS_PAGE_HEIGHT_PT) > 0.5 Then Err.Raise 5, , "page size was not applied: " & d.Sections(1).PageSetup.PageWidth & "x" & d.Sections(1).PageSetup.PageHeight
    End If
    If WU_TS_MARGIN_TOP_IN > 0 Then If Abs(d.Sections(1).PageSetup.TopMargin - InchesToPoints(WU_TS_MARGIN_TOP_IN)) > 0.5 Then Err.Raise 5, , "top margin was not applied"
    If WU_TS_MIRROR_MARGINS Then
        If d.Sections(1).PageSetup.MirrorMargins <> True Then Err.Raise 5, , "mirror margins were not applied"
    ElseIf WU_TS_MARGIN_LEFT_IN > 0 Then
        If Abs(d.Sections(1).PageSetup.LeftMargin - InchesToPoints(WU_TS_MARGIN_LEFT_IN)) > 0.5 Then Err.Raise 5, , "left margin was not applied"
    End If
    ' Styles: body leading, heading case and alignment, title alignment.
    If d.Paragraphs(4).Style.NameLocal <> WU_JOURNAL_STYLE_BODY Then Err.Raise 5, , "body paragraph was not styled: " & d.Paragraphs(4).Style.NameLocal
    If d.Paragraphs(3).Style.NameLocal <> WU_JOURNAL_STYLE_H1 Then Err.Raise 5, , "heading was not styled: " & d.Paragraphs(3).Style.NameLocal
    If WU_TS_BODY_LEADING_PT > 0 Then
        If d.Styles(WU_JOURNAL_STYLE_BODY).ParagraphFormat.LineSpacingRule <> wdLineSpaceExactly Or Abs(d.Styles(WU_JOURNAL_STYLE_BODY).ParagraphFormat.LineSpacing - WU_TS_BODY_LEADING_PT) > 0.1 Then Err.Raise 5, , "body leading was not applied"
    End If
    If WU_TS_HEADING_CASE = "upper" Then If d.Styles(WU_JOURNAL_STYLE_H1).Font.AllCaps <> True Then Err.Raise 5, , "heading all-caps was not applied"
    If WU_TS_HEADING_ALIGN = "center" Then If d.Styles(WU_JOURNAL_STYLE_H1).ParagraphFormat.Alignment <> wdAlignParagraphCenter Then Err.Raise 5, , "heading alignment was not applied"
    If WU_TS_QUOTE_INDENT_IN > 0 Then If Abs(d.Styles(WU_JOURNAL_STYLE_QUOTATION).ParagraphFormat.LeftIndent - InchesToPoints(WU_TS_QUOTE_INDENT_IN)) > 0.5 Then Err.Raise 5, , "block quote indent was not applied"
    ' Running heads carry the journal name or title when the profile has them.
    If Len(WU_TS_ODD_HEAD) > 0 Then
        ' The head holds whatever the journal prints there: its name, the
        ' article title or author, or only a volume label and page number.
        headerText = d.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text
        If Len(Trim$(Replace(Replace(headerText, vbCr, ""), vbTab, ""))) = 0 And d.Sections(1).Headers(wdHeaderFooterPrimary).Range.Fields.Count = 0 Then Err.Raise 5, , "running head was not written: [" & headerText & "]"
        If Not WU_HeadIsOwned(d.Sections(1).Headers(wdHeaderFooterPrimary).Range) Then Err.Raise 5, , "running head is not marked as template-owned"
    End If
    If Left$(WU_TS_PAGE_NUMBER, 3) = "top" And Len(WU_TS_ODD_HEAD) > 0 Then
        If d.Sections(1).Headers(wdHeaderFooterPrimary).Range.Fields.Count = 0 Then Err.Raise 5, , "page number field was not added to the header"
    End If
    ' Footnote numbers: a tab follows the mark when the journal prints a gap.
    If WU_TS_NOTE_NUMBER_STYLE = "number_gap" Or WU_TS_NOTE_NUMBER_STYLE = "number_space" Then
        For Each note In d.Footnotes
            If Left$(note.Range.Text, 1) <> vbTab Then Err.Raise 5, , "footnote does not start with a tab: [" & Left$(note.Range.Text, 12) & "]"
        Next note
    End If
    ' Quotes: straight quotes are gone from the body when the journal curls them.
    If WU_TS_QUOTES = "curly" Then
        If InStr(1, d.Paragraphs(4).Range.Text, """", vbBinaryCompare) > 0 Then Err.Raise 5, , "straight double quotes survived: " & d.Paragraphs(4).Range.Text
        If InStr(1, d.Paragraphs(4).Range.Text, ChrW(8220) & "straight quotes" & ChrW(8221), vbBinaryCompare) = 0 Then Err.Raise 5, , "curly quotes are not oriented: " & d.Paragraphs(4).Range.Text
    End If
    ' Citation audit highlighted at least one seeded deviation when rules exist.
    If WU_TypesetAuditRuleCount() > 0 Then
        For Each note In d.Footnotes
            Set rng = note.Range
            For i = 1 To rng.Characters.Count
                If rng.Characters(i).HighlightColorIndex = wdYellow Then highlighted = highlighted + 1: Exit For
            Next i
        Next note
        If highlighted = 0 Then Err.Raise 5, , "citation audit highlighted nothing although rules exist"
    End If
    ' Fixes: Ibid case follows the journal when measured.
    If InStr(1, WU_TypesetFixTargets(), "|ibid>Ibid", vbBinaryCompare) > 0 Then
        If InStr(1, d.Footnotes(1).Range.Text, "Ibid", vbBinaryCompare) = 0 Then Err.Raise 5, , "ibid case was not fixed: " & d.Footnotes(1).Range.Text
    End If
    ' One undo reverses everything the batch did. Word's Document.Undo
    ' reports False for a custom record that spans header stories although
    ' it reverts the document, so the proof checks the restored state.
    stage = "undo"
    d.Undo
    If Abs(d.Sections(1).PageSetup.PageWidth - widthBefore) > 0.5 Then Err.Raise 5, , "one undo did not restore the page width"
    If d.Content.Text <> before Then Err.Raise 5, , "one undo did not restore the text"
    headerText = d.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text
    If Len(headerText) > 1 Then Err.Raise 5, , "one undo did not clear the running head: [" & headerText & "]"
    If Left$(d.Footnotes(1).Range.Text, 1) = vbTab Then Err.Raise 5, , "one undo did not remove the footnote tab"
    Check = "PASS " & report(1)
    d.Close SaveChanges:=wdDoNotSaveChanges
    Exit Function
Failed:
    Check = "FAIL stage=" & stage & " err=" & CStr(Err.Number) & " " & Err.Description
    On Error Resume Next
    WU_EndSafeEdit updating, opened, captured
    WU_BatchEnd
    d.Close SaveChanges:=wdDoNotSaveChanges
End Function
`

func typesetSuite(name string) verify.Suite {
	return verify.Suite{Schema: 1, Name: name, RequireCompile: true, Steps: []verify.Step{
		{Name: "Open", Operation: native.Operation{Op: "open", File: "$artifact", As: "doc"}},
		{Name: "Compile", Operation: native.Operation{Op: "compile", Target: "doc", Member: "$project"}, Assert: []verify.Assertion{{Path: "/vba_compiled", Kind: "equals", Expected: true}}},
		{Name: "Typeset stages", Operation: native.Operation{Op: "run", Macro: "JournalTypesetProof.Check", TimeoutMS: 120000}, Assert: []verify.Assertion{{Kind: "contains", Expected: "PASS"}}},
	}}
}

func buildJournalWithProof(t *testing.T, root string, profile journal.Profile) string {
	t.Helper()
	report, err := journal.Create(root, profile)
	if err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "vba/JournalTypesetProof.bas", []byte(journalTypesetProof), ""); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := w.Build(report.Artifact); err != nil {
		t.Fatal(err)
	}
	return report.Artifact
}

func TestNativeJournalTypeset(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	profile, err := journal.Get("UBC-L-REV")
	if err != nil {
		t.Fatal(err)
	}
	if profile.Layout == nil || profile.Layout.PageWidthPT == 0 {
		t.Fatalf("UBC-L-REV profile carries no inferred layout: %+v", profile.EvidenceGaps)
	}
	artifact := buildJournalWithProof(t, filepath.Join(t.TempDir(), "journal"), profile)
	rep, err := verify.Run(context.Background(), artifact, typesetSuite("UBC-L-REV typeset"), nil, true)
	if err != nil {
		t.Fatalf("native journal typeset acceptance failed: %v", err)
	}
	t.Logf("typeset acceptance: %v ms", rep.DurationMS)
}

// TestNativeJournalTypesetAll runs the same proof for every journal with a
// bundled profile and writes a summary table; opt in with WORDUP_JOURNAL_ALL=1.
func TestNativeJournalTypesetAll(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" || os.Getenv("WORDUP_JOURNAL_ALL") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1 and WORDUP_JOURNAL_ALL=1")
	}
	var rows []string
	failures := 0
	for _, id := range journal.InferredIDs() {
		profile, err := journal.Get(id)
		if err != nil {
			t.Errorf("%s: %v", id, err)
			continue
		}
		root := filepath.Join(t.TempDir(), id)
		artifact := buildJournalWithProof(t, root, profile)
		started := time.Now()
		rep, err := verify.Run(context.Background(), artifact, typesetSuite(id+" typeset"), nil, true)
		status := "PASS"
		detail := ""
		if err != nil {
			status = "FAIL"
			detail = err.Error()
			failures++
		} else {
			for _, obs := range rep.Observations {
				if obs.Name == "Typeset stages" {
					if s, ok := obs.Result.(string); ok {
						detail = strings.ReplaceAll(strings.TrimPrefix(s, "PASS "), "\r\n", "; ")
					}
				}
			}
		}
		gaps := strings.Join(profile.EvidenceGaps, ", ")
		rows = append(rows, fmt.Sprintf("| %s | %s | %.1f s | %s | %s |", id, status, time.Since(started).Seconds(), gaps, strings.ReplaceAll(detail, "|", "/")))
		t.Logf("%s: %s (%s)", id, status, detail)
	}
	out := os.Getenv("WORDUP_JOURNAL_REPORT")
	if out == "" {
		out = filepath.Join("..", "..", "evidence", "journals", time.Now().Format("20060102")+"-typeset.md")
	}
	body := "# Journal typeset proofs\n\nGenerated by TestNativeJournalTypesetAll in Word; one synthetic manuscript per journal.\n\n| journal | status | wall | evidence gaps | stage timings |\n|---|---|---:|---|---|\n" + strings.Join(rows, "\n") + "\n"
	if err := os.MkdirAll(filepath.Dir(out), 0o755); err == nil {
		_ = os.WriteFile(out, []byte(body), 0o644)
	}
	if failures > 0 {
		t.Fatalf("%d journal typeset proof(s) failed; see %s", failures, out)
	}
}
