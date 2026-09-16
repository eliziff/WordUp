//go:build windows && (amd64 || arm64)

package verify_test

import (
	"context"
	"os"
	"path/filepath"
	"testing"

	"github.com/eliziff/WordUp/internal/journal"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/vbaparse"
	"github.com/eliziff/WordUp/internal/verify"
)

func TestJournalStyleProofParses(t *testing.T) {
	parsed, err := vbaparse.Parse(journalStyleProof)
	if err != nil {
		t.Fatal(err)
	}
	if diagnostics := parsed["diagnostics"].([]vbaparse.Diagnostic); len(diagnostics) != 0 {
		t.Fatalf("journal style proof diagnostics: %#v", diagnostics)
	}
}

func TestNativeNeutralJournalStylePass(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	root := filepath.Join(t.TempDir(), "journal")
	profile, err := journal.Get("UBC-L-REV")
	if err != nil {
		t.Fatal(err)
	}
	report, err := journal.Create(root, profile)
	if err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "vba/JournalStyleProof.bas", []byte(journalStyleProof), ""); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := w.Build(report.Artifact); err != nil {
		t.Fatal(err)
	}
	suite := verify.Suite{Schema: 1, Name: "Neutral journal style pass", RequireCompile: true, Steps: []verify.Step{
		{Name: "Open", Operation: native.Operation{Op: "open", File: "$artifact", As: "doc"}},
		{Name: "Compile", Operation: native.Operation{Op: "compile", Target: "doc", Member: "$project"}, Assert: []verify.Assertion{{Path: "/vba_compiled", Kind: "equals", Expected: true}}},
		{Name: "Preserve list and direct formatting", Operation: native.Operation{Op: "run", Macro: "JournalStyleProof.Check"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "PASS"}}},
	}}
	if _, err := verify.Run(context.Background(), report.Artifact, suite, nil, true); err != nil {
		t.Fatalf("native neutral journal style acceptance failed: %v", err)
	}
}

const journalStyleProof = `Attribute VB_Name = "JournalStyleProof"
Option Explicit

Public Function Check() As String
    Dim d As Document, roleDoc As Document, before As String, started As Single, elapsed As Single, italicRange As Range, directRange As Range
    Dim proofTitle As Style, proofAuthor As Style, proofQuote As Style
    Set d = Documents.Add
    d.Content.Text = "ordinary list item" & vbCr & "second list item" & vbCr & "Introduction" & vbCr & "Body with emphasis" & vbCr
    d.Content.Style = wdStyleNormal
    d.Paragraphs(1).Range.ListFormat.ApplyNumberDefault
    d.Paragraphs(2).Range.ListFormat.ApplyNumberDefault
    d.Paragraphs(3).OutlineLevel = wdOutlineLevel1
    Set italicRange = d.Paragraphs(4).Range.Duplicate
    italicRange.End = italicRange.Start + 4
    italicRange.Italic = True
    before = d.Content.Text
    started = Timer
    WU_JournalApplyStyles
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400
    If d.Paragraphs(1).Style.NameLocal = WU_JOURNAL_STYLE_H1 Or d.Paragraphs(2).Style.NameLocal = WU_JOURNAL_STYLE_H1 Then Err.Raise 5, , "numbered list was promoted to a heading"
    If d.Paragraphs(3).Style.NameLocal <> WU_JOURNAL_STYLE_H1 Then Err.Raise 5, , "native outline heading was not styled"
    If d.Paragraphs(4).Style.NameLocal <> WU_JOURNAL_STYLE_BODY Then Err.Raise 5, , "body paragraph was not styled"
    If Not d.Paragraphs(4).Range.Characters(1).Italic Then Err.Raise 5, , "direct italic formatting was lost"
    If elapsed * 1000 > 250 Then Err.Raise 5, , "neutral style pass exceeded 250 ms: " & CStr(elapsed * 1000)
    If Not d.Undo Then Err.Raise 5, , "style pass did not create one undo record"
    If d.Content.Text <> before Then Err.Raise 5, , "style pass undo did not restore text"
    d.Close SaveChanges:=wdDoNotSaveChanges
    ' The shared detector exposes generic roles, and the neutral core maps
    ' those roles to named editable styles without rewriting paragraph text or
    ' direct inline formatting. Use explicit style-name evidence here so the
    ' proof does not depend on a publisher's private style vocabulary.
    Set roleDoc = Documents.Add
    roleDoc.Content.Text = "A Neutral Article" & vbCr & "By Jane Doe" & vbCr & "Abstract" & vbCr & "Abstract text with emphasis." & vbCr & "Quoted text." & vbCr & "1. Heading" & vbCr & "Body text." & vbCr
    Set proofTitle = roleDoc.Styles.Add(Name:="Proof Title", Type:=wdStyleTypeParagraph)
    Set proofAuthor = roleDoc.Styles.Add(Name:="Proof Author", Type:=wdStyleTypeParagraph)
    Set proofQuote = roleDoc.Styles.Add(Name:="Proof Quote", Type:=wdStyleTypeParagraph)
    roleDoc.Paragraphs(1).Style = proofTitle: roleDoc.Paragraphs(1).Alignment = wdAlignParagraphCenter
    roleDoc.Paragraphs(2).Style = proofAuthor: roleDoc.Paragraphs(2).Alignment = wdAlignParagraphCenter
    roleDoc.Paragraphs(5).Style = proofQuote
    Set directRange = roleDoc.Paragraphs(4).Range.Duplicate
    directRange.End = directRange.Start + 8
    directRange.Italic = True
    WU_JournalApplyStyles
    If roleDoc.Paragraphs(1).Style.NameLocal <> WU_JOURNAL_STYLE_TITLE Then Err.Raise 5, , "title role was not mapped to the neutral title style"
    If roleDoc.Paragraphs(2).Style.NameLocal <> WU_JOURNAL_STYLE_AUTHOR Then Err.Raise 5, , "author role was not mapped to the neutral author style"
    If roleDoc.Paragraphs(3).Style.NameLocal <> WU_JOURNAL_STYLE_ABSTRACT Then Err.Raise 5, , "abstract role was not mapped to the neutral abstract style"
    If roleDoc.Paragraphs(5).Style.NameLocal <> WU_JOURNAL_STYLE_QUOTATION Then Err.Raise 5, , "quotation role was not mapped to the neutral quotation style"
    If roleDoc.Paragraphs(6).Style.NameLocal <> WU_JOURNAL_STYLE_H1 Then Err.Raise 5, , "numbered heading role was not mapped to heading style"
    If roleDoc.Paragraphs(7).Style.NameLocal <> WU_JOURNAL_STYLE_BODY Then Err.Raise 5, , "role-mapped body paragraph was not styled"
    If Not roleDoc.Paragraphs(4).Range.Characters(1).Italic Then Err.Raise 5, , "role style pass lost direct inline emphasis"
    If roleDoc.Styles(WU_JOURNAL_STYLE_CITATION).Type <> wdStyleTypeCharacter Then Err.Raise 5, , "neutral citation character style was not provisioned"
    If Not roleDoc.Undo Then Err.Raise 5, , "role style pass did not create one undo record"
    roleDoc.Close SaveChanges:=wdDoNotSaveChanges
    Check = "PASS"
End Function
`
