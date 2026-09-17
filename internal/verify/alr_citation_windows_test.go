//go:build windows

package verify_test

import (
	"context"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap/citations"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/verify"
)

// The private ALR workspace is not part of the repository; the proof runs
// only where it exists. It seeds footnotes with the pinpoint forms the
// StylePass citation pass rewrites, runs the pass in fix mode, accepts the
// tracked changes and compares every note with the ported code map's own
// normalization where the oracle parses the note, and with the guide's
// expected form otherwise.
const alrCitationSeed = `Dim d As Document, rng As Range, seeds As Variant, i As Long
Set d = ActiveDocument
seeds = Array("Ibid at pp. 400-408.", "Ibid at para. 12.", "Ibid at paras 12.", "Ibid at para 12-14.", "Ibid at p 5.", "Ibid at 1000 - 1010.", "Ibid at paras 400-408.", "Ibid at 12.", "Ibid at 12 (emphasis added).")
d.Content.Text = "Body one." & vbCr & "Body two." & vbCr
For i = LBound(seeds) To UBound(seeds)
Set rng = d.Paragraphs(1).Range.Duplicate
rng.Collapse wdCollapseEnd
rng.Move wdCharacter, -1
d.Footnotes.Add rng, , CStr(seeds(i))
Next i
Evaluate = d.Footnotes.Count`

const alrCitationRead = `Dim d As Document, note As Footnote, out As String
Set d = ActiveDocument
d.AcceptAllRevisions
For Each note In d.Footnotes
out = out & Replace(note.Range.Text, vbCr, "") & "|"
Next note
d.Close SaveChanges:=wdDoNotSaveChanges
Evaluate = out`

func TestNativeALRCitationPassMatchesOracle(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	artifact := filepath.Join("..", "..", "workspaces", "alr-refactor", "dist", "ALR_V1.dotm")
	if _, err := os.Stat(artifact); err != nil {
		t.Skip("private ALR workspace not present")
	}
	suite := verify.Suite{Schema: 1, Name: "ALR citation pinpoints", RequireCompile: true, Steps: []verify.Step{
		{Name: "Open template", Operation: native.Operation{Op: "open", File: "$artifact", As: "template"}},
		{Name: "Compile", Operation: native.Operation{Op: "compile", Target: "template", Member: "$project"}, Assert: []verify.Assertion{{Path: "/vba_compiled", Kind: "equals", Expected: true}}},
		{Name: "Create document from the template", Operation: native.Operation{Op: "new", File: "$artifact", As: "document"}},
		{Name: "Seed footnotes", Operation: native.Operation{Op: "eval", Value: alrCitationSeed}, Assert: []verify.Assertion{{Path: "/result", Kind: "equals", Expected: 9}}},
		{Name: "Run citation pass", Operation: native.Operation{Op: "run", Macro: "ALR_CitationNormalizeAll", TimeoutMS: 120000}},
		{Name: "Read notes", Operation: native.Operation{Op: "eval", Value: alrCitationRead}, Assert: []verify.Assertion{{Path: "/result", Kind: "contains", Expected: "|"}}},
	}}
	rep, err := verify.Run(context.Background(), artifact, suite, nil, true)
	if err != nil {
		t.Fatalf("native citation pass failed: %v", err)
	}
	var got []string
	for _, obs := range rep.Observations {
		if obs.Name == "Read notes" {
			if m, ok := obs.Result.(map[string]any); ok {
				if s, ok := m["result"].(string); ok {
					got = strings.Split(strings.TrimSuffix(s, "|"), "|")
				}
			}
		}
	}
	// Footnotes are added at the same anchor, so Word numbers them in
	// reverse insertion order; sort both sides by seed text to compare.
	seeds := []string{"Ibid at pp. 400-408.", "Ibid at para. 12.", "Ibid at paras 12.", "Ibid at para 12-14.", "Ibid at p 5.", "Ibid at 1000 - 1010.", "Ibid at paras 400-408.", "Ibid at 12.", "Ibid at 12 (emphasis added)."}
	want := map[string]string{
		"Ibid at pp. 400-408.":         "Ibid at 400–08.",
		"Ibid at para. 12.":            "Ibid at para 12.",
		"Ibid at paras 12.":            "Ibid at para 12.",
		"Ibid at para 12-14.":          "Ibid at paras 12–14.",
		"Ibid at p 5.":                 "Ibid at 5.",
		"Ibid at 1000 - 1010.":         "Ibid at 1000–10.",
		"Ibid at paras 400-408.":       "Ibid at paras 400–408.",
		"Ibid at 12.":                  "Ibid at 12.",
		"Ibid at 12 (emphasis added).": "Ibid at 12 [emphasis added].",
	}
	// Where the oracle parses the seed, its normalization is the expectation.
	for _, seed := range seeds {
		if res, err := citations.NormalizeNote(seed, nil, citations.Options{}); err == nil {
			want[seed] = res.Text
		}
	}
	if len(got) != len(seeds) {
		t.Fatalf("expected %d notes, got %d: %q", len(seeds), len(got), got)
	}
	// Two seeds normalize to the same form, so count expectations.
	remaining := map[string]int{}
	for _, w := range want {
		remaining[w]++
	}
	for _, g := range got {
		g = strings.TrimLeft(g, "\t ")
		if remaining[g] == 0 {
			t.Errorf("note %q is not an expected form (%v)", g, want)
			continue
		}
		remaining[g]--
	}
	for w, n := range remaining {
		if n > 0 {
			t.Errorf("expected form never produced: %q", w)
		}
	}
}
