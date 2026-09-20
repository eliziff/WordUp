package verify

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"testing"

	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

// Synthetic bookkeeping only; the Windows test separately executes Word.
func TestCorpusAttemptsAndResume(t *testing.T) {
	root := t.TempDir()
	artifact := filepath.Join(root, "fixture.dotm")
	source := filepath.Join(root, "input.docx")
	for path, text := range map[string]string{artifact: "artifact", source: "input"} {
		if err := os.WriteFile(path, []byte(text), 0600); err != nil {
			t.Fatal(err)
		}
	}
	suite := Suite{Schema: 1, Name: "generated", Inputs: map[string]string{"document": source}, Steps: []Step{{Name: "assert", Operation: native.Operation{Op: "eval", Value: "Evaluate = True"}, Assert: []Assertion{{Kind: "equals", Expected: true}}}}}
	cases := []CorpusCase{{ID: "First", Inputs: suite.Inputs}, {ID: "Second", Inputs: suite.Inputs}}
	calls := 0
	fail := true
	run := func(ctx context.Context, path string, s Suite) (*Report, error) {
		calls++
		if fail && calls == 2 {
			return nil, fmt.Errorf("deliberate second-case failure")
		}
		dir := filepath.Join(root, fmt.Sprint("evidence-", calls))
		inputs, err := readInputs(s.Inputs, nil)
		if err != nil {
			return nil, err
		}
		snapshots, err := stageInputs(dir, inputs, map[string]string{})
		if err != nil {
			return nil, err
		}
		b, err := os.ReadFile(path)
		if err != nil {
			return nil, err
		}
		frozen := filepath.Join(dir, "artifact.dotm")
		if err := os.WriteFile(frozen, b, 0600); err != nil {
			return nil, err
		}
		r := &Report{Schema: 1, ToolVersion: project.Version, Status: "passed", WordExecuted: true, OS: "synthetic", ArtifactSnapshot: frozen, EvidenceDirectory: dir, SHA256: office.Hash(b), Suite: s, SuiteSHA256: office.Hash(project.JSON(s)), Assertions: 1, InputSnapshots: snapshots, Observations: []Observation{{Name: "assert", Passed: true, Assertions: 1, Result: true}}}
		return r, SaveReport(root, r)
	}
	output := filepath.Join(root, "runs")
	first, err := RunCorpus(context.Background(), artifact, suite, cases, output, "", run)
	if err == nil || first.Status != "failed" || calls != 2 || first.Cases[0].Status != "passed" {
		t.Fatal(first, err, calls)
	}
	original, _ := os.ReadFile(first.SavedReport)
	fail = false
	second, err := RunCorpus(context.Background(), artifact, suite, cases, output, first.SavedReport, run)
	if err != nil || second.Status != "passed" || calls != 3 || !second.Cases[0].Reused || second.Cases[1].Reused {
		t.Fatal(second, err, calls)
	}
	retained, _ := os.ReadFile(first.SavedReport)
	if string(retained) != string(original) || second.SavedReport == first.SavedReport {
		t.Fatal("retry overwrote attempt")
	}
	// Damage frozen output: resume must rerun only that case, never trust status.
	loaded, _, err := LoadReport(second.Cases[0].Bundle)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(loaded.ArtifactSnapshot, []byte("tampered"), 0600); err != nil {
		t.Fatal(err)
	}
	third, err := RunCorpus(context.Background(), artifact, suite, cases, output, second.SavedReport, run)
	if err != nil || calls != 4 || third.Cases[0].Reused || !third.Cases[1].Reused {
		t.Fatal(third, err, calls)
	}
	// Changed source invalidates both matching IDs.
	if err := os.WriteFile(source, []byte("changed"), 0600); err != nil {
		t.Fatal(err)
	}
	fourth, err := RunCorpus(context.Background(), artifact, suite, cases, output, third.SavedReport, run)
	if err != nil || calls != 6 || fourth.Cases[0].Reused || fourth.Cases[1].Reused {
		t.Fatal(fourth, err, calls)
	}
	// Changes during execution are rejected, even if the native assertion passed.
	mutate := func(ctx context.Context, path string, s Suite) (*Report, error) {
		if err := os.WriteFile(source, []byte("changed during run"), 0600); err != nil {
			t.Fatal(err)
		}
		return run(ctx, path, s)
	}
	changed, err := RunCorpus(context.Background(), artifact, suite, cases[:1], output, "", mutate)
	if err == nil || changed.Cases[0].Status != "failed" {
		t.Fatal("input race accepted", changed, err)
	}
	before := calls
	bad := []CorpusCase{cases[0], {ID: "first", Inputs: suite.Inputs}}
	if _, err := RunCorpus(context.Background(), artifact, suite, bad, output, "", run); err == nil || calls != before {
		t.Fatal("case collision accepted")
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	canceled, err := RunCorpus(ctx, artifact, suite, cases, output, "", run)
	if err == nil || canceled.Status != "incomplete" || calls != before {
		t.Fatal("cancelled corpus accepted", canceled, err)
	}
}
