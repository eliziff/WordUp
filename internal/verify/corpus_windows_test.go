//go:build windows && (amd64 || arm64)

package verify_test

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"testing"

	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
)

func TestNativeCorpus(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	root := t.TempDir()
	workspace := filepath.Join(root, "workspace")
	if _, err := project.New("CorpusFixture", workspace); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(workspace)
	if err != nil {
		t.Fatal(err)
	}
	built, err := w.Build("")
	if err != nil {
		t.Fatal(err)
	}
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:r><w:t>Original</w:t></w:r></w:p></w:body></w:document>`)
	data, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	var cases []verify.CorpusCase
	for _, id := range []string{"First", "Second"} {
		path := filepath.Join(root, id+".docx")
		if err := os.WriteFile(path, data, 0600); err != nil {
			t.Fatal(err)
		}
		cases = append(cases, verify.CorpusCase{ID: id, Inputs: map[string]string{"document": path}})
	}
	suite := verify.Suite{Schema: 1, Name: "Generated corpus", Inputs: cases[0].Inputs, Steps: []verify.Step{
		{Name: "Open", Operation: native.Operation{Op: "open", File: "$input:document$", As: "doc"}},
		{Name: "Range", Operation: native.Operation{Op: "get", Target: "doc", Member: "Content", As: "body"}},
		{Name: "Original", Operation: native.Operation{Op: "get", Target: "body", Member: "Text"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "Original\r"}}},
		{Name: "Edit", Operation: native.Operation{Op: "put", Target: "body", Member: "Text", Value: "Edited"}},
		{Name: "Save", Operation: native.Operation{Op: "invoke", Target: "doc", Member: "SaveAs2", Args: []any{"$output/output.docx"}}},
		{Name: "Release", Operation: native.Operation{Op: "release", Target: "body"}},
		{Name: "Close", Operation: native.Operation{Op: "unload", Target: "doc"}},
	}}
	ctx := context.Background()
	h, err := native.Start(ctx, native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	calls := 0
	run := func(ctx context.Context, artifact string, s verify.Suite) (*verify.Report, error) {
		calls++
		r, err := verify.Run(ctx, artifact, s, h, true)
		if r != nil {
			err = errors.Join(err, verify.SaveReport(workspace, r))
		}
		return r, err
	}
	first, err := verify.RunCorpus(ctx, built.Artifact, suite, cases, filepath.Join(root, "runs"), "", run)
	if err != nil {
		t.Fatal(err)
	}
	second, err := verify.RunCorpus(ctx, built.Artifact, suite, cases, filepath.Join(root, "runs"), first.SavedReport, run)
	if err != nil || calls != 2 || !second.Cases[0].Reused || !second.Cases[1].Reused {
		t.Fatal(second, err, calls)
	}
	for _, c := range cases {
		current, err := os.ReadFile(c.Inputs["document"])
		if err != nil || office.Hash(current) != office.Hash(data) {
			t.Fatal("source changed", err)
		}
	}
	// A clean rerun uses new working files in the same warm host, not saved edits.
	third, err := verify.RunCorpus(ctx, built.Artifact, suite, cases, filepath.Join(root, "runs"), "", run)
	if err != nil || third.Status != "passed" || calls != 4 {
		t.Fatal(third, err, calls)
	}
}
