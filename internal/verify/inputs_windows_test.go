//go:build windows && (amd64 || arm64)

package verify_test

import (
	"context"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
	"os"
	"path/filepath"
	"testing"
)

func TestNativeReplayInputIsolation(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	root := t.TempDir()
	workspace := filepath.Join(root, "workspace")
	if _, err := project.New("InputIsolation", workspace); err != nil {
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
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Original</w:t></w:r></w:p></w:body></w:document>`)
	source, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	manuscript := filepath.Join(root, "manuscript.docx")
	if err := os.WriteFile(manuscript, source, 0600); err != nil {
		t.Fatal(err)
	}
	suite := verify.Suite{Schema: 1, Name: "Input isolation", Inputs: map[string]string{"manuscript": manuscript}, Steps: []verify.Step{
		{Name: "Open working copy", Operation: native.Operation{Op: "open", File: "$input:manuscript$", As: "manuscript"}},
		{Name: "Content", Operation: native.Operation{Op: "get", Target: "manuscript", Member: "Content", As: "body"}},
		{Name: "Original text", Operation: native.Operation{Op: "get", Target: "body", Member: "Text"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "Original\r"}}},
		{Name: "Edit", Operation: native.Operation{Op: "put", Target: "body", Member: "Text", Value: "Saved edit"}},
		{Name: "Save working copy", Operation: native.Operation{Op: "invoke", Target: "manuscript", Member: "Save"}},
		{Name: "Release range", Operation: native.Operation{Op: "release", Target: "body"}},
		{Name: "Close working copy", Operation: native.Operation{Op: "unload", Target: "manuscript"}},
	}}
	ctx := context.Background()
	h, err := native.Start(ctx, native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	first, err := verify.Run(ctx, built.Artifact, suite, h, true)
	if err != nil {
		t.Fatal(err)
	}
	original, err := os.ReadFile(manuscript)
	if err != nil || office.Hash(original) != office.Hash(source) {
		t.Fatal("test saved over original manuscript")
	}
	if err := os.WriteFile(manuscript, []byte("original changed outside the test"), 0600); err != nil {
		t.Fatal(err)
	}
	second, err := verify.RunWithInputs(ctx, first.ArtifactSnapshot, suite, h, true, first.InputSnapshots)
	if err != nil {
		t.Fatal(err)
	}
	if second.InputSnapshots["manuscript"].SHA256 != office.Hash(source) || second.EvidenceDirectory == first.EvidenceDirectory {
		t.Fatal("replay did not preserve input identity and isolation")
	}
	t.Logf("native input isolation and replay passed: %.1f ms then %.1f ms", first.DurationMS, second.DurationMS)
}
