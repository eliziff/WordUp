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
		{Name: "Save working copy", Operation: native.Operation{Op: "invoke", Target: "manuscript", Member: "SaveAs2", Args: []any{"$output/saved.docx"}}},
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
	if err := verify.SaveReport(workspace, first); err != nil {
		t.Fatal(err)
	}
	bundle := filepath.Join(root, "baseline")
	if _, err := verify.Freeze(first.SavedReport, bundle); err != nil {
		t.Fatal(err)
	}
	relocated := filepath.Join(root, "relocated-baseline")
	if err := os.Rename(bundle, relocated); err != nil {
		t.Fatal(err)
	}
	// Make the old evidence paths unavailable, proving no fallback to originals.
	if err := os.Rename(first.EvidenceDirectory, first.EvidenceDirectory+"-retired"); err != nil {
		t.Fatal(err)
	}
	frozen, _, err := verify.LoadReport(filepath.Join(relocated, "bundle.json"))
	if err != nil {
		t.Fatal(err)
	}
	if frozen.SuiteSHA256 != first.SuiteSHA256 {
		t.Fatal("relocation rewrote logical suite")
	}
	second, err := verify.RunWithInputs(ctx, frozen.ArtifactSnapshot, frozen.Suite, h, true, frozen.InputSnapshots)
	if err != nil {
		t.Fatal(err)
	}
	if second.InputSnapshots["manuscript"].SHA256 != office.Hash(source) || second.EvidenceDirectory == first.EvidenceDirectory {
		t.Fatal("replay did not preserve input identity and isolation")
	}
	t.Logf("native input isolation and replay passed: %.1f ms then %.1f ms", first.DurationMS, second.DurationMS)
	// Direct staging must work too, without the harness's unique input names.
	direct := filepath.Join(root, "direct.docx")
	if err := os.WriteFile(direct, source, 0600); err != nil {
		t.Fatal(err)
	}
	call := func(op native.Operation) any {
		t.Helper()
		result, err := h.Call(ctx, op)
		if err != nil {
			t.Fatal(err)
		}
		return result
	}
	call(native.Operation{Op: "open", File: direct, As: "direct"})
	call(native.Operation{Op: "invoke", Target: "direct", Member: "SaveAs2", Args: []any{filepath.Join(root, "saved-direct.docx")}})
	call(native.Operation{Op: "unload", Target: "direct"})
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:r><w:t>Changed source</w:t></w:r></w:p></w:body></w:document>`)
	changed, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(direct, changed, 0600); err != nil {
		t.Fatal(err)
	}
	call(native.Operation{Op: "open", File: direct, As: "direct"})
	call(native.Operation{Op: "get", Target: "direct", Member: "Content", As: "directBody"})
	actual := call(native.Operation{Op: "get", Target: "directBody", Member: "Text"})
	if actual != "Changed source\r" {
		t.Fatal("reopened stale source", actual)
	}
	call(native.Operation{Op: "release", Target: "directBody"})
	call(native.Operation{Op: "unload", Target: "direct"})
}
