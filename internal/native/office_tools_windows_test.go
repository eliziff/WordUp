//go:build windows && (amd64 || arm64)

package native

import (
	"archive/zip"
	"context"
	"crypto/sha256"
	"github.com/eliziff/WordUp/internal/office"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// Build this test executable into bin beside the bundled helpers. No Word is started.
func TestOfficeToolsIntegration(t *testing.T) {
	if os.Getenv("WORDUP_OFFICE_TOOLS_TEST") != "1" {
		t.Skip("set WORDUP_OFFICE_TOOLS_TEST=1 with bundled helpers beside test executable")
	}
	defer StopOfficeTools()
	makeDocument := func(name, body string) string {
		t.Helper()
		path := filepath.Join(t.TempDir(), name+".docx")
		f, err := os.Create(path)
		if err != nil {
			t.Fatal(err)
		}
		z := zip.NewWriter(f)
		for name, data := range map[string]string{
			"[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
			"_rels/.rels":         `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
			"word/document.xml":   `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` + body + `</w:body></w:document>`,
		} {
			w, e := z.Create(name)
			if e != nil {
				t.Fatal(e)
			}
			if _, e = w.Write([]byte(data)); e != nil {
				t.Fatal(e)
			}
		}
		if err = z.Close(); err != nil {
			t.Fatal(err)
		}
		if err = f.Close(); err != nil {
			t.Fatal(err)
		}
		return path
	}
	valid := makeDocument("valid", `<w:p><w:r><w:t>Hello</w:t></w:r></w:p>`)
	invalid := makeDocument("invalid", `<w:bogus/>`)
	validate := func(path string) map[string]any {
		t.Helper()
		before, err := os.ReadFile(path)
		if err != nil {
			t.Fatal(err)
		}
		v, err := OfficeTools(context.Background(), "validate", path)
		if err != nil {
			t.Fatal(err)
		}
		after, err := os.ReadFile(path)
		if err != nil || sha256.Sum256(before) != sha256.Sum256(after) {
			t.Fatal("validation changed input", err)
		}
		return v.(map[string]any)
	}
	first := validate(valid)
	if first["valid"] != true {
		t.Fatal(first)
	}
	source, err := os.ReadFile(valid)
	if err != nil {
		t.Fatal(err)
	}
	pkg, err := office.ReadPackage(source)
	if err != nil {
		t.Fatal(err)
	}
	if err = office.AddBuildingBlocks(pkg, []office.BuildingBlock{{Name: "Schema checked block", Description: "Description after category and behaviors", Blocks: []office.Block{{Text: "Reusable text"}}}}, nil); err != nil {
		t.Fatal(err)
	}
	encoded, err := pkg.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	blocks := filepath.Join(t.TempDir(), "blocks.docx")
	if err = os.WriteFile(blocks, encoded, 0600); err != nil {
		t.Fatal(err)
	}
	if result := validate(blocks); result["valid"] != true {
		t.Fatal("generated building block violates independent SDK schema", result)
	}
	bad := validate(invalid)
	if bad["valid"] != false || bad["helper_reused"] != true {
		t.Fatal(bad)
	}
	errors := bad["errors"].([]any)
	if len(errors) == 0 {
		t.Fatal("invalid OOXML produced no diagnostics")
	}
	diagnostic := errors[0].(map[string]any)
	if diagnostic["part"] != "/word/document.xml" || diagnostic["xpath"] == nil || diagnostic["xpath"] == "" || diagnostic["id"] == "" {
		t.Fatal(diagnostic)
	}
	if _, err := OfficeTools(context.Background(), "validate", filepath.Join(t.TempDir(), "missing.docx")); err == nil {
		t.Fatal("missing file accepted")
	}
	if recovered := validate(valid); recovered["valid"] != true || recovered["helper_pid"] != first["helper_pid"] {
		t.Fatal("helper failed to recover from document error", recovered)
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if _, err := OfficeTools(ctx, "validate", valid); err == nil || !strings.Contains(err.Error(), "deadline") {
		t.Fatal("cancelled helper call was not bounded", err)
	}
	if recovered := validate(valid); recovered["valid"] != true || recovered["helper_reused"] != false {
		t.Fatal("helper failed to restart after cancellation", recovered)
	}
}
