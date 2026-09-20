package agent

import (
	"context"
	"path/filepath"
	"testing"

	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

func TestXMLQueryPackageThroughAgent(t *testing.T) {
	root := t.TempDir()
	pkg := office.BlankPackage()
	pkg.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p/></w:body></w:document>`)
	data, err := pkg.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "sample.docx", data, ""); err != nil {
		t.Fatal(err)
	}
	e := &Engine{Root: root}
	params := Parameters{Path: "sample.docx", Part: "word/document.xml", Query: "count(//w:p)", Namespaces: map[string]string{"w": office.W}}
	got, err := e.Call(context.Background(), "xml.query", params)
	if err != nil {
		t.Fatal(err)
	}
	r := got.(map[string]any)
	if r["value"] != float64(1) || r["package_sha256"] != office.Hash(data) || r["source_sha256"] != office.Hash(pkg.Files[params.Part]) || e.host != nil {
		t.Fatalf("lost package identity or started Word: %+v", r)
	}
	for _, part := range []string{"../word/document.xml", "word/missing.xml"} {
		params.Part = part
		if _, err := e.Call(context.Background(), "xml.query", params); err == nil {
			t.Fatalf("accepted unsafe/missing part: %s", part)
		}
	}
	unchanged, err := project.Read(root, filepath.Base(params.Path))
	if err != nil || office.Hash(unchanged) != office.Hash(data) {
		t.Fatal("query modified the input document")
	}
}

func TestXMLReviewDoesNotStartWord(t *testing.T) {
	root := t.TempDir()
	for _, name := range []string{"before.xml", "after.xml"} {
		data := []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:r><w:t>` + name + `</w:t></w:r></w:p></w:body></w:document>`)
		if err := project.Write(root, name, data, ""); err != nil {
			t.Fatal(err)
		}
	}
	e := &Engine{Root: root}
	r, err := e.Call(context.Background(), "xml.review", Parameters{Reference: "before.xml", Path: "after.xml", Limit: 1})
	if err != nil || e.host != nil {
		t.Fatal(r, err)
	}
	if r.(map[string]any)["changed_paragraphs"] != 1 {
		t.Fatal(r)
	}
}
