package agent

import (
	"context"
	"path/filepath"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

func TestXMLPatchUsesQueryRangeAndHashGuard(t *testing.T) {
	root := filepath.Join(t.TempDir(), "work")
	if _, err := project.New("Test", root); err != nil {
		t.Fatal(err)
	}
	input := []byte(`<root><title>Before</title><keep>opaque</keep></root>`)
	if err := project.Write(root, "package/custom.xml", input, ""); err != nil {
		t.Fatal(err)
	}
	e := &Engine{Root: root}
	q, err := e.Call(context.Background(), "xml.query", Parameters{Path: "package/custom.xml", Query: "//title", Limit: 1})
	if err != nil {
		t.Fatal(err)
	}
	query := q.(map[string]any)
	match := query["matches"].([]any)[0].(map[string]any)
	start, end := match["element_start"].(int), match["element_end"].(int)
	result, err := e.Call(context.Background(), "xml.patch", Parameters{
		Path:           "package/custom.xml",
		Offset:         start,
		Length:         end - start,
		Text:           `<title>After</title>`,
		ExpectedSHA256: query["source_sha256"].(string),
	})
	if err != nil {
		t.Fatal(err)
	}
	actual := string(mustReadXML(t, root, "package/custom.xml"))
	if !strings.Contains(actual, `<title>After</title>`) || !strings.Contains(actual, `<keep>opaque</keep>`) {
		t.Fatalf("patch did not preserve direct XML: %v", result)
	}
	if _, err := e.Call(context.Background(), "xml.patch", Parameters{
		Path:           "package/custom.xml",
		Offset:         start,
		Length:         0,
		Text:           "x",
		ExpectedSHA256: query["source_sha256"].(string),
	}); err == nil {
		t.Fatal("stale XML patch accepted")
	}
}

func mustReadXML(t *testing.T, root, name string) []byte {
	t.Helper()
	b, err := project.Read(root, name)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := office.XMLSpans(b); err != nil {
		t.Fatal(err)
	}
	return b
}
