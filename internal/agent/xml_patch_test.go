package agent

import (
	"context"
	"os"
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

func TestXMLPatchRejectsAbsolutePath(t *testing.T) {
	root := filepath.Join(t.TempDir(), "work")
	if _, err := project.New("Test", root); err != nil {
		t.Fatal(err)
	}
	outside := filepath.Join(t.TempDir(), "outside.xml")
	if err := project.AtomicWrite(outside, []byte(`<root/>`)); err != nil {
		t.Fatal(err)
	}
	e := &Engine{Root: root}
	if _, err := e.Call(context.Background(), "xml.patch", Parameters{
		Path:           outside,
		Offset:         0,
		Length:         0,
		Text:           `<root/>`,
		ExpectedSHA256: office.Hash([]byte(`<root/>`)),
	}); err == nil || !strings.Contains(err.Error(), "workspace-relative") {
		t.Fatalf("absolute XML patch path was accepted: %v", err)
	}
	if got, err := os.ReadFile(outside); err != nil || string(got) != `<root/>` {
		t.Fatalf("absolute XML patch changed outside file: %v %q", err, got)
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
