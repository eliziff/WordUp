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
		Patches:        []office.XMLPatch{{Offset: start, Length: end - start, Text: `<title>After</title>`}},
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
		Patches:        []office.XMLPatch{{Offset: start, Text: "x"}},
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
		Patches:        []office.XMLPatch{{Text: `<root/>`}},
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

func TestGuardedStyleAndPatchBatches(t *testing.T) {
	root := t.TempDir()
	e := &Engine{Root: root}
	path := "package/styles.xml"
	original := []byte(`<w:styles xmlns:w="` + office.W + `"><!--untouched--></w:styles>`)
	if err := project.Write(root, path, original, ""); err != nil {
		t.Fatal(err)
	}
	recipe := &office.StyleRecipe{Styles: []office.StyleSpec{{ID: "Body", Run: map[string]any{"kerning_pt": 8, "bold_complex_script": true}}, {ID: "Heading", Paragraph: map[string]any{"bidi": true, "keep_next": true}}}}
	result, err := e.Call(context.Background(), "styles.apply", Parameters{Path: path, ExpectedSHA256: office.Hash(original), Recipe: recipe})
	if err != nil {
		t.Fatal(err)
	}
	after := mustReadXML(t, root, path)
	if result.(map[string]any)["styles_applied"] != 2 || !strings.Contains(string(after), "<!--untouched-->") || !strings.Contains(string(after), `<w:kern w:val="16"/>`) {
		t.Fatalf("wrong style result: %s", after)
	}
	for _, tc := range []struct {
		method string
		p      Parameters
	}{
		{"styles.apply", Parameters{Recipe: recipe, ExpectedSHA256: office.Hash(original)}},
		{"styles.apply", Parameters{ExpectedSHA256: office.Hash(after)}},
		{"styles.apply", Parameters{Recipe: recipe}},
		{"styles.apply", Parameters{Recipe: &office.StyleRecipe{Styles: []office.StyleSpec{{ID: "Body", Name: "rollback"}, {ID: "Heading", Run: map[string]any{"invalid": true}}}}, ExpectedSHA256: office.Hash(after)}},
		{"xml.patch", Parameters{Patches: []office.XMLPatch{{Offset: 0, Text: "<invalid>"}}, ExpectedSHA256: office.Hash(after)}},
	} {
		tc.p.Path = path
		if _, err := e.Call(context.Background(), tc.method, tc.p); err == nil {
			t.Fatalf("invalid request accepted: %+v", tc.p)
		}
		if string(mustReadXML(t, root, path)) != string(after) {
			t.Fatal("failed operation changed source")
		}
	}
	raw := []byte(`<root><a>old</a><b>old</b></root>`)
	if err := project.Write(root, path, raw, office.Hash(after)); err != nil {
		t.Fatal(err)
	}
	patches := []office.XMLPatch{{Offset: strings.LastIndex(string(raw), "old"), Length: 3, Text: "second"}, {Offset: strings.Index(string(raw), "old"), Length: 3, Text: "first"}}
	if _, err := e.Call(context.Background(), "xml.patch", Parameters{Path: path, Patches: patches, ExpectedSHA256: office.Hash(raw)}); err != nil {
		t.Fatal(err)
	}
	if got := string(mustReadXML(t, root, path)); got != `<root><a>first</a><b>second</b></root>` {
		t.Fatal(got)
	}
}
