package office

import (
	"context"
	"os"
	"path/filepath"
	"testing"
	"time"
)

// The retained private corpus is opt-in, never embedded in a release or test
// fixture. This proves parser agreement, not heading accuracy or visual fidelity.
func TestQueryXMLRetainedCorpus(t *testing.T) {
	root := os.Getenv("WORDUP_CORPUS")
	if root == "" {
		t.Skip("set WORDUP_CORPUS to retained submissions/documents")
	}
	files, err := filepath.Glob(filepath.Join(root, "*.docx"))
	if err != nil || len(files) != 109 {
		t.Fatalf("expected 109 retained documents, found %d: %v", len(files), err)
	}
	passed, paragraphs := 0, 0
	var slowest time.Duration
	started := time.Now()
	for _, file := range files {
		data, err := os.ReadFile(file)
		if err != nil {
			t.Fatal(err)
		}
		pkg, err := ReadPackage(data)
		if err != nil {
			t.Fatalf("%s: %v", filepath.Base(file), err)
		}
		xml := pkg.Files["word/document.xml"]
		spans, err := XMLSpans(xml)
		if err != nil {
			t.Fatal(err)
		}
		expected := 0
		for _, span := range spans {
			if span.Name.Space == W && span.Name.Local == "p" {
				expected++
			}
		}
		tick := time.Now()
		got, err := QueryXML(context.Background(), xml, "count(//w:p)", map[string]string{"w": W}, 20)
		slowest = max(slowest, time.Since(tick))
		if err != nil || got["value"] != float64(expected) {
			t.Fatalf("%s: paragraph count %d, XPath %v, error %v", filepath.Base(file), expected, got, err)
		}
		passed++
		paragraphs += expected
	}
	if passed != len(files) {
		t.Fatalf("parsed %d of %d documents", passed, len(files))
	}
	t.Logf("%d parsed; %d paragraphs; total %s; slowest query %s", passed, paragraphs, time.Since(started), slowest)
}
