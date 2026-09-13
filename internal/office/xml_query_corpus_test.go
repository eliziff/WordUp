package office

import (
	"context"
	"os"
	"path/filepath"
	"strings"
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
	passed, rejected, paragraphs := 0, 0, 0
	var slowest time.Duration
	started := time.Now()
	for _, file := range files {
		data, err := os.ReadFile(file)
		if err != nil {
			t.Fatal(err)
		}
		pkg, err := ReadPackage(data)
		if err != nil {
			if strings.HasSuffix(file, "--3a668c4999c4.docx") && strings.Contains(err.Error(), "unsafe package part") {
				rejected++
				continue
			}
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
	if passed != 108 || rejected != 1 {
		t.Fatalf("corpus disposition changed: %d parsed, %d rejected", passed, rejected)
	}
	t.Logf("%d parsed, %d known unsafe ZIP rejected; %d paragraphs; total %s; slowest query %s", passed, rejected, paragraphs, time.Since(started), slowest)
}
