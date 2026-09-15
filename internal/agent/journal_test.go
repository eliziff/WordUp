package agent

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"github.com/eliziff/WordUp/internal/journal"
)

func TestJournalCreateCanOverlayOneCatalogSnapshot(t *testing.T) {
	root := t.TempDir()
	article := filepath.Join(root, "catalog", "UBC-L-REV", "59", "1", "article-1")
	if err := os.MkdirAll(article, 0700); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(article, "provenance.json"), []byte(`{"band_year":2026}`), 0600); err != nil {
		t.Fatal(err)
	}
	summary, err := json.Marshal(map[string]any{
		"document_date_en": "2026-01-01",
		"font_role_body":   map[string]any{"font": "Cambria", "size": 10.5},
		"font_role_note":   map[string]any{"font": "Cambria", "size": 8.5},
	})
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(article, "digitalborn_native_summary.json"), summary, 0600); err != nil {
		t.Fatal(err)
	}

	e := &Engine{Root: root}
	defer e.Close()
	value, err := e.Call(context.Background(), "journal.create", Parameters{
		Journal: "UBC-L-REV",
		Path:    "catalog",
		Output:  "workspace",
		Years:   []int{2026},
	})
	if err != nil {
		t.Fatal(err)
	}
	report, ok := value.(journal.CreateReport)
	if !ok {
		t.Fatalf("unexpected create result type: %#v", value)
	}
	if report.Journal.BodySizePT != 10.5 || report.Journal.NoteSizePT != 8.5 || report.Journal.EvidenceStatus != "corpus-observed" {
		t.Fatalf("catalog evidence did not reach generated profile: %#v", report.Journal)
	}
	if report.Build == nil || !report.Build.PackageValidated {
		t.Fatalf("generated catalog-backed workspace was not built: %#v", report.Build)
	}
}
