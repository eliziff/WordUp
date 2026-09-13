package inspect

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

// The private corpus is opt-in and never embedded in the released package.
// This verifies structural invariants and throughput, not editorial accuracy.
func TestSubmissionCorpus(t *testing.T) {
	root := os.Getenv("WORDUP_CORPUS")
	if root == "" {
		t.Skip("set WORDUP_CORPUS to retained documents directory")
	}
	files, err := filepath.Glob(filepath.Join(root, "*.docx"))
	if err != nil {
		t.Fatal(err)
	}
	if len(files) != 109 {
		t.Fatalf("expected 109 retained manuscripts, got %d", len(files))
	}
	passed, rejected, paragraphs := 0, 0, 0
	var total, slowest time.Duration
	for _, file := range files {
		start := time.Now()
		result, err := StructureResolved(file)
		elapsed := time.Since(start)
		total += elapsed
		if elapsed > slowest {
			slowest = elapsed
		}
		if err != nil {
			if !strings.HasSuffix(file, "--3a668c4999c4.docx") || !strings.Contains(err.Error(), "unsafe package part") {
				t.Errorf("unexpected rejection of %s: %v", filepath.Base(file), err)
			}
			rejected++
			t.Logf("REJECTED %s: %v", filepath.Base(file), err)
			continue
		}
		passed++
		rows := result["paragraphs"].([]map[string]any)
		paragraphs += len(rows)
		for i, row := range rows {
			resolved := row["resolved_structure"].(map[string]any)
			if resolved["role"] != "heading" {
				continue
			}
			level := resolved["level"].(int)
			parent := resolved["parent_paragraph"].(int)
			if level < 1 || level > 9 || parent < 0 || parent > i {
				t.Fatalf("invalid hierarchy in %s paragraph %d: %v", file, i+1, resolved)
			}
			if parent > 0 {
				above := rows[parent-1]["resolved_structure"].(map[string]any)
				if above["role"] != "heading" || above["level"].(int) >= level {
					t.Fatalf("invalid parent in %s paragraph %d", file, i+1)
				}
			}
		}
	}
	t.Logf("documents=%d resolved=%d rejected=%d paragraphs=%d total=%s slowest=%s", len(files), passed, rejected, paragraphs, total, slowest)
	if rejected > 1 {
		t.Fatalf("unexpected rejected manuscripts: %d", rejected)
	}
}
