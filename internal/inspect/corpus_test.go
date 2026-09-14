package inspect

import (
	"os"
	"path/filepath"
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
	root = resolveCorpusRoot(root)
	files, err := filepath.Glob(filepath.Join(root, "*.docx"))
	if err != nil {
		t.Fatal(err)
	}
	if len(files) != 109 {
		t.Fatalf("expected 109 retained manuscripts, got %d", len(files))
	}
	passed, paragraphs := 0, 0
	var total, slowest time.Duration
	var slowestFile string
	for _, file := range files {
		start := time.Now()
		result, err := StructureResolved(file)
		elapsed := time.Since(start)
		total += elapsed
		if elapsed > slowest {
			slowest = elapsed
			slowestFile = filepath.Base(file)
		}
		if err != nil {
			t.Errorf("unexpected rejection of %s: %v", filepath.Base(file), err)
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
	t.Logf("documents=%d resolved=%d paragraphs=%d total=%s slowest=%s file=%s", len(files), passed, paragraphs, total, slowest, slowestFile)
	if passed != len(files) {
		t.Fatalf("resolved %d of %d manuscripts", passed, len(files))
	}
}

// go test runs a package test binary from that package's directory, so a
// repo-relative corpus path must be resolved against its ancestors rather
// than interpreted relative to internal/inspect.
func resolveCorpusRoot(root string) string {
	if filepath.IsAbs(root) {
		return root
	}
	if cwd, err := os.Getwd(); err == nil {
		for dir := cwd; ; dir = filepath.Dir(dir) {
			candidate := filepath.Join(dir, root)
			if info, statErr := os.Stat(candidate); statErr == nil && info.IsDir() {
				return candidate
			}
			parent := filepath.Dir(dir)
			if parent == dir {
				break
			}
		}
	}
	return root
}
