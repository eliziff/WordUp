package native

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestSessionLogsRetainLoadErrorsAndBoundOutput(t *testing.T) {
	dir := t.TempDir()
	for name, data := range map[string][]byte{"Form.log": {255, 254, 'E', 0, 'r', 0, 'r', 0}, "large.log": []byte(strings.Repeat("x", 262145)), "unrelated.docx": []byte("not a log")} {
		if err := os.WriteFile(filepath.Join(dir, name), data, 0600); err != nil {
			t.Fatal(err)
		}
	}
	r := sessionLogs(dir)
	files := r["files"].([]any)
	if len(files) != 2 {
		t.Fatal(r)
	}
	first := files[0].(map[string]any)
	if first["text"] != "Err" || first["base64"] == "" {
		t.Fatal(first)
	}
	large := files[1].(map[string]any)
	if large["truncated"] != true || len(large["text"].(string)) != 262144 {
		t.Fatal("unbounded log")
	}
}
