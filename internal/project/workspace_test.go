package project

import (
	"bytes"
	"github.com/eliziff/WordUp/internal/office"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestReadJSONWindowsBOM(t *testing.T) {
	for _, prefix := range []string{"", "\xef\xbb\xbf"} {
		var value struct{ Name string }
		if err := ReadJSON([]byte(prefix+`{"Name":"WordUp"}`), &value); err != nil || value.Name != "WordUp" {
			t.Fatalf("UTF-8 input rejected: %v", err)
		}
		for _, invalid := range []string{`{"Unknown":true}`, `{"Name":"WordUp"} {}`, "\xef\xbb\xbf{}"} {
			if prefix == "" && invalid == "\xef\xbb\xbf{}" {
				continue
			}
			if err := ReadJSON([]byte(prefix+invalid), &value); err == nil {
				t.Fatalf("strict parsing weakened for %q", prefix+invalid)
			}
		}
	}
}

func newWorkspace(t *testing.T) *Workspace {
	t.Helper()
	root := filepath.Join(t.TempDir(), "work")
	if _, err := New("TestProject", root); err != nil {
		t.Fatal(err)
	}
	w, err := Open(root)
	if err != nil {
		t.Fatal(err)
	}
	return w
}
func TestSourceBuildCacheAndTamper(t *testing.T) {
	w := newWorkspace(t)
	if err := Write(w.Root, "vba/Answer.bas", []byte("Option Explicit\nPublic Function Answer() As Long\nAnswer = 42\nEnd Function\n"), ""); err != nil {
		t.Fatal(err)
	}
	r, err := w.Build("")
	if err != nil {
		t.Fatal(err)
	}
	if r.Modules != 2 || r.WordExecuted || r.VBACompiled {
		t.Fatalf("untrue build evidence: %+v", r)
	}
	b, _ := os.ReadFile(r.Artifact)
	p, err := office.ReadPackage(b)
	if err != nil {
		t.Fatal(err)
	}
	v, err := office.ReadVBA(p.Files["word/vbaProject.bin"])
	if err != nil || len(v.Modules) != 2 {
		t.Fatalf("%v", err)
	}
	r, err = w.Build("")
	if err != nil || !r.Cached {
		t.Fatalf("cache miss: %v", err)
	}
	r.ToolVersion = "older-writer"
	if err = Write(w.Root, "reports/build.json", JSON(r), ""); err != nil {
		t.Fatal(err)
	}
	r, err = w.Build("")
	if err != nil || r.Cached || r.ToolVersion != Version {
		t.Fatalf("reused old writer output: %+v %v", r, err)
	}
	if len(r.ToolSHA256) != 64 {
		t.Fatal("missing executable identity", r.ToolSHA256)
	}
	r.ToolSHA256 = "different binary with the same version"
	if err = Write(w.Root, "reports/build.json", JSON(r), ""); err != nil {
		t.Fatal(err)
	}
	r, err = w.Build("")
	if err != nil || r.Cached || r.ToolSHA256 != ExecutableSHA256() {
		t.Fatalf("same-version writer cache survived: %+v %v", r, err)
	}
	if err = os.WriteFile(r.Artifact, []byte("tampered"), 0600); err != nil {
		t.Fatal(err)
	}
	r, err = w.Build("")
	if err != nil || r.Cached {
		t.Fatalf("trusted corrupt cache: %v", err)
	}
	b, _ = os.ReadFile(r.Artifact)
	if office.Hash(b) != r.SHA256 {
		t.Fatal("wrong artifact hash")
	}
}
func TestGuardedWrites(t *testing.T) {
	w := newWorkspace(t)
	old, err := Read(w.Root, "vba/ThisDocument.cls")
	if err != nil {
		t.Fatal(err)
	}
	if err = Write(w.Root, "vba/ThisDocument.cls", []byte("bad"), "stale"); err == nil {
		t.Fatal("stale write accepted")
	}
	b, _ := Read(w.Root, "vba/ThisDocument.cls")
	if !bytes.Equal(b, old) {
		t.Fatal("changed after failed write")
	}
	for _, p := range []string{"../escape", "/absolute", "vba/../../escape"} {
		if _, err = Under(w.Root, p); err == nil {
			t.Fatalf("accepted %s", p)
		}
	}
}
func TestSpecimenWorkspaceRoundTripAndControlledEdit(t *testing.T) {
	src := os.Getenv("WORDUP_SPECIMEN")
	if src == "" {
		t.Skip("private specimen not provided")
	}
	orig, err := os.ReadFile(src)
	if err != nil {
		t.Fatal(err)
	}
	root := filepath.Join(t.TempDir(), "alr")
	if _, err = Import(src, root); err != nil {
		t.Fatal(err)
	}
	w, err := Open(root)
	if err != nil {
		t.Fatal(err)
	}
	r, err := w.Build("")
	if err != nil {
		t.Fatal(err)
	}
	out, _ := os.ReadFile(r.Artifact)
	if !bytes.Equal(out, orig) || !r.NoChange {
		t.Fatal("non-identical no-op round trip")
	}
	before, err := office.ReadVBA(w.Baseline.Files["word/vbaProject.bin"])
	if err != nil {
		t.Fatal(err)
	}
	files, err := w.SourceFiles()
	if err != nil {
		t.Fatal(err)
	}
	chosen := ""
	for n := range files {
		if strings.HasSuffix(n, "/Tester.bas") {
			chosen = n
			break
		}
	}
	if chosen == "" {
		t.Fatal("expected specimen Tester component")
	}
	if err = Write(root, chosen, append(files[chosen], []byte("\nPublic Function WordUpProof() As Long\nWordUpProof=42\nEnd Function\n")...), ""); err != nil {
		t.Fatal(err)
	}
	r, err = w.Build("")
	if err != nil {
		t.Fatal(err)
	}
	out, _ = os.ReadFile(r.Artifact)
	p, err := office.ReadPackage(out)
	if err != nil {
		t.Fatal(err)
	}
	after, err := office.ReadVBA(p.Files["word/vbaProject.bin"])
	if err != nil {
		t.Fatal(err)
	}
	preserved := 0
	for path, entry := range before.CFB.Entries {
		if entry.Kind == 2 && !strings.HasPrefix(path, "VBA/") && strings.Contains(path, "/") {
			other, err := after.CFB.Stream(path)
			if err != nil || !bytes.Equal(entry.Data, other) {
				t.Fatalf("lost form stream %s: %v", path, err)
			}
			preserved++
		}
	}
	if preserved != 26 {
		t.Fatalf("preserved %d form streams, expected 26", preserved)
	}
	if !bytes.Equal(p.Files["word/document.xml"], w.Baseline.Files["word/document.xml"]) {
		t.Fatal("unrelated document body changed")
	}
	for _, m := range before.Modules {
		if m.Name == "Tester" {
			continue
		}
		found := false
		for _, n := range after.Modules {
			if n.Name == m.Name {
				found = true
				if n.Source != m.Source {
					t.Fatalf("unrelated module %s changed", m.Name)
				}
			}
		}
		if !found {
			t.Fatal("module disappeared")
		}
	}
	still, _ := os.ReadFile(src)
	if !bytes.Equal(still, orig) {
		t.Fatal("original specimen changed")
	}
	t.Logf("15 modules retained; 14 unedited sources and 26 form streams and document body unchanged; modified artifact %s", r.SHA256)
}
