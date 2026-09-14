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

func TestWorkspaceManifestHasOneCurrentShape(t *testing.T) {
	root := filepath.Join(t.TempDir(), "work")
	if _, err := New("TestProject", root); err != nil {
		t.Fatal(err)
	}
	b, err := Read(root, "project.json")
	if err != nil {
		t.Fatal(err)
	}
	if string(b) != "{\n  \"name\": \"TestProject\"\n}\n" {
		t.Fatalf("project manifest grew a redundant source of truth: %s", b)
	}
	if err = Write(root, "project.json", []byte(`{"format":2,"name":"TestProject"}`), ""); err != nil {
		t.Fatal(err)
	}
	if _, err = Open(root); err == nil || !strings.Contains(err.Error(), "unknown field \"format\"") {
		t.Fatalf("discarded workspace format was accepted: %v", err)
	}
}

func TestImportBuildPreservesCaseDistinctPackageParts(t *testing.T) {
	p := office.BlankPackage()
	p.Files["customXml/item1.xml"] = []byte(`<item xmlns="urn:one"/>`)
	p.Files["customXML/item3.xml"] = []byte(`<item xmlns="urn:three"/>`)
	if err := p.ContentType("customXml/item1.xml", "application/xml"); err != nil {
		t.Fatal(err)
	}
	if err := p.ContentType("customXML/item3.xml", "application/xml"); err != nil {
		t.Fatal(err)
	}
	if err := p.Relationship("word/document.xml", "rIdCase", office.R+"/customXml", "/customXML/item3.xml", ""); err != nil {
		t.Fatal(err)
	}
	if err := p.Validate(); err != nil {
		t.Fatal(err)
	}
	original, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	input := filepath.Join(t.TempDir(), "case-distinct.docx")
	if err := os.WriteFile(input, original, 0600); err != nil {
		t.Fatal(err)
	}
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := Import(input, root); err != nil {
		t.Fatal(err)
	}
	w, err := Open(root)
	if err != nil {
		t.Fatal(err)
	}
	report, err := w.Build("")
	if err != nil {
		t.Fatal(err)
	}
	built, err := os.ReadFile(report.Artifact)
	if err != nil {
		t.Fatal(err)
	}
	out, err := office.ReadPackage(built)
	if err != nil {
		t.Fatal(err)
	}
	for name, want := range map[string][]byte{
		"customXml/item1.xml": p.Files["customXml/item1.xml"],
		"customXML/item3.xml": p.Files["customXML/item3.xml"],
	} {
		if !bytes.Equal(out.Files[name], want) {
			t.Fatalf("case-distinct package part changed or disappeared: %s", name)
		}
	}
}

func TestRibbonMergeTargetUsesImportedPartSpelling(t *testing.T) {
	const namespace = "http://schemas.microsoft.com/office/2009/07/customui"
	base := office.BlankPackage()
	baseRibbon := []byte(`<customUI xmlns="` + namespace + `"><ribbon><tabs><tab id="base"/></tabs></ribbon></customUI>`)
	fragment := []byte(`<customUI xmlns="` + namespace + `"><ribbon><tabs><tab id="added"/></tabs></ribbon></customUI>`)
	if err := base.MergeRibbon("customUI/customUI14.xml", baseRibbon); err != nil {
		t.Fatal(err)
	}
	original, err := base.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	input := filepath.Join(t.TempDir(), "ribbon-case.docx")
	if err := os.WriteFile(input, original, 0600); err != nil {
		t.Fatal(err)
	}
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := Import(input, root); err != nil {
		t.Fatal(err)
	}
	if err := Write(root, "assets/fragment.xml", fragment, ""); err != nil {
		t.Fatal(err)
	}
	lock := map[string]any{"schema": 1, "components": map[string]any{
		"ribbon.case": map[string]any{"ribbon_merges": []map[string]string{{"source": "assets/fragment.xml", "target": "customui/customui14.xml"}}},
	}}
	if err := Write(root, componentLockSource, JSON(lock), ""); err != nil {
		t.Fatal(err)
	}
	w, err := Open(root)
	if err != nil {
		t.Fatal(err)
	}
	report, err := w.Build("")
	if err != nil {
		t.Fatal("case-variant Ribbon target rejected:", err)
	}
	built, err := os.ReadFile(report.Artifact)
	if err != nil {
		t.Fatal(err)
	}
	out, err := office.ReadPackage(built)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(out.Files["customUI/customUI14.xml"], []byte(`id="added"`)) {
		t.Fatal("Ribbon fragment did not merge into the imported part spelling")
	}
	if _, exists := out.Files["customui/customui14.xml"]; exists {
		t.Fatal("case-variant Ribbon merge created a second package part")
	}
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
	if err = Write(w.Root, componentLockSource, []byte(`{"schema":1,"components":{}}`), ""); err != nil {
		t.Fatal(err)
	}
	r, err = w.Build("")
	if err != nil || r.Cached {
		t.Fatalf("component lock change reused a persisted build: %v", err)
	}
	const sourcePath = "vba/Answer.bas"
	sourceInfo, err := os.Stat(filepath.Join(w.Root, filepath.FromSlash(sourcePath)))
	if err != nil {
		t.Fatal(err)
	}
	originalSource, err := Read(w.Root, sourcePath)
	if err != nil {
		t.Fatal(err)
	}
	mutatedSource := append([]byte(nil), originalSource...)
	mutatedSource[0] = 'o'
	if err = os.WriteFile(filepath.Join(w.Root, filepath.FromSlash(sourcePath)), mutatedSource, 0600); err != nil {
		t.Fatal(err)
	}
	if err = os.Chtimes(filepath.Join(w.Root, filepath.FromSlash(sourcePath)), sourceInfo.ModTime(), sourceInfo.ModTime()); err != nil {
		t.Fatal(err)
	}
	r, err = w.Build("")
	if err != nil || r.Cached {
		t.Fatalf("trusted same-size source edit with restored timestamp: %v", err)
	}
	if err = os.WriteFile(filepath.Join(w.Root, filepath.FromSlash(sourcePath)), originalSource, 0600); err != nil {
		t.Fatal(err)
	}
	if err = os.Chtimes(filepath.Join(w.Root, filepath.FromSlash(sourcePath)), sourceInfo.ModTime(), sourceInfo.ModTime()); err != nil {
		t.Fatal(err)
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
	artifactInfo, err := os.Stat(r.Artifact)
	if err != nil {
		t.Fatal(err)
	}
	tampered := append([]byte(nil), b...)
	tampered[0] ^= 1
	if err = os.WriteFile(r.Artifact, tampered, 0600); err != nil {
		t.Fatal(err)
	}
	if err = os.Chtimes(r.Artifact, artifactInfo.ModTime(), artifactInfo.ModTime()); err != nil {
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

func TestMetadataChangedDetectsSameStampHashEdit(t *testing.T) {
	previous := map[string]fileStamp{
		"project.json": {Size: 16, ModifiedNS: 7, ChangedNS: 0, Hash: "before"},
		".wordwright/base.opc": {Size: 32, ModifiedNS: 8, ChangedNS: 0, Hash: "same"},
		".wordwright/index.json": {Size: 24, ModifiedNS: 9, ChangedNS: 0, Hash: "same-index"},
	}
	current := map[string]fileStamp{
		"project.json": {Size: 16, ModifiedNS: 7, ChangedNS: 0, Hash: "after"},
		".wordwright/base.opc": {Size: 32, ModifiedNS: 8, ChangedNS: 0, Hash: "same"},
		".wordwright/index.json": {Size: 24, ModifiedNS: 9, ChangedNS: 0, Hash: "same-index"},
	}
	if !metadataChanged(previous, current) {
		t.Fatal("same-stamp metadata edit was not detected")
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

func TestSourceStampsRejectsMetadataSymlink(t *testing.T) {
	w := newWorkspace(t)
	target := filepath.Join(w.Root, "project.json")
	backup := filepath.Join(w.Root, ".project.json.wordup-test-backup")
	if err := os.Rename(target, backup); err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		_ = os.Remove(target)
		_ = os.Rename(backup, target)
	})
	if err := os.Symlink(backup, target); err != nil {
		t.Skipf("metadata symlinks unavailable: %v", err)
	}
	if _, err := sourceStamps(w.Root, nil); err == nil || !strings.Contains(err.Error(), "workspace metadata symlink is not accepted: project.json") {
		t.Fatalf("metadata symlink was followed: %v", err)
	}
}

func TestSourceStampsRejectsOversizedSourceBeforeHashing(t *testing.T) {
	w := newWorkspace(t)
	path := filepath.Join(w.Root, "assets", "oversized.bin")
	if err := os.MkdirAll(filepath.Dir(path), 0700); err != nil {
		t.Fatal(err)
	}
	file, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := file.Truncate(int64(office.Limit) + 1); err != nil {
		file.Close()
		t.Fatal(err)
	}
	if err := file.Close(); err != nil {
		t.Fatal(err)
	}
	if _, err := sourceStamps(w.Root, nil); err == nil || !strings.Contains(err.Error(), "file budget exceeded: assets/oversized.bin") {
		t.Fatalf("oversized source was hashed or accepted: %v", err)
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
