package project

import (
	"bytes"
	"encoding/json"
	"io/fs"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/office"
)

func TestRepeatedImportHasIdenticalFiles(t *testing.T) {
	w := newWorkspace(t)
	w.Manifest.Components["ExampleForm"] = "form"
	design := office.Design{Name: "ExampleForm", Mode: "replace", Controls: []office.ControlDesign{
		{Name: "Group", Type: "Frame", Controls: []office.ControlDesign{
			{Name: "Run", Type: "CommandButton", Properties: map[string]any{"Caption": "Run", "Width": 72}},
		}},
	}}
	for path, data := range map[string][]byte{
		"project.json":           JSON(w.Manifest),
		"forms/ExampleForm.json": JSON(design),
		"vba/ExampleForm.vba":    []byte("Attribute VB_Name = \"ExampleForm\"\nOption Explicit\nPrivate Sub Run_Click()\nEnd Sub\n"),
	} {
		if err := Write(w.Root, path, data, ""); err != nil {
			t.Fatal(err)
		}
	}
	if err := Write(w.Root, "vba/Example.bas", []byte("Public Sub Example()\nEnd Sub\n"), ""); err != nil {
		t.Fatal(err)
	}
	report, err := w.Build("")
	if err != nil {
		t.Fatal(err)
	}
	first, second := filepath.Join(t.TempDir(), "first"), filepath.Join(t.TempDir(), "second")
	for _, destination := range []string{first, second} {
		if _, err := Import(report.Artifact, destination); err != nil {
			t.Fatal(err)
		}
	}
	files := func(root string) map[string][]byte {
		t.Helper()
		result := map[string][]byte{}
		err := filepath.WalkDir(root, func(path string, entry fs.DirEntry, err error) error {
			if err != nil {
				return err
			}
			if entry.IsDir() {
				return nil
			}
			rel, err := filepath.Rel(root, path)
			if err != nil {
				return err
			}
			result[rel], err = os.ReadFile(path)
			return err
		})
		if err != nil {
			t.Fatal(err)
		}
		return result
	}
	a, b := files(first), files(second)
	if len(a) != len(b) {
		t.Fatal("Import changed file inventory")
	}
	for name, data := range a {
		if !bytes.Equal(data, b[name]) {
			t.Errorf("Import churn: %s", name)
		}
	}
	imported, err := Open(first)
	if err != nil {
		t.Fatal(err)
	}
	rebuilt, err := imported.Build("")
	if err != nil {
		t.Fatal(err)
	}
	third := filepath.Join(t.TempDir(), "reimport")
	if _, err := Import(rebuilt.Artifact, third); err != nil {
		t.Fatal(err)
	}
	c := files(third)
	if len(a) != len(c) {
		t.Fatal("No-op rebuild/reimport changed file inventory")
	}
	for name, data := range a {
		if !bytes.Equal(data, c[name]) {
			t.Errorf("No-op rebuild/reimport churn: %s", name)
		}
	}
}

func TestGitNoOpBuildAndIndependentModuleMerge(t *testing.T) {
	if _, err := exec.LookPath("git"); err != nil {
		t.Skip("git is required for checkout/merge verification")
	}
	w := newWorkspace(t)
	w.Manifest.Components["MergeForm"] = "form"
	design := office.Design{Name: "MergeForm", Mode: "replace", Controls: []office.ControlDesign{
		{Name: "Run", Type: "CommandButton", Properties: map[string]any{"Caption": "Before", "Width": 72}},
	}}
	ribbon := []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="MergeTab" label="Before"><group id="MergeGroup" label="Commands"><button id="MergeButton" label="Run"/></group></tab></tabs></ribbon></customUI>`)
	for path, data := range map[string][]byte{
		"project.json":                    JSON(w.Manifest),
		"forms/MergeForm.json":            JSON(design),
		"vba/MergeForm.vba":               []byte("Attribute VB_Name = \"MergeForm\"\nOption Explicit\n"),
		"package/customUI/customUI14.xml": ribbon,
	} {
		if err := Write(w.Root, path, data, ""); err != nil {
			t.Fatal(err)
		}
	}
	git := func(args ...string) string {
		t.Helper()
		cmd := exec.Command("git", append([]string{"-c", "user.name=WordUp Test", "-c", "user.email=test@example.invalid", "-c", "core.autocrlf=false"}, args...)...)
		cmd.Dir = w.Root
		out, err := cmd.CombinedOutput()
		if err != nil {
			t.Fatalf("git %v: %v\n%s", args, err, out)
		}
		return string(out)
	}
	if err := Write(w.Root, "tests/expected.xml", []byte("<root>\r\n<text>é</text>\r\n</root>\r\n"), ""); err != nil {
		t.Fatal(err)
	}
	if err := Write(w.Root, "assets/sample.bin", []byte{0, 13, 10, 255, 10}, ""); err != nil {
		t.Fatal(err)
	}
	git("init", "-q", "-b", "baseline")
	git("add", ".gitattributes", ".gitignore", "project.json", "vba", "package", "forms", ".wordwright", "tests", "assets")
	git("commit", "-qm", "baseline")
	for i := 0; i < 2; i++ {
		if _, err := w.Build(""); err != nil {
			t.Fatal(err)
		}
		if changed := git("diff", "--name-only", "HEAD"); changed != "" {
			t.Fatalf("No-op build changed tracked source: %s", changed)
		}
	}
	git("checkout", "-qb", "first")
	if err := Write(w.Root, "vba/First.bas", []byte("Public Function First() As Long\nFirst = 1\nEnd Function\n"), ""); err != nil {
		t.Fatal(err)
	}
	git("add", "vba/First.bas")
	ribbon = bytes.Replace(ribbon, []byte(`label="Before"`), []byte(`label="After merge"`), 1)
	if err := Write(w.Root, "package/customUI/customUI14.xml", ribbon, ""); err != nil {
		t.Fatal(err)
	}
	git("add", "package/customUI/customUI14.xml")
	git("commit", "-qm", "first module")
	git("checkout", "-qb", "second", "baseline")
	if err := Write(w.Root, "vba/Second.bas", []byte("Public Function Second() As Long\nSecond = 2\nEnd Function\n"), ""); err != nil {
		t.Fatal(err)
	}
	git("add", "vba/Second.bas")
	design.Controls[0].Properties["Caption"] = "After merge"
	if err := Write(w.Root, "forms/MergeForm.json", JSON(design), ""); err != nil {
		t.Fatal(err)
	}
	git("add", "forms/MergeForm.json")
	git("commit", "-qm", "second module")
	git("merge", "--no-edit", "first")
	merged, err := Open(w.Root)
	if err != nil {
		t.Fatal(err)
	}
	report, err := merged.Build("")
	if err != nil {
		t.Fatal(err)
	}
	if report.Modules != 4 {
		t.Fatalf("Merged modules missing: %+v", report)
	}
	destination := filepath.Join(t.TempDir(), "reimport")
	if _, err := Import(report.Artifact, destination); err != nil {
		t.Fatal(err)
	}
	actualRibbon, err := os.ReadFile(filepath.Join(destination, "package", "customUI", "customUI14.xml"))
	if err != nil || !bytes.Equal(ribbon, actualRibbon) {
		t.Fatalf("Merged Ribbon source changed or lost: %v", err)
	}
	actualForm, err := os.ReadFile(filepath.Join(destination, "forms", "MergeForm.json"))
	if err != nil {
		t.Fatal(err)
	}
	var importedDesign office.Design
	if err := json.Unmarshal(actualForm, &importedDesign); err != nil {
		t.Fatal(err)
	}
	if len(importedDesign.Controls) != 1 || importedDesign.Controls[0].Name != "Run" || importedDesign.Controls[0].Properties["Caption"] != "After merge" {
		t.Fatalf("Merged form edit changed or lost: %s", actualForm)
	}
	for _, name := range []string{"First", "Second"} {
		data, err := os.ReadFile(filepath.Join(destination, "vba", name+".bas"))
		if err != nil || !strings.Contains(string(data), "Function "+name+"()") {
			t.Fatalf("Merged source lost: %s %v", name, err)
		}
	}
	for _, autocrlf := range []string{"true", "input", "false"} {
		clone := filepath.Join(t.TempDir(), "checkout")
		git("-c", "core.autocrlf="+autocrlf, "clone", "-q", "--no-hardlinks", w.Root, clone)
		for _, name := range strings.Split(strings.TrimSuffix(git("ls-files", "-z"), "\x00"), "\x00") {
			before, err := os.ReadFile(filepath.Join(w.Root, name))
			if err != nil {
				t.Fatal(err)
			}
			after, err := os.ReadFile(filepath.Join(clone, name))
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(before, after) {
				t.Errorf("autocrlf=%s changed %s", autocrlf, name)
			}
		}
		checked, err := Open(clone)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := checked.Build(""); err != nil {
			t.Fatalf("autocrlf=%s checkout build: %v", autocrlf, err)
		}
	}
	// The same source line must remain an ordinary, visible Git conflict.
	// Do not install a merge driver that silently chooses one implementation.
	for _, branch := range []string{"conflict-left", "conflict-right"} {
		git("checkout", "-qb", branch, "second")
		source := "Public Function First() As Long\nFirst = 2 ' " + branch + "\nEnd Function\n"
		if err := Write(w.Root, "vba/First.bas", []byte(source), ""); err != nil {
			t.Fatal(err)
		}
		git("add", "vba/First.bas")
		git("commit", "-qm", branch)
	}
	conflict := exec.Command("git", "-c", "user.name=WordUp Test", "-c", "user.email=test@example.invalid", "merge", "--no-edit", "conflict-left")
	conflict.Dir = w.Root
	if out, err := conflict.CombinedOutput(); err == nil {
		t.Fatalf("Overlapping source edits merged silently: %s", out)
	}
	if status := git("status", "--porcelain", "--", "vba/First.bas"); !strings.HasPrefix(status, "UU ") {
		t.Fatalf("Expected an unmerged source file, got %q", status)
	}
	conflicted, err := os.ReadFile(filepath.Join(w.Root, "vba", "First.bas"))
	if err != nil {
		t.Fatal(err)
	}
	for _, marker := range []string{"<<<<<<<", "=======", ">>>>>>>", "' conflict-left", "' conflict-right"} {
		if !bytes.Contains(conflicted, []byte(marker)) {
			t.Errorf("Conflict discarded source or marker %q", marker)
		}
	}
}
