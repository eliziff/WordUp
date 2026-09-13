package project

import (
	"bytes"
	"io/fs"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

func TestRepeatedImportHasIdenticalFiles(t *testing.T) {
	w := newWorkspace(t)
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
}

func TestGitNoOpBuildAndIndependentModuleMerge(t *testing.T) {
	if _, err := exec.LookPath("git"); err != nil {
		t.Skip("git is required for checkout/merge verification")
	}
	w := newWorkspace(t)
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
	git("add", ".gitattributes", ".gitignore", "project.json", "vba", "package", "forms", "styles", "content", "building_blocks", ".wordwright", "tests", "assets")
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
	git("commit", "-qm", "first module")
	git("checkout", "-qb", "second", "baseline")
	if err := Write(w.Root, "vba/Second.bas", []byte("Public Function Second() As Long\nSecond = 2\nEnd Function\n"), ""); err != nil {
		t.Fatal(err)
	}
	git("add", "vba/Second.bas")
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
	if report.Modules != 3 {
		t.Fatalf("Merged modules missing: %+v", report)
	}
	destination := filepath.Join(t.TempDir(), "reimport")
	if _, err := Import(report.Artifact, destination); err != nil {
		t.Fatal(err)
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
}
