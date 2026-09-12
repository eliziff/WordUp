package example

import (
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestIntegratedExampleBuild(t *testing.T) {
	r, e := Studio(filepath.Join(t.TempDir(), "studio"))
	if e != nil {
		t.Fatal(e)
	}
	if r.Modules != 4 || r.Forms != 1 || r.WordExecuted || r.VBACompiled {
		t.Fatalf("bad evidence %+v", r)
	}
	b, _ := os.ReadFile(r.Artifact)
	p, e := office.ReadPackage(b)
	if e != nil {
		t.Fatal(e)
	}
	if e = p.Validate(); e != nil {
		t.Fatal(e)
	}
	for file, text := range map[string]string{"word/document.xml": "Editorial Studio", "word/styles.xml": "StudioTitle", "word/footnotes.xml": "footnoteRef", "word/glossary/document.xml": "Studio Editorial Note", "customUI/customUI14.xml": "OnRibbonLoad"} {
		if !strings.Contains(string(p.Files[file]), text) {
			t.Fatalf("missing %s in %s", text, file)
		}
	}
	v, e := office.ReadVBA(p.Files["word/vbaProject.bin"])
	if e != nil {
		t.Fatal(e)
	}
	f, e := office.ReadForm(v.CFB, "StudioForm", v.Codepage)
	if e != nil {
		t.Fatal(e)
	}
	if len(f.ControlNames()) < 6 {
		t.Fatal("missing actual nested form controls")
	}
	var suite map[string]any
	s, _ := project.Read(filepath.Dir(filepath.Dir(r.Artifact)), "tests/suite.json")
	if e = project.ReadJSON(s, &suite); e != nil {
		t.Fatal(e)
	}
}
