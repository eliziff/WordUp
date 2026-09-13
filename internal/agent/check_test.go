package agent

import (
	"context"
	"github.com/eliziff/WordUp/internal/project"
	"path/filepath"
	"testing"
)

func TestProjectCheckUsesRequestedConditionalConstants(t *testing.T) {
	root := filepath.Join(t.TempDir(), "source")
	if _, err := project.New("ConditionalDemo", root); err != nil {
		t.Fatal(err)
	}
	source := []byte("Attribute VB_Name = \"Branch\"\n#If CustomBranch Then\nPublic Sub Example()\nEnd Sub\n#Else\nPublic Sub Example()\nDim x As\nEnd Sub\n#End If\n")
	if err := project.Write(root, "vba/Branch.bas", source, ""); err != nil {
		t.Fatal(err)
	}
	e := &Engine{Root: root}
	for _, enabled := range []bool{true, false} {
		result, err := e.Call(context.Background(), "check", Parameters{CompilationConstants: map[string]any{"CustomBranch": enabled}})
		if err != nil {
			t.Fatal(err)
		}
		diagnostics := result.(map[string]any)["diagnostics"].([]map[string]any)
		found := false
		for _, d := range diagnostics {
			if d["file"] == "vba/Branch.bas" {
				found = true
			}
		}
		if found == enabled {
			t.Fatalf("wrong branch checked: enabled=%v diagnostics=%v", enabled, diagnostics)
		}
	}
}
