package agent

import (
	"context"
	"path/filepath"
	"testing"

	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

func TestResidentProjectCheckReusesSourceSnapshot(t *testing.T) {
	root := filepath.Join(t.TempDir(), "source")
	if _, err := project.New("ResidentCheck", root); err != nil {
		t.Fatal(err)
	}
	e := &Engine{Root: root}
	defer e.Close()
	if _, err := e.Call(context.Background(), "check", Parameters{}); err != nil {
		t.Fatal(err)
	}
	workspace := e.workspace
	if workspace == nil {
		t.Fatal("check did not retain its resident source snapshot")
	}
	if _, err := e.Call(context.Background(), "check", Parameters{}); err != nil {
		t.Fatal(err)
	}
	if e.workspace != workspace {
		t.Fatal("unchanged check reopened the workspace")
	}
	if err := project.Write(root, "vba/Added.bas", []byte("Attribute VB_Name = \"Added\"\nPublic Sub AddedMacro()\nEnd Sub\n"), ""); err != nil {
		t.Fatal(err)
	}
	result, err := e.Call(context.Background(), "check", Parameters{})
	if err != nil {
		t.Fatal(err)
	}
	if e.workspace != workspace {
		t.Fatal("source edit discarded the resident workspace")
	}
	if parsed, ok := result.(map[string]any)["syntax_modules_parsed"].(int); !ok || parsed != 2 {
		t.Fatalf("source edit did not refresh the resident snapshot: %#v", result)
	}
}

func TestResidentProjectCheckRefreshesWorkspaceMetadata(t *testing.T) {
	root := filepath.Join(t.TempDir(), "source")
	if _, err := project.New("MetadataCheck", root); err != nil {
		t.Fatal(err)
	}
	e := &Engine{Root: root}
	defer e.Close()
	if _, err := e.Call(context.Background(), "check", Parameters{}); err != nil {
		t.Fatal(err)
	}
	first := e.workspace
	projectJSON, err := project.Read(root, "project.json")
	if err != nil {
		t.Fatal(err)
	}
	updated := []byte("{\n  \"name\": \"MetadataCheckRenamed\"\n}\n")
	if err := project.Write(root, "project.json", updated, projectJSONHash(projectJSON)); err != nil {
		t.Fatal(err)
	}
	if _, err := e.Call(context.Background(), "check", Parameters{}); err != nil {
		t.Fatal(err)
	}
	if e.workspace != first {
		t.Fatal("metadata edit discarded the resident workspace instead of refreshing it")
	}
	if e.workspace.Manifest.Name != "MetadataCheckRenamed" {
		t.Fatalf("workspace metadata was not refreshed: %#v", e.workspace.Manifest)
	}
}

func TestResidentFilesAndSearchRefreshSource(t *testing.T) {
	root := filepath.Join(t.TempDir(), "source")
	if _, err := project.New("ResidentFiles", root); err != nil {
		t.Fatal(err)
	}
	e := &Engine{Root: root}
	defer e.Close()
	if _, err := e.Call(context.Background(), "files", Parameters{}); err != nil {
		t.Fatal(err)
	}
	workspace := e.workspace
	if workspace == nil {
		t.Fatal("files did not retain its resident workspace")
	}
	if _, err := e.Call(context.Background(), "search", Parameters{Query: "ThisDocument"}); err != nil {
		t.Fatal(err)
	}
	if e.workspace != workspace {
		t.Fatal("unchanged search reopened the workspace")
	}
	if err := project.Write(root, "vba/Added.bas", []byte("Attribute VB_Name = \"Added\"\nPublic Sub AddedMacro()\nEnd Sub\n"), ""); err != nil {
		t.Fatal(err)
	}
	result, err := e.Call(context.Background(), "search", Parameters{Query: "AddedMacro"})
	if err != nil {
		t.Fatal(err)
	}
	if e.workspace != workspace {
		t.Fatal("source edit discarded the resident files/search workspace")
	}
	if total, ok := result.(map[string]any)["total"].(int); !ok || total != 1 {
		t.Fatalf("source edit did not refresh files/search snapshot: %#v", result)
	}
}

func projectJSONHash(b []byte) string {
	// Keep this test independent of the agent's file-writing path while still
	// exercising its optimistic-concurrency guard.
	return office.Hash(b)
}

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
