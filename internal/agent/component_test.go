package agent

import (
	"context"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/component"
	"github.com/eliziff/WordUp/internal/project"
)

func TestComponentAddAcceptsLocalBundleRelativeToWorkspace(t *testing.T) {
	parent := t.TempDir()
	root, bundle := filepath.Join(parent, "workspace"), filepath.Join(parent, "bundle")
	manifest := component.Manifest{Schema: 1, ID: "agent.local", Version: "1", License: "MIT", Provenance: "agent test", Files: []component.File{{Path: "vba/AgentLocal.bas"}}}
	if err := project.Write(bundle, "component.json", project.JSON(manifest), ""); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(bundle, "vba/AgentLocal.bas", []byte("Attribute VB_Name = \"AgentLocal\"\n"), ""); err != nil {
		t.Fatal(err)
	}
	engine := &Engine{Root: root}
	result, err := engine.Call(context.Background(), "component.add", Parameters{Path: "../bundle"})
	if err != nil {
		t.Fatal(err)
	}
	if result.(component.Installed).ID != manifest.ID {
		t.Fatalf("wrong component installed: %#v", result)
	}
}

func TestComponentParametersFlowThroughAgent(t *testing.T) {
	root := t.TempDir()
	engine := &Engine{Root: root}
	result, err := engine.Call(context.Background(), "component.add", Parameters{
		Component: "batch.progress", Parameters: map[string]string{"module_prefix": "Fast"},
	})
	if err != nil {
		t.Fatal(err)
	}
	if result.(component.Installed).Parameters["module_prefix"] != "Fast" {
		t.Fatalf("parameters not returned: %#v", result)
	}
	source, err := os.ReadFile(filepath.Join(root, "vba", "WordUpBatch.bas"))
	if err != nil || !strings.Contains(string(source), "Fast_CancelRequested") {
		t.Fatalf("parameterized source missing: %v", err)
	}
}
