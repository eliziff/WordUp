package agent

import (
	"context"
	"strings"
	"testing"
)

func TestCoreProfileHidesModelFacingTools(t *testing.T) {
	defer SetToolProfile("")
	full := Tools()
	if err := SetToolProfile("core"); err != nil {
		t.Fatal(err)
	}
	core := Tools()
	if len(core) >= len(full) {
		t.Fatalf("core profile exposes %d tools, full exposes %d", len(core), len(full))
	}
	for _, tool := range core {
		if strings.HasPrefix(tool.Name, "component.") || strings.HasPrefix(tool.Name, "journal.") || strings.HasPrefix(tool.Name, "structure.") {
			t.Fatalf("core profile still advertises %s", tool.Name)
		}
	}
	for _, name := range []string{"build", "check", "native.call", "native.probe", "test", "xml.patch"} {
		found := false
		for _, tool := range core {
			found = found || tool.Name == name
		}
		if !found {
			t.Fatalf("core profile lost core tool %s", name)
		}
	}
	e := &Engine{Root: t.TempDir()}
	if _, err := e.Call(context.Background(), "component.list", Parameters{}); err == nil || !strings.Contains(err.Error(), "profile") {
		t.Fatalf("hidden tool dispatched under core profile: %v", err)
	}
	if err := SetToolProfile("bogus"); err == nil {
		t.Fatal("unknown profile accepted")
	}
	if ToolProfile() != "core" {
		t.Fatalf("unexpected profile %q", ToolProfile())
	}
}
