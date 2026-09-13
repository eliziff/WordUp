package component

import (
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/vbaparse"
	"os"
	"path/filepath"
	"testing"
)

func TestInstallationPreflightsAllFiles(t *testing.T) {
	for _, badPath := range []string{"vba/existing.bas", "../escape.bas", "vba/FIRST.bas", ".wordwright/components.json"} {
		t.Run(badPath, func(t *testing.T) {
			root := t.TempDir()
			if err := project.Write(root, "vba/existing.bas", []byte("user source"), ""); err != nil {
				t.Fatal(err)
			}
			m := Manifest{ID: "test", Version: "1", Files: []File{{Path: "vba/first.bas", Text: "first"}, {Path: badPath, Text: "second"}}}
			if _, err := install(root, m); err == nil {
				t.Fatal("invalid installation accepted")
			}
			if _, err := project.Read(root, "vba/first.bas"); !os.IsNotExist(err) {
				t.Fatalf("partial installation: %v", err)
			}
			b, err := project.Read(root, "vba/existing.bas")
			if err != nil || string(b) != "user source" {
				t.Fatal("user source changed")
			}
		})
	}
}

func TestInstallationRollsBackCompletedCopies(t *testing.T) {
	root := t.TempDir()
	// A parent file prevents directory creation after an earlier copy succeeds.
	if err := project.Write(root, "blocked", []byte("user file"), ""); err != nil {
		t.Fatal(err)
	}
	m := Manifest{ID: "test", Version: "1", Files: []File{{Path: "vba/first.bas", Text: "first"}, {Path: "blocked/second.bas", Text: "second"}}}
	if _, err := install(root, m); err == nil {
		t.Fatal("blocked installation accepted")
	}
	if _, err := project.Read(root, "vba/first.bas"); !os.IsNotExist(err) {
		t.Fatalf("partial installation: %v", err)
	}
	if _, err := project.Read(root, lockPath); !os.IsNotExist(err) {
		t.Fatalf("false provenance: %v", err)
	}
}

func TestInstallationRejectsAliasedAndReservedPathsBeforeCopying(t *testing.T) {
	for _, path := range []string{"./.wordwright/forged.json", "vba/../.wordwright/forged.json", `.wordwright\forged.json`, ".WORDWRIGHT/forged.json", ".git/config", ".GIT/hooks/pre-commit", "vba//Module.bas", "vba/./Module.bas", "vba/Module.bas:stream", "../outside.bas", "/absolute.bas", "."} {
		t.Run(path, func(t *testing.T) {
			root := t.TempDir()
			m := Manifest{ID: "test", Version: "1", Files: []File{{Path: "vba/first.bas", Text: "first"}, {Path: path, Text: "unexpected"}}}
			if _, err := install(root, m); err == nil {
				t.Fatal("unsafe component path accepted")
			}
			if _, err := project.Read(root, "vba/first.bas"); !os.IsNotExist(err) {
				t.Fatalf("preflight wrote source: %v", err)
			}
			if _, err := project.Read(root, lockPath); !os.IsNotExist(err) {
				t.Fatalf("preflight wrote provenance: %v", err)
			}
		})
	}
}

func TestSameVersionDifferentSourceIsNotIdempotent(t *testing.T) {
	root := t.TempDir()
	m := Manifest{ID: "test", Version: "1", Files: []File{{Path: "vba/first.bas", Text: "first"}}}
	if _, err := install(root, m); err != nil {
		t.Fatal(err)
	}
	m.Files[0].Text = "changed without version bump"
	if _, err := install(root, m); err == nil {
		t.Fatal("different source treated as unchanged")
	}
}

func TestInstallationLockProtectsExistingInstaller(t *testing.T) {
	root := t.TempDir()
	if err := project.Write(root, ".wordwright/component-install.lock", []byte("another installer"), ""); err != nil {
		t.Fatal(err)
	}
	if _, err := Add(root, "operation.safe-edit"); err == nil {
		t.Fatal("concurrent installation accepted")
	}
	b, err := project.Read(root, ".wordwright/component-install.lock")
	if err != nil || string(b) != "another installer" {
		t.Fatal("another installer's lock changed")
	}
}

func TestBundledVBADeclarations(t *testing.T) {
	for _, manifest := range builtin() {
		for _, file := range manifest.Files {
			result, err := vbaparse.Parse(file.Text)
			if err != nil {
				t.Fatal(err)
			}
			if diagnostics := result["diagnostics"].([]vbaparse.Diagnostic); len(diagnostics) != 0 {
				t.Errorf("%s %s: %v", manifest.ID, file.Path, diagnostics)
			}
		}
	}
}

func TestAddTracksAndProtectsEditedComponent(t *testing.T) {
	root := t.TempDir()
	installed, err := Add(root, "structure.detect")
	if err != nil {
		t.Fatal(err)
	}
	if len(installed.Files) != 1 {
		t.Fatalf("files=%d", len(installed.Files))
	}
	if _, err = Add(root, "structure.detect"); err != nil {
		t.Fatal("idempotent add:", err)
	}
	file := filepath.Join(root, "vba", "WordUpStructure.bas")
	if err = os.WriteFile(file, []byte("edited"), 0600); err != nil {
		t.Fatal(err)
	}
	status, err := Status(root, "structure.detect")
	if err != nil || status["state"] != "modified" {
		t.Fatalf("status=%v err=%v", status, err)
	}
	if _, err = Add(root, "structure.detect"); err == nil {
		t.Fatal("edited component overwritten")
	}
}

func TestBundledCatalogIsComplete(t *testing.T) {
	want := []string{"structure.detect", "operation.safe-edit", "ui.form-shell", "ui.progress-cancel", "ui.ribbon-command", "command.hotkey", "command.context-menu", "document.style-converter"}
	items := List()
	if len(items) != len(want) {
		t.Fatalf("components=%d, want %d", len(items), len(want))
	}
	for i, id := range want {
		if items[i].ID != id {
			t.Fatalf("component %d=%q, want %q", i, items[i].ID, id)
		}
		if items[i].Provenance == "" || items[i].Acceptance == "" || len(items[i].SupportedPlatforms) == 0 {
			t.Fatalf("incomplete manifest: %#v", items[i])
		}
	}
}
