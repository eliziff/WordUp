package component

import (
	"bytes"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/vbaparse"
	"os"
	"path/filepath"
	"strings"
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
	root := filepath.Join(t.TempDir(), "workspace")
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

func TestInstallationRejectsExportedIdentifierCollision(t *testing.T) {
	root := t.TempDir()
	if err := project.Write(root, "vba/Existing.bas", []byte("Public Sub WU_BeginSafeEdit()\nEnd Sub\n"), ""); err != nil {
		t.Fatal(err)
	}
	if _, err := Add(root, "operation.safe-edit"); err == nil || !strings.Contains(err.Error(), `identifier "WU_BeginSafeEdit"`) || !strings.Contains(err.Error(), "Existing.bas:1") {
		t.Fatalf("collision was not precise: %v", err)
	}
	if _, err := project.Read(root, "vba/WordUpSafeEdit.bas"); !os.IsNotExist(err) {
		t.Fatalf("component source was written after collision: %v", err)
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
	platforms, ok := status["supported_platforms"].([]string)
	if !ok || len(platforms) != 1 || platforms[0] != "windows" {
		t.Fatalf("status lost platform declaration: %#v", status)
	}
	if compatibility, ok := status["compatibility"].(string); !ok || compatibility == "" {
		t.Fatalf("status lost compatibility: %#v", status)
	}
	if _, err = Add(root, "structure.detect"); err == nil {
		t.Fatal("edited component overwritten")
	}
}

func TestAddAppliesOnlyDeclaredTypedParameters(t *testing.T) {
	root := t.TempDir()
	installed, err := AddWith(root, "operation.safe-edit", map[string]string{"module_prefix": "ACME"})
	if err != nil {
		t.Fatal(err)
	}
	source, err := project.Read(root, "vba/WordUpSafeEdit.bas")
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(source), "Public Sub ACME_BeginSafeEdit") || strings.Contains(string(source), "Public Sub WU_BeginSafeEdit") {
		t.Fatal("module prefix was not applied to the VBA entry point")
	}
	if installed.Parameters["module_prefix"] != "ACME" {
		t.Fatalf("parameters not recorded: %#v", installed.Parameters)
	}
	status, err := Status(root, installed.ID)
	if err != nil {
		t.Fatal(err)
	}
	parameters, ok := status["parameters"].(map[string]string)
	if !ok || parameters["module_prefix"] != "ACME" {
		t.Fatalf("status parameters=%#v", status["parameters"])
	}
	difference, err := Diff(root, installed.ID)
	if err != nil {
		t.Fatal(err)
	}
	if difference["equal"] != true {
		t.Fatalf("unchanged component diff is not equal: %#v", difference)
	}
	files := difference["files"].([]map[string]any)
	if len(files) != 1 || files[0]["state"] != "clean" || files[0]["bundled_sha256"] != files[0]["installed_sha256"] {
		t.Fatalf("diff did not compare installed parameterized source: %#v", difference)
	}
	if _, exists := files[0]["text"]; exists {
		t.Fatal("component diff duplicated source text")
	}
	if _, err = AddWith(root, installed.ID, map[string]string{"module_prefix": "ACME"}); err != nil {
		t.Fatal("same parameterization is not idempotent:", err)
	}
	if _, err = AddWith(root, installed.ID, map[string]string{"module_prefix": "Other"}); err == nil {
		t.Fatal("different parameterization was treated as idempotent")
	}
}

func TestModulePrefixOnlyChangesIdentifierTokens(t *testing.T) {
	m := Manifest{ID: "tokens", Version: "1", Parameters: map[string]string{"module_prefix": "prefix"}, Defaults: map[string]string{"module_prefix": "WU"}, Files: []File{{Path: "vba/Tokens.bas", Text: "Public Sub WU_Run()\n    Dim value As String\n    value = \"WU_String\" & \"WU_Command_\"\n    value = \"XWU_Unchanged\"\n    ' WU_Comment\n    Rem WU_RemComment\nEnd Sub\n"}}, Acceptance: "WU_Run is the entry point."}
	adapted, err := adapt(m, map[string]string{"module_prefix": "ACME"})
	if err != nil {
		t.Fatal(err)
	}
	source := adapted.Files[0].Text
	for _, want := range []string{"Public Sub ACME_Run", `"ACME_String"`, `"ACME_Command_"`, `"XWU_Unchanged"`, "' WU_Comment", "Rem WU_RemComment"} {
		if !strings.Contains(source, want) {
			t.Fatalf("adaptation lost %q: %s", want, source)
		}
	}
	if strings.Contains(source, "WU_Run") || adapted.Acceptance != "ACME_Run is the entry point." {
		t.Fatalf("identifier adaptation was incomplete: source=%s acceptance=%q", source, adapted.Acceptance)
	}
}

func TestAddRejectsInvalidOrUndeclaredParametersBeforeWriting(t *testing.T) {
	for name, values := range map[string]map[string]string{
		"invalid identifier": {"module_prefix": "not-valid"},
		"undeclared":         {"label": "Surprise"},
	} {
		t.Run(name, func(t *testing.T) {
			root := t.TempDir()
			if _, err := AddWith(root, "ui.progress-cancel", values); err == nil {
				t.Fatal("invalid parameters accepted")
			}
			if _, err := project.Read(root, "vba/WordUpProgress.bas"); !os.IsNotExist(err) {
				t.Fatalf("source written after rejected parameters: %v", err)
			}
		})
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

func TestAllBundledComponentsInstallWithoutCollisions(t *testing.T) {
	root := t.TempDir()
	for _, manifest := range builtin() {
		if _, err := Add(root, manifest.ID); err != nil {
			t.Fatalf("%s: %v", manifest.ID, err)
		}
	}
}

func TestLocalBundleInstallsAndDiffsWithoutEmbeddingSource(t *testing.T) {
	bundle, root := t.TempDir(), t.TempDir()
	manifest := Manifest{Schema: 1, ID: "example.local", Version: "1.2.3", License: "MIT", Provenance: "local test bundle", Files: []File{{Path: "vba/Local.bas"}, {Path: "assets/icon.bin"}, {Path: "package/media/icon.png"}}}
	if err := project.Write(bundle, "component.json", project.JSON(manifest), ""); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(bundle, "vba/Local.bas", []byte("Attribute VB_Name = \"Local\"\nOption Explicit\n"), ""); err != nil {
		t.Fatal(err)
	}
	icon := []byte{0, 1, 2, 255, 0}
	if err := project.Write(bundle, "assets/icon.bin", icon, ""); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(bundle, "package/media/icon.png", icon, ""); err != nil {
		t.Fatal(err)
	}
	loaded, err := LoadBundle(bundle)
	if err != nil || !loaded.Files[1].Binary || !bytes.Equal(loaded.Files[1].data, icon) {
		t.Fatalf("binary bundle source was not preserved: %#v (%v)", loaded.Files, err)
	}
	installed, err := AddBundle(root, bundle)
	if err != nil {
		t.Fatal(err)
	}
	if installed.ID != manifest.ID || installed.Provenance != manifest.Provenance {
		t.Fatalf("wrong installed provenance: %#v", installed)
	}
	if actual, err := project.Read(root, "assets/icon.bin"); err != nil || !bytes.Equal(actual, icon) {
		t.Fatalf("binary source was not installed unchanged: %v", err)
	}
	if actual, err := project.Read(root, "package/media/icon.png"); err != nil || !bytes.Equal(actual, icon) {
		t.Fatalf("binary package source was not installed unchanged: %v", err)
	}
	if _, err = AddBundle(root, bundle); err != nil {
		t.Fatal("unchanged local bundle is not idempotent:", err)
	}
	diff, err := DiffBundle(root, bundle)
	if err != nil || diff["component"] != manifest.ID {
		t.Fatalf("local diff=%#v err=%v", diff, err)
	}
	if err = project.Write(root, "vba/Local.bas", []byte("adapted"), ""); err != nil {
		t.Fatal(err)
	}
	diff, err = DiffBundle(root, bundle)
	if err != nil || diff["equal"] != false {
		t.Fatalf("modified local bundle diff=%#v err=%v", diff, err)
	}
	if _, err = AddBundle(root, bundle); err == nil {
		t.Fatal("local bundle overwrote adapted source")
	}
}

func TestLocalBundleRejectsUnsafeSources(t *testing.T) {
	for _, file := range []File{{Path: "../escape.bas"}, {Path: "vba/Embedded.bas", Text: "embedded"}} {
		t.Run(file.Path, func(t *testing.T) {
			bundle := t.TempDir()
			manifest := Manifest{Schema: 1, ID: "unsafe", Version: "1", License: "MIT", Provenance: "test", Files: []File{file}}
			if err := project.Write(bundle, "component.json", project.JSON(manifest), ""); err != nil {
				t.Fatal(err)
			}
			if _, err := LoadBundle(bundle); err == nil {
				t.Fatal("unsafe local bundle accepted")
			}
		})
	}
}

func TestManifestHashIncludesBinarySourceBytes(t *testing.T) {
	base := Manifest{Schema: 1, ID: "binary.hash", Version: "1", License: "MIT", Provenance: "test", Files: []File{{Path: "assets/icon.bin", Binary: true, data: []byte{1, 2, 3}}}}
	changed := base
	changed.Files = append([]File(nil), base.Files...)
	changed.Files[0].data = []byte{1, 2, 4}
	if manifestHash(base) == manifestHash(changed) {
		t.Fatal("manifest hash ignored binary component bytes")
	}
}

func TestRibbonMergeBundleIsPreflightedAndBuilt(t *testing.T) {
	const ns = "http://schemas.microsoft.com/office/2009/07/customui"
	base := []byte(`<customUI xmlns="` + ns + `"><ribbon><tabs><tab id="base"/></tabs></ribbon></customUI>`)
	fragment := []byte(`<customUI xmlns="` + ns + `"><ribbon><tabs><tab id="added"><group id="group"><button id="button" label="Run"/></group></tab></tabs></ribbon></customUI>`)
	conflict := []byte(`<customUI xmlns="` + ns + `"><ribbon><tabs><tab id="base" label="Conflict"/></tabs></ribbon></customUI>`)
	bundle := t.TempDir()
	manifest := Manifest{Schema: 1, ID: "ribbon.local", Version: "1", License: "MIT", Provenance: "local Ribbon test", Files: []File{{Path: "assets/fragment.xml"}}, RibbonMerges: []RibbonMerge{{Source: "assets/fragment.xml", Target: "customUI/customUI14.xml"}}}
	if err := project.Write(bundle, "component.json", project.JSON(manifest), ""); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(bundle, "assets/fragment.xml", fragment, ""); err != nil {
		t.Fatal(err)
	}
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("RibbonProject", root); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "package/customUI/customUI14.xml", base, ""); err != nil {
		t.Fatal(err)
	}
	installed, err := AddBundle(root, bundle)
	if err != nil {
		t.Fatal(err)
	}
	if len(installed.RibbonMerges) != 1 || installed.RibbonMerges[0].Target != "customUI/customUI14.xml" {
		t.Fatalf("Ribbon mapping was not recorded: %#v", installed)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	report, err := w.Build(filepath.Join(root, "dist", "RibbonProject.dotm"))
	if err != nil {
		t.Fatal(err)
	}
	artifact, err := os.ReadFile(report.Artifact)
	if err != nil {
		t.Fatal(err)
	}
	pkg, err := office.ReadPackage(artifact)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(pkg.Files["customUI/customUI14.xml"]), `id="button"`) {
		t.Fatal("Ribbon fragment was not composed into the built package")
	}
	status, err := Status(root, manifest.ID)
	if err != nil || status["ribbon_merges"] == nil {
		t.Fatalf("Ribbon mapping missing from status: %#v (%v)", status, err)
	}

	conflictBundle := t.TempDir()
	manifest.ID = "ribbon.conflict"
	if err := project.Write(conflictBundle, "component.json", project.JSON(manifest), ""); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(conflictBundle, "assets/fragment.xml", conflict, ""); err != nil {
		t.Fatal(err)
	}
	conflictRoot := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("RibbonConflict", conflictRoot); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(conflictRoot, "package/customUI/customUI14.xml", base, ""); err != nil {
		t.Fatal(err)
	}
	if _, err := AddBundle(conflictRoot, conflictBundle); err == nil || !strings.Contains(err.Error(), "Ribbon merge assets/fragment.xml") {
		t.Fatalf("Ribbon collision was not rejected before copy: %v", err)
	}
	if _, err := project.Read(conflictRoot, "assets/fragment.xml"); !os.IsNotExist(err) {
		t.Fatalf("Ribbon source copied after rejected preflight: %v", err)
	}
}
