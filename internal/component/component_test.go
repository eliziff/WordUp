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

func TestInstallationRejectsCaseVariantSourceDirectory(t *testing.T) {
	root := t.TempDir()
	if err := os.Mkdir(filepath.Join(root, "VBA"), 0700); err != nil {
		t.Fatal(err)
	}
	m := Manifest{ID: "case-variant", Version: "1", Files: []File{{Path: "vba/New.bas", Text: "Option Explicit\n"}}}
	if _, err := install(root, m); err == nil || !strings.Contains(err.Error(), `workspace source directory must be named "vba" (found "VBA")`) {
		t.Fatalf("case-variant source directory was accepted: %v", err)
	}
	if _, err := project.Read(root, "vba/New.bas"); !os.IsNotExist(err) {
		t.Fatalf("case-variant install wrote source: %v", err)
	}
}

func TestStatusRejectsCaseVariantSourceDirectory(t *testing.T) {
	root := t.TempDir()
	if err := os.Mkdir(filepath.Join(root, "VBA"), 0700); err != nil {
		t.Fatal(err)
	}
	if _, err := Status(root, "operation.safe-edit"); err == nil || !strings.Contains(err.Error(), `workspace source directory must be named "vba" (found "VBA")`) {
		t.Fatalf("status accepted case-variant source directory: %v", err)
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

func TestInstallationRejectsModuleNameCollision(t *testing.T) {
	root := t.TempDir()
	if err := project.Write(root, "vba/Existing.bas", []byte("Attribute VB_Name = \"WordUpSafeEdit\"\nPublic Sub OtherMacro()\nEnd Sub\n"), ""); err != nil {
		t.Fatal(err)
	}
	if _, err := Add(root, "operation.safe-edit"); err == nil || !strings.Contains(err.Error(), `module name "WordUpSafeEdit"`) || !strings.Contains(err.Error(), "Existing.bas:1") {
		t.Fatalf("module collision was not precise: %v", err)
	}
	if _, err := project.Read(root, "vba/WordUpSafeEdit.bas"); !os.IsNotExist(err) {
		t.Fatalf("component source was written after module collision: %v", err)
	}
}

func TestInstallationRejectsDerivedModuleNameCollision(t *testing.T) {
	root := t.TempDir()
	if err := project.Write(root, "vba/WordUpSafeEdit.cls", []byte("Public Sub ExistingMacro()\nEnd Sub\n"), ""); err != nil {
		t.Fatal(err)
	}
	if _, err := Add(root, "operation.safe-edit"); err == nil || !strings.Contains(err.Error(), `module name "WordUpSafeEdit"`) || !strings.Contains(err.Error(), "WordUpSafeEdit.cls:1") {
		t.Fatalf("derived module collision was not precise: %v", err)
	}
	if _, err := project.Read(root, "vba/WordUpSafeEdit.bas"); !os.IsNotExist(err) {
		t.Fatalf("component source was written after derived module collision: %v", err)
	}
}

func TestInstallationRejectsModuleNameMismatch(t *testing.T) {
	root := t.TempDir()
	m := Manifest{ID: "mismatch", Version: "1", Files: []File{{Path: "vba/Expected.bas", Text: "Attribute VB_Name = \"Other\"\nPublic Sub Run()\nEnd Sub\n"}}}
	if _, err := install(root, m); err == nil || !strings.Contains(err.Error(), `VB_Name "Other"`) || !strings.Contains(err.Error(), `Expected.bas:1`) {
		t.Fatalf("module name mismatch was not rejected precisely: %v", err)
	}
	if _, err := project.Read(root, "vba/Expected.bas"); !os.IsNotExist(err) {
		t.Fatalf("mismatched component source was written: %v", err)
	}
}

func TestInstallationRejectsInvalidVBNameAttribute(t *testing.T) {
	for name, source := range map[string]string{
		"invalid identifier": "Attribute VB_Name = \"Bad-Name\"\nOption Explicit\n",
		"malformed value":    "Attribute VB_Name = BadName\nOption Explicit\n",
	} {
		t.Run(name, func(t *testing.T) {
			root := t.TempDir()
			m := Manifest{ID: "invalid-vb-name", Version: "1", Files: []File{{Path: "vba/Good.bas", Text: source}}}
			if _, err := install(root, m); err == nil || !strings.Contains(err.Error(), "invalid VB_Name attribute") {
				t.Fatalf("invalid VB_Name attribute was accepted: %v", err)
			}
			if _, err := project.Read(root, "vba/Good.bas"); !os.IsNotExist(err) {
				t.Fatalf("invalid component source was written: %v", err)
			}
		})
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

func TestHotkeyStatusScopesTemplateContext(t *testing.T) {
	manifest, err := Get("command.hotkey")
	if err != nil {
		t.Fatal(err)
	}
	source := manifest.Files[0].Text
	start := strings.Index(source, "Public Function WU_HotkeyRegistered")
	if start < 0 {
		t.Fatal("hotkey status function is not present")
	}
	end := strings.Index(source[start:], "End Function")
	if end < 0 {
		t.Fatal("hotkey status function has no terminator")
	}
	body := source[start : start+end]
	context := strings.Index(body, "Application.CustomizationContext = ThisDocument")
	lookup := strings.Index(body, "WU_OwnedHotkey(keyCode)")
	if context < 0 || lookup < 0 || context > lookup {
		t.Fatal("hotkey status must inspect bindings in the template customization context")
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
	if status["version_state"] != "current" || status["bundled_version"] != installed.Version || status["bundled_manifest_sha256"] != installed.ManifestSHA256 {
		t.Fatalf("status lost bundled version provenance: %#v", status)
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
	want := []string{"structure.detect", "operation.safe-edit", "ui.form-shell", "ui.progress-cancel", "ui.ribbon-command", "command.hotkey", "command.context-menu", "document.style-converter", "document.text-operations"}
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

func TestStyleConverterGuardsInputsAndStateCapture(t *testing.T) {
	item, err := Get("document.style-converter")
	if err != nil {
		t.Fatal(err)
	}
	if item.Version != "1.0.4" {
		t.Fatalf("style converter version did not advance: %q", item.Version)
	}
	for _, capability := range []string{"paragraph-style conversion", "range-bounded conversion", "story-wide conversion"} {
		found := false
		for _, got := range item.Capabilities {
			if got == capability {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("style converter manifest omitted capability %q: %#v", capability, item.Capabilities)
		}
	}
	var source string
	for _, file := range item.Files {
		if file.Path == "vba/WordUpStyleConverter.bas" {
			source = file.Text
		}
	}
	for _, want := range []string{
		"Public Function WU_ConvertStyleInRange",
		"Private Function WU_ConvertStyleInStory",
		`If document Is Nothing Then Err.Raise 91, "WU_ConvertStyle", "document is required"`,
		`If Len(Trim$(fromStyle)) = 0 Then Err.Raise 5, "WU_ConvertStyle", "source style is required"`,
		`If Len(Trim$(toStyle)) = 0 Then Err.Raise 5, "WU_ConvertStyle", "target style is required"`,
		`If sourceStyle.Type <> wdStyleTypeParagraph Then Err.Raise 5, "WU_ConvertStyle", "source style is not a paragraph style"`,
		`If targetStyle.Type <> wdStyleTypeParagraph Then Err.Raise 5, "WU_ConvertStyle", "target style is not a paragraph style"`,
		"captured = True",
		"If captured Then Application.ScreenUpdating = updating",
		"If targetEnd <= targetStart Then Exit Function",
		"If story.End > story.Start Then If WU_ConvertStyleInStory",
	} {
		if !strings.Contains(source, want) {
			t.Fatalf("style converter source omitted %q", want)
		}
	}
}

func TestTextOperationsUsesBoundedStoryFindAndStateCleanup(t *testing.T) {
	item, err := Get("document.text-operations")
	if err != nil {
		t.Fatal(err)
	}
	if item.Version != "1.0.8" {
		t.Fatalf("text operations version=%q", item.Version)
	}
	if len(item.Files) != 1 || item.Files[0].Path != "vba/WordUpTextOperations.bas" {
		t.Fatalf("unexpected text operations files: %#v", item.Files)
	}
	for _, capability := range []string{"literal replacement", "literal counting", "batch replacement", "range-bounded edits", "character-style matching"} {
		found := false
		for _, got := range item.Capabilities {
			if got == capability {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("text operations manifest omitted capability %q: %#v", capability, item.Capabilities)
		}
	}
	source := item.Files[0].Text
	for _, want := range []string{
		"Public Function WU_ReplaceLiteral",
		"Public Function WU_CountLiteral",
		"Public Function WU_CountLiteralInRange",
		"Public Function WU_ReplaceLiteralBatch",
		"Public Function WU_ReplaceLiteralBatchInRange",
		"Public Function WU_ReplaceLiteralInRange",
		"Public Function WU_ApplyCharacterStyleToMatches",
		"Public Function WU_ApplyCharacterStyleToRange",
		"target range is required",
		"never opens an undo record or toggles ScreenUpdating",
		"story scope must be main, notes, or all",
		"replacements must be a two-dimensional array",
		"replacements must have exactly two columns",
		"style is not a character style",
		"find text exceeds Word's 255-character limit",
		"Application.UndoRecord.StartCustomRecord \"Replace literal text\"",
		"Application.UndoRecord.StartCustomRecord \"Replace literal text batch\"",
		"Application.UndoRecord.StartCustomRecord \"Style literal matches\"",
		"If captured Then Application.ScreenUpdating = updating",
		".Replacement.Text = WU_EscapeFindLiteral(replaceText)",
		"find text exceeds Word's escaped 255-character limit",
		"replacement text exceeds Word's escaped 255-character limit",
		"story.End > story.Start",
		"Set story = document.StoryRanges(wdMainTextStory)",
		".MatchWholeWord = wholeWord",
		".MatchSoundsLike = False",
		".MatchAllWordForms = False",
		"WU_EscapeFindLiteral = Replace(value, \"^\", \"^^\")",
		"Set firstStory = document.StoryRanges(wdFootnotesStory)",
		"Set firstStory = document.StoryRanges(wdEndnotesStory)",
		"WU_ReplaceLiteralInStoryChain",
		"With MatchCase on, identical find/replacement text is an exact no-op.",
		"If matchCase And StrComp(findText, replaceText, vbBinaryCompare) = 0 Then Exit Function",
		"targetStart = target.Start: targetEnd = target.End",
		"If targetEnd <= targetStart Then Exit Function",
		"WU_ValidateLiteralBatch(replacements, matchCase)",
		"ReDim matched(firstRow To lastRow)",
		"If matched(row) Then changed = changed + 1",
		"ByRef matched() As Boolean",
		"WU_ApplyCharacterStyleInStoryChain",
		"If StrComp(currentStyle, style.NameLocal, vbTextCompare) <> 0 Then",
		"search.SetRange Start:=nextStart, End:=story.End",
	} {
		if !strings.Contains(source, want) {
			t.Fatalf("text operations source omitted %q", want)
		}
	}
}

func TestSafeEditRequiresExplicitCaptureState(t *testing.T) {
	item, err := Get("operation.safe-edit")
	if err != nil {
		t.Fatal(err)
	}
	if item.Version != "1.0.2" {
		t.Fatalf("safe-edit version did not advance: %q", item.Version)
	}
	var source string
	for _, file := range item.Files {
		if file.Path == "vba/WordUpSafeEdit.bas" {
			source = file.Text
		}
	}
	for _, want := range []string{
		"ByRef captured As Boolean",
		"captured = False",
		"If captured Then Application.ScreenUpdating = updating",
	} {
		if !strings.Contains(source, want) {
			t.Fatalf("safe-edit source omitted %q", want)
		}
	}
}

func TestFormShellPreservesUnloadFailure(t *testing.T) {
	item, err := Get("ui.form-shell")
	if err != nil {
		t.Fatal(err)
	}
	if item.Version != "1.0.2" {
		t.Fatalf("form shell version did not advance: %q", item.Version)
	}
	var source string
	for _, file := range item.Files {
		if file.Path == "vba/WordUpFormShell.bas" {
			source = file.Text
		}
	}
	for _, want := range []string{
		"If Not instance Is Nothing Then Unload instance",
		"If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description",
		"Err.Clear",
	} {
		if !strings.Contains(source, want) {
			t.Fatalf("form shell cleanup omitted %q", want)
		}
	}
}

func TestRibbonCallbackEncodingUsesASCIIOnly(t *testing.T) {
	item, err := Get("ui.ribbon-command")
	if err != nil {
		t.Fatal(err)
	}
	if item.Version != "1.0.3" {
		t.Fatalf("Ribbon component version did not advance: %q", item.Version)
	}
	var source string
	for _, file := range item.Files {
		if file.Path == "vba/WordUpRibbon.bas" {
			source = file.Text
		}
	}
	if !strings.Contains(source, "code = AscW(character)") || !strings.Contains(source, "code >= 48 And code <= 57") || strings.Contains(source, `character Like "[A-Za-z0-9]"`) {
		t.Fatalf("Ribbon callback encoder is not locale-independent: %s", source)
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

func TestLocalBundleRejectsNonCanonicalBackslashPath(t *testing.T) {
	bundle := t.TempDir()
	manifest := Manifest{Schema: 1, ID: "noncanonical.path", Version: "1", License: "MIT", Provenance: "test", Files: []File{{Path: `assets\icon.bin`}}}
	if err := project.Write(bundle, "component.json", project.JSON(manifest), ""); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadBundle(bundle); err == nil || !strings.Contains(err.Error(), "canonical relative path") {
		t.Fatalf("non-canonical component path was accepted: %v", err)
	}
}

func TestLocalBundleRejectsCaseCollisionBeforePayloadReads(t *testing.T) {
	bundle := t.TempDir()
	manifest := Manifest{Schema: 1, ID: "duplicate.paths", Version: "1", License: "MIT", Provenance: "test", Files: []File{{Path: "assets/icon.bin"}, {Path: "assets/ICON.BIN"}}}
	if err := project.Write(bundle, "component.json", project.JSON(manifest), ""); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(bundle, "assets/icon.bin", []byte("icon"), ""); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadBundle(bundle); err == nil || !strings.Contains(err.Error(), "duplicate path") {
		t.Fatalf("case-colliding component paths were accepted: %v", err)
	}
}

func TestLocalBundleRejectsDuplicateFormControlsBeforeCopy(t *testing.T) {
	bundle, root := t.TempDir(), t.TempDir()
	manifest := Manifest{Schema: 1, ID: "form.invalid", Version: "1", License: "MIT", Provenance: "test", Files: []File{{Path: "forms/Editor.json"}}}
	if err := project.Write(bundle, "component.json", project.JSON(manifest), ""); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(bundle, "forms/Editor.json", []byte(`{"name":"Editor","controls":[{"name":"Run","type":"CommandButton"},{"name":"run","type":"Label"}]}`), ""); err != nil {
		t.Fatal(err)
	}
	if _, err := AddBundle(root, bundle); err == nil || !strings.Contains(err.Error(), "duplicate control name \"run\" in controls") {
		t.Fatalf("duplicate form control was not rejected during preflight: %v", err)
	}
	if _, err := project.Read(root, "forms/Editor.json"); !os.IsNotExist(err) {
		t.Fatalf("invalid form bundle was partially installed: %v", err)
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

func TestRibbonMergePreflightNormalizesTargetCase(t *testing.T) {
	const ns = "http://schemas.microsoft.com/office/2009/07/customui"
	base := []byte(`<customUI xmlns="` + ns + `"><ribbon><tabs><tab id="base"/></tabs></ribbon></customUI>`)
	first := []byte(`<customUI xmlns="` + ns + `"><ribbon><tabs><tab id="same"/></tabs></ribbon></customUI>`)
	second := []byte(`<customUI xmlns="` + ns + `"><ribbon><tabs><tab id="same" label="collision"/></tabs></ribbon></customUI>`)
	bundle := t.TempDir()
	manifest := Manifest{Schema: 1, ID: "ribbon.case-conflict", Version: "1", License: "MIT", Provenance: "case-normalization test", Files: []File{{Path: "assets/first.xml"}, {Path: "assets/second.xml"}}, RibbonMerges: []RibbonMerge{{Source: "assets/first.xml", Target: "customUI/customUI14.xml"}, {Source: "assets/second.xml", Target: "customui/customui14.xml"}}}
	if err := project.Write(bundle, "component.json", project.JSON(manifest), ""); err != nil {
		t.Fatal(err)
	}
	for path, data := range map[string][]byte{"assets/first.xml": first, "assets/second.xml": second} {
		if err := project.Write(bundle, path, data, ""); err != nil {
			t.Fatal(err)
		}
	}
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("RibbonCaseConflict", root); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "package/customUI/customUI14.xml", base, ""); err != nil {
		t.Fatal(err)
	}
	if _, err := AddBundle(root, bundle); err == nil || !strings.Contains(err.Error(), "Ribbon merge assets/second.xml") {
		t.Fatalf("case-variant Ribbon collision was not rejected before copy: %v", err)
	}
	if _, err := project.Read(root, "assets/first.xml"); !os.IsNotExist(err) {
		t.Fatalf("first Ribbon source copied after rejected preflight: %v", err)
	}
}
