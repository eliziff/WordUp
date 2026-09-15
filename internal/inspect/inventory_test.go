package inspect

import (
	"path/filepath"
	"testing"

	"github.com/eliziff/WordUp/internal/project"
)

func TestRibbonDeclarationShapeDefersContinuedSignature(t *testing.T) {
	actual := Symbol{Kind: "Sub", Declaration: "Public Sub Action(_"}
	if mismatch := ribbonDeclarationMismatch("Public Sub Action(control As IRibbonControl)", actual); mismatch != "" {
		t.Fatalf("continued declaration was treated as a mismatch: %s", mismatch)
	}
}

func TestCheckInventoryReportsWiringAndModifiedComponent(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("Inventory", root); err != nil {
		t.Fatal(err)
	}
	commands := "Attribute VB_Name = \"Commands\"\nOption Explicit\n' KeyBindings.Add wdKeyCategoryMacro, \"comment\"\nRem CommandBars(\"comment\").Controls.Add\nMsgBox \"KeyBindings.Add and CommandBars(\"\"Text\"\").Controls.Add\"\nPublic Sub RunThing()\n    Dim bar As CommandBar\n    Set bar = CommandBars(\"Text\")\n    KeyBindings. _\n        Add wdKeyCategoryMacro, \"Commands.RunThing\", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyU)\n    CommandBars(\"Text\").Controls. _\n        Add Type:=msoControlButton\nEnd Sub\n"
	form := "Attribute VB_Name = \"Editor\"\nOption Explicit\nPrivate Sub UserForm_AddControl(ByVal Control As MSForms.Control)\nEnd Sub\nPrivate Sub cmdSave_Click()\nEnd Sub\nPrivate Sub cmdSave_DropButtonClick()\nEnd Sub\n"
	for name, data := range map[string][]byte{"vba/Commands.bas": []byte(commands), "vba/Editor.vba": []byte(form), "forms/editor.json": []byte(`{"name":"editor","controls":[{"name":"cmdSave","type":"CommandButton"}]}`), "package/customUI14.xml": []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="tab" label="Tab"><group id="group" label="Group"><button id="run" onAction="RunThing"/></group></tab></tabs></ribbon><contextMenus><contextMenu idMso="ContextMenuText"/></contextMenus></customUI>`)} {
		if err := project.Write(root, name, data, ""); err != nil {
			t.Fatal(err)
		}
	}
	lock := []byte(`{"schema":1,"components":{"command.hotkey":{"id":"command.hotkey","version":"1.0.2","files":{"vba/Commands.bas":"wrong"},"provenance":"test","license":"MIT"}}}`)
	if err := project.Write(root, ".wordwright/components.json", lock, ""); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	result, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	inventory := result["inventory"].(map[string]any)
	public := inventory["public_macros"].([]map[string]any)
	if len(public) != 1 || public[0]["name"] != "RunThing" {
		t.Fatalf("unexpected public macro inventory: %#v", public)
	}
	events := inventory["form_events"].([]map[string]any)
	if len(events) != 3 || events[0]["wiring"] != "form_lifecycle" || events[1]["wiring"] != "control_present" || events[2]["wiring"] != "control_present" {
		t.Fatalf("unexpected form event inventory: %#v", events)
	}
	ribbons := inventory["ribbon_callbacks"].([]map[string]any)
	if len(ribbons) != 1 || ribbons[0]["state"] != "declared" || ribbons[0]["name"] != "RunThing" {
		t.Fatalf("unexpected Ribbon inventory: %#v", ribbons)
	}
	if hotkeys := inventory["hotkey_registrations"].([]map[string]any); len(hotkeys) != 1 {
		t.Fatalf("hotkey registration missing: %#v", hotkeys)
	}
	menus := inventory["context_menu_registrations"].([]map[string]any)
	if len(menus) != 2 {
		t.Fatalf("context-menu registrations missing: %#v", menus)
	}
	components := inventory["installed_components"].([]map[string]any)
	if len(components) != 1 || components[0]["state"] != "modified" {
		t.Fatalf("component state missing: %#v", components)
	}
	foundComponentDiagnostic := false
	for _, diagnostic := range result["diagnostics"].([]map[string]any) {
		if diagnostic["component"] == "command.hotkey" {
			foundComponentDiagnostic = true
		}
	}
	if !foundComponentDiagnostic {
		t.Fatal("modified component did not produce a wiring diagnostic")
	}
}

func TestCheckInventoryDoesNotTreatStandardModuleAsFormEvents(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("StandardModule", root); err != nil {
		t.Fatal(err)
	}
	for name, data := range map[string][]byte{
		"vba/Commands.bas":    []byte("Attribute VB_Name = \"Commands\"\nOption Explicit\nPublic Sub cmdSave_Click()\nEnd Sub\n"),
		"forms/Commands.json": []byte(`{"name":"Commands","controls":[{"name":"cmdSave","type":"CommandButton"}]}`),
	} {
		if err := project.Write(root, name, data, ""); err != nil {
			t.Fatal(err)
		}
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	result, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	inventory := result["inventory"].(map[string]any)
	if events := inventory["form_events"].([]map[string]any); len(events) != 0 {
		t.Fatalf("standard module was inventoried as a form event: %#v", events)
	}
}

func TestCheckInventoryReportsDuplicateModuleNames(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("DuplicateModules", root); err != nil {
		t.Fatal(err)
	}
	for name := range map[string]bool{"vba/First.bas": true, "vba/Second.bas": true} {
		if err := project.Write(root, name, []byte("Attribute VB_Name = \"SameModule\"\nOption Explicit\n"), ""); err != nil {
			t.Fatal(err)
		}
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	result, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	inventory := result["inventory"].(map[string]any)
	if modules := inventory["module_names"].([]map[string]any); len(modules) != 3 {
		t.Fatalf("module inventory=%#v", modules)
	}
	for _, diagnostic := range result["diagnostics"].([]map[string]any) {
		if diagnostic["message"] == "duplicate VBA module name: samemodule" {
			return
		}
	}
	t.Fatal("duplicate module diagnostic missing")
}

func TestCheckInventoryReportsDerivedModuleName(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("DerivedModule", root); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "vba/nested/NoAttribute.bas", []byte("Option Explicit\nPublic Sub Run()\nEnd Sub\n"), ""); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	result, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	inventory := result["inventory"].(map[string]any)
	modules := inventory["module_names"].([]map[string]any)
	for _, module := range modules {
		if module["file"] == "vba/nested/NoAttribute.bas" && module["module"] == "NoAttribute" && module["derived"] == true {
			return
		}
	}
	t.Fatalf("filename-derived module was not inventoried: %#v", modules)
}

func TestCheckInventoryReportsNestedDerivedModuleCollision(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("NestedDuplicateModules", root); err != nil {
		t.Fatal(err)
	}
	for path := range map[string]bool{"vba/first/Foo.bas": true, "vba/second/Foo.cls": true} {
		if err := project.Write(root, path, []byte("Option Explicit\n"), ""); err != nil {
			t.Fatal(err)
		}
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	result, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	for _, diagnostic := range result["diagnostics"].([]map[string]any) {
		if diagnostic["message"] == "duplicate VBA module name: foo" {
			locations := diagnostic["locations"].([]map[string]any)
			if len(locations) != 2 {
				t.Fatalf("duplicate module locations=%#v", locations)
			}
			return
		}
	}
	t.Fatalf("nested duplicate module diagnostic missing: %#v", result["diagnostics"])
}

func TestCheckInventoryReportsModuleNameMismatch(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("ModuleMismatch", root); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "vba/Expected.bas", []byte("Attribute VB_Name = \"Other\"\nOption Explicit\n"), ""); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	result, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	for _, diagnostic := range result["diagnostics"].([]map[string]any) {
		if diagnostic["file"] == "vba/Expected.bas" && diagnostic["message"] == `VB_Name "Other" does not match module filename "Expected"` {
			return
		}
	}
	t.Fatalf("module name mismatch diagnostic missing: %#v", result["diagnostics"])
}

func TestCheckInventoryReportsInvalidVBNameAttribute(t *testing.T) {
	for name, source := range map[string]string{
		"invalid identifier": "Attribute VB_Name = \"Bad-Name\"\nOption Explicit\n",
		"malformed value":    "Attribute VB_Name = BadName\nOption Explicit\n",
	} {
		t.Run(name, func(t *testing.T) {
			root := filepath.Join(t.TempDir(), "workspace")
			if _, err := project.New("InvalidVBName", root); err != nil {
				t.Fatal(err)
			}
			if err := project.Write(root, "vba/Good.bas", []byte(source), ""); err != nil {
				t.Fatal(err)
			}
			w, err := project.Open(root)
			if err != nil {
				t.Fatal(err)
			}
			result, err := Check(w)
			if err != nil {
				t.Fatal(err)
			}
			for _, diagnostic := range result["diagnostics"].([]map[string]any) {
				if diagnostic["file"] == "vba/Good.bas" && diagnostic["message"] == "invalid VB_Name attribute" {
					return
				}
			}
			t.Fatalf("invalid VB_Name diagnostic missing: %#v", result["diagnostics"])
		})
	}
}

func TestCheckInventoryReportsInvalidModuleFilename(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("InvalidModule", root); err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "vba/Bad-Name.bas", []byte("Option Explicit\nPublic Sub Run()\nEnd Sub\n"), ""); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	result, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	for _, diagnostic := range result["diagnostics"].([]map[string]any) {
		if diagnostic["file"] == "vba/Bad-Name.bas" && diagnostic["message"] == `VBA source filename does not derive a valid module name "Bad-Name"` {
			return
		}
	}
	t.Fatalf("invalid module filename diagnostic missing: %#v", result["diagnostics"])
}

func TestCheckReportsInvalidAndMisnamedFormDesigns(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("FormDesignChecks", root); err != nil {
		t.Fatal(err)
	}
	for path, data := range map[string][]byte{
		"forms/Broken.json":   []byte(`{"name":`),
		"forms/Wrong.json":    []byte(`{"name":"Other"}`),
		"forms/Standard.json": []byte(`{"name":"Standard"}`),
		"vba/Wrong.vba":       []byte("Option Explicit\n"),
		"vba/Missing.vba":     []byte("Option Explicit\n"),
		"vba/Standard.bas":    []byte("Option Explicit\n"),
	} {
		if err := project.Write(root, path, data, ""); err != nil {
			t.Fatal(err)
		}
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	result, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	foundBroken, foundMismatch, foundMissing, foundWrongSource := false, false, false, false
	for _, diagnostic := range result["diagnostics"].([]map[string]any) {
		if diagnostic["file"] == "forms/Broken.json" && diagnostic["message"] == "invalid form design: unexpected EOF" {
			foundBroken = true
		}
		if diagnostic["file"] == "forms/Wrong.json" && diagnostic["message"] == `form design name "Other" does not match filename "Wrong"` {
			foundMismatch = true
		}
		if diagnostic["file"] == "vba/Missing.vba" && diagnostic["message"] == "form source has no matching persistent design; new forms require forms/Missing.json" {
			foundMissing = true
		}
		if diagnostic["file"] == "forms/Standard.json" && diagnostic["message"] == "form design has no matching VBA source module" {
			foundWrongSource = true
		}
	}
	if !foundBroken || !foundMismatch || !foundMissing || !foundWrongSource {
		t.Fatalf("form design diagnostics missing: %#v", result["diagnostics"])
	}
}
