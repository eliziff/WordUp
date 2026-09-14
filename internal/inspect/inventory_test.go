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
	commands := "Attribute VB_Name = \"Commands\"\nOption Explicit\n' KeyBindings.Add wdKeyCategoryMacro, \"comment\"\nRem CommandBars(\"comment\").Controls.Add\nMsgBox \"KeyBindings.Add and CommandBars(\"\"Text\"\").Controls.Add\"\nPublic Sub RunThing()\n    KeyBindings.Add wdKeyCategoryMacro, \"Commands.RunThing\", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyU)\n    CommandBars(\"Text\").Controls.Add Type:=msoControlButton\nEnd Sub\n"
	form := "Attribute VB_Name = \"Editor\"\nOption Explicit\nPrivate Sub UserForm_AddControl(ByVal Control As MSForms.Control)\nEnd Sub\nPrivate Sub cmdSave_Click()\nEnd Sub\nPrivate Sub cmdSave_DropButtonClick()\nEnd Sub\n"
	for name, data := range map[string][]byte{"vba/Commands.bas": []byte(commands), "vba/Editor.vba": []byte(form), "forms/Editor.json": []byte(`{"name":"Editor","controls":[{"name":"cmdSave","type":"CommandButton"}]}`), "package/customUI14.xml": []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="tab" label="Tab"><group id="group" label="Group"><button id="run" onAction="RunThing"/></group></tab></tabs></ribbon><contextMenus><contextMenu idMso="ContextMenuText"/></contextMenus></customUI>`)} {
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
