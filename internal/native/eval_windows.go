//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"wordwright.local/internal/office"
)

func (h *wordHost) evaluate(op Operation) (any, error) {
	if !h.execute {
		return nil, Fail("execution_not_authorized", "Scratch VBA execution requires the execute capability", nil)
	}
	body, ok := op.Value.(string)
	if !ok || len(body) > 1<<20 {
		return nil, fmt.Errorf("eval value must be a bounded VBA function body; assign the result to Evaluate")
	}
	name := "WWEval" + randomID()[:16]
	v := office.NewVBA(name)
	source := "Attribute VB_Name = \"Evaluation\"\nOption Explicit\nPublic Function Evaluate() As Variant\n" + body + "\nEnd Function\n"
	vb, e := v.Rewrite([]office.Module{{Name: "ThisDocument", Kind: "document", Source: officeDocumentSource}, {Name: "Evaluation", Kind: "standard", Source: source}}, nil)
	if e != nil {
		return nil, e
	}
	p := office.BlankPackage()
	p.Files["word/vbaProject.bin"] = vb
	if e = p.ContentType("word/document.xml", office.MainDOTM); e != nil {
		return nil, e
	}
	if e = p.ContentType("word/vbaProject.bin", "application/vnd.ms-office.vbaProject"); e != nil {
		return nil, e
	}
	if e = p.Relationship("word/document.xml", "rIdVBA", office.R+"/vbaProject", "vbaProject.bin", ""); e != nil {
		return nil, e
	}
	data, e := p.Bytes()
	if e != nil {
		return nil, e
	}
	path := filepath.Join(h.cfg.Directory, name+".dotm")
	if e = os.WriteFile(path, data, 0600); e != nil {
		return nil, e
	}
	// Explicitly generated scratch add-in, not an edit to the tested artifact.
	// Its complete VBA source and package hash are returned for inspection.
	_, e = h.operation(Operation{Op: "addin", File: path, As: name})
	if e != nil {
		return nil, e
	}
	defer func() {
		if addin, ok := h.objects[name]; ok {
			_ = addin.put("Installed", false)
			addin.release()
			delete(h.objects, name)
		}
	}()
	value, e := h.app.invoke("Run", 1, []any{name + ".Evaluation.Evaluate"}, nil, h.objects)
	if e != nil {
		return nil, Fail("scratch_vba_failed", e.Error(), map[string]any{"generated_source": source, "scratch_sha256": office.Hash(data), "cause": fault(e)})
	}
	result, e := h.result(&value, op.As)
	if e != nil {
		return nil, e
	}
	return map[string]any{"result": result, "executor": "native Microsoft Word VBA", "generated_source": source, "scratch_sha256": office.Hash(data), "tested_artifact_modified": false}, nil
}

const officeDocumentSource = `Attribute VB_Name = "ThisDocument"
Attribute VB_Base = "0{00020906-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
`

func (h *wordHost) compileMenu(op Operation) (any, error) {
	// A target project name is mandatory when selecting via the native UI rather
	// than VBProject. Never compile whatever project happens to be active.
	name := op.Member
	if name == "" {
		return nil, Fail("compiler_target_required", "Supply the project's name in member for native Compile-menu verification; VBProject access was not enabled", nil)
	}
	if e := h.app.put("ShowVisualBasicEditor", true); e != nil {
		return nil, e
	}
	vbe, e := objectProperty(h.app, "VBE")
	if e != nil {
		return nil, Fail("compiler_ui_unavailable", "Word did not expose its VBE menu surface; AccessVBOM remains unchanged", fault(e))
	}
	defer vbe.release()
	bars, e := objectProperty(vbe, "CommandBars")
	if e != nil {
		return nil, e
	}
	defer bars.release()
	v, e := bars.invoke("FindControl", 1, nil, map[string]any{"ID": 578}, nil)
	if e != nil {
		return nil, e
	}
	control, e := v.object()
	v.clear()
	if e != nil {
		return nil, e
	}
	defer control.release()
	caption, e := control.get("Caption")
	if e != nil {
		return nil, e
	}
	text, _ := caption.value(0)
	caption.clear()
	captionText, _ := text.(string)
	clean := strings.ReplaceAll(captionText, "&", "")
	if !strings.HasSuffix(strings.ToLower(strings.TrimSpace(clean)), strings.ToLower(name)) {
		return nil, Fail("compiler_target_mismatch", "Select the intended project through the owned VBE's accessibility tree, then retry. No other project was compiled", map[string]any{"expected_project": name, "actual_compile_caption": captionText, "windows": windowInventory(h.process.PID)})
	}
	enabled, e := control.get("Enabled")
	if e != nil {
		return nil, e
	}
	isEnabled, _ := enabled.value(0)
	enabled.clear()
	if isEnabled != true {
		return map[string]any{"vba_compiled": false, "reason": "Compile command was already disabled; no fresh compiler run was observed", "caption": captionText}, nil
	}
	v, e = control.call("Execute")
	v.clear()
	if e != nil {
		return nil, e
	}
	enabled, e = control.get("Enabled")
	if e != nil {
		return nil, e
	}
	after, _ := enabled.value(0)
	enabled.clear()
	if after != false {
		return nil, Fail("compile_not_confirmed", "Native Compile did not reach a clean disabled state; inspect the owned VBE diagnostics", nil)
	}
	return map[string]any{"vba_compiled": true, "project": name, "verification": "target-matched native VBE Compile command executed and became disabled", "vbproject_access_enabled": false}, nil
}
