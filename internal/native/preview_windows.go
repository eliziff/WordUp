//go:build windows && (amd64 || arm64)

package native

import (
	"context"
	"fmt"
	"github.com/eliziff/WordUp/internal/office"
	"os"
	"path/filepath"
	"runtime"
	"strings"
)

// Preview hands an explicitly launched visible Word instance to the user.
// It never attaches to an existing instance or changes persistent trust settings.
func Preview(ctx context.Context, file, document string) (result any, err error) {
	phase := "start Word"
	defer func() {
		if err != nil {
			err = Fail("preview_failed", "Preview failed while attempting to "+phase, map[string]any{"phase": phase, "cause": fault(err)})
		}
	}()
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	word, err := findWord()
	if err != nil {
		return nil, err
	}
	directory := filepath.Dir(file)
	seed, err := office.BlankPackage().Bytes()
	if err != nil {
		return nil, err
	}
	if err = os.WriteFile(filepath.Join(directory, "seed.docx"), seed, 0600); err != nil {
		return nil, err
	}
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	hr, _, _ := coInit.Call(0, 2)
	if failed(hr) {
		return nil, fmt.Errorf("preview COM initialization failed")
	}
	defer coUninit.Call()
	h, err := connectWord(hostConfig{Options: Options{Directory: directory, Execute: true, Visible: true, StartupTimeoutMS: 30000}, Desktop: "Default", WordPath: word})
	if err != nil {
		return nil, err
	}
	handedOff := false
	defer func() {
		if !handedOff {
			h.close()
			return
		}
		for _, object := range h.objects {
			object.release()
		}
		closeHandle.Call(uintptr(h.process.Process))
	}()
	phase = "open working document"
	// Both new documents and manuscripts use the same attachment and style contract.
	var opened any
	if document == "" {
		opened, err = h.operation(Operation{Op: "new", File: file, As: "preview"})
	} else {
		opened, err = h.operation(Operation{Op: "open", File: document, As: "preview"})
	}
	if err != nil {
		return nil, err
	}
	templatePath, err := h.stage(file)
	if err != nil {
		return nil, err
	}
	phase = "attach template, apply styles and save"
	state, err := preparePreviewDocument(h.objects["preview"], templatePath, filepath.Join(directory, "preview-document.docx"))
	if err != nil {
		return nil, err
	}
	// Reopen the saved copy so Word loads the attached template's Ribbon too.
	phase = "close saved working document"
	if _, err = h.operation(Operation{Op: "unload", Target: "preview"}); err != nil {
		return nil, err
	}
	phase = "reopen saved working document"
	opened, err = h.operation(Operation{Op: "open", File: state["document"].(string), As: "preview"})
	if err != nil {
		return nil, err
	}
	phase = "verify reopened attachment and automatic styles"
	state, err = verifyPreviewAttachment(h.objects["preview"], templatePath, state["document"].(string))
	if err != nil {
		return nil, err
	}
	// Foreground focus is best-effort; do not discard a verified visible copy
	// merely because Word is busy activating its window. Retain the diagnostic.
	activated, activationErr := h.app.call("Activate")
	activated.clear()
	handedOff = true
	result = map[string]any{"opened": true, "pid": h.process.PID, "template": file, "document": opened, "attachment": state, "persistent_trust_changed": false, "automation_security_restored": true, "priority": "below_normal", "user_owned_after_launch": true}
	if activationErr != nil {
		result.(map[string]any)["activation_error"] = fault(activationErr)
	}
	return result, nil
}

// preparePreviewDocument makes attachment a checked operation, not a UI convention.
func preparePreviewDocument(d dispatch, templatePath, output string) (map[string]any, error) {
	if err := d.put("AttachedTemplate", templatePath); err != nil {
		return nil, err
	}
	if err := d.put("UpdateStylesOnOpen", true); err != nil {
		return nil, err
	}
	v, err := d.call("UpdateStyles")
	v.clear()
	if err != nil {
		return nil, err
	}
	v, err = d.get("HasVBProject")
	if err != nil {
		return nil, err
	}
	hasVBA, err := v.value(0)
	v.clear()
	if err != nil {
		return nil, err
	}
	format := 12 // wdFormatXMLDocument
	if hasVBA == true {
		format = 13 // Preserve a manuscript's own VBA project.
		output = strings.TrimSuffix(output, filepath.Ext(output)) + ".docm"
	}
	v, err = d.call("SaveAs2", output, format)
	v.clear()
	if err != nil {
		return nil, err
	}
	return verifyPreviewAttachment(d, templatePath, output)
}

func verifyPreviewAttachment(d dispatch, templatePath, output string) (map[string]any, error) {
	template, err := objectProperty(d, "AttachedTemplate")
	if err != nil {
		return nil, err
	}
	defer template.release()
	v, err := template.get("FullName")
	if err != nil {
		return nil, err
	}
	name, err := v.value(0)
	v.clear()
	if err != nil {
		return nil, err
	}
	v, err = d.get("UpdateStylesOnOpen")
	if err != nil {
		return nil, err
	}
	update, err := v.value(0)
	v.clear()
	if err != nil {
		return nil, err
	}
	actual, ok := name.(string)
	if !ok || !strings.EqualFold(filepath.Clean(actual), filepath.Clean(templatePath)) || update != true {
		return nil, Fail("preview_attachment_mismatch", "Word did not retain the requested template and automatic style updates", map[string]any{"expected_template": templatePath, "attached_template": name, "update_styles_on_open": update})
	}
	return map[string]any{"document": output, "attached_template": actual, "update_styles_on_open": update, "styles_applied": true, "saved": true}, nil
}
