//go:build windows && (amd64 || arm64)

package native

import (
	"context"
	"fmt"
	"github.com/eliziff/WordUp/internal/office"
	"os"
	"path/filepath"
	"runtime"
)

// Preview hands an explicitly launched visible Word instance to the user.
// It never attaches to an existing instance or changes persistent trust settings.
func Preview(ctx context.Context, file, document string) (any, error) {
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
	h, err := connectWord(hostConfig{Options: Options{Directory: directory, Execute: true, StartupTimeoutMS: 30000}, Desktop: "Default", WordPath: word})
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
	documents, err := h.app.get("Documents")
	if err != nil {
		return nil, err
	}
	docs, err := documents.object()
	documents.clear()
	if err != nil {
		return nil, err
	}
	closed, err := docs.call("Close", 0)
	closed.clear()
	docs.release()
	if err != nil {
		return nil, err
	}
	var opened any
	if document == "" {
		opened, err = h.operation(Operation{Op: "new", File: file, As: "preview"})
	} else {
		var loaded any
		loaded, err = h.operation(Operation{Op: "addin", File: file, As: "template"})
		if err != nil {
			return nil, err
		}
		templatePath := loaded.(map[string]any)["staged_path"].(string)
		opened, err = h.operation(Operation{Op: "open", File: document, As: "preview"})
		if err != nil {
			return nil, err
		}
		d := h.objects["preview"]
		if err = d.put("UpdateStylesOnOpen", true); err != nil {
			return nil, err
		}
		if err = d.put("AttachedTemplate", templatePath); err != nil {
			return nil, err
		}
		updated, updateErr := d.call("UpdateStyles")
		updated.clear()
		if updateErr != nil {
			return nil, updateErr
		}
		saved, saveErr := d.call("Save")
		saved.clear()
		if saveErr != nil {
			return nil, saveErr
		}
	}
	if err != nil {
		return nil, err
	}
	activated, err := h.app.call("Activate")
	activated.clear()
	if err != nil {
		return nil, err
	}
	handedOff = true
	return map[string]any{"opened": true, "pid": h.process.PID, "template": file, "document": opened, "persistent_trust_changed": false, "automation_security_restored": true, "priority": "below_normal", "user_owned_after_launch": true}, nil
}
