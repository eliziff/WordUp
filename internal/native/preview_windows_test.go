//go:build windows && (amd64 || arm64)

package native

import (
	"os"
	"path/filepath"
	"runtime"
	"testing"

	"github.com/eliziff/WordUp/internal/office"
)

func TestNativePreviewAttachment(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	hr, _, _ := coInit.Call(0, 2)
	if failed(hr) {
		t.Fatal("COM initialization failed")
	}
	defer coUninit.Call()
	dir := t.TempDir()
	seed, err := office.BlankPackage().Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if err = os.WriteFile(filepath.Join(dir, "seed.docx"), seed, 0600); err != nil {
		t.Fatal(err)
	}
	word, err := findWord()
	if err != nil {
		t.Fatal(err)
	}
	h, err := connectWord(hostConfig{Options: Options{Directory: dir, Execute: true, StartupTimeoutMS: 30000}, WordPath: word})
	if err != nil {
		t.Fatal(err)
	}
	defer h.close()
	call := func(op Operation) any {
		t.Helper()
		v, err := h.operation(op)
		if err != nil {
			t.Fatal(err)
		}
		return v
	}
	font := func(target string) dispatch {
		t.Helper()
		styles, err := objectProperty(h.objects[target], "Styles")
		if err != nil {
			t.Fatal(err)
		}
		defer styles.release()
		normal, err := objectProperty(styles, "Item", -1)
		if err != nil {
			t.Fatal(err)
		}
		defer normal.release()
		f, err := objectProperty(normal, "Font")
		if err != nil {
			t.Fatal(err)
		}
		return f
	}
	template := filepath.Join(dir, "house.dotx")
	call(Operation{Op: "new", As: "template"})
	f := font("template")
	err = f.put("Size", 17)
	f.release()
	if err != nil {
		t.Fatal(err)
	}
	call(Operation{Op: "invoke", Target: "template", Member: "SaveAs2", Args: []any{template, 14}})
	call(Operation{Op: "unload", Target: "template"})
	call(Operation{Op: "new", As: "preview"})
	f = font("preview")
	err = f.put("Size", 11)
	f.release()
	if err != nil {
		t.Fatal(err)
	}
	call(Operation{Op: "put", Target: "preview", Member: "UpdateStylesOnOpen", Value: false})
	output := filepath.Join(dir, "preview-document.docx")
	if _, err = preparePreviewDocument(h.objects["preview"], template, output); err != nil {
		t.Fatal(err)
	}
	call(Operation{Op: "unload", Target: "preview"})
	call(Operation{Op: "open", File: output, As: "preview"})
	if _, err = verifyPreviewAttachment(h.objects["preview"], template, output); err != nil {
		t.Fatal(err)
	}
	f = font("preview")
	size, err := scalarNumber(f, "Size")
	f.release()
	if err != nil || size != 17 {
		t.Fatalf("template style was not applied: size=%v err=%v", size, err)
	}
	call(Operation{Op: "put", Target: "preview", Member: "UpdateStylesOnOpen", Value: false})
	if _, err = verifyPreviewAttachment(h.objects["preview"], template, output); err == nil {
		t.Fatal("disabled automatic updates reported as success")
	}
	if _, err = preparePreviewDocument(h.objects["preview"], filepath.Join(dir, "missing.dotx"), output); err == nil {
		t.Fatal("missing template reported as success")
	}
}
