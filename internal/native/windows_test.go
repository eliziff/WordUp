//go:build windows && (amd64 || arm64)

package native

import (
	"encoding/json"
	"image/png"
	"os"
	"path/filepath"
	"runtime"
	"syscall"
	"testing"
	"unsafe"
)

func TestWordSafeModeStartupWindowRequiresExactOwnedDialog(t *testing.T) {
	dialog := func(message string) []any {
		return []any{map[string]any{
			"hwnd": uint64(42), "class": "#32770", "title": "Microsoft Word", "visible_on_private_desktop": true,
			"children": []any{
				map[string]any{"class": "Static", "title": message},
				map[string]any{"class": "Button", "title": "&No"},
				map[string]any{"class": "Button", "title": "&Yes"},
			},
		}}
	}
	if hwnd, ok := wordSafeModeStartupWindow(dialog(wordSafeModeStartupPrompt)); !ok || hwnd != 42 {
		t.Fatalf("exact prompt was not recognized: hwnd=%d ok=%v", hwnd, ok)
	}
	hidden := dialog(wordSafeModeStartupPrompt)
	hidden[0].(map[string]any)["visible_on_private_desktop"] = false
	if _, ok := wordSafeModeStartupWindow(hidden); !ok {
		t.Fatal("owned hidden startup prompt was not recognized")
	}
	if _, ok := wordSafeModeStartupWindow(dialog(wordSafeModeStartupPrompt + " ")); ok {
		t.Fatal("changed startup prompt was recognized")
	}
	other := dialog(wordSafeModeStartupPrompt)
	other[0].(map[string]any)["class"] = "OpusApp"
	if _, ok := wordSafeModeStartupWindow(other); ok {
		t.Fatal("non-dialog window was recognized")
	}
}

func TestCloseRetainsCleanupFailure(t *testing.T) {
	directory := t.TempDir()
	path := filepath.Join(directory, "locked.dotm")
	if err := os.WriteFile(path, []byte("lock test"), 0600); err != nil {
		t.Fatal(err)
	}
	name, err := syscall.UTF16PtrFromString(path)
	if err != nil {
		t.Fatal(err)
	}
	handle, err := syscall.CreateFile(name, syscall.GENERIC_READ, syscall.FILE_SHARE_READ, nil, syscall.OPEN_EXISTING, 0, 0)
	if err != nil {
		t.Fatal(err)
	}
	defer syscall.CloseHandle(handle)
	h := &localHost{closed: make(chan struct{})}
	h.cfg.Directory = directory
	first := h.Close()
	if first == nil {
		t.Fatal("locked workspace cleanup was reported successful")
	}
	if f, ok := first.(*Fault); !ok || f.Code != "workspace_cleanup_failed" || f.Details.(map[string]any)["directory"] != directory {
		t.Fatal("cleanup failure omitted classification or retained path", first)
	}
	if second := h.Close(); second == nil || second.Error() != first.Error() {
		t.Fatalf("repeated Close lost its failure: first=%v second=%v", first, second)
	}
	if _, err := os.Stat(path); err != nil {
		t.Fatal("locked evidence was not retained", err)
	}
}

func TestScreenshotPixelChannels(t *testing.T) {
	s, e := newSurface(2, 1)
	if e != nil {
		t.Fatal(e)
	}
	defer s.close()
	copy(unsafe.Slice((*byte)(s.bits), 8), []byte{10, 20, 30, 0, 40, 50, 60, 0})
	file := filepath.Join(t.TempDir(), "pixels.png")
	if e = s.write(file); e != nil {
		t.Fatal(e)
	}
	f, e := os.Open(file)
	if e != nil {
		t.Fatal(e)
	}
	defer f.Close()
	im, e := png.Decode(f)
	if e != nil {
		t.Fatal(e)
	}
	r, g, b, a := im.At(0, 0).RGBA()
	if r != 30*257 || g != 20*257 || b != 10*257 || a != 65535 {
		t.Fatalf("incorrect BGRA conversion: %d %d %d %d", r, g, b, a)
	}
}

func TestOwnedWindowEnumeration(t *testing.T) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	class, _ := utf("STATIC")
	create := user32.NewProc("CreateWindowExW")
	destroy := user32.NewProc("DestroyWindow")
	parent, _, err := create.Call(0, uintptr(unsafe.Pointer(class)), 0, 0, 0, 0, 100, 100, 0, 0, 0, 0)
	if parent == 0 {
		t.Fatal(err)
	}
	defer destroy.Call(parent)
	child, _, err := create.Call(0, uintptr(unsafe.Pointer(class)), 0, 0x40000000, 0, 0, 20, 20, parent, 0, 0, 0)
	if child == 0 {
		t.Fatal(err)
	}
	q := windowEnum{pid: uint32(os.Getpid())}
	enumChildren.Call(parent, collectWindowsCallback, uintptr(unsafe.Pointer(&q)))
	if len(q.windows) != 1 || q.windows[0] != child {
		t.Fatalf("child enumeration: %v, want %v", q.windows, child)
	}
	if classOf(child) != "Static" {
		t.Fatalf("child class: %q", classOf(child))
	}
	// Exercise the same callback used to find Word's innermost document window.
	type windowClass struct {
		Style                              uint32
		Proc                               uintptr
		ClassExtra, WindowExtra            int32
		Instance, Icon, Cursor, Background uintptr
		Menu, Name                         *uint16
	}
	wordClass, _ := utf("_WwG")
	wc := windowClass{Proc: user32.NewProc("DefWindowProcW").Addr(), Name: wordClass}
	atom, _, err := user32.NewProc("RegisterClassW").Call(uintptr(unsafe.Pointer(&wc)))
	if atom == 0 {
		t.Fatal(err)
	}
	defer user32.NewProc("UnregisterClassW").Call(uintptr(unsafe.Pointer(wordClass)), 0)
	wordChild, _, err := create.Call(0, uintptr(unsafe.Pointer(wordClass)), 0, 0x40000000, 0, 0, 20, 20, parent, 0, 0, 0)
	if wordChild == 0 {
		t.Fatal(err)
	}
	defer destroy.Call(wordChild)
	if got := documentWindow(uint32(os.Getpid())); got != wordChild {
		t.Fatalf("document window: %v, want %v", got, wordChild)
	}
}

func TestNativeDialogButtonSelectors(t *testing.T) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	pid := uint32(os.Getpid())
	create := func(class, text string, style, parent uintptr) uintptr {
		t.Helper()
		c, _ := utf(class)
		label, _ := utf(text)
		hwnd, _, err := user32.NewProc("CreateWindowExW").Call(0, uintptr(unsafe.Pointer(c)), uintptr(unsafe.Pointer(label)), style, 0, 0, 200, 100, parent, 0, 0, 0)
		if hwnd == 0 {
			t.Fatal(err)
		}
		return hwnd
	}
	dialog := create("#32770", "Owned test dialog", 0, 0)
	defer user32.NewProc("DestroyWindow").Call(dialog)
	create("Static", "A decision is required.", 0x40000000, dialog)
	button := create("Button", "&No", 0x40000000, dialog)
	invoked := false
	callback := syscall.NewCallback(func(hwnd uintptr, message uint32, wparam, lparam uintptr) uintptr {
		if message == 0x0111 && lparam == button {
			invoked = true
			return 0
		}
		result, _, _ := user32.NewProc("DefWindowProcW").Call(hwnd, uintptr(message), wparam, lparam)
		return result
	})
	original, _, _ := user32.NewProc("SetWindowLongPtrW").Call(dialog, ^uintptr(3), callback)
	defer user32.NewProc("SetWindowLongPtrW").Call(dialog, ^uintptr(3), original)
	op := Operation{Op: "ui.find", HWND: uint64(dialog), Named: map[string]any{"name": "No", "role": 43, "scope": "dialog", "message_pattern": "^A decision is required\\.$"}}
	found, err := uiOperation(pid, t.TempDir(), false, op)
	if err != nil {
		t.Fatal(err)
	}
	// Selectors must serialize without cycles and be reusable without a name search.
	encoded, err := json.Marshal(found)
	if err != nil {
		t.Fatal(err)
	}
	var selected struct{ Selector Operation }
	if err = json.Unmarshal(encoded, &selected); err != nil {
		t.Fatal(err)
	}
	selected.Selector.Op = "ui.invoke"
	expectFault := func(operation Operation, execute bool, code string) {
		t.Helper()
		_, err := uiOperation(pid, t.TempDir(), execute, operation)
		if f := fault(err); f == nil || f.Code != code {
			t.Fatalf("want %s, got %v", code, err)
		}
	}
	expectFault(selected.Selector, false, "execution_not_authorized")
	selected.Selector.Named["expected_name"] = "Changed"
	expectFault(selected.Selector, true, "stale_ui_selector")
	selected.Selector.Named["expected_name"] = "No"
	user32.NewProc("EnableWindow").Call(button, 0)
	expectFault(selected.Selector, true, "ui_control_disabled")
	user32.NewProc("EnableWindow").Call(button, 1)
	if _, err := uiOperation(pid+1, t.TempDir(), true, selected.Selector); fault(err).Code != "window_not_owned" {
		t.Fatal(err)
	}
	if _, err := uiOperation(pid, t.TempDir(), true, selected.Selector); err != nil {
		t.Fatal(err)
	}
	pump()
	if !invoked {
		t.Fatal("native button did not deliver its command")
	}
	op.Named["message_pattern"] = "^Different prompt$"
	expectFault(op, true, "ui_selector_not_found")
	delete(op.Named, "message_pattern")
	create("Button", "N&o", 0x40000000, dialog)
	expectFault(op, true, "ui_selector_not_unique")
	literal := create("Button", "Save && Close", 0x40000000, dialog)
	if name, ok := nativeButtonName(literal); !ok || name != "Save & Close" {
		t.Fatalf("literal caption: %q", name)
	}
	group := create("Button", "Group", 0x40000007, dialog)
	if _, ok := nativeButtonName(group); ok {
		t.Fatal("group box treated as push button")
	}
}
