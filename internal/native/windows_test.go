//go:build windows && (amd64 || arm64)

package native

import (
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
