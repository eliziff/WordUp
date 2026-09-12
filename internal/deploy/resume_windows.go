//go:build windows

package deploy

import (
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"syscall"
	"unicode/utf16"
	"unsafe"

	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

func activationName(file string) string {
	abs, _ := filepath.Abs(file)
	return "WordUp-" + office.Hash([]byte(strings.ToLower(abs)))[:24]
}

func lockActivation(file string) (func(), error) {
	runtime.LockOSThread()
	name, _ := syscall.UTF16PtrFromString("Local\\" + activationName(file))
	h, _, e := k.NewProc("CreateMutexW").Call(0, 0, uintptr(unsafe.Pointer(name)))
	if h == 0 {
		runtime.UnlockOSThread()
		return nil, e
	}
	result, _, _ := k.NewProc("WaitForSingleObject").Call(h, 0)
	if result != 0 && result != 0x80 {
		closeH.Call(h)
		runtime.UnlockOSThread()
		return nil, ErrActivationBusy
	}
	return func() { k.NewProc("ReleaseMutex").Call(h); closeH.Call(h); runtime.UnlockOSThread() }, nil
}

func resumeValue(file, command string) error {
	api := syscall.NewLazyDLL("advapi32.dll")
	path, _ := syscall.UTF16PtrFromString(`Software\Microsoft\Windows\CurrentVersion\Run`)
	var key syscall.Handle
	result, _, _ := api.NewProc("RegCreateKeyExW").Call(0x80000001, uintptr(unsafe.Pointer(path)), 0, 0, 0, 2, 0, uintptr(unsafe.Pointer(&key)), 0)
	if result != 0 {
		return syscall.Errno(result)
	}
	defer syscall.RegCloseKey(key)
	name, _ := syscall.UTF16PtrFromString(activationName(file))
	if command == "" {
		result, _, _ = api.NewProc("RegDeleteValueW").Call(uintptr(key), uintptr(unsafe.Pointer(name)))
		if result == 2 {
			return nil
		}
	} else {
		value, e := syscall.UTF16FromString(command)
		if e != nil {
			return e
		}
		result, _, _ = api.NewProc("RegSetValueExW").Call(uintptr(key), uintptr(unsafe.Pointer(name)), 0, 1, uintptr(unsafe.Pointer(&value[0])), uintptr(len(value)*2))
	}
	if result != 0 {
		return syscall.Errno(result)
	}
	return nil
}

func registerResume(file, exe string) error {
	// Keep a stable copy so moving the user's downloaded executable does not
	// strand an already authorized, exact-artifact deployment.
	dir := filepath.Dir(file)
	data, e := os.ReadFile(exe)
	if e != nil {
		return e
	}
	worker := filepath.Join(dir, "wordup-activate.exe")
	if e = project.AtomicWrite(worker, data); e != nil {
		return e
	}
	command := syscall.EscapeArg(worker) + " __activate " + syscall.EscapeArg(file)
	script := filepath.Join(dir, "resume.vbs")
	body := "CreateObject(\"WScript.Shell\").Run \"" + strings.ReplaceAll(command, "\"", "\"\"") + "\", 0, False\r\n"
	encoded := []byte{0xff, 0xfe}
	for _, c := range utf16.Encode([]rune(body)) {
		encoded = append(encoded, byte(c), byte(c>>8))
	}
	if e = project.AtomicWrite(script, encoded); e != nil {
		return e
	}
	wscript := filepath.Join(os.Getenv("SystemRoot"), "System32", "wscript.exe")
	return resumeValue(file, syscall.EscapeArg(wscript)+" //B //Nologo "+syscall.EscapeArg(script))
}

func clearResume(file string) error { return resumeValue(file, "") }
