//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"syscall"
	"unicode/utf16"
	"unsafe"
)

var advapi = syscall.NewLazyDLL("advapi32.dll")
var regOpen = advapi.NewProc("RegOpenKeyExW")
var regQuery = advapi.NewProc("RegQueryValueExW")
var regClose = advapi.NewProc("RegCloseKey")

const hkcu = uintptr(0x80000001)
const hklm = uintptr(0x80000002)
const hkcr = uintptr(0x80000000)

func registryRead(root uintptr, path, name string, view uint32) ([]byte, uint32, error) {
	p, e := utf(path)
	if e != nil {
		return nil, 0, e
	}
	var key uintptr
	code, _, _ := regOpen.Call(root, uintptr(unsafe.Pointer(p)), 0, uintptr(0x20019|view), uintptr(unsafe.Pointer(&key)))
	if code != 0 {
		return nil, 0, syscall.Errno(code)
	}
	defer regClose.Call(key)
	np, e := utf(name)
	if e != nil {
		return nil, 0, e
	}
	var typ, size uint32
	code, _, _ = regQuery.Call(key, uintptr(unsafe.Pointer(np)), 0, uintptr(unsafe.Pointer(&typ)), 0, uintptr(unsafe.Pointer(&size)))
	if code != 0 {
		return nil, 0, syscall.Errno(code)
	}
	if size > 1<<20 {
		return nil, 0, fmt.Errorf("registry value budget")
	}
	b := make([]byte, size+2)
	code, _, _ = regQuery.Call(key, uintptr(unsafe.Pointer(np)), 0, uintptr(unsafe.Pointer(&typ)), uintptr(unsafe.Pointer(&b[0])), uintptr(unsafe.Pointer(&size)))
	if code != 0 {
		return nil, 0, syscall.Errno(code)
	}
	return b[:size], typ, nil
}
func registryString(root uintptr, path, name string, view uint32) (string, error) {
	b, t, e := registryRead(root, path, name, view)
	if e != nil {
		return "", e
	}
	if t != 1 && t != 2 {
		return "", fmt.Errorf("registry value is not a string")
	}
	if len(b)%2 != 0 {
		return "", fmt.Errorf("odd registry UTF16 length")
	}
	a := make([]uint16, len(b)/2)
	for i := range a {
		a[i] = uint16(b[2*i]) | uint16(b[2*i+1])<<8
	}
	return strings.TrimRight(string(utf16.Decode(a)), "\x00"), nil
}
func registryDWORD(root uintptr, path, name string, view uint32) (uint32, error) {
	b, t, e := registryRead(root, path, name, view)
	if e != nil {
		return 0, e
	}
	if t != 4 || len(b) != 4 {
		return 0, fmt.Errorf("registry value is not DWORD")
	}
	return uint32(b[0]) | uint32(b[1])<<8 | uint32(b[2])<<16 | uint32(b[3])<<24, nil
}
func findWord() (string, error) {
	for _, root := range []uintptr{hkcu, hklm} {
		for _, view := range []uint32{0x100, 0x200} {
			s, e := registryString(root, `SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE`, "", view)
			if e == nil {
				s = strings.Trim(s, `"`)
				if info, e := os.Stat(s); e == nil && info.Mode().IsRegular() {
					return filepath.Clean(s), nil
				}
			}
		}
	}
	for _, env := range []string{"ProgramW6432", "ProgramFiles", "ProgramFiles(x86)"} {
		base := os.Getenv(env)
		if base == "" {
			continue
		}
		for _, tail := range []string{`Microsoft Office\root\Office16\WINWORD.EXE`, `Microsoft Office\Office16\WINWORD.EXE`, `Microsoft Office\Office15\WINWORD.EXE`} {
			p := filepath.Join(base, tail)
			if i, e := os.Stat(p); e == nil && i.Mode().IsRegular() {
				return p, nil
			}
		}
	}
	return "", Fail("word_not_installed", "Install/activate Microsoft Word on this computer; WordUp does not contain or emulate Microsoft's VBA runtime", nil)
}
