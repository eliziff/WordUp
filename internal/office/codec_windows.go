//go:build windows

package office

import (
	"fmt"
	"syscall"
	"unicode/utf16"
	"unsafe"
)

var kernelCodec = syscall.NewLazyDLL("kernel32.dll")

func platformDecode(b []byte, cp int) (string, error) {
	if len(b) == 0 {
		return "", nil
	}
	p := kernelCodec.NewProc("MultiByteToWideChar")
	n, _, e := p.Call(uintptr(cp), 0, uintptr(unsafe.Pointer(&b[0])), uintptr(len(b)), 0, 0)
	if n == 0 {
		return "", fmt.Errorf("decode codepage %d: %w", cp, e)
	}
	w := make([]uint16, n)
	r, _, e := p.Call(uintptr(cp), 0, uintptr(unsafe.Pointer(&b[0])), uintptr(len(b)), uintptr(unsafe.Pointer(&w[0])), n)
	if r == 0 {
		return "", e
	}
	return string(utf16.Decode(w)), nil
}
func platformEncode(s string, cp int) ([]byte, error) {
	if s == "" {
		return nil, nil
	}
	w := utf16.Encode([]rune(s))
	p := kernelCodec.NewProc("WideCharToMultiByte")
	var used uint32
	n, _, e := p.Call(uintptr(cp), 0x400, uintptr(unsafe.Pointer(&w[0])), uintptr(len(w)), 0, 0, 0, uintptr(unsafe.Pointer(&used)))
	if n == 0 {
		return nil, e
	}
	b := make([]byte, n)
	r, _, e := p.Call(uintptr(cp), 0x400, uintptr(unsafe.Pointer(&w[0])), uintptr(len(w)), uintptr(unsafe.Pointer(&b[0])), n, 0, uintptr(unsafe.Pointer(&used)))
	if r == 0 {
		return nil, e
	}
	if used != 0 {
		return nil, fmt.Errorf("lossy codepage conversion rejected")
	}
	return b, nil
}
