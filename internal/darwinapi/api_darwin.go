//go:build darwin && (amd64 || arm64)

// Package darwinapi is a small direct bridge to system frameworks. Its Darwin
// ABI trampolines are linked into the executable; no osascript, shell, Python,
// downloaded helper, or Objective-C runtime dependency is used by the runner.
package darwinapi

import (
	"runtime"
	"syscall"
	"unsafe"
)

// This is the libc-call bridge maintained by Go for syscall and x/sys. The
// released executable fixes its Go runtime; callers do not install a toolchain.
//
//go:linkname call6 syscall.syscall6
func call6(fn, a1, a2, a3, a4, a5, a6 uintptr) (r1, r2 uintptr, err syscall.Errno)
func Call(fn uintptr, args ...uintptr) uintptr {
	var a [6]uintptr
	copy(a[:], args)
	r, _, _ := call6(fn, a[0], a[1], a[2], a[3], a[4], a[5])
	runtime.KeepAlive(args)
	return r
}
func Pointer[T any](p *T) uintptr { return uintptr(unsafe.Pointer(p)) }
