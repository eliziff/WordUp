package signing

import (
	"fmt"
	"runtime"
	"syscall"
	"unsafe"
)

// Hash-only verification checks the Office SIP digest without treating an
// untrusted self-signed certificate as a trusted publisher. Trust is a separate
// SignTool /pa result. No certificate stores or security policies are changed.
func VerifyDigestWorker(file string) error {
	path, err := syscall.UTF16PtrFromString(file)
	if err != nil {
		return err
	}
	type fileInfo struct {
		Size    uint32
		Path    *uint16
		Handle  uintptr
		Subject uintptr
	}
	type trustData struct {
		Size                   uint32
		Policy, SIP            uintptr
		UI, Revocation, Choice uint32
		File                   *fileInfo
		Action                 uint32
		State                  uintptr
		URL                    uintptr
		Flags, UIContext       uint32
		Settings               uintptr
	}
	f := fileInfo{Path: path}
	f.Size = uint32(unsafe.Sizeof(f))
	d := trustData{UI: 2, Choice: 1, File: &f, Flags: 0x210, Action: 1}
	d.Size = uint32(unsafe.Sizeof(d))
	action := [16]byte{0x6b, 0xc5, 0xaa, 0x00, 0x44, 0xcd, 0xd0, 0x11, 0x8c, 0xc2, 0x00, 0xc0, 0x4f, 0xc2, 0x95, 0xee}
	verify := syscall.NewLazyDLL("wintrust.dll").NewProc("WinVerifyTrust")
	result, _, _ := verify.Call(^uintptr(0), uintptr(unsafe.Pointer(&action)), uintptr(unsafe.Pointer(&d)))
	d.Action = 2
	verify.Call(^uintptr(0), uintptr(unsafe.Pointer(&action)), uintptr(unsafe.Pointer(&d)))
	runtime.KeepAlive(f)
	runtime.KeepAlive(d)
	if uint32(result) != 0 {
		return fmt.Errorf("Office VBA digest verification failed: 0x%08x", uint32(result))
	}
	return nil
}
