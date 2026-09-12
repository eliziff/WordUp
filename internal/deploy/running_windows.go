//go:build windows

package deploy

import (
	"fmt"
	"strings"
	"syscall"
	"unsafe"
)

var k = syscall.NewLazyDLL("kernel32.dll")
var snapshot = k.NewProc("CreateToolhelp32Snapshot")
var first = k.NewProc("Process32FirstW")
var next = k.NewProc("Process32NextW")
var closeH = k.NewProc("CloseHandle")

type processEntry struct {
	Size, Usage, PID          uint32
	Heap                      uintptr
	ModuleID, Threads, Parent uint32
	Priority                  int32
	Flags                     uint32
	Exe                       [260]uint16
}

func WordRunning() (bool, error) {
	h, _, e := snapshot.Call(2, 0)
	if h == ^uintptr(0) {
		return false, e
	}
	defer closeH.Call(h)
	p := processEntry{Size: uint32(unsafe.Sizeof(processEntry{}))}
	ok, _, e := first.Call(h, uintptr(unsafe.Pointer(&p)))
	if ok == 0 {
		if e == syscall.Errno(18) {
			return false, nil
		}
		return false, fmt.Errorf("process inventory: %w", e)
	}
	for {
		if strings.EqualFold(syscall.UTF16ToString(p.Exe[:]), "WINWORD.EXE") {
			return true, nil
		}
		ok, _, e = next.Call(h, uintptr(unsafe.Pointer(&p)))
		if ok == 0 {
			if e != syscall.Errno(18) {
				return false, e
			}
			return false, nil
		}
	}
}
