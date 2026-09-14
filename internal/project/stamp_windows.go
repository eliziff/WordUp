//go:build windows

package project

import (
	"os"
	"syscall"
	"unsafe"
)

// FILE_BASIC_INFO is the stable metadata Windows updates when a file is
// changed. Last-write time alone is insufficient: an editor can restore it
// after writing a same-size file. ChangeTime lets the resident cache reuse a
// verified content hash without trusting that weaker timestamp pair.
type fileBasicInfo struct {
	CreationTime   int64
	LastAccessTime int64
	LastWriteTime  int64
	ChangeTime     int64
	Attributes     uint32
	Reserved       uint32
}

var getFileInformationByHandleEx = syscall.NewLazyDLL("kernel32.dll").NewProc("GetFileInformationByHandleEx")

func fileChangeStamp(path string) (int64, bool) {
	f, err := os.Open(path)
	if err != nil {
		return 0, false
	}
	defer f.Close()
	var info fileBasicInfo
	result, _, _ := getFileInformationByHandleEx.Call(f.Fd(), 0, uintptr(unsafe.Pointer(&info)), unsafe.Sizeof(info))
	if result == 0 {
		return 0, false
	}
	return info.ChangeTime, true
}
