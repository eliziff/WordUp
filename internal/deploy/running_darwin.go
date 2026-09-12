//go:build darwin

package deploy

import (
	"fmt"
	"runtime"
	"strings"
	"unsafe"
	"wordwright.local/internal/darwinapi"
)

func WordRunning() (bool, error) {
	pids := make([]int32, 65536)
	n := int32(darwinapi.Call(darwinapi.ProcListAllPIDs(), uintptr(unsafe.Pointer(&pids[0])), uintptr(len(pids)*4)))
	runtime.KeepAlive(pids)
	if n < 0 || int(n) >= len(pids) {
		return false, fmt.Errorf("native process inventory unavailable or beyond budget")
	}
	buf := make([]byte, 4096)
	for _, pid := range pids[:n] {
		if pid <= 0 {
			continue
		}
		size := int32(darwinapi.Call(darwinapi.ProcPIDPath(), uintptr(pid), uintptr(unsafe.Pointer(&buf[0])), uintptr(len(buf))))
		runtime.KeepAlive(buf)
		if size > 0 && int(size) <= len(buf) && strings.Contains(string(buf[:size]), "Microsoft Word.app/Contents/MacOS/Microsoft Word") {
			return true, nil
		}
	}
	return false, nil
}
