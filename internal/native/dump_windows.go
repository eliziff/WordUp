//go:build windows && (amd64 || arm64)

package native

import (
	"fmt"
	"os"
	"path/filepath"
	"sync"
	"syscall"
	"time"
	"unsafe"

	"github.com/eliziff/WordUp/internal/office"
)

var dbgHelp = syscall.NewLazyDLL("dbghelp.dll")
var dumpMu sync.Mutex

// Capture the owned process from the controller: this does not depend on Word's
// COM/UI threads responding. Normal + thread-info excludes a full heap dump.
func (h *localHost) dumpProcess(op Operation) (any, error) {
	if op.File == "" {
		return nil, fmt.Errorf("process.dump requires an output file")
	}
	h.mu.Lock()
	rawPID := h.info["pid"]
	job := h.job
	h.mu.Unlock()
	pid, ok := rawPID.(float64)
	if !ok || pid <= 0 || job == 0 {
		return nil, Fail("owned_process_unavailable", "No owned Word process is available", nil)
	}
	process, _, e := kernel.NewProc("OpenProcess").Call(0x410, 0, uintptr(uint32(pid)))
	if process == 0 {
		return nil, winError("OpenProcess for dump", e)
	}
	defer closeHandle.Call(process)
	var owned int32
	result, _, e := kernel.NewProc("IsProcessInJob").Call(process, job, uintptr(unsafe.Pointer(&owned)))
	if result == 0 {
		return nil, winError("IsProcessInJob", e)
	}
	if owned == 0 {
		return nil, Fail("word_ownership_mismatch", "Refusing to dump a process outside the owned job", nil)
	}
	if e := os.MkdirAll(filepath.Dir(op.File), 0700); e != nil {
		return nil, e
	}
	file, e := os.CreateTemp(filepath.Dir(op.File), ".wordup-dump-*")
	if e != nil {
		return nil, e
	}
	temporary := file.Name()
	defer os.Remove(temporary)
	started := time.Now()
	dumpMu.Lock()
	result, _, e = dbgHelp.NewProc("MiniDumpWriteDump").Call(process, uintptr(uint32(pid)), file.Fd(), 0x1000, 0, 0, 0)
	dumpMu.Unlock()
	closeErr := file.Close()
	if result == 0 {
		return nil, winError("MiniDumpWriteDump", e)
	}
	if closeErr != nil {
		return nil, closeErr
	}
	if e := os.Rename(temporary, op.File); e != nil {
		return nil, e
	}
	data, e := os.ReadFile(op.File)
	if e != nil {
		return nil, e
	}
	return map[string]any{"file": op.File, "pid": uint32(pid), "bytes": len(data), "sha256": office.Hash(data), "format": "Windows minidump", "dump_type": "normal_with_thread_info", "full_memory": false, "owned_job_verified": true, "duration_ms": float64(time.Since(started).Microseconds()) / 1000}, nil
}
