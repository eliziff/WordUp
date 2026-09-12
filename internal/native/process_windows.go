//go:build windows && (amd64 || arm64)

package native

import (
	"bufio"
	"context"
	"crypto/rand"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"sync/atomic"
	"syscall"
	"time"
	"unsafe"

	"wordwright.local/internal/office"
)

var kernel = syscall.NewLazyDLL("kernel32.dll")
var user32 = syscall.NewLazyDLL("user32.dll")
var createDesktop = user32.NewProc("CreateDesktopW")
var closeDesktop = user32.NewProc("CloseDesktop")
var createJob = kernel.NewProc("CreateJobObjectW")
var setJob = kernel.NewProc("SetInformationJobObject")
var assignJob = kernel.NewProc("AssignProcessToJobObject")
var terminateJob = kernel.NewProc("TerminateJobObject")
var createProcess = kernel.NewProc("CreateProcessW")
var initAttributes = kernel.NewProc("InitializeProcThreadAttributeList")
var updateAttributes = kernel.NewProc("UpdateProcThreadAttribute")
var deleteAttributes = kernel.NewProc("DeleteProcThreadAttributeList")
var resumeThread = kernel.NewProc("ResumeThread")
var waitSingle = kernel.NewProc("WaitForSingleObject")
var terminateProcess = kernel.NewProc("TerminateProcess")
var closeHandle = kernel.NewProc("CloseHandle")
var duplicateHandle = kernel.NewProc("DuplicateHandle")
var currentProcess = kernel.NewProc("GetCurrentProcess")

// The job is lifetime containment, not a security sandbox. Windows kills only
// the processes that this application explicitly owns, including on a crash.
type jobLimits struct {
	PerProcess, PerJob   int64
	Flags                uint32
	Minimum, Maximum     uintptr
	ActiveProcessLimit   uint32
	Affinity             uintptr
	Priority, Scheduling uint32
}
type ioCounters struct{ ReadOps, WriteOps, OtherOps, ReadBytes, WriteBytes, OtherBytes uint64 }
type extendedLimits struct {
	Basic                                          jobLimits
	IO                                             ioCounters
	ProcessMemory, JobMemory, PeakProcess, PeakJob uintptr
}
type startupEX struct {
	Info       syscall.StartupInfo
	Attributes uintptr
}
type childProcess struct {
	Process, Thread syscall.Handle
	PID             uint32
}

type hostConfig struct {
	Options
	Desktop  string `json:"desktop"`
	Token    string `json:"token"`
	WordPath string `json:"word_path"`
}
type localHost struct {
	cfg          hostConfig
	desktop, job uintptr
	process      childProcess
	input        *os.File
	output       *os.File
	log          *os.File
	send         sync.Mutex
	mu           sync.Mutex
	pending      map[uint64]chan Response
	sequence     atomic.Uint64
	closed       chan struct{}
	once         sync.Once
	info         map[string]any
	logMu        sync.Mutex
	logTail      []byte
}

func randomID() string {
	var b [16]byte
	if _, e := rand.Read(b[:]); e != nil {
		panic(e)
	}
	return hex.EncodeToString(b[:])
}
func utf(s string) (*uint16, error) { return syscall.UTF16PtrFromString(s) }
func winError(api string, e error) error {
	if e == nil || e == syscall.Errno(0) {
		return fmt.Errorf("%s failed", api)
	}
	return fmt.Errorf("%s: %w", api, e)
}
func duplicateInheritable(f *os.File) (syscall.Handle, error) {
	var out syscall.Handle
	self, _, _ := currentProcess.Call()
	r, _, e := duplicateHandle.Call(self, f.Fd(), self, uintptr(unsafe.Pointer(&out)), 0, 1, 2)
	if r == 0 {
		return 0, winError("DuplicateHandle", e)
	}
	return out, nil
}
func spawnOnDesktop(exe string, args []string, desktop string, stdin, stdout, stderr *os.File, job uintptr) (childProcess, error) {
	var result childProcess
	ep, e := utf(exe)
	if e != nil {
		return result, e
	}
	all := append([]string{exe}, args...)
	quoted := make([]string, len(all))
	for i, s := range all {
		quoted[i] = syscall.EscapeArg(s)
	}
	command, e := utf(strings.Join(quoted, " "))
	if e != nil {
		return result, e
	}
	dp, e := utf(desktop)
	if e != nil {
		return result, e
	}
	handles := make([]syscall.Handle, 3)
	for i, f := range []*os.File{stdin, stdout, stderr} {
		h, e := duplicateInheritable(f)
		if e != nil {
			for _, old := range handles {
				if old != 0 {
					closeHandle.Call(uintptr(old))
				}
			}
			return result, e
		}
		handles[i] = h
	}
	defer func() {
		for _, h := range handles {
			closeHandle.Call(uintptr(h))
		}
	}()
	var n uintptr
	initAttributes.Call(0, 1, 0, uintptr(unsafe.Pointer(&n)))
	if n == 0 {
		return result, fmt.Errorf("cannot size process attribute list")
	}
	storage := make([]byte, n)
	attr := uintptr(unsafe.Pointer(&storage[0]))
	r, _, e := initAttributes.Call(attr, 1, 0, uintptr(unsafe.Pointer(&n)))
	if r == 0 {
		return result, winError("InitializeProcThreadAttributeList", e)
	}
	defer deleteAttributes.Call(attr)
	r, _, e = updateAttributes.Call(attr, 0, 0x00020002, uintptr(unsafe.Pointer(&handles[0])), uintptr(len(handles))*unsafe.Sizeof(handles[0]), 0, 0)
	if r == 0 {
		return result, winError("UpdateProcThreadAttribute(handle list)", e)
	}
	si := startupEX{Attributes: attr}
	si.Info.Cb = uint32(unsafe.Sizeof(si))
	si.Info.Desktop = dp
	si.Info.Flags = 0x100
	si.Info.StdInput = handles[0]
	si.Info.StdOutput = handles[1]
	si.Info.StdErr = handles[2]
	var pi syscall.ProcessInformation
	// A process is put in the job before its first instruction, eliminating the
	// launch/timeout race that can leave a stray Word process behind.
	r, _, e = createProcess.Call(uintptr(unsafe.Pointer(ep)), uintptr(unsafe.Pointer(command)), 0, 0, 1, 0x00080000|0x00000004|0x08000000, 0, 0, uintptr(unsafe.Pointer(&si)), uintptr(unsafe.Pointer(&pi)))
	runtime.KeepAlive(storage)
	runtime.KeepAlive(handles)
	runtime.KeepAlive(command)
	runtime.KeepAlive(dp)
	if r == 0 {
		return result, winError("CreateProcessW", e)
	}
	result = childProcess{pi.Process, pi.Thread, pi.ProcessId}
	if job != 0 {
		r, _, e = assignJob.Call(job, uintptr(pi.Process))
		if r == 0 {
			terminateProcess.Call(uintptr(pi.Process), 1)
			closeHandle.Call(uintptr(pi.Thread))
			closeHandle.Call(uintptr(pi.Process))
			return childProcess{}, winError("AssignProcessToJobObject", e)
		}
	}
	r, _, e = resumeThread.Call(uintptr(pi.Thread))
	closeHandle.Call(uintptr(pi.Thread))
	result.Thread = 0
	if uint32(r) == 0xffffffff {
		terminateProcess.Call(uintptr(pi.Process), 1)
		closeHandle.Call(uintptr(pi.Process))
		return childProcess{}, winError("ResumeThread", e)
	}
	return result, nil
}
func Available() bool { _, e := findWord(); return e == nil }
func Start(ctx context.Context, opt Options) (Host, error) {
	word, e := findWord()
	if e != nil {
		return nil, e
	}
	base := opt.Directory
	if base == "" {
		base = filepath.Join(os.TempDir(), "Wordwright")
	}
	if e = os.MkdirAll(base, 0700); e != nil {
		return nil, e
	}
	stage, e := os.MkdirTemp(base, "session-")
	if e != nil {
		return nil, e
	}
	opt.Directory = stage
	h := &localHost{cfg: hostConfig{Options: opt, Desktop: "Wordwright-" + randomID(), Token: randomID(), WordPath: word}, pending: map[uint64]chan Response{}, closed: make(chan struct{}), info: map[string]any{}}
	ok := false
	defer func() {
		if !ok {
			h.Close()
		}
	}()
	dn, _ := utf(h.cfg.Desktop)
	hd, _, er := createDesktop.Call(uintptr(unsafe.Pointer(dn)), 0, 0, 0, 0x000F01FF, 0)
	if hd == 0 {
		return nil, winError("CreateDesktopW", er)
	}
	h.desktop = hd
	j, _, er := createJob.Call(0, 0)
	if j == 0 {
		return nil, winError("CreateJobObjectW", er)
	}
	h.job = j
	limits := extendedLimits{}
	limits.Basic.Flags = 0x2000
	r, _, er := setJob.Call(j, 9, uintptr(unsafe.Pointer(&limits)), unsafe.Sizeof(limits))
	if r == 0 {
		return nil, winError("SetInformationJobObject", er)
	}
	// This seed has no VBA, embedded objects or external relationships.
	seed := office.BlankPackage()
	data, e := seed.Bytes()
	if e != nil {
		return nil, e
	}
	if e = os.WriteFile(filepath.Join(stage, "seed.docx"), data, 0600); e != nil {
		return nil, e
	}
	cfg, _ := json.Marshal(h.cfg)
	configPath := filepath.Join(stage, "host.json")
	if e = os.WriteFile(configPath, cfg, 0600); e != nil {
		return nil, e
	}
	inRead, inWrite, e := os.Pipe()
	if e != nil {
		return nil, e
	}
	defer inRead.Close()
	h.input = inWrite
	outRead, outWrite, e := os.Pipe()
	if e != nil {
		return nil, e
	}
	defer outWrite.Close()
	h.output = outRead
	logRead, logWrite, e := os.Pipe()
	if e != nil {
		return nil, e
	}
	defer logWrite.Close()
	h.log = logRead
	exe, e := os.Executable()
	if e != nil {
		return nil, e
	}
	h.process, e = spawnOnDesktop(exe, []string{"__host", configPath}, "WinSta0\\"+h.cfg.Desktop, inRead, outWrite, logWrite, h.job)
	if e != nil {
		return nil, e
	}
	go h.readResponses()
	go func() {
		buf := make([]byte, 4096)
		for {
			n, e := logRead.Read(buf)
			if n > 0 {
				h.logMu.Lock()
				h.logTail = append(h.logTail, buf[:n]...)
				if len(h.logTail) > 65536 {
					h.logTail = append([]byte(nil), h.logTail[len(h.logTail)-65536:]...)
				}
				h.logMu.Unlock()
			}
			if e != nil {
				return
			}
		}
	}()
	boot, cancel := deadline(ctx, opt.StartupTimeoutMS)
	defer cancel()
	result, e := h.Call(boot, Operation{Op: "hello", Value: h.cfg.Token})
	if e != nil {
		return nil, e
	}
	m, okInfo := result.(map[string]any)
	if !okInfo {
		return nil, fmt.Errorf("worker returned invalid handshake")
	}
	h.info = m
	ok = true
	return h, nil
}
func (h *localHost) readResponses() {
	s := bufio.NewScanner(h.output)
	s.Buffer(make([]byte, 65536), 32<<20)
	for s.Scan() {
		var r Response
		if e := json.Unmarshal(s.Bytes(), &r); e != nil {
			h.breakPending(Fail("worker_protocol_error", e.Error(), nil))
			return
		}
		h.mu.Lock()
		ch := h.pending[r.ID]
		delete(h.pending, r.ID)
		h.mu.Unlock()
		if ch != nil {
			ch <- r
		}
	}
	e := s.Err()
	if e == nil {
		e = io.EOF
	}
	h.breakPending(Fail("worker_exited", e.Error(), nil))
}
func (h *localHost) breakPending(e error) {
	h.mu.Lock()
	defer h.mu.Unlock()
	for id, ch := range h.pending {
		ch <- Response{ID: id, Error: fault(e)}
		delete(h.pending, id)
	}
}
func (h *localHost) Call(ctx context.Context, op Operation) (any, error) {
	select {
	case <-h.closed:
		return nil, Fail("session_closed", "native session has been closed", nil)
	default:
	}
	id := h.sequence.Add(1)
	ch := make(chan Response, 1)
	h.mu.Lock()
	h.pending[id] = ch
	h.mu.Unlock()
	written := make(chan error, 1)
	go func() {
		h.send.Lock()
		defer h.send.Unlock()
		written <- json.NewEncoder(h.input).Encode(Request{ID: id, Operation: op})
	}()
	var e error
	select {
	case e = <-written:
	case <-ctx.Done():
		h.Close()
		return nil, Fail("native_deadline", "Native request write timed out; only the owned Word job was terminated", ctx.Err().Error())
	case <-h.closed:
		return nil, Fail("session_closed", "native session closed", nil)
	}
	if e != nil {
		h.mu.Lock()
		delete(h.pending, id)
		h.mu.Unlock()
		return nil, e
	}
	select {
	case r := <-ch:
		if r.Error != nil {
			return nil, r.Error
		}
		return r.Result, nil
	case <-ctx.Done():
		h.Close()
		h.logMu.Lock()
		tail := string(h.logTail)
		h.logMu.Unlock()
		return nil, Fail("native_deadline", "The app-owned Word job was terminated after the deadline; the user's Word processes were not touched", map[string]any{"cause": ctx.Err().Error(), "worker_log_tail": tail})
	case <-h.closed:
		return nil, Fail("session_closed", "native session closed", nil)
	}
}
func (h *localHost) Info() map[string]any {
	r := map[string]any{}
	for k, v := range h.info {
		r[k] = v
	}
	return r
}
func (h *localHost) Close() error {
	var cleanupError error
	h.once.Do(func() {
		close(h.closed)
		if h.input != nil {
			h.input.Close()
		}
		if h.process.Process != 0 {
			waitSingle.Call(uintptr(h.process.Process), 1500)
		}
		if h.job != 0 {
			terminateJob.Call(h.job, 0)
			closeHandle.Call(h.job)
		}
		if h.process.Process != 0 {
			waitSingle.Call(uintptr(h.process.Process), 5000)
			closeHandle.Call(uintptr(h.process.Process))
		}
		if h.output != nil {
			h.output.Close()
		}
		if h.log != nil {
			h.log.Close()
		}
		if h.desktop != 0 {
			closeDesktop.Call(h.desktop)
		}
		h.breakPending(Fail("session_closed", "native session closed", nil))
		if h.cfg.Directory != "" {
			cleanupError = os.RemoveAll(h.cfg.Directory)
		}
	})
	return cleanupError
}

var _ = time.Second
