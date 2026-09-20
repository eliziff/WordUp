//go:build windows && (amd64 || arm64)

package native

import (
	"bufio"
	"context"
	"crypto/rand"
	"encoding/hex"
	"encoding/json"
	"errors"
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

	"github.com/eliziff/WordUp/internal/office"
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

const (
	createSuspended     = 0x00000004
	createNoWindow      = 0x08000000
	createBreakaway     = 0x01000000
	extendedStartupInfo = 0x00080000
)

func processCreationFlags(job uintptr) uint32 {
	// Process groups are not needed for containment (the owned job supplies
	// that boundary) and some restricted logon sessions reject the extra flag.
	// Keep only the flags required for suspended handoff, hidden startup and
	// STARTUPINFOEX handle attributes.
	flags := uint32(createSuspended | createNoWindow | extendedStartupInfo)
	// The host is created from the caller and must break out of any enclosing
	// job before it is assigned to WordUp's owned job. Word itself is launched
	// by that host and must inherit the owned job; asking it to break away can
	// fail with ERROR_NO_SUCH_LOGON_SESSION in restricted interactive sessions.
	if job != 0 {
		flags |= createBreakaway
	}
	return flags
}

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
type jobAccounting struct {
	User, Kernel, PeriodUser, PeriodKernel int64
	Faults, Total, Active, Terminated      uint32
}
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
	tasks        map[string]*taskCompletion
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
	closeErr     error
	info         map[string]any
	logMu        sync.Mutex
	logTail      []byte
	termination  *Fault
	wordExit     *Fault
}

type taskCompletion struct {
	done      chan struct{}
	result    any
	err       error
	completed bool
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

func noLogonSessionError(err error) bool {
	return errors.Is(err, syscall.Errno(1312)) || strings.Contains(strings.ToLower(err.Error()), "specified logon session does not exist")
}

func callerLaunchTokenFailure(details map[string]any) error {
	details["likely_cause"] = "WordUp was launched with a restricted Windows process token"
	details["recovery"] = "Rerun the native WordUp command outside the command sandbox"
	details["session_restart_helpful"] = false
	return Fail("native_launch_token_restricted", "The caller's Windows launch token cannot start Microsoft Word", details)
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
func spawnOnDesktop(exe string, args []string, desktop string, hidden bool, stdin, stdout, stderr *os.File, job uintptr) (childProcess, error) {
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
	var dp *uint16
	if desktop != "" {
		dp, e = utf(desktop)
		if e != nil {
			return result, e
		}
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
	// Keep the extended startup-info structure and inherited-handle whitelist in
	// every branch: the creation flags always include
	// EXTENDED_STARTUPINFO_PRESENT, so Cb must always be the full
	// STARTUPINFOEXW size.
	si := startupEX{Attributes: attr}
	si.Info.Cb = uint32(unsafe.Sizeof(si))
	r, _, e = updateAttributes.Call(attr, 0, 0x00020002, uintptr(unsafe.Pointer(&handles[0])), uintptr(len(handles))*unsafe.Sizeof(handles[0]), 0, 0)
	if r == 0 {
		return result, winError("UpdateProcThreadAttribute(handle list)", e)
	}
	si.Info.Desktop = dp
	si.Info.Flags = 0x100 // STARTF_USESTDHANDLES
	if hidden {
		si.Info.Flags |= 0x1   // STARTF_USESHOWWINDOW
		si.Info.ShowWindow = 0 // SW_HIDE
	}
	si.Info.StdInput = handles[0]
	si.Info.StdOutput = handles[1]
	si.Info.StdErr = handles[2]
	var pi syscall.ProcessInformation
	// A process is put in the job before its first instruction, eliminating the
	// launch/timeout race that can leave a stray Word process behind.
	r, _, e = createProcess.Call(uintptr(unsafe.Pointer(ep)), uintptr(unsafe.Pointer(command)), 0, 0, 1, uintptr(processCreationFlags(job)), 0, 0, uintptr(unsafe.Pointer(&si)), uintptr(unsafe.Pointer(&pi)))
	runtime.KeepAlive(storage)
	runtime.KeepAlive(handles)
	runtime.KeepAlive(command)
	runtime.KeepAlive(dp)
	if r == 0 {
		return result, fmt.Errorf("CreateProcessW (%s on %s): %w", exe, desktop, e)
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

// Start launches the app-owned Word session. Restricted logon sessions can let
// CreateDesktopW succeed yet refuse every CreateProcessW attempted from that
// desktop-bound worker (ERROR_NO_SUCH_LOGON_SESSION, 1312). When the startup
// handshake reports native_session_restricted, Start rebuilds the whole session
// once with the worker and Word hidden on the inherited desktop; the handshake
// then reports the loss of desktop isolation explicitly.
func Start(ctx context.Context, opt Options) (Host, error) {
	return start(ctx, opt, false)
}

func start(ctx context.Context, opt Options, degraded bool) (Host, error) {
	started := time.Now()
	if opt.MemoryLimitMB == 0 {
		opt.MemoryLimitMB = 2048
	}
	if opt.CPUPercent == 0 {
		opt.CPUPercent = 50
	}
	if opt.MemoryLimitMB < 256 || opt.MemoryLimitMB > 16384 || opt.CPUPercent < 1 || opt.CPUPercent > 100 {
		return nil, fmt.Errorf("native limits require memory_limit_mb 256..16384 and cpu_percent 1..100")
	}
	word, e := findWord()
	if e != nil {
		return nil, e
	}
	base := opt.Directory
	if base == "" {
		base = filepath.Join(os.TempDir(), "WordUp")
	}
	if e = os.MkdirAll(base, 0700); e != nil {
		return nil, e
	}
	stage, e := os.MkdirTemp(base, "session-")
	if e != nil {
		return nil, e
	}
	opt.Directory = stage
	h := &localHost{cfg: hostConfig{Options: opt, Desktop: "WordUp-" + randomID(), Token: randomID(), WordPath: word}, pending: map[uint64]chan Response{}, tasks: map[string]*taskCompletion{}, closed: make(chan struct{}), info: map[string]any{}}
	ok := false
	defer func() {
		if !ok {
			h.Close()
		}
	}()
	if opt.Visible {
		h.cfg.Desktop = "Default"
	} else if degraded {
		// Explicit degraded-isolation retry: the previous private-desktop
		// handshake failed with native_session_restricted, so this worker
		// stays hidden on the inherited desktop instead.
		h.cfg.Desktop = ""
	} else {
		dn, _ := utf(h.cfg.Desktop)
		hd, _, er := createDesktop.Call(uintptr(unsafe.Pointer(dn)), 0, 0, 0, 0x000F01FF, 0)
		if hd == 0 {
			desktopErr := winError("CreateDesktopW", er)
			// A restricted/service-like logon can reject the desktop object
			// itself with the same 1312 boundary that may otherwise appear at
			// CreateProcessW. Fall back before creating any owned processes; the
			// handshake will report degraded isolation explicitly.
			if !noLogonSessionError(desktopErr) {
				return nil, desktopErr
			}
			degraded = true
			h.cfg.Desktop = ""
		} else {
			h.desktop = hd
		}
	}
	j, _, er := createJob.Call(0, 0)
	if j == 0 {
		return nil, winError("CreateJobObjectW", er)
	}
	h.job = j
	limits := extendedLimits{}
	limits.Basic.Flags = 0x2000 | 0x200 | 0x20 | 0x8 // kill-on-close, job memory, priority, process count
	limits.Basic.ActiveProcessLimit = 16
	limits.Basic.Priority = 0x4000 // BELOW_NORMAL_PRIORITY_CLASS
	limits.JobMemory = uintptr(opt.MemoryLimitMB) << 20
	r, _, er := setJob.Call(j, 9, uintptr(unsafe.Pointer(&limits)), unsafe.Sizeof(limits))
	if r == 0 {
		return nil, winError("SetInformationJobObject", er)
	}
	cpu := struct{ Flags, Rate uint32 }{5, uint32(opt.CPUPercent * 100)} // enable + hard cap
	r, _, er = setJob.Call(j, 15, uintptr(unsafe.Pointer(&cpu)), unsafe.Sizeof(cpu))
	if r == 0 {
		return nil, winError("SetInformationJobObject(CPU)", er)
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
	spawnStarted := time.Now()
	// Resolve the name in the current window station; see connectWord for why
	// qualifying it with WinSta0 is not portable across interactive sessions.
	h.process, e = spawnOnDesktop(exe, []string{"__host", configPath}, h.cfg.Desktop, !h.cfg.Options.Visible, inRead, outWrite, logWrite, h.job)
	if e != nil {
		if !opt.Visible && noLogonSessionError(e) {
			// Some interactive/service-like logon sessions let us create a
			// desktop object but reject CreateProcessW when it names that
			// desktop. Keep the host and Word hidden on the inherited desktop so
			// COM/runtime verification can still run. The handshake reports the
			// loss of desktop isolation explicitly.
			initial := e
			privateDesktop := h.cfg.Desktop
			if h.desktop != 0 {
				closeDesktop.Call(h.desktop)
				h.desktop = 0
			}
			h.cfg.Desktop = ""
			cfg, marshalErr := json.Marshal(h.cfg)
			if marshalErr != nil {
				return nil, marshalErr
			}
			if writeErr := os.WriteFile(configPath, cfg, 0600); writeErr != nil {
				return nil, writeErr
			}
			h.process, e = spawnOnDesktop(exe, []string{"__host", configPath}, "", true, inRead, outWrite, logWrite, h.job)
			if e != nil {
				return nil, callerLaunchTokenFailure(map[string]any{
					"desktop":                     privateDesktop,
					"word_path":                   h.cfg.WordPath,
					"create_error":                e.Error(),
					"private_create_error":        initial.Error(),
					"inherited_desktop_attempted": true,
				})
			}
		}
		if e != nil {
			return nil, e
		}
	}
	spawnFinished := time.Now()
	// The parent must release its copies of the child's pipe ends immediately.
	// Keeping them until the handshake returns hides worker EOF on startup failure.
	inRead.Close()
	outWrite.Close()
	logWrite.Close()
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
		var restricted *Fault
		if !degraded && !opt.Visible && h.desktop != 0 && errors.As(e, &restricted) && restricted.Code == "native_session_restricted" {
			// The worker ran on the private desktop but every process it tried
			// to create there (and by inheritance from that thread) failed with
			// 1312. Rebuild the whole session once on the inherited desktop so
			// verification can still run; desktop_isolation stays false in the
			// handshake info so the lost sandbox is reported, never silent.
			h.Close()
			opt.Directory = base
			retry, retryErr := start(ctx, opt, true)
			if retryErr != nil {
				var retryFault *Fault
				if !errors.As(retryErr, &retryFault) {
					retryFault = &Fault{Code: "native_error", Message: retryErr.Error()}
				}
				return nil, callerLaunchTokenFailure(map[string]any{
					"private_desktop_fault":   restricted,
					"inherited_desktop_fault": retryFault,
				})
			}
			if retryHost, okHost := retry.(*localHost); okHost {
				retryHost.mu.Lock()
				retryHost.info["degraded_isolation"] = true
				retryHost.info["degraded_isolation_cause"] = restricted.Message
				retryHost.mu.Unlock()
			}
			return retry, nil
		}
		return nil, e
	}
	m, okInfo := result.(map[string]any)
	if !okInfo {
		return nil, fmt.Errorf("worker returned invalid handshake")
	}
	h.info = m
	h.info["worker_pid"] = h.process.PID
	h.info["startup_timing_ms"] = map[string]float64{"total": float64(time.Since(started).Microseconds()) / 1000, "prepare": float64(spawnStarted.Sub(started).Microseconds()) / 1000, "spawn_host": float64(spawnFinished.Sub(spawnStarted).Microseconds()) / 1000, "handshake": float64(time.Since(spawnFinished).Microseconds()) / 1000}
	h.info["limits"] = map[string]any{"memory_limit_mb": opt.MemoryLimitMB, "cpu_percent": opt.CPUPercent, "active_process_limit": 16, "default_operation_timeout_ms": 30000, "priority": "below_normal", "scope": "owned job including child processes"}
	h.info["word_alive"] = true
	h.info["word_exit_monitor"] = false
	if wordPID, ok := m["pid"].(float64); ok {
		process, _, _ := kernel.NewProc("OpenProcess").Call(0x101000, 0, uintptr(uint32(wordPID)))
		if process != 0 {
			h.info["word_exit_monitor"] = true
			go h.observeWordExit(process, uint32(wordPID))
		}
	}
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
		if r.Task != "" {
			if r.ID != 0 {
				h.tasks[r.Task] = &taskCompletion{done: make(chan struct{})}
			} else if task := h.tasks[r.Task]; task != nil && !task.completed {
				task.result = r.Result
				task.completed = true
				close(task.done)
			}
		}
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
	h.logMu.Lock()
	tail := string(h.logTail)
	h.logMu.Unlock()
	details := map[string]any{"worker_pid": h.process.PID, "worker_log_tail": tail}
	var exitCode uint32
	if syscall.GetExitCodeProcess(h.process.Process, &exitCode) == nil {
		details["worker_running"] = exitCode == 259 // STILL_ACTIVE
		if exitCode != 259 {
			details["worker_exit_code"] = exitCode
		}
	}
	err := Fail("worker_exited", e.Error(), details)
	h.mu.Lock()
	if h.termination == nil {
		select {
		case <-h.closed:
		default:
			h.termination = fault(err)
		}
	}
	h.mu.Unlock()
	h.breakPending(err)
	h.Close()
}
func (h *localHost) breakPending(e error) {
	h.mu.Lock()
	defer h.mu.Unlock()
	if h.termination != nil {
		e = h.termination
	}
	for id, ch := range h.pending {
		ch <- Response{ID: id, Error: fault(e)}
		delete(h.pending, id)
	}
}

func (h *localHost) closedFault() error {
	h.mu.Lock()
	defer h.mu.Unlock()
	if h.termination != nil {
		return h.termination
	}
	return Fail("session_closed", "native session closed", nil)
}

func (h *localHost) observeWordExit(process uintptr, pid uint32) {
	defer closeHandle.Call(process)
	waitSingle.Call(process, 0xffffffff)
	var code uint32
	kernel.NewProc("GetExitCodeProcess").Call(process, uintptr(unsafe.Pointer(&code)))
	h.mu.Lock()
	h.info["word_alive"] = false
	h.info["word_exit_code"] = code
	select {
	case <-h.closed:
		h.mu.Unlock()
		return
	default:
	}
	f := fault(Fail("word_process_exited", fmt.Sprintf("Owned Word process %d exited with code 0x%08X", pid, code), map[string]any{"pid": pid, "exit_code": code}))
	h.wordExit = f
	for _, task := range h.tasks {
		if !task.completed {
			task.err = f
			task.completed = true
			close(task.done)
		}
	}
	h.mu.Unlock()
	h.breakPending(f)
}
func (h *localHost) Call(ctx context.Context, op Operation) (any, error) {
	h.mu.Lock()
	wordExit := h.wordExit
	h.mu.Unlock()
	if wordExit != nil && !strings.HasPrefix(op.Op, "ui.") {
		return nil, wordExit
	}
	select {
	case <-h.closed:
		h.mu.Lock()
		reason := h.termination
		h.mu.Unlock()
		if reason != nil {
			return nil, reason
		}
		return nil, Fail("session_closed", "native session has been closed", nil)
	default:
	}
	if op.Op == "process.dump" {
		return h.dumpProcess(op)
	}
	if op.TimeoutMS < 0 || op.TimeoutMS > 1800000 {
		return nil, fmt.Errorf("operation timeout_ms must be 0..1800000")
	}
	if op.Op != "hello" && op.Op != "poll" && op.Op != "forget" {
		ms := op.TimeoutMS
		if ms == 0 {
			ms = 30000
		}
		var cancel context.CancelFunc
		ctx, cancel = context.WithTimeout(ctx, time.Duration(ms)*time.Millisecond)
		defer cancel()
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
		return nil, h.closedFault()
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
		if op.Op == "begin" {
			if m, ok := r.Result.(map[string]any); ok {
				if task, ok := m["task"].(string); ok {
					ms := op.TimeoutMS
					if ms == 0 && len(op.Steps) == 1 {
						ms = op.Steps[0].TimeoutMS
					}
					if ms == 0 {
						ms = 30000
					}
					go h.watchTask(task, time.Duration(ms)*time.Millisecond, op)
				}
			}
		}
		if op.Op == "forget" {
			key, _ := op.Value.(string)
			h.mu.Lock()
			if task := h.tasks[key]; task != nil && !task.completed {
				task.err = fmt.Errorf("task forgotten")
				task.completed = true
				close(task.done)
			}
			delete(h.tasks, key)
			h.mu.Unlock()
		}
		return r.Result, nil
	case <-ctx.Done():
		h.Close()
		h.logMu.Lock()
		tail := string(h.logTail)
		h.logMu.Unlock()
		return nil, Fail("native_deadline", "The app-owned Word job was terminated after the deadline; the user's Word processes were not touched", map[string]any{"cause": ctx.Err().Error(), "worker_log_tail": tail})
	case <-h.closed:
		return nil, h.closedFault()
	}
}

func (h *localHost) watchTask(task string, limit time.Duration, operation Operation) {
	h.mu.Lock()
	completion := h.tasks[task]
	h.mu.Unlock()
	if completion == nil {
		return
	}
	end := time.NewTimer(limit)
	defer end.Stop()
	select {
	case <-completion.done:
		return
	case <-h.closed:
		return
	case <-end.C:
		h.mu.Lock()
		if completion.completed {
			h.mu.Unlock()
			return
		}
		if h.termination == nil {
			h.termination = fault(Fail("macro_deadline", "The asynchronous task exceeded its wall-clock limit; only the owned Word job was terminated", map[string]any{"task": task, "timeout_ms": limit.Milliseconds(), "operation": operation}))
		}
		h.mu.Unlock()
		h.Close()
		return
	}
}

// WaitTask receives completion directly from the worker. A caller's deadline
// leaves the independent watchdog and UI diagnostic lane alive.
func (h *localHost) WaitTask(ctx context.Context, key string) (any, error) {
	h.mu.Lock()
	task := h.tasks[key]
	h.mu.Unlock()
	if task == nil {
		return nil, fmt.Errorf("unknown task")
	}
	select {
	case <-task.done:
		return task.result, task.err
	case <-ctx.Done():
		return nil, ctx.Err()
	case <-h.closed:
		h.mu.Lock()
		reason := h.termination
		h.mu.Unlock()
		if reason != nil {
			return nil, reason
		}
		return nil, Fail("session_closed", "native session closed", nil)
	}
}
func (h *localHost) Info() map[string]any {
	h.mu.Lock()
	defer h.mu.Unlock()
	r := map[string]any{}
	for k, v := range h.info {
		r[k] = v
	}
	select {
	case <-h.closed:
		r["closed"] = true
	default:
		r["closed"] = false
	}
	if r["closed"] != true && h.job != 0 {
		query := kernel.NewProc("QueryInformationJobObject")
		var limits extendedLimits
		if ok, _, _ := query.Call(h.job, 9, uintptr(unsafe.Pointer(&limits)), unsafe.Sizeof(limits), 0); ok != 0 {
			r["job_peak_memory_bytes"] = uint64(limits.PeakJob)
		}
		var accounting jobAccounting
		if ok, _, _ := query.Call(h.job, 1, uintptr(unsafe.Pointer(&accounting)), unsafe.Sizeof(accounting), 0); ok != 0 {
			r["job_cpu_ms"] = float64(accounting.User+accounting.Kernel) / 10000
			r["job_active_processes"] = accounting.Active
		}
	}
	return r
}
func (h *localHost) Close() error {
	var cleanupError error
	h.once.Do(func() {
		defer func() { h.closeErr = cleanupError }()
		close(h.closed)
		if h.input != nil {
			h.input.Close()
		}
		if h.process.Process != 0 {
			waitSingle.Call(uintptr(h.process.Process), 1500)
		}
		if h.job != 0 {
			h.mu.Lock()
			terminateJob.Call(h.job, 0)
			// Job termination is asynchronous. The worker exiting does not prove
			// Word has released the template files it opened.
			query := kernel.NewProc("QueryInformationJobObject")
			deadline := time.Now().Add(5 * time.Second)
			for {
				var accounting jobAccounting
				ok, _, err := query.Call(h.job, 1, uintptr(unsafe.Pointer(&accounting)), unsafe.Sizeof(accounting), 0)
				if ok == 0 {
					cleanupError = fmt.Errorf("query owned job during cleanup: %w", err)
					break
				}
				if accounting.Active == 0 {
					break
				}
				if time.Now().After(deadline) {
					cleanupError = Fail("owned_job_exit_timeout", "Owned processes did not exit before workspace cleanup", map[string]any{"active_processes": accounting.Active, "directory": h.cfg.Directory})
					break
				}
				time.Sleep(10 * time.Millisecond)
			}
			closeHandle.Call(h.job)
			h.job = 0
			h.mu.Unlock()
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
		if h.cfg.Directory != "" && cleanupError == nil {
			deadline := time.Now().Add(time.Second)
			for {
				cleanupError = os.RemoveAll(h.cfg.Directory)
				if !errors.Is(cleanupError, syscall.Errno(32)) || time.Now().After(deadline) {
					break
				}
				time.Sleep(25 * time.Millisecond)
			}
			if cleanupError != nil {
				cleanupError = Fail("workspace_cleanup_failed", "Owned workspace could not be removed after process shutdown", map[string]any{"directory": h.cfg.Directory, "cause": cleanupError.Error()})
			}
		}
	})
	return h.closeErr
}

var _ = time.Second
