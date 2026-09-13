//go:build windows && (amd64 || arm64)

package native

import (
	"bufio"
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"syscall"
	"time"
	"unsafe"

	"github.com/eliziff/WordUp/internal/office"
)

var enumDesktopWindows = user32.NewProc("EnumDesktopWindows")
var enumChildren = user32.NewProc("EnumChildWindows")
var getThreadDesktop = user32.NewProc("GetThreadDesktop")
var getThreadID = kernel.NewProc("GetCurrentThreadId")
var windowPID = user32.NewProc("GetWindowThreadProcessId")
var getClass = user32.NewProc("GetClassNameW")
var windowText = user32.NewProc("GetWindowTextW")
var windowVisible = user32.NewProc("IsWindowVisible")
var peekMessage = user32.NewProc("PeekMessageW")
var translateMessage = user32.NewProc("TranslateMessage")
var dispatchMessage = user32.NewProc("DispatchMessageW")

type message struct {
	Window         uintptr
	Message        uint32
	WParam, LParam uintptr
	Time           uint32
	X, Y           int32
	Private        uint32
}

func pump() {
	var msg message
	for {
		r, _, _ := peekMessage.Call(uintptr(unsafe.Pointer(&msg)), 0, 0, 0, 1)
		if r == 0 {
			return
		}
		translateMessage.Call(uintptr(unsafe.Pointer(&msg)))
		dispatchMessage.Call(uintptr(unsafe.Pointer(&msg)))
	}
}
func pidOf(hwnd uintptr) uint32 {
	var pid uint32
	windowPID.Call(hwnd, uintptr(unsafe.Pointer(&pid)))
	return pid
}
func classOf(hwnd uintptr) string {
	var b [256]uint16
	n, _, _ := getClass.Call(hwnd, uintptr(unsafe.Pointer(&b[0])), uintptr(len(b)))
	return syscall.UTF16ToString(b[:int(n)])
}
func textOf(hwnd uintptr) string {
	var b [8192]uint16
	n, _, _ := windowText.Call(hwnd, uintptr(unsafe.Pointer(&b[0])), uintptr(len(b)))
	return syscall.UTF16ToString(b[:int(n)])
}
func currentDesktop() uintptr {
	tid, _, _ := getThreadID.Call()
	d, _, _ := getThreadDesktop.Call(tid)
	return d
}

type windowEnum struct {
	pid     uint32
	found   uintptr
	windows []uintptr
}

var collectWindowsCallback = syscall.NewCallback(func(h, l uintptr) uintptr {
	q := (*windowEnum)(unsafe.Pointer(l))
	if pidOf(h) == q.pid {
		q.windows = append(q.windows, h)
	}
	return 1
})
var findDocumentCallback = syscall.NewCallback(func(h, l uintptr) uintptr {
	q := (*windowEnum)(unsafe.Pointer(l))
	if pidOf(h) == q.pid && classOf(h) == "_WwG" {
		q.found = h
		return 0
	}
	return 1
})

func windowsFor(pid uint32) []uintptr {
	q := windowEnum{pid: pid}
	enumDesktopWindows.Call(currentDesktop(), collectWindowsCallback, uintptr(unsafe.Pointer(&q)))
	runtime.KeepAlive(&q)
	return q.windows
}
func documentWindow(pid uint32) uintptr {
	q := windowEnum{pid: pid}
	for _, h := range windowsFor(pid) {
		enumChildren.Call(h, findDocumentCallback, uintptr(unsafe.Pointer(&q)))
		runtime.KeepAlive(&q)
		if q.found != 0 {
			return q.found
		}
	}
	return 0
}

type wordHost struct {
	cfg          hostConfig
	process      childProcess
	app          dispatch
	objects      map[string]dispatch
	staged       map[string]string
	execute      bool
	uiWindows    sync.Map
	evalPrograms map[string]*evalProgram
}

func connectWord(cfg hostConfig) (*wordHost, error) {
	null, e := os.OpenFile("NUL", os.O_RDWR, 0600)
	if e != nil {
		return nil, e
	}
	defer null.Close()
	p, e := spawnOnDesktop(cfg.WordPath, []string{"/a", filepath.Join(cfg.Directory, "seed.docx")}, "WinSta0\\"+cfg.Desktop, null, null, null, 0)
	if e != nil {
		return nil, e
	}
	h := &wordHost{cfg: cfg, process: p, objects: map[string]dispatch{}, staged: map[string]string{}, execute: cfg.Execute}
	good := false
	defer func() {
		if !good {
			terminateProcess.Call(uintptr(p.Process), 1)
			closeHandle.Call(uintptr(p.Process))
		}
	}()
	duration := time.Duration(cfg.StartupTimeoutMS) * time.Millisecond
	if duration <= 0 {
		duration = 60 * time.Second
	}
	end := time.Now().Add(duration)
	var last error
	for time.Now().Before(end) {
		pump()
		state, _, _ := waitSingle.Call(uintptr(p.Process), 0)
		if state == 0 {
			return nil, Fail("word_process_exited", "The explicitly launched Word process exited before a document window appeared; no other Word instance was attached", nil)
		}
		if hwnd := documentWindow(p.PID); hwnd != 0 {
			win, e := fromNativeWindow(hwnd, 0xfffffff0)
			if e == nil {
				// Word exposes Hwnd on Window, not Application. Verify the native
				// window before extracting its Application object.
				hv, windowErr := win.get("Hwnd")
				if windowErr != nil {
					win.release()
					last = windowErr
					continue
				}
				own := pidOf(uintptr(uint32(hv.Value)))
				hv.clear()
				if own != p.PID {
					win.release()
					return nil, Fail("word_ownership_mismatch", "Refusing to automate a Word window outside the app-owned process", nil)
				}
				v, e := win.get("Application")
				win.release()
				if e == nil {
					app, e := v.object()
					v.clear()
					if e == nil {
						if e = app.put("AutomationSecurity", 3); e != nil {
							app.release()
							return nil, e
						}
						h.app = app
						h.objects["app"] = app
						_ = app.put("DisplayAlerts", 0)
						_ = app.put("Visible", true)
						if cfg.Visible {
							seedValue, err := app.get("ActiveDocument")
							if err != nil {
								return nil, err
							}
							seed, err := seedValue.object()
							seedValue.clear()
							if err != nil {
								return nil, err
							}
							closed, err := seed.call("Close", 0)
							closed.clear()
							seed.release()
							if err != nil {
								return nil, err
							}
						}
						good = true
						return h, nil
					}
					last = e
				}
				if e != nil {
					last = e
				}
			}
			if e != nil {
				last = e
			}
		}
		time.Sleep(40 * time.Millisecond)
	}
	return nil, Fail("word_startup_timeout", "The app-owned Word process did not expose its native document window", map[string]any{"last_error": fmt.Sprint(last), "windows": windowInventory(p.PID)})
}
func (h *wordHost) close() {
	for name, d := range h.objects {
		if name != "app" {
			d.release()
		}
	}
	if h.app.ptr != 0 {
		v, _ := h.app.call("Quit", 0)
		v.clear()
		h.app.release()
	}
	if h.process.Process != 0 {
		closeHandle.Call(uintptr(h.process.Process))
	}
}
func (h *wordHost) store(name string, d dispatch) error {
	if name == "" || len(name) > 128 || name == "app" {
		d.release()
		return fmt.Errorf("a nonempty object handle other than app is required")
	}
	if len(h.objects) >= 4096 {
		d.release()
		return fmt.Errorf("object handle budget; release unused objects")
	}
	if old, ok := h.objects[name]; ok {
		old.release()
	}
	h.objects[name] = d
	return nil
}
func (h *wordHost) result(v *variant, name string) (any, error) {
	defer v.clear()
	if v.VT == 9 {
		d, e := v.object()
		if e != nil {
			return nil, e
		}
		if e = h.store(name, d); e != nil {
			return nil, e
		}
		return map[string]any{"object": name}, nil
	}
	return v.value(0)
}
func (h *wordHost) object(name string) (dispatch, error) {
	if name == "" {
		name = "app"
	}
	d, ok := h.objects[name]
	if !ok {
		return dispatch{}, fmt.Errorf("unknown object handle %q", name)
	}
	return d, nil
}
func (h *wordHost) stage(source string) (string, error) {
	source, e := filepath.Abs(source)
	if e != nil {
		return "", e
	}
	info, e := os.Stat(source)
	if e != nil {
		return "", e
	}
	if !info.Mode().IsRegular() || info.Size() > office.Limit {
		return "", fmt.Errorf("native input is not a bounded file")
	}
	b, e := os.ReadFile(source)
	if e != nil {
		return "", e
	}
	p, e := office.ReadPackage(b)
	if e != nil {
		return "", e
	}
	if e = p.Validate(); e != nil {
		return "", e
	}
	base := filepath.Base(source)
	target := filepath.Join(h.cfg.Directory, base)
	hash := office.Hash(b)
	key := strings.ToLower(target)
	if prior, ok := h.staged[key]; ok {
		if prior != hash {
			return "", Fail("staged_filename_busy", "Close the previous document, then use unload before opening a new revision with this filename", map[string]any{"path": target})
		}
		current, err := os.ReadFile(target)
		if err != nil || office.Hash(current) != hash {
			return "", Fail("staged_input_changed", "The staged artifact changed or became unreadable during native execution; use a fresh host for the original candidate", map[string]any{"path": target})
		}
		// Reuse the exact existing bytes. Rewriting a template already loaded
		// in Word would conflict with its file lock, even when bytes match.
		return target, nil
	}
	if !strings.EqualFold(source, target) {
		if e = os.WriteFile(target, b, 0600); e != nil {
			return "", e
		}
	}
	h.staged[key] = hash
	return target, nil
}
func (h *wordHost) operation(op Operation) (any, error) {
	switch op.Op {
	case "profile":
		return profileOperations(op.Steps, h.operation)
	case "hello":
		if op.Value != h.cfg.Token {
			return nil, Fail("invalid_handshake", "worker handshake token mismatch", nil)
		}
		version, _ := h.app.get("Version")
		defer version.clear()
		v, _ := version.value(0)
		return map[string]any{"pid": h.process.PID, "word_version": v, "desktop": h.cfg.Desktop, "visible_desktop_switched": false, "owns_word_process": true, "user_word_attached": false, "runtime": "Microsoft Word", "macro_execution_authorized": h.execute, "registry_security_settings_modified": false, "automation_open_mode": "ForceDisable for inspection; Low only for explicitly authorized staged input", "directory": h.cfg.Directory, "private_desktop_is_security_sandbox": false}, nil
	case "release":
		if op.Target == "app" {
			return nil, fmt.Errorf("cannot release application handle")
		}
		if d, ok := h.objects[op.Target]; ok {
			d.release()
			delete(h.objects, op.Target)
		}
		return true, nil
	case "get", "invoke", "put", "putref":
		if op.Op != "get" && !h.execute {
			return nil, Fail("execution_not_authorized", "Raw Word mutation/invocation requires the session's explicit execute capability", nil)
		}
		d, e := h.object(op.Target)
		if e != nil {
			return nil, e
		}
		flag := uint16(2)
		args := op.Args
		switch op.Op {
		case "invoke":
			flag = 1
		case "put":
			flag = 4
			args = append(append([]any(nil), args...), op.Value)
		case "putref":
			flag = 8
			args = append(append([]any(nil), args...), op.Value)
		}
		v, e := d.invoke(op.Member, flag, args, op.Named, h.objects)
		if e != nil {
			return nil, e
		}
		return h.result(&v, op.As)
	case "open", "new", "addin":
		disableMacros := false
		if value, present := op.Named["disable_macros"]; present {
			var valid bool
			disableMacros, valid = value.(bool)
			if !valid || op.Op != "open" {
				return nil, fmt.Errorf("disable_macros is a boolean option for open only")
			}
		}
		var target string
		var e error
		if op.File != "" {
			target, e = h.stage(op.File)
			if e != nil {
				return nil, e
			}
		} else if op.Op != "new" {
			return nil, fmt.Errorf("file required for %s", op.Op)
		}
		if op.Op != "open" && !h.execute {
			return nil, Fail("execution_not_authorized", "Template/add-in activation requires explicit execute capability", nil)
		}
		security := 3
		if h.execute && !disableMacros {
			security = 1
		}
		if e = h.app.put("AutomationSecurity", security); e != nil {
			return nil, e
		}
		defer h.app.put("AutomationSecurity", 3)
		collection := "Documents"
		if op.Op == "addin" {
			collection = "AddIns"
		}
		cv, e := h.app.get(collection)
		if e != nil {
			return nil, e
		}
		coll, e := cv.object()
		cv.clear()
		if e != nil {
			return nil, e
		}
		defer coll.release()
		member := "Open"
		named := map[string]any{"FileName": target, "ReadOnly": false, "AddToRecentFiles": false, "Visible": true, "OpenAndRepair": false}
		if op.Op == "new" {
			member = "Add"
			named = map[string]any{"Visible": true}
			if target != "" {
				named["Template"] = target
			}
		} else if op.Op == "addin" {
			member = "Add"
			named = map[string]any{"FileName": target, "Install": true}
		}
		v, e := coll.invoke(member, 1, nil, named, h.objects)
		if e != nil {
			return nil, e
		}
		name := op.As
		if name == "" {
			name = "document"
		}
		r, e := h.result(&v, name)
		if e != nil {
			return nil, e
		}
		opened := map[string]any{"handle": r, "staged_path": target, "source_sha256": h.staged[strings.ToLower(target)], "macro_execution_authorized": h.execute, "open_and_repair": false}
		opened["macros_disabled_on_open"] = security == 3
		if op.Op != "addin" {
			if mode, modeErr := scalarNumber(h.objects[name], "CompatibilityMode"); modeErr == nil {
				opened["compatibility_mode"] = int(mode)
			} else {
				opened["compatibility_mode_error"] = fault(modeErr)
			}
			// Explicitly select the staged document. A private desktop need not
			// activate a newly opened window as the interactive desktop would.
			d := h.objects[name]
			activated, err := d.call("Activate")
			activated.clear()
			if err != nil {
				return nil, err
			}
			if win, err := objectProperty(d, "ActiveWindow"); err == nil {
				if hwnd, err := scalarNumber(win, "Hwnd"); err == nil {
					h.uiWindows.Store(name, uint64(hwnd))
				}
				win.release()
			}
		}
		return opened, nil
	case "unload":
		// Closing is explicit; never close a user's process/document. All objects
		// available here belong to the separately owned application instance.
		if op.Target == "" || op.Target == "app" {
			return nil, fmt.Errorf("document handle required")
		}
		d, e := h.object(op.Target)
		if e != nil {
			return nil, e
		}
		// Bind cleanup to the actual document, not an optional caller-supplied
		// filename. Otherwise a saved working copy stays marked as immutable
		// staged input even after it has been closed.
		fullName, e := d.get("FullName")
		if e != nil {
			return nil, e
		}
		name, nameErr := fullName.value(0)
		fullName.clear()
		if nameErr != nil {
			return nil, nameErr
		}
		v, e := d.call("Close", 0)
		v.clear()
		if e != nil {
			return nil, e
		}
		d.release()
		delete(h.objects, op.Target)
		h.uiWindows.Delete(op.Target)
		if target, ok := name.(string); ok && strings.EqualFold(filepath.Dir(target), h.cfg.Directory) {
			delete(h.staged, strings.ToLower(target))
		}
		return true, nil
	case "run":
		if !h.execute {
			return nil, Fail("execution_not_authorized", "Macro execution was not authorized for this session", nil)
		}
		if op.Macro == "" {
			return nil, fmt.Errorf("macro required")
		}
		v, e := h.app.invoke("Run", 1, append([]any{op.Macro}, op.Args...), nil, h.objects)
		if e != nil {
			f := fault(e)
			if details, ok := f.Details.(map[string]any); ok {
				details["macro"] = op.Macro
			}
			return nil, f
		}
		return h.result(&v, op.As)
	case "eval":
		return h.evaluate(op)
	case "vba.inspect", "vba.reset":
		return inspectVBA(h.app, h.execute, op.Op == "vba.reset")
	case "xml.snapshot":
		return h.xmlSnapshot(op)
	case "toc.inspect":
		return h.inspectTOC(op)
	case "eval.prepare":
		return h.prepareEvaluation(op)
	case "eval.execute":
		return h.executeEvaluation(op)
	case "eval.release":
		return h.releaseEvaluation(op.Target)
	case "context_menu":
		d, e := h.object(op.Target)
		if e != nil {
			return nil, e
		}
		w, e := objectProperty(d, "ActiveWindow")
		if e != nil {
			return nil, e
		}
		defer w.release()
		hwnd, e := scalarNumber(w, "Hwnd")
		if e != nil {
			return nil, e
		}
		return uiOperation(h.process.PID, h.cfg.Directory, h.execute, Operation{Op: "ui.context_menu", HWND: uint64(hwnd)})
	case "render":
		return h.render(op)
	case "compile":
		return h.compile(op)
	case "batch":
		if len(op.Steps) > 10000 {
			return nil, fmt.Errorf("batch step budget")
		}
		out := make([]any, 0, len(op.Steps))
		for i, step := range op.Steps {
			if step.Op == "batch" {
				return nil, fmt.Errorf("nested batches are not accepted")
			}
			r, e := h.operation(step)
			if e != nil {
				return nil, Fail("batch_step_failed", fmt.Sprintf("step %d: %v", i, e), map[string]any{"step": i, "completed": out, "cause": fault(e)})
			}
			out = append(out, r)
		}
		return out, nil
	default:
		return nil, Fail("unknown_operation", "Unknown native operation "+op.Op, nil)
	}
}
func (h *wordHost) compile(op Operation) (any, error) {
	if !h.execute {
		return nil, Fail("execution_not_authorized", "Compilation requires execute authorization", nil)
	}
	doc, e := h.object(op.Target)
	if e != nil {
		return nil, e
	}
	pv, e := doc.get("VBProject")
	if e != nil {
		// Menu automation does not require reading/writing a VBProject. The
		// target caption still has to match before any compile is attempted.
		return h.compileMenu(op)
	}
	project, e := pv.object()
	pv.clear()
	if e != nil {
		return nil, e
	}
	defer project.release()
	vv, e := h.app.get("VBE")
	if e != nil {
		return nil, e
	}
	vbe, e := vv.object()
	vv.clear()
	if e != nil {
		return nil, e
	}
	defer vbe.release()
	v, e := vbe.invoke("ActiveVBProject", 8, []any{project}, nil, nil)
	v.clear()
	if e != nil {
		return nil, e
	}
	bv, e := vbe.get("CommandBars")
	if e != nil {
		return nil, e
	}
	bars, e := bv.object()
	bv.clear()
	if e != nil {
		return nil, e
	}
	defer bars.release()
	cv, e := bars.invoke("FindControl", 1, nil, map[string]any{"ID": 578}, nil)
	if e != nil {
		return nil, e
	}
	control, e := cv.object()
	cv.clear()
	if e != nil {
		return nil, e
	}
	defer control.release()
	enabled, e := control.get("Enabled")
	if e != nil {
		return nil, e
	}
	yes, e := enabled.value(0)
	enabled.clear()
	if e != nil {
		return nil, e
	}
	if yes != true {
		return map[string]any{"vba_compiled": false, "reason": "Compile command is disabled; a precompiled state is not reclassified as a witnessed compiler pass"}, nil
	}
	v, e = control.call("Execute")
	v.clear()
	if e != nil {
		return nil, e
	}
	enabled, e = control.get("Enabled")
	if e != nil {
		return nil, e
	}
	after, _ := enabled.value(0)
	enabled.clear()
	if after != false {
		return nil, Fail("compile_not_confirmed", "Compile did not reach a disabled clean state; inspect the native VBE selection and owned UI diagnostics", compilerSelection(vbe))
	}
	return map[string]any{"vba_compiled": true, "verification": "target project selected; native Compile command executed; command became disabled"}, nil
}

type asyncTask struct {
	Status     string  `json:"status"`
	Result     any     `json:"result,omitempty"`
	Error      *Fault  `json:"error,omitempty"`
	DurationMS float64 `json:"duration_ms,omitempty"`
}
type queued struct {
	Request
	Task string
}

func HostMain(args []string) error {
	defer StopOfficeTools()
	if len(args) != 1 {
		return fmt.Errorf("private host expects exactly its generated configuration path")
	}
	b, e := os.ReadFile(args[0])
	if e != nil {
		return e
	}
	var cfg hostConfig
	if e = json.Unmarshal(b, &cfg); e != nil {
		return e
	}
	if len(cfg.Token) != 32 || (!strings.HasPrefix(cfg.Desktop, "WordUp-") && !(cfg.Visible && cfg.Desktop == "Default")) || filepath.Clean(args[0]) != filepath.Join(cfg.Directory, "host.json") {
		return fmt.Errorf("invalid private-host configuration")
	}
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	hr, _, _ := coInit.Call(0, 2)
	if failed(hr) {
		return fmt.Errorf("CoInitializeEx STA: 0x%08X", uint32(hr))
	}
	defer coUninit.Call()
	// Verify that a manually invoked private-host command cannot silently run
	// its windows on the caller's visible desktop.
	if e = verifyDesktop(cfg.Desktop); e != nil {
		return e
	}
	h, e := connectWord(cfg)
	if e != nil {
		// Startup is the first request. Preserve the full fault, including owned
		// window diagnostics, instead of reducing it to a stderr string and EOF.
		_ = json.NewEncoder(os.Stdout).Encode(Response{ID: 1, Error: fault(e)})
		return e
	}
	defer h.close()
	var output sync.Mutex
	emit := func(r Response) { output.Lock(); defer output.Unlock(); _ = json.NewEncoder(os.Stdout).Encode(r) }
	work := make(chan queued, 128)
	ui := make(chan queued, 64)
	uiAsync := make(chan queued, 64)
	ended := make(chan struct{})
	var tasksMu sync.Mutex
	tasks := map[string]*asyncTask{}
	workTasks := map[string]string{}
	perform := func(q queued, fn func(Operation) (any, error)) {
		start := time.Now()
		if q.Task != "" {
			tasksMu.Lock()
			tasks[q.Task].Status = "running"
			tasksMu.Unlock()
		}
		r, e := fn(q.Operation)
		ms := float64(time.Since(start).Microseconds()) / 1000
		if q.Task != "" {
			tasksMu.Lock()
			completed := &asyncTask{Status: "completed", Result: r, Error: fault(e), DurationMS: ms}
			tasks[q.Task] = completed
			delete(workTasks, q.Task)
			tasksMu.Unlock()
			emit(Response{Task: q.Task, Result: completed})
		} else {
			emit(Response{ID: q.ID, Result: r, Error: fault(e), DurationMS: ms})
		}
	}
	inspectionStream, inspectionErr := marshalDispatch(h.app)
	uiWorker := func(queue <-chan queued, stream uintptr) {
		runtime.LockOSThread()
		defer runtime.UnlockOSThread()
		hr, _, _ := coInit.Call(0, 2)
		if failed(hr) {
			if stream != 0 {
				ole32.NewProc("CoReleaseMarshalData").Call(stream)
				(dispatch{stream}).release()
			}
			for q := range queue {
				perform(q, func(Operation) (any, error) { return nil, fmt.Errorf("UI CoInitializeEx failed") })
			}
			return
		}
		defer coUninit.Call()
		var inspector dispatch
		inspectErr := inspectionErr
		if stream != 0 {
			inspector, inspectErr = unmarshalDispatch(stream)
			defer inspector.release()
		}
		uiTicker := time.NewTicker(20 * time.Millisecond)
		defer uiTicker.Stop()
		for {
			select {
			case q := <-queue:
				perform(q, func(op Operation) (any, error) {
					if op.Op == "ui.vba.inspect" || op.Op == "ui.vba.reset" || op.Op == "ui.vba.stack" {
						if inspectErr != nil {
							return nil, inspectErr
						}
						if op.Op == "ui.vba.stack" {
							return vbaStack(inspector, h.process.PID, cfg.Directory, h.execute)
						}
						return inspectVBA(inspector, h.execute, op.Op == "ui.vba.reset")
					}
					if op.Target != "" {
						value, ok := h.uiWindows.Load(op.Target)
						if !ok {
							return nil, fmt.Errorf("unknown document window target %s", op.Target)
						}
						op.HWND = value.(uint64)
					}
					return uiOperation(h.process.PID, cfg.Directory, h.execute, op)
				})
			case <-ended:
				return
			case <-uiTicker.C:
				pump()
			}
		}
	}
	go uiWorker(ui, inspectionStream)
	asyncUIStarted := false

	go func() {
		defer close(ended)
		s := bufio.NewScanner(os.Stdin)
		s.Buffer(make([]byte, 65536), 16<<20)
		for s.Scan() {
			var q Request
			if e = json.Unmarshal(s.Bytes(), &q); e != nil {
				emit(Response{Error: fault(e)})
				continue
			}
			if q.Operation.Op == "forget" {
				key, ok := q.Operation.Value.(string)
				tasksMu.Lock()
				task := tasks[key]
				if ok && task != nil && task.Status == "completed" {
					delete(tasks, key)
				}
				tasksMu.Unlock()
				if !ok || task == nil || task.Status != "completed" {
					emit(Response{ID: q.ID, Error: fault(fmt.Errorf("forget requires a completed task token"))})
				} else {
					emit(Response{ID: q.ID, Result: map[string]any{"forgotten": key}})
				}
				continue
			}
			if q.Operation.Op == "poll" {
				key, ok := q.Operation.Value.(string)
				tasksMu.Lock()
				task := tasks[key]
				var copyTask asyncTask
				if task != nil {
					copyTask = *task
				}
				tasksMu.Unlock()
				if !ok || task == nil {
					emit(Response{ID: q.ID, Error: fault(fmt.Errorf("unknown task"))})
				} else {
					emit(Response{ID: q.ID, Result: copyTask})
				}
				continue
			}
			item := queued{Request: q}
			if q.Operation.Op != "begin" && !strings.HasPrefix(q.Operation.Op, "ui.") {
				tasksMu.Lock()
				busy := map[string]string{}
				for key, operation := range workTasks {
					busy[key] = operation
				}
				tasksMu.Unlock()
				if len(busy) != 0 {
					emit(Response{ID: q.ID, Error: fault(Fail("word_task_pending", "Word's object-model lane has unfinished asynchronous work. Use UI operations to handle any owned dialog, then poll or wait for task completion before this call.", map[string]any{"tasks": busy, "requested_operation": q.Operation.Op, "request_executed": false}))})
					continue
				}
			}
			if q.Operation.Op == "begin" {
				if len(q.Operation.Steps) != 1 {
					emit(Response{ID: q.ID, Error: fault(fmt.Errorf("begin requires exactly one operation in steps"))})
					continue
				}
				tasksMu.Lock()
				if len(tasks) >= 256 {
					tasksMu.Unlock()
					emit(Response{ID: q.ID, Error: fault(fmt.Errorf("asynchronous task budget reached"))})
					continue
				}
				key := randomID()
				if q.Operation.As != "" {
					key = q.Operation.As
					if len(key) > 128 || tasks[key] != nil {
						tasksMu.Unlock()
						emit(Response{ID: q.ID, Error: fault(fmt.Errorf("task name already exists or exceeds 128 characters"))})
						continue
					}
				}
				tasks[key] = &asyncTask{Status: "queued"}
				if !strings.HasPrefix(q.Operation.Steps[0].Op, "ui.") {
					workTasks[key] = q.Operation.Steps[0].Op
				}
				tasksMu.Unlock()
				item.Task = key
				item.Operation = q.Operation.Steps[0]
				emit(Response{ID: q.ID, Task: key, Result: map[string]any{"task": key}})
			}
			destination := work
			if strings.HasPrefix(item.Operation.Op, "ui.") {
				destination = ui
				if item.Task != "" && !strings.HasPrefix(item.Operation.Op, "ui.vba.") {
					// A blocked accessibility action must not occupy the lane
					// used to inspect and acknowledge its modal dialogs.
					if !asyncUIStarted {
						go uiWorker(uiAsync, 0)
						asyncUIStarted = true
					}
					destination = uiAsync
				}
			}
			select {
			case destination <- item:
			case <-ended:
				return
			}
		}
	}()
	ticker := time.NewTicker(10 * time.Millisecond)
	defer ticker.Stop()
	for {
		select {
		case q := <-work:
			perform(q, h.operation)
		case <-ticker.C:
			pump()
		case <-ended:
			return nil
		}
	}
}

var getUserObjectInfo = user32.NewProc("GetUserObjectInformationW")

func verifyDesktop(want string) error {
	var size uint32
	d := currentDesktop()
	getUserObjectInfo.Call(d, 2, 0, 0, uintptr(unsafe.Pointer(&size)))
	if size == 0 || size > 65536 {
		return fmt.Errorf("cannot inspect process desktop")
	}
	b := make([]uint16, size/2+1)
	r, _, _ := getUserObjectInfo.Call(d, 2, uintptr(unsafe.Pointer(&b[0])), uintptr(size), uintptr(unsafe.Pointer(&size)))
	if r == 0 || syscall.UTF16ToString(b) != want {
		return Fail("wrong_desktop", "Private host is not on its own desktop; refusing to launch Word", nil)
	}
	return nil
}

var _ = context.Background
