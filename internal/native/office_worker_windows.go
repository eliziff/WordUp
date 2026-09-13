//go:build windows && (amd64 || arm64)

package native

import (
	"bufio"
	"context"
	"encoding/json"
	"io"
	"os/exec"
	"sync"
	"syscall"
	"time"
)

type officeWorker struct {
	cmd        *exec.Cmd
	input      io.WriteCloser
	output     *bufio.Scanner
	timer      *time.Timer
	generation uint64
}

var officeWorkers = struct {
	sync.Mutex
	all map[string]*officeWorker
}{all: map[string]*officeWorker{}}

func stopOfficeWorker(key string, w *officeWorker) {
	if w.timer != nil {
		w.timer.Stop()
	}
	w.input.Close()
	w.cmd.Process.Kill()
	w.cmd.Wait()
	delete(officeWorkers.all, key)
}
func StopOfficeTools() {
	officeWorkers.Lock()
	defer officeWorkers.Unlock()
	for key, w := range officeWorkers.all {
		stopOfficeWorker(key, w)
	}
}

func callOfficeWorker(ctx context.Context, helper string, args []string) (map[string]any, error) {
	officeWorkers.Lock()
	defer officeWorkers.Unlock()
	w := officeWorkers.all[helper]
	reused := w != nil
	if w == nil {
		cmd := exec.Command(helper, "serve")
		cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true, CreationFlags: 0x4000}
		input, e := cmd.StdinPipe()
		if e != nil {
			return nil, e
		}
		output, e := cmd.StdoutPipe()
		if e != nil {
			input.Close()
			return nil, e
		}
		if e = cmd.Start(); e != nil {
			input.Close()
			output.Close()
			return nil, e
		}
		scan := bufio.NewScanner(output)
		scan.Buffer(make([]byte, 65536), 32<<20)
		w = &officeWorker{cmd: cmd, input: input, output: scan}
		officeWorkers.all[helper] = w
	}
	if w.timer != nil {
		w.timer.Stop()
	}
	w.generation++
	generation := w.generation
	type response struct {
		value map[string]any
		err   error
	}
	done := make(chan response, 1)
	go func() {
		if e := json.NewEncoder(w.input).Encode(args); e != nil {
			done <- response{err: e}
			return
		}
		if !w.output.Scan() {
			e := w.output.Err()
			if e == nil {
				e = io.EOF
			}
			done <- response{err: e}
			return
		}
		var value map[string]any
		e := json.Unmarshal(w.output.Bytes(), &value)
		done <- response{value, e}
	}()
	select {
	case <-ctx.Done():
		stopOfficeWorker(helper, w)
		return nil, Fail("office_tools_timeout", "Owned helper exceeded its deadline", nil)
	case r := <-done:
		if r.err != nil {
			stopOfficeWorker(helper, w)
			return nil, r.err
		}
		w.timer = time.AfterFunc(30*time.Second, func() {
			officeWorkers.Lock()
			defer officeWorkers.Unlock()
			if officeWorkers.all[helper] == w && w.generation == generation {
				stopOfficeWorker(helper, w)
			}
		})
		r.value["helper_reused"] = reused
		r.value["helper_pid"] = w.cmd.Process.Pid
		r.value["idle_shutdown_ms"] = 30000
		return r.value, nil
	}
}
