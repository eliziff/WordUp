// Package localipc keeps a warm agent session reachable from ordinary CLI calls.
// Windows uses a user-only named pipe; Unix uses a mode-0600 local socket. No
// network listener, cloud service, registry change or system service is created.
package localipc

import (
	"bufio"
	"context"
	"crypto/rand"
	"crypto/subtle"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"github.com/eliziff/WordUp/internal/agent"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"sync"
	"sync/atomic"
	"time"
)

type Metadata struct {
	Schema  int    `json:"schema"`
	Address string `json:"address"`
	Token   string `json:"token"`
	Root    string `json:"root"`
	Execute bool   `json:"execute"`
	PID     int    `json:"pid"`
	Started string `json:"started"`
}
type wire struct {
	Token     string           `json:"token"`
	Method    string           `json:"method"`
	Params    agent.Parameters `json:"params"`
	TimeoutMS int              `json:"timeout_ms,omitempty"`
}
type Reply struct {
	Result     any     `json:"result,omitempty"`
	Error      any     `json:"error,omitempty"`
	DurationMS float64 `json:"duration_ms"`
}

func metaPath(root string) string { return filepath.Join(root, ".wordwright", "session.json") }
func readMeta(root string) (Metadata, error) {
	var m Metadata
	b, e := project.Read(root, ".wordwright/session.json")
	if e != nil {
		return m, e
	}
	if e = project.ReadJSON(b, &m); e != nil {
		return m, e
	}
	abs, e := filepath.Abs(root)
	if e != nil {
		return m, e
	}
	if m.Root != abs || len(m.Token) != 64 || m.Schema != 1 || !validAddress(m.Address) {
		return m, fmt.Errorf("invalid local session descriptor")
	}
	return m, nil
}
func send(ctx context.Context, m Metadata, method string, p agent.Parameters) (Reply, error) {
	var r Reply
	c, e := dial(ctx, m.Address)
	if e != nil {
		return r, e
	}
	defer c.Close()
	done := make(chan struct{})
	go func() {
		select {
		case <-ctx.Done():
			c.Close()
		case <-done:
		}
	}()
	defer close(done)
	ms := 120000
	if d, ok := ctx.Deadline(); ok {
		ms = int(time.Until(d).Milliseconds())
	}
	if e = json.NewEncoder(c).Encode(wire{m.Token, method, p, ms}); e != nil {
		return r, e
	}
	scan := bufio.NewScanner(c)
	scan.Buffer(make([]byte, 65536), 32<<20)
	if !scan.Scan() {
		if e = scan.Err(); e == nil {
			e = io.ErrUnexpectedEOF
		}
		return r, e
	}
	if e = json.Unmarshal(scan.Bytes(), &r); e != nil {
		return r, e
	}
	return r, nil
}
func Call(ctx context.Context, root, method string, p agent.Parameters) (Reply, error) {
	m, e := readMeta(root)
	if e != nil {
		return Reply{}, e
	}
	return send(ctx, m, method, p)
}
func Start(ctx context.Context, root string, execute bool) (Metadata, error) {
	root, e := filepath.Abs(root)
	if e != nil {
		return Metadata{}, e
	}
	if old, e := readMeta(root); e == nil {
		c, cancel := context.WithTimeout(ctx, time.Second)
		r, e := send(c, old, "session.info", agent.Parameters{})
		cancel()
		if e == nil && r.Error == nil {
			if old.Execute != execute {
				return Metadata{}, fmt.Errorf("existing session has a different execution capability; stop it explicitly before restarting")
			}
			return old, nil
		}
	}
	lock := filepath.Join(root, ".wordwright", "session.starting")
	if e = os.MkdirAll(filepath.Dir(lock), 0700); e != nil {
		return Metadata{}, e
	}
	f, e := os.OpenFile(lock, os.O_CREATE|os.O_EXCL|os.O_WRONLY, 0600)
	if e != nil {
		return Metadata{}, fmt.Errorf("another start is in progress (or stale session.starting exists): %w", e)
	}
	f.Close()
	defer os.Remove(lock)
	token := make([]byte, 32)
	if _, e = rand.Read(token); e != nil {
		return Metadata{}, e
	}
	exe, e := os.Executable()
	if e != nil {
		return Metadata{}, e
	}
	args := []string{"__session", root, hex.EncodeToString(token)}
	if execute {
		args = append(args, "execute")
	}
	log, e := os.OpenFile(filepath.Join(root, ".wordwright", "session.log"), os.O_CREATE|os.O_TRUNC|os.O_WRONLY, 0600)
	if e != nil {
		return Metadata{}, e
	}
	defer log.Close()
	cmd := exec.Command(exe, args...)
	cmd.Stdin = nil
	cmd.Stdout = log
	cmd.Stderr = log
	detach(cmd)
	if e = cmd.Start(); e != nil {
		return Metadata{}, e
	}
	pid := cmd.Process.Pid
	_ = cmd.Process.Release()
	deadline := time.NewTimer(10 * time.Second)
	defer deadline.Stop()
	tick := time.NewTicker(25 * time.Millisecond)
	defer tick.Stop()
	for {
		select {
		case <-ctx.Done():
			return Metadata{}, ctx.Err()
		case <-deadline.C:
			return Metadata{}, fmt.Errorf("local session %d did not publish its endpoint; inspect .wordwright/session.log", pid)
		case <-tick.C:
			if m, e := readMeta(root); e == nil && m.Token == hex.EncodeToString(token) {
				return m, nil
			}
		}
	}
}
func Serve(ctx context.Context, root, token string, execute bool) error {
	if len(token) != 64 {
		return fmt.Errorf("invalid session capability")
	}
	l, e := listen()
	if e != nil {
		return e
	}
	defer l.Close()
	m := Metadata{1, l.Address(), token, root, execute, os.Getpid(), time.Now().UTC().Format(time.RFC3339Nano)}
	if e = project.AtomicWrite(metaPath(root), project.JSON(m)); e != nil {
		return e
	}
	defer func() {
		if current, e := readMeta(root); e == nil && current.Token == token {
			os.Remove(metaPath(root))
		}
	}()
	engine := &agent.Engine{Root: root, Execute: execute}
	defer engine.Close()
	shutdown := make(chan struct{})
	var stop sync.Once
	closeServer := func() { stop.Do(func() { close(shutdown); l.Close() }) }
	defer closeServer()
	var active atomic.Int64
	var last atomic.Int64
	last.Store(time.Now().UnixNano())
	var wg sync.WaitGroup
	go func() {
		tick := time.NewTicker(time.Minute)
		defer tick.Stop()
		for {
			select {
			case <-ctx.Done():
				closeServer()
				return
			case <-shutdown:
				return
			case <-tick.C:
				if active.Load() == 0 && time.Since(time.Unix(0, last.Load())) > 10*time.Minute {
					closeServer()
					return
				}
			}
		}
	}()
	slots := make(chan struct{}, 32)
	for {
		c, e := l.Accept()
		if e != nil {
			select {
			case <-shutdown:
				wg.Wait()
				return nil
			default:
				return e
			}
		}
		select {
		case slots <- struct{}{}:
		default:
			c.Close()
			continue
		}
		active.Add(1)
		wg.Add(1)
		go func(c io.ReadWriteCloser) {
			defer wg.Done()
			defer func() { <-slots; active.Add(-1); last.Store(time.Now().UnixNano()); c.Close() }()
			idle := time.AfterFunc(5*time.Second, func() { c.Close() })
			scan := bufio.NewScanner(c)
			scan.Buffer(make([]byte, 65536), 16<<20)
			if !scan.Scan() {
				idle.Stop()
				return
			}
			idle.Stop()
			var q wire
			if project.ReadJSON(scan.Bytes(), &q) != nil || subtle.ConstantTimeCompare([]byte(q.Token), []byte(token)) != 1 {
				_ = json.NewEncoder(c).Encode(Reply{Error: map[string]any{"code": "unauthorized"}})
				return
			}
			t := time.Now()
			var r any
			var err error
			if q.Method == "session.info" {
				r = m
			} else if q.Method == "session.stop" {
				r = map[string]any{"stopping": true}
				defer closeServer()
			} else {
				ms := q.TimeoutMS
				if ms <= 0 {
					ms = 120000
				}
				if ms > 1800000 {
					ms = 1800000
				}
				callCtx, cancel := context.WithTimeout(ctx, time.Duration(ms)*time.Millisecond)
				r, err = engine.Call(callCtx, q.Method, q.Params)
				cancel()
			}
			_ = json.NewEncoder(c).Encode(Reply{r, verify.ErrorValue(err), float64(time.Since(t).Microseconds()) / 1000})
		}(c)
	}
}
