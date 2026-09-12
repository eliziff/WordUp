// Package agent exposes the same operations to a terminal, JSON-lines client,
// and stdio MCP. Native Word is lazy-started and stays warm within a session.
package agent

import (
	"context"
	"encoding/base64"
	"fmt"
	"github.com/eliziff/WordUp/internal/compat"
	"github.com/eliziff/WordUp/internal/deploy"
	"github.com/eliziff/WordUp/internal/example"
	"github.com/eliziff/WordUp/internal/inspect"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/signing"
	"github.com/eliziff/WordUp/internal/verify"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"sync"
	"time"
	"unicode/utf8"
)

type Parameters struct {
	Signing        signing.Options  `json:"signing,omitempty"`
	Path           string           `json:"path,omitempty"`
	Output         string           `json:"output,omitempty"`
	Name           string           `json:"name,omitempty"`
	Text           string           `json:"text,omitempty"`
	Base64         string           `json:"base64,omitempty"`
	ExpectedSHA256 string           `json:"expected_sha256,omitempty"`
	Query          string           `json:"query,omitempty"`
	Operation      native.Operation `json:"operation,omitempty"`
	Suite          *verify.Suite    `json:"suite,omitempty"`
	TimeoutMS      int              `json:"timeout_ms,omitempty"`
	Reference      string           `json:"reference,omitempty"`
	Tolerance      int              `json:"tolerance,omitempty"`
	Fresh          bool             `json:"fresh,omitempty"`
	Limit          int              `json:"limit,omitempty"`
	NativeOptions  *native.Options  `json:"native_options,omitempty"`
}
type Engine struct {
	Root          string
	Execute       bool
	mu            sync.Mutex
	fsMu          sync.Mutex
	host          native.Host
	nativeOptions native.Options
}

func (e *Engine) Close() error {
	e.mu.Lock()
	defer e.mu.Unlock()
	if e.host != nil {
		err := e.host.Close()
		e.host = nil
		return err
	}
	return nil
}
func (e *Engine) Host(ctx context.Context) (native.Host, error) {
	e.mu.Lock()
	defer e.mu.Unlock()
	if e.host == nil {
		options := e.nativeOptions
		options.Execute = e.Execute
		h, err := native.Start(ctx, options)
		if err != nil {
			return nil, err
		}
		e.host = h
	}
	return e.host, nil
}
func (e *Engine) path(s string) (string, error) {
	if s == "" {
		return "", fmt.Errorf("path required")
	}
	if filepath.IsAbs(s) {
		return filepath.Clean(s), nil
	}
	return project.Under(e.Root, filepath.ToSlash(s))
}
func (e *Engine) Call(ctx context.Context, method string, p Parameters) (any, error) {
	if !strings.HasPrefix(method, "native.") {
		e.fsMu.Lock()
		defer e.fsMu.Unlock()
	}
	switch method {
	case "sign", "signature.verify":
		file, err := e.path(p.Path)
		if err != nil {
			return nil, err
		}
		if method == "signature.verify" {
			return signing.Verify(ctx, file, p.Signing)
		}
		if !e.Execute {
			return nil, fmt.Errorf("signing requires --execute authority")
		}
		output, err := e.path(p.Output)
		if err != nil {
			return nil, err
		}
		return signing.Sign(ctx, file, output, p.Signing)
	case "deploy":
		if !e.Execute {
			return nil, fmt.Errorf("deployment requires --execute authority")
		}
		artifact, err := e.path(p.Path)
		if err != nil {
			return nil, err
		}
		proof, err := e.path(p.Reference)
		if err != nil {
			return nil, err
		}
		target, err := e.path(p.Output)
		if err != nil {
			return nil, err
		}
		if err = e.Close(); err != nil {
			return nil, err
		}
		planFile, plan, err := deploy.Prepare(e.Root, artifact, proof, target)
		if err != nil {
			return nil, err
		}
		busy, err := deploy.WordRunning()
		if err != nil {
			return plan, err
		}
		if !busy {
			installed, err := deploy.Activate(planFile)
			return map[string]any{"plan": planFile, "activation": installed}, err
		}
		if err = deploy.Spawn(planFile); err != nil {
			return plan, err
		}
		return map[string]any{"plan": planFile, "activation": plan, "worker": "same local executable; waits for Word to exit; does not terminate Word", "resume": "Windows registers a current-user login command to resume pending activation. Other platforms retain the plan for explicit activation."}, nil
	case "activate", "restore":
		if !e.Execute {
			return nil, fmt.Errorf("activation requires --execute authority")
		}
		file, err := e.path(p.Path)
		if err != nil {
			return nil, err
		}
		if method == "restore" {
			return deploy.Restore(file)
		}
		return deploy.Activate(file)
	case "example":
		out, err := e.path(p.Output)
		if err != nil {
			return nil, err
		}
		return example.Studio(out)
	case "selftest":
		// Known source generated locally, but still requires explicit execution authority.
		if !e.Execute {
			return nil, fmt.Errorf("selftest requires --execute; no native code was run")
		}
		out := p.Output
		var err error
		if out == "" {
			out = filepath.Join(e.Root, "reports", "selftest-"+time.Now().UTC().Format("20060102T150405.000000000"))
		} else {
			out, err = e.path(out)
			if err != nil {
				return nil, err
			}
		}
		build, err := example.Studio(out)
		if err != nil {
			return nil, err
		}
		suite := example.WindowsSuite()
		if runtime.GOOS == "darwin" {
			suite = example.MacSuite()
		}
		var host native.Host
		if !p.Fresh {
			host, err = e.Host(ctx)
			if err != nil {
				return nil, err
			}
		}
		report, err := verify.Run(ctx, build.Artifact, suite, host, true)
		if report != nil {
			if save := project.Write(out, "reports/acceptance.json", project.JSON(report), ""); save != nil {
				return report, save
			}
		}
		return map[string]any{"workspace": out, "build": build, "acceptance": report}, err
	case "doctor":
		return native.Describe(), nil
	case "help":
		return map[string]any{"instructions": project.AgentInstructions, "tools": Tools(), "native_operation_help": NativeHelp}, nil
	case "new":
		out, err := e.path(p.Output)
		if err != nil {
			return nil, err
		}
		return project.New(p.Name, out)
	case "import":
		source, err := e.path(p.Path)
		if err != nil {
			return nil, err
		}
		out, err := e.path(p.Output)
		if err != nil {
			return nil, err
		}
		return project.Import(source, out)
	case "inspect":
		path, err := e.path(p.Path)
		if err != nil {
			return nil, err
		}
		return inspect.Artifact(path)
	case "reference.document":
		path, err := e.path(p.Path)
		if err != nil {
			return nil, err
		}
		return inspect.StyleReference(path)
	case "reference.image":
		path, err := e.path(p.Path)
		if err != nil {
			return nil, err
		}
		im, b, typ, err := inspect.Image(path)
		if err != nil {
			return nil, err
		}
		mime := "image/" + typ
		if typ == "jpeg" {
			mime = "image/jpeg"
		}
		return map[string]any{"type": "image", "path": path, "mimeType": mime, "width": im.Bounds().Dx(), "height": im.Bounds().Dy(), "sha256": office.Hash(b), "data": base64.StdEncoding.EncodeToString(b), "interpretation": "Native image bytes; the connected agent supplies visual reasoning. No styles were inferred by this command."}, nil
	case "image.compare":
		a, err := e.path(p.Reference)
		if err != nil {
			return nil, err
		}
		b, err := e.path(p.Path)
		if err != nil {
			return nil, err
		}
		out := ""
		if p.Output != "" {
			out, err = e.path(p.Output)
			if err != nil {
				return nil, err
			}
		}
		return inspect.Compare(a, b, out, p.Tolerance)
	case "read":
		b, err := project.Read(e.Root, p.Path)
		if err != nil {
			return nil, err
		}
		r := map[string]any{"path": p.Path, "sha256": office.Hash(b), "bytes": len(b)}
		if utf8.Valid(b) && !strings.ContainsRune(string(b), 0) {
			r["text"] = string(b)
		} else {
			r["base64"] = base64.StdEncoding.EncodeToString(b)
		}
		return r, nil
	case "write":
		b := []byte(p.Text)
		if p.Base64 != "" {
			if p.Text != "" {
				return nil, fmt.Errorf("choose text or base64")
			}
			var err error
			b, err = base64.StdEncoding.DecodeString(p.Base64)
			if err != nil {
				return nil, err
			}
		}
		if err := project.Write(e.Root, p.Path, b, p.ExpectedSHA256); err != nil {
			return nil, err
		}
		return map[string]any{"path": p.Path, "sha256": office.Hash(b), "bytes": len(b)}, nil
	case "files", "search":
		w, err := project.Open(e.Root)
		if err != nil {
			return nil, err
		}
		files, err := w.SourceFiles()
		if err != nil {
			return nil, err
		}
		names := []string{}
		for n := range files {
			names = append(names, n)
		}
		sort.Strings(names)
		out := []any{}
		limit := p.Limit
		if limit <= 0 {
			limit = 200
		}
		if limit > 10000 {
			return nil, fmt.Errorf("result limit exceeds 10000")
		}
		total := 0
		for _, n := range names {
			b := files[n]
			if method == "files" {
				total++
				if len(out) < limit {
					out = append(out, map[string]any{"path": n, "bytes": len(b), "sha256": office.Hash(b)})
				}
				continue
			}
			if p.Query == "" {
				return nil, fmt.Errorf("nonempty search query required")
			}
			if !utf8.Valid(b) {
				continue
			}
			for i, line := range strings.Split(string(b), "\n") {
				if strings.Contains(strings.ToLower(line), strings.ToLower(p.Query)) {
					total++
					if len(out) < limit {
						out = append(out, map[string]any{"path": n, "line": i + 1, "text": line})
					}
				}
			}
		}
		return map[string]any{"results": out, "total": total, "truncated": total > len(out)}, nil
	case "build", "compile", "check", "compat":
		w, err := project.Open(e.Root)
		if err != nil {
			return nil, err
		}
		if method == "check" {
			return inspect.Check(w)
		}
		if method == "compat" {
			return compat.Review(w)
		}
		out := ""
		if p.Output != "" {
			out, err = e.path(p.Output)
			if err != nil {
				return nil, err
			}
		}
		if method == "compile" {
			return e.compile(ctx, w, out, p.Signing)
		}
		return w.Build(out)
	case "native.start":
		if p.NativeOptions != nil {
			e.mu.Lock()
			if e.host != nil {
				e.mu.Unlock()
				return nil, fmt.Errorf("stop the existing native session before changing its limits")
			}
			e.nativeOptions = *p.NativeOptions
			e.mu.Unlock()
		}
		h, err := e.Host(ctx)
		if err != nil {
			return nil, err
		}
		return h.Info(), nil
	case "native.stop":
		err := e.Close()
		return map[string]any{"closed": err == nil}, err
	case "native.call":
		if p.Operation.File != "" {
			path, err := e.path(p.Operation.File)
			if err != nil {
				return nil, err
			}
			p.Operation.File = path
		}
		h, err := e.Host(ctx)
		if err != nil {
			return nil, err
		}
		ms := p.TimeoutMS
		if ms <= 0 {
			ms = 30000
		}
		if ms > 1800000 {
			return nil, fmt.Errorf("timeout exceeds 30 minute guard")
		}
		c, cancel := context.WithTimeout(ctx, time.Duration(ms)*time.Millisecond)
		defer cancel()
		if p.Operation.TimeoutMS == 0 {
			p.Operation.TimeoutMS = ms
		}
		result, err := h.Call(c, p.Operation)
		if err != nil && h.Info()["closed"] == true {
			_ = e.Close()
		}
		return result, err
	case "test":
		artifact, err := e.path(p.Path)
		if err != nil {
			return nil, err
		}
		s := p.Suite
		if s == nil {
			b, err := project.Read(e.Root, "tests/suite.json")
			if err != nil {
				return nil, err
			}
			s = &verify.Suite{}
			if err = project.ReadJSON(b, s); err != nil {
				return nil, err
			}
		}
		var h native.Host
		if !p.Fresh {
			h, err = e.Host(ctx)
			if err != nil { // still return a durable not-run report rather than lose the reason
				r, runErr := verify.Run(ctx, artifact, *s, nil, e.Execute)
				if r != nil {
					_ = project.Write(e.Root, "reports/acceptance.json", project.JSON(r), "")
				}
				return r, runErr
			}
		}
		r, err := verify.Run(ctx, artifact, *s, h, e.Execute)
		if r != nil {
			if writeErr := project.Write(e.Root, "reports/acceptance.json", project.JSON(r), ""); writeErr != nil {
				return r, fmt.Errorf("acceptance evidence could not be saved: %w", writeErr)
			}
		}
		return r, err
	default:
		return nil, fmt.Errorf("unknown tool %q; use help", method)
	}
}

const NativeHelp = `Native operations on Windows: open/new/addin (file,as); get/invoke/put/putref (target,member,args,named,value,as); run (macro,args); eval (value = VBA function body assigning Evaluate); compile (target,member = project name for the VBE menu fallback); render (target,file = output directory,value = DPI,child = optional 1-based page); release/unload; batch (steps); ui.windows/ui.tree/ui.capture/ui.invoke/ui.set_value/ui.select (owned hwnd only). Use {"object":"handle"} to pass a retained object and {"missing":true} to omit an optional COM argument. Object-returning calls require as. begin wraps one operation in steps and returns a task token immediately; poll uses value=token; forget releases a completed task token. This allows modal forms to be driven through ui.* on a second STA. Mac: dictionary/get/put/invoke/run/ae.send; event codes come from the installed Word dictionary. Never interpret check/build/compat as VBA execution. Agent code runs with the user's OS authority, not a security sandbox. Fresh test runs prove only their recorded assertions.`
