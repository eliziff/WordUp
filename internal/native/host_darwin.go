//go:build darwin && (amd64 || arm64)

package native

import (
	"context"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"net/url"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/eliziff/WordUp/internal/office"
)

type scriptParameter struct {
	Name     string `xml:"name,attr" json:"name"`
	Code     string `xml:"code,attr" json:"code"`
	Type     string `xml:"type,attr" json:"type"`
	Optional string `xml:"optional,attr" json:"optional,omitempty"`
}
type scriptCommand struct {
	Name        string            `xml:"name,attr" json:"name"`
	Code        string            `xml:"code,attr" json:"code"`
	Description string            `xml:"description,attr" json:"description,omitempty"`
	Parameters  []scriptParameter `xml:"parameter" json:"parameters,omitempty"`
}
type scriptProperty struct {
	Name   string `xml:"name,attr" json:"name"`
	Code   string `xml:"code,attr" json:"code"`
	Type   string `xml:"type,attr" json:"type"`
	Access string `xml:"access,attr" json:"access,omitempty"`
}
type scriptClass struct {
	Name       string           `xml:"name,attr" json:"name"`
	Code       string           `xml:"code,attr" json:"code"`
	Properties []scriptProperty `xml:"property" json:"properties"`
}
type scriptDictionary struct {
	Commands []scriptCommand `json:"commands"`
	Classes  []scriptClass   `json:"classes"`
	Source   string          `json:"source"`
}

func normalizeTerm(s string) string {
	return strings.ToLower(strings.NewReplacer(" ", "", "_", "", "-", "").Replace(s))
}
func (d scriptDictionary) property(name string) (string, error) {
	if strings.HasPrefix(name, "#") && len(name) == 5 {
		return name[1:], nil
	}
	codes := map[string]bool{}
	for _, c := range d.Classes {
		for _, p := range c.Properties {
			if normalizeTerm(p.Name) == normalizeTerm(name) {
				codes[p.Code] = true
			}
		}
	}
	if len(codes) == 1 {
		for code := range codes {
			return code, nil
		}
	}
	return "", Fail("mac_property_lookup", fmt.Sprintf("Property %q is absent or ambiguous in this Word version's dictionary; use its explicit #four-character code", name), codes)
}
func (d scriptDictionary) command(name string) (scriptCommand, error) {
	if strings.HasPrefix(name, "#") && len(name) == 9 {
		return scriptCommand{Name: name, Code: name[1:]}, nil
	}
	for _, c := range d.Commands {
		if normalizeTerm(c.Name) == normalizeTerm(name) && len(c.Code) == 8 {
			return c, nil
		}
	}
	return scriptCommand{}, Fail("mac_command_lookup", "Command not in the installed Word dictionary: "+name, nil)
}
func wordBundle() (string, error) {
	home, _ := os.UserHomeDir()
	for _, base := range []string{"/Applications", filepath.Join(home, "Applications")} {
		p := filepath.Join(base, "Microsoft Word.app")
		if i, e := os.Stat(filepath.Join(p, "Contents", "MacOS", "Microsoft Word")); e == nil && i.Mode().IsRegular() {
			return p, nil
		}
	}
	return "", Fail("word_not_installed", "Microsoft Word.app was not found in Applications; no runtime is emulated", nil)
}
func readDictionary(bundle string) (scriptDictionary, error) {
	d := scriptDictionary{}
	root := filepath.Join(bundle, "Contents", "Resources")
	count := 0
	e := filepath.WalkDir(root, func(path string, entry os.DirEntry, e error) error {
		if e != nil {
			return e
		}
		count++
		if count > 20000 {
			return fmt.Errorf("Word resource traversal budget")
		}
		if entry.IsDir() {
			return nil
		}
		if strings.ToLower(filepath.Ext(path)) != ".sdef" {
			return nil
		}
		b, e := os.ReadFile(path)
		if e != nil {
			return e
		}
		if len(b) > 16<<20 {
			return fmt.Errorf("scripting dictionary size budget")
		}
		decoder := xml.NewDecoder(strings.NewReader(string(b)))
		for {
			tok, e := decoder.Token()
			if e == io.EOF {
				break
			}
			if e != nil {
				return e
			}
			start, ok := tok.(xml.StartElement)
			if !ok {
				continue
			}
			switch start.Name.Local {
			case "command":
				var c scriptCommand
				if e = decoder.DecodeElement(&c, &start); e != nil {
					return e
				}
				d.Commands = append(d.Commands, c)
			case "class", "class-extension":
				var c scriptClass
				if e = decoder.DecodeElement(&c, &start); e != nil {
					return e
				}
				d.Classes = append(d.Classes, c)
			}
		}
		if d.Source != "" {
			d.Source += "; "
		}
		d.Source += path
		return nil
	})
	if e != nil {
		return d, e
	}
	if len(d.Commands) == 0 {
		return d, Fail("mac_dictionary_unavailable", "No readable native scripting dictionary was found; WordUp will not guess event codes", nil)
	}
	return d, nil
}

type macHost struct {
	mu      sync.Mutex
	cmd     *exec.Cmd
	pid     int
	dir     string
	execute bool
	objects map[string]aeDesc
	dict    scriptDictionary
	info    map[string]any
	done    chan struct{}
	once    sync.Once
	staged  map[string]string
}

func Available() bool { _, e := wordBundle(); return e == nil }
func Start(ctx context.Context, opt Options) (Host, error) {
	bundle, e := wordBundle()
	if e != nil {
		return nil, e
	}
	dict, e := readDictionary(bundle)
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
	dir, e := os.MkdirTemp(base, "session-")
	if e != nil {
		return nil, e
	}
	cmd := exec.Command(filepath.Join(bundle, "Contents", "MacOS", "Microsoft Word"))
	cmd.Dir = dir
	log, e := os.Create(filepath.Join(dir, "word.log"))
	if e != nil {
		os.RemoveAll(dir)
		return nil, e
	}
	cmd.Stdout = log
	cmd.Stderr = log
	if e = cmd.Start(); e != nil {
		log.Close()
		os.RemoveAll(dir)
		return nil, e
	}
	h := &macHost{cmd: cmd, pid: cmd.Process.Pid, dir: dir, execute: opt.Execute, dict: dict, objects: map[string]aeDesc{}, info: map[string]any{}, done: make(chan struct{}), staged: map[string]string{}}
	// Events target only the exact launched PID, never a bundle ID or an existing
	// application returned by the running-application registry.
	go func() { _ = cmd.Wait(); log.Close(); close(h.done) }()
	boot, cancel := deadline(ctx, opt.StartupTimeoutMS)
	defer cancel()
	timer := time.NewTicker(100 * time.Millisecond)
	defer timer.Stop()
	null, _ := newDesc("null", nil)
	h.objects["app"] = null
	for {
		select {
		case <-h.done:
			h.Close()
			return nil, Fail("mac_word_process_exited", "The newly launched Word instance exited or delegated to another instance; no existing user instance was attached", nil)
		case <-boot.Done():
			h.Close()
			return nil, Fail("mac_word_startup_timeout", boot.Err().Error(), nil)
		case <-timer.C:
			p, e := dict.property("version")
			if e != nil {
				h.Close()
				return nil, e
			}
			spec, e := propertySpecifier(&null, p)
			if e != nil {
				h.Close()
				return nil, e
			}
			result, e := sendAppleEvent(h.pid, "core", "getd", map[string]aeDesc{"----": spec}, 1000)
			spec.close()
			if e != nil {
				if f, ok := e.(*Fault); ok && strings.Contains(f.Message, "-1743") {
					h.Close()
					return nil, e
				}
				continue
			}
			version, _ := result.value(0)
			result.close()
			h.info = map[string]any{"runtime": "Microsoft Word for Mac", "pid": h.pid, "word_version": version, "owns_launched_pid": true, "user_word_attached": false, "native_transport": "Apple Event Manager; direct system-framework calls", "macro_execution_authorized": opt.Execute, "registry_security_settings_modified": false, "private_desktop": false, "mac_execution_verified": false, "directory": dir}
			return h, nil
		}
	}
}
func (h *macHost) Info() map[string]any {
	r := map[string]any{}
	for k, v := range h.info {
		r[k] = v
	}
	return r
}
func (h *macHost) Close() error {
	var err error
	h.once.Do(func() {
		select {
		case <-h.done:
		default:
			if h.cmd != nil && h.cmd.Process != nil {
				err = h.cmd.Process.Kill()
			}
			select {
			case <-h.done:
			case <-time.After(3 * time.Second):
			}
		}
		h.mu.Lock()
		defer h.mu.Unlock()
		for _, d := range h.objects {
			d.close()
		}
		err2 := os.RemoveAll(h.dir)
		if err == nil {
			err = err2
		}
	})
	return err
}
func (h *macHost) Call(ctx context.Context, op Operation) (any, error) {
	h.mu.Lock()
	defer h.mu.Unlock()
	select {
	case <-h.done:
		return nil, Fail("mac_word_exited", "The owned Word process exited", nil)
	default:
	}
	ms := 120000
	if end, ok := ctx.Deadline(); ok {
		ms = int(time.Until(end).Milliseconds())
		if ms <= 0 {
			return nil, ctx.Err()
		}
	}
	result, e := h.operation(op, ms)
	if f, ok := e.(*Fault); ok && strings.Contains(f.Message, "-1712") {
		go h.Close()
	}
	return result, e
}
func (h *macHost) retain(result *aeDesc, name string) (any, error) {
	defer result.close()
	if result.Type == code("obj ") {
		if name == "" || name == "app" {
			return nil, fmt.Errorf("Apple-event object result needs an as handle")
		}
		if len(h.objects) > 4096 {
			return nil, fmt.Errorf("object handle budget")
		}
		copy, e := result.clone()
		if e != nil {
			return nil, e
		}
		if old, ok := h.objects[name]; ok {
			old.close()
		}
		h.objects[name] = copy
		return map[string]any{"object": name}, nil
	}
	return result.value(0)
}
func (h *macHost) event(class, id string, values map[string]any, as string, ms int) (any, error) {
	params := map[string]aeDesc{}
	defer func() {
		for _, d := range params {
			d.close()
		}
	}()
	for k, v := range values {
		d, e := descValue(v, h.objects)
		if e != nil {
			return nil, e
		}
		params[k] = d
	}
	result, e := sendAppleEvent(h.pid, class, id, params, ms)
	if e != nil {
		return nil, e
	}
	return h.retain(&result, as)
}
func (h *macHost) operation(op Operation, ms int) (any, error) {
	switch op.Op {
	case "hello":
		return h.Info(), nil
	case "dictionary":
		return h.dict, nil
	case "release":
		if op.Target == "app" {
			return nil, fmt.Errorf("cannot release app")
		}
		if d, ok := h.objects[op.Target]; ok {
			d.close()
			delete(h.objects, op.Target)
		}
		return true, nil
	case "get", "put":
		if op.Op == "put" && !h.execute {
			return nil, Fail("execution_not_authorized", "Native mutation requires execute capability", nil)
		}
		target := op.Target
		if target == "" {
			target = "app"
		}
		container, ok := h.objects[target]
		if !ok {
			return nil, fmt.Errorf("unknown target %s", target)
		}
		prop, e := h.dict.property(op.Member)
		if e != nil {
			return nil, e
		}
		spec, e := propertySpecifier(&container, prop)
		if e != nil {
			return nil, e
		}
		defer spec.close()
		params := map[string]aeDesc{"----": spec}
		id := "getd"
		if op.Op == "put" {
			id = "setd"
			value, e := descValue(op.Value, h.objects)
			if e != nil {
				return nil, e
			}
			defer value.close()
			params["data"] = value
		}
		result, e := sendAppleEvent(h.pid, "core", id, params, ms)
		if e != nil {
			return nil, e
		}
		return h.retain(&result, op.As)
	case "invoke", "run", "ae.send":
		if !h.execute {
			return nil, Fail("execution_not_authorized", "Native invocation requires execute capability", nil)
		}
		if op.Op == "ae.send" {
			if len(op.Member) != 8 {
				return nil, fmt.Errorf("ae.send member is an eight-character event class+ID")
			}
			return h.event(op.Member[:4], op.Member[4:], op.Named, op.As, ms)
		}
		name := op.Member
		named := op.Named
		if op.Op == "run" {
			name = "run VB macro"
			named = map[string]any{"macro name": op.Macro}
			if len(op.Args) > 0 {
				return nil, Fail("mac_macro_arguments", "The Word Apple-event macro command does not provide Application.Run's positional argument ABI. Compose a VBA test wrapper in the source tree, then run it; no arguments were silently dropped", nil)
			}
		}
		command, e := h.dict.command(name)
		if e != nil {
			return nil, e
		}
		values := map[string]any{}
		if len(op.Args) > 1 {
			return nil, fmt.Errorf("Apple-event commands have one direct parameter plus named parameters")
		}
		if len(op.Args) == 1 {
			values["----"] = op.Args[0]
		}
		for name, value := range named {
			key := ""
			if strings.HasPrefix(name, "#") && len(name) == 5 {
				key = name[1:]
			} else {
				for _, p := range command.Parameters {
					if normalizeTerm(p.Name) == normalizeTerm(name) {
						key = p.Code
						break
					}
				}
			}
			if len(key) != 4 {
				return nil, fmt.Errorf("unknown parameter %q for %q", name, command.Name)
			}
			values[key] = value
		}
		return h.event(command.Code[:4], command.Code[4:], values, op.As, ms)
	case "open":
		// Do not silently open untrusted macro documents where the platform's
		// scripting surface cannot force-disable them for an inspection session.
		if !h.execute {
			return nil, Fail("mac_inspection_not_authorized", "Use the offline inspector, or explicitly authorize native macro execution; no document was opened", nil)
		}
		source, e := filepath.Abs(op.File)
		if e != nil {
			return nil, e
		}
		b, e := os.ReadFile(source)
		if e != nil {
			return nil, e
		}
		p, e := office.ReadPackage(b)
		if e != nil {
			return nil, e
		}
		if e = p.Validate(); e != nil {
			return nil, e
		}
		path := filepath.Join(h.dir, filepath.Base(source))
		hash := office.Hash(b)
		if old, ok := h.staged[path]; ok && old != hash {
			return nil, fmt.Errorf("staged filename contains a different revision; close/restart the native session")
		}
		if e = os.WriteFile(path, b, 0600); e != nil {
			return nil, e
		}
		h.staged[path] = hash
		// Only the owned application's automation mode is changed, and only if
		// this Word version exposes it. Corporate policy remains authoritative.
		prior, pe := h.operation(Operation{Op: "get", Member: "automation security"}, ms)
		if pe == nil {
			_, e = h.operation(Operation{Op: "put", Member: "automation security", Value: float64(1)}, ms)
			if e != nil {
				return nil, e
			}
			defer h.operation(Operation{Op: "put", Member: "automation security", Value: prior}, ms)
		}
		fileURL := (&url.URL{Scheme: "file", Path: path}).String()
		fileDesc, e := newDesc("furl", []byte(fileURL))
		if e != nil {
			return nil, e
		}
		defer fileDesc.close()
		result, e := sendAppleEvent(h.pid, "aevt", "odoc", map[string]aeDesc{"----": fileDesc}, ms)
		if e != nil {
			return nil, e
		}
		result.close()
		name := op.As
		if name == "" {
			name = "document"
		}
		r, e := h.operation(Operation{Op: "get", Member: "active document", As: name}, ms)
		if e != nil {
			return nil, e
		}
		return map[string]any{"handle": r, "staged_path": path, "source_sha256": hash, "automation_security_exposed": pe == nil, "existing_office_policy_applies": true}, nil
	case "batch":
		if len(op.Steps) > 10000 {
			return nil, fmt.Errorf("batch budget")
		}
		out := []any{}
		for i, step := range op.Steps {
			if step.Op == "batch" {
				return nil, fmt.Errorf("nested batch")
			}
			r, e := h.operation(step, ms)
			if e != nil {
				return nil, Fail("batch_step_failed", fmt.Sprintf("step %d: %v", i, e), out)
			}
			out = append(out, r)
		}
		return out, nil
	case "compile":
		return nil, Fail("mac_full_compile_unavailable", "This native Mac backend does not claim a whole-project compiler result. Run explicit VBA acceptance procedures and assert their saved document effects", nil)
	default:
		return nil, Fail("mac_operation_unavailable", "Use dictionary/get/put/invoke/run/ae.send on macOS; native Windows UI/EMF operations are not emulated: "+op.Op, nil)
	}
}
func HostMain([]string) error {
	return fmt.Errorf("macOS uses a directly owned PID and Apple Events, not the Win32 desktop worker")
}

var _ = json.Valid
