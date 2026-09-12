// Package verify records observable native behavior. It contains no VBA evaluator.
package verify

import (
	"context"
	"encoding/json"
	"fmt"
	"math"
	"path/filepath"
	"regexp"
	"runtime"
	"strconv"
	"strings"
	"time"

	"wordwright.local/internal/native"
	"wordwright.local/internal/office"
	"wordwright.local/internal/project"
)

type Assertion struct {
	Path      string  `json:"path,omitempty"` // RFC 6901 JSON pointer, empty means entire result.
	Kind      string  `json:"kind"`
	Expected  any     `json:"expected,omitempty"`
	Tolerance float64 `json:"tolerance,omitempty"`
}
type Step struct {
	Name      string           `json:"name"`
	Operation native.Operation `json:"operation"`
	Assert    []Assertion      `json:"assert,omitempty"`
	TimeoutMS int              `json:"timeout_ms,omitempty"`
}
type Suite struct {
	Schema         int      `json:"schema"`
	Name           string   `json:"name"`
	Platforms      []string `json:"platforms,omitempty"`
	Steps          []Step   `json:"steps"`
	RequireCompile bool     `json:"require_compile,omitempty"`
}
type Observation struct {
	Name       string  `json:"name"`
	Passed     bool    `json:"passed"`
	Result     any     `json:"result,omitempty"`
	Error      any     `json:"error,omitempty"`
	Assertions int     `json:"assertions"`
	DurationMS float64 `json:"duration_ms"`
}
type Report struct {
	Schema       int            `json:"schema"`
	ToolVersion  string         `json:"tool_version"`
	Status       string         `json:"status"`
	Artifact     string         `json:"artifact"`
	SHA256       string         `json:"sha256"`
	SuiteSHA256  string         `json:"suite_sha256"`
	Suite        Suite          `json:"suite"`
	OS           string         `json:"os"`
	Arch         string         `json:"arch"`
	FreshProcess bool           `json:"fresh_process"`
	WordExecuted bool           `json:"word_executed"`
	MacExecuted  bool           `json:"mac_executed"`
	VBACompiled  bool           `json:"vba_compiled"`
	Assertions   int            `json:"assertions"`
	StartedUTC   string         `json:"started_utc"`
	DurationMS   float64        `json:"duration_ms"`
	Host         map[string]any `json:"host,omitempty"`
	Observations []Observation  `json:"observations"`
	Error        any            `json:"error,omitempty"`
}

func ErrorValue(e error) any {
	if e == nil {
		return nil
	}
	if f, ok := e.(*native.Fault); ok {
		return f
	}
	return map[string]any{"message": e.Error()}
}
func canonical(v any) any {
	b, _ := json.Marshal(v)
	var out any
	_ = json.Unmarshal(b, &out)
	return out
}
func Pointer(v any, path string) (any, error) {
	if path == "" {
		return v, nil
	}
	if !strings.HasPrefix(path, "/") {
		return nil, fmt.Errorf("assertion path must be a JSON pointer")
	}
	for _, s := range strings.Split(path[1:], "/") {
		s = strings.ReplaceAll(strings.ReplaceAll(s, "~1", "/"), "~0", "~")
		switch x := v.(type) {
		case map[string]any:
			var ok bool
			v, ok = x[s]
			if !ok {
				return nil, fmt.Errorf("missing result key %q", s)
			}
		case []any:
			i, e := strconv.Atoi(s)
			if e != nil || i < 0 || i >= len(x) {
				return nil, fmt.Errorf("invalid result index %q", s)
			}
			v = x[i]
		default:
			return nil, fmt.Errorf("result path crosses a scalar")
		}
	}
	return v, nil
}
func Check(v any, a Assertion) error {
	actual, e := Pointer(canonical(v), a.Path)
	if e != nil {
		return e
	}
	expected := canonical(a.Expected)
	pass := false
	switch a.Kind {
	case "equals", "not_equals":
		ab, _ := json.Marshal(actual)
		eb, _ := json.Marshal(expected)
		pass = string(ab) == string(eb)
		if a.Kind == "not_equals" {
			pass = !pass
		}
	case "contains":
		s, ok := actual.(string)
		part, pok := expected.(string)
		if !ok || !pok {
			return fmt.Errorf("contains requires two strings")
		}
		pass = strings.Contains(s, part)
	case "matches":
		s, ok := actual.(string)
		pattern, pok := expected.(string)
		if !ok || !pok || len(pattern) > 100000 {
			return fmt.Errorf("matches needs a bounded regex and a string")
		}
		r, e := regexp.Compile(pattern)
		if e != nil {
			return e
		}
		pass = r.MatchString(s)
	case "near":
		x, ok := actual.(float64)
		y, yok := expected.(float64)
		if !ok || !yok || a.Tolerance < 0 || math.IsNaN(a.Tolerance) {
			return fmt.Errorf("near requires numeric values and nonnegative tolerance")
		}
		pass = math.Abs(x-y) <= a.Tolerance
	case "greater_than":
		x, ok := actual.(float64)
		y, yok := expected.(float64)
		if !ok || !yok {
			return fmt.Errorf("greater_than requires numbers")
		}
		pass = x > y
	default:
		return fmt.Errorf("unknown assertion kind %q", a.Kind)
	}
	if !pass {
		return fmt.Errorf("assertion %s at %q failed: actual=%v expected=%v", a.Kind, a.Path, actual, expected)
	}
	return nil
}
func (s Suite) Validate() error {
	if s.Schema != 1 {
		return fmt.Errorf("suite schema must be 1")
	}
	if s.Name == "" || len(s.Steps) == 0 || len(s.Steps) > 10000 {
		return fmt.Errorf("named nonempty bounded suite required")
	}
	assertions := 0
	names := map[string]bool{}
	for _, v := range s.Steps {
		if v.Name == "" || names[v.Name] {
			return fmt.Errorf("step names must be nonempty and unique")
		}
		names[v.Name] = true
		if v.Operation.Op == "" {
			return fmt.Errorf("step %s has no operation", v.Name)
		}
		if v.TimeoutMS < 0 || v.TimeoutMS > 1800000 {
			return fmt.Errorf("step timeout outside 0..1800000ms")
		}
		assertions += len(v.Assert)
	}
	if assertions == 0 {
		return fmt.Errorf("an acceptance suite without assertions cannot pass")
	}
	return nil
}

// Run accepts an existing host for explicitly labelled warm tests. Nil creates a
// fresh native Word process. Both paths bind observations to actual input bytes.
func Run(ctx context.Context, artifact string, s Suite, existing native.Host, execute bool) (*Report, error) {
	started := time.Now()
	abs, e := filepath.Abs(artifact)
	if e != nil {
		return nil, e
	}
	b, e := project.Read(filepath.Dir(abs), filepath.Base(abs))
	if e != nil {
		return nil, e
	}
	r := &Report{Schema: 1, ToolVersion: project.Version, Status: "not_run", Artifact: abs, SHA256: office.Hash(b), SuiteSHA256: office.Hash(project.JSON(s)), Suite: s, OS: runtime.GOOS, Arch: runtime.GOARCH, FreshProcess: existing == nil, StartedUTC: started.UTC().Format(time.RFC3339Nano), Observations: []Observation{}}
	finish := func(e error) (*Report, error) {
		r.DurationMS = float64(time.Since(started).Microseconds()) / 1000
		r.Error = ErrorValue(e)
		return r, e
	}
	if e = s.Validate(); e != nil {
		return finish(e)
	}
	if len(s.Platforms) > 0 {
		ok := false
		for _, p := range s.Platforms {
			ok = ok || p == runtime.GOOS
		}
		if !ok {
			return finish(fmt.Errorf("suite does not target %s; it was not executed", runtime.GOOS))
		}
	}
	if !execute {
		return finish(fmt.Errorf("native execution must be explicitly authorized"))
	}
	p, e := office.ReadPackage(b)
	if e != nil {
		return finish(e)
	}
	if e = p.Validate(); e != nil {
		return finish(e)
	}
	h := existing
	if h == nil {
		h, e = native.Start(ctx, native.Options{Execute: true})
		if e != nil {
			return finish(e)
		}
		defer h.Close()
	}
	r.Host = h.Info()
	r.Status = "failed"
	// Copy the exact bytes once; never let a concurrent source edit change the
	// candidate between the report hash and Word's open call.
	temp, e := createCandidate(abs, b)
	if e != nil {
		return finish(e)
	}
	defer removeCandidate(temp)
	for _, step := range s.Steps {
		t := time.Now()
		op := step.Operation
		op = expand(op, map[string]string{"$artifact": temp, "$project": pProjectName(p), "$filename": filepath.Base(temp)})
		ms := step.TimeoutMS
		if ms == 0 {
			ms = 120000
		}
		sc, cancel := context.WithTimeout(ctx, time.Duration(ms)*time.Millisecond)
		value, err := h.Call(sc, op)
		cancel()
		ob := Observation{Name: step.Name, Result: value, Error: ErrorValue(err), DurationMS: float64(time.Since(t).Microseconds()) / 1000}
		if err == nil {
			r.WordExecuted = true
			r.MacExecuted = runtime.GOOS == "darwin"
			if op.Op == "compile" {
				m, _ := canonical(value).(map[string]any)
				if m["compiled"] == true || m["vba_compiled"] == true {
					r.VBACompiled = true
				}
			}
			for _, a := range step.Assert {
				if err = Check(value, a); err != nil {
					ob.Error = ErrorValue(err)
					break
				}
				ob.Assertions++
				r.Assertions++
			}
		}
		ob.Passed = err == nil
		r.Observations = append(r.Observations, ob)
		if err != nil {
			return finish(fmt.Errorf("step %q: %w", step.Name, err))
		}
	}
	if s.RequireCompile && !r.VBACompiled {
		return finish(fmt.Errorf("suite required whole-project native compilation but no verified compile observation was recorded"))
	}
	if r.Assertions == 0 || !r.WordExecuted {
		return finish(fmt.Errorf("no native assertion evidence"))
	}
	r.Status = "passed"
	return finish(nil)
}
func pProjectName(p *office.Package) string {
	if b := p.Files["word/vbaProject.bin"]; b != nil {
		if v, e := office.ReadVBA(b); e == nil {
			return v.Name
		}
	}
	return ""
}
func expand(op native.Operation, replacements map[string]string) native.Operation {
	var walk func(any) any
	walk = func(v any) any {
		switch x := v.(type) {
		case string:
			for k, s := range replacements {
				x = strings.ReplaceAll(x, k, s)
			}
			return x
		case []any:
			for i := range x {
				x[i] = walk(x[i])
			}
			return x
		case map[string]any:
			for k := range x {
				x[k] = walk(x[k])
			}
			return x
		}
		return v
	}
	b, _ := json.Marshal(op)
	var v any
	_ = json.Unmarshal(b, &v)
	b, _ = json.Marshal(walk(v))
	_ = json.Unmarshal(b, &op)
	return op
}
