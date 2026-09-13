// Package verify records observable native behavior. It contains no VBA evaluator.
package verify

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"regexp"
	"runtime"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"

	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

type Assertion struct {
	Compare   string  `json:"compare,omitempty"` // "constraint" compares pass criteria instead of volatile values.
	Path      string  `json:"path,omitempty"`    // RFC 6901 JSON pointer, empty means entire result.
	Kind      string  `json:"kind"`
	Expected  any     `json:"expected,omitempty"`
	Tolerance float64 `json:"tolerance,omitempty"`
}
type Step struct {
	EventuallyMS int              `json:"eventually_ms,omitempty"`
	Name         string           `json:"name"`
	Operation    native.Operation `json:"operation"`
	Assert       []Assertion      `json:"assert,omitempty"`
	TimeoutMS    int              `json:"timeout_ms,omitempty"`
}
type Suite struct {
	Inputs         map[string]string `json:"inputs,omitempty"`
	Schema         int               `json:"schema"`
	Name           string            `json:"name"`
	Platforms      []string          `json:"platforms,omitempty"`
	Steps          []Step            `json:"steps"`
	RequireCompile bool              `json:"require_compile,omitempty"`
}
type Observation struct {
	Attempts    int     `json:"attempts,omitempty"`
	Name        string  `json:"name"`
	Passed      bool    `json:"passed"`
	Result      any     `json:"result,omitempty"`
	Error       any     `json:"error,omitempty"`
	Assertions  int     `json:"assertions"`
	DurationMS  float64 `json:"duration_ms"`
	Diagnostics any     `json:"diagnostics,omitempty"`
	FailureXML  any     `json:"failure_xml,omitempty"`
}
type ReplaySource struct {
	ReportSHA256   string `json:"report_sha256"`
	ArtifactSHA256 string `json:"artifact_sha256"`
	SuiteSHA256    string `json:"suite_sha256"`
}

type Report struct {
	InputSnapshots    map[string]InputSnapshot `json:"input_snapshots,omitempty"`
	ArtifactSnapshot  string                   `json:"artifact_snapshot,omitempty"`
	SavedReport       string                   `json:"saved_report,omitempty"`
	CleanupErrors     map[string]any           `json:"cleanup_errors,omitempty"`
	ReplayOf          *ReplaySource            `json:"replay_of,omitempty"`
	EvidenceDirectory string                   `json:"evidence_directory,omitempty"`
	TimingMS          map[string]float64       `json:"timing_ms,omitempty"`
	Schema            int                      `json:"schema"`
	ToolVersion       string                   `json:"tool_version"`
	Status            string                   `json:"status"`
	Artifact          string                   `json:"artifact"`
	SHA256            string                   `json:"sha256"`
	SuiteSHA256       string                   `json:"suite_sha256"`
	Suite             Suite                    `json:"suite"`
	OS                string                   `json:"os"`
	Arch              string                   `json:"arch"`
	FreshProcess      bool                     `json:"fresh_process"`
	WordExecuted      bool                     `json:"word_executed"`
	MacExecuted       bool                     `json:"mac_executed"`
	VBACompiled       bool                     `json:"vba_compiled"`
	Assertions        int                      `json:"assertions"`
	StartedUTC        string                   `json:"started_utc"`
	DurationMS        float64                  `json:"duration_ms"`
	Host              map[string]any           `json:"host,omitempty"`
	Observations      []Observation            `json:"observations"`
	Error             any                      `json:"error,omitempty"`
}

func ErrorValue(e error) any {
	if e == nil {
		return nil
	}
	var f *native.Fault
	if errors.As(e, &f) {
		if e == f {
			return f
		}
		return &native.Fault{Code: f.Code, Message: e.Error(), Details: f.Details}
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
		for i := 0; i < len(s); i++ {
			if s[i] == '~' {
				if i+1 == len(s) || (s[i+1] != '0' && s[i+1] != '1') {
					return nil, fmt.Errorf("invalid JSON pointer escape")
				}
				i++
			}
		}
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
	if a.Kind == "absent" {
		split := strings.LastIndex(a.Path, "/")
		if split < 0 {
			return fmt.Errorf("absent requires an object-key JSON pointer")
		}
		parent, e := Pointer(canonical(v), a.Path[:split])
		if e != nil {
			return e
		}
		object, ok := parent.(map[string]any)
		if !ok {
			return fmt.Errorf("absent requires an existing parent object")
		}
		encoded := a.Path[split+1:]
		key := strings.ReplaceAll(strings.ReplaceAll(encoded, "~1", "/"), "~0", "~")
		// Validate the final pointer segment with the same decoder as reads.
		if _, e := Pointer(map[string]any{key: true}, "/"+encoded); e != nil {
			return e
		}
		if actual, exists := object[key]; exists {
			return fmt.Errorf("expected absent result key at %q; actual=%v", a.Path, diagnosticValue(actual))
		}
		return nil
	}
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
	case "greater_than", "less_than":
		x, ok := actual.(float64)
		y, yok := expected.(float64)
		if !ok || !yok {
			return fmt.Errorf("%s requires numbers", a.Kind)
		}
		pass = x > y
		if a.Kind == "less_than" {
			pass = x < y
		}
	default:
		return fmt.Errorf("unknown assertion kind %q", a.Kind)
	}
	if !pass {
		return fmt.Errorf("assertion %s at %q failed: actual=%v expected=%v", a.Kind, a.Path, diagnosticValue(actual), diagnosticValue(expected))
	}
	return nil
}
func assertionFailure(value any, a Assertion, index int, cause error) error {
	actual, pointerErr := Pointer(canonical(value), a.Path)
	details := map[string]any{"assertion_index": index, "path": a.Path, "kind": a.Kind, "expected": diagnosticValue(a.Expected), "actual": diagnosticValue(actual), "actual_available": pointerErr == nil}
	if pointerErr != nil {
		details["pointer_error"] = pointerErr.Error()
	}
	return native.Fail("assertion_failed", cause.Error(), details)
}

// Keep the complete value in Observation.Result (and expected value in Suite).
// Errors are repeated in several envelopes; never duplicate unbounded XML there.
func diagnosticValue(value any) any {
	data, err := json.Marshal(value)
	if err != nil {
		return map[string]any{"unavailable": err.Error()}
	}
	const limit = 1024
	if len(data) <= limit {
		return value
	}
	end := limit
	for !utf8.Valid(data[:end]) {
		end--
	}
	return map[string]any{"truncated": true, "json_bytes": len(data), "json_sha256": office.Hash(data), "json_prefix": string(data[:end]), "full_value": "observation.result or suite assertion.expected at the assertion path"}
}

func (a Assertion) validate() error {
	if a.Path != "" && !strings.HasPrefix(a.Path, "/") {
		return fmt.Errorf("assertion path must be a JSON pointer")
	}
	escaped := strings.ReplaceAll(strings.ReplaceAll(a.Path, "~0", ""), "~1", "")
	if strings.Contains(escaped, "~") {
		return fmt.Errorf("invalid JSON pointer escape")
	}
	switch a.Kind {
	case "equals", "not_equals":
	case "absent":
		if a.Path == "" {
			return fmt.Errorf("absent requires a nonempty object-key pointer")
		}
	case "contains", "matches":
		text, ok := a.Expected.(string)
		if !ok {
			return fmt.Errorf("%s requires a string expected value", a.Kind)
		}
		if a.Kind == "matches" {
			if len(text) > 100000 {
				return fmt.Errorf("matches regex exceeds length limit")
			}
			if _, err := regexp.Compile(text); err != nil {
				return err
			}
		}
	case "near", "greater_than", "less_than":
		n, ok := canonical(a.Expected).(float64)
		if !ok || math.IsNaN(n) || math.IsInf(n, 0) {
			return fmt.Errorf("%s requires a finite numeric expected value", a.Kind)
		}
		if a.Kind == "near" && (a.Tolerance < 0 || math.IsNaN(a.Tolerance) || math.IsInf(a.Tolerance, 0)) {
			return fmt.Errorf("near requires finite nonnegative tolerance")
		}
	default:
		return fmt.Errorf("unknown assertion kind %q", a.Kind)
	}
	return nil
}

func (s Suite) Validate() error {
	if len(s.Inputs) > 32 {
		return fmt.Errorf("suite input limit is 32")
	}
	inputNames := map[string]bool{}
	for name, path := range s.Inputs {
		if inputNames[strings.ToLower(name)] {
			return fmt.Errorf("suite input name collision: %s", name)
		}
		inputNames[strings.ToLower(name)] = true
		if !inputName.MatchString(name) || !filepath.IsAbs(path) {
			return fmt.Errorf("suite input %q requires a simple name and absolute source path", name)
		}
	}
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
		if v.EventuallyMS < 0 || v.EventuallyMS > 1800000 || (v.EventuallyMS > 0 && v.Operation.Op != "get" && v.Operation.Op != "ui.find" && v.Operation.Op != "poll") {
			return fmt.Errorf("eventually_ms requires a read-only get/ui.find/poll and must be 0..1800000; the step timeout still applies")
		}
		assertions += len(v.Assert)
		for _, a := range v.Assert {
			if err := a.validate(); err != nil {
				return fmt.Errorf("step %q assertion at %q: %w", v.Name, a.Path, err)
			}
			if a.Compare != "" && a.Compare != "constraint" {
				return fmt.Errorf("unknown assertion comparison mode %q", a.Compare)
			}
		}
	}
	if assertions == 0 {
		return fmt.Errorf("an acceptance suite without assertions cannot pass")
	}
	return nil
}

// Run accepts an existing host for explicitly labelled warm tests. Nil creates a
// fresh native Word process. Both paths bind observations to actual input bytes.
func Run(ctx context.Context, artifact string, s Suite, existing native.Host, execute bool) (*Report, error) {
	return RunWithInputs(ctx, artifact, s, existing, execute, nil)
}

func RunWithInputs(ctx context.Context, artifact string, s Suite, existing native.Host, execute bool, frozen map[string]InputSnapshot) (*Report, error) {
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
	h := existing
	r.TimingMS = map[string]float64{}
	preparedProgram := ""
	var initialCPU float64
	finish := func(e error) (*Report, error) {
		recordCleanup := func(stage string, cleanupErr error) {
			if cleanupErr == nil {
				return
			}
			if r.CleanupErrors == nil {
				r.CleanupErrors = map[string]any{}
			}
			r.CleanupErrors[stage] = ErrorValue(cleanupErr)
			if e == nil {
				e = cleanupErr
				r.Status = "failed"
			}
		}
		if h != nil {
			if preparedProgram != "" {
				cleanupStarted := time.Now()
				cleanupContext, cancel := context.WithTimeout(context.Background(), 3*time.Second)
				_, cleanupErr := h.Call(cleanupContext, native.Operation{Op: "eval.release", Target: preparedProgram})
				cancel()
				r.TimingMS["scratch_release"] = float64(time.Since(cleanupStarted).Microseconds()) / 1000
				recordCleanup("scratch_release", cleanupErr)
			}
			r.Host = h.Info()
			if total, ok := r.Host["job_cpu_ms"].(float64); ok {
				r.Host["run_cpu_ms"] = total - initialCPU
			}
			if existing == nil {
				cleanupStarted := time.Now()
				cleanupErr := h.Close()
				r.TimingMS["cleanup"] = float64(time.Since(cleanupStarted).Microseconds()) / 1000
				recordCleanup("host_close", cleanupErr)
			}
		}
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
	inputs, inputErr := readInputs(s.Inputs, frozen)
	if inputErr != nil {
		return finish(inputErr)
	}
	p, e := office.ReadPackage(b)
	if e != nil {
		return finish(e)
	}
	if e = p.Validate(); e != nil {
		return finish(e)
	}
	if h == nil {
		startupStarted := time.Now()
		h, e = native.Start(ctx, native.Options{Execute: true})
		r.TimingMS["startup"] = float64(time.Since(startupStarted).Microseconds()) / 1000
		if e != nil {
			return finish(e)
		}
	}
	r.Host = h.Info()
	if existing != nil {
		initialCPU, _ = r.Host["job_cpu_ms"].(float64)
	}
	r.Status = "failed"
	// Copy the exact bytes once; never let a concurrent source edit change the
	// candidate between the report hash and Word's open call.
	temp, e := createCandidate(abs, b)
	if e != nil {
		return finish(e)
	}
	defer removeCandidate(temp)
	output, e := createEvidenceDirectory(abs, r.SHA256)
	if e != nil {
		return finish(e)
	}
	r.EvidenceDirectory = output
	// Preserve the tested bytes independently of the mutable build output and
	// the temporary candidate Word opens. This is evidence, not a new build.
	r.ArtifactSnapshot = filepath.Join(output, "input"+filepath.Ext(abs))
	if e = project.AtomicWrite(r.ArtifactSnapshot, b); e != nil {
		return finish(e)
	}
	progress, e := os.OpenFile(filepath.Join(output, "progress.jsonl"), os.O_CREATE|os.O_EXCL|os.O_WRONLY, 0600)
	if e != nil {
		return finish(e)
	}
	defer progress.Close()
	progressEncoder := json.NewEncoder(progress)
	replacements := map[string]string{"$artifact": temp, "$project": pProjectName(p), "$filename": filepath.Base(temp), "$output": output}
	r.InputSnapshots, e = stageInputs(output, inputs, replacements)
	if e != nil {
		return finish(e)
	}
	preparedSteps := []native.Operation{}
	for index, step := range s.Steps {
		if step.Operation.Op == "eval" && step.Operation.Target == "" && step.Operation.Member == "" {
			op := expand(step.Operation, replacements)
			op.As = fmt.Sprint(index)
			preparedSteps = append(preparedSteps, op)
		}
	}
	for stepIndex, step := range s.Steps {
		t := time.Now()
		if err := progressEncoder.Encode(map[string]any{"event": "step_started", "step": stepIndex + 1, "name": step.Name, "utc": t.UTC(), "artifact_sha256": r.SHA256}); err != nil {
			return finish(fmt.Errorf("write acceptance progress: %w", err))
		}
		op := step.Operation
		op = expand(op, replacements)
		ms := step.TimeoutMS
		if ms == 0 {
			ms = 30000
		}
		if op.TimeoutMS == 0 {
			op.TimeoutMS = ms
		}
		sc, cancel := context.WithTimeout(ctx, time.Duration(ms)*time.Millisecond)
		var value any
		var err error
		if runtime.GOOS == "windows" && len(preparedSteps) > 1 && op.Op == "eval" && op.Target == "" && op.Member == "" {
			if preparedProgram == "" {
				preparing := time.Now()
				var prepared any
				prepared, err = h.Call(sc, native.Operation{Op: "eval.prepare", Steps: preparedSteps})
				r.TimingMS["scratch_prepare"] = float64(time.Since(preparing).Microseconds()) / 1000
				if err == nil {
					preparedProgram = prepared.(map[string]any)["program"].(string)
				}
			}
			op.Op = "eval.execute"
			op.Target = preparedProgram
			op.Member = fmt.Sprint(stepIndex)
		}
		if err == nil {
			value, err = observedCall(sc, h, op)
		}
		attempts := 1
		if (err == nil || retryObservation(op, err)) && step.EventuallyMS > 0 {
			until := time.Now().Add(time.Duration(step.EventuallyMS) * time.Millisecond)
			lastWaitError := ""
			for {
				if err != nil {
					waitError := string(project.JSON(ErrorValue(err)))
					if waitError != lastWaitError {
						if progressErr := progressEncoder.Encode(map[string]any{"event": "step_waiting", "step": stepIndex + 1, "error": ErrorValue(err)}); progressErr != nil {
							err = fmt.Errorf("write acceptance progress: %w", progressErr)
							break
						}
						lastWaitError = waitError
					}
				}
				passed := err == nil
				for _, a := range step.Assert {
					if Check(value, a) != nil {
						passed = false
						break
					}
				}
				if passed || time.Now().After(until) {
					break
				}
				err = nil
				pollInterval := 5 * time.Millisecond
				if op.Op == "ui.find" {
					pollInterval = 50 * time.Millisecond
				}
				select {
				case <-sc.Done():
					err = sc.Err()
				case <-time.After(pollInterval):
				}
				if err != nil {
					break
				}
				attempts++
				value, err = observedCall(sc, h, op)
				if err != nil && !retryObservation(op, err) {
					break
				}
			}
		}
		cancel()
		if err != nil {
			artifactSource(err, p)
		}
		ob := Observation{Name: step.Name, Result: value, Error: ErrorValue(err), Attempts: attempts, DurationMS: float64(time.Since(t).Microseconds()) / 1000}
		if err == nil {
			if op.Op != "xml.compare" && op.Op != "xml.query" && op.Op != "xml.verify" {
				r.WordExecuted = true
			}
			r.MacExecuted = runtime.GOOS == "darwin"
			if op.Op == "compile" {
				m, _ := canonical(value).(map[string]any)
				if m["compiled"] == true || m["vba_compiled"] == true {
					r.VBACompiled = true
				}
			}
			for assertionIndex, a := range step.Assert {
				if err = Check(value, a); err != nil {
					err = assertionFailure(value, a, assertionIndex, err)
					fault := ErrorValue(err).(*native.Fault)
					// Preserve the existing flat step-error fields for report consumers.
					details := fault.Details.(map[string]any)
					stepError := map[string]any{"code": fault.Code, "message": fault.Message}
					for key, value := range details {
						stepError[key] = value
					}
					ob.Error = stepError
					break
				}
				ob.Assertions++
				r.Assertions++
			}
		}
		ob.Passed = err == nil
		if err != nil && runtime.GOOS == "windows" {
			diagnosticStarted := time.Now()
			dc, stop := context.WithTimeout(context.Background(), 10*time.Second)
			var diagnosticErr error
			capture := native.Operation{Op: "ui.diagnostics", File: filepath.Join(output, fmt.Sprintf("failure-%02d", stepIndex+1))}
			if fault, ok := err.(*native.Fault); ok {
				details, _ := fault.Details.(map[string]any)
				if fault.Code == "ui_provider_timeout" {
					// PrintWindow can block on the same unresponsive UI thread.
					capture = native.Operation{Op: "ui.windows"}
					dump, dumpErr := h.Call(dc, native.Operation{Op: "process.dump", File: filepath.Join(output, fmt.Sprintf("failure-%02d.dmp", stepIndex+1))})
					if dumpErr != nil {
						details["process_dump_error"] = ErrorValue(dumpErr)
					} else {
						details["process_dump"] = dump
					}

				}
				if fault.Code == "ui_provider_timeout" || fault.Code == "scratch_vba_failed" || (fault.Code == "vba_runtime_error" && details["runtime_location"] != nil) {
					// Located failures and stalled providers retain screenshots and dialog evidence without
					// walking unrelated document/editor accessibility trees.
					capture.Named = map[string]any{"trees": false}
				}
			}
			if op.Op == "compile" {
				if fault, ok := err.(*native.Fault); ok {
					if details, ok := fault.Details.(map[string]any); ok && details["module"] != nil && details["line_text"] != nil {
						capture.Named = map[string]any{"window_class": "wndclass_desked_gsk", "trees": false}
					}
				}
			}
			ob.Diagnostics, diagnosticErr = h.Call(dc, capture)
			if capture.Op == "ui.windows" {
				ob.Diagnostics = map[string]any{"windows": ob.Diagnostics, "capture_mode": "window_inventory", "screenshot_unavailable": "Provider timed out; synchronous window painting may block on the same UI thread"}
			}
			if failure, ok := diagnosticErr.(*native.Fault); ok && failure.Code == "diagnostic_window_not_found" {
				capture.Named = nil
				ob.Diagnostics, diagnosticErr = h.Call(dc, capture)
			}
			if diagnosticErr != nil {
				ob.Diagnostics = map[string]any{"capture_error": ErrorValue(diagnosticErr), "partial": ob.Diagnostics}
			}
			stop()
			xc, cancelXML := context.WithTimeout(context.Background(), 2*time.Second)
			xml, xmlErr := h.Call(xc, native.Operation{Op: "xml.snapshot", Named: map[string]any{"failure_capture": true}, File: filepath.Join(output, fmt.Sprintf("failure-%02d.xml", stepIndex+1))})
			cancelXML()
			if xmlErr != nil {
				ob.FailureXML = map[string]any{"capture_error": ErrorValue(xmlErr)}
			} else {
				ob.FailureXML = xml
			}
			r.TimingMS["failure_diagnostics"] += float64(time.Since(diagnosticStarted).Microseconds()) / 1000
			if diagnosticErr == nil {
				if enriched := failureDialogs(err, ob.Diagnostics); enriched != err {
					err = enriched
					ob.Error = ErrorValue(err)
				}
			}
		}
		r.Observations = append(r.Observations, ob)
		if progressErr := progressEncoder.Encode(map[string]any{"event": "step_finished", "step": stepIndex + 1, "observation": ob}); progressErr != nil {
			return finish(fmt.Errorf("write acceptance progress: %w", progressErr))
		}
		r.TimingMS["steps"] += ob.DurationMS
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

func retryObservation(op native.Operation, err error) bool {
	var fault *native.Fault
	return errors.As(err, &fault) && ((op.Op == "ui.find" && fault.Code == "ui_selector_not_found") || (op.Op == "get" && fault.Code == "word_busy"))
}

func failureDialogs(err error, diagnostics any) error {
	var failure *native.Fault
	if !errors.As(err, &failure) {
		return err
	}
	observed, _ := canonical(diagnostics).(map[string]any)
	dialogs, _ := observed["dialogs"].([]any)
	if len(dialogs) == 0 {
		return err
	}
	details := map[string]any{}
	if existing, ok := failure.Details.(map[string]any); ok {
		for key, value := range existing {
			details[key] = value
		}
	} else if failure.Details != nil {
		details["cause_details"] = failure.Details
	}
	// These are observed owned dialogs, not an inferred root cause.
	details["observed_dialogs"] = dialogs
	return native.Fail(failure.Code, failure.Message, details)
}

func observedCall(ctx context.Context, h native.Host, op native.Operation) (any, error) {
	if op.Op == "xml.query" {
		var options struct {
			Part       string            `json:"part"`
			Namespaces map[string]string `json:"namespaces"`
			Limit      int               `json:"limit"`
		}
		encoded, err := json.Marshal(op.Named)
		if err != nil {
			return nil, err
		}
		if err := project.ReadJSON(encoded, &options); err != nil {
			return nil, err
		}
		data, err := project.Read(filepath.Dir(op.File), filepath.Base(op.File))
		if err != nil {
			return nil, err
		}
		return office.QueryXMLInput(ctx, data, options.Part, op.Member, options.Namespaces, options.Limit)
	}
	if runtime.GOOS == "windows" && op.Op == "poll" {
		value, err := h.Call(ctx, op)
		if err == nil {
			result, _ := canonical(value).(map[string]any)
			if result["status"] != "completed" {
				err = acknowledgeRuntimeError(ctx, h, fmt.Sprint(op.Value), op)
			}
		}
		return value, err
	}
	if op.Op == "xml.compare" || op.Op == "xml.verify" {
		var policy office.XMLComparePolicy
		if raw, ok := op.Named["xml_policy"]; ok {
			encoded, err := json.Marshal(raw)
			if err != nil {
				return nil, err
			}
			if err = project.ReadJSON(encoded, &policy); err != nil {
				return nil, err
			}
		}
		a, err := project.Read(filepath.Dir(op.Target), filepath.Base(op.Target))
		if err != nil {
			return nil, err
		}
		b, err := project.Read(filepath.Dir(op.File), filepath.Base(op.File))
		if err != nil {
			return nil, err
		}
		if op.Op == "xml.verify" {
			var options struct {
				Comparison string `json:"comparison"`
			}
			if len(op.Named) > 0 {
				encoded, err := json.Marshal(op.Named)
				if err != nil {
					return nil, err
				}
				if err := project.ReadJSON(encoded, &options); err != nil {
					return nil, err
				}
			}
			return office.VerifyXML(a, b, options.Comparison)
		}
		return office.CompareXML(a, b, policy)
	}
	if runtime.GOOS != "windows" || (op.Op != "open" && op.Op != "new" && op.Op != "run" && op.Op != "compile" && op.Op != "eval" && op.Op != "eval.execute" && op.Op != "profile" && op.Op != "ui.invoke") {
		return h.Call(ctx, op)
	}
	var beforeUI any
	if op.Op == "ui.invoke" {
		windows, captureErr := h.Call(ctx, native.Operation{Op: "ui.windows"})
		beforeUI = map[string]any{"windows": windows, "captured_utc": time.Now().UTC(), "phase": "before_operation"}
		if captureErr != nil {
			beforeUI.(map[string]any)["capture_error"] = ErrorValue(captureErr)
		}
	}
	// Allow the observer a short diagnostic window before the independent
	// parent watchdog terminates a stuck asynchronous operation.
	v, e := h.Call(ctx, native.Operation{Op: "begin", TimeoutMS: min(1800000, op.TimeoutMS+12000), Steps: []native.Operation{op}})
	if e != nil {
		return nil, e
	}
	task := v.(map[string]any)["task"]
	if waiter, ok := h.(interface {
		WaitTask(context.Context, string) (any, error)
	}); ok {
		if op.Op == "compile" {
			v, e = waitCompile(ctx, h, waiter.WaitTask, task.(string))
		} else if op.Op == "run" || op.Op == "ui.invoke" || op.Op == "eval" || op.Op == "eval.execute" {
			v, e = waitRuntime(ctx, h, waiter.WaitTask, task.(string), op)
		} else {
			v, e = waiter.WaitTask(ctx, task.(string))
		}
		var runtimeFault *native.Fault
		if errors.As(e, &runtimeFault) && (runtimeFault.Code == "vba_runtime_error" || runtimeFault.Code == "vba_compile_error") {
			if completion, ok := v.(map[string]any); ok && completion["status"] == "completed" {
				if details, ok := runtimeFault.Details.(map[string]any); ok {
					details["task_completion"] = completion
				}
				_, _ = h.Call(ctx, native.Operation{Op: "forget", Value: task})
			}
			return nil, e
		}
		if ctx.Err() != nil {
			return nil, native.Fail("native_step_timeout", "Native operation did not finish; inspect captured owned-window diagnostics", map[string]any{"task": task, "operation": op, "cause": ErrorValue(e), "host": h.Info(), "before_operation_ui": beforeUI})
		}
		if e != nil {
			return nil, e
		}
		_, _ = h.Call(ctx, native.Operation{Op: "forget", Value: task})
		m := v.(map[string]any)
		if m["error"] != nil {
			b, _ := json.Marshal(m["error"])
			var f native.Fault
			_ = json.Unmarshal(b, &f)
			return nil, &f
		}
		return m["result"], nil
	}
	tick := time.NewTicker(5 * time.Millisecond)
	defer tick.Stop()
	for {
		select {
		case <-ctx.Done():
			return nil, native.Fail("native_step_timeout", "Native operation did not finish; inspect captured owned-window diagnostics", map[string]any{"task": task, "operation": op, "cause": ErrorValue(ctx.Err()), "host": h.Info(), "before_operation_ui": beforeUI})
		case <-tick.C:
			v, e = h.Call(ctx, native.Operation{Op: "poll", Value: task})
			if e != nil {
				return nil, e
			}
			m := v.(map[string]any)
			if m["status"] != "completed" {
				continue
			}
			_, _ = h.Call(ctx, native.Operation{Op: "forget", Value: task})
			if m["error"] != nil {
				b, _ := json.Marshal(m["error"])
				var f native.Fault
				_ = json.Unmarshal(b, &f)
				return nil, &f
			}
			return m["result"], nil
		}
	}
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

func createEvidenceDirectory(artifact, hash string) (string, error) {
	parent := filepath.Join(filepath.Dir(artifact), "acceptance-assets", hash[:12])
	if err := os.MkdirAll(parent, 0700); err != nil {
		return "", err
	}
	// Reports must keep referring to their own screenshots/XML after later runs.
	return os.MkdirTemp(parent, "run-")
}
