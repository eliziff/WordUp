package verify

import (
	"fmt"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"path/filepath"
)

type Difference struct {
	Step      string `json:"step"`
	Path      string `json:"path"`
	Baseline  any    `json:"baseline,omitempty"`
	Candidate any    `json:"candidate,omitempty"`
	Error     string `json:"error"`
}
type StepTiming struct {
	Step        string  `json:"step"`
	BaselineMS  float64 `json:"baseline_ms"`
	CandidateMS float64 `json:"candidate_ms"`
	Ratio       float64 `json:"candidate_to_baseline_ratio,omitempty"`
}
type Comparison struct {
	XMLSnapshots       []map[string]any `json:"xml_snapshots,omitempty"`
	Status             string           `json:"status"`
	BaselineSHA256     string           `json:"baseline_artifact_sha256"`
	CandidateSHA256    string           `json:"candidate_artifact_sha256"`
	SuiteSHA256        string           `json:"suite_sha256"`
	AssertionsCompared int              `json:"assertions_compared"`
	Differences        []Difference     `json:"differences"`
	Timings            []StepTiming     `json:"timings"`
	TaskTimings        []StepTiming     `json:"task_timings,omitempty"`
	Coverage           string           `json:"coverage"`
}

func checkedReport(r *Report) error {
	if r.Status != "passed" || !r.WordExecuted {
		return fmt.Errorf("comparison requires passing native execution reports")
	}
	if r.Error != nil || len(r.CleanupErrors) != 0 {
		return fmt.Errorf("comparison requires reports without execution or cleanup errors")
	}
	if err := r.Suite.Validate(); err != nil {
		return err
	}
	if r.SuiteSHA256 != office.Hash(project.JSON(r.Suite)) || len(r.Observations) != len(r.Suite.Steps) {
		return fmt.Errorf("suite hash or observation count mismatch")
	}
	if r.Suite.RequireCompile && !r.VBACompiled {
		return fmt.Errorf("required native compilation is missing")
	}
	count := 0
	for i, step := range r.Suite.Steps {
		ob := r.Observations[i]
		if !ob.Passed || ob.Error != nil || ob.Name != step.Name || ob.Assertions != len(step.Assert) {
			return fmt.Errorf("invalid observation for %q", step.Name)
		}
		for _, a := range step.Assert {
			if err := Check(ob.Result, a); err != nil {
				return fmt.Errorf("stored observation %q: %w", step.Name, err)
			}
			count++
		}
	}
	if count != r.Assertions {
		return fmt.Errorf("assertion count mismatch")
	}
	if len(r.Suite.Inputs) > 0 {
		if len(r.InputSnapshots) != len(r.Suite.Inputs) {
			return fmt.Errorf("declared input snapshots missing; recapture baseline")
		}
		if _, err := readInputs(r.Suite.Inputs, r.InputSnapshots); err != nil {
			return err
		}
	}
	return nil
}

// Compare checks recorded native behavior, not source similarity. Only fields
// explicitly asserted by the same suite are parity contracts; volatile handles
// and unasserted diagnostics are not silently treated as behavioral evidence.
func Compare(baseline, candidate *Report) (*Comparison, error) {
	return CompareWithPolicy(baseline, candidate, office.XMLComparePolicy{})
}

func CompareWithPolicy(baseline, candidate *Report, policy office.XMLComparePolicy) (*Comparison, error) {
	r := &Comparison{Status: "failed", BaselineSHA256: baseline.SHA256, CandidateSHA256: candidate.SHA256, SuiteSHA256: baseline.SuiteSHA256, Differences: []Difference{}, Timings: []StepTiming{}, Coverage: "Compared asserted result fields and hash-verified XML snapshots from two recorded native runs of the same suite; does not execute Word or certify untested features. Timing ratios are observations, not performance guarantees."}
	if err := checkedReport(baseline); err != nil {
		return r, fmt.Errorf("baseline: %w", err)
	}
	if err := checkedReport(candidate); err != nil {
		return r, fmt.Errorf("candidate: %w", err)
	}
	if baseline.SuiteSHA256 != candidate.SuiteSHA256 || baseline.OS != candidate.OS || baseline.Arch != candidate.Arch {
		return r, fmt.Errorf("parity requires the same suite, OS and architecture")
	}
	for name := range baseline.Suite.Inputs {
		if baseline.InputSnapshots[name].SHA256 != candidate.InputSnapshots[name].SHA256 {
			return r, fmt.Errorf("parity input %s differs between runs", name)
		}
	}
	for i, step := range baseline.Suite.Steps {
		before, after := baseline.Observations[i], candidate.Observations[i]
		timing := StepTiming{Step: step.Name, BaselineMS: before.DurationMS, CandidateMS: after.DurationMS}
		if before.DurationMS > 0 {
			timing.Ratio = after.DurationMS / before.DurationMS
		}
		r.Timings = append(r.Timings, timing)
		if step.Operation.Op == "poll" {
			oldTask, _ := canonical(before.Result).(map[string]any)
			newTask, _ := canonical(after.Result).(map[string]any)
			oldMS, oldOK := oldTask["duration_ms"].(float64)
			newMS, newOK := newTask["duration_ms"].(float64)
			if oldTask["status"] == "completed" && newTask["status"] == "completed" && oldOK && newOK && oldMS >= 0 && newMS >= 0 {
				taskTiming := StepTiming{Step: step.Name, BaselineMS: oldMS, CandidateMS: newMS}
				if oldMS > 0 {
					taskTiming.Ratio = newMS / oldMS
				}
				r.TaskTimings = append(r.TaskTimings, taskTiming)
			}
		}
		if step.Operation.Op == "xml.snapshot" {
			a, err := snapshotBytes(before.Result)
			if err != nil {
				return r, fmt.Errorf("baseline XML %q: %w", step.Name, err)
			}
			b, err := snapshotBytes(after.Result)
			if err != nil {
				return r, fmt.Errorf("candidate XML %q: %w", step.Name, err)
			}
			evidence, err := office.CompareXML(a, b, policy)
			if err != nil {
				return r, fmt.Errorf("XML %q: %w", step.Name, err)
			}
			r.XMLSnapshots = append(r.XMLSnapshots, map[string]any{"step": step.Name, "comparison": evidence})
			if evidence["equal"] != true {
				r.Differences = append(r.Differences, Difference{Step: step.Name, Path: "/xml", Error: "captured XML differs under the explicit comparison policy"})
			}
		}
		for _, a := range step.Assert {
			if a.Compare == "constraint" {
				r.AssertionsCompared++
				continue
			}
			old, _ := Pointer(canonical(before.Result), a.Path)
			actual, _ := Pointer(canonical(after.Result), a.Path)
			kind := "equals"
			if a.Kind == "near" {
				kind = "near"
			}
			if err := Check(actual, Assertion{Kind: kind, Expected: old, Tolerance: a.Tolerance}); err != nil {
				r.Differences = append(r.Differences, Difference{Step: step.Name, Path: a.Path, Baseline: old, Candidate: actual, Error: err.Error()})
			}
			r.AssertionsCompared++
		}
	}
	if len(r.Differences) > 0 {
		return r, fmt.Errorf("%d behavior or XML differences", len(r.Differences))
	}
	r.Status = "passed"
	return r, nil
}

func snapshotBytes(result any) ([]byte, error) {
	value, ok := canonical(result).(map[string]any)
	if !ok {
		return nil, fmt.Errorf("missing snapshot result")
	}
	var raw []byte
	if path, ok := value["file"].(string); ok && path != "" {
		var err error
		raw, err = project.Read(filepath.Dir(path), filepath.Base(path))
		if err != nil {
			return nil, err
		}
	} else if text, ok := value["xml"].(string); ok {
		raw = []byte(text)
	} else {
		return nil, fmt.Errorf("snapshot has no file or XML")
	}
	if hash, ok := value["sha256"].(string); !ok || hash != office.Hash(raw) {
		return nil, fmt.Errorf("snapshot SHA256 mismatch")
	}
	return raw, nil
}
