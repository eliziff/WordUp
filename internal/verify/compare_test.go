package verify

import (
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"os"
	"path/filepath"
	"testing"
)

func TestCompareRejectsNilReportsWithoutPanicking(t *testing.T) {
	for _, reports := range [][2]*Report{{nil, nil}, {nil, &Report{}}, {&Report{}, nil}} {
		comparison, err := Compare(reports[0], reports[1])
		if err == nil || comparison == nil || comparison.Status != "failed" {
			t.Fatalf("nil report was not rejected cleanly: comparison=%#v err=%v", comparison, err)
		}
	}
}

func TestCompareRejectsEditedRetainedArtifact(t *testing.T) {
	root := t.TempDir()
	artifact := filepath.Join(root, "input.dotm")
	data := []byte("exact artifact")
	if err := os.WriteFile(artifact, data, 0600); err != nil {
		t.Fatal(err)
	}
	suite := Suite{Schema: 1, Name: "artifact identity", Steps: []Step{{Name: "value", Operation: native.Operation{Op: "get", Member: "Text"}, Assert: []Assertion{{Path: "/text", Kind: "equals", Expected: "stable"}}}}}
	report := func() *Report {
		return &Report{Status: "passed", WordExecuted: true, OS: "synthetic", Arch: "synthetic", ArtifactSnapshot: artifact, SHA256: office.Hash(data), Suite: suite, SuiteSHA256: office.Hash(project.JSON(suite)), Assertions: 1, Observations: []Observation{{Name: "value", Passed: true, Assertions: 1, Result: map[string]any{"text": "stable"}}}}
	}
	if _, err := Compare(report(), report()); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(artifact, []byte("edited artifact"), 0600); err != nil {
		t.Fatal(err)
	}
	if _, err := Compare(report(), report()); err == nil {
		t.Fatal("edited retained artifact accepted")
	}
}

func TestCompareRecordedBehavior(t *testing.T) {
	// Synthetic reports exercise comparison policy, not Word execution.
	s := Suite{Schema: 1, Name: "synthetic comparison policy", Steps: []Step{{Name: "text", Operation: native.Operation{Op: "get", Member: "Text"}, Assert: []Assertion{{Path: "/text", Kind: "contains", Expected: "stable"}}}}}
	makeReport := func(text string) *Report {
		return &Report{Status: "passed", WordExecuted: true, OS: "synthetic", Arch: "synthetic", Suite: s, SuiteSHA256: office.Hash(project.JSON(s)), Assertions: 1, Observations: []Observation{{Name: "text", Passed: true, Assertions: 1, Result: map[string]any{"text": text, "volatile": 123}, DurationMS: 2}}}
	}
	a, b := makeReport("stable text"), makeReport("stable text")
	if r, err := Compare(a, b); err != nil || r.AssertionsCompared != 1 {
		t.Fatalf("same behavior: %+v %v", r, err)
	}
	b = makeReport("stable but changed")
	if r, err := Compare(a, b); err == nil || len(r.Differences) != 1 || r.Differences[0].Path != "/text" {
		t.Fatalf("lost diff: %+v %v", r, err)
	}
	b = makeReport("broken")
	if _, err := Compare(a, b); err == nil {
		t.Fatal("forged pass accepted")
	}
	b = makeReport("stable text")
	b.WordExecuted = false
	if _, err := Compare(a, b); err == nil {
		t.Fatal("offline result accepted")
	}
	b = makeReport("stable text")
	b.SuiteSHA256 = "changed"
	if _, err := Compare(a, b); err == nil {
		t.Fatal("wrong suite accepted")
	}
	b = makeReport("stable text")
	b.CleanupErrors = map[string]any{"host_close": native.Fail("workspace_cleanup_failed", "retained workspace", nil)}
	if _, err := Compare(a, b); err == nil {
		t.Fatal("cleanup failure accepted despite contradictory pass status")
	}
	b = makeReport("stable text")
	b.Error = native.Fail("native_step_timeout", "unfinished operation", nil)
	if _, err := Compare(a, b); err == nil {
		t.Fatal("execution error accepted despite contradictory pass status")
	}
}

func TestComparePerformanceConstraint(t *testing.T) {
	s := Suite{Schema: 1, Name: "timing policy", Steps: []Step{{Name: "budget", Operation: native.Operation{Op: "eval"}, Assert: []Assertion{{Kind: "less_than", Expected: 3000, Compare: "constraint"}}}}}
	report := func(ms float64) *Report {
		return &Report{Status: "passed", WordExecuted: true, Suite: s, SuiteSHA256: office.Hash(project.JSON(s)), Assertions: 1, Observations: []Observation{{Name: "budget", Passed: true, Assertions: 1, Result: ms}}}
	}
	if _, err := Compare(report(100), report(200)); err != nil {
		t.Fatal(err)
	}
	if _, err := Compare(report(100), report(4000)); err == nil {
		t.Fatal("performance regression outside target accepted")
	}
}

func TestCompareTaskTimingIsNotPollingLatency(t *testing.T) {
	s := Suite{Schema: 1, Name: "async timing", Steps: []Step{{Name: "completion", Operation: native.Operation{Op: "poll"}, Assert: []Assertion{{Path: "/status", Kind: "equals", Expected: "completed"}}}}}
	report := func(ms any) *Report {
		return &Report{Status: "passed", WordExecuted: true, Suite: s, SuiteSHA256: office.Hash(project.JSON(s)), Assertions: 1, Observations: []Observation{{Name: "completion", Passed: true, Assertions: 1, DurationMS: 2, Result: map[string]any{"status": "completed", "duration_ms": ms}}}}
	}
	r, err := Compare(report(2000), report(1400))
	if err != nil || len(r.TaskTimings) != 1 || r.TaskTimings[0].Ratio != .7 || r.Timings[0].BaselineMS != 2 {
		t.Fatalf("task duration confused with polling cost: %+v %v", r, err)
	}
	r, err = Compare(report(nil), report(1400))
	if err != nil || len(r.TaskTimings) != 0 {
		t.Fatalf("missing task timing invented: %+v %v", r, err)
	}
}

func TestCompareAutomaticallyChecksSnapshotBytes(t *testing.T) {
	suite := Suite{Schema: 1, Name: "snapshot comparison", Steps: []Step{{Name: "capture", Operation: native.Operation{Op: "xml.snapshot"}, Assert: []Assertion{{Path: "/source", Kind: "equals", Expected: "test"}}}}}
	report := func(text string) *Report {
		return &Report{Status: "passed", WordExecuted: true, Assertions: 1, Suite: suite, SuiteSHA256: office.Hash(project.JSON(suite)), Observations: []Observation{{Name: "capture", Passed: true, Assertions: 1, Result: map[string]any{"source": "test", "xml": text, "sha256": office.Hash([]byte(text))}}}}
	}
	a, b := report("<p>before</p>"), report("<p>after</p>")
	r, err := Compare(a, b)
	if err == nil || r.Status != "failed" || len(r.XMLSnapshots) != 1 {
		t.Fatal("missed unasserted snapshot change", r, err)
	}
	if r, err = Compare(a, a); err != nil || r.Status != "passed" {
		t.Fatal(r, err)
	}
	b.Observations[0].Result.(map[string]any)["sha256"] = "tampered"
	if _, err = Compare(a, b); err == nil {
		t.Fatal("accepted modified evidence")
	}
}
