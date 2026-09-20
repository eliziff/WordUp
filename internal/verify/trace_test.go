package verify

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

const testTrace = `{"schema":1,"complete":true,"coverage":["demo writes"],"events":[
{"sequence":1,"kind":"write","stage":"replace","document":"d1","story":"main1","operation":"w1","data":{"start":0,"end":6,"story_type":1,"story_length":7,"property":"Text","requested":"After","text":"Before"}},
{"sequence":2,"kind":"result","stage":"replace","document":"d1","story":"main1","operation":"w1","data":{"start":0,"end":5,"story_type":1,"story_length":6,"error":0,"text":"After"}}]}`

func TestTraceRejectsIncompleteEvidence(t *testing.T) {
	for _, raw := range []string{
		strings.Replace(testTrace, `"complete":true`, `"complete":false`, 1),
		strings.Replace(testTrace, `"sequence":2`, `"sequence":3`, 1),
		strings.Replace(testTrace, `"kind":"result"`, `"kind":"read"`, 1),
		strings.Replace(testTrace, `"error":0`, `"outcome":0`, 1),
		strings.Replace(testTrace, `"start":0`, `"start":null`, 1),
		strings.Replace(testTrace, `"end":6`, `"end":-1`, 1),
		strings.Replace(testTrace, `"kind":"result"`, `"kind":"unknown"`, 1),
		testTrace + `{}`, testTrace[:len(testTrace)-1],
	} {
		if _, err := decodeTrace([]byte(raw)); err == nil {
			t.Fatalf("accepted bad trace: %s", raw)
		}
	}
	if _, err := decodeTrace([]byte(testTrace)); err != nil {
		t.Fatal(err)
	}
}

func TestTraceExactComparisonAndTamper(t *testing.T) {
	dir := t.TempDir()
	put := func(name, data string) *TraceEvidence {
		t.Helper()
		if err := os.WriteFile(filepath.Join(dir, name), []byte(data), 0600); err != nil {
			t.Fatal(err)
		}
		e, err := captureTrace(dir, name)
		if err != nil {
			t.Fatal(err)
		}
		return e
	}
	a := put("a.json", testTrace)
	b := put("b.json", testTrace)
	if diff, err := compareTraces(a, b); err != nil || diff != nil {
		t.Fatal(diff, err)
	}
	for _, change := range [][2]string{{`"start":0`, `"start":1`}, {`"After"`, `"Different"`}, {`"error":0`, `"error":438`}, {`"main1"`, `"header2"`}} {
		b = put("b.json", strings.ReplaceAll(testTrace, change[0], change[1]))
		if diff, err := compareTraces(a, b); err != nil || diff == nil {
			t.Fatal("lost edit difference", diff, err)
		}
	}
	if err := os.WriteFile(a.File, []byte("tampered"), 0600); err != nil {
		t.Fatal(err)
	}
	if _, err := readTrace(a); err == nil {
		t.Fatal("accepted tampered trace")
	}
}

func TestTraceCompareFreezeRelocation(t *testing.T) {
	dir := t.TempDir()
	evidence := filepath.Join(dir, "evidence")
	if err := project.AtomicWrite(filepath.Join(evidence, "trace.json"), []byte(testTrace)); err != nil {
		t.Fatal(err)
	}
	trace, err := captureTrace(evidence, "trace.json")
	if err != nil {
		t.Fatal(err)
	}
	artifact := filepath.Join(evidence, "input.dotm")
	if err := project.AtomicWrite(artifact, []byte("synthetic")); err != nil {
		t.Fatal(err)
	}
	s := Suite{Schema: 1, Name: "bookkeeping only", SerialTrace: "trace.json", Steps: []Step{{Name: "value", Operation: native.Operation{Op: "get"}, Assert: []Assertion{{Kind: "equals", Expected: true}}}}}
	r := &Report{Schema: 1, Status: "passed", WordExecuted: true, Suite: s, SuiteSHA256: office.Hash(project.JSON(s)), SHA256: office.Hash([]byte("synthetic")), ArtifactSnapshot: artifact, EvidenceDirectory: evidence, SerialTrace: trace, Assertions: 1, Observations: []Observation{{Name: "value", Passed: true, Assertions: 1, Result: true}}}
	if err := SaveReport(dir, r); err != nil {
		t.Fatal(err)
	}
	if _, err := Compare(r, r); err != nil {
		t.Fatal(err)
	}
	bundle := filepath.Join(dir, "bundle")
	if _, err := Freeze(r.SavedReport, bundle); err != nil {
		t.Fatal(err)
	}
	if err := os.Rename(evidence, evidence+"-old"); err != nil {
		t.Fatal(err)
	}
	loaded, _, err := LoadReport(filepath.Join(bundle, "bundle.json"))
	if err != nil {
		t.Fatal(err)
	}
	if _, err := Compare(loaded, loaded); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(loaded.SerialTrace.File, []byte("bad"), 0600); err != nil {
		t.Fatal(err)
	}
	if _, _, err := LoadReport(filepath.Join(bundle, "bundle.json")); err == nil {
		t.Fatal("accepted tamper")
	}
}
