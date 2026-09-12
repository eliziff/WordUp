//go:build windows && (amd64 || arm64)

package verify_test

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/eliziff/WordUp/internal/example"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
)

func TestMain(m *testing.M) {
	if len(os.Args) > 1 && os.Args[1] == "__host" {
		if native.HostMain(os.Args[2:]) != nil {
			os.Exit(1)
		}
		return
	}
	os.Exit(m.Run())
}

func TestNativeCompletion(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	ctx := context.Background()
	h, err := native.Start(ctx, native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	w := h.(interface {
		WaitTask(context.Context, string) (any, error)
	})
	if _, err = h.Call(ctx, native.Operation{Op: "new", As: "scratch"}); err != nil {
		t.Fatal(err)
	}
	begin := func(body string, ms int) string {
		t.Helper()
		v, e := h.Call(ctx, native.Operation{Op: "begin", TimeoutMS: ms, Steps: []native.Operation{{Op: "eval", Value: body}}})
		if e != nil {
			t.Fatal(e)
		}
		return v.(map[string]any)["task"].(string)
	}
	for _, body := range []string{"Evaluate = 6 * 7", "100 Err.Raise 5"} {
		key := begin(body, 5000)
		v, e := w.WaitTask(ctx, key)
		if e != nil {
			t.Fatal(e)
		}
		b, _ := json.Marshal(v)
		if body == "Evaluate = 6 * 7" && !strings.Contains(string(b), `"result":42`) {
			t.Fatalf("missing result: %s", b)
		}
		if body != "Evaluate = 6 * 7" && (!strings.Contains(string(b), `"number":5`) || !strings.Contains(string(b), `"line":100`)) {
			t.Fatalf("missing runtime fault: %s", b)
		}
		if _, e = h.Call(ctx, native.Operation{Op: "forget", Value: key}); e != nil {
			t.Fatal(e)
		}
	}
	key := begin("Do\nLoop", 1000)
	short, cancel := context.WithTimeout(ctx, 20*time.Millisecond)
	_, err = w.WaitTask(short, key)
	cancel()
	if err != context.DeadlineExceeded || h.Info()["closed"] == true {
		t.Fatalf("caller deadline killed diagnostic lane: %v", err)
	}
	deadline, cancel := context.WithTimeout(ctx, 5*time.Second)
	defer cancel()
	_, err = w.WaitTask(deadline, key)
	if err == nil || !strings.Contains(err.Error(), "wall-clock limit") {
		t.Fatalf("independent watchdog: %v", err)
	}
}

// Opt-in: these checks run real VBA in owned Microsoft Word processes.
func TestNativeFeedback(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1 to authorize native Word tests")
	}
	root := t.TempDir()
	if output := os.Getenv("WORDUP_EVIDENCE_DIR"); output != "" {
		var e error
		root, e = os.MkdirTemp(output, "feedback-")
		if e != nil {
			t.Fatal(e)
		}
	}
	write := func(name string, v any) {
		b, _ := json.MarshalIndent(v, "", "  ")
		if e := os.WriteFile(filepath.Join(root, name), b, 0600); e != nil {
			t.Fatal(e)
		}
	}
	build, e := example.Studio(filepath.Join(root, "studio"))
	if e != nil {
		t.Fatal(e)
	}
	ctx := context.Background()
	r, e := verify.Run(ctx, build.Artifact, example.WindowsSuite(), nil, true)
	write("acceptance.json", r)
	if e != nil {
		t.Fatal(e)
	}
	before, e := os.ReadFile(build.Artifact)
	if e != nil {
		t.Fatal(e)
	}
	h, e := native.Start(ctx, native.Options{Execute: true})
	if e != nil {
		t.Fatal(e)
	}
	defer h.Close()
	call := func(op native.Operation) (any, error) {
		c, cancel := context.WithTimeout(ctx, 10*time.Second)
		defer cancel()
		return h.Call(c, op)
	}
	if _, e = call(native.Operation{Op: "new", As: "sentinel"}); e != nil {
		t.Fatal(e)
	}
	latencies := []float64{}
	for i := 0; i < 5; i++ {
		start := time.Now()
		v, e := call(native.Operation{Op: "eval", Value: "Evaluate = 21 * 2"})
		if e != nil || v.(map[string]any)["result"] != float64(42) {
			t.Fatalf("warm eval: %v %v", v, e)
		}
		latencies = append(latencies, float64(time.Since(start).Microseconds())/1000)
	}
	write("warm-latencies-ms.json", latencies)
	_, e = call(native.Operation{Op: "eval", Value: `100 Err.Raise 5, "WordUp fixture", "Expected runtime failure"`})
	f, ok := e.(*native.Fault)
	if !ok || f.Code != "scratch_vba_failed" {
		t.Fatalf("runtime error was not observable: %v", e)
	}
	write("runtime-error.json", f)
	details := f.Details.(map[string]any)
	if details["line"] != float64(100) || details["number"] != float64(5) {
		t.Fatalf("lost VBA diagnostic: %v", details)
	}
	w, e := project.Open(filepath.Join(root, "studio"))
	if e != nil {
		t.Fatal(e)
	}
	path := filepath.Join(root, "studio", "vba", "Studio.bas")
	source, e := os.ReadFile(path)
	if e != nil {
		t.Fatal(e)
	}
	iterations := []map[string]any{}
	for i := 0; i < 5; i++ {
		start := time.Now()
		revision := append(append([]byte(nil), source...), []byte(fmt.Sprintf("\n' Warm revision %d\n", i))...)
		if e = os.WriteFile(path, revision, 0600); e != nil {
			t.Fatal(e)
		}
		candidate, e := w.Build(filepath.Join(root, "warm.dotm"))
		if e != nil {
			t.Fatal(e)
		}
		compileStart := time.Now()
		if _, e = call(native.Operation{Op: "open", File: candidate.Artifact, As: "warm"}); e != nil {
			t.Fatal(e)
		}
		v, e := call(native.Operation{Op: "compile", Target: "warm", Member: "WordUpStudio"})
		if e != nil || v.(map[string]any)["vba_compiled"] != true {
			t.Fatalf("warm compile: %v %v", v, e)
		}
		v, e = call(native.Operation{Op: "run", Macro: "Studio.NativeAcceptance"})
		if e != nil || v != "WORDUP_NATIVE_OK" {
			t.Fatalf("warm VBA: %v %v", v, e)
		}
		if _, e = call(native.Operation{Op: "unload", Target: "warm", File: candidate.Artifact}); e != nil {
			t.Fatal(e)
		}
		iterations = append(iterations, map[string]any{"iteration": i + 1, "sha256": candidate.SHA256, "build_ms": candidate.DurationMS, "reload_compile_test_unload_ms": float64(time.Since(compileStart).Microseconds()) / 1000, "total_ms": float64(time.Since(start).Microseconds()) / 1000})
	}
	write("warm-edit-cycles.json", iterations)
	if e = os.WriteFile(path, []byte(strings.Replace(string(source), "Option Explicit", "Option Explicit\nPublic Sub DeliberatelyBroken(\n", 1)), 0600); e != nil {
		t.Fatal(e)
	}
	bad, e := w.Build(filepath.Join(root, "broken.dotm"))
	if e != nil {
		t.Fatal(e)
	}
	suite := verify.Suite{Schema: 1, Name: "Deliberate compiler failure", Steps: []verify.Step{
		{Name: "Open broken source", Operation: native.Operation{Op: "open", File: "$artifact", As: "doc"}},
		{Name: "Reject syntax error", Operation: native.Operation{Op: "compile", Target: "doc", Member: "$project"}, TimeoutMS: 3000, Assert: []verify.Assertion{{Path: "/vba_compiled", Kind: "equals", Expected: true}}},
	}}
	failed, e := verify.Run(ctx, bad.Artifact, suite, nil, true)
	write("compiler-error.json", failed)
	if e == nil || failed.Status != "failed" || failed.VBACompiled {
		t.Fatalf("syntax error falsely passed: %v", failed)
	}
	if len(failed.Observations) != 2 || failed.Observations[1].Diagnostics == nil {
		t.Fatal("missing automatic failure diagnostics")
	}
	// The independent Word process must remain usable after the failing host is
	// terminated. No attach-to-user-Word or global process termination is used.
	if _, e = call(native.Operation{Op: "get", Target: "sentinel", Member: "Name"}); e != nil {
		t.Fatalf("other Word instance affected: %v", e)
	}
	runaway, e := native.Start(ctx, native.Options{Execute: true})
	if e != nil {
		t.Fatal(e)
	}
	defer runaway.Close()
	start := time.Now()
	_, e = runaway.Call(ctx, native.Operation{Op: "begin", As: "runaway", TimeoutMS: 1000, Steps: []native.Operation{{Op: "eval", Value: "Do\nLoop"}}})
	if e != nil {
		t.Fatal(e)
	}
	// No polling: the parent watchdog must enforce the limit independently.
	for time.Since(start) < 6*time.Second && runaway.Info()["closed"] != true {
		time.Sleep(50 * time.Millisecond)
	}
	if runaway.Info()["closed"] != true {
		t.Fatal("unpolled infinite loop outlived its watchdog")
	}
	if e = runaway.Close(); e != nil {
		t.Fatal(e)
	}
	_, e = runaway.Call(ctx, native.Operation{Op: "poll", Value: "runaway"})
	f, ok = e.(*native.Fault)
	if !ok || f.Code != "macro_deadline" {
		t.Fatalf("lost runaway termination reason: %v", e)
	}
	write("runaway.json", map[string]any{"elapsed_ms": float64(time.Since(start).Microseconds()) / 1000, "fault": f, "limits": runaway.Info()["limits"]})
	if _, e = call(native.Operation{Op: "get", Target: "sentinel", Member: "Name"}); e != nil {
		t.Fatalf("runaway affected other Word: %v", e)
	}
	after, e := os.ReadFile(build.Artifact)
	if e != nil || office.Hash(before) != office.Hash(after) {
		t.Fatal("original artifact changed")
	}
	write("isolation.json", map[string]any{"other_owned_word_survived": true, "original_sha256": office.Hash(after), "original_unchanged": true})
	t.Log("native evidence:", root)
}
