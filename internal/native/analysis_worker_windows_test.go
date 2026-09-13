//go:build windows && (amd64 || arm64)

package native

import (
	"context"
	"crypto/sha256"
	"encoding/json"
	"fmt"
	"os"
	"testing"
	"time"
)

func TestAnalysisWorkerIntegration(t *testing.T) {
	helper := os.Getenv("WORDUP_ANALYSIS_TEST_HELPER")
	if helper == "" {
		t.Skip("set WORDUP_ANALYSIS_TEST_HELPER to the built analysis worker")
	}
	defer StopOfficeTools()
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	const source = "Attribute VB_Name = \"Example\"\nPublic Sub Run()\nEnd Sub\n"
	request, err := json.Marshal(map[string]any{
		"Modules":     []any{map[string]any{"Name": "Example", "Kind": "standard", "Path": "vba/Example.bas", "Source": source}},
		"Inspections": []string{"OptionExplicitInspection"},
	})
	if err != nil {
		t.Fatal(err)
	}
	first, err := callOfficeWorker(ctx, helper, []string{"analyze", string(request)})
	if err != nil {
		t.Fatal(err)
	}
	if first["error"] != nil || first["scope"] != "source-only" || first["native_compilation"] != false {
		t.Fatalf("unexpected analysis: %v", first)
	}
	findings, ok := first["diagnostics"].([]any)
	if !ok || len(findings) != 1 {
		t.Fatalf("missing finding: %v", first)
	}
	finding, ok := findings[0].(map[string]any)
	if !ok || finding["SHA256"] != fmt.Sprintf("%x", sha256.Sum256([]byte(source))) || finding["Path"] != "vba/Example.bas" || finding["Line"] != float64(2) {
		t.Fatalf("source identity or mapped line changed: %v", findings)
	}
	if first["helper_reused"] != false {
		t.Fatal("first request unexpectedly reused worker")
	}
	bad, err := callOfficeWorker(ctx, helper, []string{"analyze", "{}"})
	if err != nil || bad["error"] == nil {
		t.Fatalf("invalid request accepted: %v %v", bad, err)
	}
	second, err := callOfficeWorker(ctx, helper, []string{"analyze", string(request)})
	if err != nil || second["error"] != nil || second["helper_pid"] != first["helper_pid"] || second["helper_reused"] != true {
		t.Fatalf("warm worker recovery failed: %v %v", second, err)
	}
	if err := lockOfficeWorkers(ctx); err != nil {
		t.Fatal(err)
	}
	worker := officeWorkers.all[helper]
	unlockOfficeWorkers()
	StopOfficeTools()
	if worker.cmd.ProcessState == nil {
		t.Fatal("stop did not reap owned worker")
	}
	third, err := callOfficeWorker(ctx, helper, []string{"analyze", string(request)})
	if err != nil || third["error"] != nil || third["helper_reused"] != false {
		t.Fatalf("fresh worker after stop failed: %v %v", third, err)
	}
}
