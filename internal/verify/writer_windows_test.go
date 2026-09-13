//go:build windows && (amd64 || arm64)

package verify_test

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/verify"
)

func TestNativeWriterCaptionOutputs(t *testing.T) {
	path := os.Getenv("WORDUP_WRITER_REPORT")
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" || path == "" {
		t.Skip("requires native opt-in and writer comparison report")
	}
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	var report struct {
		Results []struct {
			Operation, Status, Output string
			Control                   []string
		}
	}
	if err := json.Unmarshal(data, &report); err != nil {
		t.Fatal(err)
	}
	count := 0
	for _, row := range report.Results {
		if row.Operation != "control-caption" || row.Status != "passed" {
			continue
		}
		count++
		if len(row.Control) != 2 {
			t.Fatal("missing form/control identity")
		}
		t.Run(filepath.Base(row.Output), func(t *testing.T) {
			host, err := native.Start(context.Background(), native.Options{Directory: t.TempDir(), Execute: true})
			if err != nil {
				t.Fatal(err)
			}
			defer host.Close()
			suite := verify.Suite{Schema: 1, Name: "Writer caption native inspection", Steps: []verify.Step{
				{Name: "Open", Operation: native.Operation{Op: "open", File: "$artifact", As: "doc", Named: map[string]any{"disable_macros": true}}, Assert: []verify.Assertion{{Path: "/macros_disabled_on_open", Kind: "equals", Expected: true}}},
				{Name: "Project", Operation: native.Operation{Op: "get", Target: "doc", Member: "VBProject", As: "project"}},
				{Name: "Components", Operation: native.Operation{Op: "get", Target: "project", Member: "VBComponents", As: "components"}},
				{Name: "Form component", Operation: native.Operation{Op: "invoke", Target: "components", Member: "Item", Args: []any{row.Control[0]}, As: "component"}},
				{Name: "Designer", Operation: native.Operation{Op: "get", Target: "component", Member: "Designer", As: "designer"}},
				{Name: "Controls", Operation: native.Operation{Op: "get", Target: "designer", Member: "Controls", As: "controls"}},
				{Name: "Control", Operation: native.Operation{Op: "invoke", Target: "controls", Member: "Item", Args: []any{row.Control[1]}, As: "control"}},
				{Name: "Caption", Operation: native.Operation{Op: "get", Target: "control", Member: "Caption"}, Assert: []verify.Assertion{{Path: "", Kind: "equals", Expected: "Writer comparison"}}},
			}}
			// Permit the VBComponents.Item method, but open with macros disabled.
			// This suite does not run VBA or mutate the form designer.
			result, err := verify.Run(context.Background(), row.Output, suite, host, true)
			if saveErr := verify.SaveReport(filepath.Dir(path), result); saveErr != nil {
				t.Fatal(saveErr)
			}
			if err != nil {
				t.Fatalf("Native caption failed; report %s: %v", result.SavedReport, err)
			}
			t.Logf("Native caption confirmed; report %s", result.SavedReport)
		})
	}
	if count == 0 {
		t.Fatal("No candidate captions tested")
	}
}
