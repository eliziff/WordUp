package verify

import (
	"github.com/eliziff/WordUp/internal/project"
	"os"
	"path/filepath"
	"testing"
)

func TestSavedReportsSurviveLaterRuns(t *testing.T) {
	root := t.TempDir()
	first := &Report{Schema: 1, Status: "failed", Error: "first failure"}
	if err := SaveReport(root, first); err != nil {
		t.Fatal(err)
	}
	before, err := os.ReadFile(first.SavedReport)
	if err != nil {
		t.Fatal(err)
	}
	second := &Report{Schema: 1, Status: "not_run", Error: "second failure"}
	if err := SaveReport(root, second); err != nil {
		t.Fatal(err)
	}
	after, err := os.ReadFile(first.SavedReport)
	if err != nil || string(before) != string(after) || first.SavedReport == second.SavedReport {
		t.Fatal("earlier run overwritten")
	}
	latest, err := os.ReadFile(filepath.Join(root, "reports", "acceptance.json"))
	if err != nil {
		t.Fatal(err)
	}
	var decoded Report
	if err := DecodeReport(latest, &decoded); err != nil || decoded.SavedReport != second.SavedReport {
		t.Fatalf("latest pointer invalid: %v", err)
	}
}

func TestArchiveSurvivesLatestWriteFailure(t *testing.T) {
	root := t.TempDir()
	if err := os.MkdirAll(filepath.Join(root, "reports", "acceptance.json"), 0700); err != nil {
		t.Fatal(err)
	}
	report := &Report{Schema: 1, Status: "failed"}
	if err := SaveReport(root, report); err == nil {
		t.Fatal("latest failure swallowed")
	}
	if _, err := os.ReadFile(report.SavedReport); err != nil {
		t.Fatalf("durable report lost: %v", err)
	}
}

func TestDecodeRawAndWrappedReports(t *testing.T) {
	original := Report{Schema: 1, Status: "passed", SHA256: "artifact"}
	for _, data := range [][]byte{project.JSON(original), project.JSON(map[string]any{"duration_ms": 12, "result": original})} {
		var decoded Report
		if err := DecodeReport(data, &decoded); err != nil || decoded.SHA256 != "artifact" {
			t.Fatal(decoded, err)
		}
	}
	for _, data := range []string{`null`, `{"result":null}`, `{"result":{"schema":1,"status":"passed"},"error":{"message":"failed"}}`, `{"schema":1,"unknown":true}`} {
		decoded := original
		if err := DecodeReport([]byte(data), &decoded); err == nil {
			t.Fatal("accepted invalid evidence", data)
		}
	}
	var failed Report
	if err := DecodeReport([]byte(`{"result":{"schema":1,"status":"failed"},"error":{"message":"failed"}}`), &failed); err != nil {
		t.Fatal("failed report must remain available for replay", err)
	}
}
