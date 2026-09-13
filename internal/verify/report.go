package verify

import (
	"bytes"
	"encoding/json"
	"fmt"
	"github.com/eliziff/WordUp/internal/project"
	"os"
	"path/filepath"
)

// SaveReport archives each run before replacing the convenience latest pointer.
// Failed/not-run evidence is retained too; archiving does not certify a pass.
func SaveReport(root string, report *Report) error {
	if report == nil {
		return fmt.Errorf("cannot save a nil acceptance report")
	}
	dir, err := project.Under(root, "reports/acceptance-runs")
	if err != nil {
		return err
	}
	if err = os.MkdirAll(dir, 0700); err != nil {
		return err
	}
	run, err := os.MkdirTemp(dir, "run-")
	if err != nil {
		return err
	}
	report.SavedReport = filepath.Join(run, "acceptance.json")
	data := project.JSON(report)
	if err = project.AtomicWrite(report.SavedReport, data); err != nil {
		return fmt.Errorf("archive acceptance report: %w", err)
	}
	if err = project.Write(root, "reports/acceptance.json", data, ""); err != nil {
		return fmt.Errorf("report retained at %s; latest report could not be updated: %w", report.SavedReport, err)
	}
	return nil
}

// DecodeReport accepts acceptance.json or a saved CLI call result.
// Validation of execution, assertions and artifact identity stays with the caller.
func DecodeReport(data []byte, report *Report) error {
	// Decode into fresh storage so omitted fields cannot reuse earlier evidence.
	var decoded Report
	if err := decodeReport(data, &decoded); err != nil {
		return err
	}
	*report = decoded
	return nil
}

func decodeReport(data []byte, report *Report) error {
	var fields map[string]json.RawMessage
	if err := project.ReadJSON(data, &fields); err != nil {
		return err
	}
	if _, wrapped := fields["result"]; wrapped {
		var envelope struct {
			DurationMS float64         `json:"duration_ms"`
			Result     json.RawMessage `json:"result"`
			Error      json.RawMessage `json:"error"`
		}
		if err := project.ReadJSON(data, &envelope); err != nil {
			return err
		}
		if err := project.ReadJSON(envelope.Result, report); err != nil {
			return err
		}
		if report.Status == "passed" && len(envelope.Error) > 0 && !bytes.Equal(bytes.TrimSpace(envelope.Error), []byte("null")) {
			return fmt.Errorf("CLI error conflicts with passing acceptance report")
		}
	} else if err := project.ReadJSON(data, report); err != nil {
		return err
	}
	if report.Schema != 1 {
		return fmt.Errorf("expected schema 1 acceptance report")
	}
	return nil
}
