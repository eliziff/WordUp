package verify

import (
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"os"
	"path/filepath"
	"testing"
)

func TestFreezeRelocationAndTamper(t *testing.T) {
	// Synthetic bookkeeping test; native replay is tested separately in Word.
	root := t.TempDir()
	evidence := filepath.Join(root, "evidence")
	artifact := filepath.Join(evidence, "input.dotm")
	snapshot := filepath.Join(evidence, "output.xml")
	xml := []byte(`<w:p xmlns:w="urn:test"><w:r><w:i/><w:t>Untouched</w:t></w:r></w:p>`)
	for path, data := range map[string][]byte{artifact: []byte("synthetic artifact"), snapshot: xml} {
		if err := project.AtomicWrite(path, data); err != nil {
			t.Fatal(err)
		}
	}
	suite := Suite{Schema: 1, Name: "synthetic freeze", Steps: []Step{{Name: "capture", Operation: native.Operation{Op: "xml.snapshot"}, Assert: []Assertion{{Path: "/text", Kind: "equals", Expected: snapshot}}}}}
	report := &Report{Schema: 1, Status: "passed", WordExecuted: true, OS: "synthetic", Arch: "synthetic", ArtifactSnapshot: artifact, EvidenceDirectory: evidence, SHA256: office.Hash([]byte("synthetic artifact")), Suite: suite, SuiteSHA256: office.Hash(project.JSON(suite)), Assertions: 1, Observations: []Observation{{Name: "capture", Passed: true, Assertions: 1, Result: map[string]any{"file": snapshot, "sha256": office.Hash(xml), "text": snapshot}}}}
	if err := SaveReport(root, report); err != nil {
		t.Fatal(err)
	}
	bundle := filepath.Join(root, "bundle")
	if _, err := Freeze(report.SavedReport, bundle); err != nil {
		t.Fatal(err)
	}
	if _, err := Freeze(report.SavedReport, bundle); err == nil {
		t.Fatal("existing bundle overwritten")
	}
	moved := filepath.Join(root, "moved")
	if err := os.Rename(bundle, moved); err != nil {
		t.Fatal(err)
	}
	if err := os.Rename(evidence, evidence+"-retired"); err != nil {
		t.Fatal(err)
	}
	loaded, _, err := LoadReport(filepath.Join(moved, "bundle.json"))
	if err != nil {
		t.Fatal(err)
	}
	if loaded.SuiteSHA256 != report.SuiteSHA256 || loaded.Observations[0].Result.(map[string]any)["text"] != snapshot {
		t.Fatal("relocation changed suite or non-path data")
	}
	actual, err := snapshotBytes(loaded.Observations[0].Result)
	if err != nil || string(actual) != string(xml) {
		t.Fatal("frozen XML changed", err)
	}
	if err := os.WriteFile(loaded.ArtifactSnapshot, []byte("tampered"), 0600); err != nil {
		t.Fatal(err)
	}
	if _, _, err := LoadReport(filepath.Join(moved, "bundle.json")); err == nil {
		t.Fatal("tampered bundle accepted")
	}
}

func TestBundleRejectsTraversal(t *testing.T) {
	root := t.TempDir()
	raw := project.JSON(Report{Schema: 1})
	if err := project.Write(root, "report.json", raw, ""); err != nil {
		t.Fatal(err)
	}
	bundle := frozenBundle{Schema: 1, Kind: "wordup-parity", ReportSHA256: office.Hash(raw), Files: []frozenFile{{Source: "old", Path: "../outside", SHA256: "invalid"}}}
	if err := project.Write(root, "bundle.json", project.JSON(bundle), ""); err != nil {
		t.Fatal(err)
	}
	if _, _, err := LoadReport(filepath.Join(root, "bundle.json")); err == nil {
		t.Fatal("unsafe bundle path accepted")
	}
}
