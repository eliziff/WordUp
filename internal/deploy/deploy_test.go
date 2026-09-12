package deploy

import (
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
	"os"
	"path/filepath"
	"testing"
)

// Synthetic evidence only tests deployment policy; it is never native evidence.
func proofFixture(t *testing.T) (string, string, string, verify.Report) {
	t.Helper()
	root := t.TempDir()
	artifact := filepath.Join(root, "artifact.dotm")
	proof := filepath.Join(root, "proof.json")
	b := []byte("synthetic-artifact-for-policy-tests")
	os.WriteFile(artifact, b, 0600)
	suite := verify.Suite{Schema: 1, Name: "MOCK policy fixture NOT Word verification", Steps: []verify.Step{{Name: "mock", Operation: native.Operation{Op: "get", Member: "Version"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "MOCK"}}}}}
	r := verify.Report{Schema: 1, Status: "passed", WordExecuted: true, FreshProcess: true, SHA256: office.Hash(b), Suite: suite, SuiteSHA256: office.Hash(project.JSON(suite)), Assertions: 1, Observations: []verify.Observation{{Name: "mock", Passed: true, Assertions: 1, Result: "MOCK"}}}
	os.WriteFile(proof, project.JSON(r), 0600)
	return root, artifact, proof, r
}
func TestDeploymentRefusesOfflineAndHashMismatch(t *testing.T) {
	root, a, p, r := proofFixture(t)
	r.WordExecuted = false
	os.WriteFile(p, project.JSON(r), 0600)
	if _, _, e := Prepare(root, a, p, filepath.Join(root, "target.dotm")); e == nil {
		t.Fatal("offline proof accepted")
	}
	r.WordExecuted = true
	r.SHA256 = "wrong"
	os.WriteFile(p, project.JSON(r), 0600)
	if _, _, e := Prepare(root, a, p, filepath.Join(root, "target.dotm")); e == nil {
		t.Fatal("wrong artifact hash accepted")
	}
}
func TestDeploymentBackupAndStaleTarget(t *testing.T) {
	root, a, p, _ := proofFixture(t)
	target := filepath.Join(root, "installed.dotm")
	os.WriteFile(target, []byte("old"), 0600)
	file, _, e := Prepare(root, a, p, target)
	if e != nil {
		t.Fatal(e)
	}
	os.WriteFile(target, []byte("concurrent"), 0600)
	if _, e = Activate(file); e == nil {
		t.Fatal("overwrote concurrent change")
	}
	os.WriteFile(target, []byte("old"), 0600)
	plan, e := Activate(file)
	if e != nil {
		t.Fatal(e)
	}
	b, _ := os.ReadFile(plan.Backup)
	if string(b) != "old" {
		t.Fatal("backup differs")
	}
	b, _ = os.ReadFile(target)
	expected, _ := os.ReadFile(a)
	if string(b) != string(expected) || plan.State != "installed" {
		t.Fatal("installation mismatch")
	}
}

func TestInterruptedInstallAndRestore(t *testing.T) {
	root, a, proof, _ := proofFixture(t)
	target := filepath.Join(root, "installed.dotm")
	if err := os.WriteFile(target, []byte("previous"), 0600); err != nil {
		t.Fatal(err)
	}
	file, p, err := Prepare(root, a, proof, target)
	if err != nil {
		t.Fatal(err)
	}
	// Simulate interruption between replacing the target and recording success.
	if err = os.WriteFile(filepath.Join(filepath.Dir(file), "previous.dotm"), []byte("previous"), 0600); err != nil {
		t.Fatal(err)
	}
	data, _ := os.ReadFile(p.Candidate)
	if err = os.WriteFile(target, data, 0600); err != nil {
		t.Fatal(err)
	}
	p, err = Activate(file)
	if err != nil || p.State != "installed" {
		t.Fatalf("recover: %+v %v", p, err)
	}
	if err = os.WriteFile(target, []byte("user edit"), 0600); err != nil {
		t.Fatal(err)
	}
	if _, err = Restore(file); err == nil {
		t.Fatal("overwrote user edit")
	}
	if err = os.WriteFile(target, data, 0600); err != nil {
		t.Fatal(err)
	}
	p, err = Restore(file)
	if err != nil || p.State != "restored" {
		t.Fatalf("restore: %+v %v", p, err)
	}
	got, _ := os.ReadFile(target)
	if string(got) != "previous" {
		t.Fatal("wrong restored bytes")
	}
	if _, err = Restore(file); err != nil {
		t.Fatal("restore is not idempotent", err)
	}
}
