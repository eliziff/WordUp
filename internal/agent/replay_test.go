package agent

import (
	"context"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestReplayRejectsTamperedSnapshotBeforeWord(t *testing.T) {
	root := t.TempDir()
	snapshot := filepath.Join(root, "input.dotm")
	if err := os.WriteFile(snapshot, []byte("changed"), 0600); err != nil {
		t.Fatal(err)
	}
	report := verify.Report{Schema: 1, SHA256: "original", ArtifactSnapshot: snapshot, Suite: verify.Suite{Schema: 1, Name: "replay", Steps: []verify.Step{{Name: "observe", Operation: native.Operation{Op: "get", Member: "Version"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "16"}}}}}}
	if err := project.Write(root, "report.json", project.JSON(report), ""); err != nil {
		t.Fatal(err)
	}
	e := &Engine{Root: root, Execute: true}
	_, err := e.Call(context.Background(), "test.replay", Parameters{Reference: "report.json"})
	if err == nil || !strings.Contains(err.Error(), "snapshot hash mismatch") || e.host != nil {
		t.Fatalf("unsafe replay: %v", err)
	}
}

func TestReplayRejectsInvalidRequestsBeforeStartingWord(t *testing.T) {
	e := &Engine{Root: t.TempDir()}
	if _, err := e.Call(context.Background(), "test.replay", Parameters{}); err == nil || !strings.Contains(err.Error(), "--execute") {
		t.Fatal(err)
	}
	e.Execute = true
	if _, err := e.Call(context.Background(), "test.replay", Parameters{Suite: &verify.Suite{}}); err == nil || !strings.Contains(err.Error(), "recorded suite") {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(e.Root, "bad.json"), []byte(`{"suite":{},"unrecognized":true}`), 0600); err != nil {
		t.Fatal(err)
	}
	if _, err := e.Call(context.Background(), "test.replay", Parameters{Reference: "bad.json"}); err == nil {
		t.Fatal("invalid report accepted")
	}
	if e.host != nil {
		t.Fatal("invalid replay started Word")
	}
}

func TestInvalidAssertionDoesNotStartWarmHost(t *testing.T) {
	e := &Engine{Root: t.TempDir(), Execute: true}
	suite := &verify.Suite{Schema: 1, Name: "invalid", Steps: []verify.Step{{Name: "edit", Operation: native.Operation{Op: "eval", Value: "ActiveDocument.Content.Text = 1"}, Assert: []verify.Assertion{{Kind: "matches", Expected: "["}}}}}
	_, err := e.Call(context.Background(), "test", Parameters{Path: "missing.dotm", Suite: suite})
	if err == nil || !strings.Contains(err.Error(), "regexp") {
		t.Fatal(err)
	}
	if e.host != nil {
		t.Fatal("malformed assertion started Word")
	}
}
