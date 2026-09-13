package verify

import (
	"github.com/eliziff/WordUp/internal/native"
	"os"
	"path/filepath"
	"testing"
)

func TestReplayInputsUseFreshWorkingCopies(t *testing.T) {
	root := t.TempDir()
	source := filepath.Join(root, "original.docx")
	if err := os.WriteFile(source, []byte("original bytes"), 0600); err != nil {
		t.Fatal(err)
	}
	declared := map[string]string{"manuscript": source}
	inputs, err := readInputs(declared, nil)
	if err != nil {
		t.Fatal(err)
	}
	replacements := map[string]string{}
	frozen, err := stageInputs(filepath.Join(root, "run1"), inputs, replacements)
	if err != nil {
		t.Fatal(err)
	}
	working := expand(native.Operation{File: "$input:manuscript$"}, replacements).File
	if working == source || working == frozen["manuscript"].File {
		t.Fatal("test edits original or baseline")
	}
	if err := os.WriteFile(working, []byte("saved by operation"), 0600); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(source, []byte("changed after capture"), 0600); err != nil {
		t.Fatal(err)
	}
	replay, err := readInputs(declared, frozen)
	if err != nil {
		t.Fatal(err)
	}
	if string(replay["manuscript"].data) != "original bytes" {
		t.Fatal("replay used changed input")
	}
	again := map[string]string{}
	if _, err := stageInputs(filepath.Join(root, "run2"), replay, again); err != nil {
		t.Fatal(err)
	}
	b, err := os.ReadFile(again["$input:manuscript$"])
	if err != nil || string(b) != "original bytes" || again["$input:manuscript$"] == working {
		t.Fatal("working copy not reset", err)
	}
	if err := os.WriteFile(frozen["manuscript"].File, []byte("tampered baseline"), 0600); err != nil {
		t.Fatal(err)
	}
	if _, err := readInputs(declared, frozen); err == nil {
		t.Fatal("tampered input accepted")
	}
	if _, err := readInputs(declared, map[string]InputSnapshot{}); err == nil {
		t.Fatal("uncaptured input accepted")
	}
}
