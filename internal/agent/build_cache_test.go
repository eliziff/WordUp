package agent

import (
	"context"
	"path/filepath"
	"testing"

	"github.com/eliziff/WordUp/internal/project"
)

func TestResidentBuildCacheObservesExternalSourceEdit(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("CacheProof", root); err != nil {
		t.Fatal(err)
	}
	engine := &Engine{Root: root}
	first, err := engine.Call(context.Background(), "build", Parameters{})
	if err != nil {
		t.Fatal(err)
	}
	second, err := engine.Call(context.Background(), "build", Parameters{})
	if err != nil || !second.(*project.BuildReport).Cached {
		t.Fatalf("resident cache miss: %#v %v", second, err)
	}
	if err = project.Write(root, "vba/Added.bas", []byte("Attribute VB_Name = \"Added\"\nPublic Sub AddedMacro()\nEnd Sub\n"), ""); err != nil {
		t.Fatal(err)
	}
	third, err := engine.Call(context.Background(), "build", Parameters{})
	if err != nil {
		t.Fatal(err)
	}
	if third.(*project.BuildReport).Cached || third.(*project.BuildReport).SHA256 == first.(*project.BuildReport).SHA256 {
		t.Fatal("external source edit was hidden by resident cache")
	}
}
