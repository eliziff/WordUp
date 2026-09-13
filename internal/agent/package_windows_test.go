//go:build windows && (amd64 || arm64)

package agent_test

import (
	"encoding/json"
	"os"
	"os/exec"
	"path/filepath"
	"testing"
)

func TestPackagedAnalysisRunsOutsideRepository(t *testing.T) {
	source := os.Getenv("WORDUP_PACKAGE")
	if source == "" {
		t.Skip("set WORDUP_PACKAGE to a staged distribution directory")
	}
	packageDir := filepath.Join(t.TempDir(), "package")
	if err := os.CopyFS(packageDir, os.DirFS(source)); err != nil {
		t.Fatal(err)
	}
	exe := filepath.Join(packageDir, "wordup.exe")
	workspace := filepath.Join(t.TempDir(), "workspace")
	runPackage(t, exe, "new", "PackageAnalysis", workspace)
	if err := os.WriteFile(filepath.Join(workspace, "vba", "AnalysisProof.bas"), []byte("Attribute VB_Name = \"AnalysisProof\"\nOption Explicit\nPublic Sub Run()\nDim unused As Long\nEnd Sub\n"), 0600); err != nil {
		t.Fatal(err)
	}
	params := filepath.Join(t.TempDir(), "parameters.json")
	if err := os.WriteFile(params, []byte(`{"paths":["vba/AnalysisProof.bas"]}`), 0600); err != nil {
		t.Fatal(err)
	}
	output := runPackage(t, exe, "-w", workspace, "call", "vba.analyze", "@"+params)
	var response struct {
		Result struct {
			Engine string `json:"engine"`
		} `json:"result"`
	}
	if err := json.Unmarshal(output, &response); err != nil {
		t.Fatal(err)
	}
	if response.Result.Engine != "Rubberduck" {
		t.Fatalf("packaged analysis helper did not run: %s", output)
	}
}

func runPackage(t *testing.T, exe string, args ...string) []byte {
	t.Helper()
	command := exec.Command(exe, args...)
	command.Env = append(os.Environ(), `PATH=C:\Windows\System32;C:\Windows`)
	output, err := command.CombinedOutput()
	if err != nil {
		t.Fatalf("%s: %v\n%s", args, err, output)
	}
	return output
}
