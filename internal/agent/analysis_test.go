package agent

import (
	"encoding/json"
	"github.com/eliziff/WordUp/internal/project"
	"path/filepath"
	"strings"
	"testing"
)

func TestAnalysisRequestCopiesSourceAndRejectsUnsupportedInputs(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("AnalysisTest", root); err != nil {
		t.Fatal(err)
	}
	const source = "Attribute VB_Name = \"Example\"\r\nOption Explicit\r\n"
	if err := project.Write(root, "vba/Example.bas", []byte(source), ""); err != nil {
		t.Fatal(err)
	}
	request, err := analysisRequest(root, []string{"vba/Example.bas"})
	if err != nil {
		t.Fatal(err)
	}
	var decoded struct{ Modules []map[string]string }
	if err := json.Unmarshal([]byte(request), &decoded); err != nil {
		t.Fatal(err)
	}
	if len(decoded.Modules) != 1 || decoded.Modules[0]["Source"] != source || decoded.Modules[0]["Kind"] != "standard" {
		t.Fatal("source changed", request)
	}
	for _, paths := range [][]string{nil, {"vba/ThisDocument.cls"}, {"vba/../project.json"}, {"project.json"}, {"vba/Example.bas", "vba/Example.bas"}} {
		if _, err := analysisRequest(root, paths); err == nil {
			t.Errorf("accepted unsupported paths %v", paths)
		}
	}
	if _, err := analysisRequest(root, []string{"vba/thisdocument.cls"}); err == nil || !strings.Contains(err.Error(), "host/designer metadata") {
		t.Fatalf("case alias bypassed document-module classification: %v", err)
	}
}
