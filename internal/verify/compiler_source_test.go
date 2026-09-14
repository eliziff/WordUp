package verify

import (
	"fmt"
	"testing"

	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
)

func TestExportedCompilerLineSkipsOnlyHiddenAttributes(t *testing.T) {
	source := "Attribute VB_Name = \"Example\"\r\nOption Explicit\r\nPublic Sub Example()\r\nAttribute Example.VB_Description = \"Description\"\r\n' Attribute comment stays visible\r\nBroken declaration\r\nEnd Sub\r\n"
	if got := exportedLine(source, 4, "Broken declaration"); got != 6 {
		t.Fatal(got)
	}
	if got := exportedLine(source, 3, "' Attribute comment stays visible"); got != 5 {
		t.Fatal(got)
	}
	if got := exportedLine(source, 4, "Different source"); got != 0 {
		t.Fatal("guessed location", got)
	}
}

func TestProcedureAtLineUsesContainingSourceProcedure(t *testing.T) {
	source := "Attribute VB_Name = \"Example\"\r\nOption Explicit\r\nPublic Sub First()\r\n    Dim value As Long\r\nEnd Sub\r\nPrivate Function Second() As Long\r\n    Second = 42\r\nEnd Function\r\n"
	if got := procedureAtLine(source, 4); got != "First" {
		t.Fatalf("first procedure=%q", got)
	}
	if got := procedureAtLine(source, 7); got != "Second" {
		t.Fatalf("second procedure=%q", got)
	}
	if got := procedureAtLine(source, 9); got != "" {
		t.Fatalf("end line retained procedure=%q", got)
	}
	if got := procedureAtLine(source, 2); got != "" {
		t.Fatalf("module header assigned procedure=%q", got)
	}
}

func TestArtifactSourceUnwrapsWrappedRuntimeLocation(t *testing.T) {
	v := office.NewVBA("Example")
	data, err := v.Rewrite([]office.Module{{Name: "ExampleModule", Kind: "standard", Source: "Attribute VB_Name = \"ExampleModule\"\nOption Explicit\nPublic Sub First()\nBroken\nEnd Sub\n"}}, nil)
	if err != nil {
		t.Fatal(err)
	}
	artifact := &office.Package{Files: map[string][]byte{"word/vbaProject.bin": data}}
	failure := native.Fail("vba_runtime_error", "broken", map[string]any{"runtime_location": map[string]any{
		"location_kind": "native_vbe_selection", "module": "ExampleModule", "line": 3, "line_text": "Broken",
	}})
	artifactSource(fmt.Errorf("wrapped runtime: %w", failure), artifact)
	details := failure.(*native.Fault).Details.(map[string]any)["runtime_location"].(map[string]any)
	if details["source_file"] != "vba/ExampleModule.bas" || details["source_line"] != 4 || details["procedure"] != "First" {
		t.Fatalf("wrapped location was not mapped: %#v", details)
	}
}
