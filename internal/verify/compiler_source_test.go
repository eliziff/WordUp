package verify

import "testing"

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
