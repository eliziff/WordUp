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
