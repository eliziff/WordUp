package agent

import (
	"context"
	"github.com/eliziff/WordUp/internal/project"
	"strings"
	"testing"
)

func TestDirectExpectedXML(t *testing.T) {
	root := t.TempDir()
	// Copy source; change only the intended property, never reconstruct runs.
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:bookmarkStart w:id="1" w:name="keep"/><w:r><w:rPr><w:i/></w:rPr><w:t>Résumé</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:bookmarkEnd w:id="1"/></w:p>`
	expected := strings.Replace(input, `w:val="Normal"`, `w:val="Heading1"`, 1)
	variant := strings.ReplaceAll(strings.ReplaceAll(expected, "w:", "x:"), "xmlns:w", "xmlns:x")
	if err := project.Write(root, "expected.xml", []byte(expected), ""); err != nil {
		t.Fatal(err)
	}
	e := &Engine{Root: root}
	for _, tc := range []struct {
		name, actual, mode string
		pass               bool
	}{
		{"match", expected, "", true},
		{"unchanged", input, "", false},
		{"lost italic", strings.Replace(expected, "<w:i/>", "", 1), "", false},
		{"lost bookmark", strings.Replace(expected, `<w:bookmarkEnd w:id="1"/>`, "", 1), "semantic", false},
		{"wrong note", strings.Replace(expected, `w:footnoteReference w:id="2"`, `w:footnoteReference w:id="3"`, 1), "semantic", false},
		{"serialization exact", variant, "", false},
		{"serialization semantic", variant, "semantic", true},
		{"malformed", "<broken>", "", false},
		{"invalid mode", expected, "guess", false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			if err := project.Write(root, "actual.xml", []byte(tc.actual), ""); err != nil {
				t.Fatal(err)
			}
			got, err := e.Call(context.Background(), "xml.verify", Parameters{Reference: "expected.xml", Path: "actual.xml", Comparison: tc.mode})
			if (err == nil) != tc.pass {
				t.Fatalf("result=%v err=%v", got, err)
			}
			if tc.pass && got.(map[string]any)["matches_expected"] != true {
				t.Fatal(got)
			}
		})
	}
	if e.host != nil {
		t.Fatal("comparison started Word")
	}
	unchanged, err := project.Read(root, "expected.xml")
	if err != nil || string(unchanged) != expected {
		t.Fatal("comparison changed expected XML")
	}
}

func TestExpectedXMLToolDiscovery(t *testing.T) {
	found := false
	for _, tool := range Tools() {
		if strings.HasPrefix(tool.Name, "gold.") {
			t.Fatalf("annotation workflow still advertised: %s", tool.Name)
		}
		if tool.Name == "xml.verify" {
			found = true
		}
	}
	if !found {
		t.Fatal("direct XML verification not discoverable")
	}
}
