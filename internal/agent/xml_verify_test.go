package agent

import (
	"context"
	"github.com/eliziff/WordUp/internal/office"
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

func TestXMLVerifySelectsPackagePartWithoutExtraction(t *testing.T) {
	root := t.TempDir()
	pkg := office.BlankPackage()
	pkg.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:r><w:t>keep</w:t></w:r></w:p></w:body></w:document>`)
	data, err := pkg.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "sample.docx", data, ""); err != nil {
		t.Fatal(err)
	}
	expected := pkg.Files["word/document.xml"]
	if err := project.Write(root, "expected.xml", expected, ""); err != nil {
		t.Fatal(err)
	}
	e := &Engine{Root: root}
	params := Parameters{Reference: "expected.xml", Path: "sample.docx", Part: "word/document.xml"}
	if result, err := e.Call(context.Background(), "xml.verify", params); err != nil || result.(map[string]any)["matches_expected"] != true {
		t.Fatalf("package part did not compare directly: result=%v err=%v", result, err)
	}
	flat := `<pkg:package xmlns:pkg="` + "http://schemas.microsoft.com/office/2006/xmlPackage" + `"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/xml"><pkg:xmlData>` + string(expected) + `</pkg:xmlData></pkg:part></pkg:package>`
	if err := project.Write(root, "sample-flat.xml", []byte(flat), ""); err != nil {
		t.Fatal(err)
	}
	params.Path = "sample-flat.xml"
	if result, err := e.Call(context.Background(), "xml.verify", params); err != nil || result.(map[string]any)["matches_expected"] != true {
		t.Fatalf("Flat OPC part did not compare directly: result=%v err=%v", result, err)
	}
	params.Path = "sample.docx"
	pkg.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:r><w:t>changed</w:t></w:r></w:p></w:body></w:document>`)
	changed, err := pkg.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "sample.docx", changed, ""); err != nil {
		t.Fatal(err)
	}
	if result, err := e.Call(context.Background(), "xml.compare", params); err != nil || result.(map[string]any)["equal"] == true {
		t.Fatalf("changed package part unexpectedly compared equal: result=%v err=%v", result, err)
	}
	for _, part := range []string{"../word/document.xml", "word/missing.xml"} {
		params.Part = part
		if _, err := e.Call(context.Background(), "xml.verify", params); err == nil {
			t.Fatal("unsafe or missing package part accepted", part)
		} else if !strings.Contains(err.Error(), "XML package part") && !strings.Contains(err.Error(), "unsafe") {
			t.Fatalf("unexpected part error for %s: %v", part, err)
		}
	}
	params.Path, params.Part = "sample-flat.xml", "word/missing.xml"
	if _, err := e.Call(context.Background(), "xml.verify", params); err == nil {
		t.Fatal("missing Flat OPC part accepted")
	}
}
