package gold

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

func TestStructureRolesAreExact(t *testing.T) {
	for _, role := range []string{"", "body|title", "heading|body", "Body", "body "} {
		if validRole(role, 0) {
			t.Fatalf("invalid role accepted: %q", role)
		}
	}
	if !validRole("body", 0) || !validRole("metadata", 0) || validRole("metadata", 1) || validRole("body", 1) || !validRole("heading", 9) || validRole("heading", 10) {
		t.Fatal("role/level contract violated")
	}
}

func TestManualCitationFixture(t *testing.T) {
	root := filepath.Join("..", "..")
	path := "testdata/gold/citation-direct/fixture.json"
	f, r, err := ValidateFixture(root, path)
	if err != nil {
		t.Fatal(err)
	}
	if r["negative_controls_rejected"] != 3 || r["editorially_approved"] != false {
		t.Fatal(r)
	}
	if _, err := VerifyFixture(root, path, f.Expected.Path); err != nil {
		t.Fatal(err)
	}
	for _, c := range f.Counterexamples {
		r, err := VerifyFixture(root, path, c.Artifact.Path)
		if err == nil || r["matches_expected"] != false {
			t.Fatalf("false pass: %s", c.Name)
		}
	}
	if _, err := VerifyFixture(root, path, f.Input.Path); err == nil {
		t.Fatal("unchanged operation passed")
	}
}

func TestStructureValidationAndAgreement(t *testing.T) {
	root := t.TempDir()
	pkg := office.BlankPackage()
	pkg.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:r><w:t>Heading</w:t></w:r></w:p><w:p><w:r><w:t>Body é</w:t></w:r></w:p></w:body></w:document>`)
	data, err := pkg.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if err := project.Write(root, "source.docx", data, ""); err != nil {
		t.Fatal(err)
	}
	d := Document{Source: "source.docx", SHA256: office.Hash(data), Part: "word/document.xml", Labels: []Label{{Paragraph: 1, Text: "Heading", Role: "heading", Level: 1, Rationale: "Title of this section"}, {Paragraph: 2, Text: "Body é", Role: "body", Parent: 1, Rationale: "Section prose"}}}
	d.Coverage.First = 1
	d.Coverage.Last = 2
	a := Annotation{Schema: 1, Annotator: "a", Status: "proposed", Documents: []Document{d}}
	write := func(name string, v any) {
		t.Helper()
		if err := os.WriteFile(filepath.Join(root, name), project.JSON(v), 0600); err != nil {
			t.Fatal(err)
		}
	}
	write("a.json", a)
	a.Status = "silver"
	write("silver.json", a)
	if _, report, err := ValidateStructure(root, "silver.json"); err != nil || report["status"] != "silver" || report["editorially_approved"] != false {
		t.Fatalf("direct structure silver rejected or certified: %v %v", report, err)
	}
	a.Status = "proposed"
	a.Annotator = "b"
	write("b.json", a)
	r, err := CompareStructure(root, "a.json", "b.json")
	if err != nil || r["primary_label_agreements"] != 2 {
		t.Fatalf("%v %v", r, err)
	}
	a.Documents[0].Labels[1].Role = "quotation"
	write("b.json", a)
	r, err = CompareStructure(root, "a.json", "b.json")
	if err != nil || r["primary_label_disagreements"] != 1 {
		t.Fatalf("%v %v", r, err)
	}
	for _, mutate := range []func(*Annotation){
		func(a *Annotation) { a.Status = "frozen" },
		func(a *Annotation) { a.Documents[0].SHA256 = "wrong" },
		func(a *Annotation) { a.Documents[0].Labels[1].Text = "wrong" },
		func(a *Annotation) { a.Documents[0].Labels[1].Parent = 2 },
		func(a *Annotation) { a.Documents[0].Labels[0].Level = 0 },
		func(a *Annotation) { a.Documents[0].Labels = a.Documents[0].Labels[:1] },
		func(a *Annotation) { a.Documents[0].Source = "../outside.docx" },
	} {
		var fresh Annotation
		if _, err := read(root, "a.json", &fresh); err != nil {
			t.Fatal(err)
		}
		mutate(&fresh)
		write("bad.json", fresh)
		if _, _, err := ValidateStructure(root, "bad.json"); err == nil {
			t.Fatal("invalid annotation accepted")
		}
	}
}

func TestNestedParagraphLocations(t *testing.T) {
	data := []byte(`<w:p xmlns:w="` + office.W + `"><w:r><w:t>outer</w:t><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:r></w:p>`)
	text, err := paragraphTexts(data)
	if err != nil || strings.Join(text, "|") != "outerinner|inner" {
		t.Fatalf("%v %v", text, err)
	}
}

func TestSemanticXMLVariation(t *testing.T) {
	a := []byte(`<w:p xmlns:w="urn:w"><w:r><w:i/></w:r></w:p>`)
	b := []byte(`<x:p xmlns:x="urn:w"><x:r><x:i/></x:r></x:p>`)
	matched, err := equal(a, b, "semantic_xml")
	if err != nil || !matched {
		t.Fatal("prefix-only change rejected")
	}
	bad := []byte(`<x:p xmlns:x="urn:w"><x:r/></x:p>`)
	matched, err = equal(a, bad, "semantic_xml")
	if err != nil || matched {
		t.Fatal("missing italics escaped semantic comparison")
	}
}

func TestMinimalFixtureAndOptionalSafeguards(t *testing.T) {
	root := t.TempDir()
	data := []byte(`<p>Expected</p>`)
	if err := os.WriteFile(filepath.Join(root, "expected.xml"), data, 0600); err != nil {
		t.Fatal(err)
	}
	minimal := map[string]any{"schema": 1, "expected": Artifact{Path: "expected.xml", SHA256: office.Hash(data)}}
	write := func() {
		t.Helper()
		if err := os.WriteFile(filepath.Join(root, "fixture.json"), project.JSON(minimal), 0600); err != nil {
			t.Fatal(err)
		}
	}
	write()
	f, report, err := ValidateFixture(root, "fixture.json")
	if err != nil || f.Comparison != "exact_xml" || report["negative_controls_rejected"] != 0 || report["editorially_approved"] != false {
		t.Fatalf("minimal fixture: %v %v", report, err)
	}
	if _, err := VerifyFixture(root, "fixture.json", "expected.xml"); err != nil {
		t.Fatal(err)
	}
	minimal["status"] = "silver"
	write()
	if _, r, err := ValidateFixture(root, "fixture.json"); err != nil || r["status"] != "silver" || r["editorially_approved"] != false {
		t.Fatalf("direct model silver rejected or overclaimed: %v %v", r, err)
	}
	for key, value := range map[string]any{
		"expected":        Artifact{Path: "expected.xml", SHA256: "wrong"},
		"input":           Artifact{Path: "expected.xml", SHA256: "wrong"},
		"counterexamples": []Counterexample{{Name: "actually correct", Artifact: Artifact{Path: "expected.xml", SHA256: office.Hash(data)}}},
		"revision_policy": "typo",
		"comparison":      "ignore_everything",
		"status":          "approved",
	} {
		old, exists := minimal[key]
		minimal[key] = value
		write()
		if _, _, err := ValidateFixture(root, "fixture.json"); err == nil {
			t.Fatalf("invalid supplied %s accepted", key)
		}
		if exists {
			minimal[key] = old
		} else {
			delete(minimal, key)
		}
	}
}
