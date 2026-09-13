package office

import (
	"bytes"
	"testing"
)

func TestExistingPackageMetadataIsNotReserialized(t *testing.T) {
	empty := &Package{Files: map[string][]byte{}}
	if err := empty.ContentType("custom.xml", "application/xml"); err != nil || len(empty.Files["[Content_Types].xml"]) == 0 {
		t.Fatalf("New default content type not saved: %v", err)
	}
	ct := []byte(`<Types xmlns="` + CT + `"><Default ContentType="application/vnd.ms-office.vbaProject" Extension="bin"/><Override ContentType="` + MainDOTM + `" PartName="/word/document.xml"/></Types>`)
	rels := []byte(`<Relationships xmlns="` + RelNS + `"><Relationship Target="vbaProject.bin" Type="` + VBAProjectRel + `" Id="rId1" /></Relationships>`)
	p := &Package{Files: map[string][]byte{"[Content_Types].xml": ct, RelPart("word/document.xml"): rels}}
	if err := p.ContentType("word/vbaProject.bin", "application/vnd.ms-office.vbaProject"); err != nil {
		t.Fatal(err)
	}
	if err := p.ContentType("word/document.xml", MainDOTM); err != nil {
		t.Fatal(err)
	}
	if err := p.Relationship("word/document.xml", "rId1", VBAProjectRel, "vbaProject.bin", ""); err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(ct, p.Files["[Content_Types].xml"]) || !bytes.Equal(rels, p.Files[RelPart("word/document.xml")]) {
		t.Fatal("Unchanged metadata rewritten")
	}
	if err := p.ContentType("word/vbaProject.bin", "different"); err != nil {
		t.Fatal(err)
	}
	if bytes.Equal(ct, p.Files["[Content_Types].xml"]) {
		t.Fatal("Changed type ignored")
	}
	// An explicit override must win over the matching default on subsequent calls.
	if err := p.ContentType("word/vbaProject.bin", "application/vnd.ms-office.vbaProject"); err != nil {
		t.Fatal(err)
	}
	spans, err := XMLSpans(p.Files["[Content_Types].xml"])
	if err != nil {
		t.Fatal(err)
	}
	for _, s := range spans {
		if s.Name.Local == "Override" && s.Attribute("", "PartName") == "/word/vbaProject.bin" && s.Attribute("", "ContentType") != "application/vnd.ms-office.vbaProject" {
			t.Fatal("Override precedence lost")
		}
	}
	if err := p.Relationship("word/document.xml", "rId1", VBAProjectRel, "other.bin", ""); err != nil {
		t.Fatal(err)
	}
	if bytes.Equal(rels, p.Files[RelPart("word/document.xml")]) {
		t.Fatal("Changed target ignored")
	}
}
