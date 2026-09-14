package office

import (
	"archive/zip"
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
	if err := p.Relationship("word/document.xml", "rId1", VBAProjectRel, "other.bin", ""); err == nil {
		t.Fatal("relationship collision was silently replaced")
	}
	if !bytes.Equal(rels, p.Files[RelPart("word/document.xml")]) {
		t.Fatal("relationship collision changed the original part")
	}
}

func TestKnownChangedPartsSkipReadsWithoutTrustingIncompleteSet(t *testing.T) {
	original, err := BlankPackage().Bytes()
	if err != nil {
		t.Fatal(err)
	}
	p, err := ReadPackage(original)
	if err != nil {
		t.Fatal(err)
	}
	const part = "word/document.xml"
	p.Files[part] = bytes.Replace(p.Files[part], []byte("</w:body>"), []byte("<w:p/></w:body>"), 1)
	if _, err = p.BytesChanged(nil); err == nil {
		t.Fatal("incomplete changed-part set accepted")
	}
	fast, err := p.BytesChanged([]string{part})
	if err != nil {
		t.Fatal(err)
	}
	ordinary, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(fast, ordinary) {
		t.Fatal("known-change serialization differs from checked serialization")
	}
}

func TestChangedPackagePreservesOriginalBackslashMemberNames(t *testing.T) {
	var raw bytes.Buffer
	z := zip.NewWriter(&raw)
	entries := map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="` + CT + `"/>`),
		`customXml\item1.xml`: []byte("opaque"),
		"word/document.xml":   []byte("before"),
	}
	for name, data := range entries {
		w, err := z.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err = w.Write(data); err != nil {
			t.Fatal(err)
		}
	}
	if err := z.Close(); err != nil {
		t.Fatal(err)
	}

	p, err := ReadPackage(raw.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	if got := string(p.Files["customXml/item1.xml"]); got != "opaque" {
		t.Fatalf("logical backslash member lookup = %q", got)
	}
	p.Files["word/document.xml"] = []byte("after")
	out, err := p.BytesChanged([]string{"word/document.xml"})
	if err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(out), int64(len(out)))
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range zr.File {
		if f.Name == `customXml\item1.xml` {
			return
		}
	}
	t.Fatal("changed package rewrote untouched backslash member name")
}
