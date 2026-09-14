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
	hashes := map[string]string{}
	for name, data := range p.Files {
		hashes[name] = Hash(data)
	}
	fastWithHashes, err := p.BytesChangedWithHashes([]string{part}, hashes)
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
	if !bytes.Equal(fastWithHashes, ordinary) {
		t.Fatal("prehashed known-change serialization differs from checked serialization")
	}
}

func TestChangedPartMarkersHandleLeadingHyphenNames(t *testing.T) {
	base := BlankPackage()
	base.Files["-asset.bin"] = []byte("before")
	base.Files["asset.bin"] = []byte("also before")
	if err := base.ContentType("-asset.bin", "application/octet-stream"); err != nil {
		t.Fatal(err)
	}
	if err := base.ContentType("asset.bin", "application/octet-stream"); err != nil {
		t.Fatal(err)
	}
	original, err := base.Bytes()
	if err != nil {
		t.Fatal(err)
	}

	p, err := ReadPackage(original)
	if err != nil {
		t.Fatal(err)
	}
	p.Files["-asset.bin"] = []byte("after")
	changed, err := p.BytesChanged([]string{"-asset.bin"})
	if err != nil {
		t.Fatalf("modified leading-hyphen part rejected: %v", err)
	}
	read, err := ReadPackage(changed)
	if err != nil || string(read.Files["-asset.bin"]) != "after" {
		t.Fatalf("modified leading-hyphen part was not written: %v", err)
	}

	delete(p.Files, "asset.bin")
	deleted, err := p.BytesChanged([]string{"-asset.bin", DeletedPartPrefix + "asset.bin"})
	if err != nil {
		t.Fatalf("deleted part rejected: %v", err)
	}
	read, err = ReadPackage(deleted)
	if err != nil {
		t.Fatal(err)
	}
	if _, exists := read.Files["asset.bin"]; exists {
		t.Fatal("deleted part remained in package")
	}

	// A new -asset.bin and a deleted asset.bin must be representable in the
	// same changed set; the old '-' marker made these two entries identical.
	p.Files["asset.bin"] = []byte("restored")
	delete(p.Files, "-asset.bin")
	delete(p.Files, "asset.bin")
	p.Files["-asset.bin"] = []byte("new")
	combined, err := p.BytesChanged([]string{"-asset.bin", DeletedPartPrefix + "asset.bin"})
	if err != nil {
		t.Fatalf("combined add/delete changed set rejected: %v", err)
	}
	read, err = ReadPackage(combined)
	if err != nil || string(read.Files["-asset.bin"]) != "new" {
		t.Fatalf("combined add/delete result incorrect: %v", err)
	}
	if _, exists := read.Files["asset.bin"]; exists {
		t.Fatal("combined add/delete retained deleted part")
	}
}

func TestFailedVBAAttachDoesNotMutatePackage(t *testing.T) {
	p := BlankPackage()
	p.Files[RelPart("word/document.xml")] = []byte(`<Relationships xmlns="` + RelNS + `"><Relationship Id="rIdVBA" Type="` + VBAProjectRel + `" Target="existing.bin"/></Relationships>`)
	before, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if err = p.SetVBA([]byte("new project")); err == nil {
		t.Fatal("VBA relationship collision was not rejected")
	}
	after, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(before, after) {
		t.Fatal("failed VBA attach mutated the package")
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
