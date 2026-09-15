package office

import (
	"archive/zip"
	"bytes"
	"crypto/rand"
	"fmt"
	"io"
	"math"
	"os"
	"strings"
	"testing"
)

func TestCompression(t *testing.T) {
	for _, n := range []int{0, 1, 64, 3640, 4095, 4096, 4097, 10000, 65536} {
		for _, random := range []bool{false, true} {
			b := bytes.Repeat([]byte("This is a repeated VBA line.\r\n"), n/29+1)[:n]
			if random {
				rand.Read(b)
			}
			c := Compress(b)
			d, e := Decompress(c)
			if e != nil || len(d) < len(b) || !bytes.Equal(b, d[:len(b)]) || !bytes.Equal(d[len(b):], make([]byte, len(d)-len(b))) {
				t.Fatalf("n=%d random=%v: %v (%d bytes)", n, random, e, len(d))
			}
		}
	}
}
func TestCompound(t *testing.T) {
	c := NewCompound()
	for _, n := range []int{0, 1, 63, 64, 65, 4095, 4096, 4097, 65536, 9 << 20} {
		b := make([]byte, n)
		rand.Read(b)
		if e := c.Set("VBA/"+strings.Repeat("x", n%23+1)+string(rune('A'+n%26)), b); e != nil {
			t.Fatal(e)
		}
	}
	b, e := c.Bytes()
	if e != nil {
		t.Fatal(e)
	}
	d, e := ReadCompound(b)
	if e != nil {
		t.Fatal(e)
	}
	for p, x := range c.Entries {
		if x.Kind == 2 {
			y, e := d.Stream(p)
			if e != nil || !bytes.Equal(x.Data, y) {
				t.Fatalf("%s mismatch: %v", p, e)
			}
		}
	}
}
func specimen(t *testing.T) *VBA {
	t.Helper()
	p := os.Getenv("WORDUP_SPECIMEN")
	if p == "" {
		t.Skip("private specimen path not set")
	}
	z, e := zip.OpenReader(p)
	if e != nil {
		t.Fatal(e)
	}
	defer z.Close()
	for _, f := range z.File {
		if f.Name == "word/vbaProject.bin" {
			r, e := f.Open()
			if e != nil {
				t.Fatal(e)
			}
			b, e := io.ReadAll(r)
			r.Close()
			if e != nil {
				t.Fatal(e)
			}
			v, e := ReadVBA(b)
			if e != nil {
				t.Fatal(e)
			}
			return v
		}
	}
	t.Fatal("no vba")
	return nil
}
func TestSpecimen(t *testing.T) {
	v := specimen(t)
	t.Logf("%s %d modules", v.Name, len(v.Modules))
	for _, m := range v.Modules {
		if m.Kind == "form" {
			f, e := ReadForm(v.CFB, m.Name, v.Codepage)
			if e != nil {
				t.Errorf("%s: %v", m.Name, e)
				continue
			}
			t.Logf("%s: %v", m.Name, f.ControlNames())
			s, e := f.Streams()
			if e != nil {
				t.Fatal(e)
			}
			for p, b := range s {
				orig, e := v.CFB.Stream(m.Name + "/" + p)
				if e == nil && !bytes.Equal(orig, b) {
					t.Errorf("unchanged form stream differs %s/%s", m.Name, p)
				}
			}
		}
	}
	b, e := v.Rewrite(v.Modules, nil)
	if e != nil {
		t.Fatal(e)
	}
	if _, e = ReadVBA(b); e != nil {
		t.Fatal(e)
	}
}
func TestFormTabOrderShiftsSiblings(t *testing.T) {
	f, err := NewForm("TabOrder", 1252)
	if err != nil {
		t.Fatal(err)
	}
	err = f.Apply(Design{Name: "TabOrder", Controls: []ControlDesign{
		{Name: "title", Type: "Label"}, {Name: "hint", Type: "Label"},
		{Name: "first", Type: "CommandButton", Properties: map[string]any{"TabIndex": 0}},
		{Name: "second", Type: "CommandButton", Properties: map[string]any{"TabIndex": 1}},
	}})
	if err != nil {
		t.Fatal(err)
	}
	want := map[string]int64{"first": 0, "second": 1, "title": 2, "hint": 3}
	for _, c := range f.Design().Controls {
		if c.Properties["TabIndex"] != want[c.Name] {
			t.Fatalf("%s TabIndex=%v want %d", c.Name, c.Properties["TabIndex"], want[c.Name])
		}
	}
}

func TestFormRejectsUnrepresentableGeometry(t *testing.T) {
	for name, value := range map[string]any{
		"negative": -1.0,
		"zero":     0.0,
		"tiny":     0.0001,
		"huge int": int64(100001),
	} {
		t.Run(name, func(t *testing.T) {
			r, err := defaultControl("Label", "Sample")
			if err != nil {
				t.Fatal(err)
			}
			if err := applyRecord(r, map[string]any{"Width": value, "Height": 20.0}, "Size"); err == nil {
				t.Fatal("invalid control geometry was accepted")
			}
		})
	}
	r, err := defaultControl("Label", "Sample")
	if err != nil {
		t.Fatal(err)
	}
	if err := applyRecord(r, map[string]any{"Font": map[string]any{"Size": 0.0}}, "Size"); err == nil {
		t.Fatal("zero font size was accepted")
	}
}

func TestFormCreate(t *testing.T) {
	f, e := NewForm("TestForm", 1252)
	if e != nil {
		t.Fatal(e)
	}
	types := []string{"Label", "CommandButton", "TextBox", "ListBox", "ComboBox", "CheckBox", "OptionButton", "ToggleButton", "Image", "SpinButton", "ScrollBar", "TabStrip", "Frame"}
	d := Design{Name: "TestForm", Properties: map[string]any{"Caption": "Unicode résumé", "Width": 420., "Height": 300.}}
	for _, k := range types {
		d.Controls = append(d.Controls, ControlDesign{Name: "ctl" + k, Type: k, Properties: map[string]any{"Left": 10., "Top": 20., "Width": 100., "Height": 25.}})
	}
	if e = f.Apply(d); e != nil {
		t.Fatal(e)
	}
	s, e := f.Streams()
	if e != nil {
		t.Fatal(e)
	}
	c := NewCompound()
	for p, b := range s {
		c.Set("TestForm/"+p, b)
	}
	g, e := ReadForm(c, "TestForm", 1252)
	if e != nil {
		t.Fatal(e)
	}
	if len(g.ControlNames()) != len(types) {
		t.Fatal("control count")
	}
}
func TestNewProject(t *testing.T) {
	v := NewVBA("Example")
	m := []Module{{Name: "ThisDocument", Kind: "document", Source: "Attribute VB_Name = \"ThisDocument\"\n"}, {Name: "Main", Kind: "standard", Source: "Option Explicit\nPublic Function Answer() As Long\nAnswer=42\nEnd Function\n"}}
	b, e := v.Rewrite(m, nil)
	if e != nil {
		t.Fatal(e)
	}
	w, e := ReadVBA(b)
	if e != nil {
		t.Fatal(e)
	}
	w.Modules = append(w.Modules, Module{Name: "Other", Kind: "standard", Source: "Public Sub Hello()\nEnd Sub\n"})
	if _, e = w.Rewrite(w.Modules, nil); e != nil {
		t.Fatal(e)
	}
}
func FuzzDecompress(f *testing.F) {
	f.Add([]byte{1})
	f.Add(Compress([]byte("Hello\n")))
	f.Fuzz(func(t *testing.T, b []byte) { Decompress(b) })
}

func TestFormUnicodeControl(t *testing.T) {
	f, e := NewForm("UnicodeForm", 1252)
	if e != nil {
		t.Fatal(e)
	}
	d := Design{Name: "UnicodeForm", Controls: []ControlDesign{{Name: "Label1", Type: "Label", Properties: map[string]any{"Caption": "標題 — résumé"}}}}
	if e = f.Apply(d); e != nil {
		t.Fatal(e)
	}
	streams, e := f.Streams()
	if e != nil {
		t.Fatal(e)
	}
	c := NewCompound()
	for p, b := range streams {
		if e = c.Set("UnicodeForm/"+p, b); e != nil {
			t.Fatal(e)
		}
	}
	g, e := ReadForm(c, "UnicodeForm", 1252)
	if e != nil {
		t.Fatal(e)
	}
	if got := g.Design().Controls[0].Properties["Caption"]; got != "標題 — résumé" {
		t.Fatalf("caption=%v", got)
	}
	if e = f.Apply(Design{Name: "UnicodeForm", Properties: map[string]any{"Caption": "標題"}}); e == nil {
		t.Fatal("unrepresentable VBFrame caption silently accepted")
	}
}
func TestFailedRewriteIsTransactional(t *testing.T) {
	v := NewVBA("Trial")
	before, _ := v.CFB.Bytes()
	_, e := v.Rewrite([]Module{{Name: "A", Kind: "standard", Source: "Option Explicit"}, {Name: "BAD-NAME", Kind: "class"}}, nil)
	if e == nil {
		t.Fatal("expected invalid identifier")
	}
	after, _ := v.CFB.Bytes()
	if !bytes.Equal(before, after) {
		t.Fatal("failed build mutated baseline")
	}
}
func TestCompressionRatio(t *testing.T) {
	b := bytes.Repeat([]byte("Public Function Answer() As Long\r\nAnswer = 42\r\nEnd Function\r\n"), 1000)
	c := Compress(b)
	if len(c) > len(b)/5 {
		t.Fatalf("poor compression %d/%d", len(c), len(b))
	}
}
func BenchmarkCompression(b *testing.B) {
	src := bytes.Repeat([]byte("Public Function Answer() As Long\r\nAnswer = 42\r\nEnd Function\r\n"), 5000)
	b.SetBytes(int64(len(src)))
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		Compress(src)
	}
}
func BenchmarkNewProject(b *testing.B) {
	mods := []Module{{Name: "ThisDocument", Kind: "document", Source: "Option Explicit\n"}, {Name: "Main", Kind: "standard", Source: strings.Repeat("' source line\n", 1000)}}
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		if _, e := NewVBA("Bench").Rewrite(mods, nil); e != nil {
			b.Fatal(e)
		}
	}
}

// This deliberately examines each wire chunk separately: a self-roundtrip
// alone would miss a writer+reader that share the same boundary error.
func TestCompressionWireBoundaries(t *testing.T) {
	for _, n := range []int{3641, 4095, 4096, 8191, 12287} {
		raw := make([]byte, n)
		rand.Read(raw)
		wire := Compress(raw)
		for p := 1; p < len(wire); {
			total := int(U16(wire, p)&4095) + 3
			if p+total > len(wire) {
				t.Fatal("chunk exceeds container")
			}
			chunk, err := Decompress(append([]byte{1}, wire[p:p+total]...))
			if err != nil {
				t.Fatal(err)
			}
			if p+total < len(wire) && len(chunk) != 4096 {
				t.Fatalf("short interior chunk: %d", len(chunk))
			}
			p += total
		}
	}
}

func TestMultiPageCreateEditRemove(t *testing.T) {
	f, e := NewForm("Editor", 1252)
	if e != nil {
		t.Fatal(e)
	}
	d := Design{Name: "Editor", Controls: []ControlDesign{{Name: "Tabs", Type: "MultiPage", Pages: []ControlDesign{
		{Name: "One", Type: "Page", Properties: map[string]any{"Caption": "Résumé"}, Controls: []ControlDesign{{Name: "Text1", Type: "TextBox", Properties: map[string]any{"MultiLine": true, "EnterKeyBehavior": true}}}},
		{Name: "Two", Type: "Page", Controls: []ControlDesign{{Name: "OK", Type: "CommandButton"}}},
	}}}}
	if e = f.Apply(d); e != nil {
		t.Fatal(e)
	}
	c := NewCompound()
	if e = f.WriteBack(c); e != nil {
		t.Fatal(e)
	}
	for p, x := range c.Entries {
		if x.Kind == 1 && p != "Editor" {
			if bytes.Equal(x.Raw[80:96], make([]byte, 16)) {
				t.Fatalf("container CLSID absent %s", p)
			}
		}
	}
	g, e := ReadForm(c, "Editor", 1252)
	if e != nil {
		t.Fatal(e)
	}
	exp := g.Design()
	if len(exp.Controls) != 1 || len(exp.Controls[0].Pages) != 2 {
		t.Fatalf("bad exported pages: %+v", exp)
	}
	exp.Controls[0].Pages = append(exp.Controls[0].Pages, ControlDesign{Name: "Three", Type: "Page", Properties: map[string]any{"Caption": "中文"}})
	if e = g.Apply(exp); e != nil {
		t.Fatal(e)
	}
	if e = g.WriteBack(c); e != nil {
		t.Fatal(e)
	}
	before := len(c.Entries)
	g, e = ReadForm(c, "Editor", 1252)
	if e != nil {
		t.Fatal(e)
	}
	if e = g.Apply(Design{Name: "Editor", Controls: []ControlDesign{{Name: "Tabs", Type: "MultiPage", Remove: []string{"One"}}}}); e != nil {
		t.Fatal(e)
	}
	if e = g.WriteBack(c); e != nil {
		t.Fatal(e)
	}
	if len(c.Entries) >= before {
		t.Fatal("deleted page storage survived")
	}
	g, e = ReadForm(c, "Editor", 1252)
	if e != nil {
		t.Fatal(e)
	}
	if len(g.Design().Controls[0].Pages) != 2 {
		t.Fatal("wrong surviving page count")
	}
}
func TestSpecimenFormsEditable(t *testing.T) {
	v := specimen(t)
	for _, m := range v.Modules {
		if m.Kind != "form" {
			continue
		}
		f, e := ReadForm(v.CFB, m.Name, v.Codepage)
		if e != nil {
			t.Fatal(e)
		}
		d := f.Design()
		d.Properties["Caption"] = "Verified binary edit"
		minimal, e := ReadForm(v.CFB, m.Name, v.Codepage)
		if e != nil {
			t.Fatal(e)
		}
		if e = minimal.Apply(Design{Name: m.Name, Properties: map[string]any{"Caption": "Verified binary edit"}}); e != nil {
			t.Fatal(e)
		}
		if e = f.Apply(d); e != nil {
			t.Fatalf("%s: %v", m.Name, e)
		}
		fullStreams, e := f.Streams()
		if e != nil {
			t.Fatal(e)
		}
		minimalStreams, e := minimal.Streams()
		if e != nil {
			t.Fatal(e)
		}
		if len(fullStreams) != len(minimalStreams) {
			t.Fatal("unchanged properties altered stream inventory")
		}
		for name, data := range fullStreams {
			if !bytes.Equal(data, minimalStreams[name]) {
				t.Fatalf("%s/%s: full design rewrote unchanged data", m.Name, name)
			}
		}
		c := v.CFB.Clone()
		if e = f.WriteBack(c); e != nil {
			t.Fatalf("%s: %v", m.Name, e)
		}
		g, e := ReadForm(c, m.Name, v.Codepage)
		if e != nil {
			t.Fatal(e)
		}
		if g.Design().Properties["Caption"] != "Verified binary edit" {
			t.Fatal("caption edit missing")
		}
	}
}
func TestFormFailedEditTransactional(t *testing.T) {
	f, _ := NewForm("F", 1252)
	before, _ := f.Streams()
	err := f.Apply(Design{Name: "F", Controls: []ControlDesign{{Name: "New", Type: "Label"}, {Name: "Bad", Type: "MissingType"}}})
	if err == nil {
		t.Fatal("expected error")
	}
	after, _ := f.Streams()
	if !sameStreams(before, after) {
		t.Fatal("failed edit mutated form")
	}
}

func TestFootnoteEmptyParagraphAndRelationships(t *testing.T) {
	p := BlankPackage()
	recipe := ContentRecipe{Blocks: []Block{{Inlines: []Inline{{Text: "Body"}, {Footnote: []Block{{Type: "xml", XML: "<w:p/>"}, {Inlines: []Inline{{Text: "Source", URL: "https://example.test/citation"}}}}}}}}}
	if e := Compose(p, recipe, nil); e != nil {
		t.Fatal(e)
	}
	if e := p.Validate(); e != nil {
		t.Fatal(e)
	}
	for _, x := range []string{"<w:footnoteRef/>", "<w:hyperlink"} {
		if !strings.Contains(string(p.Files["word/footnotes.xml"]), x) {
			t.Fatalf("missing footnote structure %s", x)
		}
	}
	if !strings.Contains(string(p.Files["word/_rels/footnotes.xml.rels"]), "https://example.test/citation") {
		t.Fatal("link relationship attached to wrong part")
	}
}

func TestComposePreservesExistingFootnotesAndAllocatesNewIDs(t *testing.T) {
	p := BlankPackage()
	p.Files["word/footnotes.xml"] = []byte(`<w:footnotes xmlns:w="` + W + `"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="7"><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Existing note</w:t></w:r></w:p></w:footnote></w:footnotes>`)
	if err := Compose(p, ContentRecipe{Blocks: []Block{{Inlines: []Inline{{Text: "Body"}, {Footnote: []Block{{Inlines: []Inline{{Text: "New note"}}}}}}}}}, nil); err != nil {
		t.Fatal(err)
	}
	footnotes := string(p.Files["word/footnotes.xml"])
	if !strings.Contains(footnotes, `w:id="7"`) || !strings.Contains(footnotes, "Existing note") || !strings.Contains(footnotes, `<w:i/>`) {
		t.Fatalf("existing footnote was not preserved: %s", footnotes)
	}
	if !strings.Contains(footnotes, `w:id="8"`) || !strings.Contains(string(p.Files["word/document.xml"]), `w:id="8"`) {
		t.Fatalf("new footnote did not allocate the next ID: %s", footnotes)
	}
}

func TestBuildingBlocksPreserveExistingGlossaryDefinitions(t *testing.T) {
	p := BlankPackage()
	styles := []byte(`<w:styles xmlns:w="` + W + `"><w:style w:type="paragraph" w:styleId="GlossaryOnly"><w:name w:val="Glossary only"/></w:style></w:styles>`)
	numbering := []byte(`<w:numbering xmlns:w="` + W + `"><w:abstractNum w:abstractNumId="91"/></w:numbering>`)
	p.Files["word/glossary/styles.xml"] = styles
	p.Files["word/glossary/numbering.xml"] = numbering
	if err := AddBuildingBlocks(p, []BuildingBlock{{Name: "Snippet", Blocks: []Block{{Text: "Saved"}}}}, nil); err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(p.Files["word/glossary/styles.xml"], styles) || !bytes.Equal(p.Files["word/glossary/numbering.xml"], numbering) {
		t.Fatal("adding a building block overwrote existing glossary definitions")
	}
	if err := p.Validate(); err != nil {
		t.Fatal("preserved glossary package is invalid:", err)
	}
}

func TestBuildingBlocksRejectCaseCollisions(t *testing.T) {
	p := BlankPackage()
	before, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	err = AddBuildingBlocks(p, []BuildingBlock{
		{Name: "Snippet", Blocks: []Block{{Text: "one"}}},
		{Name: "snippet", Blocks: []Block{{Text: "two"}}},
	}, nil)
	if err == nil {
		t.Fatal("case-colliding building blocks were accepted")
	}
	after, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(before, after) {
		t.Fatal("failed building-block collision mutated the package")
	}
}

func TestPackageValidationRejectsCaseCollidingParts(t *testing.T) {
	p := BlankPackage()
	p.Files["word/STYLES.xml"] = append([]byte(nil), p.Files["word/styles.xml"]...)
	if err := p.ContentType("word/STYLES.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"); err != nil {
		t.Fatal(err)
	}
	if err := p.Validate(); err == nil || !strings.Contains(err.Error(), "case-colliding package parts") {
		t.Fatalf("case-colliding package parts were accepted: %v", err)
	}
	if _, err := p.Bytes(); err == nil || !strings.Contains(err.Error(), "case-colliding package parts") {
		t.Fatalf("low-level package writer bypassed case-collision guard: %v", err)
	}
}

func TestComposeFailureDoesNotMutatePackage(t *testing.T) {
	p := BlankPackage()
	p.Files[RelPart("word/document.xml")] = []byte(`<Relationships xmlns="` + RelNS + `"><Relationship Id="wwheader" Type="other" Target="existing.xml"/></Relationships>`)
	before, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	err = Compose(p, ContentRecipe{Page: &PageSpec{Header: []Block{{Text: "new"}}}}, nil)
	if err == nil {
		t.Fatal("relationship collision was not rejected")
	}
	after, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(before, after) {
		t.Fatal("failed composition mutated the package")
	}
}

func TestBlankPackageHasDOCXContentType(t *testing.T) {
	p := BlankPackage()
	if !strings.Contains(string(p.Files["[Content_Types].xml"]), "wordprocessingml.document.main+xml") {
		t.Fatal("seed.docx has wrong package content type")
	}
}

func TestWordInteroperabilityMetadata(t *testing.T) {
	p := BlankPackage()
	if e := p.SetVBA([]byte("binary fixture")); e != nil {
		t.Fatal(e)
	}
	if !bytes.Contains(p.Files["word/_rels/document.xml.rels"], []byte(VBAProjectRel)) {
		t.Fatal("Word requires the Microsoft VBA relationship namespace")
	}
	if len(p.Files["word/vbaData.xml"]) == 0 {
		t.Fatal("missing supplemental VBA part")
	}
	if e := AddBuildingBlocks(p, []BuildingBlock{{Name: "Note", Blocks: []Block{{Text: "Reusable note"}}}}, nil); e != nil {
		t.Fatal(e)
	}
	if bytes.Contains(p.Files["word/glossary/document.xml"], []byte("<w:types>")) {
		t.Fatal("explicit types hide ordinary entries in Word's BuildingBlockEntries")
	}
	if bytes.Contains(p.Files["word/glossary/_rels/document.xml.rels"], []byte("../styles.xml")) {
		t.Fatal("main and glossary documents must own separate style parts")
	}
	f, e := NewForm("TestForm", 1252)
	if e != nil {
		t.Fatal(e)
	}
	if e = f.Apply(Design{Name: "TestForm", Mode: "replace", Controls: []ControlDesign{{Name: "Run", Type: "CommandButton"}}}); e != nil {
		t.Fatal(e)
	}
	streams, e := f.Streams()
	if e != nil {
		t.Fatal(e)
	}
	expected := fmt.Sprintf("TypeInfoVer = %d", f.root.record.values["ShapeCookie"])
	if !bytes.Contains(streams["\x03VBFrame"], []byte(expected)) {
		t.Fatal("designer type information is out of sync")
	}
	for _, s := range []string{"Run", "Café — 文"} {
		raw := encodeStrings([]string{s})
		values, e := arrayStrings(raw, 1252)
		if e != nil || len(values) != 1 || values[0] != s {
			t.Fatalf("MSForms string roundtrip: %q %v", values, e)
		}
	}
}
func TestTypeInfoVersionUpdatePreservesDesignerFormatting(t *testing.T) {
	frame := "VERSION 5.00\r\n   TypeInfoVer     =   97   ' retained\r\nEnd\r\n"
	want := "VERSION 5.00\r\n   TypeInfoVer     =   123   ' retained\r\nEnd\r\n"
	if got := syncTypeInfoVersion(frame, 123); got != want {
		t.Fatalf("designer formatting changed:\nwant %q\n got %q", want, got)
	}
}

func TestRecipeRejectsFractionalAndInvalidNumberingFields(t *testing.T) {
	for name, spec := range map[string]map[string]any{
		"fractional outline":    {"outline_level": 1.5},
		"fractional list id":    {"list_id": 2.5},
		"zero list id":          {"list_id": 0},
		"negative list id":      {"list_id": -1},
		"fractional list level": {"list_id": 2, "list_level": 1.5},
		"negative list level":   {"list_id": 2, "list_level": -1},
		"deep list level":       {"list_id": 2, "list_level": 9},
	} {
		t.Run(name, func(t *testing.T) {
			if _, err := props(nil, "paragraph", spec); err == nil {
				t.Fatal("invalid integer field was rounded or accepted")
			}
		})
	}
	raw, err := props(nil, "paragraph", map[string]any{"outline_level": 1.0, "list_id": 7.0, "list_level": 8.0})
	if err != nil || !strings.Contains(string(raw), `w:outlineLvl w:val="1"`) || !strings.Contains(string(raw), `w:numId w:val="7"`) || !strings.Contains(string(raw), `w:ilvl w:val="8"`) {
		t.Fatalf("valid integer fields were not emitted: %s (%v)", raw, err)
	}
}

func TestNumberingRecipePreservesExistingLevelOverrides(t *testing.T) {
	p := BlankPackage()
	p.Files["word/numbering.xml"] = []byte(`<w:numbering xmlns:w="` + W + `"><w:abstractNum w:abstractNumId="2"/><w:num w:numId="7"><w:abstractNumId w:val="2"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride></w:num></w:numbering>`)
	err := ApplyStyles(p, StyleRecipe{Numbering: []NumberingSpec{{ID: 7, Levels: []NumberLevel{{Level: 0, Format: "decimal", Text: "%1."}}}}})
	if err != nil {
		t.Fatal(err)
	}
	text := string(p.Files["word/numbering.xml"])
	if !strings.Contains(text, `<w:abstractNumId w:val="3"/>`) || !strings.Contains(text, `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride>`) {
		t.Fatalf("numbering update discarded existing override: %s", text)
	}
}

func TestNumberingRecipeRejectsUnusableLevelValues(t *testing.T) {
	for _, level := range []NumberLevel{
		{Level: 0, Start: -1, Format: "decimal", Text: "%1."},
		{Level: 0, Format: "", Text: "%1."},
		{Level: 0, Format: "decimal", Text: ""},
	} {
		p := BlankPackage()
		err := ApplyStyles(p, StyleRecipe{Numbering: []NumberingSpec{{ID: 4, Levels: []NumberLevel{level}}}})
		if err == nil {
			t.Fatalf("accepted unusable numbering level: %#v", level)
		}
	}
}

func TestComposeRejectsInvalidLayoutMeasurements(t *testing.T) {
	cases := []struct {
		name   string
		recipe ContentRecipe
	}{
		{name: "page width", recipe: ContentRecipe{Page: &PageSpec{WidthPT: math.NaN()}}},
		{name: "page height", recipe: ContentRecipe{Page: &PageSpec{HeightPT: math.Inf(1)}}},
		{name: "tiny page width", recipe: ContentRecipe{Page: &PageSpec{WidthPT: 0.001}}},
		{name: "page margin", recipe: ContentRecipe{Page: &PageSpec{MarginsPT: map[string]float64{"left": math.NaN()}}}},
		{name: "table column", recipe: ContentRecipe{Blocks: []Block{{Type: "table", ColumnsPT: []float64{math.Inf(1)}, Rows: [][]Cell{{{Text: "cell"}}}}}}},
		{name: "huge table column", recipe: ContentRecipe{Blocks: []Block{{Type: "table", ColumnsPT: []float64{math.MaxFloat64}, Rows: [][]Cell{{{Text: "cell"}}}}}}},
		{name: "table cell", recipe: ContentRecipe{Blocks: []Block{{Type: "table", ColumnsPT: []float64{72}, Rows: [][]Cell{{{Text: "cell", WidthPT: math.NaN()}}}}}}},
		{name: "negative table cell", recipe: ContentRecipe{Blocks: []Block{{Type: "table", ColumnsPT: []float64{72}, Rows: [][]Cell{{{Text: "cell", WidthPT: -1}}}}}}},
		{name: "image size", recipe: ContentRecipe{Blocks: []Block{{Inlines: []Inline{{Image: "figure.png", WidthPT: math.NaN(), HeightPT: 12}}}}}},
		{name: "tiny image size", recipe: ContentRecipe{Blocks: []Block{{Inlines: []Inline{{Image: "figure.png", WidthPT: 0.000001, HeightPT: 12}}}}}},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			p := BlankPackage()
			before, err := p.Bytes()
			if err != nil {
				t.Fatal(err)
			}
			asset := func(string) ([]byte, error) { return []byte("image"), nil }
			if err := Compose(p, tc.recipe, asset); err == nil {
				t.Fatal("non-finite layout measurement was accepted")
			}
			after, err := p.Bytes()
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(before, after) {
				t.Fatal("failed composition mutated the package")
			}
		})
	}
}

func TestPackageRefusesSymlinkEntries(t *testing.T) {
	var b bytes.Buffer
	z := zip.NewWriter(&b)
	h := &zip.FileHeader{Name: "word/document.xml", Method: zip.Store}
	h.SetMode(os.ModeSymlink | 0777)
	w, e := z.CreateHeader(h)
	if e != nil {
		t.Fatal(e)
	}
	w.Write([]byte("../../target"))
	z.Close()
	if _, e = ReadPackage(b.Bytes()); e == nil {
		t.Fatal("accepted symlink ZIP entry")
	}
}
func FuzzCompound(f *testing.F) {
	c := NewCompound()
	c.Set("test", []byte("hi"))
	b, _ := c.Bytes()
	f.Add(b)
	f.Add([]byte("not-cfb"))
	f.Fuzz(func(t *testing.T, b []byte) {
		if len(b) > 1<<20 {
			return
		}
		_, _ = ReadCompound(b)
	})
}
