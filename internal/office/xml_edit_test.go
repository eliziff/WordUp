package office

import (
	"bytes"
	"encoding/json"
	"fmt"
	"math"
	"reflect"
	"strings"
	"testing"
)

func TestXMLPatchPlan(t *testing.T) {
	source := []byte(`<r>é<a/>tail</r>`)
	for _, tc := range []struct {
		name    string
		patches []XMLPatch
		want    string
	}{
		{"original offsets", []XMLPatch{{9, 4, "end"}, {5, 4, "<b/>"}}, `<r>é<b/>end</r>`},
		{"ordered inserts", []XMLPatch{{5, 0, "one"}, {5, 0, "two"}, {5, 4, "<b/>"}}, `<r>éonetwo<b/>tail</r>`},
		{"overlap", []XMLPatch{{5, 4, "<b/>"}, {6, 0, "x"}}, ""},
		{"negative", []XMLPatch{{-1, 0, ""}}, ""},
		{"overflow", []XMLPatch{{math.MaxInt, 1, ""}}, ""},
		{"split character", []XMLPatch{{4, 0, "x"}}, ""},
		{"invalid XML", []XMLPatch{{5, 4, "<b>"}}, ""},
		{"invalid UTF8", []XMLPatch{{5, 0, "\xff"}}, ""},
		{"DTD", []XMLPatch{{0, 0, "<!DOCTYPE r>"}}, ""},
	} {
		t.Run(tc.name, func(t *testing.T) {
			plan := append([]XMLPatch(nil), tc.patches...)
			got, err := PatchXML(source, tc.patches)
			if tc.want == "" {
				if err == nil || got != nil {
					t.Fatalf("accepted invalid plan: %q %v", got, err)
				}
			} else if err != nil || string(got) != tc.want {
				t.Fatalf("got %q: %v; want %q", got, err, tc.want)
			}
			if string(source) != `<r>é<a/>tail</r>` || !reflect.DeepEqual(plan, tc.patches) {
				t.Fatal("mutated caller input")
			}
		})
	}
}

func TestStyleBatchPreservesNamespaceAndUnknownXML(t *testing.T) {
	for _, prefix := range []string{"w:", "x:", ""} {
		t.Run(prefix, func(t *testing.T) {
			p := BlankPackage()
			ns := `xmlns="` + W + `" xmlns:w="` + W + `" xmlns:x="` + W + `" xmlns:y="` + W + `" xmlns:e="urn:extension"`
			p.Files["word/styles.xml"] = []byte(`<` + prefix + `styles ` + ns + `><!--keep--><` + prefix + `style x:styleId="A" x:type="paragraph"><` + prefix + `rPr><e:b e:val="keep"/><y:b x:val="1" e:val="opaque"/><x:i/><e:unknown>unchanged</e:unknown></` + prefix + `rPr></` + prefix + `style></` + prefix + `styles>`)
			recipe := StyleRecipe{Styles: []StyleSpec{{ID: "A", Run: map[string]any{"bold": false, "size_pt": 12}}, {ID: "A", Run: map[string]any{"italic": false}}, {ID: "B", Name: "New", Paragraph: map[string]any{"bidi": true}}}}
			if err := ApplyStyles(p, recipe); err != nil {
				t.Fatal(err)
			}
			raw := p.Files["word/styles.xml"]
			for _, keep := range []string{`<!--keep-->`, `<e:b e:val="keep"/>`, `<e:unknown>unchanged</e:unknown>`, `e:val="opaque"`} {
				if !bytes.Contains(raw, []byte(keep)) {
					t.Fatalf("lost %s: %s", keep, raw)
				}
			}
			spans, err := XMLSpans(raw)
			if err != nil {
				t.Fatal(err)
			}
			counts := map[string]int{}
			for _, n := range spans {
				if n.Name.Space == W {
					counts[n.Name.Local]++
					if n.Name.Local == "b" || n.Name.Local == "i" {
						if n.Attribute(W, "val") != "0" {
							t.Fatalf("wrong toggle: %#v", n)
						}
						count := 0
						for _, a := range n.Attr {
							if a.Name.Space == W && a.Name.Local == "val" {
								count++
							}
						}
						if count != 1 {
							t.Fatal("duplicate expanded attribute")
						}
					}
				}
			}
			if counts["style"] != 2 || counts["rPr"] != 1 || counts["b"] != 1 || counts["i"] != 1 {
				t.Fatalf("duplicate/missing children: %v; %s", counts, raw)
			}
			before := p.Clone()
			if err := ApplyStyles(p, StyleRecipe{Styles: []StyleSpec{{ID: "A", Name: "must rollback"}, {ID: "B", Run: map[string]any{"size_pt": json.Number("NaN")}}}}); err == nil {
				t.Fatal("invalid recipe accepted")
			}
			if !reflect.DeepEqual(before.Files, p.Files) {
				t.Fatal("failed transaction modified package")
			}
		})
	}
}

func TestExtendedProperties(t *testing.T) {
	for _, group := range []struct {
		kind  string
		names map[string]string
	}{
		{"run", map[string]string{"double_strike": "dstrike", "outline": "outline", "shadow": "shadow", "emboss": "emboss", "imprint": "imprint", "no_proof": "noProof", "rtl": "rtl", "complex_script": "cs", "bold_complex_script": "bCs", "italic_complex_script": "iCs", "snap_to_grid": "snapToGrid"}},
		{"paragraph", map[string]string{"contextual_spacing": "contextualSpacing", "mirror_indents": "mirrorIndents", "suppress_line_numbers": "suppressLineNumbers", "suppress_auto_hyphens": "suppressAutoHyphens", "bidi": "bidi", "snap_to_grid": "snapToGrid"}},
	} {
		for _, enabled := range []bool{true, false} {
			spec := map[string]any{}
			for key := range group.names {
				spec[key] = enabled
			}
			raw, err := props(nil, group.kind, spec)
			if err != nil {
				t.Fatal(err)
			}
			spans, err := XMLSpans(raw)
			if err != nil {
				t.Fatal(err)
			}
			found := map[string]string{}
			for _, n := range spans {
				if n.Depth == 1 {
					found[n.Name.Local] = n.Attribute(W, "val")
				}
			}
			want := "0"
			if enabled {
				want = "1"
			}
			for _, name := range group.names {
				if found[name] != want {
					t.Fatalf("%s %s: %s", group.kind, name, raw)
				}
			}
		}
	}
	raw, err := props(nil, "run", map[string]any{"character_spacing_pt": -0.5, "position_pt": 2.5, "kerning_pt": 8})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{`<w:spacing w:val="-10"/>`, `<w:position w:val="5"/>`, `<w:kern w:val="16"/>`} {
		if !bytes.Contains(raw, []byte(want)) {
			t.Fatalf("missing %s in %s", want, raw)
		}
	}
	for _, spec := range []map[string]any{{"line_pt": 12, "line_multiple": 1}, {"first_line_pt": 1, "hanging_pt": 1}, {"list_level": 2}, {"line_pt": math.Inf(1)}} {
		if _, err := props(nil, "paragraph", spec); err == nil {
			t.Fatalf("accepted %v", spec)
		}
	}
	if _, err := props(nil, "run", map[string]any{"superscript": true, "subscript": false}); err == nil {
		t.Fatal("ambiguous vertical alignment accepted")
	}
}

func TestNumberingBatchAndSelfClosingRoots(t *testing.T) {
	p := BlankPackage()
	p.Files["word/numbering.xml"] = []byte(`<w:numbering xmlns:w="` + W + `"><w:abstractNum w:abstractNumId="7"/><w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num><!--keep--></w:numbering>`)
	recipe := StyleRecipe{Numbering: []NumberingSpec{{ID: 3, Levels: []NumberLevel{{Format: "decimal", Text: "%1."}}}, {ID: 4, Levels: []NumberLevel{{Level: 1, Format: "lowerLetter", Text: "%2)", Run: map[string]any{"no_proof": true}}}}}}
	if err := ApplyStyles(p, recipe); err != nil {
		t.Fatal(err)
	}
	raw := p.Files["word/numbering.xml"]
	spans, err := XMLSpans(raw)
	if err != nil {
		t.Fatal(err)
	}
	nums, abstracts := 0, 0
	seenNum := false
	for _, n := range spans {
		if n.Depth != 1 {
			continue
		}
		switch n.Name.Local {
		case "num":
			nums++
			seenNum = true
		case "abstractNum":
			abstracts++
			if seenNum {
				t.Fatal("abstract definition after num")
			}
		}
	}
	if nums != 2 || abstracts != 3 || !bytes.Contains(raw, []byte("<!--keep-->")) {
		t.Fatalf("wrong numbering: %s", raw)
	}
	for _, root := range []string{`<w:styles xmlns:w="` + W + `"/>`, `<w:styles xmlns:w="` + W + `"></w:styles>`} {
		got, err := editStyles([]byte(root), []StyleSpec{{ID: "A", Run: map[string]any{"bold": true}}})
		if err != nil || !strings.Contains(string(got), `w:styleId="A"`) {
			t.Fatalf("self-closing root: %s %v", got, err)
		}
	}
}

func BenchmarkStyleBatch(b *testing.B) {
	for _, count := range []int{10, 100} {
		b.Run(fmt.Sprint(count), func(b *testing.B) {
			base := BlankPackage()
			recipe := StyleRecipe{}
			for i := 0; i < count; i++ {
				recipe.Styles = append(recipe.Styles, StyleSpec{ID: fmt.Sprintf("Style%d", i), Run: map[string]any{"font": "Times New Roman", "size_pt": 12, "bold": true, "italic": false, "color": "123456"}, Paragraph: map[string]any{"before_pt": 6, "after_pt": 8, "line_multiple": 1.5, "left_pt": 12, "keep_next": true}})
			}
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				p := base.Clone()
				if err := ApplyStyles(p, recipe); err != nil {
					b.Fatal(err)
				}
			}
		})
	}
}
