package ooxml_test

import (
	"bytes"
	"fmt"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap"
	"github.com/eliziff/WordUp/internal/alrmap/alrtest"
	"github.com/eliziff/WordUp/internal/alrmap/kernels"
	"github.com/eliziff/WordUp/internal/alrmap/ooxml"
)

func fldChar(kind string) string {
	return `<w:r><w:fldChar w:fldCharType="` + kind + `"/></w:r>`
}

func TestXmlProtection2(t *testing.T) {
	t.Run("test_nested_fields", func(t *testing.T) {
		unchanged(t, alrtest.Project(t, alrtest.Para(fldChar("begin")+alrtest.Run("towards ", "", true)+fldChar("begin")+alrtest.Run("judgement", "", true)+fldChar("end")+fldChar("end"), ""), nil))
	})
	t.Run("test_unmatched_structure_and_orphan_code", func(t *testing.T) {
		for _, content := range []string{`<w:bookmarkStart w:id="1"/>`, `<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
			`<w:r><w:instrText>REF unknown</w:instrText></w:r>`, `<w:customXmlInsRangeStart w:id="3"/>`} {
			unchanged(t, alrtest.Project(t, alrtest.Para(content+alrtest.Run("towards", "", true), ""), nil))
		}
	})
	t.Run("test_bookmark_crosses_paragraphs", func(t *testing.T) {
		sc := alrtest.Project(t, alrtest.Para(`<w:bookmarkStart w:id="1" w:name="keep"/>`+alrtest.Run("towards", "", true), "")+
			alrtest.Para(alrtest.Run("judgement", "", true)+`<w:bookmarkEnd w:id="1"/>`+alrtest.Run(" towards", "", true), ""), nil)
		output, _ := lexicalOut(t, sc)
		mustText(t, alrtest.Text(t, output, sc.Styles, sc.Part, 0), "towards\rjudgement toward\r")
	})
	t.Run("test_permission_and_comment_ranges", func(t *testing.T) {
		for _, pair := range [][2]string{{"permStart", "permEnd"}, {"commentRangeStart", "commentRangeEnd"}, {"moveToRangeStart", "moveToRangeEnd"}} {
			unchanged(t, alrtest.Project(t, alrtest.Para(`<w:`+pair[0]+` w:id="2"/>`+alrtest.Run("towards", "", true)+`<w:`+pair[1]+` w:id="2"/>`, ""), nil))
		}
	})
	t.Run("test_boundary_inside_token_prevents_partial_delete", func(t *testing.T) {
		unchanged(t, alrtest.Project(t, alrtest.Para(alrtest.Run("towar", "", true)+`<w:bookmarkStart w:id="2"/>`+alrtest.Run("d", "", true)+`<w:bookmarkEnd w:id="2"/>`+alrtest.Run("s", "", true), ""), nil))
	})
	t.Run("test_format_revision_blocks_whole_paragraph", func(t *testing.T) {
		unchanged(t, alrtest.Project(t, alrtest.Para(alrtest.Run("towards", `<w:rPrChange w:id="1"><w:rPr/></w:rPrChange>`, true)+alrtest.Run(" judgement", "", true), ""), nil))
	})
	t.Run("test_mixed_run_replacement_refused_deletion_allowed", func(t *testing.T) {
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run("May ", `<w:color w:val="000001"/>`, true)+alrtest.Run("2, 2024", "", true), ""), nil)
		if _, _, err := alrtest.Lexical(t, sc); err == nil || !strings.Contains(err.Error(), "mixed run") {
			t.Fatalf("got %v", err)
		}
		sc = alrtest.Project(t, alrtest.Para(alrtest.Run("toward", `<w:color w:val="000001"/>`, true)+alrtest.Run("s", "", true), ""), nil)
		out, _ := lexicalOut(t, sc)
		if bytes.Equal(out, sc.Raw) {
			t.Fatal("cross-run deletion was refused")
		}
	})
	t.Run("test_stale_styles_refused", func(t *testing.T) {
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run("towards", "", true), ""), nil)
		c := sc.Containers[0]
		p, err := alrmap.PlanAuthorText(c.Text, alrmap.AuthorTextOptions{Authority: "selected"})
		if err != nil {
			t.Fatal(err)
		}
		bound, err := ooxml.Bind(sc, c.Key, p.Edits, "lexical", "selected")
		if err != nil {
			t.Fatal(err)
		}
		other, err := ooxml.ScanPart(sc.Raw, alrtest.Styles(t, "", "<w:docDefaults/>"), sc.Part)
		if err != nil {
			t.Fatal(err)
		}
		if _, _, err := ooxml.ApplyBound(other, bound); err == nil || !strings.Contains(err.Error(), "stale") {
			t.Fatalf("got %v", err)
		}
	})
	t.Run("test_forged_empty_masks_do_not_bypass_xml", func(t *testing.T) {
		sc := alrtest.Project(t, alrtest.Para("<w:hyperlink>"+alrtest.Run("towards", "", true)+"</w:hyperlink>", ""), nil)
		c := sc.Containers[0]
		p, err := alrmap.PlanAuthorText(c.Text, alrmap.AuthorTextOptions{Authority: "selected"})
		if err != nil {
			t.Fatal(err)
		}
		bound := &ooxml.BoundPlan{Fingerprint: sc.Fingerprint(), Container: c.Key, Family: "lexical", Plan: p}
		if _, _, err := ooxml.ApplyBound(sc, bound); err == nil {
			t.Fatal("forged empty masks bypassed the XML guards")
		}
	})
	t.Run("test_whitespace_attribute_not_silently_added", func(t *testing.T) {
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run("a", "", false)+alrtest.Run("b", "", true), ""), nil)
		c := sc.Containers[0]
		e := kernels.Edit{Start: 1, End: 1, Old: "", New: " ", Rule: "selected-space", ReadStart: 0, ReadEnd: 2}
		bound, err := ooxml.Bind(sc, c.Key, []kernels.Edit{e}, "lexical", "synthetic space insertion")
		if err != nil {
			t.Fatal(err)
		}
		if _, _, err := ooxml.ApplyBound(sc, bound); err == nil || !strings.Contains(err.Error(), "xml:space") {
			t.Fatalf("got %v", err)
		}
	})
	t.Run("test_supplementary_unicode_and_entities", func(t *testing.T) {
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run("😀 & towards", "", true), ""), nil)
		output, _ := lexicalOut(t, sc)
		mustText(t, alrtest.Text(t, output, sc.Styles, sc.Part, 0), "😀 & toward\r")
		if !bytes.Contains(output, []byte("😀 &amp; toward")) {
			t.Fatalf("output %s", output)
		}
	})
	t.Run("test_dtd_utf16_and_strict_namespace_refused", func(t *testing.T) {
		for _, raw := range []string{`<!DOCTYPE x [<!ENTITY x "towards">]><x>&x;</x>`, `<?xml version="1.0" encoding="UTF-16"?><x/>`} {
			if _, err := ooxml.ParseXML([]byte(raw)); err == nil {
				t.Fatalf("parsed %q", raw)
			}
		}
		if _, err := ooxml.ScanPart([]byte(`<document xmlns="urn:strict"/>`), alrtest.Styles(t, "", ""), "word/document.xml"); err == nil {
			t.Fatal("foreign namespace accepted")
		}
	})
	t.Run("test_notes_are_independent", func(t *testing.T) {
		raw := []byte(`<w:footnotes xmlns:w="` + ooxml.W + `"><w:footnote w:id="1">` + alrtest.Para(alrtest.Run("“towards", "", true), "") +
			`</w:footnote><w:footnote w:id="2">` + alrtest.Para(alrtest.Run("judgement", "", true), "") + `</w:footnote></w:footnotes>`)
		sc, err := ooxml.ScanPart(raw, alrtest.Styles(t, "", ""), "word/footnotes.xml")
		if err != nil {
			t.Fatal(err)
		}
		c := sc.Containers[1]
		p, err := alrmap.PlanAuthorText(c.Text, alrmap.AuthorTextOptions{Authority: "selected"})
		if err != nil {
			t.Fatal(err)
		}
		bound, err := ooxml.Bind(sc, c.Key, p.Edits, "lexical", "selected")
		if err != nil {
			t.Fatal(err)
		}
		output, _, err := ooxml.ApplyBound(sc, bound)
		if err != nil {
			t.Fatal(err)
		}
		mustText(t, alrtest.Text(t, output, sc.Styles, sc.Part, 1), "judgment\r")
	})
	t.Run("test_node_segmentation_metamorphic", func(t *testing.T) {
		text := "We move towards judgment."
		for cut := 1; cut < len(text); cut++ {
			t.Run(fmt.Sprintf("cut_%02d", cut), func(t *testing.T) {
				sc := alrtest.Project(t, alrtest.Para(alrtest.Run(text[:cut], "", true)+alrtest.Run(text[cut:], "", true), ""), nil)
				out, _ := lexicalOut(t, sc)
				mustText(t, alrtest.Text(t, out, sc.Styles, sc.Part, 0), "We move toward judgment.\r")
			})
		}
	})
	t.Run("test_two_edits_same_text_node_keep_original_offsets", func(t *testing.T) {
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run("towards judgement towards judgement", "", true), ""), nil)
		out, _ := lexicalOut(t, sc)
		mustText(t, alrtest.Text(t, out, sc.Styles, sc.Part, 0), "toward judgment toward judgment\r")
	})
}
