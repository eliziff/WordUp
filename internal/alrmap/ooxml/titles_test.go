package ooxml_test

import (
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap"
	"github.com/eliziff/WordUp/internal/alrmap/alrtest"
	"github.com/eliziff/WordUp/internal/alrmap/citations"
	"github.com/eliziff/WordUp/internal/alrmap/kernels"
	"github.com/eliziff/WordUp/internal/alrmap/ooxml"
)

// note mirrors TitleCapabilities.note: one footnote holding one paragraph.
func note(t *testing.T, text, props string) *ooxml.Scan {
	t.Helper()
	raw := []byte(`<w:footnotes xmlns:w="` + ooxml.W + `"><w:footnote w:id="1">` + alrtest.Para(alrtest.Run(text, "", true), props) + `</w:footnote></w:footnotes>`)
	sc, err := ooxml.ScanPart(raw, alrtest.Styles(t, "", ""), "word/footnotes.xml")
	if err != nil {
		t.Fatal(err)
	}
	return sc
}

func TestTitleCapabilities(t *testing.T) {
	t.Run("test_title_capitalization_is_not_spelling_permission", func(t *testing.T) {
		sc := note(t, "A Author, “towards a judgement” (2024) 1 Example Law Review 2.", "")
		bound, err := alrmap.PlanSourceTitles(sc, "footnote:1", alrtest.Reg, "en", "refuse")
		if err != nil {
			t.Fatal(err)
		}
		out, _, err := ooxml.ApplyBound(sc, bound)
		if err != nil {
			t.Fatal(err)
		}
		mustText(t, alrtest.Text(t, out, sc.Styles, sc.Part, 0), "A Author, “Towards a Judgement” (2024) 1 Example Law Review 2.\r")
	})
	t.Run("test_nested_quotation_in_title_stays_exact", func(t *testing.T) {
		sc := note(t, "A Author, “towards ‘a judgement’ in canada” (2024) 1 Example Law Review 2.", "")
		bound, err := alrmap.PlanSourceTitles(sc, "footnote:1", alrtest.Reg, "en", "refuse")
		if err != nil {
			t.Fatal(err)
		}
		out, _, err := ooxml.ApplyBound(sc, bound)
		if err != nil {
			t.Fatal(err)
		}
		if got := alrtest.Text(t, out, sc.Styles, sc.Part, 0); !strings.Contains(got, "“Towards ‘a judgement’ in Canada”") {
			t.Fatalf("got %q", got)
		}
	})
	t.Run("test_prose_quotation_is_not_a_title", func(t *testing.T) {
		sc := note(t, "She wrote “towards a judgement”.", "")
		if _, err := alrmap.PlanSourceTitles(sc, "footnote:1", alrtest.Reg, "en", "refuse"); err == nil {
			t.Fatal("prose quotation accepted as a title")
		}
	})
	t.Run("test_forged_lexical_edit_with_title_rule_refused", func(t *testing.T) {
		sc := note(t, "A Author, “towards a judgement” (2024) 1 Example Law Review 2.", "")
		c := sc.Containers[0]
		a := len([]rune(c.Text[:strings.Index(c.Text, "towards")]))
		e := kernels.MinimalEdit(c.Text, a, a+7, "toward", "TITLE-EN")
		p := kernels.Plan{Digest: kernels.Digest(c.Text), Edits: []kernels.Edit{*e}, Authorized: true, AuthorizationBasis: "fake-title-rule"}
		bound := &ooxml.BoundPlan{Fingerprint: sc.Fingerprint(), Container: c.Key, Family: "title", Plan: p, Registry: alrtest.Reg}
		if _, _, err := ooxml.ApplyBound(sc, bound); err == nil || !strings.Contains(err.Error(), "case changes only") {
			t.Fatalf("got %v", err)
		}
	})
	t.Run("test_case_change_outside_title_refused", func(t *testing.T) {
		sc := note(t, "A Author, “towards a judgement” (2024) 1 Example Law Review 2.", "")
		c := sc.Containers[0]
		e := kernels.MinimalEdit(c.Text, 0, 1, "a", "TITLE-EN")
		p := kernels.Plan{Digest: kernels.Digest(c.Text), Edits: []kernels.Edit{*e}, Authorized: true, AuthorizationBasis: "fake-range"}
		bound := &ooxml.BoundPlan{Fingerprint: sc.Fingerprint(), Container: c.Key, Family: "title", Plan: p, Registry: alrtest.Reg}
		if _, _, err := ooxml.ApplyBound(sc, bound); err == nil || !strings.Contains(err.Error(), "outside a parsed title") {
			t.Fatalf("got %v", err)
		}
	})
	t.Run("test_statute_title_not_in_general_title_policy", func(t *testing.T) {
		reg := &citations.Registry{Statutes: map[string]bool{"The example Act": true}}
		sc := note(t, "The example Act, SA 2024, c A-1, s 2.", "")
		bound, err := alrmap.PlanSourceTitles(sc, "footnote:1", reg, "en", "refuse")
		if err != nil {
			t.Fatal(err)
		}
		if len(bound.Plan.Edits) != 0 {
			t.Fatalf("edits = %v", bound.Plan.Edits)
		}
	})
	t.Run("test_case_dictionary_cannot_rename_a_source", func(t *testing.T) {
		reg := &citations.Registry{Cases: map[string][]string{"Alpha v Beta": {"Different v Source"}}}
		if _, err := citations.NormalizeNote("Alpha v Beta, 2024 SCC 1.", reg, citations.Options{}); err == nil {
			t.Fatal("case dictionary renamed a source")
		}
	})
}

func TestXMLCharacterDataBoundaryTests(t *testing.T) {
	t.Run("test_noncharacter_markup_in_text_nodes_refused", func(t *testing.T) {
		for _, inner := range []string{"one<!--keep-->two", "one<?keep value?>two"} {
			t.Run(inner, func(t *testing.T) {
				_, err := ooxml.ParseXML([]byte(`<w:t xmlns:w="` + ooxml.W + `">` + inner + `</w:t>`))
				if err == nil || !strings.Contains(err.Error(), "non-character XML") {
					t.Fatalf("got %v", err)
				}
			})
		}
	})
}
