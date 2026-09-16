package citations_test

import (
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap/alrtest"
	"github.com/eliziff/WordUp/internal/alrmap/citations"
	"github.com/eliziff/WordUp/internal/alrmap/ooxml"
)

// The three CitationGrammar tests that bind parsed edits to actual XML.
func TestCitationGrammarXML(t *testing.T) {
	t.Run("test_citation_applied_to_actual_xml_preserves_title", func(t *testing.T) {
		text := "A Author, “A Title; Towards 20%” (2024) 61:1 Example Law Review 400 at pp. 401-408."
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run(text, "", true), ""), nil)
		c := sc.Containers[0]
		result := normalize(t, text, alrtest.Reg, citations.Options{})
		bound, err := ooxml.Bind(sc, c.Key, result.Edits, "citation", "full synthetic citation grammar")
		if err != nil {
			t.Fatal(err)
		}
		out, _, err := ooxml.ApplyBound(sc, bound)
		if err != nil {
			t.Fatal(err)
		}
		if got := alrtest.Text(t, out, sc.Styles, sc.Part, 0); got != result.Text+"\r" {
			t.Fatalf("got %q", got)
		}
	})
	t.Run("test_field_result_digits_are_not_text", func(t *testing.T) {
		text := "Ibid at pp. 400-408."
		content := alrtest.Run("Ibid at pp. ", "", true) + `<w:fldSimple w:instr="REF book">` + alrtest.Run("400-408", "", true) + `</w:fldSimple>` + alrtest.Run(".", "", true)
		sc := alrtest.Project(t, alrtest.Para(content, ""), nil)
		c := sc.Containers[0]
		result := normalize(t, text, nil, citations.Options{})
		if _, err := ooxml.Bind(sc, c.Key, result.Edits, "citation", "syntax"); err == nil {
			t.Fatal("field result digits were treated as text")
		}
	})
	t.Run("test_terminal_repair_on_text_atom", func(t *testing.T) {
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run("Ibid", "", true), ""), nil)
		c := sc.Containers[0]
		result := normalize(t, "Ibid", nil, citations.Options{})
		bound, err := ooxml.Bind(sc, c.Key, result.Edits, "citation", "syntax")
		if err != nil {
			t.Fatal(err)
		}
		out, _, err := ooxml.ApplyBound(sc, bound)
		if err != nil {
			t.Fatal(err)
		}
		if got := alrtest.Text(t, out, sc.Styles, sc.Part, 0); got != "Ibid.\r" {
			t.Fatalf("got %q", got)
		}
	})
}
