package alrmap_test

import (
	"bytes"
	"errors"
	"reflect"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap"
	"github.com/eliziff/WordUp/internal/alrmap/alrtest"
	"github.com/eliziff/WordUp/internal/alrmap/citations"
	"github.com/eliziff/WordUp/internal/alrmap/kernels"
	"github.com/eliziff/WordUp/internal/alrmap/ooxml"
)

func authored(t *testing.T, text, authority string) string {
	t.Helper()
	p, err := alrmap.PlanAuthorText(text, alrmap.AuthorTextOptions{Authority: authority})
	if err != nil {
		t.Fatal(err)
	}
	out, err := kernels.ApplyPlan(text, p)
	if err != nil {
		t.Fatal(err)
	}
	return out
}

func title(t *testing.T, text string, opts alrmap.TitleOptions) alrmap.TitleResult {
	t.Helper()
	r, err := alrmap.TitleCase(text, opts)
	if err != nil {
		t.Fatalf("TitleCase(%q): %v", text, err)
	}
	return r
}

func TestFiniteRules(t *testing.T) {
	t.Run("test_expanded_spelling_not_substrings", func(t *testing.T) {
		got := authored(t, "honor honorary honorific defense defensive cell phone by-law e-mail", "selected")
		if got != "honour honorary honorific defence defensive cellphone bylaw email" {
			t.Fatalf("got %q", got)
		}
	})
	t.Run("test_spelling_stays_outside_sources", func(t *testing.T) {
		got := authored(t, "behavior “behavior” judgement https://example.ca/judgement [behavior]", "selected")
		if got != "behaviour “behavior” judgment https://example.ca/judgement [behavior]" {
			t.Fatalf("got %q", got)
		}
	})
	t.Run("test_unauthorized_still_advisory", func(t *testing.T) {
		p, err := alrmap.PlanAuthorText("towards", alrmap.AuthorTextOptions{})
		if err != nil {
			t.Fatal(err)
		}
		_, err = kernels.ApplyPlan("towards", p)
		var perm *kernels.PermissionError
		if !errors.As(err, &perm) {
			t.Fatalf("got %v", err)
		}
	})
	t.Run("test_title_hyphen_colon_verbs_and_four_letter_words", func(t *testing.T) {
		r := title(t, "over and from: what is treaty-making with practice", alrmap.TitleOptions{Language: "en"})
		if r.Text != "Over and From: What Is Treaty-Making with Practice" {
			t.Fatalf("got %q", r.Text)
		}
	})
	t.Run("test_title_case_preserves_unknown_acronyms", func(t *testing.T) {
		r := title(t, "an ALR case about iPhone", alrmap.TitleOptions{Language: "en"})
		if r.Text != "An ALR Case About iPhone" || !reflect.DeepEqual(r.Unresolved, []string{"ALR", "iPhone"}) {
			t.Fatalf("got %q, unresolved %v", r.Text, r.Unresolved)
		}
	})
	t.Run("test_title_with_conflict_and_language", func(t *testing.T) {
		if _, err := alrmap.TitleCase("the title: with others", alrmap.TitleOptions{Language: "en"}); err == nil {
			t.Fatal("guide conflict not refused")
		}
		if r := title(t, "the title: with others", alrmap.TitleOptions{Language: "en", WithAfterColon: "lower"}); r.Text != "The Title: with Others" {
			t.Fatalf("got %q", r.Text)
		}
		if _, err := alrmap.TitleCase("le droit de Paris", alrmap.TitleOptions{Language: "fr"}); err == nil {
			t.Fatal("French policy accepted")
		}
	})
	t.Run("test_exact_case_cannot_replace_identity", func(t *testing.T) {
		if _, err := alrmap.TitleCase("title", alrmap.TitleOptions{Language: "en", ExactCase: map[string]string{"title": "different"}}); err == nil {
			t.Fatal("identity change accepted")
		}
	})
	t.Run("test_selected_counts_ratios_ordinals", func(t *testing.T) {
		for _, c := range []struct{ raw, role, want string }{{"7", "count", "seven"}, {"1986", "count", "1,986"},
			{"21st", "ordinal", "twenty-first"}, {"20th", "ordinal", "twentieth"},
			{"5:1", "ratio", "5 to 1"}, {"2-4", "count-range", "two to four"}} {
			got, err := alrmap.SelectedNumber(c.raw, c.role)
			if err != nil || got != c.want {
				t.Fatalf("SelectedNumber(%q,%q) = %q, %v", c.raw, c.role, got, err)
			}
		}
		for _, c := range []struct{ raw, role string }{{"11st", "ordinal"}, {"01", "count"}, {"5-2", "count-range"}, {"7", "unknown"}} {
			if _, err := alrmap.SelectedNumber(c.raw, c.role); err == nil {
				t.Fatalf("SelectedNumber(%q,%q) accepted", c.raw, c.role)
			}
		}
	})
	t.Run("test_selected_section_keeps_identifier", func(t *testing.T) {
		if got, err := alrmap.SelectedLegalReference("subparagraph 4(1)(h)", "prose", "statute"); err != nil || got != "section 4(1)(h)" {
			t.Fatalf("got %q, %v", got, err)
		}
		if got, err := alrmap.SelectedLegalReference("ss. 3-2 to 3-5", "citation", "statute"); err != nil || got != "ss 3-2 to 3-5" {
			t.Fatalf("got %q, %v", got, err)
		}
		if _, err := alrmap.SelectedLegalReference("subparagraph g", "prose", "statute"); err == nil {
			t.Fatal("incomplete reference accepted")
		}
	})
	t.Run("test_question_exclamation_never_moved", func(t *testing.T) {
		for _, suffix := range []string{"?", "!"} {
			edits, err := alrmap.QuoteSeparator("Did she say “go”"+suffix, kernels.Span{Start: 12, End: 16, Reason: "quote"}, "prose-quotation")
			if err != nil || len(edits) != 0 {
				t.Fatalf("got %v, %v", edits, err)
			}
		}
	})
	t.Run("test_typed_quote_boundary_only", func(t *testing.T) {
		text := "She said “go”, then left."
		edits, err := alrmap.QuoteSeparator(text, kernels.Span{Start: 9, End: 13, Reason: "quote"}, "prose-quotation")
		if err != nil || len(edits) == 0 {
			t.Fatalf("got %v, %v", edits, err)
		}
		e := edits[0]
		runes := []rune(text)
		if got := string(runes[:e.Start]) + e.New + string(runes[e.End:]); got != "She said “go,” then left." {
			t.Fatalf("got %q", got)
		}
		if _, err := alrmap.QuoteSeparator(text, kernels.Span{Start: 9, End: 13, Reason: "title"}, "citation-title"); err == nil {
			t.Fatal("untyped role accepted")
		}
	})
	t.Run("test_qualifications_do_not_infer_provenance", func(t *testing.T) {
		if got, err := alrmap.QualificationFlags([]string{"emphasis added", "footnotes omitted"}); err != nil || got != "[emphasis added, footnotes omitted]" {
			t.Fatalf("got %q, %v", got, err)
		}
		if _, err := alrmap.QualificationFlags([]string{"emphasis added", "emphasis in original"}); err == nil {
			t.Fatal("combined emphasis accepted")
		}
	})
}

func TestAddedGrammarEdges(t *testing.T) {
	t.Run("test_reporter_abbreviation_and_pinpoint_only", func(t *testing.T) {
		r := &citations.Registry{Reporters: map[string][]string{"S.C.R.": {"SCR"}, "SCR": {"SCR"}}}
		result, err := citations.NormalizeNote("[1993] 3 S.C.R. 519 at pp. 520-528 (SCC).", r, citations.Options{})
		if err != nil || result.Text != "[1993] 3 SCR 519 at 520–28 (SCC)." {
			t.Fatalf("got %q, %v", result.Text, err)
		}
		again, err := citations.NormalizeNote(result.Text, r, citations.Options{BasisHints: []string{"page"}})
		if err != nil || result.Citations[0].Meaning() != again.Citations[0].Meaning() {
			t.Fatalf("meaning changed: %v", err)
		}
	})
	t.Run("test_reporter_series_position", func(t *testing.T) {
		r := &citations.Registry{Reporters: map[string][]string{"D.L.R.": {"DLR"}, "DLR": {"DLR"}}}
		result, err := citations.NormalizeNote("1960, 23 D.L.R. (2d) 689 at pp. 690-699.", r, citations.Options{})
		if err != nil || result.Text != "1960, 23 DLR (2d) 689 at 690–99." {
			t.Fatalf("got %q, %v", result.Text, err)
		}
	})
	t.Run("test_legislative_identifier_hyphens_unchanged", func(t *testing.T) {
		for _, c := range [][2]string{{"RSC 1985, c C-46, s. 718.", "RSC 1985, c C-46, s 718."},
			{"RSA 2000, c L-7, ss. 3-2 to 3-5.", "RSA 2000, c L-7, ss 3-2 to 3-5."},
			{"SOR/94-688, s. 3.", "SOR/94-688, s 3."}} {
			result, err := citations.NormalizeNote(c[0], nil, citations.Options{})
			if err != nil || result.Text != c[1] {
				t.Fatalf("got %q, %v; want %q", result.Text, err, c[1])
			}
		}
	})
	t.Run("test_semicolon_inside_url_not_a_citation_separator", func(t *testing.T) {
		text := "A Author, “Title” (May 2, 2024), online: <https://example.ca/a;b?q=1>; Ibid at para. 7."
		result, err := citations.NormalizeNote(text, nil, citations.Options{})
		if err != nil || result.Text != "A Author, “Title” (2 May 2024), online: <https://example.ca/a;b?q=1>; Ibid at para 7." {
			t.Fatalf("got %q, %v", result.Text, err)
		}
	})
	t.Run("test_leaf_locator_preserves_different_number_formatting", func(t *testing.T) {
		content := alrtest.Run("Ibid at para. ", "", true) + alrtest.Run("400", `<w:color w:val="123456"/>`, true) + alrtest.Run("-", "", true) +
			alrtest.Run("408", `<w:sz w:val="18"/>`, true) + alrtest.Run(".", "", true)
		sc := alrtest.Project(t, alrtest.Para(content, ""), nil)
		c := sc.Containers[0]
		result, err := citations.NormalizeNote("Ibid at para. 400-408.", nil, citations.Options{})
		if err != nil {
			t.Fatal(err)
		}
		bound, err := ooxml.Bind(sc, c.Key, result.Edits, "citation", "synthetic full citation")
		if err != nil {
			t.Fatal(err)
		}
		out, _, err := ooxml.ApplyBound(sc, bound)
		if err != nil {
			t.Fatal(err)
		}
		if got := alrtest.Text(t, out, sc.Styles, sc.Part, 0); got != "Ibid at paras 400–408.\r" {
			t.Fatalf("got %q", got)
		}
		if !bytes.Contains(out, []byte(alrtest.Run("400", `<w:color w:val="123456"/>`, true))) || !bytes.Contains(out, []byte(alrtest.Run("408", `<w:sz w:val="18"/>`, true))) {
			t.Fatalf("original runs not preserved: %s", out)
		}
	})
	t.Run("test_original_occurrence_not_another_equal_token", func(t *testing.T) {
		text := "“300-308” then 300-308"
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run(text, "", true), ""), nil)
		c := sc.Containers[0]
		start := len([]rune(text[:strings.LastIndex(text, "300-308")]))
		e := kernels.MinimalEdit(c.Text, start, start+7, "300-08", "SELECTED-PAGE")
		bound, err := ooxml.Bind(sc, c.Key, []kernels.Edit{*e}, "citation", "explicit synthetic page selection")
		if err != nil {
			t.Fatal(err)
		}
		out, _, err := ooxml.ApplyBound(sc, bound)
		if err != nil {
			t.Fatal(err)
		}
		if got := alrtest.Text(t, out, sc.Styles, sc.Part, 0); got != "“300-308” then 300-08\r" {
			t.Fatalf("got %q", got)
		}
	})
	t.Run("test_batch_part_is_same_as_independent_noninterfering_patches", func(t *testing.T) {
		raw := `<w:footnotes xmlns:w="` + ooxml.W + `">`
		for _, i := range []string{"1", "2"} {
			raw += `<w:footnote w:id="` + i + `">` + alrtest.Para(alrtest.Run("Ibid at para. 5-8.", "", true), "") + `</w:footnote>`
		}
		raw += `</w:footnotes>`
		sc, err := ooxml.ScanPart([]byte(raw), alrtest.Styles(t, "", ""), "word/footnotes.xml")
		if err != nil {
			t.Fatal(err)
		}
		var bounds []*ooxml.BoundPlan
		for _, c := range sc.Containers {
			bound, _, err := ooxml.PlanNote(sc, c.Key, nil, nil)
			if err != nil {
				t.Fatal(err)
			}
			bounds = append(bounds, bound)
		}
		out, writes, err := ooxml.ApplyBatch(sc, bounds)
		if err != nil {
			t.Fatal(err)
		}
		after, err := ooxml.ScanPart(out, sc.Styles, sc.Part)
		if err != nil {
			t.Fatal(err)
		}
		for _, c := range after.Containers {
			if c.Text != "Ibid at paras 5–8.\r" {
				t.Fatalf("container %s = %q", c.Key, c.Text)
			}
		}
		if len(writes) != 2 {
			t.Fatalf("writes = %d", len(writes))
		}
		if _, _, err := ooxml.ApplyBatch(sc, []*ooxml.BoundPlan{bounds[0], bounds[0]}); err == nil {
			t.Fatal("same-container plans accepted")
		}
	})
}
