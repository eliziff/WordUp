package citations_test

import (
	"errors"
	"fmt"
	"sort"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap/citations"
	"github.com/eliziff/WordUp/internal/alrmap/kernels"
)

// reg mirrors the shared REG fixture of the reference suite.
var reg = &citations.Registry{
	Journals: map[string][]string{"Example Law Review": {"Ex L Rev"}, "Ex L Rev": {"Ex L Rev"}},
	Cases:    map[string][]string{"R. v. Example Ltd.": {"R v Example Ltd"}, "R v Example Ltd": {"R v Example Ltd"}},
	Aliases:  map[string]bool{"Example": true, "Author, Treatise": true},
	Books:    map[string]bool{"Towards Judgement": true},
}

func normalize(t *testing.T, text string, r *citations.Registry, opts citations.Options) citations.Result {
	t.Helper()
	result, err := citations.NormalizeNote(text, r, opts)
	if err != nil {
		t.Fatalf("NormalizeNote(%q): %v", text, err)
	}
	return result
}

func meanings(cs []citations.Citation) []string {
	out := make([]string, len(cs))
	for i, c := range cs {
		out[i] = c.Meaning()
	}
	return out
}

func bases(cs []citations.Citation) []string {
	out := make([]string, len(cs))
	for i, c := range cs {
		out[i] = c.Basis
	}
	return out
}

func wantValueError(t *testing.T, err error) {
	t.Helper()
	var perm *kernels.PermissionError
	if err == nil || errors.As(err, &perm) {
		t.Fatalf("expected a value error, got %v", err)
	}
}

// checkRoundtrip mirrors CitationGrammar.check_roundtrip.
func checkRoundtrip(t *testing.T, text, expected string) citations.Result {
	t.Helper()
	result := normalize(t, text, reg, citations.Options{})
	if result.Text != expected {
		t.Fatalf("normalize(%q) = %q, want %q", text, result.Text, expected)
	}
	again := normalize(t, result.Text, reg, citations.Options{BasisHints: bases(result.Citations)})
	if again.Text != result.Text {
		t.Fatalf("second pass changed %q to %q", result.Text, again.Text)
	}
	if strings.Join(meanings(result.Citations), "\n") != strings.Join(meanings(again.Citations), "\n") {
		t.Fatalf("meaning changed on reparse: %v vs %v", meanings(result.Citations), meanings(again.Citations))
	}
	runes := []rune(text)
	for _, e := range result.Edits {
		if string(runes[e.Start:e.End]) != e.Old {
			t.Fatalf("edit %+v does not address its old text", e)
		}
	}
	return result
}

func TestCitationGrammar(t *testing.T) {
	t.Run("test_pages_paragraphs_sections_distinct", func(t *testing.T) {
		for _, c := range [][2]string{
			{"Ibid at pp. 400-408.", "Ibid at 400–08."},
			{"Ibid at paras 1847–1857.", "Ibid at paras 1847–1857."},
			{"Ibid, ss. 3-2 to 3-5.", "Ibid, ss 3-2 to 3-5."},
			{"Ibid at para. 5-8 and 10.", "Ibid at paras 5–8, 10."},
		} {
			checkRoundtrip(t, c[0], c[1])
		}
	})
	t.Run("test_canlii_order_owns_only_court_and_locator", func(t *testing.T) {
		checkRoundtrip(t, "R. v. Example Ltd., 1999 CanLII 42 (ONCA) at para. 4-8.", "R v Example Ltd, 1999 CanLII 42 at paras 4–8 (ONCA).")
	})
	t.Run("test_multi_citation_nested_title", func(t *testing.T) {
		text := "A Author, “A Title; Towards 20%” (2024) 61:1 Example Law Review 400 at pp. 401-408; 2024 SCC 7 at paras. 5."
		checkRoundtrip(t, text, "A Author, “A Title; Towards 20%” (2024) 61:1 Ex L Rev 400 at 401–08; 2024 SCC 7 at para 5.")
	})
	t.Run("test_book_title_is_not_prose", func(t *testing.T) {
		checkRoundtrip(t, "A Author, Towards Judgement, 2nd ed (Town: Press, 2024) at pp. 401-408.",
			"A Author, Towards Judgement, 2nd ed (Town: Press, 2024) at 401–08.")
	})
	t.Run("test_web_date_not_date_inside_title", func(t *testing.T) {
		checkRoundtrip(t, "A Author, “Towards May 1, 2024” (May 2, 2024), online (blog): [perma.cc/ABCD-1234].",
			"A Author, “Towards May 1, 2024” (2 May 2024), online (blog): [perma.cc/ABCD-1234].")
	})
	t.Run("test_qualifier_and_alias_retained", func(t *testing.T) {
		checkRoundtrip(t, "2024 SCC 7 at para. 5-8 [emphasis added, footnotes omitted] [Example].",
			"2024 SCC 7 at paras 5–8 [emphasis added, footnotes omitted] [Example].")
	})
	t.Run("test_named_and_unnamed_supra", func(t *testing.T) {
		for _, text := range []string{"Example, supra note 7 at para. 5.", "Supra note 7 at para. 5."} {
			checkRoundtrip(t, text, strings.ReplaceAll(text, "para.", "para"))
		}
	})
	t.Run("test_unknown_suffix_rejects_whole_note", func(t *testing.T) {
		for _, text := range []string{"Ibid at para. 5. This is a different judgement.",
			"Ibid at para. 5 [unknown].", "Ibid at para. 5; unsupported source.",
			"Ibid at para. 5 (discussion).", "Ibid at para. 5;"} {
			t.Run(text, func(t *testing.T) {
				_, err := citations.NormalizeNote(text, reg, citations.Options{})
				wantValueError(t, err)
			})
		}
	})
	t.Run("test_dictionary_collision_refused", func(t *testing.T) {
		registry := &citations.Registry{Journals: map[string][]string{"Review": {"A Rev", "B Rev"}}}
		_, err := citations.NormalizeNote("A Author, “Title” (2024) 1 Review 2.", registry, citations.Options{})
		wantValueError(t, err)
	})
	t.Run("test_alias_and_source_basis_not_guessed", func(t *testing.T) {
		for _, text := range []string{"Unknown, supra note 7 at para 5.", "Ibid at 400-408.", "Ibid at paras 1847-57.", "Ibid at para 9-2."} {
			t.Run(text, func(t *testing.T) {
				_, err := citations.NormalizeNote(text, reg, citations.Options{})
				wantValueError(t, err)
			})
		}
	})
	t.Run("test_all_disabled_is_exact_identity", func(t *testing.T) {
		text := "R. v. Example Ltd., 1999 CanLII 42 (ONCA) at para. 4-8."
		if got := normalize(t, text, reg, citations.Options{Enabled: kernels.NewSwitches()}).Text; got != text {
			t.Fatalf("got %q", got)
		}
	})
	t.Run("test_each_switch_is_independent", func(t *testing.T) {
		text := "Ibid at pp. 400 - 408 and 410."
		expected := map[string]string{"labels": "Ibid at 400 - 408 and 410.",
			"range-dashes":     "Ibid at pp. 400–408 and 410.",
			"page-contraction": "Ibid at pp. 400 - 08 and 410.",
			"pinpoint-and":     "Ibid at pp. 400 - 408, 410."}
		for _, sw := range citations.AllSwitches {
			want, ok := expected[sw]
			if !ok {
				want = text
			}
			if got := normalize(t, text, reg, citations.Options{Enabled: kernels.NewSwitches(sw)}).Text; got != want {
				t.Fatalf("switch %s: got %q, want %q", sw, got, want)
			}
		}
	})
	t.Run("test_all_switch_subsets_reparse_same_meaning", func(t *testing.T) {
		text := "R. v. Example Ltd., 1999 CanLII 42 (ONCA) at para. 400-408 and 410."
		switches := append([]string(nil), citations.AllSwitches...)
		sort.Strings(switches)
		meaning := normalize(t, text, reg, citations.Options{Enabled: kernels.NewSwitches()}).Citations[0].Meaning()
		// itertools.product((False, True), repeat=10): every one of the 1,024
		// subsets, enumerated as bit patterns over the sorted switch names.
		for bits := 0; bits < 1<<len(switches); bits++ {
			var names []string
			for i, name := range switches {
				if bits&(1<<i) != 0 {
					names = append(names, name)
				}
			}
			t.Run(fmt.Sprintf("subset_%04d", bits), func(t *testing.T) {
				result := normalize(t, text, reg, citations.Options{Enabled: kernels.NewSwitches(names...)})
				again := normalize(t, result.Text, reg, citations.Options{Enabled: kernels.NewSwitches(), BasisHints: []string{result.Citations[0].Basis}})
				if again.Citations[0].Meaning() != meaning {
					t.Fatalf("subset %v: %q reparses to %q, want %q", names, result.Text, again.Citations[0].Meaning(), meaning)
				}
			})
		}
	})
	t.Run("test_case_identifier_never_becomes_name_guess", func(t *testing.T) {
		for _, text := range []string{"Unknown v Unknown, 2024 SCC 7 at para. 5.", "2024 UNKNOWN 7 at para. 5.", "2024 SCC 7 (ONCA) at para. 5."} {
			_, err := citations.NormalizeNote(text, reg, citations.Options{})
			wantValueError(t, err)
		}
	})
}

func TestCanonicalDictionaryTests(t *testing.T) {
	t.Run("test_canonical_outputs_need_no_duplicate_registry_entries", func(t *testing.T) {
		r := &citations.Registry{Journals: map[string][]string{"Example Law Review": {"Ex L Rev"}},
			Cases: map[string][]string{"Alpha v. Beta": {"Alpha v Beta"}}, Reporters: map[string][]string{"S.C.R.": {"SCR"}}}
		for _, raw := range []string{"A Author, “Title” (2024) 1 Example Law Review 2.", "Alpha v. Beta, 2024 SCC 1.", "[2024] 1 S.C.R. 2."} {
			t.Run(raw, func(t *testing.T) {
				first := normalize(t, raw, r, citations.Options{})
				again := normalize(t, first.Text, r, citations.Options{})
				if first.Text != again.Text {
					t.Fatalf("%q then %q", first.Text, again.Text)
				}
				if strings.Join(meanings(first.Citations), "\n") != strings.Join(meanings(again.Citations), "\n") {
					t.Fatal("meaning changed")
				}
				if len(again.Edits) != 0 {
					t.Fatalf("canonical text still has edits: %v", again.Edits)
				}
			})
		}
	})
	t.Run("test_chained_or_cyclic_dictionary_refused", func(t *testing.T) {
		for i, mapping := range []map[string][]string{
			{"Example Law Review": {"Ex L Rev"}, "Ex L Rev": {"Other"}},
			{"Example Law Review": {"Ex L Rev"}, "Ex L Rev": {"Example Law Review"}},
		} {
			t.Run(fmt.Sprint(i), func(t *testing.T) {
				_, err := citations.NormalizeNote("A Author, “Title” (2024) 1 Example Law Review 2.", &citations.Registry{Journals: mapping}, citations.Options{})
				if err == nil || !strings.Contains(err.Error(), "not terminal") {
					t.Fatalf("got %v", err)
				}
			})
		}
	})
}

func TestSpeculativeRootTests(t *testing.T) {
	t.Run("test_failed_case_or_statute_prefix_cannot_own_journal_author", func(t *testing.T) {
		raw := "A. Author, “Title” (2024) 1 Example Law Review 2."
		for name, r := range map[string]*citations.Registry{
			"cases":    {Journals: map[string][]string{"Example Law Review": {"Example Law Review"}}, Cases: map[string][]string{"A. Author": {"A Author"}}},
			"statutes": {Journals: map[string][]string{"Example Law Review": {"Example Law Review"}}, Statutes: map[string]bool{"A. Author": true}},
		} {
			t.Run(name, func(t *testing.T) {
				result := normalize(t, raw, r, citations.Options{})
				if result.Text != raw {
					t.Fatalf("got %q", result.Text)
				}
				var titles []string
				for _, tok := range result.Citations[0].Tokens {
					if tok.Kind == "source-title" {
						titles = append(titles, tok.Value)
					}
				}
				if len(titles) != 1 || titles[0] != "“Title”" {
					t.Fatalf("source-title tokens = %v", titles)
				}
			})
		}
	})
}
