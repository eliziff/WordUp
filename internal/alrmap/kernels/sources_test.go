package kernels_test

import (
	"reflect"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap/kernels"
)

func ibid(t *testing.T, previous, current kernels.ResolvedCitation, single, adjacent bool) (string, bool) {
	t.Helper()
	return kernels.IbidPayload(previous, current, single, adjacent)
}

func TestIdentityAndPropertyTests(t *testing.T) {
	base := kernels.ResolvedCitation{SourceID: "book1", EditionID: "ed3", LocatorBasis: "page", Locator: []kernels.Range{{9, 9}}}
	t.Run("test_longest_phrase_blocks_shorter_italic_match", func(t *testing.T) {
		s := "ultra vires; vires; obiter dictum; dictum"
		rules := map[string]bool{"vires": true, "ultra vires": false, "obiter dictum": true, "dictum": false}
		got, err := kernels.PhraseItalics(s, rules, nil)
		if err != nil {
			t.Fatal(err)
		}
		type pair struct {
			text string
			v    bool
		}
		var pairs []pair
		runes := []rune(s)
		for _, it := range got {
			pairs = append(pairs, pair{string(runes[it.Start:it.End]), it.Italic})
		}
		want := []pair{{"ultra vires", false}, {"vires", true}, {"obiter dictum", true}, {"dictum", false}}
		if !reflect.DeepEqual(pairs, want) {
			t.Fatalf("got %v, want %v", pairs, want)
		}
	})
	t.Run("test_lexical_italics_do_not_touch_quote", func(t *testing.T) {
		got, err := kernels.PhraseItalics("“vires”", map[string]bool{"vires": true}, nil)
		if err != nil || len(got) != 0 {
			t.Fatalf("got %v, %v", got, err)
		}
	})
	t.Run("test_ibid_same_locator_can_be_omitted", func(t *testing.T) {
		if got, ok := ibid(t, base, base, true, true); !ok || got != "Ibid" {
			t.Fatalf("got %q, %v", got, ok)
		}
	})
	t.Run("test_ibid_sourcewide_does_not_inherit_wrong_pinpoint", func(t *testing.T) {
		b := base
		b.Locator = nil
		if got, ok := ibid(t, base, b, true, true); ok {
			t.Fatalf("got %q", got)
		}
	})
	t.Run("test_ibid_unknown_identity_never_equal_by_accident", func(t *testing.T) {
		for _, c := range []struct{ source, edition string }{{"", "ed3"}, {"book1", ""}, {" ", "ed3"}} {
			a := kernels.ResolvedCitation{SourceID: c.source, EditionID: c.edition, LocatorBasis: "page", Locator: []kernels.Range{{9, 9}}}
			if got, ok := ibid(t, a, a, true, true); ok {
				t.Fatalf("got %q for %+v", got, c)
			}
		}
	})
	t.Run("test_ibid_equal_invalid_locators_not_elided", func(t *testing.T) {
		// The reference also checks a ((True, True),) locator; Go's typed
		// Range cannot express a non-integer endpoint, so that value is
		// unrepresentable rather than tested.
		for _, value := range [][]kernels.Range{{}, {{0, 0}}, {{9, 8}}} {
			a := kernels.ResolvedCitation{SourceID: "book1", EditionID: "ed3", LocatorBasis: "page", Locator: value}
			if got, ok := ibid(t, a, a, true, true); ok {
				t.Fatalf("got %q for %v", got, value)
			}
		}
	})
	t.Run("test_perma_mapping_target_is_checked_at_token_boundary", func(t *testing.T) {
		_, _, err := kernels.PermaTokenChange("https://example.ca/a", map[string]string{"https://example.ca/a": "https://example.ca/not-perma"})
		wantValueError(t, err)
	})
	t.Run("test_ibid_changed_pinpoint_retained", func(t *testing.T) {
		b := base
		b.Locator = []kernels.Range{{12, 12}}
		if got, ok := ibid(t, base, b, true, true); !ok || got != "Ibid at 12" {
			t.Fatalf("got %q, %v", got, ok)
		}
	})
	t.Run("test_ibid_multiple_sources_version_or_adjacency_denied", func(t *testing.T) {
		if got, ok := ibid(t, base, base, false, true); ok {
			t.Fatalf("got %q", got)
		}
		if got, ok := ibid(t, base, base, true, false); ok {
			t.Fatalf("got %q", got)
		}
		other := base
		other.EditionID = "ed4"
		if got, ok := ibid(t, base, other, true, true); ok {
			t.Fatalf("got %q", got)
		}
	})
	t.Run("test_perma_exact_complete_token_not_prefix", func(t *testing.T) {
		d := map[string]string{"https://example.ca/A?q=1": "https://perma.cc/ABCD-1234"}
		got, ok, err := kernels.PermaTokenChange("https://example.ca/A?q=1", d)
		if err != nil || !ok || got != "[perma.cc/ABCD-1234]" {
			t.Fatalf("got %q, %v, %v", got, ok, err)
		}
		if _, ok, err := kernels.PermaTokenChange("https://example.ca/A?q=10", d); ok || err != nil {
			t.Fatalf("prefix token matched: %v %v", ok, err)
		}
		if _, ok, err := kernels.PermaTokenChange("https://example.ca/a?q=1", d); ok || err != nil {
			t.Fatalf("case-folded token matched: %v %v", ok, err)
		}
	})
	t.Run("test_csv_record_newline_not_line_split", func(t *testing.T) {
		s := "submitted_url,status,perma_url,title\r\nhttps://example.ca/a,success,https://perma.cc/ABCD-1234,\"a title\nwith a newline\"\r\n"
		got, err := kernels.PermaMapping(s)
		if err != nil || !reflect.DeepEqual(got, map[string]string{"https://example.ca/a": "https://perma.cc/ABCD-1234"}) {
			t.Fatalf("got %v, %v", got, err)
		}
	})
	t.Run("test_perma_conflicting_duplicate_denied", func(t *testing.T) {
		s := "submitted_url,status,perma_url\nhttps://example.ca/a,success,https://perma.cc/ABCD-1234\nhttps://example.ca/a,success,https://perma.cc/EFGH-5678\n"
		_, err := kernels.PermaMapping(s)
		wantValueError(t, err)
	})
}

func TestLegacyCounterexampleTests(t *testing.T) {
	t.Run("test_legacy_find_relocates_approved_match_to_protected_title", func(t *testing.T) {
		// A source-shaped reduction of the actual Pair 085/id35 and 105/id90 bug.
		s := "“Pricing 2023-2030”, online: https://canada.ca/federal-benchmark-2023-2030.html."
		runes := []rune(s)
		approved := len([]rune(s[:strings.LastIndex(s, "2023-2030")]))
		actual := len([]rune(s[:strings.Index(s, string(runes[approved:approved+9]))]))
		if approved == actual {
			t.Fatal("legacy find did not relocate")
		}
		hit := false
		for _, g := range kernels.QuotationGuards(s) {
			if g.Overlaps(actual, actual+9) {
				hit = true
			}
		}
		if !hit {
			t.Fatal("relocated match is not inside the protected title")
		}
		result := string(runes[:actual+5]) + string(runes[actual+7:])
		if !strings.Contains(result, "“Pricing 2023-30”") {
			t.Fatalf("result %q", result)
		}
	})
	t.Run("test_raw_unmarked_quote_is_indistinguishable_not_a_safety_claim", func(t *testing.T) {
		// Exactly the same bytes can be the author's prose or an unmarked quote.
		s := "We move towards a 20% threshold."
		p, err := kernels.PlanProse(s, kernels.ProseOptions{})
		if err != nil {
			t.Fatal(err)
		}
		if len(p.Edits) == 0 || p.Authorized {
			t.Fatalf("plan %+v", p)
		}
		_, err = kernels.ApplyPlan(s, p)
		wantPermissionError(t, err)
	})
}
