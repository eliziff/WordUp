package kernels_test

import (
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap/kernels"
)

func short(t *testing.T, text string, opts kernels.ShortCitationOptions) string {
	t.Helper()
	r, err := kernels.NormalizeShortCitation(text, opts)
	if err != nil {
		t.Fatalf("NormalizeShortCitation(%q): %v", text, err)
	}
	return r.Text
}

func TestLocatorTests(t *testing.T) {
	t.Run("test_page_contraction_round_trip_and_boundary", func(t *testing.T) {
		for _, c := range []struct{ a, b, want string }{
			{"400", "408", "08"}, {"1847", "1857", "57"}, {"199", "201", "201"}, {"1000", "1005", "05"}, {"1000", "1015", "15"},
		} {
			t.Run(c.a+"-"+c.b, func(t *testing.T) {
				got, err := kernels.ShortenPageRange(c.a, c.b)
				if err != nil || got != c.want {
					t.Fatalf("ShortenPageRange(%q,%q) = %q, %v; want %q", c.a, c.b, got, err, c.want)
				}
			})
		}
	})
	t.Run("test_paragraph_not_page", func(t *testing.T) {
		mustEqual(t, short(t, "Ibid at paras 1847–1857.", kernels.ShortCitationOptions{}), "Ibid at paras 1847–1857.")
		mustEqual(t, short(t, "Ibid at pp. 1847-1857.", kernels.ShortCitationOptions{}), "Ibid at 1847–57.")
	})
	t.Run("test_compound_section_identifier_never_en_dash", func(t *testing.T) {
		mustEqual(t, short(t, "Ibid, s. 3-2.", kernels.ShortCitationOptions{}), "Ibid, s 3-2.")
		mustEqual(t, short(t, "Ibid, ss. 3-2 to 3-5.", kernels.ShortCitationOptions{}), "Ibid, ss 3-2 to 3-5.")
	})
	t.Run("test_para_label_cardinality", func(t *testing.T) {
		mustEqual(t, short(t, "Ibid at para. 5-8, 10.", kernels.ShortCitationOptions{}), "Ibid at paras 5–8, 10.")
		mustEqual(t, short(t, "Ibid at paras. 5.", kernels.ShortCitationOptions{}), "Ibid at para 5.")
	})
	t.Run("test_no_bare_at_type_guess", func(t *testing.T) {
		_, err := kernels.NormalizeShortCitation("Ibid at 400-408.", kernels.ShortCitationOptions{})
		wantValueError(t, err)
		mustEqual(t, short(t, "Ibid at 400-408.", kernels.ShortCitationOptions{BareAtMeansPage: true}), "Ibid at 400–08.")
	})
	t.Run("test_entire_production_must_parse", func(t *testing.T) {
		for _, s := range []string{"“Ibid at pp. 400-408.”", "The years 2023-2030 matter.", "Ibid at pp. 400-408; another source.", "Ibid at paras 3-2 to 3-5.", "Ibid at para 5 of the report."} {
			t.Run(s, func(t *testing.T) {
				_, err := kernels.NormalizeShortCitation(s, kernels.ShortCitationOptions{})
				wantValueError(t, err)
			})
		}
	})
	t.Run("test_exact_supra_alias_not_partial_name", func(t *testing.T) {
		s := "Smith, supra note 7 at p. 400-408."
		mustEqual(t, short(t, s, kernels.ShortCitationOptions{Aliases: map[string]bool{"Smith": true}}), "Smith, supra note 7 at 400–08.")
		_, err := kernels.NormalizeShortCitation(s, kernels.ShortCitationOptions{Aliases: map[string]bool{"Smithson": true}})
		wantValueError(t, err)
	})
	t.Run("test_italicization_targets_syntax_word_not_punctuation", func(t *testing.T) {
		r, err := kernels.NormalizeShortCitation("Ibid at para 5.", kernels.ShortCitationOptions{})
		if err != nil {
			t.Fatal(err)
		}
		if len(r.TokenFormat) != 1 || r.TokenFormat[0] != (kernels.TokenFormat{Start: 0, End: 4, Format: "italic"}) {
			t.Fatalf("token_format = %v", r.TokenFormat)
		}
	})
	t.Run("test_no_downgoing_range", func(t *testing.T) {
		_, err := kernels.NormalizeShortCitation("Ibid at pp. 408-400.", kernels.ShortCitationOptions{})
		wantValueError(t, err)
	})
}
