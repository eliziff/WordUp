package kernels_test

import (
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap/kernels"
)

// These cases are not part of the 149 reference tests. They pin the manual
// boundary checks that replace the Python lookaround assertions
// ((?<!\w), (?!\w), (?<![\w.,+\-]), (?![%\w]), \b, (?<=\S), (?=\S) and the
// OPAQUE (?<![\w@]) lookbehind) so each RE2 substitution is exercised on
// both sides of its boundary.
func TestLookaroundSubstitutions(t *testing.T) {
	t.Run("date_word_boundaries", func(t *testing.T) {
		for in, want := range map[string]string{
			"xMay 1, 2024":   "xMay 1, 2024",
			"May 1, 2024x":   "May 1, 2024x",
			"éMay 1, 2024":  "éMay 1, 2024",
			"_May 1, 2024":   "_May 1, 2024",
			"(May 1, 2024)":  "(1 May 2024)",
			"on may 1 2024.": "on 1 May 2024.",
			"1May 1, 2024":   "1May 1, 2024",
		} {
			mustEqual(t, edit(t, in, kernels.ProseOptions{Enabled: kernels.NewSwitches("date")}), want)
		}
	})
	t.Run("percent_boundaries", func(t *testing.T) {
		for in, want := range map[string]string{
			".20%": ".20%", ",20%": ",20%", "+20%": "+20%", "-20%": "-20%", "−20%": "−20%",
			"a20%": "a20%", "_20%": "_20%", "20%%": "20%%", "20%x": "20%x", "20%_": "20%_",
			"(20%)": "(20 percent)", "20 %.": "20 percent.", "é20%": "é20%", "20%é": "20%é",
		} {
			mustEqual(t, edit(t, in, kernels.ProseOptions{Enabled: kernels.NewSwitches("percent")}), want)
		}
	})
	t.Run("spaces_boundaries", func(t *testing.T) {
		for in, want := range map[string]string{
			"a  b": "a b", "  a": "  a", "a  ": "a  ", "a\t  b": "a\t  b", "a   b": "a   b",
			"a  \tb": "a  \tb", "a   b   c": "a b c", "a  b  ": "a b  ",
		} {
			mustEqual(t, edit(t, in, kernels.ProseOptions{Enabled: kernels.NewSwitches("spaces")}), want)
		}
	})
	t.Run("spelling_word_boundaries", func(t *testing.T) {
		for in, want := range map[string]string{
			"étowards": "étowards", "towardsé": "towardsé", "towards-x": "towards-x", "x/towards": "x/towards",
			"(towards)": "(toward)", "TOWARDS": "TOWARD", "Towards": "Toward", "tOwards": "tOwards",
			"1towards": "1towards", "towards_": "towards_", "caselaw.": "case law.",
		} {
			mustEqual(t, edit(t, in, kernels.ProseOptions{Enabled: kernels.NewSwitches("spelling")}), want)
		}
	})
	t.Run("opaque_naked_host_lookbehind", func(t *testing.T) {
		urls := func(text string) []kernels.Span {
			guards, err := kernels.Protection(text, nil)
			if err != nil {
				t.Fatal(err)
			}
			var out []kernels.Span
			for _, g := range guards {
				if g.Reason == "url-or-email" {
					out = append(out, g)
				}
			}
			return out
		}
		check := func(text string, want []kernels.Span) {
			t.Helper()
			got := urls(text)
			if len(got) != len(want) {
				t.Fatalf("%q: got %v, want %v", text, got, want)
			}
			for i := range got {
				if got[i] != want[i] {
					t.Fatalf("%q: got %v, want %v", text, got, want)
				}
			}
		}
		// A host preceded by a word character or '@' is skipped, but Python
		// resumes one position later and still finds the host after the dot.
		check("@sub.canada.ca/a x", []kernels.Span{{Start: 5, End: 16, Reason: "url-or-email"}})
		check("xcanada.ca/a", []kernels.Span{{Start: 0, End: 12, Reason: "url-or-email"}})
		check("écanada.ca/a", nil)
		check("_canada.ca/a", nil)
		check("see canada.ca/a.", []kernels.Span{{Start: 4, End: 16, Reason: "url-or-email"}})
		check("mail a.b@c.ca now", []kernels.Span{{Start: 5, End: 13, Reason: "url-or-email"}})
		check("www.x.ca/y https://z.ca", []kernels.Span{{Start: 0, End: 10, Reason: "url-or-email"}, {Start: 11, End: 23, Reason: "url-or-email"}})
		check("a canada.ca/a", []kernels.Span{{Start: 2, End: 13, Reason: "url-or-email"}})
	})
	t.Run("supra_note_lookahead", func(t *testing.T) {
		aliases := map[string]bool{"A": true, "A, supra note 7": true}
		if _, err := kernels.NormalizeShortCitation("A, supra note 7x.", kernels.ShortCitationOptions{Aliases: aliases}); err == nil {
			t.Fatal("digit run followed by a letter must not parse")
		}
		r, err := kernels.NormalizeShortCitation("A, supra note 7, supra note 8 at para 1.", kernels.ShortCitationOptions{Aliases: aliases})
		if err != nil {
			t.Fatal(err)
		}
		// Backtracking selects the longer alias only when the lookahead holds;
		// the italic intention then targets the first "supra" in the head, as
		// the reference str.index does.
		mustEqual(t, r.Text, "A, supra note 7, supra note 8 at para 1.")
		if r.TokenFormat[0].Start != 3 {
			t.Fatalf("italic supra at %d", r.TokenFormat[0].Start)
		}
	})
}
