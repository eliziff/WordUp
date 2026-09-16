package kernels_test

import (
	"errors"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap/kernels"
)

// edit mirrors the reference test helper: an explicitly authored plan applied
// to a string.
func edit(t *testing.T, s string, opts kernels.ProseOptions) string {
	t.Helper()
	opts.Authorized = true
	opts.AuthorizationBasis = "synthetic fixture explicitly authored"
	p, err := kernels.PlanProse(s, opts)
	if err != nil {
		t.Fatalf("PlanProse(%q): %v", s, err)
	}
	out, err := kernels.ApplyPlan(s, p)
	if err != nil {
		t.Fatalf("ApplyPlan(%q): %v", s, err)
	}
	return out
}

func mustEqual(t *testing.T, got, want string) {
	t.Helper()
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

// wantValueError asserts a refusal that the reference raised as ValueError
// (anything but a PermissionError).
func wantValueError(t *testing.T, err error) {
	t.Helper()
	var perm *kernels.PermissionError
	if err == nil || errors.As(err, &perm) {
		t.Fatalf("expected a value error, got %v", err)
	}
}

func wantPermissionError(t *testing.T, err error) {
	t.Helper()
	var perm *kernels.PermissionError
	if !errors.As(err, &perm) {
		t.Fatalf("expected a permission error, got %v", err)
	}
}

func runeIndex(s, sub string) int {
	return len([]rune(s[:strings.Index(s, sub)]))
}

func TestProtectionTests(t *testing.T) {
	t.Run("test_outside_quote_changes_inside_does_not", func(t *testing.T) {
		mustEqual(t, edit(t, "towards “towards 20%” towards", kernels.ProseOptions{}), "toward “towards 20%” toward")
	})
	t.Run("test_source_title_is_protected_without_classifying_quote_kind", func(t *testing.T) {
		mustEqual(t, edit(t, "Read “Towards Judgment” and move towards it.", kernels.ProseOptions{}), "Read “Towards Judgment” and move toward it.")
	})
	t.Run("test_nested_quotation", func(t *testing.T) {
		mustEqual(t, edit(t, "towards “She wrote ‘towards’ and 20%.” towards", kernels.ProseOptions{}), "toward “She wrote ‘towards’ and 20%.” toward")
	})
	t.Run("test_cross_paragraph_quote", func(t *testing.T) {
		mustEqual(t, edit(t, "towards “towards\rtowards” towards", kernels.ProseOptions{}), "toward “towards\rtowards” toward")
	})
	t.Run("test_continued_paragraph_openers", func(t *testing.T) {
		mustEqual(t, edit(t, "towards “towards\r“towards” towards", kernels.ProseOptions{}), "toward “towards\r“towards” toward")
	})
	t.Run("test_unclosed_quote_blocks_entire_container", func(t *testing.T) {
		s := "towards “towards\rA new paragraph towards"
		mustEqual(t, edit(t, s, kernels.ProseOptions{}), s)
	})
	t.Run("test_unmatched_close_blocks_entire_container", func(t *testing.T) {
		s := "towards” and towards"
		mustEqual(t, edit(t, s, kernels.ProseOptions{}), s)
	})
	t.Run("test_straight_nested_system_is_not_toggle_guessed", func(t *testing.T) {
		s := `He said "this "towards" text" towards.`
		mustEqual(t, edit(t, s, kernels.ProseOptions{}), s)
	})
	t.Run("test_straight_single_nested_system_not_toggle_guessed", func(t *testing.T) {
		s := "He said 'this 'towards' text' towards."
		mustEqual(t, edit(t, s, kernels.ProseOptions{}), s)
	})
	t.Run("test_coincident_insertions_are_rejected", func(t *testing.T) {
		s := "abc"
		a := kernels.Edit{Start: 1, End: 1, Old: "", New: "x", Rule: "one", ReadStart: 0, ReadEnd: 2}
		b := kernels.Edit{Start: 1, End: 1, Old: "", New: "y", Rule: "two", ReadStart: 0, ReadEnd: 2}
		p := kernels.Plan{Digest: kernels.Digest(s), Edits: []kernels.Edit{a, b}, Authorized: true, AuthorizationBasis: "fixture"}
		_, err := kernels.ApplyPlan(s, p)
		wantValueError(t, err)
	})
	t.Run("test_empty_authority_origin_is_rejected", func(t *testing.T) {
		p := kernels.Plan{Digest: kernels.Digest("abc"), Authorized: true, AuthorizationBasis: ""}
		_, err := kernels.ApplyPlan("abc", p)
		wantPermissionError(t, err)
	})
	t.Run("test_invalid_protection_bounds_are_rejected", func(t *testing.T) {
		p := kernels.Plan{Digest: kernels.Digest("abc"), Protected: []kernels.Span{{Start: -1, End: 2}}, Authorized: true, AuthorizationBasis: "fixture"}
		_, err := kernels.ApplyPlan("abc", p)
		wantValueError(t, err)
	})
	t.Run("test_contraction_does_not_open_quote", func(t *testing.T) {
		mustEqual(t, edit(t, "Don't move towards it.", kernels.ProseOptions{}), "Don't move toward it.")
	})
	t.Run("test_curly_single_quote", func(t *testing.T) {
		mustEqual(t, edit(t, "‘towards’ towards", kernels.ProseOptions{}), "‘towards’ toward")
	})
	t.Run("test_url_including_naked_host", func(t *testing.T) {
		s := "https://example.ca/towards?q=20%25 towards canada.ca/a/2023-2030"
		mustEqual(t, edit(t, s, kernels.ProseOptions{}), "https://example.ca/towards?q=20%25 toward canada.ca/a/2023-2030")
	})
	t.Run("test_explicit_blockquote_and_roman_title_masks", func(t *testing.T) {
		s := "towards Old Judgements towards"
		mustEqual(t, edit(t, s, kernels.ProseOptions{Extra: []kernels.Span{{Start: 8, End: 22, Reason: "source-title"}}}), "toward Old Judgements toward")
		mustEqual(t, edit(t, s, kernels.ProseOptions{Extra: []kernels.Span{{Start: 0, End: len([]rune(s)), Reason: "unmarked-source-passage"}}}), s)
	})
	t.Run("test_field_revision_and_math_masks", func(t *testing.T) {
		s := "towards 20% towards"
		for _, reason := range []string{"field", "revision", "math", "locked-sdt", "text-frame", "comment"} {
			t.Run(reason, func(t *testing.T) {
				mustEqual(t, edit(t, s, kernels.ProseOptions{Extra: []kernels.Span{{Start: 0, End: len([]rune(s)), Reason: reason}}}), s)
			})
		}
	})
	t.Run("test_bracketed_shortform", func(t *testing.T) {
		mustEqual(t, edit(t, "towards [Towards Judgment] towards", kernels.ProseOptions{}), "toward [Towards Judgment] toward")
	})
	t.Run("test_utf16_not_python_offset", func(t *testing.T) {
		s := "😀 towards and étowards"
		p, err := kernels.PlanProse(s, kernels.ProseOptions{Authorized: true, AuthorizationBasis: "fixture"})
		if err != nil {
			t.Fatal(err)
		}
		out, err := kernels.ApplyPlan(s, p)
		if err != nil {
			t.Fatal(err)
		}
		mustEqual(t, out, "😀 toward and étowards")
		if off, err := kernels.UTF16Offset(s, 2); err != nil || off != 3 {
			t.Fatalf("UTF16Offset = %d, %v; want 3", off, err)
		}
	})
	t.Run("test_advisory_never_mutates", func(t *testing.T) {
		p, err := kernels.PlanProse("towards", kernels.ProseOptions{})
		if err != nil {
			t.Fatal(err)
		}
		if len(p.Edits) != 1 {
			t.Fatalf("edits = %v", p.Edits)
		}
		_, err = kernels.ApplyPlan("towards", p)
		wantPermissionError(t, err)
	})
	t.Run("test_stale_plan_is_rejected", func(t *testing.T) {
		p, err := kernels.PlanProse("towards", kernels.ProseOptions{Authorized: true, AuthorizationBasis: "fixture"})
		if err != nil {
			t.Fatal(err)
		}
		_, err = kernels.ApplyPlan("towards elsewhere", p)
		wantValueError(t, err)
	})
	t.Run("test_forged_protected_patch_is_rejected", func(t *testing.T) {
		s := "“towards”"
		e := kernels.Edit{Start: 7, End: 8, Old: "s", New: "", Rule: "forged", ReadStart: 1, ReadEnd: 8}
		p := kernels.Plan{Digest: kernels.Digest(s), Edits: []kernels.Edit{e}, Protected: kernels.QuotationGuards(s), Authorized: true, AuthorizationBasis: "fixture"}
		_, err := kernels.ApplyPlan(s, p)
		wantValueError(t, err)
	})
	t.Run("test_same_value_wrong_occurrence_not_relocated", func(t *testing.T) {
		s := "“towards” towards"
		p, err := kernels.PlanProse(s, kernels.ProseOptions{Authorized: true, AuthorizationBasis: "fixture"})
		if err != nil {
			t.Fatal(err)
		}
		if len(p.Edits) != 1 {
			t.Fatalf("edits = %v", p.Edits)
		}
		if p.Edits[0].Start <= runeIndex(s, "”") {
			t.Fatalf("edit start %d is not after the closing quote", p.Edits[0].Start)
		}
	})
	t.Run("test_valid_date_only_and_no_quote_date", func(t *testing.T) {
		mustEqual(t, edit(t, "February 29, 2024; February 29, 2023; “May 1, 2024”.", kernels.ProseOptions{}),
			"29 February 2024; February 29, 2023; “May 1, 2024”.")
	})
	t.Run("test_date_internal_whitespace_is_one_rule", func(t *testing.T) {
		mustEqual(t, edit(t, "May  1, 2024", kernels.ProseOptions{}), "1 May 2024")
	})
	t.Run("test_percent_number_value", func(t *testing.T) {
		mustEqual(t, edit(t, "20% 20.50% 20%% 20%20", kernels.ProseOptions{}), "20 percent 20.50 percent 20%% 20%20")
	})
	t.Run("test_combining_boundaries_and_unicode_minus_not_reclassified", func(t *testing.T) {
		s := "May 1, 2024́ é20% 20%́ −20%"
		mustEqual(t, edit(t, s, kernels.ProseOptions{}), s)
	})
	t.Run("test_percent_internal_whitespace", func(t *testing.T) {
		mustEqual(t, edit(t, "20  %", kernels.ProseOptions{}), "20 percent")
	})
	t.Run("test_idempotence", func(t *testing.T) {
		s := "towards judgement  May 2, 2024: 20% “towards”"
		once := edit(t, s, kernels.ProseOptions{})
		mustEqual(t, edit(t, once, kernels.ProseOptions{}), once)
	})
	t.Run("test_disabled_feature_is_independent", func(t *testing.T) {
		mustEqual(t, edit(t, "towards 20%", kernels.ProseOptions{Enabled: kernels.NewSwitches("percent")}), "towards 20 percent")
	})
}
