package structure

import "testing"

func TestLadderAmbiguityAndParentRestart(t *testing.T) {
	rows := [][]Interpretation{{{"roman", 1}, {"alpha", 9}}, {{"alpha", 1}}, {{"alpha", 2}}, {{"roman", 2}}, {{"alpha", 1}}}
	got := HeadingLadder(rows)
	for i, want := range []int{1, 2, 2, 1, 2} {
		if got[i].Level != want {
			t.Fatalf("%d: %#v", i, got)
		}
	}
	if got[4].Action != "open_level" {
		t.Fatal("child restart after parent increment must be legal", got)
	}
}
func TestLadderRetainsViolationsAndGaps(t *testing.T) {
	got := HeadingLadder([][]Interpretation{{{"numeric", 3}}, {{"numeric", 5}}, {{"numeric", 1}}, {{"numeric", 1}}, {}})
	for i, want := range []string{"open_midcounter", "jump_forward", "illegal_restart", "illegal_restart", "violation"} {
		if got[i].Action != want {
			t.Fatalf("%d: %#v", i, got)
		}
	}
}

func TestMarkerChoicesPreserveAmbiguity(t *testing.T) {
	got := MarkerChoices("I. Introduction")
	if len(got) != 2 || got[0].Family != "roman_." || got[1].Family != "upper_alpha_." {
		t.Fatal(got)
	}
	lower := MarkerChoices("i. introduction")
	if len(lower) != 2 || lower[0].Family != "roman_." || lower[1].Family != "lower_alpha_." {
		t.Fatalf("lowercase Roman/alpha ambiguity was lost: %#v", lower)
	}
	if len(MarkerChoices("1.5 percent")) != 0 {
		t.Fatal("decimal treated as marker")
	}
	if len(MarkerChoices("Part IV - Contractual Rights")) != 1 {
		t.Fatal("named part not recognized")
	}
}

func TestNamedHeadingDashVariants(t *testing.T) {
	for _, text := range []string{"Part I \u2013 Introduction", "Part II - Background", "Part IV \u2014Contractual Rights"} {
		got := MarkerChoices(text)
		if len(got) == 0 || got[0].Family != "roman_part_named_section" {
			t.Fatal(text, got)
		}
	}
	if got := MarkerChoices("Part iv - Lowercase Roman"); len(got) != 1 || got[0].Family != "roman_part_named_section" || got[0].Value != 4 {
		t.Fatalf("lowercase named Roman marker was not recognized: %#v", got)
	}
}

func TestGenericNamedHeadingPrefixes(t *testing.T) {
	for _, text := range []string{"Theme 1: Pricing", "Section II - Scope", "Appendix A: Sources"} {
		if got := MarkerChoices(text); len(got) == 0 {
			t.Fatalf("named heading prefix was not recognized: %q", text)
		}
	}
}

func TestNamedHeadingPrefixesCreateNestedFamilies(t *testing.T) {
	got := HeadingLadder([][]Interpretation{
		MarkerChoices("Part I - Introduction"),
		MarkerChoices("Theme 1: Pricing"),
	})
	if len(got) != 2 || got[0].Level != 1 || got[1].Level != 2 {
		t.Fatalf("named prefixes collapsed into one ladder family: %#v", got)
	}
}

func TestPartAndChapterKeepSeparateFamilies(t *testing.T) {
	part := MarkerChoices("Part II - Introduction")
	chapter := MarkerChoices("Chapter II - Scope")
	if len(part) != 1 || len(chapter) != 1 || part[0].Family == chapter[0].Family {
		t.Fatalf("named prefixes were merged: part=%#v chapter=%#v", part, chapter)
	}
	got := HeadingLadder([][]Interpretation{part, chapter, MarkerChoices("Chapter III - Detail")})
	if got[0].Level != 1 || got[1].Level != 2 || got[2].Level != 2 {
		t.Fatalf("separate named ladders did not nest predictably: %#v", got)
	}
}
