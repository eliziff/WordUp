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
		if len(got) == 0 || got[0].Family != "roman_named_section" {
			t.Fatal(text, got)
		}
	}
}
