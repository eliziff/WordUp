package vbaparse

import (
	"math"
	"testing"
)

func TestConditionalCompilationRejectsInvalidConditions(t *testing.T) {
	for _, source := range []string{
		"#If \"not a number\" Then\nOption Explicit\n#End If\n",
		"#If False Then\nOption Explicit\n#ElseIf \"not a number\" Then\nOption Explicit\n#End If\n",
	} {
		if result, err := Parse(source); err == nil {
			t.Fatalf("invalid condition silently selected a branch: %v", result)
		}
	}
	for _, value := range []any{math.NaN(), math.Inf(1), math.Inf(-1)} {
		if _, err := ccTruth(value); err == nil {
			t.Fatalf("accepted non-finite condition %v", value)
		}
	}
}

func TestConditionalCompilationPreservesActiveBranchCoordinates(t *testing.T) {
	source := "#Const n = 2 + 3 * 4\n#If VBA7 And Win64 And n = 14 Then\nPublic Sub Valid()\nEnd Sub\n#Else\nPublic Sub Broken()\nDim x As\nEnd Sub\n#End If\n"
	r, e := ParseWithConstants(source, map[string]any{"Win64": true})
	if e != nil || r["syntax_valid"] != true {
		t.Fatal(r, e)
	}
	r, e = ParseWithConstants(source, map[string]any{"Win64": false})
	if e != nil || r["syntax_valid"] != false {
		t.Fatal(r, e)
	}
	ds := r["diagnostics"].([]Diagnostic)
	if len(ds) == 0 || ds[0].Line != 7 {
		t.Fatal(ds)
	}
	nested := "#If False Then\n#Const ignored = 1\n#If True Then\ninvalid VBA\n#End If\n#ElseIf True Then\nOption Explicit\n#Else\ninvalid\n#End If\n"
	r, e = Parse(nested)
	if e != nil || r["syntax_valid"] != true {
		t.Fatal(r, e)
	}
}
