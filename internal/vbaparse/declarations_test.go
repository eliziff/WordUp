package vbaparse

import "testing"

func TestDeclarationCollisions(t *testing.T) {
	for _, tc := range []struct {
		source     string
		collisions int
	}{
		{"Public Const WU_PARENT As Long = 2\nPrivate Function WU_Parent() As Long\nEnd Function\n", 1},
		{"Private value As Long, Other As String\nPublic Sub other()\nEnd Sub\n", 1},
		{"Private value As Long\nPublic Sub Work()\nDim value As Long\nEnd Sub\n", 0},
		{"Public Property Get Value() As Long\nEnd Property\nPublic Property Let Value(ByVal n As Long)\nEnd Property\n", 0},
		{"#If False Then\nPublic Const same = 1\n#End If\nPublic Sub same()\nEnd Sub\n", 0},
	} {
		result, err := Parse(tc.source)
		if err != nil {
			t.Fatal(err)
		}
		if result["syntax_valid"] != true {
			t.Fatalf("syntax: %v", result)
		}
		if ds := result["diagnostics"].([]Diagnostic); len(ds) != tc.collisions {
			t.Fatalf("source=%s diagnostics=%v", tc.source, ds)
		}
	}
}
