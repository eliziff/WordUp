package verify

import "testing"

func TestOnlyCompilerErrorsAreAcknowledged(t *testing.T) {
	for _, tc := range []struct {
		title, message string
		want           bool
	}{
		{"Microsoft Visual Basic for Applications", "Compile error:\nUser-defined type not defined", true},
		{"Microsoft Visual Basic for Applications", "Run-time error 5", false},
		{"Microsoft Word", "Compile error:", false},
		{"Microsoft Visual Basic for Applications", "Save changes?", false},
	} {
		if got := compilerErrorDialog(map[string]any{"title": tc.title, "messages": []any{tc.message}}); got != tc.want {
			t.Fatal(tc, got)
		}
	}
}
