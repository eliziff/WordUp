package native

import (
	"strings"
	"testing"
)

func TestEvaluationSourceLocations(t *testing.T) {
	selectSource, _ := evaluationSource("Probe", "tag", "Select Case 2\nCase 1, 2\nEvaluate = True\nCase Else\nEvaluate = False\nEnd Select")
	if !strings.Contains(selectSource, "\nCase 1, 2\n3 Evaluate = True\nCase Else\n5 Evaluate = False") {
		t.Fatal("instrumentation changed Case syntax", selectSource)
	}
	body := "Dim x As String\nx = \"a'\" & _\n    \"b\"\n' comment _\nRem comment _\n#If VBA7 Then\nagain:\nErr.Raise 5\n#End If"
	source, mode := evaluationSource("Probe", "tag", body)
	for _, want := range []string{"1 Dim x", "2 x =", "\n    \"b\"\n", "\n' comment _\n", "\nRem comment _\n", "\n#If VBA7 Then\nagain:\n8 Err.Raise 5\n#End If"} {
		if !strings.Contains(source, want) {
			t.Fatalf("lost syntax or mapping %q in %s", want, source)
		}
	}
	details := map[string]any{}
	evaluationLineDetails(details, body, mode, float64(8))
	if details["body_line"] != 8 || details["body_line_text"] != "Err.Raise 5" {
		t.Fatalf("wrong original location: %v", details)
	}
	source, mode = evaluationSource("Probe", "tag", "100 Err.Raise 5\nEvaluate = Erl")
	if mode != "caller_numbered" || !strings.Contains(source, "\n100 Err.Raise 5\nEvaluate = Erl\n") {
		t.Fatal("rewrote explicit line numbers")
	}
	source, _ = evaluationSource("Probe", "tag", "' continued comment _\nignored text\nErr.Raise 5")
	if !strings.Contains(source, "\nignored text\n3 Err.Raise 5") {
		t.Fatal("numbered a continued comment")
	}
	for _, line := range []string{"x = \"Rem _\"", "x = \"a\"\"'b\"", "x = 1: Rem ignored _", "x = 1 ' ignored _"} {
		code := scratchCode(line)
		if strings.HasSuffix(strings.TrimSpace(code), "_") {
			t.Fatalf("comment/string mistaken for continuation: %s", code)
		}
	}
}
