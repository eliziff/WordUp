package structure

import "testing"

func TestResolveKeepsOutlineAndDemotesOrdinaryList(t *testing.T) {
	rows := []map[string]any{
		{"context":"body", "text":"I. INTRODUCTION", "style_id":"Heading1", "outline_evidence":map[string]any{"level":1}, "marker_interpretations":[]Interpretation{{Family:"roman",Value:1}}, "sequence_evidence":Assignment{Family:"roman",Value:1,Level:1,Action:"open_level"}},
		{"context":"body", "text":"1. Buy milk.", "style_id":"ListParagraph", "marker_interpretations":[]Interpretation{{Family:"integer",Value:1}}, "sequence_evidence":Assignment{Family:"integer",Value:1,Level:2,Action:"open_level"}},
		{"context":"table", "text":"A. Table value", "style_id":"Heading1"},
	}
	Resolve(rows)
	if rows[0]["resolved_structure"].(map[string]any)["role"] != "heading" { t.Fatal(rows[0]) }
	if rows[1]["resolved_structure"].(map[string]any)["role"] == "heading" { t.Fatal("ordinary sentence became heading") }
	if rows[2]["resolved_structure"].(map[string]any)["role"] != "table" { t.Fatal(rows[2]) }
}
