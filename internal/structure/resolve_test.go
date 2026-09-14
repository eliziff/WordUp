package structure

import "testing"

func TestResolveKeepsOutlineAndDemotesOrdinaryList(t *testing.T) {
	rows := []map[string]any{
		{"context": "body", "text": "I. INTRODUCTION", "style_id": "Heading1", "outline_evidence": map[string]any{"level": 1}, "marker_interpretations": []Interpretation{{Family: "roman", Value: 1}}, "sequence_evidence": Assignment{Family: "roman", Value: 1, Level: 1, Action: "open_level"}},
		{"context": "body", "text": "1. Buy milk.", "style_id": "ListParagraph", "marker_interpretations": []Interpretation{{Family: "integer", Value: 1}}, "sequence_evidence": Assignment{Family: "integer", Value: 1, Level: 2, Action: "open_level"}},
		{"context": "table", "text": "A. Table value", "style_id": "Heading1"},
		{"context": "body", "text": "AUTOMATICALLY NUMBERED HEADING", "style_id": "Heading2", "numbering_evidence": map[string]any{"level": 2, "family": "upperLetter"}},
	}
	Resolve(rows)
	if rows[0]["resolved_structure"].(map[string]any)["role"] != "heading" {
		t.Fatal(rows[0])
	}
	if rows[1]["resolved_structure"].(map[string]any)["role"] == "heading" {
		t.Fatal("ordinary sentence became heading")
	}
	table := rows[2]["resolved_structure"].(map[string]any)
	if table["role"] != "table" {
		t.Fatal(rows[2])
	}
	if table["level"] != 0 || table["parent_paragraph"] != 1 || table["ambiguous"] != false {
		t.Fatalf("non-body result did not preserve contract defaults: %#v", table)
	}
	if rows[3]["resolved_structure"].(map[string]any)["role"] != "heading" || rows[3]["resolved_structure"].(map[string]any)["level"] != 2 {
		t.Fatal(rows[3])
	}
}

func TestResolveDoesNotTreatOneBoldWordAsHeadingEmphasis(t *testing.T) {
	row := map[string]any{"context": "body", "text": "Important words inside an ordinary sentence", "style_id": "Normal", "direct_formatting_evidence": map[string]int{"text_units": 42, "bold_units": 9}}
	Resolve([]map[string]any{row})
	evidence := row["resolved_structure"].(map[string]any)["evidence"].([]string)
	for _, item := range evidence {
		if item == "emphasis" {
			t.Fatal("minor inline emphasis became paragraph-level heading evidence")
		}
	}
}

func TestResolveRetainsAlternativesScoreAndContradictions(t *testing.T) {
	row := map[string]any{
		"context":                "body",
		"text":                   "I. Introduction",
		"style_id":               "Normal",
		"marker_interpretations": []Interpretation{{Family: "roman_.", Value: 1}, {Family: "upper_alpha_.", Value: 9}},
		"sequence_evidence":      Assignment{Family: "roman_.", Value: 1, Level: 1, Action: "open_level"},
		"outline_evidence":       map[string]any{"level": 1},
		"style_family_evidence":  map[string]any{"coherent": true, "corroborated_level": 2},
	}
	Resolve([]map[string]any{row})
	resolved := row["resolved_structure"].(map[string]any)
	if resolved["candidate_score"].(int) <= 0 || resolved["confidence"].(int) <= 0 {
		t.Fatalf("missing score fields: %#v", resolved)
	}
	if resolved["ambiguous"] != true {
		t.Fatalf("contradictory alternatives were not marked ambiguous: %#v", resolved)
	}
	if got, ok := resolved["alternatives"].([]Interpretation); !ok || len(got) != 2 {
		t.Fatalf("marker alternatives were dropped: %#v", resolved["alternatives"])
	}
	if got, ok := resolved["contradictions"].([]string); !ok || len(got) != 1 {
		t.Fatalf("contradictions were dropped: %#v", resolved["contradictions"])
	}
}

func TestResolveDetectsLayoutFrontMatter(t *testing.T) {
	rows := []map[string]any{
		{"context": "body", "text": "A General Title", "style_id": "Normal", "paragraph_alignment": "center", "direct_formatting_evidence": map[string]int{"text_units": 14, "bold_units": 14}},
		{"context": "body", "text": "Jane Doe, Editor", "style_id": "Normal", "paragraph_alignment": "center"},
		{"context": "body", "text": "Introduction", "style_id": "Heading 1", "outline_evidence": map[string]any{"level": 1}},
	}
	Resolve(rows)
	if got := rows[0]["resolved_structure"].(map[string]any)["role"]; got != "title" {
		t.Fatalf("centered first paragraph was not title: %#v", rows[0])
	}
	if got := rows[1]["resolved_structure"].(map[string]any)["role"]; got != "author" {
		t.Fatalf("centered author line was not author: %#v", rows[1])
	}
}
