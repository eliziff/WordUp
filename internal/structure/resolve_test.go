package structure

import (
	"encoding/json"
	"math"
	"testing"
)

func TestIntValueAcceptsNativeAndReportNumberTypes(t *testing.T) {
	for _, test := range []struct {
		value any
		want  int
	}{{int8(3), 3}, {int16(4), 4}, {int32(5), 5}, {int64(6), 6}, {uint(7), 7}, {uint8(8), 8}, {uint16(9), 9}, {uint32(10), 10}, {uint64(11), 11}, {json.Number("12"), 12}} {
		got, ok := intValue(test.value)
		if !ok || got != test.want {
			t.Fatalf("%T: got %d, %v; want %d, true", test.value, got, ok, test.want)
		}
	}
	for _, test := range []struct {
		value any
		want  int
	}{{float32(13), 13}, {float64(14), 14}, {json.Number("15.0"), 15}} {
		if got, ok := intValue(test.value); !ok || got != test.want {
			t.Fatalf("%T: got %d, %v; want %d, true", test.value, got, ok, test.want)
		}
	}
	for _, value := range []any{uint64(^uint64(0)), float64(1.5), math.NaN(), math.Inf(1), json.Number("not-a-number")} {
		if _, ok := intValue(value); ok {
			t.Fatalf("%T unexpectedly accepted as an integer: %v", value, value)
		}
	}
}

func TestFormattingSignalUsesParagraphMarkEvidence(t *testing.T) {
	for _, evidence := range []any{
		map[string]bool{"bold": true},
		map[string]any{"small_caps": true},
	} {
		if !formattingSignal(map[string]any{"paragraph_mark_formatting": evidence}) {
			t.Fatalf("paragraph-mark evidence was ignored: %#v", evidence)
		}
	}
	if formattingSignal(map[string]any{"paragraph_mark_formatting": map[string]bool{"underline": true}}) {
		t.Fatal("underline alone is not paragraph-level heading evidence")
	}
	if !formattingSignal(map[string]any{"direct_formatting_evidence": map[string]any{"text_units": 10.0, "bold_units": 7.0}}) {
		t.Fatal("decoded direct-formatting counts were ignored")
	}
}

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

func TestResolveRetainsNumberingFormattingConflict(t *testing.T) {
	body := map[string]any{"context": "body", "text": "Ordinary body paragraph", "style_id": "BodyText"}
	row := map[string]any{
		"context":                    "body",
		"text":                       "1. Directly formatted heading",
		"style_id":                   "BodyText",
		"numbering_evidence":         map[string]any{"level": 1, "family": "decimal"},
		"direct_formatting_evidence": map[string]int{"text_units": 32, "bold_units": 32},
	}
	Resolve([]map[string]any{body, row})
	resolved := row["resolved_structure"].(map[string]any)
	contradictions, ok := resolved["contradictions"].([]string)
	if !ok || len(contradictions) != 1 || resolved["ambiguous"] != true {
		t.Fatalf("numbering/formatting conflict was dropped: %#v", resolved)
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

func TestResolveAcceptsDecodedEvidenceShapes(t *testing.T) {
	row := map[string]any{
		"context":     "body",
		"text":        "I. Introduction",
		"style_id":    "Normal",
		"style_names": []any{"Outline 2"},
		"marker_interpretations": []any{
			map[string]any{"family": "roman_.", "value": float64(1)},
			map[string]any{"family": "upper_alpha_.", "value": float64(9)},
		},
		"sequence_evidence": map[string]any{
			"family": "roman_.", "value": float64(1), "level": float64(2), "action": "open_level",
		},
	}
	Resolve([]map[string]any{row})
	resolved := row["resolved_structure"].(map[string]any)
	if resolved["role"] != "heading" || resolved["level"] != 2 {
		t.Fatalf("decoded style/sequence evidence was ignored: %#v", resolved)
	}
	if alternatives, ok := resolved["alternatives"].([]Interpretation); !ok || len(alternatives) != 2 {
		t.Fatalf("decoded marker alternatives were dropped: %#v", resolved["alternatives"])
	}
}

func TestResolveRejectsOutOfRangeHierarchyEvidence(t *testing.T) {
	cases := []map[string]any{
		{"context": "body", "text": "Body prose", "outline_evidence": map[string]any{"level": 99}},
		{"context": "body", "text": "Body prose", "numbering_evidence": map[string]any{"level": 0, "family": "decimal"}},
		{"context": "body", "text": "Body prose", "sequence_evidence": Assignment{Level: 10, Action: "open_level"}},
	}
	for i, row := range cases {
		Resolve([]map[string]any{row})
		resolved, ok := row["resolved_structure"].(map[string]any)
		if !ok {
			t.Fatalf("case %d did not resolve: %#v", i, row)
		}
		if level, _ := resolved["level"].(int); level < 0 || level > maxHierarchyLevel {
			t.Fatalf("case %d emitted invalid level: %#v", i, resolved)
		}
		if resolved["ambiguous"] != true {
			t.Fatalf("case %d discarded malformed hierarchy evidence: %#v", i, resolved)
		}
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

func TestResolveRecognizesCamelCaseSemanticStyles(t *testing.T) {
	rows := []map[string]any{
		{"context": "body", "text": "Document name", "style_id": "DocumentTitle"},
		{"context": "body", "text": "Plain title", "style_id": "Title"},
		{"context": "body", "text": "Contents", "style_id": "TOCHeading"},
		{"context": "body", "text": "Quoted material", "style_id": "BlockQuote"},
		{"context": "body", "text": "Header text", "style_id": "Header"},
	}
	Resolve(rows)
	want := []string{"title", "title", "toc", "quotation", "body"}
	for i, role := range want {
		if got := rows[i]["resolved_structure"].(map[string]any)["role"]; got != role {
			t.Fatalf("row %d role=%v want %s", i, got, role)
		}
	}
}

func TestResolveDoesNotTrustRepurposedTOCStyle(t *testing.T) {
	rows := []map[string]any{
		{"context": "body", "text": "A Journal Article", "style_id": "Normal", "paragraph_alignment": "center", "direct_formatting_evidence": map[string]int{"text_units": 17, "bold_units": 17}},
		{"context": "body", "text": "By Jane Doe", "style_id": "Normal", "paragraph_alignment": "center"},
		{"context": "body", "text": "This long italic paragraph describes the article and its scope. It is deliberately long enough to look like front matter rather than a table-of-contents label, and it contains no page-number tabs or outline metadata.", "style_id": "TOCHeading", "direct_formatting_evidence": map[string]int{"text_units": 211, "italic_units": 211}},
	}
	Resolve(rows)
	got := rows[2]["resolved_structure"].(map[string]any)
	if got["role"] != "body" {
		t.Fatalf("repurposed TOC style was promoted without TOC text evidence: %#v", got)
	}
}

func TestResolveRecognizesTOCEntriesWithRepurposedStyle(t *testing.T) {
	rows := []map[string]any{
		{"context": "body", "text": "Table of Contents", "style_id": "TOCHeading"},
		{"context": "body", "text": "I. INTRODUCTION\t1", "style_id": "TOCHeading"},
	}
	Resolve(rows)
	for i, row := range rows {
		if got := row["resolved_structure"].(map[string]any)["role"]; got != "toc" {
			t.Fatalf("TOC row %d role=%v", i, got)
		}
	}
}
