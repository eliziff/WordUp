package inspect

import (
	"encoding/xml"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/eliziff/WordUp/internal/project"
)

type structureSilver struct {
	Schema    int              `xml:"schema,attr"`
	Status    string           `xml:"status,attr"`
	Documents []silverDocument `xml:"documents>document"`
}

type silverDocument struct {
	Source     string            `xml:"source,attr"`
	SHA256     string            `xml:"source_sha256,attr"`
	Paragraphs []silverParagraph `xml:"p"`
}

type silverParagraph struct {
	ParagraphID string `xml:"paraId,attr"`
	XMLPath     string `xml:"xml_path,attr"`
	Role        string `xml:"role,attr"`
	Level       int    `xml:"level,attr"`
	Parent      string `xml:"parent,attr"`
	// These are optional so existing silver remains valid. When present,
	// contradiction is a presence contract: the detector's detailed wording
	// is evidence, not a second model-authored taxonomy.
	Ambiguous     string `xml:"ambiguous,attr"`
	Ambiguity     string `xml:"ambiguity,attr"`
	Contradiction string `xml:"contradiction,attr"`
}

// CompareStructureSilver compares silver labels with package-derived resolution.
func CompareStructureSilver(root, reference string, limit int) (map[string]any, error) {
	b, err := os.ReadFile(reference)
	if err != nil {
		return nil, err
	}
	var silver structureSilver
	if err := xml.Unmarshal(b, &silver); err != nil {
		return nil, fmt.Errorf("read structure silver: %w", err)
	}
	if silver.Schema != 1 || len(silver.Documents) == 0 {
		return nil, fmt.Errorf("structure silver must use schema 1 and name at least one document")
	}
	if limit <= 0 || limit > 1000 {
		limit = 100
	}
	mismatches := []map[string]any{}
	compared := 0
	mismatchCount := 0
	mismatchFields := map[string]int{}
	roleConfusions := map[string]int{}
	roleConfusionStyles := map[string]map[string]int{}
	roleExamples := map[string]any{}
	uncertaintyExamples := map[string]any{}
	parentMismatches := map[string]int{}
	for _, expectedDocument := range silver.Documents {
		path, err := project.Under(root, filepath.ToSlash(expectedDocument.Source))
		if err != nil {
			return nil, fmt.Errorf("silver source %q: %w", expectedDocument.Source, err)
		}
		actualDocument, err := StructureResolved(path)
		if err != nil {
			return nil, fmt.Errorf("resolve %q: %w", expectedDocument.Source, err)
		}
		if actualDocument["source_sha256"] != expectedDocument.SHA256 {
			return nil, fmt.Errorf("silver source hash changed for %q", expectedDocument.Source)
		}
		rows := actualDocument["paragraphs"].([]map[string]any)
		actual := make(map[string]map[string]any, len(rows))
		for _, row := range rows {
			actual[row["source_id"].(string)] = row
		}
		for _, expected := range expectedDocument.Paragraphs {
			id := "word/document.xml#" + expected.XMLPath
			if expected.ParagraphID != "" {
				id = "word/document.xml#paraId=" + expected.ParagraphID
			}
			row, ok := actual[id]
			compared++
			if !ok {
				appendStructureMismatch(&mismatches, limit, expectedDocument.Source, id, "identity", "present", "missing")
				mismatchCount++
				mismatchFields["identity"]++
				continue
			}
			got := row["resolved_structure"].(map[string]any)
			if compareStructureValue(&mismatches, limit, expectedDocument.Source, id, "role", expected.Role, got["role"]) {
				mismatchCount++
				mismatchFields["role"]++
				confusion := expected.Role + " -> " + fmt.Sprint(got["role"])
				roleConfusions[confusion]++
				if roleConfusionStyles[confusion] == nil {
					roleConfusionStyles[confusion] = map[string]int{}
				}
				style := fmt.Sprint(row["style_id"])
				if style == "" {
					style = "(none)"
				}
				roleConfusionStyles[confusion][style]++
				if roleExamples[confusion] == nil {
					roleExamples[confusion] = map[string]any{
						"document":                  expectedDocument.Source,
						"source_id":                 id,
						"xml_path":                  row["xml_path"],
						"xml_start":                 row["xml_start"],
						"xml_end":                   row["xml_end"],
						"context":                   row["context"],
						"style_id":                  row["style_id"],
						"style_chain":               styleNames(row),
						"formatting":                row["direct_formatting_evidence"],
						"paragraph_mark_formatting": row["paragraph_mark_formatting"],
						"alignment":                 row["paragraph_alignment"],
						"resolved_level":            got["level"],
						"evidence":                  got["evidence"],
					}
				}
			}
			if expected.Role == "heading" {
				if compareStructureValue(&mismatches, limit, expectedDocument.Source, id, "level", expected.Level, got["level"]) {
					mismatchCount++
					mismatchFields["level"]++
				}
			}
			ambiguity := expected.Ambiguous
			if ambiguity == "" {
				ambiguity = expected.Ambiguity
			}
			if ambiguity != "" {
				actualAmbiguous, _ := got["ambiguous"].(bool)
				if !ambiguitySatisfied(ambiguity, row, got) {
					mismatchCount++
					mismatchFields["ambiguous"]++
					if _, exists := uncertaintyExamples["ambiguous"]; !exists {
						uncertaintyExamples["ambiguous"] = structureEvidenceExample(expectedDocument.Source, id, row, got, true, actualAmbiguous)
					}
					appendStructureMismatch(&mismatches, limit, expectedDocument.Source, id, "ambiguous", true, actualAmbiguous)
				}
			}
			if expected.Contradiction != "" {
				value := strings.TrimSpace(expected.Contradiction)
				want := !strings.EqualFold(value, "false") && !strings.EqualFold(value, "none") && !strings.EqualFold(value, "absent")
				contradictions, _ := got["contradictions"].([]string)
				actual := len(contradictions) > 0
				if compareStructureValue(&mismatches, limit, expectedDocument.Source, id, "contradiction", want, actual) {
					mismatchCount++
					mismatchFields["contradiction"]++
					if _, exists := uncertaintyExamples["contradiction"]; !exists {
						uncertaintyExamples["contradiction"] = structureEvidenceExample(expectedDocument.Source, id, row, got, want, actual)
					}
				}
			}
			expectedParent := expected.Parent
			if expectedParent == "none" {
				expectedParent = ""
			} else if expectedParent != "" {
				if expectedParent[0] == '(' {
					expectedParent = "word/document.xml#" + expectedParent
				} else {
					expectedParent = "word/document.xml#paraId=" + expectedParent
				}
			}
			actualParent, _ := got["parent_source_id"].(string)
			if compareStructureValue(&mismatches, limit, expectedDocument.Source, id, "parent", expectedParent, actualParent) {
				mismatchCount++
				mismatchFields["parent"]++
				switch {
				case expectedParent == "":
					parentMismatches["unexpected"]++
				case actualParent == "":
					parentMismatches["missing"]++
				default:
					parentMismatches["wrong"]++
				}
			}
		}
	}
	return map[string]any{
		"schema": silver.Schema, "silver_status": silver.Status,
		"documents": len(silver.Documents), "paragraphs_compared": compared,
		"mismatches": mismatches, "mismatch_count": mismatchCount,
		"mismatch_fields": mismatchFields, "role_confusions": roleConfusions, "role_confusion_styles": roleConfusionStyles, "role_confusion_examples": roleExamples, "uncertainty_examples": uncertaintyExamples, "parent_mismatches": parentMismatches,
		"exact": mismatchCount == 0,
	}, nil
}

func structureEvidenceExample(document, id string, row, resolved map[string]any, expected, actual any) map[string]any {
	return map[string]any{
		"document":                 document,
		"source_id":                id,
		"xml_path":                 row["xml_path"],
		"xml_start":                row["xml_start"],
		"xml_end":                  row["xml_end"],
		"expected":                 expected,
		"actual":                   actual,
		"style_id":                 row["style_id"],
		"style_chain":              styleNames(row),
		"outline_evidence":         row["outline_evidence"],
		"numbering_evidence":       row["numbering_evidence"],
		"style_family_evidence":    row["style_family_evidence"],
		"paragraph_properties_xml": row["paragraph_properties_xml"],
		"resolved_evidence":        resolved["evidence"],
		"resolved_alternatives":    resolved["alternatives"],
		"resolved_contradictions":  resolved["contradictions"],
	}
}

func styleNames(row map[string]any) []string {
	names, _ := row["style_names"].([]string)
	return names
}

func ambiguitySatisfied(expected string, row, resolved map[string]any) bool {
	value := strings.TrimSpace(strings.ToLower(expected))
	if parsed, err := strconv.ParseBool(value); err == nil {
		actual, _ := resolved["ambiguous"].(bool)
		return parsed == actual
	}
	if value == "none" || value == "absent" {
		actual, _ := resolved["ambiguous"].(bool)
		return !actual
	}
	labels := strings.FieldsFunc(value, func(r rune) bool { return r == ',' || r == ';' || r == ' ' })
	if len(labels) == 0 {
		return false
	}
	marker, direct, unknown := false, false, false
	for _, label := range labels {
		switch {
		case strings.Contains(label, "marker"), strings.Contains(label, "roman"):
			marker = true
		case strings.Contains(label, "format"):
			direct = true
		default:
			unknown = true
		}
	}
	if unknown {
		actual, _ := resolved["ambiguous"].(bool)
		return actual
	}
	if marker && resolved["alternatives"] == nil {
		return false
	}
	if direct && !directFormattingEvidence(row) {
		return false
	}
	return true
}

func directFormattingEvidence(row map[string]any) bool {
	for _, key := range []string{"direct_formatting_evidence", "paragraph_mark_formatting"} {
		switch values := row[key].(type) {
		case map[string]any:
			if formattingMapHasValue(values) {
				return true
			}
		case map[string]int:
			for _, value := range values {
				if value > 0 {
					return true
				}
			}
		case map[string]bool:
			for _, value := range values {
				if value {
					return true
				}
			}
		}
	}
	return false
}

func formattingMapHasValue(values map[string]any) bool {
	for _, value := range values {
		switch typed := value.(type) {
		case bool:
			if typed {
				return true
			}
		case int:
			if typed > 0 {
				return true
			}
		case float64:
			if typed > 0 {
				return true
			}
		}
	}
	return false
}

func compareStructureValue(out *[]map[string]any, limit int, document, id, field string, expected, actual any) bool {
	if fmt.Sprint(expected) != fmt.Sprint(actual) {
		appendStructureMismatch(out, limit, document, id, field, expected, actual)
		return true
	}
	return false
}

func appendStructureMismatch(out *[]map[string]any, limit int, document, id, field string, expected, actual any) {
	if len(*out) < limit {
		*out = append(*out, map[string]any{"document": document, "source_id": id, "field": field, "expected": expected, "actual": actual})
	}
}
