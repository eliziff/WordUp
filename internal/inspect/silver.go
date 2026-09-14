package inspect

import (
	"encoding/xml"
	"fmt"
	"os"
	"path/filepath"

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
	roleExamples := map[string]any{}
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
				if roleExamples[confusion] == nil {
					roleExamples[confusion] = map[string]any{"document": expectedDocument.Source, "source_id": id, "context": row["context"], "style_id": row["style_id"], "style_chain": styleNames(row), "text_units": formattingUnits(row), "resolved_level": got["level"]}
				}
			}
			if expected.Role == "heading" {
				if compareStructureValue(&mismatches, limit, expectedDocument.Source, id, "level", expected.Level, got["level"]) {
					mismatchCount++
					mismatchFields["level"]++
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
		"mismatch_fields": mismatchFields, "role_confusions": roleConfusions, "role_confusion_examples": roleExamples, "parent_mismatches": parentMismatches,
		"exact": mismatchCount == 0,
	}, nil
}

func styleNames(row map[string]any) []string {
	names, _ := row["style_names"].([]string)
	return names
}

func formattingUnits(row map[string]any) int {
	evidence, _ := row["direct_formatting_evidence"].(map[string]int)
	return evidence["text_units"]
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
