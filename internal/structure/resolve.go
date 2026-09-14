package structure

import (
	"fmt"
	"strings"
	"unicode"
)

// Resolve fuses document-wide style, outline, marker and formatting evidence.
// It mutates the supplied observation rows so callers retain exact source facts.
func Resolve(rows []map[string]any) map[string]any {
	parents := [10]int{}
	headings, candidates, ambiguities := 0, 0, 0
	for i, row := range rows {
		context, _ := row["context"].(string)
		text, _ := row["text"].(string)
		style, _ := row["style_id"].(string)
		score, level := 0, 0
		evidence := []string{}
		if role := semanticRole(row); role != "" {
			resolved := map[string]any{"role": role, "confidence": 100, "evidence": []string{"semantic-style-or-context"}}
			if role == "quotation" || role == "abstract" {
				attachParent(resolved, rows, parents)
			}
			row["resolved_structure"] = resolved
			continue
		}
		if context != "body" || strings.TrimSpace(text) == "" {
			resolved := map[string]any{"role": fallbackRole(context, text), "confidence": 100, "evidence": evidence}
			if resolved["role"] != "blank" {
				attachParent(resolved, rows, parents)
			}
			row["resolved_structure"] = resolved
			continue
		}
		if outline, ok := row["outline_evidence"].(map[string]any); ok && outline["body_text"] != true {
			if n, ok := intValue(outline["level"]); ok {
				level = n
				score += 55
				evidence = append(evidence, "native-outline")
			}
		}
		if family, ok := row["style_family_evidence"].(map[string]any); ok && family["coherent"] == true {
			if n, ok := intValue(family["corroborated_level"]); ok {
				if level > 0 && level != n {
					row["structure_contradiction"] = fmt.Sprintf("native level %d conflicts with coherent style-family level %d", level, n)
				} else if level == 0 {
					level = n
				}
				score += 15
				evidence = append(evidence, "coherent-style-family")
			}
		}
		if level == 0 {
			if n := headingLevelFromStyle(row); n > 0 {
				level = n
				evidence = append(evidence, "heading-style-level")
			}
		}
		if headingStyle(style) {
			score += 30
			evidence = append(evidence, "heading-style-name")
		}
		if _, ok := row["marker_interpretations"]; ok {
			score += 25
			evidence = append(evidence, "marker-grammar")
		}
		if numbering, ok := row["numbering_evidence"].(map[string]any); ok && numbering["family"] != "bullet" {
			if n, ok := intValue(numbering["level"]); ok && level == 0 {
				level = n
			}
			score += 12
			evidence = append(evidence, "native-numbering")
		}
		if sequence, ok := row["sequence_evidence"].(Assignment); ok && sequence.Action != "violation" {
			if level == 0 {
				level = sequence.Level
			}
			score += 15
			evidence = append(evidence, "sequence-"+sequence.Action)
		}
		// A marker sequence is evidence, not permission to promote ordinary list
		// paragraphs. Native outline or a heading-style family can override this.
		if level > 0 && !hasNativeHeading(row) && listStyle(style) {
			score -= 30
			evidence = append(evidence, "ordinary-list-risk")
		}
		if len([]rune(text)) <= 160 {
			score += 8
		} else {
			score -= 25
			evidence = append(evidence, "long-paragraph")
		}
		props, _ := row["paragraph_properties_xml"].(string)
		if strings.Contains(props, ":keepNext") {
			score += 8
			evidence = append(evidence, "keep-next")
		}
		if formattingSignal(row) {
			score += 8
			evidence = append(evidence, "emphasis")
		}
		if upperShare(text) >= .8 {
			score += 8
			evidence = append(evidence, "uppercase-text")
		}
		if level >= 5 && !headingStyle(style) && row["marker_interpretations"] == nil {
			score -= 45
			evidence = append(evidence, "deep-outline-body-risk")
		}
		if sentenceEnding(text) && level == 0 {
			score -= 18
			evidence = append(evidence, "sentence-ending")
		}
		ambiguous := false
		role := "body"
		if level > 0 && score >= 35 {
			role = "heading"
			headings++
			parent := 0
			for n := level - 1; n >= 1; n-- {
				if parents[n] != 0 {
					parent = parents[n]
					break
				}
			}
			resolved := map[string]any{"role": role, "level": level, "parent_paragraph": parent, "confidence": clamp(score), "evidence": evidence}
			if parent > 0 {
				resolved["parent_source_id"] = rows[parent-1]["source_id"]
			}
			row["resolved_structure"] = resolved
			parents[level] = i + 1
			for n := level + 1; n <= 9; n++ {
				parents[n] = 0
			}
			continue
		}
		if score >= 25 {
			role = "candidate"
			ambiguous = true
			candidates++
			ambiguities++
		}
		if row["structure_contradiction"] != nil {
			ambiguous = true
			ambiguities++
		}
		resolved := map[string]any{"role": role, "level": level, "confidence": clamp(score), "ambiguous": ambiguous, "evidence": evidence}
		attachParent(resolved, rows, parents)
		row["resolved_structure"] = resolved
	}
	return map[string]any{"contract_version": ContractVersion, "paragraphs": len(rows), "headings": headings, "candidates": candidates, "ambiguities": ambiguities, "editorial_hierarchy_verified": false}
}

func fallbackRole(context, text string) string {
	if context != "" && context != "body" {
		return context
	}
	if strings.TrimSpace(text) == "" {
		return "blank"
	}
	return "body"
}

func semanticRole(row map[string]any) string {
	if text, _ := row["text"].(string); strings.TrimSpace(text) == "" {
		return ""
	}
	if row["context"] == "contents" {
		return "toc"
	}
	parts := []string{}
	if id, _ := row["style_id"].(string); id != "" {
		parts = append(parts, id)
	}
	if names, ok := row["style_names"].([]string); ok {
		parts = append(parts, names...)
	}
	style := strings.ToLower(strings.Join(parts, " "))
	switch {
	case strings.Contains(style, "toc heading"):
		return "toc"
	case strings.Contains(style, "quotation") || strings.Contains(style, "quote") || strings.Contains(style, "block text"):
		return "quotation"
	case strings.Contains(style, "abstract"):
		return "abstract"
	case strings.Contains(style, "author") || strings.Contains(style, "byline"):
		return "author"
	case !strings.Contains(style, "heading") && (strings.Contains(style, "document title") || strings.TrimSpace(style) == "title normal"):
		return "title"
	}
	return ""
}

func attachParent(resolved map[string]any, rows []map[string]any, parents [10]int) {
	for level := len(parents) - 1; level > 0; level-- {
		if paragraph := parents[level]; paragraph > 0 {
			resolved["parent_paragraph"] = paragraph
			resolved["parent_source_id"] = rows[paragraph-1]["source_id"]
			return
		}
	}
}
func clamp(n int) int {
	if n < 0 {
		return 0
	}
	if n > 100 {
		return 100
	}
	return n
}
func intValue(v any) (int, bool) {
	switch n := v.(type) {
	case int:
		return n, true
	case float64:
		return int(n), n == float64(int(n))
	}
	return 0, false
}
func headingStyle(s string) bool {
	s = strings.ToLower(s)
	return strings.Contains(s, "heading") || strings.Contains(s, "title") || strings.Contains(s, "head")
}

func headingLevelFromStyle(row map[string]any) int {
	parts := []string{}
	if id, _ := row["style_id"].(string); id != "" {
		parts = append(parts, id)
	}
	if names, ok := row["style_names"].([]string); ok && len(names) > 0 {
		parts = append(parts, names[0])
	}
	name := strings.ToLower(strings.Join(parts, " "))
	if !strings.Contains(name, "heading") && !strings.Contains(name, "outline") {
		return 0
	}
	for i := len(name) - 1; i >= 0; i-- {
		if name[i] >= '1' && name[i] <= '9' {
			return int(name[i] - '0')
		}
	}
	return 0
}
func listStyle(s string) bool { return strings.Contains(strings.ToLower(s), "list") }
func hasNativeHeading(row map[string]any) bool {
	outline, ok := row["outline_evidence"].(map[string]any)
	return ok && outline["body_text"] != true
}
func formattingSignal(row map[string]any) bool {
	if evidence, ok := row["direct_formatting_evidence"].(map[string]int); ok {
		total := evidence["text_units"]
		if total == 0 {
			return false
		}
		return evidence["bold_units"]*100 >= total*65 || evidence["caps_units"]*100 >= total*65 || evidence["small_caps_units"]*100 >= total*65
	}
	runs, _ := row["run_properties"].([]map[string]any)
	for _, run := range runs {
		x, _ := run["properties_xml"].(string)
		if strings.Contains(x, ":b") || strings.Contains(x, ":smallCaps") || strings.Contains(x, ":caps") {
			return true
		}
	}
	return false
}
func upperShare(s string) float64 {
	letters, upper := 0, 0
	for _, r := range s {
		if unicode.IsLetter(r) {
			letters++
			if unicode.IsUpper(r) {
				upper++
			}
		}
	}
	if letters == 0 {
		return 0
	}
	return float64(upper) / float64(letters)
}
func sentenceEnding(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	return strings.ContainsRune(".,;:!?", []rune(s)[len([]rune(s))-1])
}
