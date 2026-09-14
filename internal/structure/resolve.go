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
	frontMatter := frontMatterRoles(rows)
	headings, candidates, ambiguities := 0, 0, 0
	for i, row := range rows {
		context, _ := row["context"].(string)
		text, _ := row["text"].(string)
		style, _ := row["style_id"].(string)
		score, level := 0, 0
		evidence := []string{}
		if role := semanticRole(row); role != "" {
			resolved := map[string]any{"role": role, "level": 0, "parent_paragraph": 0, "candidate_score": 100, "confidence": 100, "ambiguous": false, "evidence": []string{"semantic-style-or-context"}}
			if role == "quotation" || role == "abstract" {
				attachParent(resolved, rows, parents)
			}
			row["resolved_structure"] = resolved
			continue
		}
		if role := frontMatter[i]; role != "" {
			row["resolved_structure"] = map[string]any{"role": role, "level": 0, "parent_paragraph": 0, "candidate_score": 85, "confidence": 85, "ambiguous": false, "evidence": []string{"front-matter-layout"}}
			continue
		}
		if context != "body" || strings.TrimSpace(text) == "" {
			resolved := map[string]any{"role": fallbackRole(context, text), "level": 0, "parent_paragraph": 0, "candidate_score": score, "confidence": 100, "ambiguous": false, "evidence": evidence}
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
		numbered := false
		if numbering, ok := row["numbering_evidence"].(map[string]any); ok && numbering["family"] != "bullet" {
			if numbering["disabled"] == true {
				if level > 0 || headingStyle(style) || headingLevelFromStyle(row) > 0 {
					row["structure_contradiction"] = "heading evidence conflicts with explicit numId=0 (numbering disabled)"
				}
				evidence = append(evidence, "numbering-disabled")
			} else {
				numbered = true
				if n, ok := intValue(numbering["level"]); ok {
					if level == 0 {
						level = n
					} else if level != n {
						row["structure_contradiction"] = fmt.Sprintf("heading level %d conflicts with numbering level %d", level, n)
					}
				}
				score += 12
				evidence = append(evidence, "native-numbering")
			}
		}
		if sequence, ok := row["sequence_evidence"].(Assignment); ok && sequence.Action != "violation" {
			if level == 0 {
				level = sequence.Level
			}
			score += 15
			evidence = append(evidence, "sequence-"+sequence.Action)
		}
		props, _ := row["paragraph_properties_xml"].(string)
		keepNext := strings.Contains(props, ":keepNext")
		emphasized := formattingSignal(row)
		// A marker sequence is evidence, not permission to promote ordinary list
		// paragraphs. Native outline or a heading-style family can override this.
		listRisk := listStyle(style) || numbered && !headingStyle(style) && !keepNext && !emphasized
		if level > 0 && !hasNativeHeading(row) && listRisk {
			score -= 30
			evidence = append(evidence, "ordinary-list-risk")
		}
		if len([]rune(text)) <= 160 {
			score += 8
		} else {
			score -= 25
			evidence = append(evidence, "long-paragraph")
		}
		if keepNext {
			score += 8
			evidence = append(evidence, "keep-next")
		}
		if emphasized {
			score += 8
			evidence = append(evidence, "emphasis")
			familyCoherent := false
			if family, ok := row["style_family_evidence"].(map[string]any); ok {
				familyCoherent = family["coherent"] == true
			}
			if numbered && !familyCoherent && !headingStyle(style) && !hasNativeHeading(row) {
				row["structure_contradiction"] = "numbering conflicts with direct-format heading evidence"
			}
		}
		if upperShare(text) >= .8 {
			score += 8
			evidence = append(evidence, "uppercase-text")
		}
		if level >= 5 && !headingStyle(style) && row["marker_interpretations"] == nil {
			// Deep outline levels on an ordinary body style are commonly stale
			// residue from an editor's outline view. Keep the evidence, but do
			// not let the native-outline bonus promote the paragraph by itself.
			score -= 65
			evidence = append(evidence, "deep-outline-body-risk")
		}
		if sentenceEnding(text) && level == 0 {
			score -= 18
			evidence = append(evidence, "sentence-ending")
		}
		alternatives := markerAlternatives(row)
		contradictions := structureContradictions(row)
		ambiguous := len(alternatives) > 1 || len(contradictions) > 0
		if len(alternatives) > 1 {
			evidence = append(evidence, "ambiguous-marker")
		}
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
			resolved := map[string]any{"role": role, "level": level, "parent_paragraph": parent, "candidate_score": score, "confidence": clamp(score), "ambiguous": ambiguous, "evidence": evidence}
			addResolutionAlternatives(resolved, alternatives, contradictions)
			if parent > 0 {
				if sourceID := rows[parent-1]["source_id"]; sourceID != nil && sourceID != "" {
					resolved["parent_source_id"] = sourceID
				}
			}
			if ambiguous {
				ambiguities++
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
		}
		if row["structure_contradiction"] != nil {
			ambiguous = true
		}
		if ambiguous {
			ambiguities++
		}
		resolved := map[string]any{"role": role, "level": level, "candidate_score": score, "confidence": clamp(score), "ambiguous": ambiguous, "evidence": evidence}
		addResolutionAlternatives(resolved, alternatives, contradictions)
		attachParent(resolved, rows, parents)
		row["resolved_structure"] = resolved
	}
	return map[string]any{"contract_version": ContractVersion, "paragraphs": len(rows), "headings": headings, "candidates": candidates, "ambiguities": ambiguities, "editorial_hierarchy_verified": false}
}

func markerAlternatives(row map[string]any) []Interpretation {
	choices, ok := row["marker_interpretations"].([]Interpretation)
	if !ok || len(choices) == 0 {
		return nil
	}
	return append([]Interpretation(nil), choices...)
}

func structureContradictions(row map[string]any) []string {
	if text := stringValue(row["structure_contradiction"]); text != "" {
		return []string{text}
	}
	return nil
}

func addResolutionAlternatives(resolved map[string]any, alternatives []Interpretation, contradictions []string) {
	if len(alternatives) > 0 {
		resolved["alternatives"] = alternatives
	}
	if len(contradictions) > 0 {
		resolved["contradictions"] = contradictions
	}
}

func frontMatterRoles(rows []map[string]any) map[int]string {
	roles := map[int]string{}
	limit := len(rows)
	for i, row := range rows {
		if row["context"] != "body" || strings.TrimSpace(stringValue(row["text"])) == "" {
			continue
		}
		style := styleText(row)
		if (hasNativeHeading(row) || headingLevelFromStyle(row) > 0) && !strings.Contains(strings.ToLower(style), "title") {
			limit = i
			break
		}
	}
	if limit > 64 {
		limit = 64
	}
	firstContent := -1
	for i := 0; i < limit; i++ {
		if rows[i]["context"] == "body" && strings.TrimSpace(stringValue(rows[i]["text"])) != "" {
			firstContent = i
			break
		}
	}
	titleSeen, authorSeen := false, false
	for i := 0; i < limit; i++ {
		row := rows[i]
		text := strings.TrimSpace(stringValue(row["text"]))
		if text == "" || row["context"] != "body" {
			continue
		}
		style := styleText(row)
		normalizedStyle := strings.ToLower(style)
		centered := strings.EqualFold(stringValue(row["paragraph_alignment"]), "center") || strings.Contains(normalizedStyle, "centred") || strings.Contains(normalizedStyle, "centered")
		titleStyle := (styleTokenPresent(style, "document") && styleTokenPresent(style, "title")) ||
			(styleTokenPresent(style, "heading") && styleTokenPresent(style, "title")) || strings.EqualFold(strings.TrimSpace(style), "title")
		if !titleSeen && len([]rune(text)) <= 240 && (titleStyle || i == firstContent && (centered || formattingSignal(row))) {
			roles[i], titleSeen = "title", true
			continue
		}
		if titleSeen && !authorSeen && titleStyle {
			roles[i] = "title"
			continue
		}
		if titleSeen && !authorSeen && centered && len([]rune(text)) <= 240 && likelyAuthorLine(text) {
			roles[i], authorSeen = "author", true
			continue
		}
		if titleSeen && !authorSeen && centered && len([]rune(text)) <= 240 && formattingSignal(row) {
			roles[i] = "title"
		}
	}
	return roles
}

func styleText(row map[string]any) string {
	parts := []string{stringValue(row["style_id"])}
	if names, ok := row["style_names"].([]string); ok {
		parts = append(parts, names...)
	}
	return strings.Join(parts, " ")
}

func stringValue(value any) string {
	text, _ := value.(string)
	return text
}

func likelyAuthorLine(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	if strings.HasPrefix(lower, "by ") {
		return true
	}
	for _, word := range []string{"draft", "revised", "january", "february", "march", "april", "may ", "june", "july", "august", "september", "october", "november", "december"} {
		if strings.Contains(lower, word) {
			return false
		}
	}
	return strings.Contains(text, ",") || strings.Contains(text, "*") && strings.Contains(lower, " and ")
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
	text, _ := row["text"].(string)
	plain := strings.ToLower(strings.Trim(strings.TrimSpace(text), ":"))
	if plain == "" {
		return ""
	}
	if plain == "abstract" && !hasNativeHeading(row) && headingLevelFromStyle(row) == 0 {
		return "abstract"
	}
	if plain == "contents" || plain == "table of contents" {
		return "toc"
	}
	if row["context"] == "contents" {
		return "toc"
	}
	style := styleText(row)
	normalizedStyle := strings.ToLower(style)
	switch {
	case (strings.Contains(normalizedStyle, "tocheading") || (styleTokenPresent(style, "toc") && styleTokenPresent(style, "heading"))) && likelyTOCText(row, text):
		return "toc"
	case styleTokenPresent(style, "quotation") || styleTokenPresent(style, "quote") || (styleTokenPresent(style, "block") && styleTokenPresent(style, "text")):
		return "quotation"
	case styleTokenPresent(style, "abstract"):
		return "abstract"
	case styleTokenPresent(style, "author") || styleTokenPresent(style, "byline"):
		return "author"
	case !styleTokenPresent(style, "heading") && ((styleTokenPresent(style, "document") && styleTokenPresent(style, "title")) || strings.EqualFold(strings.TrimSpace(style), "title normal")):
		return "title"
	}
	return ""
}

// likelyTOCText prevents a repurposed TOC style from overriding stronger
// content/layout evidence. Word templates frequently reuse a built-in style
// for an abstract or other front matter. A real TOC heading is a short label,
// a TOC entry has tabular/page-number text, or the inspector has already
// classified the paragraph as part of the contents story.
func likelyTOCText(row map[string]any, text string) bool {
	if row["context"] == "contents" {
		return true
	}
	plain := strings.ToLower(strings.TrimSpace(text))
	if plain == "contents" || plain == "table of contents" || plain == "table of contents:" {
		return true
	}
	if strings.Contains(text, "\t") {
		return true
	}
	trimmed := strings.TrimSpace(text)
	return len([]rune(trimmed)) <= 96 && trailingPageNumber(trimmed)
}

func trailingPageNumber(text string) bool {
	runes := []rune(strings.TrimSpace(text))
	if len(runes) == 0 {
		return false
	}
	i := len(runes) - 1
	for i >= 0 && unicode.IsDigit(runes[i]) {
		i--
	}
	return i < len(runes)-1
}

func attachParent(resolved map[string]any, rows []map[string]any, parents [10]int) {
	for level := len(parents) - 1; level > 0; level-- {
		if paragraph := parents[level]; paragraph > 0 {
			resolved["parent_paragraph"] = paragraph
			if sourceID := rows[paragraph-1]["source_id"]; sourceID != nil && sourceID != "" {
				resolved["parent_source_id"] = sourceID
			}
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
	for _, token := range styleTokens(s) {
		switch token {
		case "heading", "title", "head":
			return true
		}
	}
	return false
}

// styleTokens treats separators, digit transitions and camel-case boundaries
// as style-name boundaries. Substring matching makes unrelated names such as
// Header, AheadBody and Headnote look like heading styles.
func styleTokens(s string) []string {
	runes := []rune(s)
	words := make([]string, 0, 4)
	word := make([]rune, 0, 12)
	flush := func() {
		if len(word) == 0 {
			return
		}
		words = append(words, strings.ToLower(string(word)))
		word = word[:0]
	}
	for i, r := range runes {
		if !unicode.IsLetter(r) && !unicode.IsDigit(r) {
			flush()
			continue
		}
		if len(word) > 0 {
			previous := runes[i-1]
			next := rune(0)
			if i+1 < len(runes) {
				next = runes[i+1]
			}
			if unicode.IsUpper(r) && (unicode.IsLower(previous) || unicode.IsUpper(previous) && unicode.IsLower(next)) ||
				unicode.IsDigit(r) != unicode.IsDigit(previous) {
				flush()
			}
		}
		word = append(word, r)
	}
	flush()
	return words
}

func headingLevelFromStyle(row map[string]any) int {
	parts := []string{}
	if id, _ := row["style_id"].(string); id != "" {
		parts = append(parts, id)
	}
	if names, ok := row["style_names"].([]string); ok && len(names) > 0 {
		parts = append(parts, names[0])
	}
	name := strings.Join(parts, " ")
	if !styleTokenPresent(name, "heading") && !styleTokenPresent(name, "outline") {
		return 0
	}
	for i := len(name) - 1; i >= 0; i-- {
		if name[i] >= '1' && name[i] <= '9' {
			return int(name[i] - '0')
		}
	}
	return 0
}

func styleTokenPresent(s, wanted string) bool {
	for _, token := range styleTokens(s) {
		if token == wanted {
			return true
		}
	}
	return false
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
