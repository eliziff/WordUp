package inspect

import (
	"fmt"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/structure"
	"path/filepath"
	"strconv"
	"strings"
)

type numberingLevel struct {
	Family, Pattern string
	Start           int
	StartSpecified  bool
	RestartAfter    int
}

type numberingDefinition struct {
	AbstractID     string
	Levels         map[int]numberingLevel
	Overrides      map[int]int
	LevelOverrides map[int]numberingLevel
}

type numberingState struct {
	values [9]int
	seen   [9]bool
}

func numberingLabel(definition numberingDefinition, level int, state *numberingState) (string, bool) {
	if level < 0 || level >= len(state.values) {
		return "", false
	}
	item, ok := definition.Levels[level]
	if !ok || item.Pattern == "" || item.Family == "bullet" {
		return "", false
	}
	start := item.Start
	if override, exists := definition.Overrides[level]; exists {
		start = override
	}
	if state.seen[level] {
		state.values[level]++
	} else {
		state.values[level], state.seen[level] = start, true
	}
	for deeper := level + 1; deeper < len(state.values); deeper++ {
		item, exists := definition.Levels[deeper]
		if !exists {
			continue
		}
		// Word's lvlRestart value is one-based: a positive value restarts
		// after that level or any higher level. Zero means never restart;
		// omission means the previous level or any higher level. Values
		// greater than this level are ignored by Word.
		restart := item.RestartAfter
		reset := restart < 0 || restart > 0 && restart <= deeper && level <= restart-1
		if !reset {
			continue
		}
		state.values[deeper], state.seen[deeper] = 0, false
	}
	label, certain := item.Pattern, true
	for index := 0; index < len(state.values); index++ {
		placeholder := "%" + strconv.Itoa(index+1)
		if !strings.Contains(label, placeholder) {
			continue
		}
		part, exists := definition.Levels[index]
		if !exists || !state.seen[index] {
			certain = false
			continue
		}
		value, supported := numberingValue(state.values[index], part.Family)
		if !supported {
			certain = false
		}
		label = strings.ReplaceAll(label, placeholder, value)
	}
	return label, certain
}

func numberingValue(value int, family string) (string, bool) {
	switch family {
	case "decimal":
		return strconv.Itoa(value), true
	case "decimalZero":
		if value >= 0 && value < 10 {
			return "0" + strconv.Itoa(value), true
		}
		return strconv.Itoa(value), true
	case "upperLetter", "lowerLetter":
		if value < 1 {
			return strconv.Itoa(value), false
		}
		var out string
		for value > 0 {
			value--
			out = string(rune('A'+value%26)) + out
			value /= 26
		}
		if family == "lowerLetter" {
			out = strings.ToLower(out)
		}
		return out, true
	case "upperRoman", "lowerRoman":
		if value < 1 || value > 3999 {
			return strconv.Itoa(value), false
		}
		var out strings.Builder
		for _, part := range []struct {
			value int
			text  string
		}{{1000, "M"}, {900, "CM"}, {500, "D"}, {400, "CD"}, {100, "C"}, {90, "XC"}, {50, "L"}, {40, "XL"}, {10, "X"}, {9, "IX"}, {5, "V"}, {4, "IV"}, {1, "I"}} {
			for value >= part.value {
				out.WriteString(part.text)
				value -= part.value
			}
		}
		if family == "lowerRoman" {
			return strings.ToLower(out.String()), true
		}
		return out.String(), true
	default:
		return strconv.Itoa(value), false
	}
}

const drawingML = "http://schemas.openxmlformats.org/drawingml/2006/main"

func themeFonts(p *office.Package) (map[string]string, error) {
	b := p.Files["word/theme/theme1.xml"]
	if len(b) == 0 {
		return map[string]string{}, nil
	}
	spans, err := office.XMLSpans(b)
	if err != nil {
		return nil, err
	}
	out := map[string]string{}
	ancestors := []office.XMLSpan{}
	for _, span := range spans {
		for len(ancestors) > 0 && ancestors[len(ancestors)-1].End <= span.Start {
			ancestors = ancestors[:len(ancestors)-1]
		}
		if span.Name.Space == drawingML && (span.Name.Local == "latin" || span.Name.Local == "ea" || span.Name.Local == "cs") {
			family := ""
			for i := len(ancestors) - 1; i >= 0; i-- {
				if ancestors[i].Name.Space == drawingML && (ancestors[i].Name.Local == "majorFont" || ancestors[i].Name.Local == "minorFont") {
					family = strings.TrimSuffix(ancestors[i].Name.Local, "Font")
					break
				}
			}
			if family != "" && span.Attribute("", "typeface") != "" {
				key := map[string]string{"latin": "HAnsi", "ea": "EastAsia", "cs": "Bidi"}[span.Name.Local]
				out[family+key] = span.Attribute("", "typeface")
			}
		}
		ancestors = append(ancestors, span)
	}
	return out, nil
}

func resolvedFonts(values, theme map[string]string) map[string]string {
	out := map[string]string{}
	for key, value := range values {
		out[key] = value
	}
	for _, key := range []string{"asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme", "cstheme"} {
		if token := values[key]; token != "" && theme[token] != "" {
			out[strings.TrimSuffix(strings.TrimSuffix(key, "Theme"), "theme")] = theme[token]
		}
	}
	return out
}

func numberingDefinitions(p *office.Package) (map[string]numberingDefinition, error) {
	b := p.Files["word/numbering.xml"]
	if len(b) == 0 {
		return map[string]numberingDefinition{}, nil
	}
	spans, err := office.XMLSpans(b)
	if err != nil {
		return nil, err
	}
	abstracts := map[string]map[int]numberingLevel{}
	definitions := map[string]numberingDefinition{}
	for i, s := range spans {
		if s.Name.Space != office.W || s.Depth != 1 {
			continue
		}
		switch s.Name.Local {
		case "abstractNum":
			levels := map[int]numberingLevel{}
			for j, level := range spans[i+1:] {
				if level.Start >= s.End {
					break
				}
				if level.Name.Space != office.W || level.Name.Local != "lvl" || level.Depth != s.Depth+1 {
					continue
				}
				ilvl, parseErr := strconv.Atoi(level.Attribute(office.W, "ilvl"))
				if parseErr != nil {
					continue
				}
				item := numberingLevel{Start: 1, RestartAfter: -1}
				for _, x := range spans[i+1+j+1:] {
					if x.Start >= level.End {
						break
					}
					if x.Name.Space != office.W {
						continue
					}
					switch x.Name.Local {
					case "start":
						if n, e := strconv.Atoi(x.Attribute(office.W, "val")); e == nil {
							item.Start, item.StartSpecified = n, true
						}
					case "numFmt":
						item.Family = x.Attribute(office.W, "val")
					case "lvlText":
						item.Pattern = x.Attribute(office.W, "val")
					case "lvlRestart":
						if n, e := strconv.Atoi(x.Attribute(office.W, "val")); e == nil && n >= 0 {
							item.RestartAfter = n
						}
					}
				}
				levels[ilvl] = item
			}
			abstracts[s.Attribute(office.W, "abstractNumId")] = levels
		case "num":
			definition := numberingDefinition{Levels: map[int]numberingLevel{}, Overrides: map[int]int{}, LevelOverrides: map[int]numberingLevel{}}
			for j, x := range spans[i+1:] {
				if x.Start >= s.End {
					break
				}
				if x.Name.Space != office.W {
					continue
				}
				if x.Name.Local == "abstractNumId" && x.Depth == s.Depth+1 {
					definition.AbstractID = x.Attribute(office.W, "val")
				}
				if x.Name.Local != "lvlOverride" || x.Depth != s.Depth+1 {
					continue
				}
				ilvl, parseErr := strconv.Atoi(x.Attribute(office.W, "ilvl"))
				if parseErr != nil {
					continue
				}
				for k, override := range spans[i+1+j+1:] {
					if override.Start >= x.End {
						break
					}
					if override.Name.Space == office.W && override.Name.Local == "startOverride" {
						if n, e := strconv.Atoi(override.Attribute(office.W, "val")); e == nil {
							definition.Overrides[ilvl] = n
						}
					}
					if override.Name.Space != office.W || override.Name.Local != "lvl" || override.Depth != x.Depth+1 {
						continue
					}
					item := numberingLevel{Start: 1, RestartAfter: -1}
					for _, child := range spans[i+1+j+1+k+1:] {
						if child.Start >= override.End {
							break
						}
						if child.Name.Space != office.W || child.Depth != override.Depth+1 {
							continue
						}
						switch child.Name.Local {
						case "start":
							if n, e := strconv.Atoi(child.Attribute(office.W, "val")); e == nil {
								item.Start, item.StartSpecified = n, true
							}
						case "numFmt":
							item.Family = child.Attribute(office.W, "val")
						case "lvlText":
							item.Pattern = child.Attribute(office.W, "val")
						}
					}
					definition.LevelOverrides[ilvl] = item
				}
			}
			definitions[s.Attribute(office.W, "numId")] = definition
		}
	}
	for id, definition := range definitions {
		baseLevels := abstracts[definition.AbstractID]
		definition.Levels = make(map[int]numberingLevel, len(baseLevels))
		for level, item := range baseLevels {
			definition.Levels[level] = item
		}
		for level, override := range definition.LevelOverrides {
			item := definition.Levels[level]
			if override.Family != "" {
				item.Family = override.Family
			}
			if override.Pattern != "" {
				item.Pattern = override.Pattern
			}
			if override.StartSpecified {
				item.Start, item.StartSpecified = override.Start, true
			}
			definition.Levels[level] = item
		}
		definitions[id] = definition
	}
	return definitions, nil
}

// StructureReference resolves paragraph outline inheritance without starting Word.
// It retains the raw evidence; an outline claim is not editorial acceptance.
func StructureReference(file string) (map[string]any, error) {
	b, err := project.Read(filepath.Dir(file), filepath.Base(file))
	if err != nil {
		return nil, err
	}
	p, err := office.ReadPackage(b)
	if err != nil {
		return nil, err
	}
	documentSpans, err := office.XMLSpans(p.Files["word/document.xml"])
	if err != nil {
		return nil, err
	}
	rows, _, err := textObservationsPartWithSpans(p, "word/document.xml", documentSpans)
	if err != nil {
		return nil, err
	}
	var styles []office.XMLSpan
	if rawStyles := p.Files["word/styles.xml"]; len(rawStyles) > 0 {
		styles, err = office.XMLSpans(rawStyles)
		if err != nil {
			return nil, err
		}
	}
	numbering, err := numberingDefinitions(p)
	if err != nil {
		return nil, err
	}
	fonts, err := themeFonts(p)
	if err != nil {
		return nil, err
	}
	type style struct {
		parent, outline, name, paragraphProperties, runProperties string
		fonts                                                     map[string]string
		numID                                                     string
		ilvl                                                      int
		hasIlvl                                                   bool
	}
	byID := map[string]style{}
	styleEvidence := map[string]map[string]any{}
	styleSources := map[string]map[string]any{}
	styleDuplicates := map[string][]map[string]any{}
	styleDuplicateOrder := []string{}
	styleConflicts := map[string]bool{}
	defaultID := ""
	for i, s := range styles {
		if s.Name.Space != office.W || s.Name.Local != "style" || s.Depth != 1 {
			continue
		}
		id := s.Attribute(office.W, "styleId")
		if s.Attribute(office.W, "type") == "paragraph" && s.Attribute(office.W, "default") == "1" {
			defaultID = id
		}
		def := style{fonts: map[string]string{}}
		for _, x := range styles[i+1:] {
			if x.Start >= s.End {
				break
			}
			if x.Name.Space != office.W {
				continue
			}
			switch {
			case x.Depth == s.Depth+1 && x.Name.Local == "basedOn":
				def.parent = x.Attribute(office.W, "val")
			case x.Depth == s.Depth+1 && x.Name.Local == "name":
				def.name = x.Attribute(office.W, "val")
			case x.Depth == s.Depth+1 && x.Name.Local == "pPr":
				def.paragraphProperties = string(p.Files["word/styles.xml"][x.Start:x.End])
			case x.Depth == s.Depth+1 && x.Name.Local == "rPr":
				def.runProperties = string(p.Files["word/styles.xml"][x.Start:x.End])
			case x.Depth == s.Depth+2 && x.Name.Local == "rFonts":
				for _, key := range []string{"ascii", "hAnsi", "eastAsia", "cs", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme", "cstheme"} {
					if value := x.Attribute(office.W, key); value != "" {
						def.fonts[key] = value
					}
				}
			case x.Depth == s.Depth+2 && x.Name.Local == "outlineLvl":
				def.outline = x.Attribute(office.W, "val")
			case x.Depth == s.Depth+3 && x.Name.Local == "numId":
				def.numID = x.Attribute(office.W, "val")
			case x.Depth == s.Depth+3 && x.Name.Local == "ilvl":
				if n, e := strconv.Atoi(x.Attribute(office.W, "val")); e == nil {
					def.ilvl, def.hasIlvl = n, true
				}
			}
		}
		byID[id] = def
		raw := p.Files["word/styles.xml"][s.Start:s.End]
		source := map[string]any{"start": s.Start, "end": s.End, "sha256": office.Hash(raw)}
		if first, exists := styleSources[id]; exists {
			if _, recorded := styleDuplicates[id]; !recorded {
				styleDuplicateOrder = append(styleDuplicateOrder, id)
				styleDuplicates[id] = []map[string]any{first}
			}
			styleDuplicates[id] = append(styleDuplicates[id], source)
			if first["sha256"] != source["sha256"] {
				styleConflicts[id] = true
			}
		} else {
			styleSources[id] = source
		}
		evidence := map[string]any{"name": def.name, "based_on": def.parent, "paragraph_properties_xml": def.paragraphProperties, "run_properties_xml": def.runProperties, "fonts": resolvedFonts(def.fonts, fonts)}
		if def.numID != "" {
			evidence["num_id"] = def.numID
			if def.hasIlvl {
				evidence["numbering_level"] = def.ilvl + 1
			}
		}
		styleEvidence[id] = evidence
	}
	defaultRunProperties := ""
	defaultFonts := map[string]string{}
	for i, s := range styles {
		if s.Name.Space != office.W || s.Name.Local != "rPrDefault" {
			continue
		}
		for _, x := range styles[i+1:] {
			if x.Start >= s.End {
				break
			}
			if x.Name.Space == office.W && x.Name.Local == "rPr" && x.Depth == s.Depth+1 {
				defaultRunProperties = string(p.Files["word/styles.xml"][x.Start:x.End])
			}
			if x.Name.Space == office.W && x.Name.Local == "rFonts" && x.Depth == s.Depth+2 {
				for _, key := range []string{"ascii", "hAnsi", "eastAsia", "cs", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme", "cstheme"} {
					if value := x.Attribute(office.W, key); value != "" {
						defaultFonts[key] = value
					}
				}
			}
		}
		break
	}
	for _, item := range rows {
		runs, _ := item["run_properties"].([]map[string]any)
		for _, run := range runs {
			values, _ := run["fonts"].(map[string]string)
			resolved := resolvedFonts(values, fonts)
			if len(resolved) > 0 {
				run["fonts"] = resolved
			}
		}
	}
	spans := documentSpans
	ancestors := []office.XMLSpan{}
	numberingStates := map[string]*numberingState{}
	row := 0
	for i, s := range spans {
		for len(ancestors) > 0 && ancestors[len(ancestors)-1].End <= s.Start {
			ancestors = ancestors[:len(ancestors)-1]
		}
		if s.Name.Space == office.W && s.Name.Local == "p" {
			item := rows[row]
			row++
			context := "body"
			for _, a := range ancestors {
				if a.Name.Space != office.W {
					continue
				}
				if a.Name.Local == "tbl" {
					context = "table"
				}
				if a.Name.Local == "txbxContent" {
					context = "textbox"
				}
			}
			item["context"] = context
			id := item["style_id"].(string)
			if id == "" {
				id = defaultID
			}
			chain := []string{}
			styleNames := []string{}
			seen := map[string]bool{}
			styleAmbiguities := []string{}
			outline, origin := "", ""
			styleNumID, styleNumOrigin, styleIlvl := "", "", 0
			styleHasIlvl := false
			for at := id; at != ""; {
				if seen[at] {
					item["style_error"] = "cyclic style inheritance at " + at
					break
				}
				seen[at] = true
				def, ok := byID[at]
				if !ok {
					item["style_error"] = "missing style " + at
					break
				}
				if styleConflicts[at] {
					styleAmbiguities = append(styleAmbiguities, at)
				}
				chain = append(chain, at)
				styleNames = append(styleNames, def.name)
				if outline == "" && def.outline != "" {
					outline = def.outline
					origin = "style:" + at
				}
				if styleNumID == "" && def.numID != "" {
					styleNumID, styleNumOrigin = def.numID, "style:"+at
					styleIlvl, styleHasIlvl = def.ilvl, def.hasIlvl
				}
				at = def.parent
			}
			item["style_chain"] = chain
			item["style_names"] = styleNames
			if len(styleAmbiguities) > 0 {
				item["style_ambiguity"] = map[string]any{
					"kind":       "conflicting_duplicate_style_id",
					"style_ids":  styleAmbiguities,
					"resolution": "source-order fallback; inspect style_duplicates before treating inherited properties as effective",
				}
			}
			if context == "body" {
				for index, ancestor := range chain {
					name := ""
					if index < len(styleNames) {
						name = styleNames[index]
					}
					for level := 1; level <= 9; level++ {
						if strings.EqualFold(ancestor, fmt.Sprintf("TOC%d", level)) || strings.EqualFold(name, fmt.Sprintf("toc %d", level)) {
							item["context"] = "contents"
						}
					}
				}
			}
			for _, x := range spans[i+1:] {
				if x.Start >= s.End {
					break
				}
				if x.Name.Space == office.W && x.Depth == s.Depth+2 {
					switch x.Name.Local {
					case "outlineLvl":
						outline = x.Attribute(office.W, "val")
						origin = "direct paragraph formatting"
					case "jc":
						item["paragraph_alignment"] = x.Attribute(office.W, "val")
					}
				}
			}
			numID, ilvl := "", 0
			numIDSpecified, ilvlSpecified := false, false
			for _, x := range spans[i+1:] {
				if x.Start >= s.End {
					break
				}
				if x.Name.Space != office.W || x.Depth != s.Depth+3 {
					continue
				}
				switch x.Name.Local {
				case "numId":
					numID = x.Attribute(office.W, "val")
					numIDSpecified = true
				case "ilvl":
					if n, parseErr := strconv.Atoi(x.Attribute(office.W, "val")); parseErr == nil {
						ilvl = n
						ilvlSpecified = true
					}
				}
			}
			numberingOrigin := "direct paragraph formatting"
			if !numIDSpecified {
				numID, numberingOrigin = styleNumID, styleNumOrigin
			}
			if !ilvlSpecified && styleHasIlvl {
				ilvl = styleIlvl
			}
			if numID != "" && numID != "0" {
				evidence := map[string]any{"num_id": numID, "level": ilvl + 1, "provenance": "package XML", "origin": numberingOrigin}
				if definition, ok := numbering[numID]; ok {
					evidence["abstract_num_id"] = definition.AbstractID
					if level, exists := definition.Levels[ilvl]; exists {
						evidence["family"] = level.Family
						evidence["label_pattern"] = level.Pattern
						evidence["start"] = level.Start
						if level.RestartAfter >= 0 {
							evidence["restart_after_level"] = level.RestartAfter
						}
					}
					if start, overridden := definition.Overrides[ilvl]; overridden {
						evidence["start"] = start
						evidence["start_override"] = true
					}
					if _, overridden := definition.LevelOverrides[ilvl]; overridden {
						evidence["level_override"] = true
					}
					state := numberingStates[numID]
					if state == nil {
						state = &numberingState{}
						numberingStates[numID] = state
					}
					if label, certain := numberingLabel(definition, ilvl, state); label != "" {
						evidence["displayed_label"] = label
						evidence["displayed_label_provenance"] = "package XML reconstruction"
						if !certain {
							evidence["displayed_label_uncertain"] = true
						}
					}
				} else {
					evidence["unresolved"] = true
				}
				item["numbering_evidence"] = evidence
			} else if numID == "0" && numIDSpecified {
				// Word uses an explicit numId=0 to disable numbering. Keep that
				// fact instead of dropping it: a heading style/outline paired
				// with disabled numbering is useful contradiction evidence.
				item["numbering_evidence"] = map[string]any{
					"num_id": numID, "level": ilvl + 1, "disabled": true,
					"provenance": "package XML", "origin": numberingOrigin,
				}
			}
			if outline != "" {
				level, e := strconv.Atoi(outline)
				if e != nil || level < 0 || level > 9 {
					item["outline_error"] = fmt.Sprintf("invalid outline level %q", outline)
				} else {
					item["outline_evidence"] = map[string]any{"level": level + 1, "origin": origin, "body_text": level == 9}
				}
			}
		}
		ancestors = append(ancestors, s)
	}
	// Repeated native headings can corroborate a style family, including
	// unnumbered headings. Conflicting levels remain visible, never averaged.
	votes := map[string]map[int]int{}
	for _, item := range rows {
		if item["context"] != "body" {
			continue
		}
		outline, ok := item["outline_evidence"].(map[string]any)
		if !ok || outline["body_text"] == true {
			continue
		}
		id := item["style_id"].(string)
		if id == "" {
			continue
		}
		if votes[id] == nil {
			votes[id] = map[int]int{}
		}
		votes[id][outline["level"].(int)]++
	}
	for _, item := range rows {
		if item["context"] != "body" {
			continue
		}
		family := votes[item["style_id"].(string)]
		if len(family) == 0 {
			continue
		}
		evidence := map[string]any{"level_votes": family, "coherent": false}
		if len(family) == 1 {
			for level, count := range family {
				if count >= 2 {
					evidence["coherent"] = true
					evidence["corroborated_level"] = level
					if outline, ok := item["outline_evidence"].(map[string]any); !ok || outline["level"] != level {
						evidence["requires_review"] = true
						evidence["reason"] = "paragraph outline differs from repeated native headings using the same style; direct override may be intentional"
					}
				}
			}
		} else {
			evidence["requires_review"] = true
			evidence["reason"] = "same style has conflicting native heading levels"
		}
		item["style_family_evidence"] = evidence
	}
	// Sequence observations are kept separate from native outline claims.
	candidates := [][]structure.Interpretation{}
	candidateRows := []int{}
	for i, item := range rows {
		if item["context"] != "body" {
			continue
		}
		choices := structure.MarkerChoices(item["text"].(string))
		if len(choices) == 0 {
			continue
		}
		item["marker_interpretations"] = choices
		candidateRows = append(candidateRows, i)
		candidates = append(candidates, choices)
	}
	for i, assignment := range structure.HeadingLadder(candidates) {
		rows[candidateRows[i]]["sequence_evidence"] = assignment
		rows[candidateRows[i]]["sequence_is_heading_claim"] = false
	}
	out := map[string]any{"source_sha256": office.Hash(b), "paragraphs": rows, "styles": styleEvidence, "document_default_run_properties_xml": defaultRunProperties, "document_default_fonts": resolvedFonts(defaultFonts, fonts), "theme_fonts": fonts, "xml_namespaces": map[string]string{"w": office.W}, "locator_units": "UTF-8 XML byte offsets; paragraph indexes include nested paragraphs, not native Word range positions", "editorial_hierarchy_verified": false, "scope": "main document XML; native outline evidence, style and formatting ancestry, theme fonts, and table/textbox containment; heading inference remains generic"}
	// Keep the exact style inputs beside the derived evidence. These are the
	// authoring source of truth for minute formatting (including properties this
	// generic detector does not interpret), not a second recipe representation.
	for _, source := range []struct {
		part    string
		key     string
		hashKey string
	}{
		{part: "word/styles.xml", key: "styles_xml", hashKey: "styles_sha256"},
		{part: "word/numbering.xml", key: "numbering_xml", hashKey: "numbering_sha256"},
		{part: "word/theme/theme1.xml", key: "theme_xml", hashKey: "theme_sha256"},
		{part: "word/fontTable.xml", key: "font_table_xml", hashKey: "font_table_sha256"},
		{part: "word/settings.xml", key: "settings_xml", hashKey: "settings_sha256"},
	} {
		if raw := p.Files[source.part]; len(raw) > 0 {
			out[source.key] = string(raw)
			out[source.hashKey] = office.Hash(raw)
		}
	}
	if document := p.Files["word/document.xml"]; len(document) > 0 {
		out["document_xml_sha256"] = office.Hash(document)
	}
	if len(styleDuplicateOrder) > 0 {
		duplicates := make([]map[string]any, 0, len(styleDuplicateOrder))
		for _, id := range styleDuplicateOrder {
			duplicates = append(duplicates, map[string]any{"style_id": id, "conflicting": styleConflicts[id], "definitions": styleDuplicates[id]})
		}
		out["style_duplicates"] = duplicates
	}
	return out, nil
}

func StructureResolved(file string) (map[string]any, error) {
	out, err := StructureReference(file)
	if err != nil {
		return nil, err
	}
	rows, ok := out["paragraphs"].([]map[string]any)
	if !ok {
		return nil, fmt.Errorf("structure paragraph contract unavailable")
	}
	out["resolution"] = structure.Resolve(rows)
	out["editorial_hierarchy_verified"] = false
	out["scope"] = "main document XML; generic structure resolution with explicit evidence and ambiguity; publication mapping remains separate"
	return out, nil
}
