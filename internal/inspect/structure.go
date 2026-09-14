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
}

type numberingDefinition struct {
	AbstractID string
	Levels     map[int]numberingLevel
	Overrides  map[int]int
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
				item := numberingLevel{Start: 1}
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
							item.Start = n
						}
					case "numFmt":
						item.Family = x.Attribute(office.W, "val")
					case "lvlText":
						item.Pattern = x.Attribute(office.W, "val")
					}
				}
				levels[ilvl] = item
			}
			abstracts[s.Attribute(office.W, "abstractNumId")] = levels
		case "num":
			definition := numberingDefinition{Levels: map[int]numberingLevel{}, Overrides: map[int]int{}}
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
				for _, override := range spans[i+1+j+1:] {
					if override.Start >= x.End {
						break
					}
					if override.Name.Space == office.W && override.Name.Local == "startOverride" {
						if n, e := strconv.Atoi(override.Attribute(office.W, "val")); e == nil {
							definition.Overrides[ilvl] = n
						}
					}
				}
			}
			definitions[s.Attribute(office.W, "numId")] = definition
		}
	}
	for id, definition := range definitions {
		definition.Levels = abstracts[definition.AbstractID]
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
	rows, err := TextObservations(p)
	if err != nil {
		return nil, err
	}
	styles, err := office.XMLSpans(p.Files["word/styles.xml"])
	if err != nil {
		return nil, err
	}
	numbering, err := numberingDefinitions(p)
	if err != nil {
		return nil, err
	}
	type style struct {
		parent, outline, name, paragraphProperties, runProperties string
	}
	byID := map[string]style{}
	styleEvidence := map[string]map[string]any{}
	defaultID := ""
	for i, s := range styles {
		if s.Name.Space != office.W || s.Name.Local != "style" || s.Depth != 1 {
			continue
		}
		id := s.Attribute(office.W, "styleId")
		if s.Attribute(office.W, "type") == "paragraph" && s.Attribute(office.W, "default") == "1" {
			defaultID = id
		}
		def := style{}
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
			case x.Depth == s.Depth+2 && x.Name.Local == "outlineLvl":
				def.outline = x.Attribute(office.W, "val")
			}
		}
		byID[id] = def
		styleEvidence[id] = map[string]any{"name": def.name, "based_on": def.parent, "paragraph_properties_xml": def.paragraphProperties, "run_properties_xml": def.runProperties}
	}
	defaultRunProperties := ""
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
				break
			}
		}
		break
	}
	spans, err := office.XMLSpans(p.Files["word/document.xml"])
	if err != nil {
		return nil, err
	}
	ancestors := []office.XMLSpan{}
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
			outline, origin := "", ""
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
				chain = append(chain, at)
				styleNames = append(styleNames, def.name)
				if outline == "" && def.outline != "" {
					outline = def.outline
					origin = "style:" + at
				}
				at = def.parent
			}
			item["style_chain"] = chain
			item["style_names"] = styleNames
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
				if x.Name.Space == office.W && x.Depth == s.Depth+2 && x.Name.Local == "outlineLvl" {
					outline = x.Attribute(office.W, "val")
					origin = "direct paragraph formatting"
				}
			}
			numID, ilvl := "", 0
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
				case "ilvl":
					if n, parseErr := strconv.Atoi(x.Attribute(office.W, "val")); parseErr == nil {
						ilvl = n
					}
				}
			}
			if numID != "" && numID != "0" {
				evidence := map[string]any{"num_id": numID, "level": ilvl + 1, "provenance": "package XML"}
				if definition, ok := numbering[numID]; ok {
					evidence["abstract_num_id"] = definition.AbstractID
					if level, exists := definition.Levels[ilvl]; exists {
						evidence["family"] = level.Family
						evidence["label_pattern"] = level.Pattern
						evidence["start"] = level.Start
					}
					if start, overridden := definition.Overrides[ilvl]; overridden {
						evidence["start"] = start
						evidence["start_override"] = true
					}
				} else {
					evidence["unresolved"] = true
				}
				item["numbering_evidence"] = evidence
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
	return map[string]any{"source_sha256": office.Hash(b), "paragraphs": rows, "styles": styleEvidence, "document_default_run_properties_xml": defaultRunProperties, "numbering_xml": string(p.Files["word/numbering.xml"]), "xml_namespaces": map[string]string{"w": office.W}, "locator_units": "UTF-8 XML byte offsets; paragraph indexes include nested paragraphs, not native Word range positions", "editorial_hierarchy_verified": false, "scope": "main document XML; native outline evidence, style and formatting ancestry, and table/textbox containment; heading inference remains generic"}, nil
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
