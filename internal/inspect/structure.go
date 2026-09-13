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
	type style struct {
		parent     string
		outline    string
		name       string
		properties string
	}
	byID := map[string]style{}
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
				def.properties = string(p.Files["word/styles.xml"][x.Start:x.End])
			case x.Depth == s.Depth+2 && x.Name.Local == "outlineLvl":
				def.outline = x.Attribute(office.W, "val")
			}
		}
		byID[id] = def
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
			chain := []map[string]any{}
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
				chain = append(chain, map[string]any{"id": at, "name": def.name, "paragraph_properties_xml": def.properties})
				if outline == "" && def.outline != "" {
					outline = def.outline
					origin = "style:" + at
				}
				at = def.parent
			}
			item["style_chain"] = chain
			if context == "body" {
				for _, ancestor := range chain {
					for level := 1; level <= 9; level++ {
						if strings.EqualFold(ancestor["id"].(string), fmt.Sprintf("TOC%d", level)) || strings.EqualFold(ancestor["name"].(string), fmt.Sprintf("toc %d", level)) {
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
	return map[string]any{"source_sha256": office.Hash(b), "paragraphs": rows, "numbering_xml": string(p.Files["word/numbering.xml"]), "xml_namespaces": map[string]string{"w": office.W}, "locator_units": "UTF-8 XML byte offsets; paragraph indexes include nested paragraphs, not native Word range positions", "editorial_hierarchy_verified": false, "scope": "main document XML; native outline evidence, style ancestry and table/textbox containment; not a complete formatting cascade or heading inference engine"}, nil
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
