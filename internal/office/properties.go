package office

import (
	"fmt"
	"sort"
	"strings"
)

type propertyRule struct {
	elements, attributes string
	convert              func(any) (string, error)
}

func textValue(v any) (string, error) {
	text, ok := v.(string)
	if !ok || text == "" {
		return "", fmt.Errorf("non-empty string required")
	}
	return text, nil
}

func measurement(scale float64) func(any) (string, error) {
	return func(v any) (string, error) { return scaled(v, scale) }
}

func switchValue(on, off string) func(any) (string, error) {
	return func(v any) (string, error) {
		n, err := boolVal(v)
		if n == "1" {
			return on, err
		}
		return off, err
	}
}

// The same conversion and OOXML mapping is used for styles, paragraph/run
// content, numbering levels and building blocks. No per-surface option lists.
var runProperties = map[string]propertyRule{
	"style":      {"rStyle", "val", textValue},
	"font":       {"rFonts", "ascii hAnsi eastAsia cs", textValue},
	"size_pt":    {"sz szCs", "val", measurement(2)},
	"bold":       {"b", "val", boolVal},
	"italic":     {"i", "val", boolVal},
	"small_caps": {"smallCaps", "val", boolVal},
	"all_caps":   {"caps", "val", boolVal},
	"strike":     {"strike", "val", boolVal},
	"hidden":     {"vanish", "val", boolVal},
	"color":      {"color", "val", textValue},
	"highlight":  {"highlight", "val", textValue},
	"language":   {"lang", "val", textValue},
	"underline": {"u", "val", func(v any) (string, error) {
		if _, ok := v.(bool); ok {
			return switchValue("single", "none")(v)
		}
		return textValue(v)
	}},
	"superscript":           {"vertAlign", "val", switchValue("superscript", "baseline")},
	"subscript":             {"vertAlign", "val", switchValue("subscript", "baseline")},
	"double_strike":         {"dstrike", "val", boolVal},
	"outline":               {"outline", "val", boolVal},
	"shadow":                {"shadow", "val", boolVal},
	"emboss":                {"emboss", "val", boolVal},
	"imprint":               {"imprint", "val", boolVal},
	"no_proof":              {"noProof", "val", boolVal},
	"rtl":                   {"rtl", "val", boolVal},
	"complex_script":        {"cs", "val", boolVal},
	"bold_complex_script":   {"bCs", "val", boolVal},
	"italic_complex_script": {"iCs", "val", boolVal},
	"snap_to_grid":          {"snapToGrid", "val", boolVal},
	"character_spacing_pt":  {"spacing", "val", measurement(20)},
	"position_pt":           {"position", "val", measurement(2)},
	"kerning_pt":            {"kern", "val", measurement(2)},
}

var paragraphProperties = map[string]propertyRule{
	"style": {"pStyle", "val", textValue},
	"alignment": {"jc", "val", func(v any) (string, error) {
		text, err := textValue(v)
		if err != nil || !strings.Contains("|left|right|center|both|distribute|start|end|", "|"+text+"|") {
			return "", fmt.Errorf("invalid paragraph alignment")
		}
		return text, nil
	}},
	"keep_next":         {"keepNext", "val", boolVal},
	"keep_lines":        {"keepLines", "val", boolVal},
	"page_break_before": {"pageBreakBefore", "val", boolVal},
	"widow_control":     {"widowControl", "val", boolVal},
	"before_pt":         {"spacing", "before", measurement(20)},
	"after_pt":          {"spacing", "after", measurement(20)},
	"line_pt":           {"spacing", "line", measurement(20)},
	"line_multiple":     {"spacing", "line", measurement(240)},
	"left_pt":           {"ind", "left", measurement(20)},
	"right_pt":          {"ind", "right", measurement(20)},
	"first_line_pt":     {"ind", "firstLine", measurement(20)},
	"hanging_pt":        {"ind", "hanging", measurement(20)},
	"outline_level": {"outlineLvl", "val", func(v any) (string, error) {
		n, err := integer(v)
		if err != nil || n < 0 || n > 9 {
			return "", fmt.Errorf("invalid outline level")
		}
		return fmt.Sprint(n), nil
	}},
	"contextual_spacing":    {"contextualSpacing", "val", boolVal},
	"mirror_indents":        {"mirrorIndents", "val", boolVal},
	"suppress_line_numbers": {"suppressLineNumbers", "val", boolVal},
	"suppress_auto_hyphens": {"suppressAutoHyphens", "val", boolVal},
	"bidi":                  {"bidi", "val", boolVal},
	"snap_to_grid":          {"snapToGrid", "val", boolVal},
}

func props(raw []byte, kind string, spec map[string]any, scope ...map[string]string) ([]byte, error) {
	tag, order, rules := "rPr", runOrder, runProperties
	if kind == "paragraph" {
		tag, order, rules = "pPr", paraOrder, paragraphProperties
	}
	if len(raw) == 0 {
		raw = []byte(`<w:` + tag + ` xmlns:w="` + W + `"/>`)
	}
	for _, pair := range [][2]string{{"line_pt", "line_multiple"}, {"first_line_pt", "hanging_pt"}, {"superscript", "subscript"}} {
		if _, a := spec[pair[0]]; a {
			if _, b := spec[pair[1]]; b {
				return nil, fmt.Errorf("%s and %s are mutually exclusive", pair[0], pair[1])
			}
		}
	}
	changes := map[string]childChange{}
	attr := func(element, name, value string) {
		change := changes[element]
		if change.attrs == nil {
			change.attrs = map[string]string{}
		}
		change.attrs["w:"+name] = value
		changes[element] = change
	}
	keys := make([]string, 0, len(spec))
	for key := range spec {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	for _, key := range keys {
		if kind == "paragraph" && (key == "tabs" || key == "list_id" || key == "list_level") {
			continue
		}
		rule, ok := rules[key]
		if !ok {
			return nil, fmt.Errorf("unknown %s property %q; edit raw XML for other features", kind, key)
		}
		value, err := rule.convert(spec[key])
		if err != nil {
			return nil, fmt.Errorf("%s: %w", key, err)
		}
		for _, element := range strings.Fields(rule.elements) {
			for _, name := range strings.Fields(rule.attributes) {
				attr(element, name, value)
			}
		}
	}
	if _, ok := spec["font"]; ok {
		for _, name := range strings.Fields("asciiTheme hAnsiTheme eastAsiaTheme cstheme csTheme") {
			attr("rFonts", name, "\x00")
		}
	}
	for key, counterpart := range map[string]string{"first_line_pt": "hanging", "hanging_pt": "firstLine"} {
		if _, ok := spec[key]; ok {
			attr("ind", counterpart, "\x00")
			attr("ind", counterpart+"Chars", "\x00")
		}
	}
	for key, rule := range map[string]string{"line_pt": "exact", "line_multiple": "auto"} {
		if _, ok := spec[key]; ok {
			attr("spacing", "lineRule", rule)
		}
	}
	if kind == "paragraph" {
		if _, hasLevel := spec["list_level"]; hasLevel && spec["list_id"] == nil {
			return nil, fmt.Errorf("list_level requires list_id")
		}
		if value, ok := spec["list_id"]; ok {
			id, err := integer(value)
			if err != nil || id < 1 {
				return nil, fmt.Errorf("invalid list id")
			}
			level := 0
			if value, ok := spec["list_level"]; ok {
				level, err = integer(value)
				if err != nil || level < 0 || level > 8 {
					return nil, fmt.Errorf("invalid list level")
				}
			}
			changes["numPr"] = childChange{xml: fmt.Sprintf(`<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>`, level, id)}
		}
		if value, ok := spec["tabs"]; ok {
			tabs, ok := value.([]any)
			if !ok {
				return nil, fmt.Errorf("tabs must be array")
			}
			var xml strings.Builder
			xml.WriteString("<w:tabs>")
			for _, tab := range tabs {
				m, ok := tab.(map[string]any)
				if !ok {
					return nil, fmt.Errorf("invalid tab")
				}
				if err := KnownKeys(m, "position_pt alignment leader"); err != nil {
					return nil, err
				}
				position, err := scaled(m["position_pt"], 20)
				if err != nil {
					return nil, err
				}
				attrs := map[string]string{"w:pos": position, "w:val": "left"}
				for key, name := range map[string]string{"alignment": "w:val", "leader": "w:leader"} {
					if v, ok := m[key]; ok {
						attrs[name], err = textValue(v)
						if err != nil {
							return nil, err
						}
					}
				}
				xml.WriteString(attrsXML("w:tab", attrs))
			}
			xml.WriteString("</w:tabs>")
			changes["tabs"] = childChange{xml: xml.String()}
		}
	}
	return mergeChildren(raw, order, changes, scope...)
}
