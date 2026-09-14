package office

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"regexp"
	"sort"
	"unicode/utf8"
)

type XMLAttributeEquivalence struct {
	Element   xml.Name `json:"element"`
	Attribute xml.Name `json:"attribute"`
	Values    []string `json:"values"`
}

type XMLComparePolicy struct {
	GeneratedCommentIDs   bool                      `json:"generated_comment_ids,omitempty"`
	AttributeEquivalences []XMLAttributeEquivalence `json:"attribute_equivalences,omitempty"`
	GeneratedTOCBookmarks bool                      `json:"generated_toc_bookmarks,omitempty"`
	Attributes            []xml.Name                `json:"attributes,omitempty"`
	Elements              []xml.Name                `json:"elements,omitempty"`
	ContextParts          []string                  `json:"context_parts,omitempty"`
}

type comparisonToken struct {
	text       string
	start, end int64
}

// Compare namespace-aware tokens, ignoring only explicitly supplied names.
func CompareXML(a, b []byte, policy XMLComparePolicy) (map[string]any, error) {
	tokens := func(raw []byte) ([]comparisonToken, error) {
		if len(raw) > Limit {
			return nil, fmt.Errorf("XML budget exceeded")
		}
		var comments *commentIDs
		if policy.GeneratedCommentIDs {
			var err error
			comments, err = readCommentIDs(raw)
			if err != nil {
				return nil, err
			}
		}
		tocNames := map[string]string{}
		if policy.GeneratedTOCBookmarks {
			scan := xml.NewDecoder(bytes.NewReader(raw))
			for {
				token, err := scan.Token()
				if err == io.EOF {
					break
				}
				if err != nil {
					return nil, err
				}
				if el, ok := token.(xml.StartElement); ok && el.Name == (xml.Name{Space: wordXMLNamespace, Local: "bookmarkStart"}) {
					for _, at := range el.Attr {
						if at.Name == (xml.Name{Space: wordXMLNamespace, Local: "name"}) && generatedTOCName.MatchString(at.Value) {
							if _, exists := tocNames[at.Value]; exists {
								return nil, fmt.Errorf("duplicate generated TOC bookmark %q", at.Value)
							}
							tocNames[at.Value] = fmt.Sprintf("\x00WordUpTOC%d", len(tocNames))
						}
					}
				}
			}
		}
		d := xml.NewDecoder(bytes.NewReader(raw))
		inInstruction := false
		structureDepth, roots := 0, 0
		ignored := func(names []xml.Name, n xml.Name) bool {
			for _, x := range names {
				if x == n {
					return true
				}
			}
			return false
		}
		out := []comparisonToken{}
		skip := 0
		for {
			start := d.InputOffset()
			appendToken := func(text string) { out = append(out, comparisonToken{text, start, d.InputOffset()}) }
			t, e := d.Token()
			if e == io.EOF {
				break
			}
			if e != nil {
				return nil, e
			}
			switch t := t.(type) {
			case xml.Directive:
				return nil, fmt.Errorf("XML directives/DTD are not accepted")
			case xml.StartElement:
				if structureDepth == 0 {
					roots++
					if roots > 1 {
						return nil, fmt.Errorf("multiple XML roots")
					}
				}
				structureDepth++
				inInstruction = t.Name == (xml.Name{Space: wordXMLNamespace, Local: "instrText"})
				if skip > 0 || ignored(policy.Elements, t.Name) {
					skip++
					continue
				}
				attrs := []string{}
				for _, at := range t.Attr {
					if comments != nil {
						var err error
						at.Value, err = comments.canonical(t.Name, at)
						if err != nil {
							return nil, err
						}
					}
					if at.Name.Space == "xmlns" || at.Name.Space == "" && at.Name.Local == "xmlns" || ignored(policy.Attributes, at.Name) {
						continue
					}
					if (t.Name == (xml.Name{Space: wordXMLNamespace, Local: "bookmarkStart"}) && at.Name == (xml.Name{Space: wordXMLNamespace, Local: "name"})) || (t.Name == (xml.Name{Space: wordXMLNamespace, Local: "hyperlink"}) && at.Name == (xml.Name{Space: wordXMLNamespace, Local: "anchor"})) {
						if canonical, ok := tocNames[at.Value]; ok {
							at.Value = canonical
						}
					}
					for ruleIndex, rule := range policy.AttributeEquivalences {
						if rule.Element != t.Name || rule.Attribute != at.Name {
							continue
						}
						for _, value := range rule.Values {
							if at.Value == value {
								at.Value = fmt.Sprintf("\x00WordUpEquivalent%d", ruleIndex)
								break
							}
						}
					}
					attrs = append(attrs, fmt.Sprintf("{%s}%s=%q", at.Name.Space, at.Name.Local, at.Value))
				}
				sort.Strings(attrs)
				appendToken(fmt.Sprintf("start {%s}%s %q", t.Name.Space, t.Name.Local, attrs))
			case xml.EndElement:
				if structureDepth == 0 {
					return nil, fmt.Errorf("XML stack underflow")
				}
				structureDepth--
				inInstruction = false
				if skip > 0 {
					skip--
					continue
				}
				appendToken(fmt.Sprintf("end {%s}%s", t.Name.Space, t.Name.Local))
			case xml.CharData:
				if structureDepth == 0 && len(bytes.TrimSpace(t)) > 0 {
					return nil, fmt.Errorf("text outside XML root")
				}
				if skip == 0 {
					value := string(t)
					if inInstruction && len(tocNames) > 0 {
						if match := generatedTOCReference.FindStringSubmatchIndex(value); match != nil {
							start, end := match[2], match[3]
							if canonical, ok := tocNames[value[start:end]]; ok {
								value = value[:start] + canonical + value[end:]
							}
						}
					}
					appendToken(fmt.Sprintf("text %q", value))
				}
			case xml.Comment:
				if skip == 0 {
					appendToken(fmt.Sprintf("comment %q", string(t)))
				}
			case xml.ProcInst:
				if skip == 0 && t.Target != "xml" {
					appendToken(fmt.Sprintf("instruction %s %q", t.Target, t.Inst))
				}
			}
		}
		if roots != 1 || structureDepth != 0 {
			return nil, fmt.Errorf("incomplete XML document")
		}
		return out, nil
	}
	left, e := tokens(a)
	if e != nil {
		return nil, e
	}
	right, e := tokens(b)
	if e != nil {
		return nil, e
	}
	result := map[string]any{"equal": true, "reference_sha256": Hash(a), "candidate_sha256": Hash(b), "byte_identical": bytes.Equal(a, b), "policy": policy}
	if bytes.Contains(a, []byte("http://schemas.microsoft.com/office/2006/xmlPackage")) || bytes.Contains(b, []byte("http://schemas.microsoft.com/office/2006/xmlPackage")) {
		parts := func(raw []byte) (map[string]string, map[string]XMLSpan, error) {
			spans, err := XMLSpans(raw)
			if err != nil {
				return nil, nil, err
			}
			out := map[string]string{}
			locations := map[string]XMLSpan{}
			for _, s := range spans {
				if s.Name.Space != "http://schemas.microsoft.com/office/2006/xmlPackage" || s.Name.Local != "part" {
					continue
				}
				name := s.Attribute(s.Name.Space, "name")
				if name == "" || out[name] != "" {
					return nil, nil, fmt.Errorf("Flat OPC part name missing or duplicated: %q", name)
				}
				out[name] = Hash(raw[s.Start:s.End])
				locations[name] = s
			}
			return out, locations, nil
		}
		before, beforeLocations, err := parts(a)
		if err != nil {
			return nil, err
		}
		after, afterLocations, err := parts(b)
		if err != nil {
			return nil, err
		}
		names := map[string]bool{}
		for name := range before {
			names[name] = true
		}
		for name := range after {
			names[name] = true
		}
		ordered := []string{}
		for name := range names {
			ordered = append(ordered, name)
		}
		sort.Strings(ordered)
		rows := []map[string]any{}
		contextCount := 0
		for _, name := range ordered {
			row := map[string]any{"name": name, "reference_sha256": before[name], "candidate_sha256": after[name], "byte_identical": before[name] == after[name]}
			selected := len(policy.ContextParts) == 0
			for _, requested := range policy.ContextParts {
				if requested == name {
					selected = true
				}
			}
			if before[name] != after[name] && selected && contextCount < 8 {
				l, lok := beforeLocations[name]
				r, rok := afterLocations[name]
				if lok && rok {
					row["first_difference"] = partTokenDifference(a, b, left, right, l, r)
				} else {
					row["presence"] = map[string]bool{"reference": lok, "candidate": rok}
				}
				contextCount++
			} else if before[name] != after[name] {
				row["context_omitted"] = true
			}
			rows = append(rows, row)
		}
		result["parts"] = rows
	}
	for i := 0; i < len(left) || i < len(right); i++ {
		l, r := "<end>", "<end>"
		if i < len(left) {
			l = left[i].text
		}
		if i < len(right) {
			r = right[i].text
		}
		if l != r {
			result["equal"] = false
			offset := tokenMismatch(l, r)
			result["first_difference"] = map[string]any{"token": i, "canonical_token_byte": offset, "reference": tokenPreview(l, offset), "candidate": tokenPreview(r, offset), "reference_token_bytes": len(l), "candidate_token_bytes": len(r)}
			result["difference_context"] = map[string]any{
				"reference": comparisonContext(a, left, i),
				"candidate": comparisonContext(b, right, i),
			}
			break
		}
	}
	return result, nil
}

// Use the original token stream: extracted XML fragments may inherit namespaces.
func partTokenDifference(a, b []byte, left, right []comparisonToken, l, r XMLSpan) any {
	part := func(tokens []comparisonToken, span XMLSpan) []comparisonToken {
		start := sort.Search(len(tokens), func(i int) bool { return tokens[i].start >= int64(span.Start) })
		end := start + sort.Search(len(tokens)-start, func(i int) bool { return tokens[start+i].start >= int64(span.End) })
		return tokens[start:end]
	}
	lt, rt := part(left, l), part(right, r)
	context := func(raw []byte, tokens []comparisonToken, i int, span XMLSpan, offset int) any {
		if i >= len(tokens) {
			return map[string]any{"end_of_part": true, "byte_start": span.End}
		}
		t := tokens[i]
		lo, hi := max(int64(span.Start), t.start-128), min(int64(span.End), t.end+128)
		hi = min(hi, lo+1024)
		return map[string]any{"byte_start": t.start, "byte_end": t.end, "token": tokenPreview(t.text, offset), "token_bytes": len(t.text), "xml": string(raw[lo:hi]), "xml_start": lo}
	}
	for i := 0; i < len(lt) || i < len(rt); i++ {
		if i >= len(lt) || i >= len(rt) || lt[i].text != rt[i].text {
			offset := 0
			if i < len(lt) && i < len(rt) {
				offset = tokenMismatch(lt[i].text, rt[i].text)
			}
			return map[string]any{"canonical_token_byte": offset, "reference": context(a, lt, i, l, offset), "candidate": context(b, rt, i, r, offset)}
		}
	}
	return nil // Raw bytes differ, but the caller's comparison policy found no token difference.
}

// Only inspect structure on a mismatch; successful comparisons pay no span scan.
func comparisonContext(raw []byte, tokens []comparisonToken, index int) map[string]any {
	start, end := int64(len(raw)), int64(len(raw))
	if index < len(tokens) {
		start, end = tokens[index].start, tokens[index].end
	}
	lo, hi := max(0, start-256), min(int64(len(raw)), end+256)
	// Bound even an individual token containing a very large text node.
	hi = min(hi, lo+2048)
	context := map[string]any{"byte_start": start, "byte_end": end,
		"xml_start": lo, "xml_end": hi, "xml": string(raw[lo:hi]),
		"xml_escaped": fmt.Sprintf("%q", raw[lo:hi])}
	spans, err := XMLSpans(raw)
	if err != nil {
		context["context_error"] = err.Error()
		return context
	}
	path := []map[string]any{}
	for _, span := range spans {
		if int64(span.Start) <= start && int64(span.End) >= end {
			path = append(path, map[string]any{"namespace": span.Name.Space, "name": span.Name.Local,
				"byte_start": span.Start, "byte_end": span.End})
		}
	}
	context["ancestors"] = path
	context["formatting_scope"] = "serialized XML; effective inherited Word formatting is not resolved"
	return context
}

// Offsets here belong to the canonical diagnostic token, not the source XML.
func tokenMismatch(left, right string) int {
	i := 0
	for i < len(left) && i < len(right) && left[i] == right[i] {
		i++
	}
	return i
}

// Limit displayed tokens only; equality always compares their full contents.
func tokenPreview(text string, offset int) string {
	const limit = 1024
	if len(text) <= limit {
		return text
	}
	start := max(0, min(offset-256, len(text)-limit))
	for start > 0 && !utf8.RuneStart(text[start]) {
		start--
	}
	end := min(len(text), start+limit)
	for end < len(text) && end > start && !utf8.RuneStart(text[end]) {
		end--
	}
	preview := text[start:end]
	if start > 0 {
		preview = "[truncated]..." + preview
	}
	if end < len(text) {
		preview += "...[truncated]"
	}
	return preview
}

const wordXMLNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

var generatedTOCName = regexp.MustCompile(`^_Toc[0-9]+$`)
var generatedTOCReference = regexp.MustCompile(`(?i)^\s*(?:PAGEREF|REF)\s+"?(_Toc[0-9]+)"?(?:\s|$)`)
