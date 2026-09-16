package office

import (
	"bytes"
	"fmt"
	"sort"
	"strings"
	"unicode/utf8"
)

// XMLPatch addresses original UTF-8 bytes, never offsets shifted by another
// edit. Insertions at the same offset retain their request order and precede
// replacement at that offset. Overlapping replacements are rejected.
type XMLPatch struct {
	Offset int    `json:"offset"`
	Length int    `json:"length"`
	Text   string `json:"text"`
}

// PatchXML applies a bounded edit plan with one output allocation, then checks
// the whole result. No serialization touches XML outside the supplied ranges.
func PatchXML(source []byte, patches []XMLPatch) ([]byte, error) {
	out, err := spliceXML(source, patches)
	if err == nil {
		_, err = XMLSpans(out)
	}
	if err != nil {
		return nil, err
	}
	return out, nil
}

func spliceXML(source []byte, patches []XMLPatch) ([]byte, error) {
	if len(source) > Limit || len(patches) > 65536 {
		return nil, fmt.Errorf("XML edit budget exceeded")
	}
	ordered := append([]XMLPatch(nil), patches...)
	sort.SliceStable(ordered, func(i, j int) bool {
		if ordered[i].Offset == ordered[j].Offset {
			return ordered[i].Length < ordered[j].Length
		}
		return ordered[i].Offset < ordered[j].Offset
	})
	size, cursor := int64(len(source)), 0
	boundary := func(offset int) bool { return offset == len(source) || utf8.RuneStart(source[offset]) }
	for _, p := range ordered {
		if p.Offset < cursor || p.Offset > len(source) || p.Length < 0 || p.Length > len(source)-p.Offset {
			return nil, fmt.Errorf("XML patches overlap or exceed source bounds")
		}
		if !boundary(p.Offset) || !boundary(p.Offset+p.Length) || !utf8.ValidString(p.Text) {
			return nil, fmt.Errorf("XML patch must respect UTF-8 boundaries")
		}
		size += int64(len(p.Text)) - int64(p.Length)
		if size > Limit {
			return nil, fmt.Errorf("XML edit output budget exceeded")
		}
		cursor = p.Offset + p.Length
	}
	out := make([]byte, 0, int(size))
	cursor = 0
	for _, p := range ordered {
		out = append(out, source[cursor:p.Offset]...)
		out = append(out, p.Text...)
		cursor = p.Offset + p.Length
	}
	return append(out, source[cursor:]...), nil
}

type childChange struct {
	attrs map[string]string
	xml   string
}

// mergeChildren plans all direct-child edits against one immutable parse.
// Only explicitly edited opening tags are serialized; unknown children,
// comments, whitespace and extension attributes remain in the source bytes.
func mergeChildren(parent []byte, order string, changes map[string]childChange, scope ...map[string]string) ([]byte, error) {
	if len(changes) == 0 {
		return parent, nil
	}
	var inherited map[string]string
	if len(scope) > 0 {
		inherited = scope[0]
	}
	spans, err := xmlSpans(parent, inherited)
	if err != nil {
		return nil, err
	}
	ranks, existing := map[string]int{}, map[string]XMLSpan{}
	for i, name := range strings.Fields(order) {
		ranks[name] = i
	}
	for _, span := range spans {
		if span.Depth == 1 && (span.Name.Space == W || span.Name.Space == "w") {
			if _, found := existing[span.Name.Local]; !found {
				existing[span.Name.Local] = span
			}
		}
	}
	patches := []XMLPatch{}
	var tail strings.Builder
	for _, name := range strings.Fields(order) {
		change, ok := changes[name]
		if !ok {
			continue
		}
		xml := change.xml
		span, found := existing[name]
		if change.attrs != nil {
			if found {
				raw, err := patchAttrs(parent[span.Start:span.End], change.attrs, span.namespaces)
				if err != nil {
					return nil, err
				}
				xml = string(raw)
			} else {
				attrs := map[string]string{}
				for key, value := range change.attrs {
					if value != "\x00" {
						attrs[key] = value
					}
				}
				xml = attrsXML("w:"+name, attrs)
			}
		}
		if found {
			patches = append(patches, XMLPatch{span.Start, span.End - span.Start, xml})
			continue
		}
		at := spans[0].CloseStart
		for _, child := range spans {
			if child.Depth == 1 && (child.Name.Space == W || child.Name.Space == "w") {
				if rank, ok := ranks[child.Name.Local]; ok && rank > ranks[name] {
					at = child.Start
					break
				}
			}
		}
		if at == spans[0].CloseStart {
			tail.WriteString(xml)
		} else {
			patches = append(patches, XMLPatch{Offset: at, Text: xml})
		}
	}
	if tail.Len() > 0 {
		patches = append(patches, insertAtRoot(parent, spans[0], tail.String()))
	}
	return spliceXML(parent, patches)
}

func insertAtRoot(raw []byte, root XMLSpan, fragment string) XMLPatch {
	if bytes.HasSuffix(raw[root.Start:root.OpenEnd], []byte("/>")) {
		name := raw[root.Start+1 : root.OpenEnd-2]
		if at := bytes.IndexAny(name, " \t\r\n"); at >= 0 {
			name = name[:at]
		}
		return XMLPatch{root.OpenEnd - 2, 2, ">" + fragment + "</" + string(name) + ">"}
	}
	return XMLPatch{Offset: root.CloseStart, Text: fragment}
}
