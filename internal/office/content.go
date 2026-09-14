package office

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"math"
	"path"
	"sort"
	"strconv"
	"strings"
)

const styleOrder = "name aliases basedOn next link autoRedefine hidden uiPriority semiHidden unhideWhenUsed qFormat locked personal personalCompose personalReply rsid pPr rPr tblPr trPr tcPr tblStylePr"
const paraOrder = "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange"
const runOrder = "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange"

func KnownKeys(m map[string]any, allowed string) error {
	a := map[string]bool{}
	for _, s := range strings.Fields(allowed) {
		a[s] = true
	}
	for k := range m {
		if !a[k] {
			return fmt.Errorf("unknown property %q; edit the raw package XML for features outside this recipe", k)
		}
	}
	return nil
}
func numeric(v any) (float64, error) {
	switch n := v.(type) {
	case float64:
		if !math.IsNaN(n) && !math.IsInf(n, 0) {
			return n, nil
		}
	case int:
		return float64(n), nil
	case json.Number:
		return n.Float64()
	}
	return 0, fmt.Errorf("finite numeric measurement required")
}
func scaled(v any, mult float64) (string, error) {
	n, e := numeric(v)
	if e != nil || math.Abs(n) > 1000000 {
		return "", fmt.Errorf("measurement outside supported range")
	}
	return strconv.FormatInt(int64(math.Round(n*mult)), 10), nil
}
func boolVal(v any) (string, error) {
	b, ok := v.(bool)
	if !ok {
		return "", fmt.Errorf("boolean required")
	}
	if b {
		return "1", nil
	}
	return "0", nil
}
func qname(n xml.Name) string {
	if n.Space != "" {
		return n.Space + ":" + n.Local
	}
	return n.Local
}
func attrsXML(tag string, attrs map[string]string) string {
	keys := make([]string, 0, len(attrs))
	for k := range attrs {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	var b strings.Builder
	b.WriteString("<" + tag)
	for _, k := range keys {
		b.WriteString(" " + k + `="` + Esc(attrs[k]) + `"`)
	}
	b.WriteString("/>")
	return b.String()
}
func patchAttrs(b []byte, patch map[string]string) ([]byte, error) {
	d := xml.NewDecoder(bytes.NewReader(b))
	t, e := d.RawToken()
	if e != nil {
		return nil, e
	}
	s, ok := t.(xml.StartElement)
	if !ok {
		return nil, fmt.Errorf("XML element required")
	}
	attrs := map[string]string{}
	for _, a := range s.Attr {
		attrs[qname(a.Name)] = a.Value
	}
	for k, v := range patch {
		if v == "\x00" {
			delete(attrs, k)
		} else {
			attrs[k] = v
		}
	}
	end := int(d.InputOffset())
	open := attrsXML(qname(s.Name), attrs)
	if end >= 2 && bytes.Equal(b[end-2:end], []byte("/>")) {
		return append([]byte(open), b[end:]...), nil
	}
	return append([]byte(strings.TrimSuffix(open, "/>")+">"), b[end:]...), nil
}
func mergeChild(parent []byte, local, fragment, order string, mergeAttributes bool) ([]byte, error) {
	spans, e := XMLSpans(parent)
	if e != nil {
		return nil, e
	}
	for _, s := range spans {
		if s.Depth == 1 && s.Name.Local == local && (s.Name.Space == W || s.Name.Space == "w") {
			node := []byte(fragment)
			if mergeAttributes {
				d := xml.NewDecoder(strings.NewReader(fragment))
				t, e := d.RawToken()
				if e != nil {
					return nil, e
				}
				attrs := map[string]string{}
				for _, a := range t.(xml.StartElement).Attr {
					attrs[qname(a.Name)] = a.Value
				}
				if local == "rFonts" {
					if _, set := attrs["w:ascii"]; set {
						for _, k := range []string{"asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "csTheme"} {
							attrs["w:"+k] = "\x00"
						}
					}
				}
				if local == "ind" {
					if _, ok := attrs["w:firstLine"]; ok {
						attrs["w:hanging"] = "\x00"
						attrs["w:hangingChars"] = "\x00"
					}
					if _, ok := attrs["w:hanging"]; ok && attrs["w:hanging"] != "\x00" {
						attrs["w:firstLine"] = "\x00"
						attrs["w:firstLineChars"] = "\x00"
					}
				}
				node, e = patchAttrs(parent[s.Start:s.End], attrs)
				if e != nil {
					return nil, e
				}
			}
			return bytes.Join([][]byte{parent[:s.Start], node, parent[s.End:]}, nil), nil
		}
	}
	ranks := map[string]int{}
	for i, s := range strings.Fields(order) {
		ranks[s] = i
	}
	rank, ordered := ranks[local]
	if ordered {
		for _, s := range spans {
			if s.Depth == 1 {
				if n, ok := ranks[s.Name.Local]; ok && n > rank {
					return bytes.Join([][]byte{parent[:s.Start], []byte(fragment), parent[s.Start:]}, nil), nil
				}
			}
		}
	}
	return InsertXML(parent, fragment)
}
func props(raw []byte, kind string, spec map[string]any) ([]byte, error) {
	tag, order := "rPr", runOrder
	if kind == "paragraph" {
		tag, order = "pPr", paraOrder
	}
	if len(raw) == 0 {
		raw = []byte("<w:" + tag + " xmlns:w=\"" + W + "\"/>")
	}
	patches := map[string]map[string]string{}
	full := map[string]string{}
	attr := func(t, a, v string) {
		if patches[t] == nil {
			patches[t] = map[string]string{}
		}
		patches[t]["w:"+a] = v
	}
	val := func(t string, v any) { attr(t, "val", fmt.Sprint(v)) }
	if kind == "run" {
		if e := KnownKeys(spec, "style font size_pt bold italic small_caps all_caps strike color underline superscript subscript highlight language hidden"); e != nil {
			return nil, e
		}
		if v, ok := spec["style"]; ok {
			val("rStyle", v)
		}
		if v, ok := spec["font"]; ok {
			for _, a := range []string{"ascii", "hAnsi", "eastAsia", "cs"} {
				attr("rFonts", a, fmt.Sprint(v))
			}
		}
		if v, ok := spec["size_pt"]; ok {
			n, e := scaled(v, 2)
			if e != nil {
				return nil, e
			}
			val("sz", n)
			val("szCs", n)
		}
		for k, t := range map[string]string{"bold": "b", "italic": "i", "small_caps": "smallCaps", "all_caps": "caps", "strike": "strike", "hidden": "vanish"} {
			if v, ok := spec[k]; ok {
				n, e := boolVal(v)
				if e != nil {
					return nil, e
				}
				val(t, n)
			}
		}
		for k, t := range map[string]string{"color": "color", "highlight": "highlight"} {
			if v, ok := spec[k]; ok {
				val(t, v)
			}
		}
		if v, ok := spec["language"]; ok {
			attr("lang", "val", fmt.Sprint(v))
		}
		if v, ok := spec["underline"]; ok {
			if b, ok := v.(bool); ok {
				if b {
					val("u", "single")
				} else {
					val("u", "none")
				}
			} else {
				val("u", v)
			}
		}
		if v, ok := spec["superscript"]; ok {
			b, ok := v.(bool)
			if !ok {
				return nil, fmt.Errorf("superscript boolean required")
			}
			if b {
				val("vertAlign", "superscript")
			} else {
				val("vertAlign", "baseline")
			}
		}
		if v, ok := spec["subscript"]; ok {
			b, ok := v.(bool)
			if !ok {
				return nil, fmt.Errorf("subscript boolean required")
			}
			if b {
				val("vertAlign", "subscript")
			} else {
				val("vertAlign", "baseline")
			}
		}
	} else {
		if e := KnownKeys(spec, "style alignment before_pt after_pt line_pt line_multiple left_pt right_pt first_line_pt hanging_pt keep_next keep_lines page_break_before widow_control outline_level list_id list_level tabs"); e != nil {
			return nil, e
		}
		if _, ok := spec["line_pt"]; ok {
			if _, ok := spec["line_multiple"]; ok {
				return nil, fmt.Errorf("choose exact line spacing or a multiple")
			}
		}
		if _, ok := spec["first_line_pt"]; ok {
			if _, ok := spec["hanging_pt"]; ok {
				return nil, fmt.Errorf("first line and hanging indents are mutually exclusive")
			}
		}
		if v, ok := spec["style"]; ok {
			val("pStyle", v)
		}
		if v, ok := spec["alignment"]; ok {
			if !strings.Contains("|left|right|center|both|distribute|start|end|", "|"+fmt.Sprint(v)+"|") {
				return nil, fmt.Errorf("invalid paragraph alignment")
			}
			val("jc", v)
		}
		for k, t := range map[string]string{"keep_next": "keepNext", "keep_lines": "keepLines", "page_break_before": "pageBreakBefore", "widow_control": "widowControl"} {
			if v, ok := spec[k]; ok {
				n, e := boolVal(v)
				if e != nil {
					return nil, e
				}
				val(t, n)
			}
		}
		for k, a := range map[string]string{"before_pt": "before", "after_pt": "after", "line_pt": "line"} {
			if v, ok := spec[k]; ok {
				n, e := scaled(v, 20)
				if e != nil {
					return nil, e
				}
				attr("spacing", a, n)
				if k == "line_pt" {
					attr("spacing", "lineRule", "exact")
				}
			}
		}
		if v, ok := spec["line_multiple"]; ok {
			n, e := scaled(v, 240)
			if e != nil {
				return nil, e
			}
			attr("spacing", "line", n)
			attr("spacing", "lineRule", "auto")
		}
		for k, a := range map[string]string{"left_pt": "left", "right_pt": "right", "first_line_pt": "firstLine", "hanging_pt": "hanging"} {
			if v, ok := spec[k]; ok {
				n, e := scaled(v, 20)
				if e != nil {
					return nil, e
				}
				attr("ind", a, n)
			}
		}
		if v, ok := spec["outline_level"]; ok {
			n, e := numeric(v)
			if e != nil || n < 0 || n > 9 {
				return nil, fmt.Errorf("invalid outline level")
			}
			val("outlineLvl", int(n))
		}
		if v, ok := spec["list_id"]; ok {
			n, e := scaled(v, 1)
			if e != nil {
				return nil, e
			}
			level := "0"
			if x, ok := spec["list_level"]; ok {
				level, e = scaled(x, 1)
				if e != nil {
					return nil, e
				}
			}
			full["numPr"] = `<w:numPr><w:ilvl w:val="` + level + `"/><w:numId w:val="` + n + `"/></w:numPr>`
		}
		if v, ok := spec["tabs"]; ok {
			tabs, ok := v.([]any)
			if !ok {
				return nil, fmt.Errorf("tabs must be array")
			}
			var b strings.Builder
			b.WriteString("<w:tabs>")
			for _, t := range tabs {
				m, ok := t.(map[string]any)
				if !ok {
					return nil, fmt.Errorf("invalid tab")
				}
				if e := KnownKeys(m, "position_pt alignment leader"); e != nil {
					return nil, e
				}
				pos, e := scaled(m["position_pt"], 20)
				if e != nil {
					return nil, e
				}
				a := map[string]string{"w:pos": pos, "w:val": "left"}
				if x, ok := m["alignment"]; ok {
					a["w:val"] = fmt.Sprint(x)
				}
				if x, ok := m["leader"]; ok {
					a["w:leader"] = fmt.Sprint(x)
				}
				b.WriteString(attrsXML("w:tab", a))
			}
			b.WriteString("</w:tabs>")
			full["tabs"] = b.String()
		}
	}
	for _, t := range strings.Fields(order) {
		var e error
		if a, ok := patches[t]; ok {
			raw, e = mergeChild(raw, t, attrsXML("w:"+t, a), order, true)
		}
		if f, ok := full[t]; ok {
			raw, e = mergeChild(raw, t, f, order, false)
		}
		if e != nil {
			return nil, e
		}
	}
	return raw, nil
}

type StyleSpec struct {
	ID        string         `json:"id"`
	Name      string         `json:"name,omitempty"`
	Type      string         `json:"type,omitempty"`
	BasedOn   string         `json:"based_on,omitempty"`
	Next      string         `json:"next,omitempty"`
	Linked    string         `json:"linked,omitempty"`
	Quick     *bool          `json:"quick,omitempty"`
	Run       map[string]any `json:"run,omitempty"`
	Paragraph map[string]any `json:"paragraph,omitempty"`
	XML       string         `json:"xml,omitempty"`
}
type NumberLevel struct {
	Level     int            `json:"level"`
	Start     int            `json:"start,omitempty"`
	Format    string         `json:"format"`
	Text      string         `json:"text"`
	Suffix    string         `json:"suffix,omitempty"`
	Paragraph map[string]any `json:"paragraph,omitempty"`
	Run       map[string]any `json:"run,omitempty"`
}
type NumberingSpec struct {
	ID     int           `json:"id"`
	Levels []NumberLevel `json:"levels"`
}
type StyleRecipe struct {
	Styles    []StyleSpec     `json:"styles,omitempty"`
	Numbering []NumberingSpec `json:"numbering,omitempty"`
}

func ApplyStyles(p *Package, r StyleRecipe) error {
	work := p.Clone()
	if err := applyStyles(work, r); err != nil {
		return err
	}
	*p = *work
	return nil
}

func applyStyles(p *Package, r StyleRecipe) error {
	if len(r.Styles) > 0 {
		part := "word/styles.xml"
		b := p.Files[part]
		if b == nil {
			b = []byte(`<w:styles xmlns:w="` + W + `"/>`)
		}
		var e error
		b, e = EnsureNamespace(b, "w", W)
		if e != nil {
			return e
		}
		for _, s := range r.Styles {
			if s.ID == "" {
				return fmt.Errorf("style id required")
			}
			spans, e := XMLSpans(b)
			if e != nil {
				return e
			}
			var node []byte
			for _, n := range spans {
				if n.Depth == 1 && n.Name.Local == "style" && n.Attribute(W, "styleId") == s.ID {
					node = b[n.Start:n.End]
					if s.Type != "" && n.Attribute(W, "type") != s.Type {
						return fmt.Errorf("style type conversion requires a new ID")
					}
					break
				}
			}
			existing := len(node) > 0
			if len(node) == 0 {
				kind := s.Type
				if kind == "" {
					kind = "paragraph"
				}
				node = []byte(`<w:style xmlns:w="` + W + `" w:type="` + Esc(kind) + `" w:styleId="` + Esc(s.ID) + `" w:customStyle="1"/>`)
			}
			if s.XML != "" {
				node = []byte(s.XML)
				nodes, e := XMLSpans(node)
				if e != nil {
					return e
				}
				if nodes[0].Name.Local != "style" || nodes[0].Attribute("*", "styleId") != s.ID {
					return fmt.Errorf("raw style identity mismatch")
				}
			} else {
				name := s.Name
				if name == "" && !existing {
					name = s.ID
				}
				for _, kv := range [][2]string{{"name", name}, {"basedOn", s.BasedOn}, {"next", s.Next}, {"link", s.Linked}} {
					if kv[1] != "" {
						node, e = mergeChild(node, kv[0], attrsXML("w:"+kv[0], map[string]string{"w:val": kv[1]}), styleOrder, true)
						if e != nil {
							return e
						}
					}
				}
				if s.Quick != nil {
					v := "0"
					if *s.Quick {
						v = "1"
					}
					node, e = mergeChild(node, "qFormat", `<w:qFormat w:val="`+v+`"/>`, styleOrder, true)
					if e != nil {
						return e
					}
				}
				for _, it := range []struct {
					tag, kind string
					spec      map[string]any
				}{{"pPr", "paragraph", s.Paragraph}, {"rPr", "run", s.Run}} {
					if it.spec == nil {
						continue
					}
					var raw []byte
					nodes, e := XMLSpans(node)
					if e != nil {
						return e
					}
					for _, n := range nodes {
						if n.Depth == 1 && n.Name.Local == it.tag {
							raw = node[n.Start:n.End]
							break
						}
					}
					raw, e = props(raw, it.kind, it.spec)
					if e != nil {
						return e
					}
					node, e = mergeChild(node, it.tag, string(raw), styleOrder, false)
					if e != nil {
						return e
					}
				}
			}
			b, e = UpsertXML(b, W, "style", W, "styleId", s.ID, string(node))
			if e != nil {
				return e
			}
		}
		p.Files[part] = b
		if e = p.ContentType(part, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"); e != nil {
			return e
		}
		if e = p.Relationship("word/document.xml", "rIdStyles", R+"/styles", "styles.xml", ""); e != nil {
			return e
		}
	}
	if len(r.Numbering) > 0 {
		b := p.Files["word/numbering.xml"]
		if b == nil {
			b = []byte(`<w:numbering xmlns:w="` + W + `"/>`)
		}
		var e error
		b, e = EnsureNamespace(b, "w", W)
		if e != nil {
			return e
		}
		spans, e := XMLSpans(b)
		if e != nil {
			return e
		}
		next := 0
		for _, n := range spans {
			if n.Name.Local == "abstractNum" {
				id, _ := strconv.Atoi(n.Attribute(W, "abstractNumId"))
				next = max(next, id+1)
			}
		}
		seen := map[int]bool{}
		for _, n := range r.Numbering {
			if n.ID < 1 || len(n.Levels) == 0 || len(n.Levels) > 9 || seen[n.ID] {
				return fmt.Errorf("invalid numbering recipe")
			}
			seen[n.ID] = true
			aid := next
			next++
			var x strings.Builder
			x.WriteString(`<w:abstractNum w:abstractNumId="` + strconv.Itoa(aid) + `"><w:multiLevelType w:val="multilevel"/>`)
			levels := map[int]bool{}
			for _, l := range n.Levels {
				if l.Level < 0 || l.Level > 8 || levels[l.Level] {
					return fmt.Errorf("invalid numbering level")
				}
				levels[l.Level] = true
				start := l.Start
				if start == 0 {
					start = 1
				}
				x.WriteString(fmt.Sprintf(`<w:lvl w:ilvl="%d"><w:start w:val="%d"/><w:numFmt w:val="%s"/><w:lvlText w:val="%s"/>`, l.Level, start, Esc(l.Format), Esc(l.Text)))
				if l.Suffix != "" {
					x.WriteString(`<w:suff w:val="` + Esc(l.Suffix) + `"/>`)
				}
				pp, e := props(nil, "paragraph", l.Paragraph)
				if e != nil {
					return e
				}
				rp, e := props(nil, "run", l.Run)
				if e != nil {
					return e
				}
				x.Write(pp)
				x.Write(rp)
				x.WriteString(`</w:lvl>`)
			}
			x.WriteString(`</w:abstractNum>`)
			// Abstract definitions precede concrete numbering instances.
			b, e = mergeChildAtEndOfKind(b, "abstractNum", x.String())
			if e != nil {
				return e
			}
			b, e = UpsertXML(b, W, "num", W, "numId", strconv.Itoa(n.ID), fmt.Sprintf(`<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>`, n.ID, aid))
			if e != nil {
				return e
			}
		}
		p.Files["word/numbering.xml"] = b
		if e := p.ContentType("word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"); e != nil {
			return e
		}
		if e := p.Relationship("word/document.xml", "rIdNumbering", R+"/numbering", "numbering.xml", ""); e != nil {
			return e
		}
	}
	return nil
}
func mergeChildAtEndOfKind(b []byte, kind, fragment string) ([]byte, error) {
	spans, e := XMLSpans(b)
	if e != nil {
		return nil, e
	}
	for _, s := range spans {
		if s.Depth == 1 && s.Name.Local == "num" {
			return bytes.Join([][]byte{b[:s.Start], []byte(fragment), b[s.Start:]}, nil), nil
		}
	}
	return InsertXML(b, fragment)
}

// The recipe is an authoring convenience, not a substitute Word object model.
// XML remains available for every structure outside these conveniences.
type Inline struct {
	Text     string         `json:"text,omitempty"`
	Style    string         `json:"style,omitempty"`
	Run      map[string]any `json:"run,omitempty"`
	Field    string         `json:"field,omitempty"`
	URL      string         `json:"url,omitempty"`
	Bookmark string         `json:"bookmark,omitempty"`
	Tab      bool           `json:"tab,omitempty"`
	Break    string         `json:"break,omitempty"`
	Footnote []Block        `json:"footnote,omitempty"`
	Image    string         `json:"image,omitempty"`
	WidthPT  float64        `json:"width_pt,omitempty"`
	HeightPT float64        `json:"height_pt,omitempty"`
	Alt      string         `json:"alt,omitempty"`
	XML      string         `json:"xml,omitempty"`
}
type Cell struct {
	Blocks        []Block `json:"blocks,omitempty"`
	Text          string  `json:"text,omitempty"`
	WidthPT       float64 `json:"width_pt,omitempty"`
	Span          int     `json:"span,omitempty"`
	Shade         string  `json:"shade,omitempty"`
	VerticalMerge string  `json:"vertical_merge,omitempty"`
}
type Block struct {
	Type       string         `json:"type,omitempty"`
	Text       string         `json:"text,omitempty"`
	Style      string         `json:"style,omitempty"`
	Paragraph  map[string]any `json:"paragraph,omitempty"`
	Run        map[string]any `json:"run,omitempty"`
	Inlines    []Inline       `json:"inlines,omitempty"`
	Blocks     []Block        `json:"blocks,omitempty"`
	Rows       [][]Cell       `json:"rows,omitempty"`
	ColumnsPT  []float64      `json:"columns_pt,omitempty"`
	HeaderRows int            `json:"header_rows,omitempty"`
	Tag        string         `json:"tag,omitempty"`
	Title      string         `json:"title,omitempty"`
	Bookmark   string         `json:"bookmark,omitempty"`
	XML        string         `json:"xml,omitempty"`
}
type PageSpec struct {
	WidthPT   float64            `json:"width_pt,omitempty"`
	HeightPT  float64            `json:"height_pt,omitempty"`
	MarginsPT map[string]float64 `json:"margins_pt,omitempty"`
	Landscape bool               `json:"landscape,omitempty"`
	Header    []Block            `json:"header,omitempty"`
	Footer    []Block            `json:"footer,omitempty"`
	PageStart int                `json:"page_start,omitempty"`
	Columns   int                `json:"columns,omitempty"`
}
type ContentRecipe struct {
	Blocks []Block   `json:"blocks"`
	Page   *PageSpec `json:"page,omitempty"`
}
type BuildingBlock struct {
	Name        string  `json:"name"`
	Gallery     string  `json:"gallery,omitempty"`
	Category    string  `json:"category,omitempty"`
	Description string  `json:"description,omitempty"`
	Blocks      []Block `json:"blocks"`
}
type composer struct {
	p          *Package
	source     string
	asset      func(string) ([]byte, error)
	next       int
	nextNoteID int
	noteIDs    bool
	notes      []string
	depth      int
}

func (c *composer) textRun(text string, style string, spec map[string]any) (string, error) {
	if spec == nil {
		spec = map[string]any{}
	}
	m := map[string]any{}
	for k, v := range spec {
		m[k] = v
	}
	if style != "" {
		m["style"] = style
	}
	rp, e := props(nil, "run", m)
	if e != nil {
		return "", e
	}
	var b strings.Builder
	b.WriteString("<w:r>")
	if len(m) > 0 {
		b.Write(rp)
	}
	parts := strings.Split(strings.ReplaceAll(text, "\r\n", "\n"), "\n")
	for i, s := range parts {
		if i > 0 {
			b.WriteString("<w:br/>")
		}
		for j, t := range strings.Split(s, "\t") {
			if j > 0 {
				b.WriteString("<w:tab/>")
			}
			b.WriteString(`<w:t xml:space="preserve">` + Esc(t) + `</w:t>`)
		}
	}
	b.WriteString("</w:r>")
	return b.String(), nil
}
func (c *composer) inline(in Inline) (string, error) {
	if in.XML != "" {
		if _, e := XMLSpans([]byte(`<root xmlns:w="` + W + `" xmlns:r="` + R + `">` + in.XML + `</root>`)); e != nil {
			return "", e
		}
		return in.XML, nil
	}
	if in.Tab {
		return "<w:r><w:tab/></w:r>", nil
	}
	if in.Break != "" {
		if !strings.Contains("|textWrapping|page|column|", "|"+in.Break+"|") {
			return "", fmt.Errorf("invalid break type")
		}
		return `<w:r><w:br w:type="` + Esc(in.Break) + `"/></w:r>`, nil
	}
	if len(in.Footnote) > 0 {
		if c.source != "word/document.xml" {
			return "", fmt.Errorf("footnotes in saved parts/headers need explicit relationship-aware XML")
		}
		if !c.noteIDs {
			next, err := nextFootnoteID(c.p)
			if err != nil {
				return "", err
			}
			c.nextNoteID, c.noteIDs = next, true
		}
		id := c.nextNoteID
		c.nextNoteID++
		prior := c.source
		c.source = "word/footnotes.xml"
		body, e := c.blocks(in.Footnote)
		c.source = prior
		if e != nil {
			return "", e
		}
		// Add the real note reference after pPr in the first paragraph.
		spans, e := XMLSpans([]byte(`<root xmlns:w="` + W + `" xmlns:r="` + R + `">` + body + `</root>`))
		if e != nil {
			return "", e
		}
		wrapper := `<root xmlns:w="` + W + `" xmlns:r="` + R + `">`
		pos := -1
		for _, sp := range spans {
			if sp.Depth == 1 && sp.Name.Local == "p" {
				pos = sp.OpenEnd - len(wrapper)
				// Raw XML may contain a self-closing first paragraph.
				if strings.HasSuffix(string([]byte(body)[sp.Start-len(wrapper):pos]), "/>") {
					start := sp.Start - len(wrapper)
					body = body[:pos-2] + "></w:p>" + body[pos:]
					pos = start + (sp.OpenEnd - sp.Start) - 1
				}
				for _, ch := range spans {
					if ch.Depth == 2 && ch.Start >= sp.Start && ch.End < sp.End && ch.Name.Local == "pPr" {
						pos = ch.End - len(wrapper)
					}
				}
				break
			}
		}
		mark := `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:tab/></w:r>`
		if pos < 0 {
			body = "<w:p>" + mark + "</w:p>" + body
		} else {
			body = body[:pos] + mark + body[pos:]
		}
		c.notes = append(c.notes, fmt.Sprintf(`<w:footnote w:id="%d">%s</w:footnote>`, id, body))
		return fmt.Sprintf(`<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="%d"/></w:r>`, id), nil
	}
	if in.Image != "" {
		return c.image(in)
	}
	run, e := c.textRun(in.Text, in.Style, in.Run)
	if e != nil {
		return "", e
	}
	if in.Field != "" {
		return `<w:fldSimple w:instr="` + Esc(in.Field) + `" w:dirty="1">` + run + `</w:fldSimple>`, nil
	}
	if in.URL != "" {
		c.next++
		id := fmt.Sprintf("wwLink%d", c.next)
		if e = c.p.Relationship(c.source, id, R+"/hyperlink", in.URL, "External"); e != nil {
			return "", e
		}
		return `<w:hyperlink r:id="` + id + `">` + run + `</w:hyperlink>`, nil
	}
	if in.Bookmark != "" {
		return `<w:hyperlink w:anchor="` + Esc(in.Bookmark) + `">` + run + `</w:hyperlink>`, nil
	}
	return run, nil
}
func (c *composer) image(in Inline) (string, error) {
	if c.asset == nil {
		return "", fmt.Errorf("asset reader unavailable")
	}
	data, e := c.asset(in.Image)
	if e != nil {
		return "", e
	}
	ext := strings.ToLower(path.Ext(in.Image))
	ct := map[string]string{".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif", ".bmp": "image/bmp", ".svg": "image/svg+xml", ".emf": "image/x-emf", ".wmf": "image/x-wmf"}[ext]
	if ct == "" {
		return "", fmt.Errorf("unsupported image content type")
	}
	if in.WidthPT <= 0 || in.HeightPT <= 0 || in.WidthPT > 10000 || in.HeightPT > 10000 {
		return "", fmt.Errorf("explicit positive image dimensions required")
	}
	c.next++
	id := fmt.Sprintf("wwImage%d", c.next)
	file := "word/media/ww-" + Hash(data)[:20] + ext
	c.p.Files[file] = data
	if e = c.p.ContentType(file, ct); e != nil {
		return "", e
	}
	rel, e := pathRelative(path.Dir(c.source), file)
	if e != nil {
		return "", e
	}
	if e = c.p.Relationship(c.source, id, R+"/image", rel, ""); e != nil {
		return "", e
	}
	cx, cy := int64(math.Round(in.WidthPT*12700)), int64(math.Round(in.HeightPT*12700))
	return fmt.Sprintf(`<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="%s" descr="%s"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="%s" descr="%s"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`, cx, cy, c.next, Esc(path.Base(in.Image)), Esc(in.Alt), c.next, Esc(path.Base(in.Image)), Esc(in.Alt), id, cx, cy), nil
}
func pathRelative(base, target string) (string, error) {
	a := strings.Split(path.Clean(base), "/")
	b := strings.Split(path.Clean(target), "/")
	i := 0
	for i < len(a) && i < len(b) && a[i] == b[i] {
		i++
	}
	out := append([]string{}, b[i:]...)
	for j := i; j < len(a); j++ {
		out = append([]string{".."}, out...)
	}
	return strings.Join(out, "/"), nil
}
func (c *composer) blocks(blocks []Block) (string, error) {
	c.depth++
	defer func() { c.depth-- }()
	if c.depth > 64 {
		return "", fmt.Errorf("content nesting budget exceeded")
	}
	var out strings.Builder
	for _, b := range blocks {
		switch b.Type {
		case "", "paragraph":
			p := map[string]any{}
			for k, v := range b.Paragraph {
				p[k] = v
			}
			if b.Style != "" {
				p["style"] = b.Style
			}
			pp, e := props(nil, "paragraph", p)
			if e != nil {
				return "", e
			}
			out.WriteString("<w:p>")
			if len(p) > 0 {
				out.Write(pp)
			}
			mark := 0
			if b.Bookmark != "" {
				if !ValidIdentifier(b.Bookmark) {
					return "", fmt.Errorf("bookmark must be a stable identifier")
				}
				c.next++
				mark = c.next
				out.WriteString(fmt.Sprintf(`<w:bookmarkStart w:id="%d" w:name="%s"/>`, mark, Esc(b.Bookmark)))
			}
			if b.Text != "" {
				x, e := c.textRun(b.Text, "", b.Run)
				if e != nil {
					return "", e
				}
				out.WriteString(x)
			}
			for _, in := range b.Inlines {
				x, e := c.inline(in)
				if e != nil {
					return "", e
				}
				out.WriteString(x)
			}
			if mark > 0 {
				out.WriteString(fmt.Sprintf(`<w:bookmarkEnd w:id="%d"/>`, mark))
			}
			out.WriteString("</w:p>")
		case "content_control":
			c.next++
			body, e := c.blocks(b.Blocks)
			if e != nil {
				return "", e
			}
			out.WriteString(fmt.Sprintf(`<w:sdt><w:sdtPr><w:alias w:val="%s"/><w:tag w:val="%s"/><w:id w:val="%d"/></w:sdtPr><w:sdtContent>%s</w:sdtContent></w:sdt>`, Esc(b.Title), Esc(b.Tag), c.next, body))
		case "table":
			if len(b.Rows) == 0 {
				return "", fmt.Errorf("table needs rows")
			}
			out.WriteString(`<w:tbl><w:tblPr>`)
			if b.Style != "" {
				out.WriteString(`<w:tblStyle w:val="` + Esc(b.Style) + `"/>`)
			}
			out.WriteString(`<w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>`)
			widths := b.ColumnsPT
			if len(widths) == 0 {
				widths = make([]float64, len(b.Rows[0]))
				for i := range widths {
					widths[i] = 72
				}
			}
			for _, w := range widths {
				if w <= 0 {
					return "", fmt.Errorf("positive column width required")
				}
				out.WriteString(fmt.Sprintf(`<w:gridCol w:w="%.0f"/>`, w*20))
			}
			out.WriteString(`</w:tblGrid>`)
			for i, row := range b.Rows {
				out.WriteString(`<w:tr>`)
				if i < b.HeaderRows {
					out.WriteString(`<w:trPr><w:tblHeader/></w:trPr>`)
				}
				occupied := 0
				for _, cell := range row {
					span := max(1, cell.Span)
					occupied += span
					out.WriteString(`<w:tc><w:tcPr>`)
					if cell.WidthPT > 0 {
						out.WriteString(fmt.Sprintf(`<w:tcW w:w="%.0f" w:type="dxa"/>`, cell.WidthPT*20))
					}
					if span > 1 {
						out.WriteString(fmt.Sprintf(`<w:gridSpan w:val="%d"/>`, span))
					}
					if cell.VerticalMerge != "" {
						if cell.VerticalMerge != "restart" && cell.VerticalMerge != "continue" {
							return "", fmt.Errorf("invalid vertical merge")
						}
						out.WriteString(`<w:vMerge w:val="` + cell.VerticalMerge + `"/>`)
					}
					if cell.Shade != "" {
						out.WriteString(`<w:shd w:val="clear" w:fill="` + Esc(cell.Shade) + `"/>`)
					}
					out.WriteString(`</w:tcPr>`)
					blocks := cell.Blocks
					if len(blocks) == 0 {
						blocks = []Block{{Text: cell.Text}}
					}
					x, e := c.blocks(blocks)
					if e != nil {
						return "", e
					}
					out.WriteString(x)
					last := blocks[len(blocks)-1].Type
					if last != "paragraph" && last != "" {
						out.WriteString(`<w:p/>`)
					}
					out.WriteString(`</w:tc>`)
				}
				if occupied != len(widths) {
					return "", fmt.Errorf("table row does not span the grid")
				}
				out.WriteString(`</w:tr>`)
			}
			out.WriteString(`</w:tbl>`)
		case "xml":
			if _, e := XMLSpans([]byte(`<root xmlns:w="` + W + `" xmlns:r="` + R + `">` + b.XML + `</root>`)); e != nil {
				return "", e
			}
			out.WriteString(b.XML)
		default:
			return "", fmt.Errorf("unknown content block type %q", b.Type)
		}
		if out.Len() > Limit {
			return "", fmt.Errorf("content budget exceeded")
		}
	}
	return out.String(), nil
}
func Compose(p *Package, r ContentRecipe, asset func(string) ([]byte, error)) error {
	work := p.Clone()
	if err := compose(work, r, asset); err != nil {
		return err
	}
	*p = *work
	return nil
}

func compose(p *Package, r ContentRecipe, asset func(string) ([]byte, error)) error {
	c := &composer{p: p, source: "word/document.xml", asset: asset, next: nextDocumentID(p)}
	body, e := c.blocks(r.Blocks)
	if e != nil {
		return e
	}
	section := `<w:sectPr/>`
	if r.Page != nil {
		pg := r.Page
		var b strings.Builder
		b.WriteString(`<w:sectPr>`)
		for _, item := range []struct {
			typ    string
			blocks []Block
		}{{"header", pg.Header}, {"footer", pg.Footer}} {
			if item.blocks == nil {
				continue
			}
			old := c.source
			part := "word/ww-" + item.typ + ".xml"
			c.source = part
			x, e := c.blocks(item.blocks)
			c.source = old
			if e != nil {
				return e
			}
			tag := "hdr"
			if item.typ == "footer" {
				tag = "ftr"
			}
			p.Files[part] = []byte(`<w:` + tag + ` xmlns:w="` + W + `" xmlns:r="` + R + `">` + x + `</w:` + tag + `>`)
			if e := p.ContentType(part, "application/vnd.openxmlformats-officedocument.wordprocessingml."+item.typ+"+xml"); e != nil {
				return e
			}
			id := "ww" + item.typ
			if e := p.Relationship(c.source, id, R+"/"+item.typ, path.Base(part), ""); e != nil {
				return e
			}
			b.WriteString(`<w:` + item.typ + `Reference w:type="default" r:id="` + id + `"/>`)
		}
		w, h := pg.WidthPT, pg.HeightPT
		if w == 0 {
			w = 612
		}
		if h == 0 {
			h = 792
		}
		if w <= 0 || h <= 0 || w > 1584 || h > 1584 {
			return fmt.Errorf("page dimensions outside Word's supported 22-inch range")
		}
		orient := ""
		if pg.Landscape {
			orient = ` w:orient="landscape"`
		}
		b.WriteString(fmt.Sprintf(`<w:pgSz w:w="%.0f" w:h="%.0f"%s/>`, w*20, h*20, orient))
		m := map[string]float64{"top": 72, "bottom": 72, "left": 72, "right": 72, "header": 36, "footer": 36, "gutter": 0}
		for k, v := range pg.MarginsPT {
			if _, ok := m[k]; !ok {
				return fmt.Errorf("unknown page margin %s", k)
			}
			if v < 0 || v > 1584 {
				return fmt.Errorf("invalid margin")
			}
			m[k] = v
		}
		a := map[string]string{}
		for k, v := range m {
			a["w:"+k] = fmt.Sprintf("%.0f", v*20)
		}
		b.WriteString(attrsXML("w:pgMar", a))
		if pg.PageStart > 0 {
			b.WriteString(fmt.Sprintf(`<w:pgNumType w:start="%d"/>`, pg.PageStart))
		}
		if pg.Columns > 0 {
			b.WriteString(fmt.Sprintf(`<w:cols w:num="%d"/>`, pg.Columns))
		}
		b.WriteString(`</w:sectPr>`)
		section = b.String()
	}
	p.Files["word/document.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="` + W + `" xmlns:r="` + R + `"><w:body>` + body + section + `</w:body></w:document>`)
	if len(c.notes) > 0 {
		if existing := p.Files["word/footnotes.xml"]; len(existing) > 0 {
			spans, err := XMLSpans(existing)
			if err != nil {
				return err
			}
			if len(spans) == 0 || spans[0].Name.Local != "footnotes" || spans[0].CloseStart <= 0 {
				return fmt.Errorf("invalid existing footnotes part")
			}
			updated := make([]byte, 0, len(existing)+len(strings.Join(c.notes, "")))
			updated = append(updated, existing[:spans[0].CloseStart]...)
			updated = append(updated, strings.Join(c.notes, "")...)
			updated = append(updated, existing[spans[0].CloseStart:]...)
			p.Files["word/footnotes.xml"] = updated
		} else {
			p.Files["word/footnotes.xml"] = []byte(`<w:footnotes xmlns:w="` + W + `" xmlns:r="` + R + `"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` + strings.Join(c.notes, "") + `</w:footnotes>`)
		}
		if e := p.ContentType("word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"); e != nil {
			return e
		}
		if e := p.Relationship("word/document.xml", "rIdFootnotes", R+"/footnotes", "footnotes.xml", ""); e != nil {
			return e
		}
	}
	return nil
}

func nextFootnoteID(p *Package) (int, error) {
	b := p.Files["word/footnotes.xml"]
	if len(b) == 0 {
		return 1, nil
	}
	spans, err := XMLSpans(b)
	if err != nil {
		return 0, err
	}
	maxID := 0
	for _, span := range spans {
		if span.Name.Space != W || span.Name.Local != "footnote" || span.Depth != 1 {
			continue
		}
		value := span.Attribute(W, "id")
		if value == "" {
			return 0, fmt.Errorf("existing footnote has no id")
		}
		id, err := strconv.Atoi(value)
		if err != nil {
			return 0, fmt.Errorf("invalid existing footnote id %q", value)
		}
		if id > maxID {
			maxID = id
		}
	}
	return maxID + 1, nil
}
func AddBuildingBlocks(p *Package, blocks []BuildingBlock, asset func(string) ([]byte, error)) error {
	work := p.Clone()
	if err := addBuildingBlocks(work, blocks, asset); err != nil {
		return err
	}
	*p = *work
	return nil
}

func addBuildingBlocks(p *Package, blocks []BuildingBlock, asset func(string) ([]byte, error)) error {
	if len(blocks) == 0 {
		return nil
	}
	part := "word/glossary/document.xml"
	b := p.Files[part]
	if b == nil {
		b = []byte(`<w:glossaryDocument xmlns:w="` + W + `" xmlns:r="` + R + `"><w:docParts/></w:glossaryDocument>`)
	}
	var e error
	b, e = EnsureNamespace(b, "w", W)
	if e != nil {
		return e
	}
	b, e = EnsureNamespace(b, "r", R)
	if e != nil {
		return e
	}
	c := &composer{p: p, source: part, asset: asset, next: nextDocumentID(p)}
	seenNames := map[string]bool{}
	for _, block := range blocks {
		if block.Name == "" {
			return fmt.Errorf("saved part name required")
		}
		nameKey := strings.ToLower(block.Name)
		if seenNames[nameKey] {
			return fmt.Errorf("duplicate saved part name %q", block.Name)
		}
		seenNames[nameKey] = true
		body, e := c.blocks(block.Blocks)
		if e != nil {
			return e
		}
		gallery := block.Gallery
		if gallery == "" {
			gallery = "autoTxt"
		}
		cat := block.Category
		if cat == "" {
			cat = "General"
		}
		id := strings.ToUpper(Hash([]byte(block.Name))[:32])
		guid := fmt.Sprintf("{%s-%s-%s-%s-%s}", id[:8], id[8:12], id[12:16], id[16:20], id[20:])
		node := `<w:docPart><w:docPartPr><w:name w:val="` + Esc(block.Name) + `"/><w:category><w:name w:val="` + Esc(cat) + `"/><w:gallery w:val="` + Esc(gallery) + `"/></w:category><w:behaviors><w:behavior w:val="content"/></w:behaviors><w:description w:val="` + Esc(block.Description) + `"/><w:guid w:val="` + guid + `"/></w:docPartPr><w:docPartBody>` + body + `</w:docPartBody></w:docPart>`
		spans, e := XMLSpans(b)
		if e != nil {
			return e
		}
		var parent *XMLSpan
		replaced := false
		for i := range spans {
			sp := spans[i]
			if sp.Name.Local == "docParts" {
				copy := sp
				parent = &copy
			}
			if sp.Name.Local != "docPart" {
				continue
			}
			for _, n := range spans {
				if n.Start > sp.Start && n.End < sp.End && n.Name.Local == "name" && n.Depth == sp.Depth+2 && strings.EqualFold(n.Attribute(W, "val"), block.Name) {
					b = bytes.Join([][]byte{b[:sp.Start], []byte(node), b[sp.End:]}, nil)
					replaced = true
					break
				}
			}
			if replaced {
				break
			}
		}
		if !replaced {
			if parent == nil {
				return fmt.Errorf("glossary has no docParts container")
			}
			raw, e := InsertXML(b[parent.Start:parent.End], node)
			if e != nil {
				return e
			}
			b = bytes.Join([][]byte{b[:parent.Start], raw, b[parent.End:]}, nil)
		}
	}
	p.Files[part] = b
	if e := p.ContentType(part, "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"); e != nil {
		return e
	}
	if e := p.Relationship("word/document.xml", "rIdGlossary", R+"/glossaryDocument", "glossary/document.xml", ""); e != nil {
		return e
	}
	// Glossary content uses its own relationships, including references to styles.
	if p.Files["word/styles.xml"] != nil {
		if _, exists := p.Files["word/glossary/styles.xml"]; !exists {
			p.Files["word/glossary/styles.xml"] = append([]byte(nil), p.Files["word/styles.xml"]...)
		}
		if e := p.ContentType("word/glossary/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"); e != nil {
			return e
		}
		if e := p.Relationship(part, "wwStyles", R+"/styles", "styles.xml", ""); e != nil {
			return e
		}
	}
	if p.Files["word/numbering.xml"] != nil {
		if _, exists := p.Files["word/glossary/numbering.xml"]; !exists {
			p.Files["word/glossary/numbering.xml"] = append([]byte(nil), p.Files["word/numbering.xml"]...)
		}
		if e := p.ContentType("word/glossary/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"); e != nil {
			return e
		}
		if e := p.Relationship(part, "wwNumbering", R+"/numbering", "numbering.xml", ""); e != nil {
			return e
		}
	}
	return nil
}
func Catalog(p *Package) map[string]any {
	styles := []map[string]any{}
	if b := p.Files["word/styles.xml"]; len(b) > 0 {
		if spans, e := XMLSpans(b); e == nil {
			for i, n := range spans {
				if n.Depth == 1 && n.Name.Local == "style" {
					name := ""
					for _, ch := range spans[i+1:] {
						if ch.Start >= n.End {
							break
						}
						if ch.Start > n.Start && ch.End < n.End && ch.Name.Local == "name" && ch.Depth == 2 {
							name = ch.Attribute(W, "val")
							break
						}
					}
					styles = append(styles, map[string]any{"id": n.Attribute(W, "styleId"), "name": name, "type": n.Attribute(W, "type"), "xml": string(b[n.Start:n.End])})
				}
			}
		}
	}
	blocks := []string{}
	if b := p.Files["word/glossary/document.xml"]; len(b) > 0 {
		if spans, e := XMLSpans(b); e == nil {
			for _, n := range spans {
				if n.Name.Local == "name" && n.Depth == 4 {
					blocks = append(blocks, n.Attribute(W, "val"))
				}
			}
		}
	}
	return map[string]any{"styles": styles, "building_blocks": blocks, "note": "Raw native definitions, not a simulated layout cascade. Inspect document.xml and actual Word renders for effective layout."}
}

// Avoid new drawing/content-control/bookmark IDs colliding with preserved parts.
func nextDocumentID(p *Package) int {
	n := 1000
	for name, b := range p.Files {
		if !strings.HasSuffix(name, ".xml") {
			continue
		}
		spans, e := XMLSpans(b)
		if e != nil {
			continue
		}
		for _, sp := range spans {
			if sp.Name.Local != "docPr" && sp.Name.Local != "cNvPr" && sp.Name.Local != "bookmarkStart" && sp.Name.Local != "id" {
				continue
			}
			for _, a := range sp.Attr {
				if a.Name.Local == "id" || a.Name.Local == "val" {
					x, e := strconv.Atoi(a.Value)
					if e == nil && x >= n && x < 1<<30 {
						n = x + 1
					}
				}
			}
		}
	}
	return n
}
