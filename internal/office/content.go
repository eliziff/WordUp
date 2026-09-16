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
	var n float64
	var err error
	switch value := v.(type) {
	case float64:
		n = value
	case float32:
		n = float64(value)
	case int:
		n = float64(value)
	case int8:
		n = float64(value)
	case int16:
		n = float64(value)
	case int32:
		n = float64(value)
	case int64:
		n = float64(value)
	case uint:
		n = float64(value)
	case uint8:
		n = float64(value)
	case uint16:
		n = float64(value)
	case uint32:
		n = float64(value)
	case uint64:
		n = float64(value)
	case json.Number:
		n, err = value.Float64()
	default:
		err = fmt.Errorf("finite numeric measurement required")
	}
	if err != nil || !finiteFloat(n) {
		return 0, fmt.Errorf("finite numeric measurement required")
	}
	return n, nil
}

func finiteFloat(v float64) bool {
	return !math.IsNaN(v) && !math.IsInf(v, 0)
}

func twips(v float64) (string, error) {
	if !finiteFloat(v) {
		return "", fmt.Errorf("finite measurement required")
	}
	n := math.Round(v * 20)
	if !finiteFloat(n) || n < math.MinInt32 || n > math.MaxInt32 {
		return "", fmt.Errorf("measurement outside 32-bit twip range")
	}
	return strconv.FormatInt(int64(n), 10), nil
}

// integer is for OOXML fields whose schema is an integer, not a measurement.
// Do not round a caller's value here: a fractional outline or numbering level
// would otherwise produce a different document than the recipe requested.
func integer(v any) (int, error) {
	n, err := numeric(v)
	if err != nil || math.Trunc(n) != n || n < -2147483648 || n > 2147483647 {
		return 0, fmt.Errorf("32-bit integer required")
	}
	return int(n), nil
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
func patchAttrs(b []byte, patch map[string]string, namespaces map[string]string) ([]byte, error) {
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
		// Match expanded Word names, not the spelling of an inherited prefix.
		if strings.HasPrefix(k, "w:") {
			for _, a := range s.Attr {
				if a.Name.Local == k[2:] && (namespaces[a.Name.Space] == W || a.Name.Space == "w" && namespaces["w"] == "") {
					delete(attrs, qname(a.Name))
				}
			}
		}
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

// ApplyStyleXML edits one native styles or numbering source part. It shares
// the package author's batch editor and returns no candidate on invalid input.
// The caller owns persistence and optimistic concurrency.
func ApplyStyleXML(source []byte, recipe StyleRecipe) ([]byte, error) {
	nodes, err := XMLSpans(source)
	if err != nil {
		return nil, err
	}
	if nodes[0].Name.Space != W {
		return nil, fmt.Errorf("Word styles or numbering XML required")
	}
	source, err = EnsureNamespace(source, "w", W)
	if err != nil {
		return nil, err
	}
	switch nodes[0].Name.Local {
	case "styles":
		if len(recipe.Styles) == 0 || len(recipe.Numbering) != 0 {
			return nil, fmt.Errorf("styles source requires only a nonempty styles recipe")
		}
		source, err = editStyles(source, recipe.Styles)
	case "numbering":
		if len(recipe.Numbering) == 0 || len(recipe.Styles) != 0 {
			return nil, fmt.Errorf("numbering source requires only a nonempty numbering recipe")
		}
		source, err = editNumbering(source, recipe.Numbering)
	default:
		return nil, fmt.Errorf("Word styles or numbering XML required")
	}
	if err == nil {
		_, err = XMLSpans(source)
	}
	if err != nil {
		return nil, err
	}
	return source, nil
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
	for _, part := range []struct {
		name  string
		count int
		edit  func([]byte) ([]byte, error)
	}{
		{"styles", len(r.Styles), func(b []byte) ([]byte, error) { return editStyles(b, r.Styles) }},
		{"numbering", len(r.Numbering), func(b []byte) ([]byte, error) { return editNumbering(b, r.Numbering) }},
	} {
		if part.count == 0 {
			continue
		}
		name := "word/" + part.name + ".xml"
		b := p.Files[name]
		if b == nil {
			b = []byte(`<w:` + part.name + ` xmlns:w="` + W + `"/>`)
		}
		b, err := EnsureNamespace(b, "w", W)
		if err != nil {
			return err
		}
		b, err = part.edit(b)
		if err != nil {
			return err
		}
		p.Files[name] = b
		if err = p.ContentType(name, "application/vnd.openxmlformats-officedocument.wordprocessingml."+part.name+"+xml"); err != nil {
			return err
		}
		id := "rId" + strings.ToUpper(part.name[:1]) + part.name[1:]
		if err = p.Relationship("word/document.xml", id, R+"/"+part.name, part.name+".xml", ""); err != nil {
			return err
		}
	}
	return nil
}

func editStyles(b []byte, styles []StyleSpec) ([]byte, error) {
	spans, err := XMLSpans(b)
	if err != nil {
		return nil, err
	}
	existing := map[string]XMLSpan{}
	for _, n := range spans {
		if n.Depth == 1 && n.Name.Space == W && n.Name.Local == "style" {
			id := n.Attribute(W, "styleId")
			if _, ok := existing[id]; ok {
				return nil, fmt.Errorf("duplicate style id %q", id)
			}
			existing[id] = n
		}
	}
	// A repeated recipe ID sees its preceding edit, but the package part is
	// parsed and copied only once regardless of the number of style updates.
	updated, order := map[string][]byte{}, []string{}
	for _, s := range styles {
		raw, seen := updated[s.ID]
		if !seen {
			if n, ok := existing[s.ID]; ok {
				raw = b[n.Start:n.End]
			}
			order = append(order, s.ID)
		}
		scope := spans[0].namespaces
		if n, ok := existing[s.ID]; ok {
			scope = n.namespaces
		}
		updated[s.ID], err = editStyle(raw, s, scope)
		if err != nil {
			return nil, fmt.Errorf("style %q: %w", s.ID, err)
		}
	}
	patches := []XMLPatch{}
	var added strings.Builder
	for _, id := range order {
		if n, ok := existing[id]; ok {
			patches = append(patches, XMLPatch{n.Start, n.End - n.Start, string(updated[id])})
		} else {
			added.Write(updated[id])
		}
	}
	if added.Len() > 0 {
		patches = append(patches, insertAtRoot(b, spans[0], added.String()))
	}
	return spliceXML(b, patches)
}

func editStyle(raw []byte, s StyleSpec, scope map[string]string) ([]byte, error) {
	if s.ID == "" {
		return nil, fmt.Errorf("style id required")
	}
	existing := len(raw) > 0
	if !existing {
		kind := s.Type
		if kind == "" {
			kind = "paragraph"
		}
		raw = []byte(`<w:style xmlns:w="` + W + `" w:type="` + Esc(kind) + `" w:styleId="` + Esc(s.ID) + `" w:customStyle="1"/>`)
	}
	nodes, err := xmlSpans(raw, scope)
	if err != nil {
		return nil, err
	}
	if s.Type != "" && nodes[0].Attribute(W, "type") != s.Type {
		return nil, fmt.Errorf("style type conversion requires a new ID")
	}
	if s.XML != "" {
		node := []byte(s.XML)
		spans, err := xmlSpans(node, scope)
		if err != nil {
			return nil, err
		}
		if spans[0].Name.Local != "style" || (spans[0].Name.Space != W && spans[0].Name.Space != "w") || spans[0].Attribute(W, "styleId") != s.ID {
			return nil, fmt.Errorf("raw style identity mismatch")
		}
		return node, nil
	}
	changes := map[string]childChange{}
	name := s.Name
	if name == "" && !existing {
		name = s.ID
	}
	for _, kv := range [][2]string{{"name", name}, {"basedOn", s.BasedOn}, {"next", s.Next}, {"link", s.Linked}} {
		if kv[1] != "" {
			changes[kv[0]] = childChange{attrs: map[string]string{"w:val": kv[1]}}
		}
	}
	if s.Quick != nil {
		value, _ := boolVal(*s.Quick)
		changes["qFormat"] = childChange{attrs: map[string]string{"w:val": value}}
	}
	for _, it := range []struct {
		tag, kind string
		spec      map[string]any
	}{{"pPr", "paragraph", s.Paragraph}, {"rPr", "run", s.Run}} {
		if it.spec == nil {
			continue
		}
		var child []byte
		for _, n := range nodes {
			if n.Depth == 1 && n.Name.Local == it.tag && (n.Name.Space == W || n.Name.Space == "w") {
				child = raw[n.Start:n.End]
				break
			}
		}
		childScope := nodes[0].namespaces
		for _, n := range nodes {
			if n.Depth == 1 && n.Name.Local == it.tag && n.Name.Space == W {
				childScope = n.namespaces
				break
			}
		}
		child, err = props(child, it.kind, it.spec, childScope)
		if err != nil {
			return nil, err
		}
		changes[it.tag] = childChange{xml: string(child)}
	}
	return mergeChildren(raw, styleOrder, changes, scope)
}

func editNumbering(b []byte, numbering []NumberingSpec) ([]byte, error) {
	spans, err := XMLSpans(b)
	if err != nil {
		return nil, err
	}
	next, at := 0, spans[0].CloseStart
	existing := map[int]XMLSpan{}
	for _, n := range spans {
		if n.Depth != 1 || n.Name.Space != W {
			continue
		}
		switch n.Name.Local {
		case "abstractNum":
			id, err := strconv.Atoi(n.Attribute(W, "abstractNumId"))
			if err != nil || id < 0 || id >= math.MaxInt32 {
				return nil, fmt.Errorf("invalid abstract numbering id")
			}
			next = max(next, id+1)
		case "num":
			id, err := strconv.Atoi(n.Attribute(W, "numId"))
			if err != nil {
				return nil, err
			}
			if _, ok := existing[id]; ok {
				return nil, fmt.Errorf("duplicate numbering id %d", id)
			}
			existing[id] = n
			at = min(at, n.Start)
		}
	}
	var abstracts, added strings.Builder
	patches, seen := []XMLPatch{}, map[int]bool{}
	for _, n := range numbering {
		if n.ID < 1 || n.ID > math.MaxInt32 || len(n.Levels) == 0 || len(n.Levels) > 9 || seen[n.ID] || next >= math.MaxInt32 {
			return nil, fmt.Errorf("invalid numbering recipe")
		}
		seen[n.ID] = true
		abstracts.WriteString(fmt.Sprintf(`<w:abstractNum w:abstractNumId="%d"><w:multiLevelType w:val="multilevel"/>`, next))
		levels := map[int]bool{}
		for _, l := range n.Levels {
			if l.Level < 0 || l.Level > 8 || levels[l.Level] {
				return nil, fmt.Errorf("invalid numbering level")
			}
			levels[l.Level] = true
			start := l.Start
			if start == 0 {
				start = 1
			}
			if start < 1 || l.Format == "" || l.Text == "" {
				return nil, fmt.Errorf("numbering level needs a positive start, format, and text")
			}
			abstracts.WriteString(fmt.Sprintf(`<w:lvl w:ilvl="%d"><w:start w:val="%d"/><w:numFmt w:val="%s"/><w:lvlText w:val="%s"/>`, l.Level, start, Esc(l.Format), Esc(l.Text)))
			if l.Suffix != "" {
				abstracts.WriteString(`<w:suff w:val="` + Esc(l.Suffix) + `"/>`)
			}
			for _, it := range []struct {
				kind string
				spec map[string]any
			}{{"paragraph", l.Paragraph}, {"run", l.Run}} {
				x, err := props(nil, it.kind, it.spec)
				if err != nil {
					return nil, err
				}
				abstracts.Write(x)
			}
			abstracts.WriteString(`</w:lvl>`)
		}
		abstracts.WriteString(`</w:abstractNum>`)
		num := fmt.Sprintf(`<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>`, n.ID, next)
		if old, ok := existing[n.ID]; ok {
			// A concrete numbering ID may already carry Word-authored level
			// overrides (a restart or per-level start). Update only its abstract
			// definition link; replacing the whole <w:num> node would silently
			// discard those unrelated children.
			node, err := mergeChildren(b[old.Start:old.End], "abstractNumId lvlOverride", map[string]childChange{"abstractNumId": {attrs: map[string]string{"w:val": strconv.Itoa(next)}}}, old.namespaces)
			if err != nil {
				return nil, err
			}
			patches = append(patches, XMLPatch{old.Start, old.End - old.Start, string(node)})
		} else {
			added.WriteString(num)
		}
		next++
	}
	if at == spans[0].CloseStart {
		patches = append(patches, insertAtRoot(b, spans[0], abstracts.String()+added.String()))
	} else {
		patches = append(patches, XMLPatch{Offset: at, Text: abstracts.String()})
		if added.Len() > 0 {
			patches = append(patches, insertAtRoot(b, spans[0], added.String()))
		}
	}
	return spliceXML(b, patches)
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
	if !finiteFloat(in.WidthPT) || !finiteFloat(in.HeightPT) || in.WidthPT <= 0 || in.HeightPT <= 0 || in.WidthPT > 10000 || in.HeightPT > 10000 {
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
	if cx <= 0 || cy <= 0 {
		return "", fmt.Errorf("image dimensions must be at least one EMU")
	}
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
				columnWidth, err := twips(w)
				if err != nil || w <= 0 || columnWidth == "0" {
					return "", fmt.Errorf("positive column width required")
				}
				out.WriteString(`<w:gridCol w:w="` + columnWidth + `"/>`)
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
					if !finiteFloat(cell.WidthPT) || cell.WidthPT < 0 {
						return "", fmt.Errorf("finite cell width required")
					}
					if cell.WidthPT > 0 {
						cellWidth, err := twips(cell.WidthPT)
						if err != nil || cellWidth == "0" {
							return "", fmt.Errorf("cell width outside supported range")
						}
						out.WriteString(`<w:tcW w:w="` + cellWidth + `" w:type="dxa"/>`)
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
		if !finiteFloat(w) || !finiteFloat(h) || w <= 0 || h <= 0 || w > 1584 || h > 1584 {
			return fmt.Errorf("page dimensions outside Word's supported 22-inch range")
		}
		orient := ""
		if pg.Landscape {
			orient = ` w:orient="landscape"`
		}
		pageWidth, widthErr := twips(w)
		pageHeight, heightErr := twips(h)
		if widthErr != nil || heightErr != nil || pageWidth == "0" || pageHeight == "0" {
			return fmt.Errorf("page dimensions outside Word's supported 32-bit twip range")
		}
		b.WriteString(`<w:pgSz w:w="` + pageWidth + `" w:h="` + pageHeight + `"` + orient + `/>`)
		m := map[string]float64{"top": 72, "bottom": 72, "left": 72, "right": 72, "header": 36, "footer": 36, "gutter": 0}
		for k, v := range pg.MarginsPT {
			if _, ok := m[k]; !ok {
				return fmt.Errorf("unknown page margin %s", k)
			}
			if !finiteFloat(v) || v < 0 || v > 1584 {
				return fmt.Errorf("invalid margin")
			}
			m[k] = v
		}
		a := map[string]string{}
		for k, v := range m {
			value, err := twips(v)
			if err != nil {
				return fmt.Errorf("margin outside supported range")
			}
			a["w:"+k] = value
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
