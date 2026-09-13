package office

import (
	"encoding/hex"
	"encoding/json"
	"fmt"
	"math"
	"regexp"
	"sort"
	"strings"
)

var siteSpec = recordSpec{name: "Site", major: 0, fields: fields(field{0, "Name", 4, 's'}, field{1, "Tag", 4, 's'}, field{2, "ID", 4, 'i'}, field{3, "HelpContextID", 4, 'i'}, field{4, "BitFlags", 4, 0}, field{5, "ObjectStreamSize", 4, 0}, field{6, "TabIndex", 2, 'i'}, field{7, "ClsidCacheIndex", 2, 0}, field{9, "GroupID", 2, 0}, field{11, "ControlTipText", 4, 's'}, field{12, "RuntimeLicKey", 4, 's'}, field{13, "ControlSource", 4, 's'}, field{14, "RowSource", 4, 's'}), extra: []extra{{0, "Name", 's', ""}, {1, "Tag", 's', ""}, {8, "Position", 'z', ""}, {11, "ControlTipText", 's', ""}, {12, "RuntimeLicKey", 's', ""}, {13, "ControlSource", 's', ""}, {14, "RowSource", 's', ""}}}

type Design struct {
	Name       string          `json:"name"`
	Mode       string          `json:"mode,omitempty"`
	Properties map[string]any  `json:"properties,omitempty"`
	Controls   []ControlDesign `json:"controls,omitempty"`
	Remove     []string        `json:"remove,omitempty"`
	Pages      []ControlDesign `json:"pages,omitempty"`
}
type ControlDesign struct {
	Name       string          `json:"name"`
	Type       string          `json:"type"`
	Properties map[string]any  `json:"properties,omitempty"`
	Controls   []ControlDesign `json:"controls,omitempty"`
	Pages      []ControlDesign `json:"pages,omitempty"`
	Remove     []string        `json:"remove,omitempty"`
	Opaque     bool            `json:"opaque,omitempty"`
}
type formControl struct {
	site   *formRecord
	record *formRecord
	raw    []byte
	child  *formLevel
	kind   string
}
type formLevel struct {
	path                               string
	record                             *formRecord
	streams, classes, depths, trailing []byte
	controls                           []*formControl
	rawF, rawO                         []byte
	structural                         bool
	kind                               string
	x                                  []byte
}
type Form struct {
	Name     string
	Codepage int
	CFB      *Compound
	root     *formLevel
	vbframe  string
}

func ReadForm(c *Compound, name string, cp int) (f *Form, err error) {
	defer func() {
		if r := recover(); r != nil {
			f = nil
			err = fmt.Errorf("invalid %s form: %v", name, r)
		}
	}()
	f = &Form{Name: name, Codepage: cp, CFB: c}
	var e error
	f.root, e = readLevel(c, name, cp, 0)
	if e != nil {
		return nil, e
	}
	if data, e := c.Stream(name + "/\x03VBFrame"); e == nil {
		f.vbframe, e = Decode(data, cp)
		if e != nil {
			return nil, e
		}
	}
	return f, nil
}
func readLevel(c *Compound, path string, cp, depth int) (*formLevel, error) {
	if depth > 64 {
		return nil, fmt.Errorf("form nesting limit")
	}
	data, e := c.Stream(path + "/f")
	if e != nil {
		return nil, e
	}
	if len(data) < 8 {
		return nil, fmt.Errorf("short form")
	}
	bound := 4 + int(U16(data, 2))
	if bound > len(data) {
		return nil, fmt.Errorf("form boundary outside data")
	}
	r, e := readRecord(data[:bound], formSpec, cp)
	if e != nil {
		return nil, e
	}
	p := bound
	blob := func(font bool) error {
		if p+16 > len(data) {
			return fmt.Errorf("truncated form blob")
		}
		tag := U32(data, p)
		p += 16
		if font {
			switch tag {
			case 0x0be35203:
				if p+11 > len(data) || data[p] != 1 {
					return fmt.Errorf("invalid StdFont")
				}
				n := int(data[p+10])
				p += 11 + n
			case 0xafc20920:
				if p+4 > len(data) {
					return fmt.Errorf("short TextProps")
				}
				p += 4 + int(U16(data, p+2))
			default:
				return fmt.Errorf("unknown font GUID")
			}
		} else {
			if p+8 > len(data) {
				return fmt.Errorf("short picture")
			}
			n := int(U32(data, p+4))
			p += 8 + n
		}
		if p > len(data) {
			return fmt.Errorf("blob outside form")
		}
		return nil
	}
	for _, b := range []struct {
		bit  uint
		font bool
	}{{15, false}, {20, true}, {21, false}} {
		if r.mask&(1<<b.bit) != 0 {
			if e := blob(b.font); e != nil {
				return nil, e
			}
		}
	}
	l := &formLevel{path: path, record: r, streams: append([]byte(nil), data[bound:p]...), rawF: append([]byte(nil), data...)}
	try := func(table bool) (sites []*formRecord, classes, depths, trailing []byte, ok bool) {
		defer func() {
			if recover() != nil {
				ok = false
			}
		}()
		q := p
		if table {
			count := int(U16(data, q))
			q += 2
			for i := 0; i < count; i++ {
				if U16(data, q) != 0 {
					return
				}
				n := int(U16(data, q+2))
				q += 4 + n
				if q > len(data) {
					return
				}
			}
		}
		classes = append([]byte(nil), data[p:q]...)
		n, nbytes := int(U32(data, q)), int(U32(data, q+4))
		q += 8
		if n < 0 || n > 10000 || nbytes < 0 || q+nbytes > len(data) {
			return
		}
		stop := q + nbytes
		if stop < len(data) {
			if stop+4 > len(data) || data[stop] != 0 || data[stop+1] != 2 || stop+4+int(U16(data, stop+2)) != len(data) {
				return
			}
		}
		dstart := q
		accounted := 0
		for accounted < n {
			typ := data[q+1]
			q += 2
			if typ&0x80 != 0 {
				run := int(typ & 127)
				if run == 0 {
					return
				}
				accounted += run
				q++
			} else {
				accounted++
			}
		}
		if accounted != n {
			return
		}
		q += (4 - (q-dstart)%4) % 4
		depths = append([]byte(nil), data[dstart:q]...)
		for i := 0; i < n; i++ {
			size := 4 + int(U16(data, q+2))
			if size < 8 || q+size > stop {
				return
			}
			s, err := readRecord(data[q:q+size], siteSpec, cp)
			if err != nil {
				return
			}
			sites = append(sites, s)
			q += size
		}
		if q != stop {
			return
		}
		trailing = append([]byte(nil), data[stop:]...)
		ok = true
		return
	}
	sites, cl, ds, tr, ok := try(true)
	if !ok {
		sites, cl, ds, tr, ok = try(false)
	}
	if !ok {
		return nil, fmt.Errorf("form site counts do not reconcile in %s", path)
	}
	l.classes, l.depths, l.trailing = cl, ds, tr
	objects, e := c.Stream(path + "/o")
	if e != nil {
		objects = nil
	}
	l.rawO = append([]byte(nil), objects...)
	l.x, _ = c.Stream(path + "/x")
	op := 0
	names := map[string]bool{}
	for _, s := range sites {
		idx := int(s.values["ClsidCacheIndex"])
		ctl := &formControl{site: s, kind: "Unknown"}
		for k, i := range kindIndex {
			if i == idx {
				ctl.kind = k
				break
			}
		}
		name := s.strings["Name"]
		if names[strings.ToLower(name)] {
			return nil, fmt.Errorf("duplicate control name %s", name)
		}
		names[strings.ToLower(name)] = true
		n := int(s.values["ObjectStreamSize"])
		if n < 0 || op+n > len(objects) {
			return nil, fmt.Errorf("object stream size mismatch")
		}
		ctl.raw = append([]byte(nil), objects[op:op+n]...)
		op += n
		if idx == 7 || idx == 14 || idx == 57 {
			child := fmt.Sprintf("%s/i%02d", path, s.values["ID"])
			ctl.child, e = readLevel(c, child, cp, depth+1)
			if e != nil {
				return nil, e
			}
		} else if spec, known := specFor(idx); known {
			ctl.record, e = readRecord(ctl.raw, spec, cp)
			if e != nil {
				return nil, fmt.Errorf("%s.%s: %w", path, name, e)
			}
		}
		l.controls = append(l.controls, ctl)
	}
	if op != len(objects) {
		return nil, fmt.Errorf("unclaimed designer object bytes in %s", path)
	}
	return l, nil
}
func properties(r *formRecord, size string) map[string]any {
	p := map[string]any{}
	for _, f := range r.spec.fields {
		if r.mask&(1<<f.bit) == 0 || f.kind == 'b' {
			continue
		}
		if f.kind == 's' {
			p[f.name] = r.strings[f.name]
		} else {
			p[f.name] = r.values[f.name]
		}
	}
	if size != "" {
		if _, ok := r.extra[size]; ok {
			w, h := r.dimensions(size)
			p["Width"], p["Height"] = w, h
		}
	}
	return p
}
func designControls(l *formLevel, cp int) []ControlDesign {
	out := []ControlDesign{}
	for _, c := range l.controls {
		if c.site.strings["Name"] == "" {
			continue
		} // MultiPage's internal TabStrip is not a designer control.
		p := map[string]any{}
		if c.record != nil {
			p = properties(c.record, "Size")
		}
		if c.child != nil {
			p = properties(c.child.record, "DisplayedSize")
			delete(p, "NextAvailableID")
			delete(p, "BooleanProperties")
		}
		left, top := c.site.dimensions("Position")
		p["Left"], p["Top"] = left, top
		for _, k := range []string{"Tag", "ControlTipText", "ControlSource", "RowSource"} {
			if x, ok := c.site.strings[k]; ok {
				p[k] = x
			}
		}
		if x, ok := c.site.values["TabIndex"]; ok {
			p["TabIndex"] = x
		}
		d := ControlDesign{Name: c.site.strings["Name"], Type: c.kind, Properties: p, Opaque: c.record == nil && c.child == nil}
		if c.child != nil {
			d.Controls = designControls(c.child, cp)
			if c.kind == "MultiPage" {
				d.Pages = d.Controls
				d.Controls = nil
				captions, _ := pageStrings(c.child, "Items", cp)
				for i := range d.Pages {
					if i < len(captions) {
						d.Pages[i].Properties["Caption"] = captions[i]
					}
				}
			}
		}
		out = append(out, d)
	}
	return out
}
func (f *Form) Design() Design {
	p := properties(f.root.record, "DisplayedSize")
	delete(p, "NextAvailableID")
	delete(p, "BooleanProperties")
	if f.vbframe != "" {
		for _, kv := range []struct{ name, field string }{{"Caption", "Caption"}, {"Width", "ClientWidth"}, {"Height", "ClientHeight"}, {"StartUpPosition", "StartUpPosition"}} {
			pattern := regexp.MustCompile(`(?m)^\s*` + kv.field + `\s*=\s*(.+?)\s*(?:'.*)?$`)
			m := pattern.FindStringSubmatch(f.vbframe)
			if len(m) == 2 {
				if kv.name == "Caption" {
					p[kv.name] = strings.ReplaceAll(strings.Trim(m[1], `"`), `""`, `"`)
				} else {
					var n float64
					if _, e := fmt.Sscanf(m[1], "%f", &n); e == nil {
						if kv.name == "Width" || kv.name == "Height" {
							n /= 20
						}
						p[kv.name] = n
					}
				}
			}
		}
	}
	return Design{Name: f.Name, Mode: "patch", Properties: p, Controls: designControls(f.root, f.Codepage)}
}
func NewForm(name string, cp int) (*Form, error) {
	if !ValidIdentifier(name) {
		return nil, fmt.Errorf("invalid form identifier")
	}
	r, e := defaultControl("Form", name)
	if e != nil {
		return nil, e
	}
	_ = r.size("DisplayedSize", 360, 240)
	_ = r.size("LogicalSize", 0, 0)
	_ = r.set("BooleanProperties", nil)
	// Word compares the designer's ShapeCookie with VBFrame.TypeInfoVer.
	// A missing cookie defaults to zero and makes a newly saved form unloadable.
	_ = r.set("ShapeCookie", 1)
	return &Form{Name: name, Codepage: cp, CFB: NewCompound(), root: &formLevel{path: name, record: r, structural: true, classes: []byte{0, 0}}}, nil
}
func floatValue(m map[string]any, key string, def float64) (float64, error) {
	x, ok := m[key]
	if !ok {
		return def, nil
	}
	switch v := x.(type) {
	case float64:
		if math.IsNaN(v) || math.IsInf(v, 0) || math.Abs(v) > 100000 {
			return 0, fmt.Errorf("invalid geometry")
		}
		return v, nil
	case int:
		return float64(v), nil
	case int64:
		return float64(v), nil
	}
	return 0, fmt.Errorf("numeric %s required", key)
}
func applyRecord(r *formRecord, p map[string]any, size string) error {
	w, h := r.dimensions(size)
	var e error
	if _, ok := p["Width"]; ok {
		w, e = floatValue(p, "Width", w)
		if e != nil {
			return e
		}
	}
	if _, ok := p["Height"]; ok {
		h, e = floatValue(p, "Height", h)
		if e != nil {
			return e
		}
	}
	if _, ok := p["Width"]; ok {
		if e = r.size(size, w, h); e != nil {
			return e
		}
	}
	if _, ok := p["Height"]; ok {
		if e = r.size(size, w, h); e != nil {
			return e
		}
	}
	for k, v := range p {
		switch k {
		case "Width", "Height", "Left", "Top", "TabIndex", "Tag", "ControlTipText", "ControlSource", "RowSource", "StartUpPosition":
			continue
		case "Font":
			m, ok := v.(map[string]any)
			if !ok {
				return fmt.Errorf("Font must be object")
			}
			if e := r.applyFont(m); e != nil {
				return e
			}
			continue
		case "TextAlign":
			if e := r.applyFont(map[string]any{"TextAlign": v}); e != nil {
				return e
			}
			continue
		case "PictureBase64":
			if e := r.applyPicture(v); e != nil {
				return e
			}
			continue
		case "Enabled", "Locked", "MultiLine", "WordWrap", "AutoSize", "TabKeyBehavior", "EnterKeyBehavior":
			bits := map[string]uint{"Enabled": 1, "Locked": 2, "MultiLine": 31, "WordWrap": 23, "AutoSize": 28, "TabKeyBehavior": 22, "EnterKeyBehavior": 20}
			flag, ok := v.(bool)
			if !ok {
				return fmt.Errorf("boolean %s required", k)
			}
			n := r.values["VariousPropertyBits"]
			if _, has := r.values["VariousPropertyBits"]; !has {
				n = 0x1b
				if r.spec.wide {
					n = 0x2c80081b
				} else if r.spec.name == "Label" {
					n = 0x80001b
				}
			}
			if flag {
				n |= 1 << bits[k]
			} else {
				n &^= 1 << bits[k]
			}
			if e = r.set("VariousPropertyBits", n); e != nil {
				return e
			}
			continue
		}
		if e = r.set(k, v); e != nil {
			return e
		}
	}
	return nil
}
func pictureEnd(tail []byte, p int) (int, error) {
	if p+24 > len(tail) {
		return 0, fmt.Errorf("truncated control picture")
	}
	n := int(U32(tail, p+20))
	if n < 0 || p+24+n > len(tail) {
		return 0, fmt.Errorf("picture size out of bounds")
	}
	return p + 24 + n, nil
}
func (r *formRecord) applyFont(m map[string]any) error {
	if !r.spec.font {
		return fmt.Errorf("this record has no TextProps; root form fonts use the native stream asset")
	}
	p := 0
	var e error
	for _, b := range r.spec.blobs {
		if r.mask&(1<<b.bit) != 0 {
			p, e = pictureEnd(r.tail, p)
			if e != nil {
				return e
			}
		}
	}
	if p+4 > len(r.tail) {
		return fmt.Errorf("missing TextProps")
	}
	end := p + 4 + int(U16(r.tail, p+2))
	if end > len(r.tail) {
		return fmt.Errorf("TextProps outside record")
	}
	f, e := readRecord(r.tail[p:end], textSpec, 1252)
	if e != nil {
		return e
	}
	for k, v := range m {
		switch k {
		case "Name":
			e = f.set("FontName", v)
		case "TextAlign":
			// MSForms uses left/center/right = 1/2/3; MS-OFORMS
			// ParagraphAlign stores PFA_LEFT/RIGHT/CENTER = 1/2/3.
			var n float64
			n, e = floatValue(m, k, 1)
			if e == nil {
				if n != 1 && n != 2 && n != 3 {
					return fmt.Errorf("TextAlign must be 1 (left), 2 (center), or 3 (right)")
				}
				e = f.set("ParagraphAlign", map[int]int{1: 1, 2: 3, 3: 2}[int(n)])
			}
		case "Size":
			var n float64
			n, e = floatValue(m, k, 8.25)
			if e == nil {
				e = f.set("FontHeight", int(math.Round(n*20)))
			}
		case "Bold", "Italic", "Underline", "Strikethrough":
			on, ok := v.(bool)
			if !ok {
				return fmt.Errorf("font boolean required")
			}
			bit := map[string]uint{"Bold": 0, "Italic": 1, "Underline": 2, "Strikethrough": 3}[k]
			n := f.values["FontEffects"]
			if on {
				n |= 1 << bit
			} else {
				n &^= 1 << bit
			}
			e = f.set("FontEffects", n)
		default:
			return fmt.Errorf("unknown font property %s", k)
		}
		if e != nil {
			return e
		}
	}
	b, e := f.bytes(1252)
	if e != nil {
		return e
	}
	r.tail = append(append(append([]byte(nil), r.tail[:p]...), b...), r.tail[end:]...)
	r.dirty = true
	return nil
}
func (f *Form) Apply(d Design) error {
	if d.Mode != "replace" {
		original := f.Design()
		d.Properties = changedFormProperties(d.Properties, original.Properties)
		d.Controls = changedFormControls(d.Controls, original.Controls)
	}
	work := *f
	work.root = cloneLevel(f.root)
	if e := work.apply(d); e != nil {
		return e
	}
	if _, e := work.Streams(); e != nil {
		return e
	}
	*f = work
	return nil
}

// Imported designs include observed properties that must not be rewritten when
// unchanged. Compare their JSON values so decoded numbers and native integers agree.
func changedFormProperties(wanted, original map[string]any) map[string]any {
	out := map[string]any{}
	for key, value := range wanted {
		before, exists := original[key]
		a, ea := json.Marshal(value)
		b, eb := json.Marshal(before)
		if exists && ea == nil && eb == nil && string(a) == string(b) {
			continue
		}
		out[key] = value
	}
	return out
}

func changedFormControls(wanted, original []ControlDesign) []ControlDesign {
	byName := map[string]ControlDesign{}
	for _, c := range original {
		byName[strings.ToLower(c.Name)] = c
	}
	out := append([]ControlDesign(nil), wanted...)
	for i, c := range out {
		if before, ok := byName[strings.ToLower(c.Name)]; ok {
			out[i].Properties = changedFormProperties(c.Properties, before.Properties)
			out[i].Controls = changedFormControls(c.Controls, before.Controls)
			out[i].Pages = changedFormControls(c.Pages, before.Pages)
		}
	}
	return out
}
func (f *Form) apply(d Design) error {
	if d.Name != f.Name {
		return fmt.Errorf("form identity mismatch")
	}
	if d.Mode == "preserve" {
		return nil
	}
	if d.Mode == "replace" {
		fresh, e := NewForm(d.Name, f.Codepage)
		if e != nil {
			return e
		}
		f.root = fresh.root
		f.vbframe = ""
	}
	canvas := map[string]any{}
	for key, value := range d.Properties {
		if key != "Caption" {
			canvas[key] = value
		}
	}
	if e := applyRecord(f.root.record, canvas, "DisplayedSize"); e != nil {
		return e
	}
	if e := f.applyControls(f.root, d.Controls, d.Remove, f.Codepage); e != nil {
		return e
	}
	// The host UserForm frame is separate from the MSForms canvas. Update both.
	if f.vbframe == "" {
		f.vbframe = fmt.Sprintf("VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} %s\r\n   Caption = \"%s\"\r\n   ClientHeight = 4800\r\n   ClientLeft = 0\r\n   ClientTop = 0\r\n   ClientWidth = 7200\r\n   StartUpPosition = 1\r\n   TypeInfoVer = 1\r\nEnd\r\n", f.Name, f.Name)
	}
	version := regexp.MustCompile(`(?m)^\s*TypeInfoVer\s*=.*\r?$`)
	f.vbframe = version.ReplaceAllStringFunc(f.vbframe, func(string) string {
		return fmt.Sprintf("   TypeInfoVer = %d\r", f.root.record.values["ShapeCookie"])
	})
	for _, item := range []struct {
		key, field string
		scale      float64
	}{{"Caption", "Caption", 1}, {"Width", "ClientWidth", 20}, {"Height", "ClientHeight", 20}, {"StartUpPosition", "StartUpPosition", 1}} {
		v, ok := d.Properties[item.key]
		if !ok {
			continue
		}
		value := ""
		if item.key == "Caption" {
			s, ok := v.(string)
			if !ok {
				return fmt.Errorf("Caption string required")
			}
			value = `"` + strings.ReplaceAll(s, `"`, `""`) + `"`
		} else {
			n, e := floatValue(d.Properties, item.key, 0)
			if e != nil {
				return e
			}
			value = fmt.Sprintf("%.3f", n*item.scale)
		}
		re := regexp.MustCompile(`(?m)^\s*` + item.field + `\s*=.*\r?$`)
		if re.MatchString(f.vbframe) {
			f.vbframe = re.ReplaceAllStringFunc(f.vbframe, func(string) string { return "   " + item.field + " = " + value + "\r" })
		} else {
			f.vbframe = strings.Replace(f.vbframe, "End\r\n", "   "+item.field+" = "+value+"\r\nEnd\r\n", 1)
		}
	}
	return nil
}
func (f *Form) applyControls(l *formLevel, designs []ControlDesign, remove []string, cp int) error {
	if l.kind == "MultiPage" || len(l.x) > 0 {
		return f.applyPages(l, designs, remove, cp)
	}
	for _, name := range remove {
		found := false
		for i, c := range l.controls {
			if c.site.strings["Name"] == name {
				l.controls = append(l.controls[:i], l.controls[i+1:]...)
				l.structural = true
				found = true
				break
			}
		}
		if !found {
			return fmt.Errorf("cannot remove absent control %s", name)
		}
	}
	names := map[string]bool{}
	for _, d := range designs {
		if !ValidIdentifier(d.Name) {
			return fmt.Errorf("invalid control name %s", d.Name)
		}
		if names[strings.ToLower(d.Name)] {
			return fmt.Errorf("duplicate control edit")
		}
		names[strings.ToLower(d.Name)] = true
		var ctl *formControl
		for _, c := range l.controls {
			if strings.EqualFold(c.site.strings["Name"], d.Name) {
				ctl = c
				break
			}
		}
		if ctl == nil {
			next := f.allocateID(l)
			record, e := defaultControl(d.Type, d.Name)
			if e != nil {
				return e
			}
			idx := kindIndex[d.Type]
			s := newRecord(siteSpec)
			_ = s.set("Name", d.Name)
			_ = s.set("ID", next)
			_ = s.set("TabIndex", len(l.controls))
			_ = s.set("ClsidCacheIndex", idx)
			_ = s.size("Position", 0, 0)
			ctl = &formControl{site: s, record: record, kind: d.Type}
			if d.Type == "Frame" || d.Type == "Form" || d.Type == "MultiPage" {
				_ = s.set("BitFlags", 262179)
				ctl.record = nil
				ctl.child = &formLevel{path: fmt.Sprintf("%s/i%02d", l.path, next), record: record, structural: true, kind: d.Type}
			} else {
				_ = s.set("ObjectStreamSize", 0)
			}
			l.controls = append(l.controls, ctl)
			l.structural = true
			if d.Type == "MultiPage" {
				if e := f.initMultiPage(ctl.child); e != nil {
					return e
				}
			}
		} else if d.Type != "" && d.Type != ctl.kind {
			return fmt.Errorf("cannot silently convert control %s from %s to %s", d.Name, ctl.kind, d.Type)
		}
		if ctl.record == nil && ctl.child == nil {
			if !d.Opaque {
				return fmt.Errorf("opaque control %s cannot be rewritten by the portable serializer", d.Name)
			}
			continue
		}
		p := d.Properties
		if p == nil {
			p = map[string]any{}
		}
		left, top := ctl.site.dimensions("Position")
		left, e := floatValue(p, "Left", left)
		if e != nil {
			return e
		}
		top, e = floatValue(p, "Top", top)
		if e != nil {
			return e
		}
		if e = ctl.site.size("Position", left, top); e != nil {
			return e
		}
		for _, k := range []string{"TabIndex", "Tag", "ControlTipText", "ControlSource", "RowSource"} {
			if x, ok := p[k]; ok {
				oldIndex := ctl.site.values["TabIndex"]
				if e = ctl.site.set(k, x); e != nil {
					return e
				}
				if k == "TabIndex" {
					index := ctl.site.values[k]
					if index < 0 || index >= int64(len(l.controls)) {
						return fmt.Errorf("%s: TabIndex must be within its container", d.Name)
					}
					for _, sibling := range l.controls {
						if sibling == ctl {
							continue
						}
						at := sibling.site.values[k]
						if index < oldIndex && at >= index && at < oldIndex {
							_ = sibling.site.set(k, at+1)
						}
						if index > oldIndex && at > oldIndex && at <= index {
							_ = sibling.site.set(k, at-1)
						}
					}
				}
			}
		}
		if ctl.child != nil {
			if e = applyRecord(ctl.child.record, p, "DisplayedSize"); e != nil {
				return e
			}
			ds := d.Controls
			if ctl.kind == "MultiPage" && len(d.Pages) > 0 {
				ds = d.Pages
			}
			if ctl.kind == "MultiPage" && len(ds) == 0 && len(ctl.child.controls) == 1 {
				ds = []ControlDesign{{Name: "Page1", Type: "Form"}, {Name: "Page2", Type: "Form"}}
			}
			if e = f.applyControls(ctl.child, ds, d.Remove, cp); e != nil {
				return e
			}
		} else if e = applyRecord(ctl.record, p, "Size"); e != nil {
			return fmt.Errorf("%s: %w", d.Name, e)
		}
	}
	return nil
}
func (f *Form) Streams() (map[string][]byte, error) {
	out := map[string][]byte{}
	var write func(*formLevel) error
	write = func(l *formLevel) error {
		objects := []byte{}
		sites := []byte{}
		for _, c := range l.controls {
			raw := c.raw
			if c.child != nil {
				if e := write(c.child); e != nil {
					return e
				}
			} else if c.record != nil {
				var e error
				raw, e = c.record.bytes(f.Codepage)
				if e != nil {
					return e
				}
			}
			objects = append(objects, raw...)
			if c.child == nil {
				if e := c.site.set("ObjectStreamSize", len(raw)); e != nil {
					return e
				}
			}
			s, e := c.site.bytes(f.Codepage)
			if e != nil {
				return e
			}
			sites = append(sites, s...)
		}
		record, e := l.record.bytes(f.Codepage)
		if e != nil {
			return e
		}
		depths := l.depths
		if l.structural {
			depths = nil
			remaining := len(l.controls)
			for remaining > 0 {
				n := min(127, remaining)
				depths = append(depths, 0, byte(128+n), 1)
				remaining -= n
			}
			depths = pad(depths, 4)
		}
		b := append(record, l.streams...)
		b = append(b, l.classes...)
		b = append(b, dword(uint32(len(l.controls)))...)
		b = append(b, dword(uint32(len(depths)+len(sites)))...)
		b = append(b, depths...)
		b = append(b, sites...)
		b = append(b, l.trailing...)
		rel := strings.TrimPrefix(l.path, f.Name)
		rel = strings.TrimPrefix(rel, "/")
		if rel != "" {
			rel += "/"
		}
		out[rel+"f"] = b
		out[rel+"o"] = objects
		if len(l.x) > 0 {
			out[rel+"x"] = l.x
		}
		if l.kind != "" && l.path != f.Name {
			out[rel+"\x01CompObj"] = containerCompObj(l.kind)
		}
		return nil
	}
	if e := write(f.root); e != nil {
		return nil, e
	}
	if f.vbframe != "" {
		b, e := Encode(f.vbframe, f.Codepage)
		if e != nil {
			return nil, e
		}
		out["\x03VBFrame"] = b
	}
	if _, e := f.CFB.Stream(f.Name + "/\x01CompObj"); e != nil {
		b, _ := hex.DecodeString("0100feff030a0000ffffffff00000000000000000000000000000000190000004d6963726f736f667420466f726d7320322e3020466f726d0010000000456d626564646564204f626a6563740000000000f439b271000000000000000000000000")
		out["\x01CompObj"] = b
	}
	return out, nil
}
func (f *Form) ControlNames() []string {
	var out []string
	var walk func(*formLevel)
	walk = func(l *formLevel) {
		for _, c := range l.controls {
			if c.site.strings["Name"] != "" {
				out = append(out, c.site.strings["Name"])
			}
			if c.child != nil {
				walk(c.child)
			}
		}
	}
	walk(f.root)
	sort.Strings(out)
	return out
}
