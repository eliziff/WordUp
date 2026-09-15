package office

// MS-OFORMS data tables. The organization and interoperability details were
// cross-checked against pyOpenVBA (MIT, William Smith); see THIRD_PARTY_NOTICES.
import (
	"fmt"
	"math"
	"strings"
)

type field struct {
	bit  uint
	name string
	size int
	kind byte
}
type extra struct {
	bit    uint
	name   string
	kind   byte
	length string
}
type recordSpec struct {
	name       string
	major      byte
	wide, font bool
	fields     []field
	extra      []extra
	blobs      []field
}
type formRecord struct {
	spec    recordSpec
	mask    uint64
	values  map[string]int64
	strings map[string]string
	extra   map[string][]byte
	tail    []byte
	raw     []byte
	dirty   bool
}

func fields(a ...field) []field { return a }

var textSpec = recordSpec{name: "TextProps", major: 2, fields: fields(field{0, "FontName", 4, 's'}, field{1, "FontEffects", 4, 0}, field{2, "FontHeight", 4, 0}, field{4, "FontCharSet", 1, 0}, field{5, "FontPitchAndFamily", 1, 0}, field{6, "ParagraphAlign", 1, 0}, field{7, "FontWeight", 2, 0}), extra: []extra{{0, "FontName", 's', ""}}}
var formSpec = recordSpec{name: "Form", major: 4, fields: fields(field{1, "BackColor", 4, 0}, field{2, "ForeColor", 4, 0}, field{3, "NextAvailableID", 4, 0}, field{6, "BooleanProperties", 4, 0}, field{7, "BorderStyle", 1, 0}, field{8, "MousePointer", 1, 0}, field{9, "ScrollBars", 1, 0}, field{13, "GroupCnt", 4, 'i'}, field{15, "MouseIcon", 2, 'b'}, field{16, "Cycle", 1, 0}, field{17, "SpecialEffect", 1, 0}, field{18, "BorderColor", 4, 0}, field{19, "Caption", 4, 's'}, field{20, "Font", 2, 'b'}, field{21, "Picture", 2, 'b'}, field{22, "Zoom", 4, 0}, field{23, "PictureAlignment", 1, 0}, field{25, "PictureSizeMode", 1, 0}, field{26, "ShapeCookie", 4, 0}, field{27, "DrawBuffer", 4, 0}), extra: []extra{{10, "DisplayedSize", 'z', ""}, {11, "LogicalSize", 'z', ""}, {12, "ScrollPosition", 'z', ""}, {19, "Caption", 's', ""}}}
var morphSpec = recordSpec{name: "MorphData", major: 2, wide: true, font: true, fields: fields(field{0, "VariousPropertyBits", 4, 0}, field{1, "BackColor", 4, 0}, field{2, "ForeColor", 4, 0}, field{3, "MaxLength", 4, 0}, field{4, "BorderStyle", 1, 0}, field{5, "ScrollBars", 1, 0}, field{6, "DisplayStyle", 1, 0}, field{7, "MousePointer", 1, 0}, field{9, "PasswordChar", 2, 0}, field{10, "ListWidth", 4, 0}, field{11, "BoundColumn", 2, 0}, field{12, "TextColumn", 2, 'i'}, field{13, "ColumnCount", 2, 'i'}, field{14, "ListRows", 2, 0}, field{15, "cColumnInfo", 2, 0}, field{16, "MatchEntry", 1, 0}, field{17, "ListStyle", 1, 0}, field{18, "ShowDropButtonWhen", 1, 0}, field{20, "DropButtonStyle", 1, 0}, field{21, "MultiSelect", 1, 0}, field{22, "Value", 4, 's'}, field{23, "Caption", 4, 's'}, field{24, "PicturePosition", 4, 0}, field{25, "BorderColor", 4, 0}, field{26, "SpecialEffect", 4, 0}, field{27, "MouseIcon", 2, 'b'}, field{28, "Picture", 2, 'b'}, field{29, "Accelerator", 2, 0}, field{32, "GroupName", 4, 's'}), extra: []extra{{8, "Size", 'z', ""}, {22, "Value", 's', ""}, {23, "Caption", 's', ""}, {32, "GroupName", 's', ""}}, blobs: []field{{27, "MouseIcon", 0, 0}, {28, "Picture", 0, 0}}}
var buttonSpec = recordSpec{name: "CommandButton", major: 2, font: true, fields: fields(field{0, "ForeColor", 4, 0}, field{1, "BackColor", 4, 0}, field{2, "VariousPropertyBits", 4, 0}, field{3, "Caption", 4, 's'}, field{4, "PicturePosition", 4, 0}, field{6, "MousePointer", 1, 0}, field{7, "Picture", 2, 'b'}, field{8, "Accelerator", 2, 0}, field{10, "MouseIcon", 2, 'b'}), extra: []extra{{3, "Caption", 's', ""}, {5, "Size", 'z', ""}}, blobs: []field{{7, "Picture", 0, 0}, {10, "MouseIcon", 0, 0}}}
var labelSpec = recordSpec{name: "Label", major: 2, font: true, fields: fields(field{0, "ForeColor", 4, 0}, field{1, "BackColor", 4, 0}, field{2, "VariousPropertyBits", 4, 0}, field{3, "Caption", 4, 's'}, field{4, "PicturePosition", 4, 0}, field{6, "MousePointer", 1, 0}, field{7, "BorderColor", 4, 0}, field{8, "BorderStyle", 2, 0}, field{9, "SpecialEffect", 2, 0}, field{10, "Picture", 2, 'b'}, field{11, "Accelerator", 2, 0}, field{12, "MouseIcon", 2, 'b'}), extra: []extra{{3, "Caption", 's', ""}, {5, "Size", 'z', ""}}, blobs: []field{{10, "Picture", 0, 0}, {12, "MouseIcon", 0, 0}}}
var imageSpec = recordSpec{name: "Image", major: 2, fields: fields(field{3, "BorderColor", 4, 0}, field{4, "BackColor", 4, 0}, field{5, "BorderStyle", 1, 0}, field{6, "MousePointer", 1, 0}, field{7, "PictureSizeMode", 1, 0}, field{8, "SpecialEffect", 1, 0}, field{10, "Picture", 2, 'b'}, field{11, "PictureAlignment", 1, 0}, field{13, "VariousPropertyBits", 4, 0}, field{14, "MouseIcon", 2, 'b'}), extra: []extra{{9, "Size", 'z', ""}}, blobs: []field{{10, "Picture", 0, 0}, {14, "MouseIcon", 0, 0}}}
var spinSpec = recordSpec{name: "SpinButton", major: 2, fields: fields(field{0, "ForeColor", 4, 0}, field{1, "BackColor", 4, 0}, field{2, "VariousPropertyBits", 4, 0}, field{5, "Min", 4, 'i'}, field{6, "Max", 4, 'i'}, field{7, "Position", 4, 'i'}, field{8, "PrevEnabled", 4, 'i'}, field{9, "NextEnabled", 4, 'i'}, field{10, "SmallChange", 4, 'i'}, field{11, "Orientation", 4, 'i'}, field{12, "Delay", 4, 0}, field{13, "MouseIcon", 2, 'b'}, field{14, "MousePointer", 1, 0}), extra: []extra{{3, "Size", 'z', ""}}, blobs: []field{{13, "MouseIcon", 0, 0}}}
var scrollSpec = recordSpec{name: "ScrollBar", major: 2, fields: fields(field{0, "ForeColor", 4, 0}, field{1, "BackColor", 4, 0}, field{2, "VariousPropertyBits", 4, 0}, field{4, "MousePointer", 1, 0}, field{5, "Min", 4, 'i'}, field{6, "Max", 4, 'i'}, field{7, "Position", 4, 'i'}, field{9, "PrevEnabled", 4, 'i'}, field{10, "NextEnabled", 4, 'i'}, field{11, "SmallChange", 4, 'i'}, field{12, "LargeChange", 4, 'i'}, field{13, "Orientation", 4, 'i'}, field{14, "ProportionalThumb", 2, 'i'}, field{15, "Delay", 4, 0}, field{16, "MouseIcon", 2, 'b'}), extra: []extra{{3, "Size", 'z', ""}}, blobs: []field{{16, "MouseIcon", 0, 0}}}
var tabSpec = recordSpec{name: "TabStrip", major: 2, font: true, fields: fields(field{0, "ListIndex", 4, 'i'}, field{1, "BackColor", 4, 0}, field{2, "ForeColor", 4, 0}, field{5, "ItemsSize", 4, 0}, field{6, "MousePointer", 1, 0}, field{8, "TabOrientation", 4, 0}, field{9, "TabStyle", 4, 0}, field{11, "TabFixedWidth", 4, 0}, field{12, "TabFixedHeight", 4, 0}, field{15, "TipStringsSize", 4, 0}, field{17, "NamesSize", 4, 0}, field{18, "VariousPropertyBits", 4, 0}, field{20, "TabsAllocated", 4, 0}, field{21, "TagsSize", 4, 0}, field{22, "TabData", 4, 0}, field{23, "AcceleratorsSize", 4, 0}, field{24, "MouseIcon", 2, 'b'}), extra: []extra{{4, "Size", 'z', ""}, {5, "Items", 'a', "ItemsSize"}, {15, "TipStrings", 'a', "TipStringsSize"}, {17, "TabNames", 'a', "NamesSize"}, {21, "Tags", 'a', "TagsSize"}, {23, "Accelerators", 'a', "AcceleratorsSize"}}, blobs: []field{{24, "MouseIcon", 0, 0}}}
var kindIndex = map[string]int{"Form": 7, "Image": 12, "Frame": 14, "MorphData": 15, "SpinButton": 16, "CommandButton": 17, "TabStrip": 18, "Label": 21, "TextBox": 23, "ListBox": 24, "ComboBox": 25, "CheckBox": 26, "OptionButton": 27, "ToggleButton": 28, "ScrollBar": 47, "MultiPage": 57}

func specFor(index int) (recordSpec, bool) {
	switch index {
	case 7, 14, 57:
		return formSpec, true
	case 12:
		return imageSpec, true
	case 15, 23, 24, 25, 26, 27, 28:
		return morphSpec, true
	case 16:
		return spinSpec, true
	case 17:
		return buttonSpec, true
	case 18:
		return tabSpec, true
	case 21:
		return labelSpec, true
	case 47:
		return scrollSpec, true
	}
	return recordSpec{}, false
}
func newRecord(spec recordSpec) *formRecord {
	return &formRecord{spec: spec, values: map[string]int64{}, strings: map[string]string{}, extra: map[string][]byte{}, dirty: true}
}
func readRecord(data []byte, spec recordSpec, cp int) (r *formRecord, err error) {
	defer func() {
		if p := recover(); p != nil {
			r = nil
			err = fmt.Errorf("invalid %s record: %v", spec.name, p)
		}
	}()
	if len(data) < 8 || data[0] != 0 || data[1] != spec.major {
		return nil, fmt.Errorf("unexpected %s record version", spec.name)
	}
	bound := 4 + int(U16(data, 2))
	if bound > len(data) {
		return nil, fmt.Errorf("short %s record", spec.name)
	}
	r = newRecord(spec)
	r.mask = uint64(U32(data, 4))
	p := 8
	if spec.wide {
		r.mask |= uint64(U32(data, 8)) << 32
		p = 12
	}
	for _, f := range spec.fields {
		if r.mask&(1<<f.bit) == 0 {
			continue
		}
		p += (f.size - p%f.size) % f.size
		if p+f.size > bound {
			return nil, fmt.Errorf("field outside record")
		}
		var n int64
		switch f.size {
		case 1:
			n = int64(data[p])
		case 2:
			n = int64(U16(data, p))
		case 4:
			n = int64(U32(data, p))
		}
		if f.kind == 'i' {
			switch f.size {
			case 2:
				n = int64(int16(n))
			case 4:
				n = int64(int32(n))
			}
		}
		r.values[f.name] = n
		p += f.size
	}
	p += (4 - p%4) % 4
	for _, ex := range spec.extra {
		if r.mask&(1<<ex.bit) == 0 {
			continue
		}
		switch ex.kind {
		case 'z':
			if p+8 > bound {
				return nil, fmt.Errorf("size outside record")
			}
			r.extra[ex.name] = append([]byte(nil), data[p:p+8]...)
			p += 8
		case 's':
			n := uint32(r.values[ex.name])
			size := int(n & 0x7fffffff)
			if p+size > bound {
				return nil, fmt.Errorf("string outside record")
			}
			var s string
			if n&0x80000000 != 0 {
				s, err = Decode(data[p:p+size], cp)
			} else {
				if size%2 != 0 {
					return nil, fmt.Errorf("odd unicode string")
				}
				s = utf16str(data[p : p+size])
			}
			if err != nil {
				return nil, err
			}
			r.strings[ex.name] = s
			p += size
			p += (4 - p%4) % 4
		case 'a':
			size := int(r.values[ex.length])
			if size < 0 || p+size > bound {
				return nil, fmt.Errorf("array outside record")
			}
			r.extra[ex.name] = append([]byte(nil), data[p:p+size]...)
			p += size
			p += (4 - p%4) % 4
		}
	}
	if p > bound {
		return nil, fmt.Errorf("record overflow")
	}
	if p < bound {
		r.extra["_unparsed"] = append([]byte(nil), data[p:bound]...)
	}
	r.tail = append([]byte(nil), data[bound:]...)
	r.raw = append([]byte(nil), data...)
	r.dirty = false
	return r, nil
}
func (r *formRecord) set(name string, value any) error {
	for _, f := range r.spec.fields {
		if f.name != name {
			continue
		}
		if value == nil {
			r.mask &^= 1 << f.bit
			delete(r.values, name)
			delete(r.strings, name)
			r.dirty = true
			return nil
		}
		if f.kind == 's' {
			s, ok := value.(string)
			if !ok {
				return fmt.Errorf("%s needs string", name)
			}
			if r.strings[name] == s && r.mask&(1<<f.bit) != 0 {
				return nil
			}
			r.strings[name] = s
		} else if f.kind == 'b' {
			return fmt.Errorf("binary asset %s requires asset API", name)
		} else {
			n, err := number(value)
			if err != nil {
				return fmt.Errorf("%s: %w", name, err)
			}
			lower, upper := int64(0), int64(1)<<(f.size*8)
			if f.kind == 'i' {
				lower = -(upper / 2)
				upper /= 2
			}
			if n < lower || n >= upper {
				return fmt.Errorf("%s value %d exceeds %d-bit field", name, n, f.size*8)
			}
			if r.values[name] == n && r.mask&(1<<f.bit) != 0 {
				return nil
			}
			r.values[name] = n
		}
		r.mask |= 1 << f.bit
		r.dirty = true
		return nil
	}
	return fmt.Errorf("property %s is not a stored %s field", name, r.spec.name)
}
func number(v any) (int64, error) {
	switch n := v.(type) {
	case float64:
		if math.IsNaN(n) || math.IsInf(n, 0) || math.Trunc(n) != n || n > math.MaxInt32*2+1 || n < math.MinInt32 {
			return 0, fmt.Errorf("number outside 32-bit property range")
		}
		return int64(n), nil
	case int:
		return int64(n), nil
	case int64:
		return n, nil
	case uint32:
		return int64(n), nil
	case bool:
		if n {
			return 1, nil
		}
		return 0, nil
	}
	return 0, fmt.Errorf("expected numeric property, received %T", v)
}
func (r *formRecord) size(name string, w, h float64) error {
	for _, e := range r.spec.extra {
		if e.name == name && e.kind == 'z' {
			width := math.Round(w * 2540 / 72)
			height := math.Round(h * 2540 / 72)
			if math.IsNaN(width) || math.IsInf(width, 0) || math.IsNaN(height) || math.IsInf(height, 0) || width < math.MinInt32 || width > math.MaxInt32 || height < math.MinInt32 || height > math.MaxInt32 {
				return fmt.Errorf("size outside 32-bit form geometry range")
			}
			if (name == "DisplayedSize" || name == "Size") && (width <= 0 || height <= 0) {
				return fmt.Errorf("form size must be positive")
			}
			b := append(dword(uint32(int32(width))), dword(uint32(int32(height)))...)
			if string(r.extra[name]) == string(b) && r.mask&(1<<e.bit) != 0 {
				return nil
			}
			r.extra[name] = b
			r.mask |= 1 << e.bit
			r.dirty = true
			return nil
		}
	}
	return fmt.Errorf("size not stored in %s", r.spec.name)
}
func (r *formRecord) dimensions(name string) (float64, float64) {
	b := r.extra[name]
	if len(b) != 8 {
		return 0, 0
	}
	return float64(int32(U32(b, 0))) * 72 / 2540, float64(int32(U32(b, 4))) * 72 / 2540
}
func (r *formRecord) bytes(cp int) ([]byte, error) {
	if !r.dirty && r.raw != nil {
		return append([]byte(nil), r.raw...), nil
	}
	b := []byte{0, r.spec.major, 0, 0}
	b = append(b, dword(uint32(r.mask))...)
	if r.spec.wide {
		b = append(b, dword(uint32(r.mask>>32))...)
	}
	for _, f := range r.spec.fields {
		if r.mask&(1<<f.bit) == 0 {
			continue
		}
		b = pad(b, f.size)
		n := r.values[f.name]
		if f.kind == 's' {
			_, size := formString(r.strings[f.name])
			n = int64(size)
		}
		switch f.size {
		case 1:
			b = append(b, byte(n))
		case 2:
			b = append(b, word(uint16(n))...)
		case 4:
			b = append(b, dword(uint32(n))...)
		}
	}
	b = pad(b, 4)
	for _, e := range r.spec.extra {
		if r.mask&(1<<e.bit) == 0 {
			continue
		}
		switch e.kind {
		case 's':
			raw, _ := formString(r.strings[e.name])
			b = append(b, raw...)
			b = pad(b, 4)
		case 'z':
			x := r.extra[e.name]
			if len(x) != 8 {
				return nil, fmt.Errorf("missing geometry")
			}
			b = append(b, x...)
		case 'a':
			b = append(b, r.extra[e.name]...)
			b = pad(b, 4)
		}
	}
	b = append(b, r.extra["_unparsed"]...)
	if len(b)-4 > 65535 {
		return nil, fmt.Errorf("form record size overflow")
	}
	put16(b, 2, uint16(len(b)-4))
	b = append(b, r.tail...)
	return b, nil
}
func defaultControl(kind, name string) (*formRecord, error) {
	idx, ok := kindIndex[kind]
	if !ok {
		return nil, fmt.Errorf("unknown MSForms kind %s", kind)
	}
	spec, ok := specFor(idx)
	if !ok {
		return nil, fmt.Errorf("no record codec for %s", kind)
	}
	r := newRecord(spec)
	if kind == "Frame" || kind == "Form" || kind == "MultiPage" {
		_ = r.size("DisplayedSize", 180, 120)
		_ = r.size("LogicalSize", 0, 0)
		_ = r.set("NextAvailableID", 1)
		_ = r.set("ShapeCookie", 0)
		_ = r.set("BooleanProperties", 0x8004)
		if kind == "MultiPage" {
			_ = r.set("BooleanProperties", 0xc004)
		}
		_ = r.set("DrawBuffer", 32000)
		if kind == "Frame" {
			_ = r.set("Caption", name)
		}
		return r, nil
	}
	_ = r.size("Size", 72, 24)
	if strings.Contains("|CommandButton|Label|ToggleButton|CheckBox|OptionButton|", "|"+kind+"|") {
		_ = r.set("Caption", name)
	}
	if spec.wide {
		r.mask |= 1 << 31
		style := map[string]int{"TextBox": 1, "ListBox": 2, "ComboBox": 3, "CheckBox": 4, "OptionButton": 5, "ToggleButton": 6}[kind]
		if style > 1 {
			_ = r.set("DisplayStyle", style)
		}
		switch kind {
		case "TextBox", "ComboBox":
			_ = r.set("VariousPropertyBits", int64(0x2c80481b))
			if kind == "ComboBox" {
				_ = r.set("MatchEntry", 1)
				_ = r.set("ShowDropButtonWhen", 2)
			}
		case "ListBox":
			_ = r.set("ScrollBars", 3)
			_ = r.set("MatchEntry", 0)
		case "CheckBox", "OptionButton", "ToggleButton":
			_ = r.set("BackColor", int64(0x8000000f))
			_ = r.set("ForeColor", int64(0x80000012))
			_ = r.set("Value", "0")
		}
	}
	if kind == "SpinButton" || kind == "ScrollBar" {
		_ = r.set("Orientation", -1)
	}
	if spec.font {
		f := newRecord(textSpec)
		_ = f.set("FontName", "Tahoma")
		_ = f.set("FontHeight", 165)
		_ = f.set("FontCharSet", 0)
		_ = f.set("FontPitchAndFamily", 2)
		if kind == "CommandButton" || kind == "ToggleButton" {
			_ = f.set("ParagraphAlign", 3)
		}
		r.tail, _ = f.bytes(1252)
	}
	return r, nil
}
