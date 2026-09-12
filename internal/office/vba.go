package office

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"regexp"
	"strings"
	"unicode"
	"unicode/utf8"
)

type Module struct {
	Name    string `json:"name"`
	Kind    string `json:"kind"`
	Source  string `json:"source"`
	Stream  string `json:"stream,omitempty"`
	Offset  uint32 `json:"-"`
	Records []byte `json:"-"`
}
type VBA struct {
	CFB                       *Compound
	Directory, Prefix, Suffix []byte
	Modules                   []Module
	Name                      string
	Codepage                  int
	ProjectText               string
	References                []map[string]any
}

func Normalize(s string) string {
	return strings.ReplaceAll(strings.ReplaceAll(s, "\r\n", "\n"), "\r", "\n")
}

// Control/procedure names have a different bound from 31-UTF16-unit CFB stream names.
func ValidIdentifier(s string) bool {
	if s == "" || !utf8.ValidString(s) || utf8.RuneCountInString(s) > 255 {
		return false
	}
	for i, c := range s {
		if c == '_' || unicode.IsLetter(c) || (i > 0 && (unicode.IsDigit(c) || unicode.IsMark(c))) {
			continue
		}
		return false
	}
	return true
}
func ReadVBA(b []byte) (v *VBA, err error) {
	defer func() {
		if r := recover(); r != nil {
			v = nil
			err = fmt.Errorf("invalid VBA project: %v", r)
		}
	}()
	c, e := ReadCompound(b)
	if e != nil {
		return nil, e
	}
	d, e := c.Stream("VBA/dir")
	if e != nil {
		return nil, e
	}
	d, e = Decompress(d)
	if e != nil {
		return nil, e
	}
	v = &VBA{CFB: c, Directory: d, Codepage: 1252, Name: "Project"}
	p, count := 0, -1
	for p+6 <= len(d) {
		tag, size := U16(d, p), int(U32(d, p+2))
		if size < 0 || p+6+size > len(d) {
			return nil, fmt.Errorf("truncated dir record")
		}
		if tag == 15 {
			if size != 2 {
				return nil, fmt.Errorf("invalid module count")
			}
			count = int(U16(d, p+6))
			v.Prefix = append([]byte(nil), d[:p]...)
			p += 8
			break
		}
		n := size
		if tag == 9 {
			n = 6
		}
		if p+6+n > len(d) {
			return nil, fmt.Errorf("truncated version")
		}
		if tag == 3 {
			v.Codepage = int(U16(d, p+6))
		}
		p += 6 + n
	}
	if count < 0 || count > 4096 {
		return nil, fmt.Errorf("no bounded VBA module table")
	}
	if p+8 > len(d) || U16(d, p) != 19 || U32(d, p+2) != 2 {
		return nil, fmt.Errorf("no project cookie")
	}
	p += 8
	pb, e := c.Stream("PROJECT")
	if e != nil {
		return nil, e
	}
	v.ProjectText, e = Decode(pb, v.Codepage)
	if e != nil {
		return nil, e
	}
	kinds := map[string]string{}
	for _, line := range strings.Split(Normalize(v.ProjectText), "\n") {
		kv := strings.SplitN(line, "=", 2)
		if len(kv) != 2 {
			continue
		}
		s := strings.Split(kv[1], "/")[0]
		switch kv[0] {
		case "Module":
			kinds[strings.ToLower(s)] = "standard"
		case "Class":
			kinds[strings.ToLower(s)] = "class"
		case "Document":
			kinds[strings.ToLower(s)] = "document"
		case "BaseClass":
			kinds[strings.ToLower(s)] = "form"
		case "Name":
			v.Name = strings.Trim(kv[1], `"`)
		}
	}
	seen := map[string]bool{}
	for i := 0; i < count; i++ {
		start := p
		m := Module{Kind: "class"}
		stream := ""
		hasOffset, finished := false, false
		for p+6 <= len(d) {
			tag, size := U16(d, p), int(U32(d, p+2))
			if p+6+size > len(d) {
				return nil, fmt.Errorf("short module record")
			}
			value := d[p+6 : p+6+size]
			switch tag {
			case 0x19:
				m.Name, e = Decode(value, v.Codepage)
			case 0x47:
				m.Name = utf16str(value)
			case 0x1a:
				stream, e = Decode(value, v.Codepage)
			case 0x32:
				stream = utf16str(value)
			case 0x31:
				if size != 4 {
					return nil, fmt.Errorf("invalid source offset")
				}
				m.Offset = U32(value, 0)
				hasOffset = true
			}
			if e != nil {
				return nil, e
			}
			p += 6 + size
			if tag == 0x2b {
				finished = true
				break
			}
		}
		if m.Name == "" || stream == "" || !hasOffset || !finished {
			return nil, fmt.Errorf("incomplete VBA component")
		}
		key := strings.ToLower(m.Name)
		if seen[key] {
			return nil, fmt.Errorf("duplicate module")
		}
		seen[key] = true
		if k := kinds[key]; k != "" {
			m.Kind = k
		}
		for path := range c.Entries {
			if strings.EqualFold(path, "VBA/"+stream) {
				m.Stream = path
				break
			}
		}
		raw, e := c.Stream(m.Stream)
		if e != nil {
			return nil, e
		}
		if uint64(m.Offset) > uint64(len(raw)) {
			return nil, fmt.Errorf("source offset outside stream")
		}
		raw, e = Decompress(raw[m.Offset:])
		if e != nil {
			return nil, e
		}
		m.Source, e = Decode(raw, v.Codepage)
		if e != nil {
			return nil, e
		}
		m.Source = Normalize(m.Source)
		m.Records = append([]byte(nil), d[start:p]...)
		v.Modules = append(v.Modules, m)
	}
	v.Suffix = append([]byte(nil), d[p:]...)
	return v, nil
}
func NewVBA(name string) *VBA {
	v := &VBA{CFB: NewCompound(), Codepage: 1252, Name: name, Suffix: tlv(0x10, nil)}
	prefix := []byte{}
	add := func(t uint16, b []byte) { prefix = append(prefix, tlv(t, b)...) }
	add(1, dword(1))
	add(2, dword(0x409))
	add(0x14, dword(0x409))
	add(3, word(1252))
	add(4, []byte(name))
	add(5, nil)
	add(0x40, nil)
	add(6, nil)
	add(0x3d, nil)
	add(7, dword(0))
	add(8, dword(0))
	prefix = append(prefix, []byte{9, 0, 4, 0, 0, 0, 1, 0, 0, 0, 0, 0}...)
	add(0x0c, nil)
	add(0x3c, nil)
	v.Prefix = prefix
	v.AddReference("VBA", "{000204EF-0000-0000-C000-000000000046}", "4.2", "", "Visual Basic For Applications")
	v.AddReference("Word", "{00020905-0000-0000-C000-000000000046}", "8.7", "", "Microsoft Word Object Library")
	v.AddReference("Office", "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", "2.8", "", "Microsoft Office Object Library")
	sum := sha256.Sum256([]byte("Wordwright project:" + name))
	id := strings.ToUpper(hex.EncodeToString(sum[:16]))
	v.ProjectText = fmt.Sprintf("ID=\"{%s-%s-%s-%s-%s}\"\r\nName=\"%s\"\r\nHelpContextID=\"0\"\r\nVersionCompatible32=\"393222000\"\r\n\r\n[Host Extender Info]\r\n&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n", id[:8], id[8:12], id[12:16], id[16:20], id[20:], name)
	return v
}
func (v *VBA) AddReference(name, guid, version, path, description string) error {
	if !ValidIdentifier(name) || !regexp.MustCompile(`^\{[A-Fa-f0-9-]{36}\}$`).MatchString(guid) || !regexp.MustCompile(`^[0-9A-Fa-f]+\.[0-9A-Fa-f]+$`).MatchString(version) {
		return fmt.Errorf("invalid type library reference")
	}
	if bytes.Contains(bytes.ToUpper(v.Prefix), []byte(strings.ToUpper(guid)+"#"+strings.ToUpper(version)+"#")) {
		return nil
	}
	libid := fmt.Sprintf("*\\G%s#%s#0#%s#%s", guid, version, path, description)
	raw, e := Encode(libid, v.Codepage)
	if e != nil {
		return e
	}
	nm, e := Encode(name, v.Codepage)
	if e != nil {
		return e
	}
	v.Prefix = append(v.Prefix, tlv(0x16, nm)...)
	v.Prefix = append(v.Prefix, tlv(0x3e, utf16bytes(name))...)
	payload := append(dword(uint32(len(raw))), raw...)
	payload = append(payload, make([]byte, 6)...)
	v.Prefix = append(v.Prefix, tlv(0x0d, payload)...)
	return nil
}
func moduleRecords(m Module, cp int) ([]byte, error) {
	nm, e := Encode(m.Name, cp)
	if e != nil {
		return nil, e
	}
	b := []byte{}
	for _, r := range []struct {
		t uint16
		b []byte
	}{{0x19, nm}, {0x47, utf16bytes(m.Name)}, {0x1a, nm}, {0x32, utf16bytes(m.Name)}, {0x1c, nil}, {0x48, nil}, {0x31, dword(0)}, {0x1e, dword(0)}, {0x2c, word(0)}} {
		b = append(b, tlv(r.t, r.b)...)
	}
	typ := uint16(0x22)
	if m.Kind == "standard" {
		typ = 0x21
	}
	b = append(b, tlv(typ, nil)...)
	b = append(b, tlv(0x2b, nil)...)
	return b, nil
}
func StableGUID(scope string) string {
	sum := sha256.Sum256([]byte("Wordwright:" + scope))
	b := sum[:16]
	b[6] = (b[6] & 15) | 0x50
	b[8] = (b[8] & 63) | 0x80
	h := strings.ToUpper(hex.EncodeToString(b))
	return "{" + h[:8] + "-" + h[8:12] + "-" + h[12:16] + "-" + h[16:20] + "-" + h[20:] + "}"
}
func NewModuleSource(m Module, project string) string {
	s := Normalize(m.Source)
	if !regexp.MustCompile(`(?im)^Attribute VB_Name\s*=`).MatchString(s) {
		s = `Attribute VB_Name = "` + m.Name + "\"\n" + s
	}
	attrs := []string{}
	switch m.Kind {
	case "form":
		attrs = []string{`VB_Base = "0` + StableGUID(project+"/"+m.Name+"/base") + StableGUID(project+"/"+m.Name+"/type") + `"`, "VB_GlobalNameSpace = False", "VB_Creatable = False", "VB_PredeclaredId = True", "VB_Exposed = False", "VB_TemplateDerived = False", "VB_Customizable = False"}
	case "document":
		attrs = []string{`VB_Base = "0{00020906-0000-0000-C000-000000000046}"`, "VB_GlobalNameSpace = False", "VB_Creatable = False", "VB_PredeclaredId = True", "VB_Exposed = True", "VB_TemplateDerived = False", "VB_Customizable = True"}
	case "class":
		attrs = []string{"VB_GlobalNameSpace = False", "VB_Creatable = False", "VB_PredeclaredId = False", "VB_Exposed = False"}
	}
	head := ""
	for _, a := range attrs {
		name := strings.SplitN(a, " =", 2)[0]
		if !regexp.MustCompile(`(?im)^Attribute ` + name + `\s*=`).MatchString(s) {
			head += "Attribute " + a + "\n"
		}
	}
	if head != "" {
		pos := strings.Index(s, "\n")
		if pos < 0 {
			s += "\n"
			pos = len(s) - 1
		}
		s = s[:pos+1] + head + s[pos+1:]
	}
	return s
}

func (v *VBA) Rewrite(modules []Module, forms map[string]map[string][]byte) ([]byte, error) {
	copyV := *v
	copyV.CFB = v.CFB.Clone()
	v = &copyV
	expectedSources := map[string]string{}
	for _, m := range modules {
		if m.Kind == "form" {
			if e := v.AddReference("MSForms", "{0D452EE1-E08F-101A-852E-02608C4D0BB4}", "2.0", "", "Microsoft Forms 2.0 Object Library"); e != nil {
				return nil, e
			}
			break
		}
	}
	if !ValidIdentifier(v.Name) {
		return nil, fmt.Errorf("invalid project identifier")
	}
	if len(modules) > 4096 {
		return nil, fmt.Errorf("too many modules")
	}
	seen := map[string]bool{}
	old := map[string]Module{}
	for _, m := range v.Modules {
		if m.Stream != "" && len(m.Records) != 0 {
			old[strings.ToLower(m.Name)] = m
		}
	}
	b := append([]byte(nil), v.Prefix...)
	// Refresh the project name without reconstructing unrecognized reference data.
	for p := 0; p+6 <= len(b); {
		t, n := U16(b, p), int(U32(b, p+2))
		if t == 9 {
			n = 6
		}
		if p+6+n > len(b) {
			return nil, fmt.Errorf("invalid project prefix")
		}
		if t == 4 {
			nm, e := Encode(v.Name, v.Codepage)
			if e != nil {
				return nil, e
			}
			b = append(append(append([]byte(nil), b[:p]...), tlv(4, nm)...), b[p+6+n:]...)
			break
		}
		p += 6 + n
	}
	b = append(b, tlv(15, word(uint16(len(modules))))...)
	b = append(b, tlv(19, word(0))...)
	wm := []byte{}
	decl := []string{}
	formCount := 0
	for _, m := range modules {
		if !ValidIdentifier(m.Name) {
			return nil, fmt.Errorf("invalid VBA identifier %q", m.Name)
		}
		key := strings.ToLower(m.Name)
		if seen[key] {
			return nil, fmt.Errorf("duplicate component %q", m.Name)
		}
		seen[key] = true
		if !strings.Contains("|standard|class|document|form|", "|"+m.Kind+"|") {
			return nil, fmt.Errorf("invalid module kind")
		}
		prior, existed := old[key]
		if existed && prior.Kind != m.Kind {
			return nil, fmt.Errorf("component type conversion requires a distinct component name")
		}
		src := Normalize(m.Source)
		if !existed {
			src = NewModuleSource(m, v.Name)
		}
		att := regexp.MustCompile(`(?im)^Attribute VB_Name\s*=\s*"([^"]+)"`).FindStringSubmatch(src)
		if len(att) == 0 {
			src = `Attribute VB_Name = "` + m.Name + "\"\n" + src
		} else if !strings.EqualFold(att[1], m.Name) {
			return nil, fmt.Errorf("VB_Name mismatch in %s", m.Name)
		}
		expectedSources[m.Name] = src
		raw, e := Encode(strings.ReplaceAll(src, "\n", "\r\n"), v.Codepage)
		if e != nil {
			return nil, fmt.Errorf("%s: %w", m.Name, e)
		}
		path := "VBA/" + m.Name
		records, e := moduleRecords(m, v.Codepage)
		if e != nil {
			return nil, e
		}
		if existed {
			path = prior.Stream
			records = append([]byte(nil), prior.Records...)
			for p := 0; p+6 <= len(records); {
				t, n := U16(records, p), int(U32(records, p+2))
				if p+6+n > len(records) {
					return nil, fmt.Errorf("bad preserved module record")
				}
				if t == 0x31 {
					put32(records, p+6, 0)
				}
				p += 6 + n
			}
		}
		compressed, e := CompressExact(raw)
		if e != nil {
			return nil, fmt.Errorf("%s: %w", m.Name, e)
		}
		if e = v.CFB.Set(path, compressed); e != nil {
			return nil, e
		}
		b = append(b, records...)
		nm, _ := Encode(m.Name, v.Codepage)
		wm = append(wm, nm...)
		wm = append(wm, 0)
		wm = append(wm, utf16bytes(m.Name)...)
		wm = append(wm, 0, 0)
		typ := map[string]string{"standard": "Module", "class": "Class", "document": "Document", "form": "BaseClass"}[m.Kind]
		name := m.Name
		if m.Kind == "document" {
			name += "/&H00000000"
		}
		decl = append(decl, typ+"="+name)
		if m.Kind == "form" {
			formCount++
		}
	}
	for key, m := range old {
		if !seen[key] {
			if m.Kind == "document" {
				return nil, fmt.Errorf("cannot remove host document module")
			}
			v.CFB.Delete(m.Stream)
			if m.Kind == "form" {
				v.CFB.Delete(m.Name)
			}
		}
	}
	for name, streams := range forms {
		if !seen[strings.ToLower(name)] {
			return nil, fmt.Errorf("form %s has no code component", name)
		}
		for path, data := range streams {
			if e := v.CFB.Set(name+"/"+path, data); e != nil {
				return nil, e
			}
		}
	}
	for _, m := range modules {
		if m.Kind == "form" {
			if _, e := v.CFB.Stream(m.Name + "/f"); e != nil {
				return nil, fmt.Errorf("form %s has no persisted designer; refusing a code-only imitation", m.Name)
			}
		}
	}
	b = append(b, v.Suffix...)
	directory, e := CompressExact(b)
	if e != nil {
		return nil, fmt.Errorf("VBA directory: %w", e)
	}
	if e = v.CFB.Set("VBA/dir", directory); e != nil {
		return nil, e
	}
	_ = v.CFB.Set("VBA/_VBA_PROJECT", []byte{0xcc, 0x61, 0xff, 0xff, 0, 0, 0})
	for p := range v.CFB.Entries {
		if strings.HasPrefix(p, "VBA/__SRP_") {
			v.CFB.Delete(p)
		}
	}
	wm = append(wm, 0, 0)
	_ = v.CFB.Set("PROJECTwm", wm)
	lines := []string{}
	inWorkspace := false
	inserted := false
	for _, s := range strings.Split(Normalize(v.ProjectText), "\n") {
		if s == "[Workspace]" {
			inWorkspace = true
			continue
		}
		if inWorkspace {
			if strings.HasPrefix(s, "[") {
				inWorkspace = false
			} else {
				continue
			}
		}
		if strings.HasPrefix(s, "Module=") || strings.HasPrefix(s, "Class=") || strings.HasPrefix(s, "BaseClass=") || strings.HasPrefix(s, "Document=") || strings.HasPrefix(s, "Package=") {
			continue
		}
		if strings.HasPrefix(s, "Name=") {
			s = `Name="` + v.Name + `"`
		}
		lines = append(lines, s)
		if strings.HasPrefix(s, "ID=") && !inserted {
			lines = append(lines, decl...)
			if formCount > 0 {
				lines = append(lines, "Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}")
			}
			inserted = true
		}
	}
	if !inserted {
		lines = append(decl, lines...)
	}
	pt, e := Encode(strings.Join(lines, "\r\n"), v.Codepage)
	if e != nil {
		return nil, e
	}
	_ = v.CFB.Set("PROJECT", pt)
	out, e := v.CFB.Bytes()
	if e != nil {
		return nil, e
	}
	check, e := ReadVBA(out)
	if e != nil {
		return nil, fmt.Errorf("self-parse of rebuilt VBA failed: %w", e)
	}
	if len(check.Modules) != len(modules) {
		return nil, fmt.Errorf("module roundtrip mismatch")
	}
	for i, m := range modules {
		if check.Modules[i].Source != expectedSources[m.Name] {
			return nil, fmt.Errorf("source roundtrip mismatch: %s", m.Name)
		}
	}
	return out, nil
}
