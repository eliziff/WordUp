package office

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/xml"
	"fmt"
	"io"
	"net/url"
	"os"
	"path"
	"sort"
	"strings"
	"time"
)

const (
	W             = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	R             = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	RelNS         = "http://schemas.openxmlformats.org/package/2006/relationships"
	CT            = "http://schemas.openxmlformats.org/package/2006/content-types"
	MainDOTM      = "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"
	MainDOCX      = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
	VBAProjectRel = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
)

// SetVBA attaches the native project using Word's extension relationship, which
// is not in the standard OOXML relationship namespace. Retain existing part IDs.
func (p *Package) SetVBA(data []byte) error {
	p.Files["word/vbaProject.bin"] = data
	for part, typ := range map[string]string{"word/document.xml": MainDOTM, "word/vbaProject.bin": "application/vnd.ms-office.vbaProject"} {
		if err := p.ContentType(part, typ); err != nil {
			return err
		}
	}
	id := "rIdVBA"
	if b := p.Files[RelPart("word/document.xml")]; len(b) > 0 {
		spans, err := XMLSpans(b)
		if err != nil {
			return err
		}
		for _, s := range spans {
			if s.Name.Local == "Relationship" && (s.Attribute("", "Type") == VBAProjectRel || s.Attribute("", "Type") == R+"/vbaProject") {
				id = s.Attribute("", "Id")
				break
			}
		}
	}
	if err := p.Relationship("word/document.xml", id, VBAProjectRel, "vbaProject.bin", ""); err != nil {
		return err
	}
	const supplemental = "http://schemas.microsoft.com/office/2006/relationships/wordVbaData"
	if b := p.Files[RelPart("word/vbaProject.bin")]; len(b) > 0 {
		spans, err := XMLSpans(b)
		if err != nil {
			return err
		}
		for _, s := range spans {
			if s.Attribute("", "Type") == supplemental {
				return nil
			}
		}
	}
	const part = "word/vbaData.xml"
	if len(p.Files[part]) == 0 {
		p.Files[part] = []byte(`<wne:vbaSuppData xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"/>`)
	}
	if err := p.ContentType(part, "application/vnd.ms-word.vbaData+xml"); err != nil {
		return err
	}
	return p.Relationship("word/vbaProject.bin", "rIdData", supplemental, "vbaData.xml", "")
}

type Package struct {
	Original []byte
	Files    map[string][]byte
	archive  *zip.Reader
	// archiveLogical maps the original ZIP member spelling to the canonical
	// slash-separated key used by Files. Keep the original spelling when an
	// unrelated part changes so untouched members remain byte-preserved.
	archiveLogical map[string]string
	hashes         map[string]string
}

// WithFiles reuses immutable ZIP metadata with a caller-owned part map.
func (p *Package) WithFiles(files map[string][]byte) *Package {
	return &Package{Original: p.Original, Files: files, archive: p.archive, archiveLogical: p.archiveLogical, hashes: p.hashes}
}

func Hash(b []byte) string { s := sha256.Sum256(b); return hex.EncodeToString(s[:]) }
func SafePart(name string) bool {
	if name == "" || strings.ContainsAny(name, "\\\x00:") || strings.HasPrefix(name, "/") || path.Clean(name) != name {
		return false
	}
	for _, s := range strings.Split(name, "/") {
		if s == "." || s == ".." || s == "" || strings.TrimRight(s, ". ") != s {
			return false
		}
		base := strings.ToUpper(strings.SplitN(s, ".", 2)[0])
		if base == "CON" || base == "PRN" || base == "AUX" || base == "NUL" || len(base) == 4 && (strings.HasPrefix(base, "COM") || strings.HasPrefix(base, "LPT")) && base[3] >= '1' && base[3] <= '9' {
			return false
		}
	}
	return true
}
func ReadPackage(data []byte) (*Package, error) {
	if len(data) > Limit {
		return nil, fmt.Errorf("package exceeds %d byte limit", Limit)
	}
	z, e := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if e != nil {
		return nil, e
	}
	if len(z.File) > 65536 {
		return nil, fmt.Errorf("too many package entries")
	}
	p := &Package{Original: data, Files: map[string][]byte{}, archive: z, archiveLogical: map[string]string{}, hashes: map[string]string{}}
	names := map[string]bool{}
	var total uint64
	for _, f := range z.File {
		if f.FileInfo().IsDir() {
			continue
		}
		originalName := f.Name
		name := strings.ReplaceAll(originalName, "\\", "/")
		if !SafePart(name) || f.Mode()&os.ModeSymlink != 0 {
			return nil, fmt.Errorf("unsafe package part %q", originalName)
		}
		key := strings.ToLower(name)
		if names[key] {
			return nil, fmt.Errorf("case-colliding or duplicate ZIP part")
		}
		names[key] = true
		total += f.UncompressedSize64
		if total > Limit {
			return nil, fmt.Errorf("expanded package budget exceeded")
		}
		r, e := f.Open()
		if e != nil {
			return nil, e
		}
		b, e := io.ReadAll(io.LimitReader(r, int64(Limit)+1))
		r.Close()
		if e != nil {
			return nil, e
		}
		if len(b) > Limit || uint64(len(b)) != f.UncompressedSize64 {
			return nil, fmt.Errorf("invalid ZIP entry size")
		}
		p.Files[name] = b
		p.archiveLogical[originalName] = name
		p.hashes[name] = Hash(b)
	}
	if _, ok := p.Files["[Content_Types].xml"]; !ok {
		return nil, fmt.Errorf("not an Open Packaging Conventions file")
	}
	return p, nil
}
func (p *Package) Bytes() ([]byte, error) {
	return p.bytes(nil)
}

// BytesChanged skips re-reading unchanged ZIP members after the caller has
// already established the complete changed-part set by content hash.
func (p *Package) BytesChanged(changed []string) ([]byte, error) {
	if p.hashes == nil {
		return p.Bytes()
	}
	dirty := make(map[string]bool, len(changed))
	for _, name := range changed {
		dirty[strings.TrimPrefix(name, "-")] = true
	}
	for name, original := range p.hashes {
		data, exists := p.Files[name]
		if !exists {
			if !dirty[name] {
				return nil, fmt.Errorf("changed-part set omits deleted part %s", name)
			}
			continue
		}
		if !dirty[name] && Hash(data) != original {
			return nil, fmt.Errorf("changed-part set omits modified part %s", name)
		}
	}
	for name := range p.Files {
		if _, existed := p.hashes[name]; !existed && !dirty[name] {
			return nil, fmt.Errorf("changed-part set omits new part %s", name)
		}
	}
	return p.bytes(dirty)
}

func (p *Package) bytes(knownChanges map[string]bool) ([]byte, error) {
	logicalName := func(f *zip.File) string {
		if p.archiveLogical != nil {
			if name, ok := p.archiveLogical[f.Name]; ok {
				return name
			}
		}
		return f.Name
	}
	if p.archive != nil && len(p.Files) == len(p.archive.File) {
		if knownChanges != nil && len(knownChanges) == 0 {
			return append([]byte(nil), p.Original...), nil
		}
		if knownChanges == nil {
			equal := true
			for _, f := range p.archive.File {
				b, ok := p.Files[logicalName(f)]
				if !ok {
					equal = false
					break
				}
				r, e := f.Open()
				if e != nil {
					return nil, e
				}
				orig, e := io.ReadAll(r)
				r.Close()
				if e != nil {
					return nil, e
				}
				if !bytes.Equal(b, orig) {
					equal = false
					break
				}
			}
			if equal {
				return append([]byte(nil), p.Original...), nil
			}
		}
	}
	var out bytes.Buffer
	z := zip.NewWriter(&out)
	done := map[string]bool{}
	if p.archive != nil {
		for _, f := range p.archive.File {
			name := logicalName(f)
			b, ok := p.Files[name]
			if !ok {
				continue
			}
			if knownChanges != nil && !knownChanges[name] {
				if e := z.Copy(f); e != nil {
					return nil, e
				}
				done[name] = true
				continue
			}
			r, e := f.Open()
			if e != nil {
				return nil, e
			}
			orig, e := io.ReadAll(r)
			r.Close()
			if e != nil {
				return nil, e
			}
			if bytes.Equal(orig, b) {
				if e = z.Copy(f); e != nil {
					return nil, e
				}
				done[name] = true
			}
		}
	}
	keys := make([]string, 0, len(p.Files))
	for name := range p.Files {
		if !done[name] {
			keys = append(keys, name)
		}
	}
	sort.Strings(keys)
	for _, name := range keys {
		if !SafePart(name) {
			return nil, fmt.Errorf("invalid package output name")
		}
		header := &zip.FileHeader{Name: name, Method: zip.Deflate}
		header.SetModTime(time.Date(1980, 1, 1, 0, 0, 0, 0, time.UTC))
		w, e := z.CreateHeader(header)
		if e != nil {
			return nil, e
		}
		if _, e = w.Write(p.Files[name]); e != nil {
			return nil, e
		}
	}
	if e := z.Close(); e != nil {
		return nil, e
	}
	if out.Len() > Limit {
		return nil, fmt.Errorf("output package budget exceeded")
	}
	return out.Bytes(), nil
}
func Esc(s string) string { var b bytes.Buffer; xml.EscapeText(&b, []byte(s)); return b.String() }

// XMLSpan provides edit boundaries without round-tripping namespace prefixes,
// compatibility declarations, unknown extensions, whitespace, or opaque XML.
type XMLSpan struct {
	Name                                   xml.Name
	Attr                                   []xml.Attr
	Start, OpenEnd, CloseStart, End, Depth int
}

func (s XMLSpan) Attribute(ns, local string) string {
	for _, a := range s.Attr {
		if a.Name.Local == local && (ns == "*" || a.Name.Space == ns) {
			return a.Value
		}
	}
	return ""
}
func XMLSpans(b []byte) ([]XMLSpan, error) {
	if len(b) > Limit {
		return nil, fmt.Errorf("XML budget exceeded")
	}
	d := xml.NewDecoder(bytes.NewReader(b))
	d.Strict = true
	spans := []XMLSpan{}
	stack := []int{}
	roots := 0
	for {
		start := int(d.InputOffset())
		t, e := d.Token()
		if e == io.EOF {
			break
		}
		if e != nil {
			return nil, e
		}
		end := int(d.InputOffset())
		switch v := t.(type) {
		case xml.Directive:
			return nil, fmt.Errorf("XML directives/DTD are not accepted")
		case xml.StartElement:
			if len(stack) > 256 {
				return nil, fmt.Errorf("XML nesting budget exceeded")
			}
			if len(stack) == 0 {
				roots++
			}
			if roots > 1 {
				return nil, fmt.Errorf("multiple XML roots")
			}
			spans = append(spans, XMLSpan{Name: v.Name, Attr: v.Attr, Start: start, OpenEnd: end, Depth: len(stack)})
			stack = append(stack, len(spans)-1)
		case xml.EndElement:
			if len(stack) == 0 {
				return nil, fmt.Errorf("XML stack underflow")
			}
			i := stack[len(stack)-1]
			stack = stack[:len(stack)-1]
			spans[i].CloseStart = start
			spans[i].End = end
		case xml.CharData:
			if start == 0 && bytes.HasPrefix(v, []byte{0xef, 0xbb, 0xbf}) {
				v = v[3:]
			}
			if len(stack) == 0 && strings.TrimSpace(string(v)) != "" {
				return nil, fmt.Errorf("text outside XML root")
			}
		}
	}
	if roots != 1 || len(stack) != 0 {
		return nil, fmt.Errorf("incomplete XML document")
	}
	return spans, nil
}
func EnsureNamespace(b []byte, prefix, uri string) ([]byte, error) {
	spans, e := XMLSpans(b)
	if e != nil {
		return nil, e
	}
	root := spans[0]
	for _, a := range root.Attr {
		if a.Name.Space == "xmlns" && a.Name.Local == prefix {
			if a.Value != uri {
				return nil, fmt.Errorf("namespace prefix collision: %s", prefix)
			}
			return b, nil
		}
	}
	i := root.OpenEnd - 1
	if i > 0 && b[i-1] == '/' {
		i--
	}
	return bytes.Join([][]byte{b[:i], []byte(` xmlns:` + prefix + `="` + Esc(uri) + `"`), b[i:]}, nil), nil
}
func InsertXML(b []byte, fragment string) ([]byte, error) {
	spans, e := XMLSpans(b)
	if e != nil {
		return nil, e
	}
	r := spans[0]
	if r.OpenEnd >= 2 && bytes.Equal(b[r.OpenEnd-2:r.OpenEnd], []byte("/>")) {
		// Preserve the root QName verbatim when expanding a self-closing root.
		nameStart := r.Start + 1
		nameEnd := nameStart
		for nameEnd < len(b) && !strings.ContainsRune(" \t\r\n/>", rune(b[nameEnd])) {
			nameEnd++
		}
		return []byte(string(b[:r.OpenEnd-2]) + ">" + fragment + "</" + string(b[nameStart:nameEnd]) + ">" + string(b[r.End:])), nil
	}
	return bytes.Join([][]byte{b[:r.CloseStart], []byte(fragment), b[r.CloseStart:]}, nil), nil
}
func UpsertXML(b []byte, ns, local, attrNS, attr, value, fragment string) ([]byte, error) {
	spans, e := XMLSpans(b)
	if e != nil {
		return nil, e
	}
	for _, s := range spans {
		if s.Depth == 1 && s.Name.Space == ns && s.Name.Local == local && s.Attribute(attrNS, attr) == value {
			return bytes.Join([][]byte{b[:s.Start], []byte(fragment), b[s.End:]}, nil), nil
		}
	}
	return InsertXML(b, fragment)
}
func (p *Package) ContentType(part, kind string) error {
	b := p.Files["[Content_Types].xml"]
	if len(b) == 0 {
		b = []byte(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="` + CT + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`)
	}
	spans, err := XMLSpans(b)
	if err != nil {
		return err
	}
	existing, overridden := "", false
	for _, s := range spans {
		if s.Depth != 1 || s.Name.Space != CT {
			continue
		}
		if s.Name.Local == "Override" && s.Attribute("", "PartName") == "/"+part {
			existing, overridden = s.Attribute("", "ContentType"), true
			break
		}
	}
	if !overridden {
		for _, s := range spans {
			if s.Depth == 1 && s.Name.Space == CT && s.Name.Local == "Default" && s.Attribute("", "Extension") == strings.TrimPrefix(path.Ext(part), ".") {
				existing = s.Attribute("", "ContentType")
				break
			}
		}
	}
	if existing == kind && kind != "" {
		if len(p.Files["[Content_Types].xml"]) == 0 {
			p.Files["[Content_Types].xml"] = b
		}
		return nil
	}
	b, e := UpsertXML(b, CT, "Override", "", "PartName", "/"+part, `<Override xmlns="`+CT+`" PartName="/`+Esc(part)+`" ContentType="`+Esc(kind)+`"/>`)
	if e == nil {
		p.Files["[Content_Types].xml"] = b
	}
	return e
}
func RelPart(source string) string {
	if source == "" {
		return "_rels/.rels"
	}
	return path.Join(path.Dir(source), "_rels", path.Base(source)+".rels")
}
func (p *Package) Relationship(source, id, typ, target, mode string) error {
	rel := RelPart(source)
	b := p.Files[rel]
	if len(b) == 0 {
		b = []byte(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="` + RelNS + `"/>`)
	}
	spans, err := XMLSpans(b)
	if err != nil {
		return err
	}
	for _, s := range spans {
		if s.Depth == 1 && s.Name.Space == RelNS && s.Name.Local == "Relationship" && s.Attribute("", "Id") == id {
			if s.Attribute("", "Type") == typ && s.Attribute("", "Target") == target && s.Attribute("", "TargetMode") == mode {
				return nil
			}
			return fmt.Errorf("relationship ID collision in %s: %s targets %s, requested %s", rel, id, s.Attribute("", "Target"), target)
		}
	}
	m := ""
	if mode != "" {
		m = ` TargetMode="` + Esc(mode) + `"`
	}
	b, e := UpsertXML(b, RelNS, "Relationship", "", "Id", id, `<Relationship xmlns="`+RelNS+`" Id="`+Esc(id)+`" Type="`+Esc(typ)+`" Target="`+Esc(target)+`"`+m+`/>`)
	if e == nil {
		p.Files[rel] = b
	}
	return e
}
func (p *Package) HasSignatures() bool {
	for k := range p.Files {
		l := strings.ToLower(k)
		if strings.Contains(l, "vbaprojectsignature") || strings.HasPrefix(l, "_xmlsignatures/") {
			return true
		}
	}
	return false
}
func (p *Package) Validate() error {
	defaults, overrides := map[string]string{}, map[string]string{}
	ct, e := validationXMLSpans(p.Files["[Content_Types].xml"])
	if e != nil {
		return fmt.Errorf("content types: %w", e)
	}
	for _, s := range ct {
		if s.Depth == 1 {
			switch s.Name.Local {
			case "Default":
				defaults[s.Attribute("", "Extension")] = s.Attribute("", "ContentType")
			case "Override":
				overrides[strings.TrimPrefix(s.Attribute("", "PartName"), "/")] = s.Attribute("", "ContentType")
			}
		}
	}
	for name, b := range p.Files {
		if !SafePart(name) {
			return fmt.Errorf("invalid part path")
		}
		if name != "[Content_Types].xml" && overrides[name] == "" && defaults[strings.TrimPrefix(path.Ext(name), ".")] == "" {
			return fmt.Errorf("no content type for %s", name)
		}
		if strings.HasSuffix(name, ".xml") || strings.HasSuffix(name, ".rels") {
			nodes, e := validationXMLSpans(b)
			if e != nil {
				return fmt.Errorf("%s: %w", name, e)
			}
			if !strings.HasSuffix(name, ".rels") {
				continue
			}
			base := ""
			if name != "_rels/.rels" {
				if !strings.Contains(name, "/_rels/") {
					return fmt.Errorf("invalid relationship part")
				}
				base = path.Dir(path.Dir(name))
			}
			ids := map[string]bool{}
			for _, s := range nodes {
				if s.Name.Space != RelNS || s.Name.Local != "Relationship" {
					continue
				}
				id := s.Attribute("", "Id")
				if id == "" || ids[id] {
					return fmt.Errorf("duplicate/missing relationship ID in %s", name)
				}
				ids[id] = true
				target := s.Attribute("", "Target")
				if s.Attribute("", "TargetMode") == "External" {
					continue
				}
				u, e := url.Parse(target)
				if e != nil || u.Scheme != "" || u.Host != "" {
					return fmt.Errorf("invalid internal relationship %s", target)
				}
				decoded, e := url.PathUnescape(u.EscapedPath())
				if e != nil {
					return e
				}
				resolved := path.Join(base, decoded)
				if strings.HasPrefix(decoded, "/") {
					resolved = strings.TrimPrefix(path.Clean(decoded), "/")
				}
				if !SafePart(resolved) {
					return fmt.Errorf("unsafe relationship %q", target)
				}
				if _, ok := p.Files[resolved]; !ok {
					return fmt.Errorf("missing relationship target: %s -> %s", name, resolved)
				}
			}
		}
	}
	if p.Files["word/document.xml"] == nil {
		return fmt.Errorf("this version requires the Word main part at word/document.xml")
	}
	return nil
}
func BlankPackage() *Package {
	p := &Package{Files: map[string][]byte{}}
	p.Files["word/document.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="` + W + `" xmlns:r="` + R + `"><w:body><w:p/><w:sectPr/></w:body></w:document>`)
	p.Files["word/styles.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="` + W + `"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`)
	p.ContentType("word/document.xml", MainDOCX)
	p.ContentType("word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")
	p.Relationship("", "rIdDocument", R+"/officeDocument", "word/document.xml", "")
	p.Relationship("word/document.xml", "rIdStyles", R+"/styles", "styles.xml", "")
	return p
}
