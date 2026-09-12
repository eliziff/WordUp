package office

import (
	"bytes"
	"encoding/hex"
	"fmt"
	"strings"
)

// Explicit MS-OFORMS container metadata; these are format identifiers, not
// injected executable stubs. Interoperability constants: see THIRD_PARTY_NOTICES.
var containerIDs = map[string]string{
	"Frame":     "2020186e60f4ce119bcd00aa00608e01",
	"Form":      "f0692ac6dc16ce119e9800aa00574a4f",
	"MultiPage": "7013e3467a3fce11bed600aa00611080",
}

func containerCompObj(kind string) []byte {
	id, _ := hex.DecodeString(containerIDs[kind])
	b, _ := hex.DecodeString("0100feff030a0000ffffffff")
	b = append(b, id...)
	label := "Microsoft Forms 2.0 Form"
	if kind == "Frame" {
		label = "Microsoft Forms 2.0 Frame"
	}
	for _, s := range []string{label, "Embedded Object", "Forms." + kind + ".1"} {
		raw := append([]byte(s), 0)
		b = append(b, dword(uint32(len(raw)))...)
		b = append(b, raw...)
	}
	tail, _ := hex.DecodeString("f439b271000000000000000000000000")
	return append(b, tail...)
}
func cloneRecord(r *formRecord) *formRecord {
	if r == nil {
		return nil
	}
	q := *r
	q.values = map[string]int64{}
	for k, v := range r.values {
		q.values[k] = v
	}
	q.strings = map[string]string{}
	for k, v := range r.strings {
		q.strings[k] = v
	}
	q.extra = map[string][]byte{}
	for k, v := range r.extra {
		q.extra[k] = append([]byte(nil), v...)
	}
	q.tail = append([]byte(nil), r.tail...)
	q.raw = append([]byte(nil), r.raw...)
	return &q
}
func cloneLevel(l *formLevel) *formLevel {
	if l == nil {
		return nil
	}
	q := *l
	q.record = cloneRecord(l.record)
	q.controls = nil
	q.x = append([]byte(nil), l.x...)
	for _, c := range l.controls {
		d := *c
		d.site = cloneRecord(c.site)
		d.record = cloneRecord(c.record)
		d.child = cloneLevel(c.child)
		q.controls = append(q.controls, &d)
	}
	return &q
}
func (f *Form) allocateID(target *formLevel) int {
	next := int(f.root.record.values["NextAvailableID"])
	var walk func(*formLevel)
	walk = func(l *formLevel) {
		for _, c := range l.controls {
			next = max(next, int(c.site.values["ID"]))
			if c.child != nil {
				walk(c.child)
			}
		}
	}
	walk(f.root)
	next++
	var mark func(*formLevel)
	mark = func(l *formLevel) {
		if target.path == l.path || strings.HasPrefix(target.path, l.path+"/") {
			_ = l.record.set("NextAvailableID", next)
			if _, ok := l.record.values["ShapeCookie"]; ok {
				_ = l.record.set("ShapeCookie", l.record.values["ShapeCookie"]+1)
			}
		}
		for _, c := range l.controls {
			if c.child != nil {
				mark(c.child)
			}
		}
	}
	mark(f.root)
	return next
}
func internalTabs(l *formLevel) (*formRecord, error) {
	for _, c := range l.controls {
		if c.kind == "TabStrip" && c.record != nil {
			return c.record, nil
		}
	}
	return nil, fmt.Errorf("MultiPage has no internal TabStrip")
}

var tabArrays = []struct{ name, size string }{{"Items", "ItemsSize"}, {"TipStrings", "TipStringsSize"}, {"TabNames", "NamesSize"}, {"Tags", "TagsSize"}, {"Accelerators", "AcceleratorsSize"}}

func arrayStrings(b []byte, cp int) (out []string, err error) {
	for p := 0; p < len(b); {
		if p+4 > len(b) {
			return nil, fmt.Errorf("truncated tab string length")
		}
		n := U32(b, p)
		p += 4
		size := int(n & 0x7fffffff)
		if size > len(b)-p {
			return nil, fmt.Errorf("tab string outside array")
		}
		s := ""
		if n&0x80000000 != 0 {
			s, err = Decode(b[p:p+size], cp)
		} else {
			if size%2 != 0 {
				return nil, fmt.Errorf("odd Unicode tab string")
			}
			s = utf16str(b[p : p+size])
		}
		if err != nil {
			return nil, err
		}
		out = append(out, s)
		p += size
		p += (4 - p%4) % 4
		if p > len(b) {
			return nil, fmt.Errorf("tab string padding outside array")
		}
	}
	return out, nil
}
func encodeStrings(a []string) []byte {
	b := []byte{}
	for _, s := range a {
		raw := utf16bytes(s)
		b = append(b, dword(uint32(len(raw)))...)
		b = append(b, raw...)
		b = pad(b, 4)
	}
	return b
}
func pageStrings(l *formLevel, name string, cp int) ([]string, error) {
	t, e := internalTabs(l)
	if e != nil {
		return nil, e
	}
	return arrayStrings(t.extra[name], cp)
}

type pageBook struct {
	props    [][]byte
	mask, id uint32
	ids      []uint32
}

func readPageBook(b []byte) (q pageBook, err error) {
	for p := 0; p < len(b); {
		if p+8 > len(b) || b[p] != 0 || b[p+1] != 2 {
			return q, fmt.Errorf("invalid MultiPage x record")
		}
		end := p + 4 + int(U16(b, p+2))
		if end > len(b) || end < p+8 {
			return q, fmt.Errorf("MultiPage x length")
		}
		mask := U32(b, p+4)
		if mask&2 == 0 {
			q.props = append(q.props, append([]byte(nil), b[p:end]...))
			p = end
			continue
		}
		if p+12 > end {
			return q, fmt.Errorf("missing page count")
		}
		count := int(U32(b, p+8))
		q.mask = mask
		if mask&4 != 0 {
			if p+16 > end {
				return q, fmt.Errorf("missing page id")
			}
			q.id = U32(b, p+12)
		}
		if count > 10000 || end+4*count != len(b) {
			return q, fmt.Errorf("page ID count mismatch")
		}
		for p = end; p < len(b); p += 4 {
			q.ids = append(q.ids, U32(b, p))
		}
		if len(q.props) != len(q.ids)+1 {
			return q, fmt.Errorf("page-property count mismatch")
		}
		return q, nil
	}
	return q, fmt.Errorf("no MultiPage properties in x")
}

var emptyPage = []byte{0, 2, 4, 0, 0, 0, 0, 0}

func (q pageBook) bytes() []byte {
	b := []byte{}
	for _, p := range q.props {
		b = append(b, p...)
	}
	n := uint16(8)
	if q.mask&4 != 0 {
		n += 4
	}
	b = append(b, 0, 2)
	b = append(b, word(n)...)
	b = append(b, dword(q.mask)...)
	b = append(b, dword(uint32(len(q.ids)))...)
	if q.mask&4 != 0 {
		b = append(b, dword(q.id)...)
	}
	for _, id := range q.ids {
		b = append(b, dword(id)...)
	}
	return b
}
func (f *Form) initMultiPage(l *formLevel) error {
	l.kind = "MultiPage"
	l.trailing, _ = hex.DecodeString("00020c0019000000fc8f0000ff010000")
	id := f.allocateID(l)
	s := newRecord(siteSpec)
	for k, v := range map[string]any{"ID": id, "TabIndex": 0, "ClsidCacheIndex": 18, "ObjectStreamSize": 0} {
		if e := s.set(k, v); e != nil {
			return e
		}
	}
	_ = s.size("Position", 0, 0)
	r, e := defaultControl("TabStrip", "")
	if e != nil {
		return e
	}
	r.mask |= 1 << 19
	_ = r.size("Size", 144, 108)
	_ = r.set("ListIndex", 0)
	_ = r.set("TabsAllocated", 2)
	_ = r.set("TabData", 0)
	for _, a := range tabArrays {
		_ = r.set(a.size, 0)
		r.extra[a.name] = nil
	}
	l.controls = append(l.controls, &formControl{site: s, record: r, kind: "TabStrip"})
	l.x = (pageBook{props: [][]byte{emptyPage}, mask: 6, id: uint32(id)}).bytes()
	return nil
}
func tabTail(r *formRecord) (font []byte, flags []byte, err error) {
	p := 0
	for _, blob := range r.spec.blobs {
		if r.mask&(1<<blob.bit) != 0 {
			var e error
			p, e = pictureEnd(r.tail, p)
			if e != nil {
				return nil, nil, e
			}
		}
	}
	if p+4 > len(r.tail) {
		return nil, nil, fmt.Errorf("missing TabStrip font")
	}
	end := p + 4 + int(U16(r.tail, p+2))
	if end > len(r.tail) {
		return nil, nil, fmt.Errorf("TabStrip font boundary")
	}
	return r.tail[:end], r.tail[end:], nil
}
func (f *Form) applyPages(l *formLevel, designs []ControlDesign, remove []string, cp int) error {
	tabs, e := internalTabs(l)
	if e != nil {
		return e
	}
	book, e := readPageBook(l.x)
	if e != nil {
		return e
	}
	arrays := map[string][]string{}
	for _, a := range tabArrays {
		v, e := arrayStrings(tabs.extra[a.name], cp)
		if e != nil {
			return e
		}
		if len(v) != len(book.ids) {
			return fmt.Errorf("MultiPage %s array count mismatch", a.name)
		}
		arrays[a.name] = v
	}
	font, flagRaw, e := tabTail(tabs)
	if e != nil {
		return e
	}
	if len(flagRaw) != 4*len(book.ids) {
		return fmt.Errorf("MultiPage flags count mismatch")
	}
	flags := []uint32{}
	for i := 0; i < len(flagRaw); i += 4 {
		flags = append(flags, U32(flagRaw, i))
	}
	pages := []*formControl{}
	for _, c := range l.controls {
		if c.kind == "Form" {
			pages = append(pages, c)
		}
	}
	if len(pages) != len(book.ids) {
		return fmt.Errorf("page site count mismatch")
	}
	for i, c := range pages {
		if uint32(c.site.values["ID"]) != book.ids[i] {
			return fmt.Errorf("page ID order differs from site order")
		}
	}
	modified := false
	for _, name := range remove {
		at := -1
		for i, c := range pages {
			if strings.EqualFold(c.site.strings["Name"], name) {
				at = i
				break
			}
		}
		if at < 0 {
			return fmt.Errorf("no page %s", name)
		}
		pages = append(pages[:at], pages[at+1:]...)
		book.ids = append(book.ids[:at], book.ids[at+1:]...)
		book.props = append(book.props[:at+1], book.props[at+2:]...)
		flags = append(flags[:at], flags[at+1:]...)
		for k, a := range arrays {
			arrays[k] = append(a[:at], a[at+1:]...)
		}
		modified = true
	}
	seen := map[string]bool{}
	for _, d := range designs {
		if !ValidIdentifier(d.Name) || seen[strings.ToLower(d.Name)] {
			return fmt.Errorf("invalid/duplicate page name %s", d.Name)
		}
		seen[strings.ToLower(d.Name)] = true
		if d.Type != "" && d.Type != "Form" && d.Type != "Page" {
			return fmt.Errorf("MultiPage children must be Pages")
		}
		at := -1
		for i, c := range pages {
			if strings.EqualFold(c.site.strings["Name"], d.Name) {
				at = i
				break
			}
		}
		if at < 0 {
			id := f.allocateID(l)
			s := newRecord(siteSpec)
			for k, v := range map[string]any{"Name": d.Name, "ID": id, "TabIndex": len(pages) + 1, "ClsidCacheIndex": 7, "BitFlags": 262177} {
				if e = s.set(k, v); e != nil {
					return e
				}
			}
			_ = s.size("Position", 53.*72/2540, 556.*72/2540)
			r, e := defaultControl("Form", d.Name)
			if e != nil {
				return e
			}
			_ = r.size("DisplayedSize", 144, 108)
			pages = append(pages, &formControl{site: s, kind: "Form", child: &formLevel{path: fmt.Sprintf("%s/i%02d", l.path, id), record: r, structural: true, kind: "Form"}})
			l.controls = append(l.controls, pages[len(pages)-1])
			at = len(pages) - 1
			book.ids = append(book.ids, uint32(id))
			book.props = append(book.props, emptyPage)
			flags = append(flags, 3)
			for k, a := range arrays {
				s := ""
				if k == "Items" {
					s = d.Name
				}
				if k == "TabNames" {
					s = fmt.Sprintf("Tab%d", id)
				}
				arrays[k] = append(a, s)
			}
			modified = true
		}
		p := map[string]any{}
		for k, v := range d.Properties {
			p[k] = v
		}
		for key, array := range map[string]string{"Caption": "Items", "ControlTipText": "TipStrings", "Tag": "Tags", "Accelerator": "Accelerators"} {
			if v, ok := p[key]; ok {
				s, ok := v.(string)
				if !ok {
					return fmt.Errorf("page %s requires string", key)
				}
				if arrays[array][at] != s {
					arrays[array][at] = s
					modified = true
				}
				delete(p, key)
			}
		}
		if e = applyRecord(pages[at].child.record, p, "DisplayedSize"); e != nil {
			return e
		}
		if e = f.applyControls(pages[at].child, d.Controls, d.Remove, cp); e != nil {
			return e
		}
	}
	if !modified {
		return nil
	}
	l.controls = nil
	// The internal tab strip is always the first site and owns the o stream.
	ts := newRecord(siteSpec)
	_ = ts.set("ID", book.id)
	_ = ts.set("TabIndex", 0)
	_ = ts.set("ClsidCacheIndex", 18)
	_ = ts.set("ObjectStreamSize", 0)
	_ = ts.size("Position", 0, 0)
	l.controls = append(l.controls, &formControl{site: ts, record: tabs, kind: "TabStrip"})
	l.controls = append(l.controls, pages...)
	for i, c := range pages {
		_ = c.site.set("TabIndex", i+1)
		flag := 262177
		if i == 0 {
			flag = 262179
		}
		_ = c.site.set("BitFlags", flag)
	}
	for _, a := range tabArrays {
		raw := encodeStrings(arrays[a.name])
		tabs.extra[a.name] = raw
		_ = tabs.set(a.size, len(raw))
	}
	_ = tabs.set("TabData", len(pages))
	_ = tabs.set("TabsAllocated", max(int(tabs.values["TabsAllocated"]), len(pages)+2))
	idx := int(tabs.values["ListIndex"])
	if idx >= len(pages) {
		_ = tabs.set("ListIndex", max(0, len(pages)-1))
	}
	tabs.tail = append([]byte(nil), font...)
	for _, flag := range flags {
		tabs.tail = append(tabs.tail, dword(flag)...)
	}
	tabs.dirty = true
	l.structural = true
	l.x = book.bytes()
	return nil
}

// WriteBack preserves opaque streams and storage metadata, removes only
// designer subtrees that were explicitly deleted, and assigns correct CLSIDs
// to newly created containers. It is transactional on the destination CFB.
func (f *Form) WriteBack(c *Compound) error {
	streams, e := f.Streams()
	if e != nil {
		return e
	}
	work := c.Clone()
	levels := map[string]*formLevel{}
	var walk func(*formLevel)
	walk = func(l *formLevel) {
		levels[l.path] = l
		for _, ctl := range l.controls {
			if ctl.child != nil {
				walk(ctl.child)
			}
		}
	}
	walk(f.root)
	for p, x := range work.Entries {
		if x.Kind == 1 && strings.HasPrefix(p, f.Name+"/") {
			if _, known := f.CFB.Entries[p+"/f"]; known {
				if _, keep := levels[p]; !keep {
					work.Delete(p)
				}
			}
		}
	}
	for rel, b := range streams {
		if e = work.Set(f.Name+"/"+rel, b); e != nil {
			return e
		}
	}
	for path, l := range levels {
		if l.kind != "" && path != f.Name {
			entry := work.Entries[path]
			if entry == nil {
				return fmt.Errorf("missing container storage")
			}
			id, _ := hex.DecodeString(containerIDs[l.kind])
			copy(entry.Raw[80:96], id)
		}
	}
	// Reparse *all* storage references before committing.
	if _, e = ReadForm(work, f.Name, f.Codepage); e != nil {
		return fmt.Errorf("persisted form check: %w", e)
	}
	*c = *work
	return nil
}

// Stream comparison helper used by preservation tests.
func sameStreams(a, b map[string][]byte) bool {
	if len(a) != len(b) {
		return false
	}
	for k, v := range a {
		if !bytes.Equal(v, b[k]) {
			return false
		}
	}
	return true
}
