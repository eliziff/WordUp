// Package office implements bounded Office containers without activating Office.
package office

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"math/bits"
	"sort"
	"strings"
	"unicode/utf16"
)

const Limit = 256 << 20
const (
	free      uint32 = 0xffffffff
	end       uint32 = 0xfffffffe
	fatSector uint32 = 0xfffffffd
	difSector uint32 = 0xfffffffc
)

var cfbMagic = []byte{0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1}

func U16(b []byte, p int) uint16      { return binary.LittleEndian.Uint16(b[p:]) }
func U32(b []byte, p int) uint32      { return binary.LittleEndian.Uint32(b[p:]) }
func put16(b []byte, p int, v uint16) { binary.LittleEndian.PutUint16(b[p:], v) }
func put32(b []byte, p int, v uint32) { binary.LittleEndian.PutUint32(b[p:], v) }
func utf16bytes(s string) []byte {
	a := utf16.Encode([]rune(s))
	b := make([]byte, len(a)*2)
	for i, v := range a {
		put16(b, i*2, v)
	}
	return b
}
func utf16str(b []byte) string {
	a := make([]uint16, len(b)/2)
	for i := range a {
		a[i] = U16(b, i*2)
	}
	return string(utf16.Decode(a))
}
func pad(b []byte, n int) []byte {
	for len(b)%n != 0 {
		b = append(b, 0)
	}
	return b
}
func tlv(t uint16, b []byte) []byte {
	h := make([]byte, 6)
	put16(h, 0, t)
	put32(h, 2, uint32(len(b)))
	return append(h, b...)
}
func dword(n uint32) []byte { b := make([]byte, 4); put32(b, 0, n); return b }
func word(n uint16) []byte  { b := make([]byte, 2); put16(b, 0, n); return b }

type Entry struct {
	Name, Path string
	Kind       byte
	Raw        [128]byte
	Data       []byte
}
type Compound struct {
	Entries  map[string]*Entry
	Root     [128]byte
	Original []byte
}

func ReadCompound(data []byte) (out *Compound, err error) {
	defer func() {
		if r := recover(); r != nil {
			out = nil
			err = fmt.Errorf("invalid compound file: %v", r)
		}
	}()
	if len(data) < 512 || len(data) > Limit || !bytes.Equal(data[:8], cfbMagic) {
		return nil, fmt.Errorf("not a bounded CFB file")
	}
	major, shift := U16(data, 26), U16(data, 30)
	if !((major == 3 && shift == 9) || (major == 4 && shift == 12)) || U16(data, 28) != 0xfffe || U16(data, 32) != 6 || U32(data, 56) != 4096 {
		return nil, fmt.Errorf("unsupported CFB layout")
	}
	ss := 1 << shift
	if len(data)%ss != 0 {
		return nil, fmt.Errorf("truncated sector")
	}
	ns := len(data)/ss - 1
	sector := func(s uint32) []byte {
		if uint64(s) >= uint64(ns) {
			panic("sector out of bounds")
		}
		return data[(int(s)+1)*ss : (int(s)+2)*ss]
	}
	ids := []uint32{}
	for p := 76; p < 512; p += 4 {
		x := U32(data, p)
		if x != free {
			ids = append(ids, x)
		}
	}
	visited := map[uint32]bool{}
	di := U32(data, 68)
	nd := U32(data, 72)
	if uint64(nd) > uint64(ns) {
		return nil, fmt.Errorf("DIFAT count out of bounds")
	}
	for i := uint32(0); i < nd; i++ {
		if visited[di] {
			return nil, fmt.Errorf("DIFAT cycle")
		}
		visited[di] = true
		s := sector(di)
		for p := 0; p < ss-4; p += 4 {
			x := U32(s, p)
			if x != free {
				ids = append(ids, x)
			}
		}
		di = U32(s, ss-4)
	}
	if len(ids) != int(U32(data, 44)) {
		return nil, fmt.Errorf("FAT length mismatch")
	}
	table := []uint32{}
	visited = map[uint32]bool{}
	for _, id := range ids {
		if visited[id] {
			return nil, fmt.Errorf("duplicate FAT")
		}
		visited[id] = true
		s := sector(id)
		for p := 0; p < ss; p += 4 {
			table = append(table, U32(s, p))
		}
	}
	chain := func(start uint32, tab []uint32, max int, read func(uint32) []byte, size int) []byte {
		if size == 0 {
			return nil
		}
		b := []byte{}
		seen := map[uint32]bool{}
		for start != end {
			if uint64(start) >= uint64(max) || int(start) >= len(tab) || seen[start] {
				panic("invalid or cyclic chain")
			}
			seen[start] = true
			b = append(b, read(start)...)
			if len(b) > Limit {
				panic("stream budget exceeded")
			}
			start = tab[start]
		}
		if size >= 0 {
			if len(b) < size {
				panic("short stream")
			}
			b = b[:size]
		}
		return b
	}
	directory := chain(U32(data, 48), table, ns, sector, -1)
	if len(directory)%128 != 0 || len(directory) == 0 {
		return nil, fmt.Errorf("invalid directory")
	}
	if len(directory)/128 > 65536 {
		return nil, fmt.Errorf("CFB directory entry budget exceeded")
	}
	entries := make([]*Entry, len(directory)/128)
	totalStreams := 0
	starts := make([]uint32, len(entries))
	sizes := make([]int, len(entries))
	for i := range entries {
		r := directory[i*128 : (i+1)*128]
		e := &Entry{Kind: r[66]}
		copy(e.Raw[:], r)
		n := int(U16(r, 64))
		if e.Kind != 0 {
			if n < 2 || n > 64 || n%2 != 0 {
				return nil, fmt.Errorf("invalid name length")
			}
			e.Name = utf16str(r[:n-2])
		}
		entries[i] = e
		starts[i] = U32(r, 116)
		sz := binary.LittleEndian.Uint64(r[120:])
		if major == 3 {
			sz &= 0xffffffff
		}
		if sz > Limit {
			return nil, fmt.Errorf("stream size limit")
		}
		sizes[i] = int(sz)
		if e.Kind == 2 {
			totalStreams += int(sz)
			if totalStreams > Limit {
				return nil, fmt.Errorf("CFB aggregate stream budget exceeded")
			}
		}
	}
	if entries[0].Kind != 5 {
		return nil, fmt.Errorf("root storage absent")
	}
	mini := chain(starts[0], table, ns, sector, sizes[0])
	mfCount := U32(data, 64)
	if uint64(mfCount) > uint64(ns) {
		return nil, fmt.Errorf("MiniFAT count out of bounds")
	}
	mf := chain(U32(data, 60), table, ns, sector, int(mfCount)*ss)
	mt := []uint32{}
	for p := 0; p < len(mf); p += 4 {
		mt = append(mt, U32(mf, p))
	}
	out = &Compound{Entries: map[string]*Entry{}, Root: entries[0].Raw, Original: append([]byte(nil), data...)}
	type visit struct {
		idx    uint32
		parent string
	}
	stack := []visit{{U32(entries[0].Raw[:], 76), ""}}
	seen := map[uint32]bool{}
	names := map[string]bool{}
	for len(stack) > 0 {
		v := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		if v.idx == free {
			continue
		}
		if v.idx == 0 || int(v.idx) >= len(entries) || seen[v.idx] {
			return nil, fmt.Errorf("invalid directory tree")
		}
		seen[v.idx] = true
		e := entries[v.idx]
		if e.Kind != 1 && e.Kind != 2 {
			return nil, fmt.Errorf("invalid directory kind")
		}
		if e.Name == "" || strings.ContainsAny(e.Name, "/\\") {
			return nil, fmt.Errorf("invalid entry name")
		}
		e.Path = v.parent + e.Name
		key := strings.ToUpper(e.Path)
		if names[key] {
			return nil, fmt.Errorf("duplicate CFB path")
		}
		names[key] = true
		out.Entries[e.Path] = e
		stack = append(stack, visit{U32(e.Raw[:], 68), v.parent}, visit{U32(e.Raw[:], 72), v.parent})
		if e.Kind == 1 {
			stack = append(stack, visit{U32(e.Raw[:], 76), e.Path + "/"})
		} else if sizes[v.idx] >= 4096 {
			e.Data = chain(starts[v.idx], table, ns, sector, sizes[v.idx])
		} else {
			e.Data = chain(starts[v.idx], mt, len(mini)/64, func(s uint32) []byte { return mini[int(s)*64 : (int(s)+1)*64] }, sizes[v.idx])
		}
	}
	for i, e := range entries {
		if i > 0 && e.Kind != 0 && !seen[uint32(i)] {
			return nil, fmt.Errorf("unreachable entry %q", e.Name)
		}
	}
	return out, nil
}
func NewCompound() *Compound {
	c := &Compound{Entries: map[string]*Entry{}}
	e := newEntry("Root Entry", 5)
	c.Root = e.Raw
	return c
}
func newEntry(name string, kind byte) *Entry {
	e := &Entry{Name: name, Kind: kind}
	b := utf16bytes(name + "\x00")
	copy(e.Raw[:64], b)
	put16(e.Raw[:], 64, uint16(len(b)))
	e.Raw[66] = kind
	e.Raw[67] = 1
	for _, p := range []int{68, 72, 76} {
		put32(e.Raw[:], p, free)
	}
	return e
}

// Clone keeps failed builds from partially modifying the imported project.
func (c *Compound) Clone() *Compound {
	d := &Compound{Entries: make(map[string]*Entry, len(c.Entries)), Root: c.Root, Original: c.Original}
	for p, e := range c.Entries {
		x := *e
		d.Entries[p] = &x
	}
	return d
}
func (c *Compound) Stream(path string) ([]byte, error) {
	e := c.Entries[path]
	if e == nil || e.Kind != 2 {
		return nil, fmt.Errorf("missing stream %q", path)
	}
	return e.Data, nil
}
func (c *Compound) Set(path string, data []byte) error {
	if len(data) > Limit {
		return fmt.Errorf("stream too large")
	}
	for p := range c.Entries {
		if p != path && strings.EqualFold(p, path) {
			return fmt.Errorf("case-colliding CFB path %q", path)
		}
	}
	parts := strings.Split(path, "/")
	for _, s := range parts {
		if s == "" || strings.ContainsAny(s, "\\\x00") || len(utf16bytes(s)) > 62 || s == "." || s == ".." {
			return fmt.Errorf("invalid CFB name %q", s)
		}
	}
	for i := 1; i < len(parts); i++ {
		p := strings.Join(parts[:i], "/")
		if e := c.Entries[p]; e == nil {
			e = newEntry(parts[i-1], 1)
			e.Path = p
			c.Entries[p] = e
		} else if e.Kind != 1 {
			return fmt.Errorf("storage expected")
		}
	}
	e := c.Entries[path]
	if e == nil {
		e = newEntry(parts[len(parts)-1], 2)
		e.Path = path
		c.Entries[path] = e
	}
	if e.Kind != 2 {
		return fmt.Errorf("stream expected")
	}
	e.Data = append([]byte(nil), data...)
	return nil
}
func (c *Compound) Delete(path string) {
	for p := range c.Entries {
		if p == path || strings.HasPrefix(p, path+"/") {
			delete(c.Entries, p)
		}
	}
}
func (c *Compound) Bytes() ([]byte, error) {
	keys := make([]string, 0, len(c.Entries))
	total := 0
	for p, e := range c.Entries {
		keys = append(keys, p)
		total += len(e.Data)
	}
	if total > Limit-(1<<20) {
		return nil, fmt.Errorf("CFB output budget exceeded")
	}
	sort.Strings(keys)
	entries := []*Entry{{Name: "Root Entry", Kind: 5, Raw: c.Root}}
	for _, p := range keys {
		entries = append(entries, c.Entries[p])
	}
	raws := make([][128]byte, len(entries))
	children := map[string][]int{}
	for i, e := range entries {
		raws[i] = e.Raw
		for _, p := range []int{68, 72, 76} {
			put32(raws[i][:], p, free)
		}
		raws[i][67] = 1
		if i > 0 {
			parent := ""
			if j := strings.LastIndex(e.Path, "/"); j >= 0 {
				parent = e.Path[:j]
			}
			children[parent] = append(children[parent], i)
		}
	}
	for i, e := range entries {
		if e.Kind != 1 && e.Kind != 5 {
			continue
		}
		kids := children[e.Path]
		sort.Slice(kids, func(a, b int) bool {
			x, y := entries[kids[a]].Name, entries[kids[b]].Name
			lx, ly := len(utf16bytes(x)), len(utf16bytes(y))
			if lx != ly {
				return lx < ly
			}
			return strings.ToUpper(x) < strings.ToUpper(y)
		})
		red := bits.Len(uint(len(kids)+1)) - 1
		var tree func(int, int, int) uint32
		tree = func(lo, hi, level int) uint32 {
			if lo >= hi {
				return free
			}
			mid := (lo + hi) / 2
			k := kids[mid]
			if level == red {
				raws[k][67] = 0
			}
			put32(raws[k][:], 68, tree(lo, mid, level+1))
			put32(raws[k][:], 72, tree(mid+1, hi, level+1))
			return uint32(k)
		}
		put32(raws[i][:], 76, tree(0, len(kids), 0))
	}
	sectors := [][]byte{}
	chains := [][]uint32{}
	alloc := func(data []byte) uint32 {
		if len(data) == 0 {
			return end
		}
		start := uint32(len(sectors))
		chain := []uint32{}
		for p := 0; p < len(data); p += 512 {
			s := make([]byte, 512)
			copy(s, data[p:min(p+512, len(data))])
			chain = append(chain, uint32(len(sectors)))
			sectors = append(sectors, s)
		}
		chains = append(chains, chain)
		return start
	}
	mini := []byte{}
	mt := []uint32{}
	for i, e := range entries {
		if e.Kind != 2 {
			continue
		}
		var start uint32
		if len(e.Data) > 0 && len(e.Data) < 4096 {
			start = uint32(len(mt))
			n := (len(e.Data) + 63) / 64
			mini = append(mini, pad(append([]byte(nil), e.Data...), 64)...)
			for j := 0; j < n; j++ {
				next := start + uint32(j) + 1
				if j == n-1 {
					next = end
				}
				mt = append(mt, next)
			}
		} else {
			start = alloc(e.Data)
		}
		put32(raws[i][:], 116, start)
		binary.LittleEndian.PutUint64(raws[i][120:], uint64(len(e.Data)))
	}
	put32(raws[0][:], 116, alloc(mini))
	binary.LittleEndian.PutUint64(raws[0][120:], uint64(len(mini)))
	mfb := []byte{}
	for _, v := range mt {
		mfb = append(mfb, dword(v)...)
	}
	for len(mfb)%512 != 0 {
		mfb = append(mfb, 255)
	}
	mfStart := alloc(mfb)
	db := []byte{}
	for _, r := range raws {
		db = append(db, r[:]...)
	}
	ds := alloc(db)
	ndata := len(sectors)
	nf, nd := 1, 0
	for {
		f := (ndata + nf + nd + 127) / 128
		d := (max(0, f-109) + 126) / 127
		if f == nf && d == nd {
			break
		}
		nf, nd = f, d
	}
	ft := make([]uint32, nf*128)
	for i := range ft {
		ft[i] = free
	}
	for _, ch := range chains {
		for i, v := range ch {
			ft[v] = end
			if i+1 < len(ch) {
				ft[v] = ch[i+1]
			}
		}
	}
	for i := 0; i < nf; i++ {
		ft[ndata+i] = fatSector
	}
	for i := 0; i < nd; i++ {
		ft[ndata+nf+i] = difSector
	}
	for i := 0; i < nf; i++ {
		b := make([]byte, 512)
		for j := 0; j < 128; j++ {
			put32(b, j*4, ft[i*128+j])
		}
		sectors = append(sectors, b)
	}
	for i := 0; i < nd; i++ {
		b := bytes.Repeat([]byte{255}, 512)
		for j := 0; j < 127; j++ {
			k := 109 + i*127 + j
			if k < nf {
				put32(b, j*4, uint32(ndata+k))
			}
		}
		next := end
		if i+1 < nd {
			next = uint32(ndata + nf + i + 1)
		}
		put32(b, 508, next)
		sectors = append(sectors, b)
	}
	h := make([]byte, 512)
	copy(h, cfbMagic)
	put16(h, 24, 0x3e)
	put16(h, 26, 3)
	put16(h, 28, 0xfffe)
	put16(h, 30, 9)
	put16(h, 32, 6)
	put32(h, 44, uint32(nf))
	put32(h, 48, ds)
	put32(h, 56, 4096)
	put32(h, 60, mfStart)
	put32(h, 64, uint32(len(mfb)/512))
	df := end
	if nd > 0 {
		df = uint32(ndata + nf)
	}
	put32(h, 68, df)
	put32(h, 72, uint32(nd))
	for i := 0; i < 109; i++ {
		v := free
		if i < nf {
			v = uint32(ndata + i)
		}
		put32(h, 76+4*i, v)
	}
	for _, s := range sectors {
		h = append(h, s...)
	}
	return h, nil
}

func Decompress(src []byte) (out []byte, err error) {
	defer func() {
		if r := recover(); r != nil {
			out = nil
			err = fmt.Errorf("invalid OVBA compression: %v", r)
		}
	}()
	if len(src) == 0 || src[0] != 1 {
		return nil, fmt.Errorf("bad OVBA signature")
	}
	p := 1
	for p < len(src) {
		if p+2 > len(src) {
			return nil, fmt.Errorf("truncated chunk")
		}
		h := U16(src, p)
		p += 2
		size := int(h&0xfff) + 1
		stop := p + size
		if (h>>12)&7 != 3 || stop > len(src) {
			return nil, fmt.Errorf("invalid chunk header")
		}
		start := len(out)
		if h&0x8000 == 0 {
			if size != 4096 {
				return nil, fmt.Errorf("invalid raw chunk size")
			}
			out = append(out, src[p:stop]...)
			p = stop
		} else {
			for p < stop {
				flags := src[p]
				p++
				for bit := 0; bit < 8 && p < stop; bit++ {
					if flags&(1<<bit) == 0 {
						out = append(out, src[p])
						p++
					} else {
						if p+2 > stop {
							return nil, fmt.Errorf("truncated copy token")
						}
						token := U16(src, p)
						p += 2
						difference := len(out) - start
						if difference <= 0 {
							return nil, fmt.Errorf("copy token before output")
						}
						width := max(4, bits.Len(uint(difference-1)))
						lengthBits := 16 - width
						length := int(token&uint16((1<<lengthBits)-1)) + 3
						offset := int(token>>lengthBits) + 1
						if offset > difference || difference+length > 4096 {
							return nil, fmt.Errorf("invalid copy token bounds")
						}
						for i := 0; i < length; i++ {
							out = append(out, out[len(out)-offset])
						}
					}
					if len(out)-start > 4096 {
						return nil, fmt.Errorf("oversized chunk")
					}
				}
			}
		}
		if len(out) > Limit {
			return nil, fmt.Errorf("OVBA output budget exceeded")
		}
	}
	return out, nil
}

// Compress emits interoperable MS-OVBA LZ chunks, with a bounded hash chain.
// Incompressible blocks are stored raw. MS-OVBA 2.4.1.3.10 requires a
// partial raw final block to be zero-padded to 4096. Callers that require exact
// source bytes must use CompressExact; never split a short non-final chunk.
func Compress(src []byte) []byte {
	out := make([]byte, 1, 1+len(src))
	out[0] = 1
	for len(src) > 0 {
		n := min(4096, len(src))
		chunk := src[:n]
		body := compressChunk(chunk)
		if len(body) <= 4096 {
			out = append(out, word(uint16(len(body)-1)|0xb000)...)
			out = append(out, body...)
		} else {
			out = append(out, 0xff, 0x3f)
			out = append(out, chunk...)
			out = append(out, make([]byte, 4096-n)...)
		}
		src = src[n:]
	}
	return out
}

// CompressExact refuses the format's implicit padding rather than silently
// adding NULs to VBA source or emitting invalid short interior chunks.
func CompressExact(src []byte) ([]byte, error) {
	n := len(src) % 4096
	if n != 0 && len(compressChunk(src[len(src)-n:])) > 4096 {
		return nil, fmt.Errorf("incompressible final OVBA block (%d bytes) would require NUL padding; add a trailing blank line or comment to the source", n)
	}
	return Compress(src), nil
}

func compressChunk(src []byte) []byte {
	var heads [4096]int16
	var prev [4096]int16
	// Index+1 permits zero-initialized fixed tables; no heap/hash-map allocation.
	hash := func(i int) int { return int((uint32(src[i])*251 + uint32(src[i+1])*31 + uint32(src[i+2])) & 4095) }
	add := func(i int) {
		if i+2 < len(src) {
			h := hash(i)
			prev[i] = heads[h]
			heads[h] = int16(i + 1)
		}
	}
	out := make([]byte, 0, len(src)+len(src)/8+1)
	for pos := 0; pos < len(src); {
		flags := len(out)
		out = append(out, 0)
		for bit := uint(0); bit < 8 && pos < len(src); bit++ {
			best, offset := 0, 0
			width := max(4, bits.Len(uint(max(0, pos-1))))
			maximum := min((1<<(16-width))+2, len(src)-pos)
			if pos+2 < len(src) {
				at := int(heads[hash(pos)]) - 1
				for attempts := 0; at >= 0 && attempts < 32; attempts++ {
					distance := pos - at
					if distance > (1 << width) {
						break
					}
					length := 0
					for length < maximum && src[at+length] == src[pos+length] {
						length++
					}
					if length >= 3 && length > best {
						best = length
						offset = distance
						if best == maximum {
							break
						}
					}
					at = int(prev[at]) - 1
				}
			}
			if best >= 3 {
				out[flags] |= 1 << bit
				token := uint16((offset-1)<<(16-width) | (best - 3))
				out = append(out, word(token)...)
				for j := 0; j < best; j++ {
					add(pos + j)
				}
				pos += best
			} else {
				out = append(out, src[pos])
				add(pos)
				pos++
			}
		}
	}
	return out
}
