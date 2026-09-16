package ooxml

import (
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"strings"

	"github.com/eliziff/WordUp/internal/alrmap/kernels"
)

// Atom is one editable w:t text node projected onto code-point positions.
type Atom struct {
	Start, End int
	Node       *Node
	Signature  string
}

// Container is one logical text container (the body, or one note) with its
// derived protections.
type Container struct {
	Key        string
	Text       string
	Atoms      []Atom
	Hard       []kernels.Span
	Lexical    []kernels.Span
	Paragraphs [][2]int
}

// Guards derives the protections for an operation family. Citation parsers
// have their own title/qualification boundaries, but they may not bypass
// delimited source text either; only syntax terminals are edited.
func (c *Container) Guards(family string) ([]kernels.Span, error) {
	if family != "lexical" && family != "citation" {
		return nil, errors.New("unsupported operation family")
	}
	extra := append([]kernels.Span(nil), c.Hard...)
	if family == "lexical" {
		extra = append(extra, c.Lexical...)
	}
	return kernels.Protection(c.Text, extra)
}

// Scan is the snapshot of one story part against one styles part.
type Scan struct {
	Raw        []byte
	Styles     *Styles
	Part       string
	Containers []*Container
	Root       *Node
}

// Fingerprint identifies the complete part/styles/part-name snapshot.
func (s *Scan) Fingerprint() string {
	h := sha256.New()
	h.Write(s.Raw)
	h.Write([]byte(s.Styles.Digest))
	h.Write([]byte(s.Part))
	return hex.EncodeToString(h.Sum(nil))
}

// Container returns the container with the given key.
func (s *Scan) Container(key string) (*Container, error) {
	for _, c := range s.Containers {
		if c.Key == key {
			return c, nil
		}
	}
	return nil, errors.New("unknown container " + key)
}

var rangedTags = map[string]struct {
	group   string
	isStart bool
}{
	"bookmarkStart": {"bookmark", true}, "bookmarkEnd": {"bookmark", false},
	"commentRangeStart": {"comment", true}, "commentRangeEnd": {"comment", false},
	"permStart": {"permission", true}, "permEnd": {"permission", false},
	"moveFromRangeStart": {"moveFrom", true}, "moveFromRangeEnd": {"moveFrom", false},
	"moveToRangeStart": {"moveTo", true}, "moveToRangeEnd": {"moveTo", false},
}

var safeTags = func() map[string]bool {
	m := map[string]bool{}
	for _, s := range []string{"body", "footnote", "endnote", "p", "r", "proofErr"} {
		m[Q(s)] = true
	}
	return m
}()

type fieldState struct {
	start     int
	separated bool
}

type rangeKey struct{ group, id string }

// ScanPart projects actual story/note structure into protected intervals and
// original text atoms.
func ScanPart(raw []byte, styles *Styles, part string) (*Scan, error) {
	root, err := ParseXML(raw)
	if err != nil {
		return nil, err
	}
	type owner struct {
		key  string
		node *Node
	}
	var owners []owner
	switch root.Tag {
	case Q("document"):
		b := root.Child("body")
		if b == nil {
			return nil, errors.New("missing body")
		}
		owners = []owner{{"body", b}}
	case Q("footnotes"), Q("endnotes"):
		typ := "endnote"
		if root.Tag == Q("footnotes") {
			typ = "footnote"
		}
		seen := map[string]bool{}
		for _, n := range root.Children {
			if n.Tag != Q(typ) || Attr(n, "type", "normal") != "normal" {
				continue
			}
			key := typ + ":" + Attr(n, "id", "")
			owners = append(owners, owner{key, n})
			seen[key] = true
		}
		if len(seen) != len(owners) {
			return nil, errors.New("duplicate note ID")
		}
		for _, o := range owners {
			if strings.HasSuffix(o.key, ":") {
				return nil, errors.New("missing note ID")
			}
		}
	default:
		return nil, errors.New("unsupported story part")
	}
	var containers []*Container
	for _, o := range owners {
		w := &walker{styles: styles, ranges: map[rangeKey]int{}}
		w.walk(o.node, nil, nil)
		text := string(w.chunks)
		if len(w.fields) > 0 || len(w.ranges) > 0 || len(w.errors) > 0 {
			w.hard = append(w.hard, kernels.Span{Start: 0, End: w.pos, Reason: "malformed-or-unclosed-structure"})
		}
		containers = append(containers, &Container{Key: o.key, Text: text, Atoms: w.atoms, Hard: w.hard, Lexical: w.lexical, Paragraphs: w.paragraphs})
	}
	return &Scan{Raw: raw, Styles: styles, Part: part, Containers: containers, Root: root}, nil
}

type walker struct {
	styles     *Styles
	chunks     []rune
	pos        int
	atoms      []Atom
	hard       []kernels.Span
	lexical    []kernels.Span
	paragraphs [][2]int
	fields     []fieldState
	ranges     map[rangeKey]int
	errors     []string
}

func (w *walker) emit(s string) int {
	lo := w.pos
	r := []rune(s)
	w.chunks = append(w.chunks, r...)
	w.pos += len(r)
	return lo
}

func (w *walker) walk(n, p, r *Node) {
	lo := w.pos
	local := strings.TrimPrefix(n.Tag, W+"}")
	if n.Tag == Q("p") {
		p = n
	}
	if n.Tag == Q("r") {
		r = n
	}
	if n.Tag == Q("pPr") || n.Tag == Q("rPr") {
		return
	}
	if n.Tag == Q("t") {
		start := w.emit(n.Text())
		signature := "unknown"
		var err error
		if p == nil || r == nil || len(n.Children) > 0 {
			err = errors.New("text outside simple run")
		} else {
			var risk []string
			risk, signature, err = w.styles.Run(p, r)
			if err == nil {
				if len(risk) > 0 {
					w.lexical = append(w.lexical, kernels.Span{Start: start, End: w.pos, Reason: "run:" + strings.Join(risk, ",")})
				}
				for _, name := range risk {
					if name == "vanish" || name == "webHidden" {
						w.hard = append(w.hard, kernels.Span{Start: start, End: w.pos, Reason: "hidden"})
						break
					}
				}
			}
		}
		if err != nil {
			signature = "unknown"
			w.hard = append(w.hard, kernels.Span{Start: start, End: w.pos, Reason: "unknown-format:" + err.Error()})
		}
		w.atoms = append(w.atoms, Atom{Start: start, End: w.pos, Node: n, Signature: signature})
		return
	}
	if n.Tag == Q("fldChar") {
		switch Attr(n, "fldCharType", "") {
		case "begin":
			w.fields = append(w.fields, fieldState{w.pos, false})
		case "separate":
			if len(w.fields) == 0 || w.fields[len(w.fields)-1].separated {
				w.errors = append(w.errors, "malformed-field")
			} else {
				w.fields[len(w.fields)-1].separated = true
			}
		case "end":
			if len(w.fields) == 0 {
				w.errors = append(w.errors, "unmatched-field-end")
			} else {
				f := w.fields[len(w.fields)-1]
				w.fields = w.fields[:len(w.fields)-1]
				w.hard = append(w.hard, kernels.Span{Start: f.start, End: w.pos, Reason: "complex-field"})
			}
		default:
			w.errors = append(w.errors, "unknown-field-token")
		}
		w.hard = append(w.hard, kernels.Span{Start: w.pos, End: w.pos, Reason: "field-boundary"})
		return
	}
	if rt, ok := rangedTags[local]; ok && n.Tag == Q(local) {
		rid := rangeKey{rt.group, Attr(n, "id", "")}
		if rid.id == "" {
			w.errors = append(w.errors, "missing-range-id")
		} else if rt.isStart {
			if _, dup := w.ranges[rid]; dup {
				w.errors = append(w.errors, "duplicate-range-start")
			}
			w.ranges[rid] = w.pos
		} else if start, ok := w.ranges[rid]; !ok {
			w.errors = append(w.errors, "unmatched-range-end")
		} else {
			delete(w.ranges, rid)
			w.hard = append(w.hard, kernels.Span{Start: start, End: w.pos, Reason: rt.group})
		}
		w.hard = append(w.hard, kernels.Span{Start: w.pos, End: w.pos, Reason: rt.group + "-boundary"})
		return
	}
	if strings.HasSuffix(local, "RangeStart") || strings.HasSuffix(local, "RangeEnd") {
		w.errors = append(w.errors, "unsupported-range-pair")
		return
	}
	if n.Tag == Q("instrText") {
		if len(w.fields) == 0 {
			w.errors = append(w.errors, "orphan-field-instruction")
		}
		// No projection of hidden field code into ordinary visible text.
		w.hard = append(w.hard, kernels.Span{Start: w.pos, End: w.pos, Reason: "instruction"})
		return
	}
	if n.Tag == Q("delText") || n.Tag == Q("delInstrText") {
		w.emit(n.Text())
		w.hard = append(w.hard, kernels.Span{Start: lo, End: w.pos, Reason: "deleted-text"})
		return
	}
	if n.Tag == Q("tab") || n.Tag == Q("br") || n.Tag == Q("cr") {
		if local == "tab" {
			w.emit("\t")
		} else {
			w.emit("\n")
		}
		w.hard = append(w.hard, kernels.Span{Start: lo, End: w.pos, Reason: local})
		return
	}
	if n.Tag == Q("footnoteRef") || n.Tag == Q("endnoteRef") {
		// Native leading note marks are separate from note text, not digits.
		w.hard = append(w.hard, kernels.Span{Start: w.pos, End: w.pos, Reason: "native-note-mark"})
		return
	}
	switch n.Tag {
	case Q("footnoteReference"), Q("endnoteReference"), Q("sym"), Q("noBreakHyphen"), Q("softHyphen"):
		w.emit(string(rune(0xFFFC)))
		w.hard = append(w.hard, kernels.Span{Start: lo, End: w.pos, Reason: local})
		return
	}
	opaque := !safeTags[n.Tag]
	for _, child := range n.Children {
		w.walk(child, p, r)
	}
	if opaque {
		w.hard = append(w.hard, kernels.Span{Start: lo, End: w.pos, Reason: "opaque:" + local})
	}
	if n.Tag == Q("p") {
		if risk, err := w.styles.ParagraphRisk(n); err != nil {
			w.hard = append(w.hard, kernels.Span{Start: lo, End: w.pos, Reason: "unknown-paragraph-style"})
		} else if risk {
			w.lexical = append(w.lexical, kernels.Span{Start: lo, End: w.pos, Reason: "paragraph-role"})
		}
		for _, d := range n.Descendants() {
			if strings.HasSuffix(d.Tag, "PrChange") {
				w.hard = append(w.hard, kernels.Span{Start: lo, End: w.pos, Reason: "format-revision"})
				break
			}
		}
		w.paragraphs = append(w.paragraphs, [2]int{lo, w.pos})
		w.emit("\r")
	}
}
