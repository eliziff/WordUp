// Package ooxml inspects WordprocessingML parts directly: a byte-addressed
// element tree, protective style inheritance, structural scanning into
// protected intervals and text atoms, a lossless reference-copy writer and
// the separate title permission. It is not Word's object model, and it never
// writes a file in place.
//
// Text projections are addressed by code point (as in the Python reference)
// while nodes retain their original UTF-8 byte boundaries; the two address
// systems are never interchanged.
package ooxml

import (
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"regexp"
	"sort"
	"strings"
	"unicode/utf8"
)

// W is the WordprocessingML main namespace.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

// XML is the namespace of the reserved xml: prefix.
const XML = "http://www.w3.org/XML/1998/namespace"

// MaxXML bounds the accepted part size.
const MaxXML = 32 * 1024 * 1024

// Q returns the expanded WordprocessingML name of a local name.
func Q(local string) string { return W + "}" + local }

// Attr reads the w:-namespaced attribute, or def when the node or attribute
// is absent.
func Attr(n *Node, local, def string) string {
	if n == nil {
		return def
	}
	if v, ok := n.Attrs[Q(local)]; ok {
		return v
	}
	return def
}

// Node is one element with its original byte coordinates. Tag is the
// expat-style "namespace}local" (or a bare local name outside any namespace);
// namespace declarations are not attributes.
type Node struct {
	Tag        string
	Attrs      map[string]string
	Start      int // byte offset of '<'
	OpenEnd    int // byte offset after the start tag's '>'
	CloseStart int // byte offset of the end tag ('</'); equals OpenEnd when self-closing
	End        int // byte offset after the end tag's '>'
	Children   []*Node
	parent     *Node
	text       []byte
}

// Parent returns the enclosing element, or nil for the root.
func (n *Node) Parent() *Node { return n.parent }

// Text returns the element's direct character data (not its descendants').
func (n *Node) Text() string { return string(n.text) }

// Child returns the first direct child with the given w: local name.
func (n *Node) Child(local string) *Node {
	tag := Q(local)
	for _, c := range n.Children {
		if c.Tag == tag {
			return c
		}
	}
	return nil
}

// Descendants lists every descendant in document order.
func (n *Node) Descendants() []*Node {
	var out []*Node
	var walk func(*Node)
	walk = func(x *Node) {
		for _, c := range x.Children {
			out = append(out, c)
			walk(c)
		}
	}
	walk(n)
	return out
}

var (
	xmlDeclRe  = regexp.MustCompile(`^\x{FEFF}?<\?xml(?s:.*?)\?>`)
	encodingRe = regexp.MustCompile(`(?i)encoding\s*=\s*['"]([^'"]+)`)
)

// ParseXML builds the tree of one UTF-8 part. No DTD/entity execution.
func ParseXML(raw []byte) (*Node, error) {
	if len(raw) == 0 || len(raw) > MaxXML {
		return nil, errors.New("XML size outside supported bounds")
	}
	if !utf8.Valid(raw) {
		return nil, errors.New("XML part is not valid UTF-8")
	}
	if decl := xmlDeclRe.Find(raw); decl != nil {
		if enc := encodingRe.FindSubmatch(decl); enc != nil {
			if e := strings.ToLower(string(enc[1])); e != "utf-8" && e != "utf8" {
				return nil, errors.New("unsupported XML encoding")
			}
		}
	}
	p := &parser{raw: raw, d: xml.NewDecoder(bytes.NewReader(raw))}
	p.d.Strict = true
	return p.parse()
}

type frame struct {
	node    *Node
	rawName string            // prefix:local as written, for end-tag matching
	ns      map[string]string // in-scope prefix bindings ("" is the default namespace)
}

type parser struct {
	raw   []byte
	d     *xml.Decoder
	stack []frame
	roots []*Node
	count int
}

func (p *parser) scope() map[string]string {
	if len(p.stack) == 0 {
		return map[string]string{"xml": XML}
	}
	return p.stack[len(p.stack)-1].ns
}

func rawName(n xml.Name) string {
	if n.Space == "" {
		return n.Local
	}
	return n.Space + ":" + n.Local
}

func (p *parser) parse() (*Node, error) {
	for {
		before := p.d.InputOffset()
		tok, err := p.d.RawToken()
		after := p.d.InputOffset()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if err := p.start(t, int(before), int(after)); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if len(p.stack) == 0 {
				return nil, errors.New("unexpected end tag")
			}
			top := p.stack[len(p.stack)-1]
			if rawName(t.Name) != top.rawName {
				return nil, fmt.Errorf("mismatched end tag </%s>", rawName(t.Name))
			}
			p.stack = p.stack[:len(p.stack)-1]
			top.node.CloseStart, top.node.End = int(before), int(after)
		case xml.CharData:
			if len(p.stack) > 0 {
				top := p.stack[len(p.stack)-1].node
				top.text = append(top.text, t...)
			} else if !(before == 0 && string(t) == string(rune(0xFEFF))) && strings.TrimSpace(string(t)) != "" {
				return nil, errors.New("text outside the root element")
			}
		case xml.Comment, xml.ProcInst:
			if len(p.stack) > 0 && p.stack[len(p.stack)-1].node.Tag == Q("t") {
				return nil, errors.New("non-character XML content inside text node")
			}
		case xml.Directive:
			return nil, errors.New("DTD/entities are not supported")
		}
	}
	if len(p.stack) > 0 {
		return nil, errors.New("unclosed XML element")
	}
	if len(p.roots) != 1 {
		return nil, errors.New("one XML root required")
	}
	return p.roots[0], nil
}

func (p *parser) start(t xml.StartElement, before, after int) error {
	p.count++
	if p.count > 500_000 || len(p.stack) >= 256 {
		return errors.New("XML complexity limit")
	}
	ns := p.scope()
	declared := false
	for _, a := range t.Attr {
		if a.Name.Space == "xmlns" || (a.Name.Space == "" && a.Name.Local == "xmlns") {
			if !declared {
				copied := make(map[string]string, len(ns)+1)
				for k, v := range ns {
					copied[k] = v
				}
				ns, declared = copied, true
			}
			if a.Name.Space == "xmlns" {
				ns[a.Name.Local] = a.Value
			} else {
				ns[""] = a.Value
			}
		}
	}
	resolve := func(name xml.Name, element bool) (string, error) {
		if name.Space == "" {
			if element && ns[""] != "" {
				return ns[""] + "}" + name.Local, nil
			}
			return name.Local, nil
		}
		uri, ok := ns[name.Space]
		if !ok {
			return "", fmt.Errorf("unbound prefix %q", name.Space)
		}
		return uri + "}" + name.Local, nil
	}
	tag, err := resolve(t.Name, true)
	if err != nil {
		return err
	}
	attrs := make(map[string]string, len(t.Attr))
	for _, a := range t.Attr {
		if a.Name.Space == "xmlns" || (a.Name.Space == "" && a.Name.Local == "xmlns") {
			continue
		}
		key, err := resolve(a.Name, false)
		if err != nil {
			return err
		}
		if _, dup := attrs[key]; dup {
			return fmt.Errorf("duplicate attribute %q", key)
		}
		attrs[key] = a.Value
	}
	n := &Node{Tag: tag, Attrs: attrs, Start: before, OpenEnd: after}
	if len(p.stack) > 0 {
		parent := p.stack[len(p.stack)-1].node
		n.parent = parent
		parent.Children = append(parent.Children, n)
	} else {
		p.roots = append(p.roots, n)
	}
	p.stack = append(p.stack, frame{node: n, rawName: rawName(t.Name), ns: ns})
	return nil
}

// NodeShape renders a canonical description of a subtree (tag, sorted
// attributes, direct text, children) used for structural fingerprints.
func NodeShape(n *Node) string {
	var b strings.Builder
	var write func(*Node)
	write = func(x *Node) {
		keys := make([]string, 0, len(x.Attrs))
		for k := range x.Attrs {
			keys = append(keys, k)
		}
		sort.Strings(keys)
		fmt.Fprintf(&b, "(%q, (", x.Tag)
		for _, k := range keys {
			fmt.Fprintf(&b, "(%q, %q), ", k, x.Attrs[k])
		}
		fmt.Fprintf(&b, "), %q, (", x.Text())
		for _, c := range x.Children {
			write(c)
			b.WriteString(", ")
		}
		b.WriteString("))")
	}
	write(n)
	return b.String()
}
