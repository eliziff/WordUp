package ooxml

import (
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"sort"
	"strconv"
	"strings"

	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
)

// OnOff reads a w:val on/off attribute (absent means true).
func OnOff(n *Node) (bool, error) {
	value := Attr(n, "val", "true")
	switch value {
	case "true", "on", "1":
		return true, nil
	case "false", "off", "0":
		return false, nil
	}
	return false, errors.New("unsupported on/off value")
}

// toggles are the run properties with documented toggle semantics: a true
// value in a style flips the inherited state, false leaves it unchanged, and
// direct formatting sets the absolute value.
var toggles = []string{"i", "iCs", "b", "bCs", "vanish", "caps", "smallCaps", "strike"}

var toggleTags = func() map[string]string {
	m := map[string]string{}
	for _, t := range toggles {
		m[Q(t)] = t
	}
	return m
}()

// Styles resolves protective run flags through actual basedOn chains.
// Missing/cyclic styles, table/numbering inheritance and unsupported script
// features never yield a claim of a known plain run. Full font/layout
// resolution is NOT done.
type Styles struct {
	Raw      []byte
	Digest   string
	Root     *Node
	Items    map[string]*Node
	Defaults map[string]string
	base     *Node
	pbase    *Node
	cache    map[[2]string][]*Node
}

// NewStyles parses a styles part.
func NewStyles(raw []byte) (*Styles, error) {
	root, err := ParseXML(raw)
	if err != nil {
		return nil, err
	}
	if root.Tag != Q("styles") {
		return nil, errors.New("unsupported styles namespace")
	}
	sum := sha256.Sum256(raw)
	s := &Styles{Raw: raw, Digest: hex.EncodeToString(sum[:]), Root: root, Items: map[string]*Node{}, Defaults: map[string]string{}, cache: map[[2]string][]*Node{}}
	for _, n := range root.Children {
		if n.Tag != Q("style") {
			continue
		}
		key := Attr(n, "styleId", "")
		if key == "" || s.Items[key] != nil {
			return nil, errors.New("missing/duplicate style ID")
		}
		s.Items[key] = n
		switch Attr(n, "default", "") {
		case "1", "true", "on":
			typ := Attr(n, "type", "")
			if _, dup := s.Defaults[typ]; dup {
				return nil, errors.New("duplicate default style")
			}
			s.Defaults[typ] = key
		}
	}
	if dd := root.Child("docDefaults"); dd != nil {
		if rd := dd.Child("rPrDefault"); rd != nil {
			s.base = rd.Child("rPr")
		}
		if pd := dd.Child("pPrDefault"); pd != nil {
			s.pbase = pd.Child("pPr")
		}
	}
	return s, nil
}

// Chain returns the basedOn ancestry of a style, root ancestor first.
func (s *Styles) Chain(key, kind string) ([]*Node, error) {
	ck := [2]string{key, kind}
	if out, ok := s.cache[ck]; ok {
		return out, nil
	}
	seen := map[string]bool{}
	var out []*Node
	for key != "" {
		n, ok := s.Items[key]
		if seen[key] || !ok {
			return nil, errors.New("missing/cyclic style")
		}
		seen[key] = true
		if Attr(n, "type", "") != kind {
			return nil, errors.New("style type mismatch")
		}
		out = append(out, n)
		key = Attr(n.Child("basedOn"), "val", "")
	}
	for i, j := 0, len(out)-1; i < j; i, j = i+1, j-1 {
		out[i], out[j] = out[j], out[i]
	}
	s.cache[ck] = out
	return out, nil
}

// Run resolves the protective flags of run r inside paragraph p and returns
// the sorted risk names plus a property signature. The signature is
// deliberately stricter than effective equality: it hashes the styles
// snapshot and the raw pPr/rPr trees, so insertion into merely similar runs
// is refused rather than guessed.
func (s *Styles) Run(p, r *Node) ([]string, string, error) {
	ppr, rpr := p.Child("pPr"), r.Child("rPr")
	var pStyle, rStyle *Node
	if ppr != nil {
		pStyle = ppr.Child("pStyle")
	}
	if rpr != nil {
		rStyle = rpr.Child("rStyle")
	}
	pk := Attr(pStyle, "val", s.Defaults["paragraph"])
	rk := Attr(rStyle, "val", s.Defaults["character"])
	pc, err := s.Chain(pk, "paragraph")
	if err != nil {
		return nil, "", err
	}
	rc, err := s.Chain(rk, "character")
	if err != nil {
		return nil, "", err
	}
	state := map[string]bool{}
	for _, t := range toggles {
		state[t] = false
	}
	extra := map[string]string{}
	add := func(props *Node, toggle bool) error {
		if props == nil {
			return nil
		}
		for _, n := range props.Children {
			name := strings.TrimPrefix(n.Tag, W+"}")
			if t, ok := toggleTags[n.Tag]; ok {
				val, err := OnOff(n)
				if err != nil {
					return err
				}
				if !toggle {
					state[t] = val
				} else if val {
					state[t] = !state[t]
				}
				continue
			}
			switch name {
			case "webHidden", "dstrike":
				val, err := OnOff(n)
				if err != nil {
					return err
				}
				extra[name] = strconv.FormatBool(val)
			case "u", "vertAlign", "lang":
				extra[name] = Attr(n, "val", "")
			case "rPrChange", "rtl", "cs", "eastAsianLayout", "specVanish":
				return errors.New("unsupported or revision-dependent run formatting")
			}
			// Other properties are not interpreted but remain in the raw fingerprint.
		}
		return nil
	}
	if err := add(s.base, false); err != nil {
		return nil, "", err
	}
	for _, n := range append(append([]*Node(nil), pc...), rc...) {
		if err := add(n.Child("rPr"), true); err != nil {
			return nil, "", err
		}
	}
	if err := add(rpr, false); err != nil {
		return nil, "", err
	}
	riskSet := map[string]bool{}
	for name, v := range state {
		if v {
			riskSet[name] = true
		}
	}
	for _, k := range []string{"webHidden", "dstrike"} {
		if extra[k] == "true" {
			riskSet[k] = true
		}
	}
	if u, ok := extra["u"]; ok && u != "" && u != "none" {
		riskSet["underline"] = true
	}
	if v, ok := extra["vertAlign"]; ok && v != "baseline" {
		riskSet["vertical"] = true
	}
	if lang := extra["lang"]; lang != "" && !strings.HasPrefix(strings.ToLower(lang), "en") {
		riskSet["language"] = true
	}
	risk := make([]string, 0, len(riskSet))
	for name := range riskSet {
		risk = append(risk, name)
	}
	sort.Strings(risk)
	fragments := []string{s.Digest, pk, rk}
	for _, n := range []*Node{ppr, rpr} {
		if n != nil {
			// p/r live in the story part, not styles.Raw: canonical attrs/tree.
			fragments = append(fragments, NodeShape(n))
		}
	}
	sum := sha256.Sum256([]byte(strings.Join(fragments, "\x00")))
	return risk, hex.EncodeToString(sum[:]), nil
}

// ParagraphRisk reports whether inherited or direct paragraph formatting
// (numbering, frames, formatting revisions, outline roles, indentation or a
// known quote-style name) vetoes ordinary lexical operations.
func (s *Styles) ParagraphRisk(p *Node) (bool, error) {
	direct := p.Child("pPr")
	var pStyle *Node
	if direct != nil {
		pStyle = direct.Child("pStyle")
	}
	key := Attr(pStyle, "val", s.Defaults["paragraph"])
	chain, err := s.Chain(key, "paragraph")
	if err != nil {
		return false, err
	}
	props := map[string]*Node{}
	ind := map[string]string{}
	sources := []*Node{s.pbase}
	for _, st := range chain {
		sources = append(sources, st.Child("pPr"))
	}
	sources = append(sources, direct)
	for _, pp := range sources {
		if pp == nil {
			continue
		}
		for _, n := range pp.Children {
			if n.Tag == Q("ind") {
				for k, v := range n.Attrs {
					ind[k] = v
				}
			} else {
				props[n.Tag] = n
			}
		}
	}
	for _, k := range []string{"numPr", "framePr", "pPrChange"} {
		if _, ok := props[Q(k)]; ok {
			return true, nil
		}
	}
	if outline, ok := props[Q("outlineLvl")]; ok && Attr(outline, "val", "") != "9" {
		return true, nil
	}
	for _, k := range []string{"left", "right", "start", "end", "leftChars", "rightChars"} {
		if v, ok := ind[Q(k)]; ok {
			n, err := strconv.Atoi(strings.TrimSpace(v))
			if err != nil {
				return false, errors.New("invalid literal for int(): " + v)
			}
			if n != 0 {
				return true, nil
			}
		}
	}
	for _, st := range chain {
		if strings.Contains(pytext.Casefold(Attr(st.Child("name"), "val", "")), "quote") {
			return true, nil
		}
	}
	return false, nil
}
