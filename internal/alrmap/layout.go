package alrmap

import (
	"errors"
	"math/big"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/eliziff/WordUp/internal/alrmap/ooxml"
)

var (
	nonFiniteRe      = regexp.MustCompile(`(?i)^[+-]?(inf|infinity|nan|snan)$`)
	decimalLiteralRe = regexp.MustCompile(`^[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)([eE][+-]?[0-9]+)?$`)
)

// Twips converts an explicit physical measurement to twips with exact
// decimal arithmetic and round-half-up. The guide's unitless "1.5" cannot be
// made into a constant by guessing.
func Twips(value, unit string) (int, error) {
	var factor *big.Rat
	switch unit {
	case "in":
		factor = big.NewRat(1440, 1)
	case "pt":
		factor = big.NewRat(20, 1)
	case "cm":
		factor = big.NewRat(144000, 254)
	default:
		return 0, errors.New("an explicit physical unit is required")
	}
	v := strings.TrimSpace(value)
	if nonFiniteRe.MatchString(v) {
		return 0, errors.New("non-finite geometry")
	}
	if !decimalLiteralRe.MatchString(v) {
		return 0, errors.New("invalid decimal literal")
	}
	number, ok := new(big.Rat).SetString(v)
	if !ok {
		return 0, errors.New("invalid decimal literal")
	}
	product := new(big.Rat).Mul(number, factor)
	negative := product.Sign() < 0
	product.Abs(product)
	// ROUND_HALF_UP on the magnitude: floor(x + 1/2).
	product.Add(product, big.NewRat(1, 2))
	q := new(big.Int).Quo(product.Num(), product.Denom())
	if !q.IsInt64() {
		return 0, errors.New("geometry outside supported range")
	}
	result := int(q.Int64())
	if negative {
		result = -result
	}
	return result, nil
}

// AffectedStyles walks basedOn descendants and linked partners of a style.
// Linked partners are included conservatively because native Word's style
// editing may affect the partner.
func AffectedStyles(styles *ooxml.Styles, changed string) (map[string]bool, error) {
	if styles.Items[changed] == nil {
		return nil, errors.New("unknown style")
	}
	keys := make([]string, 0, len(styles.Items))
	for k := range styles.Items {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	affected := map[string]bool{changed: true}
	for {
		before := len(affected)
		for _, key := range keys {
			node := styles.Items[key]
			base := ooxml.Attr(node.Child("basedOn"), "val", "")
			linked := ooxml.Attr(node.Child("link"), "val", "")
			if (base != "" && affected[base]) || (linked != "" && affected[linked]) {
				affected[key] = true
			}
			if affected[key] && linked != "" {
				if styles.Items[linked] == nil {
					return nil, errors.New("unresolved linked style")
				}
				affected[linked] = true
			}
		}
		if len(affected) == before {
			return affected, nil
		}
	}
}

// StyleUse is one paragraph/run/table style reference in a story part.
type StyleUse struct {
	Part         string
	ElementIndex int
	ElementKind  string
	Style        string
}

// ElementRef addresses one element of one part for ownership declarations.
type ElementRef struct {
	Part         string
	ElementIndex int
}

// StyleUses inspects actual paragraph/character/table style references,
// including defaults, in the supplied package parts. Unknown or cyclic
// references are errors, not silently ignored.
func StyleUses(parts map[string][]byte, styles *ooxml.Styles) ([]StyleUse, error) {
	names := make([]string, 0, len(parts))
	for part := range parts {
		names = append(names, part)
	}
	sort.Strings(names)
	var result []StyleUse
	for _, part := range names {
		root, err := ooxml.ParseXML(parts[part])
		if err != nil {
			return nil, err
		}
		for i, n := range append([]*ooxml.Node{root}, root.Descendants()...) {
			var kind, prop, child string
			switch n.Tag {
			case ooxml.Q("p"):
				kind, prop, child = "paragraph", "pPr", "pStyle"
			case ooxml.Q("r"):
				kind, prop, child = "character", "rPr", "rStyle"
			case ooxml.Q("tbl"):
				kind, prop, child = "table", "tblPr", "tblStyle"
			default:
				continue
			}
			var ref *ooxml.Node
			if pp := n.Child(prop); pp != nil {
				ref = pp.Child(child)
			}
			key := ooxml.Attr(ref, "val", styles.Defaults[kind])
			if key != "" {
				if _, err := styles.Chain(key, kind); err != nil {
					return nil, err
				}
			}
			result = append(result, StyleUse{part, i, kind, key})
		}
	}
	return result, nil
}

// RequireStyleOwnership refuses a definition change when an affected use lies
// outside the declared owned set. The caller must supply every relevant
// story part; omitted parts cannot be certified.
func RequireStyleOwnership(parts map[string][]byte, styles *ooxml.Styles, changed string, owned map[ElementRef]bool) ([]StyleUse, error) {
	affected, err := AffectedStyles(styles, changed)
	if err != nil {
		return nil, err
	}
	uses, err := StyleUses(parts, styles)
	if err != nil {
		return nil, err
	}
	var hits []StyleUse
	for _, u := range uses {
		if affected[u.Style] {
			hits = append(hits, u)
		}
	}
	for _, u := range hits {
		if !owned[ElementRef{u.Part, u.ElementIndex}] {
			return nil, errors.New("shared style affects text outside the declared ownership set")
		}
	}
	return hits, nil
}

// Heading is a stable supplied heading identity with its mapped level.
type Heading struct {
	ID           string
	Level        int
	Introduction bool
}

func roman(n int) (string, error) {
	if !(0 < n && n < 4000) {
		return "", errors.New("Roman label outside supported range")
	}
	var b strings.Builder
	for _, step := range []struct {
		value  int
		symbol string
	}{{1000, "M"}, {900, "CM"}, {500, "D"}, {400, "CD"}, {100, "C"}, {90, "XC"}, {50, "L"}, {40, "XL"}, {10, "X"}, {9, "IX"}, {5, "V"}, {4, "IV"}, {1, "I"}} {
		for n >= step.value {
			b.WriteString(step.symbol)
			n -= step.value
		}
	}
	return b.String(), nil
}

func letters(n int) string {
	out := ""
	for n > 0 {
		n--
		out = string(rune('A'+n%26)) + out
		n /= 26
	}
	return out
}

// HeadingLabels computes labels for stable supplied heading IDs under an
// explicit numbered/unnumbered Introduction policy. It never discovers
// headings from short bold paragraphs or rewrites text.
func HeadingLabels(headings []Heading, numberedIntroduction bool) (map[string]string, error) {
	ids := map[string]bool{}
	for _, h := range headings {
		if h.ID == "" || ids[h.ID] {
			return nil, errors.New("unique nonempty heading identities required")
		}
		ids[h.ID] = true
	}
	counts := [3]int{}
	previous := 0
	labels := map[string]string{}
	for i, h := range headings {
		if h.Level < 1 || h.Level > 3 {
			return nil, errors.New("only the mapped first three levels are supported")
		}
		if h.Introduction && (i != 0 || h.Level != 1) {
			return nil, errors.New("Introduction must be the first level-one heading")
		}
		if h.Introduction && !numberedIntroduction {
			labels[h.ID] = ""
			previous = 0
			continue
		}
		if h.Level > previous+1 {
			return nil, errors.New("missing parent heading")
		}
		counts[h.Level-1]++
		for level := h.Level; level < 3; level++ {
			counts[level] = 0
		}
		first, err := roman(counts[0])
		if err != nil {
			return nil, err
		}
		parts := []string{first, letters(counts[1]), strconv.Itoa(counts[2])}
		labels[h.ID] = strings.Join(parts[:h.Level], ".")
		previous = h.Level
	}
	return labels, nil
}
