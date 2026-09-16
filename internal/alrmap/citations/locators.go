package citations

import (
	"errors"
	"regexp"
	"strconv"
	"strings"

	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
	"github.com/eliziff/WordUp/internal/alrmap/kernels"
)

var (
	numberSeparatorRe = regexp.MustCompile(`, *| +and +`)
	legalSeparatorRe  = regexp.MustCompile(` to |, `)
)

var (
	legalUnits  = map[string]string{"s": "section", "ss": "section", "art": "article", "arts": "article", "rule": "rule", "rules": "rule"}
	legalPlural = map[string]string{"s": "ss", "ss": "ss", "art": "arts", "arts": "arts", "rule": "rules", "rules": "rules"}
	unitLabels  = map[string][2]string{"section": {"s", "ss"}, "article": {"art", "arts"}, "rule": {"rule", "rules"}}
)

// numberPiece is one NUMBER match with code-point group offsets inside raw.
type numberPiece struct {
	start, end   int // piece bounds
	leftEnd      int
	dashStart    int
	dashEnd      int
	rightStart   int
	rightEnd     int
	left, right  string
	dash         string
	hasRange     bool
	renderedText string
}

func atoi(s string) (int, error) {
	n, err := strconv.Atoi(s)
	if err != nil {
		return 0, errors.New("unsupported numeric locator")
	}
	return n, nil
}

// splitNumbers finds the separator matches and the piece bounds of a numeric
// locator value, all as code-point offsets.
func splitNumbers(values string) (seps [][2]int, pieces [][2]int) {
	offsets := pytext.RuneOffsets(values)
	for _, m := range numberSeparatorRe.FindAllStringIndex(values, -1) {
		seps = append(seps, [2]int{offsets[m[0]], offsets[m[1]]})
	}
	start := 0
	for _, s := range seps {
		pieces = append(pieces, [2]int{start, s[0]})
		start = s[1]
	}
	pieces = append(pieces, [2]int{start, len([]rune(values))})
	return seps, pieces
}

func matchNumber(piece string) (numberPiece, bool) {
	m := numberRe.FindStringSubmatchIndex(piece)
	if m == nil {
		return numberPiece{}, false
	}
	offsets := pytext.RuneOffsets(piece)
	n := numberPiece{left: piece[m[2]:m[3]], leftEnd: offsets[m[3]]}
	if m[4] >= 0 {
		n.hasRange = true
		n.dash, n.right = piece[m[4]:m[5]], piece[m[6]:m[7]]
		n.dashStart, n.dashEnd = offsets[m[4]], offsets[m[5]]
		n.rightStart, n.rightEnd = offsets[m[6]], offsets[m[7]]
	}
	return n, true
}

// numbers renders and interprets a numeric locator value under the switches.
func numbers(raw, basis string, enabled kernels.Switches) (string, []kernels.Range, error) {
	rawRunes := []rune(raw)
	seps, pieces := splitNumbers(raw)
	var out []string
	var meaning []kernels.Range
	for _, p := range pieces {
		piece := string(rawRunes[p[0]:p[1]])
		n, ok := matchNumber(piece)
		if !ok {
			return "", nil, errors.New("unsupported numeric locator")
		}
		lo, err := atoi(n.left)
		if err != nil {
			return "", nil, err
		}
		hi := lo
		rendered := piece
		if n.hasRange {
			if basis == "page" && len(n.left) >= 3 && len(n.right) == 2 {
				hi, err = atoi(n.left[:len(n.left)-2] + n.right)
			} else {
				hi, err = atoi(n.right)
			}
			if err != nil {
				return "", nil, err
			}
			if hi <= lo || (len(n.right) > 2 && strings.HasPrefix(n.right, "0")) {
				return "", nil, errors.New("invalid/descending locator range")
			}
			if basis == "paragraph" && (strings.HasPrefix(n.right, "0") || len(n.right) < len(n.left)) {
				return "", nil, errors.New("contracted paragraph locator is not supported")
			}
			newRight := n.right
			if basis == "page" && enabled.Has("page-contraction") {
				if newRight, err = kernels.ShortenPageRange(n.left, n.right); err != nil {
					return "", nil, err
				}
			}
			// Spaces surrounding the delimiter are owned only by range-dashes.
			if enabled.Has("range-dashes") {
				rendered = n.left + "–" + newRight
			} else {
				rendered = string([]rune(piece)[:n.rightStart]) + newRight
			}
		}
		meaning = append(meaning, kernels.Range{Lo: lo, Hi: hi})
		out = append(out, rendered)
	}
	result := out[0]
	for i, sep := range seps {
		sepText := string(rawRunes[sep[0]:sep[1]])
		if enabled.Has("pinpoint-and") && strings.Contains(sepText, "and") {
			sepText = ", "
		}
		result += sepText + out[i+1]
	}
	return result, meaning, nil
}

// locatorMatch is the LOCATOR production with code-point group offsets.
type locatorMatch struct {
	prefix, label, dot, values string
	prefixEnd, valuesStart     int
}

func matchLocator(raw string) (locatorMatch, bool) {
	m := locatorRe.FindStringSubmatchIndex(raw)
	if m == nil {
		return locatorMatch{}, false
	}
	offsets := pytext.RuneOffsets(raw)
	lm := locatorMatch{prefix: raw[m[2]:m[3]], dot: raw[m[6]:m[7]], values: raw[m[8]:m[9]], prefixEnd: offsets[m[3]], valuesStart: offsets[m[8]]}
	if m[4] >= 0 {
		lm.label = raw[m[4]:m[5]]
	}
	return lm, true
}

// locator parses one locator tail and returns its rendering under the
// switches, its basis and its typed meaning (nil for an empty tail).
func locator(raw, basisHint string, enabled kernels.Switches) (string, string, *Locator, error) {
	if raw == "" {
		return raw, "", nil, nil
	}
	m, ok := matchLocator(raw)
	if !ok {
		return "", "", nil, errors.New("unconsumed suffix: no complete locator")
	}
	var basis string
	switch {
	case strings.HasPrefix(m.label, "para") && m.prefix == " at ":
		basis = "paragraph"
	case (m.label == "p" || m.label == "pp" || m.label == "page" || m.label == "pages") && m.prefix == " at ":
		basis = "page"
	case m.label == "" && m.prefix == " at " && (basisHint == "page" || basisHint == "paragraph"):
		basis = basisHint
	case legalUnits[m.label] != "" && m.prefix == ", ":
		basis = "section"
	default:
		return "", "", nil, errors.New("locator basis not established")
	}
	var rendered, desired string
	var meaning *Locator
	if basis == "section" {
		ids := legalSeparatorRe.Split(m.values, -1)
		for _, id := range ids {
			if !identRe.MatchString(id) {
				return "", "", nil, errors.New("unsupported legal identifier")
			}
		}
		if len(ids) == 0 {
			return "", "", nil, errors.New("unsupported legal identifier")
		}
		meaning = &Locator{Legal: &LegalLocator{Unit: legalUnits[m.label], Identifiers: ids, Separators: legalSeparatorRe.FindAllString(m.values, -1)}}
		rendered = m.values
		desired = m.label
		if len(ids) > 1 {
			desired = legalPlural[m.label]
		}
	} else {
		var ranges []kernels.Range
		var err error
		rendered, ranges, err = numbers(m.values, basis, enabled)
		if err != nil {
			return "", "", nil, err
		}
		meaning = &Locator{Ranges: ranges}
		desired = ""
		if basis != "page" {
			desired = "para"
			if multipleRanges(ranges) {
				desired = "paras"
			}
		}
	}
	if enabled.Has("labels") {
		if desired != "" {
			desired += " "
		}
		return m.prefix + desired + rendered, basis, meaning, nil
	}
	return string([]rune(raw)[:m.valuesStart]) + rendered, basis, meaning, nil
}

func multipleRanges(ranges []kernels.Range) bool {
	if len(ranges) > 1 {
		return true
	}
	for _, r := range ranges {
		if r.Lo != r.Hi {
			return true
		}
	}
	return false
}

// LocatorEdits emits token-specific edits; it never replaces a mixed-format
// locator tail. The full-production parse supplies the shared applicability
// evidence; individual read/write spans identify labels, delimiter glyphs and
// endpoint prefixes, so retained digits are never deleted and reinserted just
// to change a dash. offset is the code-point position of raw inside text.
func LocatorEdits(text string, offset int, raw, basisHint string, enabled kernels.Switches) ([]kernels.Edit, error) {
	target, basis, meaning, err := locator(raw, basisHint, enabled)
	if err != nil {
		return nil, err
	}
	if raw == "" {
		return nil, nil
	}
	m, _ := matchLocator(raw)
	var out []kernels.Edit
	change := func(a, b int, new, rule string) {
		if e := kernels.MinimalEdit(text, offset+a, offset+b, new, rule); e != nil {
			out = append(out, *e)
		}
	}
	if enabled.Has("labels") {
		// Locator values are unchanged for section IDs; numeric desired labels
		// are determined from the already parsed cardinality, not a word regex.
		var desired string
		switch basis {
		case "section":
			desired = m.label
			if len(meaning.Legal.Identifiers) > 1 {
				desired = unitLabels[meaning.Legal.Unit][1]
			}
		case "page":
			desired = ""
		default:
			desired = "para"
			if multipleRanges(meaning.Ranges) {
				desired = "paras"
			}
		}
		if desired != "" {
			desired += " "
		}
		change(m.prefixEnd, m.valuesStart, desired, "CIT-LABEL")
	}
	if basis != "section" {
		values, base := m.values, m.valuesStart
		valueRunes := []rune(values)
		seps, pieces := splitNumbers(values)
		for _, p := range pieces {
			n, _ := matchNumber(string(valueRunes[p[0]:p[1]]))
			at := base + p[0]
			if !n.hasRange {
				continue
			}
			if enabled.Has("range-dashes") {
				change(at+n.dashStart, at+n.dashEnd, "–", "CIT-RANGE-DASH")
				change(at+n.leftEnd, at+n.dashStart, "", "CIT-RANGE-SPACE")
				change(at+n.dashEnd, at+n.rightStart, "", "CIT-RANGE-SPACE")
			}
			if basis == "page" && enabled.Has("page-contraction") {
				shortened, err := kernels.ShortenPageRange(n.left, n.right)
				if err != nil {
					return nil, err
				}
				change(at+n.rightStart, at+n.rightEnd, shortened, "CIT-PAGE-END")
			}
		}
		if enabled.Has("pinpoint-and") {
			for _, sep := range seps {
				if strings.Contains(string(valueRunes[sep[0]:sep[1]]), "and") {
					change(base+sep[0], base+sep[1], ", ", "CIT-PINPOINT-COMMA")
				}
			}
		}
	}
	check := []rune(raw)
	for _, e := range sortedByStartDesc(out) {
		check = append(append(append([]rune(nil), check[:e.Start-offset]...), []rune(e.New)...), check[e.End-offset:]...)
	}
	if string(check) != target {
		return nil, &kernels.InvariantError{Reason: "locator leaf edits disagree with serializer"}
	}
	return out, nil
}

func sortedByStartDesc(edits []kernels.Edit) []kernels.Edit {
	out := append([]kernels.Edit(nil), edits...)
	// Stable descending order by start, mirroring sorted(..., reverse=True).
	for i := 1; i < len(out); i++ {
		for j := i; j > 0 && out[j].Start > out[j-1].Start; j-- {
			out[j], out[j-1] = out[j-1], out[j]
		}
	}
	return out
}
