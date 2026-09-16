package kernels

import (
	"errors"
	"regexp"
	"strings"

	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
)

// LegalIdent is the opaque statutory identifier grammar shared by the
// citation and rule serializers. A hyphen inside it is identity, never a range.
const LegalIdent = `[1-9][0-9]*(?:\.[0-9]+)*(?:-[0-9]+)*(?:\([A-Za-z0-9]+\))*`

var (
	decimalStartRe = regexp.MustCompile(`^[1-9][0-9]*$`)
	decimalRe      = regexp.MustCompile(`^[0-9]+$`)
	locatorPieceRe = regexp.MustCompile(`^ *([1-9][0-9]*)(?: *([-–]) *([0-9]+))? *$`)
	// (?= |,|$) after the note number is replaced by a consuming group whose
	// end is ignored; with nothing following it, the chosen match is identical.
	supraRe        = regexp.MustCompile(`^(.+), supra note ([1-9][0-9]*)( |,|$)`)
	atLabelRe      = regexp.MustCompile(`^ at (paras?|pp?|pages?)\.? +(.+)$`)
	atBareRe       = regexp.MustCompile(`^ at (.+)$`)
	sectionRe      = regexp.MustCompile(`^, (s|ss|art|arts)\.? +(.+)$`)
	sectionIdentRe = regexp.MustCompile(`^` + LegalIdent + `(?: to ` + LegalIdent + `)?$`)
)

// ShortenPageRange implements only the supplied two-final-digit/same-prefix
// contraction; no broader McGill inference.
func ShortenPageRange(start, end string) (string, error) {
	if !decimalStartRe.MatchString(start) || !decimalRe.MatchString(end) {
		return "", errors.New("not decimal endpoint")
	}
	// Already shortened endpoints can remain, never expand an uncertain range.
	if len(start) >= 3 && len(end) == 2 {
		if pytext.CompareDigits(start[:len(start)-2]+end, start) > 0 {
			return end, nil
		}
		return "", errors.New("descending or ambiguous shortened page range")
	}
	if strings.HasPrefix(end, "0") || pytext.CompareDigits(end, start) <= 0 {
		return "", errors.New("nonascending page range")
	}
	if len(start) >= 3 && len(start) == len(end) && start[:len(start)-2] == end[:len(end)-2] {
		result := end[len(end)-2:]
		if pytext.CompareDigits(start[:len(start)-2]+result, end) != 0 {
			return "", &InvariantError{"round-trip failure"}
		}
		return result, nil
	}
	return end, nil
}

// LocatorList renders a finite list of decimal endpoints/ranges. No section
// IDs or years are classified here. The second result reports plurality.
func LocatorList(raw string, pages bool) (string, bool, error) {
	var rendered []string
	multiple := false
	pieces := strings.Split(raw, ",")
	for _, piece := range pieces {
		m := locatorPieceRe.FindStringSubmatch(piece)
		if m == nil {
			return "", false, errors.New("unknown locator syntax")
		}
		a, separator, b := m[1], m[2], m[3]
		if separator != "" {
			if pages {
				var err error
				if b, err = ShortenPageRange(a, b); err != nil {
					return "", false, err
				}
			} else if strings.HasPrefix(b, "0") || pytext.CompareDigits(b, a) <= 0 {
				return "", false, errors.New("ambiguous paragraph endpoints")
			}
			rendered = append(rendered, a+"–"+b)
			multiple = true
		} else {
			rendered = append(rendered, a)
		}
	}
	multiple = multiple || len(pieces) > 1
	return strings.Join(rendered, ", "), multiple, nil
}

// TokenFormat is a formatting intention for a syntax word (e.g. italic Ibid).
type TokenFormat struct {
	Start, End int
	Format     string
}

// CitationResult is the output of the bounded short-citation kernel.
type CitationResult struct {
	Text        string
	Kind        string
	TokenFormat []TokenFormat
}

// ShortCitationOptions are the keyword arguments of normalize_short_citation.
type ShortCitationOptions struct {
	Aliases         map[string]bool
	BareAtMeansPage bool
}

// NormalizeShortCitation is a whole-production parser: Ibid or exact alias +
// supra note, then one locator. It does not parse a citation hidden in a
// quoted passage, arbitrary prose, a full case name, multi-citation notes,
// qualifications, fields or titles. Exact aliases establish syntax only.
func NormalizeShortCitation(text string, opts ShortCitationOptions) (CitationResult, error) {
	runes := []rune(text)
	if len(quotationGuards(runes)) > 0 || strings.ContainsAny(text, "\r\n\x02\x13\x14\x15") {
		return CitationResult{}, errors.New("quoted, multiline or structural input")
	}
	if !strings.HasSuffix(text, ".") || strings.HasSuffix(text, "..") {
		return CitationResult{}, errors.New("whole citation terminal required")
	}
	s := runes[:len(runes)-1]
	var head, tail string
	var fmt []TokenFormat
	if len(s) >= 4 && string(s[:4]) == "Ibid" && (len(s) == 4 || s[4] == ' ' || s[4] == ',') {
		head, tail = "Ibid", string(s[4:])
		fmt = []TokenFormat{{0, 4, "italic"}}
	} else {
		body := string(s)
		m := supraRe.FindStringSubmatchIndex(body)
		if m == nil || !opts.Aliases[body[m[2]:m[3]]] {
			return CitationResult{}, errors.New("no complete supported citation root")
		}
		head, tail = body[:m[5]], body[m[5]:]
		i := len([]rune(head[:strings.Index(head, "supra")]))
		fmt = []TokenFormat{{i, i + 5, "italic"}}
	}
	if tail == "" {
		return CitationResult{head + ".", "source-only", fmt}, nil
	}
	if m := atLabelRe.FindStringSubmatch(tail); m != nil {
		label, raw := m[1], m[2]
		pages := !strings.HasPrefix(label, "para")
		normalized, multiple, err := LocatorList(raw, pages)
		if err != nil {
			return CitationResult{}, err
		}
		kind := "page"
		label = ""
		if !pages {
			kind = "paragraph"
			label = "para "
			if multiple {
				label = "paras "
			}
		}
		return CitationResult{head + " at " + label + normalized + ".", kind, fmt}, nil
	}
	if m := atBareRe.FindStringSubmatch(tail); m != nil && opts.BareAtMeansPage {
		normalized, _, err := LocatorList(m[1], true)
		if err != nil {
			return CitationResult{}, err
		}
		return CitationResult{head + " at " + normalized + ".", "page", fmt}, nil
	}
	if m := sectionRe.FindStringSubmatch(tail); m != nil {
		label, raw := m[1], m[2]
		if !sectionIdentRe.MatchString(raw) {
			return CitationResult{}, errors.New("unsupported section identifier/list")
		}
		// A hyphen is part of the identifier. It is NEVER range evidence here.
		if strings.Contains(raw, " to ") {
			if strings.HasPrefix(label, "s") {
				label = "ss"
			} else {
				label = "arts"
			}
		}
		return CitationResult{head + ", " + label + " " + raw + ".", "section-opaque-identifier", fmt}, nil
	}
	return CitationResult{}, errors.New("unconsumed citation suffix / unknown locator type")
}
