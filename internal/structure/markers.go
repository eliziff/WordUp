package structure

import (
	"regexp"
	"strconv"
	"strings"
	"unicode"
)

var namedHeadingPrefix = regexp.MustCompile(`^\s*((?i:part|chapter|theme|section|article|appendix|schedule|division|book|title))\s+((?i:[IVXLCDM]){1,7}|[A-Za-z]|[0-9]{1,3}|(?i:one|two|three|four|five|six|seven|eight|nine|ten))\s*([:\-\x{2013}\x{2014}])\s*(.+)$`)
var namedPeriodHeadingPrefix = regexp.MustCompile(`^\s*((?i:theme|section|article|appendix|schedule|division|book|title))\s+((?i:[IVXLCDM]){1,7}|[A-Za-z]|[0-9]{1,3}|(?i:one|two|three|four|five|six|seven|eight|nine|ten))\s*([.)])\s+(.+)$`)

var headingPrefix = regexp.MustCompile(`^\s*(?:(?i:part|chapter)\s+)?((?i:[IVXLCDM]){1,7}|[A-Za-z]|[0-9]{1,3})([.)]|\s*[-\x{2013}\x{2014}])\s+(.+)$`)
var parenthesizedHeadingPrefix = regexp.MustCompile(`^\s*\(((?i:[IVXLCDM]){1,7}|[A-Za-z]|[0-9]{1,3})\)\s+(.+)$`)

// MarkerChoices retains ambiguous Roman/letter interpretations. It recognizes
// candidates only; prose, lists and quoted instruments can have these prefixes.
func MarkerChoices(text string) []Interpretation {
	// Word frequently emits non-breaking or narrow spaces when a heading is
	// pasted from a PDF or publisher template. The grammar is intentionally
	// ASCII-shaped, so normalize Unicode whitespace at this boundary while
	// leaving the paragraph's original text untouched in the evidence row.
	text = normalizeMarkerWhitespace(text)
	parts := namedHeadingPrefix.FindStringSubmatch(text)
	named := parts != nil
	parenthesized := false
	if !named {
		parts = namedPeriodHeadingPrefix.FindStringSubmatch(text)
		named = parts != nil
	}
	if !named {
		parts = headingPrefix.FindStringSubmatch(text)
		if parts == nil {
			parts = parenthesizedHeadingPrefix.FindStringSubmatch(text)
			parenthesized = parts != nil
		}
	}
	if parts == nil {
		return nil
	}
	namedPrefix, value, punct := "", parts[1], strings.TrimSpace(parts[2])
	if named {
		namedPrefix, value, punct = strings.ToLower(parts[1]), parts[2], strings.TrimSpace(parts[3])
		// Keep each named prefix in its own ladder family. A document can
		// legitimately nest Part, Chapter, and Appendix sequences; merging
		// them makes a restart look like a malformed counter.
		punct = namedPrefix + "_named_section"
	} else if parenthesized {
		punct = ")"
	}
	if n, err := strconv.Atoi(value); err == nil {
		if n > 0 {
			return []Interpretation{{"numeric_" + punct, n}}
		}
		return nil
	}
	if n := map[string]int{"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10}[strings.ToLower(value)]; n > 0 {
		return []Interpretation{{"numeric_" + punct, n}}
	}
	out := []Interpretation{}
	roman := true
	total, prior := 0, 0
	romanValue := strings.ToUpper(value)
	for i := len(value) - 1; i >= 0; i-- {
		n := map[byte]int{'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}[romanValue[i]]
		if n == 0 {
			roman = false
			break
		}
		if n < prior {
			total -= n
		} else {
			total += n
			prior = n
		}
	}
	if roman && total > 0 {
		out = append(out, Interpretation{"roman_" + punct, total})
	}
	if len(value) == 1 {
		c := value[0]
		if c >= 'A' && c <= 'Z' {
			out = append(out, Interpretation{"upper_alpha_" + punct, int(c-'A') + 1})
		}
		if c >= 'a' && c <= 'z' {
			out = append(out, Interpretation{"lower_alpha_" + punct, int(c-'a') + 1})
		}
	}
	return out
}

func normalizeMarkerWhitespace(text string) string {
	var out strings.Builder
	out.Grow(len(text))
	space := false
	for _, r := range text {
		if unicode.IsSpace(r) {
			if !space {
				out.WriteByte(' ')
				space = true
			}
			continue
		}
		out.WriteRune(r)
		space = false
	}
	return strings.TrimSpace(out.String())
}
