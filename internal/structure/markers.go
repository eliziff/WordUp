package structure

import (
	"regexp"
	"strconv"
	"strings"
)

var namedHeadingPrefix = regexp.MustCompile(`^\s*(?i:part|chapter)\s+([IVXLCDM]{1,7}|[0-9]{1,3}|(?i:one|two|three|four|five|six|seven|eight|nine|ten))\s*([:\-\x{2013}\x{2014}])\s*(.+)$`)

var headingPrefix = regexp.MustCompile(`^\s*(?:(?i:part|chapter)\s+)?([IVXLCDM]{1,7}|[A-Za-z]|[0-9]{1,3})([.)]|\s*[-\x{2013}\x{2014}])\s+(.+)$`)

// MarkerChoices retains ambiguous Roman/letter interpretations. It recognizes
// candidates only; prose, lists and quoted instruments can have these prefixes.
func MarkerChoices(text string) []Interpretation {
	parts := namedHeadingPrefix.FindStringSubmatch(text)
	named := parts != nil
	if !named {
		parts = headingPrefix.FindStringSubmatch(text)
	}
	if parts == nil {
		return nil
	}
	value, punct := parts[1], strings.TrimSpace(parts[2])
	if named {
		punct = "named_section"
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
	for i := len(value) - 1; i >= 0; i-- {
		n := map[byte]int{'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}[value[i]]
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
