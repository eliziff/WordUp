package citations

import (
	"errors"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
	"github.com/eliziff/WordUp/internal/alrmap/kernels"
)

// Python's \s in [^<>\s]+ is Unicode whitespace; spelled out for RE2.
var wrappedURIRe = regexp.MustCompile(`<https?://[^<>\t\n\x0B\x0C\r \x1C-\x1F\x{85}\p{Z}]+>`)

// Piece is a top-level component of a note, as code-point offsets.
type Piece struct{ Start, End int }

// SplitTop splits only outside balanced quotations, parentheses and brackets.
func SplitTop(text string, separator rune) ([]Piece, error) {
	runes := []rune(text)
	quotes := kernels.QuotationGuards(text)
	for _, s := range quotes {
		if s.Reason != "delimited-source-or-title" {
			return nil, errors.New("ambiguous quotation syntax")
		}
	}
	covered := make([]bool, len(runes))
	for _, s := range quotes {
		for i := s.Start; i < s.End; i++ {
			covered[i] = true
		}
	}
	offsets := pytext.RuneOffsets(text)
	for _, m := range wrappedURIRe.FindAllStringIndex(text, -1) {
		for i := offsets[m[0]]; i < offsets[m[1]]; i++ {
			covered[i] = true
		}
	}
	var stack []rune
	start := 0
	var pieces []Piece
	for i, ch := range runes {
		if covered[i] {
			continue
		}
		switch ch {
		case '(', '[':
			stack = append(stack, ch)
		case ')', ']':
			want := '('
			if ch == ']' {
				want = '['
			}
			if len(stack) == 0 || stack[len(stack)-1] != want {
				return nil, errors.New("unbalanced citation punctuation")
			}
			stack = stack[:len(stack)-1]
		default:
			if ch == separator && len(stack) == 0 {
				pieces = append(pieces, Piece{start, i})
				start = i + 1
			}
		}
	}
	if len(stack) > 0 {
		return nil, errors.New("unclosed citation punctuation")
	}
	pieces = append(pieces, Piece{start, len(runes)})
	for _, p := range pieces {
		if strings.Trim(string(runes[p.Start:p.End]), " ") == "" {
			return nil, errors.New("empty citation component")
		}
	}
	return pieces, nil
}

// Exact resolves one key through an exact multi-map; a missing key is its own
// canonical value, and a key with several values is ambiguous.
func Exact(key string, mapping map[string][]string) (string, error) {
	values, ok := mapping[key]
	if !ok {
		values = []string{key}
	}
	if len(values) != 1 || values[0] == "" {
		return "", errors.New("ambiguous exact dictionary key")
	}
	return values[0], nil
}

// Symbols returns every recognizable name (keys and canonical outputs),
// longest first, and refuses chained rewrite tables.
func Symbols(mapping map[string][]string) ([]string, error) {
	names := map[string]bool{}
	for key := range mapping {
		names[key] = true
	}
	for key, values := range mapping {
		if key == "" || len(values) == 0 {
			return nil, errors.New("invalid exact dictionary record")
		}
		for _, v := range values {
			if v == "" {
				return nil, errors.New("invalid exact dictionary record")
			}
		}
		for _, value := range values {
			target, ok := mapping[value]
			if ok && !(len(target) == 1 && target[0] == value) {
				return nil, errors.New("dictionary target is not terminal/canonical")
			}
			names[value] = true
		}
	}
	out := make([]string, 0, len(names))
	for n := range names {
		out = append(out, n)
	}
	sortLongestFirst(out)
	return out, nil
}

// sortLongestFirst orders by (-len(code points), name), the reference key.
func sortLongestFirst(names []string) {
	sort.Slice(names, func(i, j int) bool {
		li, lj := len([]rune(names[i])), len([]rune(names[j]))
		if li != lj {
			return li > lj
		}
		return names[i] < names[j]
	})
}

// Date is a validated calendar date.
type Date struct{ Year, Month, Day int }

var (
	dateMDYRe = regexp.MustCompile(`^(?i)(` + kernels.MonthAlternation + `) +([1-9]|[12][0-9]|3[01]),? +([12][0-9]{3})$`)
	dateDMYRe = regexp.MustCompile(`^(?i)([1-9]|[12][0-9]|3[01]) +(` + kernels.MonthAlternation + `) +([12][0-9]{3})$`)
)

// ParseDate accepts a complete named month/day/year in either order and
// renders it day-month-year with the canonical month spelling.
func ParseDate(raw string) (string, Date, error) {
	var mon, day, year int
	if m := dateMDYRe.FindStringSubmatch(raw); m != nil {
		mon = kernels.Months[strings.ToLower(m[1])]
		day, _ = strconv.Atoi(m[2])
		year, _ = strconv.Atoi(m[3])
	} else if m := dateDMYRe.FindStringSubmatch(raw); m != nil {
		day, _ = strconv.Atoi(m[1])
		mon = kernels.Months[strings.ToLower(m[2])]
		year, _ = strconv.Atoi(m[3])
	} else {
		return "", Date{}, errors.New("unsupported complete date")
	}
	if !kernels.ValidDate(year, mon, day) {
		return "", Date{}, errors.New("day is out of range for month")
	}
	canonical := kernels.MonthNames[mon-1]
	return strconv.Itoa(day) + " " + canonical + " " + strconv.Itoa(year), Date{year, mon, day}, nil
}
