// Package pytext reproduces the few Python str semantics the alrmap port
// depends on. The reference implementation addresses text by code point and
// classifies characters with Python's Unicode predicates, so the Go port keeps
// rune offsets everywhere and routes every classification through this file
// instead of the ASCII-oriented helpers in regexp and strings.
package pytext

import (
	"strings"
	"unicode"
	"unicode/utf8"
)

// IsWord mirrors kernels._word: str.isalnum(), underscore, or a combining mark.
// It deliberately accepts marks so that a token glued to a combining accent is
// never treated as a free-standing word.
func IsWord(r rune) bool {
	return r >= 0 && (unicode.IsLetter(r) || unicode.IsNumber(r) || r == '_' || unicode.IsMark(r))
}

// IsRegexWord mirrors Python's Unicode regex \w (str.isalnum() or '_'), which
// unlike IsWord excludes combining marks. Lookaround substitutes use it.
func IsRegexWord(r rune) bool {
	return r >= 0 && (unicode.IsLetter(r) || unicode.IsNumber(r) || r == '_')
}

// IsSpace mirrors str.isspace() for one code point (bidi WS/B/S or Zs).
func IsSpace(r rune) bool {
	switch {
	case r >= 0x09 && r <= 0x0D, r >= 0x1C && r <= 0x20:
		return true
	case r == 0x85, r == 0xA0, r == 0x1680, r >= 0x2000 && r <= 0x200A:
		return true
	case r == 0x2028, r == 0x2029, r == 0x202F, r == 0x205F, r == 0x3000:
		return true
	}
	return false
}

// Strip mirrors str.strip() with no arguments.
func Strip(s string) string {
	return strings.TrimFunc(s, IsSpace)
}

func isCased(r rune) bool { return unicode.IsUpper(r) || unicode.IsLower(r) || unicode.IsTitle(r) }

// IsLower mirrors str.islower(): every cased character is lowercase and at
// least one cased character exists.
func IsLower(s string) bool {
	cased := false
	for _, r := range s {
		if unicode.IsUpper(r) || unicode.IsTitle(r) {
			return false
		}
		if unicode.IsLower(r) {
			cased = true
		}
	}
	return cased
}

// IsUpper mirrors str.isupper().
func IsUpper(s string) bool {
	cased := false
	for _, r := range s {
		if unicode.IsLower(r) || unicode.IsTitle(r) {
			return false
		}
		if unicode.IsUpper(r) {
			cased = true
		}
	}
	return cased
}

// IsTitle mirrors str.istitle(): uppercase/titlecase letters only follow
// uncased characters and lowercase letters only follow cased ones.
func IsTitle(s string) bool {
	cased, previousCased := false, false
	for _, r := range s {
		switch {
		case unicode.IsUpper(r) || unicode.IsTitle(r):
			if previousCased {
				return false
			}
			previousCased, cased = true, true
		case unicode.IsLower(r):
			if !previousCased {
				return false
			}
			previousCased, cased = true, true
		default:
			previousCased = false
		}
	}
	return cased
}

// Casefold approximates str.casefold() with simple (one-to-one) case folding.
// Full foldings that change length (ß → ss, ﬁ → fi) are not expanded; only
// equality comparisons use the result, and no ALR fixture relies on them.
func Casefold(s string) string {
	return strings.Map(func(r rune) rune {
		if !isCased(r) {
			return r
		}
		// The smallest member of the SimpleFold orbit is a canonical
		// representative, so every case variant maps to the same rune.
		min := r
		for f := unicode.SimpleFold(r); f != r; f = unicode.SimpleFold(f) {
			if f < min {
				min = f
			}
		}
		return min
	}, s)
}

// Capitalize returns the first code point upper-cased followed by the rest
// unchanged (value[0].upper() + value[1:]).
func Capitalize(s string) string {
	if s == "" {
		return s
	}
	r, size := utf8.DecodeRuneInString(s)
	return string(unicode.ToUpper(r)) + s[size:]
}

// RuneOffsets returns, for every byte offset 0..len(s), the code-point index
// at that byte, so regexp byte positions can be mapped to rune positions.
func RuneOffsets(s string) []int {
	offsets := make([]int, len(s)+1)
	count := 0
	for i := range s {
		offsets[i] = count
		count++
	}
	offsets[len(s)] = count
	// Continuation bytes never start a match; fill them so lookups are total.
	for i := len(s) - 1; i >= 0; i-- {
		if !utf8.RuneStart(s[i]) {
			offsets[i] = offsets[i+1]
		}
	}
	return offsets
}

// RuneAt returns the code point at rune index i, or -1 outside the text.
func RuneAt(runes []rune, i int) rune {
	if i < 0 || i >= len(runes) {
		return -1
	}
	return runes[i]
}

// IsDecimalDigits mirrors str.isdecimal() restricted to ASCII, which is what
// the reference grammars accept for labels and page numbers.
func IsDecimalDigits(s string) bool {
	if s == "" {
		return false
	}
	for i := 0; i < len(s); i++ {
		if s[i] < '0' || s[i] > '9' {
			return false
		}
	}
	return true
}

// CompareDigits compares two ASCII digit strings numerically without any
// magnitude limit, mirroring Python's arbitrary-precision int comparison.
func CompareDigits(a, b string) int {
	a, b = strings.TrimLeft(a, "0"), strings.TrimLeft(b, "0")
	if len(a) != len(b) {
		if len(a) < len(b) {
			return -1
		}
		return 1
	}
	return strings.Compare(a, b)
}
