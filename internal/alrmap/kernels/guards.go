package kernels

import (
	"regexp"
	"unicode"
	"unicode/utf8"

	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
)

var quoteOpening = map[rune]rune{'“': '”', '‘': '’', '«': '»', '‹': '›', '„': '“', '‚': '‘', '「': '」', '『': '』'}

var quoteClosing = func() map[rune]bool {
	m := map[rune]bool{}
	for _, c := range quoteOpening {
		m[c] = true
	}
	return m
}()

// QuotationGuards is a conservative delimiter scanner; all quoted material is
// protected alike. It recognizes curly double/single quotes and guillemets,
// carries state across paragraphs, treats a repeated opener at the start of a
// quoted paragraph as a continuation, and quarantines the entire logical
// container on malformed delimiters. More than one non-apostrophe straight
// quote pair also quarantines the container: a toggle parser cannot tell
// successive quotations from unescaped nesting. It does not identify unmarked
// quotations; paragraph/style/source guards are additional inputs.
func QuotationGuards(text string) []Span {
	return quotationGuards([]rune(text))
}

func quotationGuards(text []rune) []Span {
	n := len(text)
	singles, doubles := 0, 0
	for i, ch := range text {
		if ch == '"' {
			doubles++
		}
		if ch == '\'' && !(pytext.IsWord(pytext.RuneAt(text, i-1)) && pytext.IsWord(pytext.RuneAt(text, i+1))) {
			singles++
		}
	}
	if doubles > 2 || singles > 2 {
		return []Span{{0, n, "ambiguous-straight-quotation-system"}}
	}
	type frame struct {
		opener, closer rune
		start          int
	}
	var stack []frame
	var spans []Span
	for i, ch := range text {
		before, after := pytext.RuneAt(text, i-1), pytext.RuneAt(text, i+1)
		if (ch == '\'' || ch == '’') && pytext.IsWord(before) && pytext.IsWord(after) {
			continue
		}
		if len(stack) > 0 && ch == stack[len(stack)-1].closer {
			top := stack[len(stack)-1]
			stack = stack[:len(stack)-1]
			spans = append(spans, Span{top.start, i + 1, "delimited-source-or-title"})
			continue
		}
		switch {
		case ch == '"' || ch == '\'':
			if len(stack) > 0 && stack[len(stack)-1].opener == ch {
				top := stack[len(stack)-1]
				stack = stack[:len(stack)-1]
				spans = append(spans, Span{top.start, i + 1, "delimited-source-or-title"})
			} else {
				stack = append(stack, frame{ch, ch, i})
			}
		case quoteOpening[ch] != 0:
			if len(stack) > 0 && ch == stack[len(stack)-1].opener {
				pstart := 0
				for j := i - 1; j >= 0; j-- {
					if text[j] == '\n' || text[j] == '\r' {
						pstart = j + 1
						break
					}
				}
				if pstart > stack[len(stack)-1].start && pytext.Strip(string(text[pstart:i])) == "" {
					continue
				}
				return []Span{{0, n, "ambiguous-same-level-quote"}}
			}
			stack = append(stack, frame{ch, quoteOpening[ch], i})
		case quoteClosing[ch]:
			return []Span{{0, n, "unmatched-closing-quote"}}
		case unicode.Is(unicode.Pi, ch) || unicode.Is(unicode.Pf, ch):
			return []Span{{0, n, "unsupported-quote-delimiter"}}
		}
	}
	if len(stack) > 0 {
		return []Span{{0, n, "unclosed-quote"}}
	}
	return SortSpans(spans)
}

// notURIChar is the reference [^\s<>\[\]"“”‘’] with Python's Unicode \s
// spelled out, because RE2's \s is ASCII-only.
const notURIChar = `[^\t\n\x0B\x0C\r \x1C-\x1F\x{85}\p{Z}<>\[\]"“”‘’]`

// The reference OPAQUE pattern has a lookbehind (?<![\w@]) on its second
// alternative (a naked host with a path). RE2 has no lookbehind, so that
// alternative is captured separately and opaqueMatches re-applies the
// assertion by hand. The e-mail alternative needs an '@' before any '/', so
// it can never match at a position where the host alternative matched; the
// only Python behaviour left to reproduce is resuming one code point later.
var opaqueRe = regexp.MustCompile(`((?:https?://|ftp://|www\.)` + notURIChar + `+)` +
	`|((?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,24}/` + notURIChar + `+)` +
	`|([\p{L}\p{N}_.+-]+@(?:[\p{L}\p{N}_-]+\.)+[A-Za-z]{2,24})`)

// opaqueMatches returns byte ranges of URL/e-mail tokens in scan order.
func opaqueMatches(text string) [][2]int {
	var out [][2]int
	pos := 0
	for pos <= len(text) {
		m := opaqueRe.FindStringSubmatchIndex(text[pos:])
		if m == nil {
			break
		}
		start, end := pos+m[0], pos+m[1]
		if m[4] >= 0 && start > 0 {
			prev, _ := utf8.DecodeLastRuneInString(text[:start])
			if pytext.IsRegexWord(prev) || prev == '@' {
				// Lookbehind failure: Python moves one code point right
				// rather than past the candidate, so a host inside it
				// (after a dot) can still be found.
				_, size := utf8.DecodeRuneInString(text[start:])
				pos = start + size
				continue
			}
		}
		out = append(out, [2]int{start, end})
		pos = end
	}
	return out
}

// Protection combines validated caller spans, quotation guards, URL/e-mail
// masks and bracketed short-form/qualification masks. Lexical operations never
// edit inside any of them.
func Protection(text string, extra []Span) ([]Span, error) {
	return protection([]rune(text), text, extra)
}

func protection(runes []rune, text string, extra []Span) ([]Span, error) {
	guards, err := validateSpans(len(runes), extra)
	if err != nil {
		return nil, err
	}
	guards = append(guards, quotationGuards(runes)...)
	if len(text) > 0 {
		offsets := pytext.RuneOffsets(text)
		for _, m := range opaqueMatches(text) {
			guards = append(guards, Span{offsets[m[0]], offsets[m[1]], "url-or-email"})
		}
	}
	depth, start := 0, 0
	unmatched := false
	for i, ch := range runes {
		if ch == '[' {
			if depth == 0 {
				start = i
			}
			depth++
		} else if ch == ']' {
			if depth == 0 {
				guards = append(guards, Span{0, len(runes), "unmatched-bracket"})
				unmatched = true
				break
			}
			depth--
			if depth == 0 {
				guards = append(guards, Span{start, i + 1, "bracketed-content"})
			}
		}
	}
	if depth > 0 && !unmatched {
		guards = append(guards, Span{0, len(runes), "unclosed-bracket"})
	}
	return SortSpans(guards), nil
}
