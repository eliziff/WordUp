// SPDX-License-Identifier: GPL-3.0-or-later
// Go adaptation of Rubberduck's VBABaseParser/VBABaseLexer predicates.
package generated

import (
	"github.com/antlr4-go/antlr/v4"
	"regexp"
	"strings"
)

func (p *VBAParser) TokenAtRelativePosition(i int) antlr.Token { return p.GetTokenStream().LT(i) }
func (p *VBAParser) TokenTypeAtRelativePosition(i int) int     { return p.GetTokenStream().LA(i) }
func (p *VBAParser) TextOf(t antlr.Token) string {
	if t == nil {
		return ""
	}
	return t.GetText()
}
func (p *VBAParser) MatchesRegex(s, pattern string) bool {
	ok, _ := regexp.MatchString(pattern, s)
	return ok
}
func (p *VBAParser) EqualsString(a string, options ...string) bool {
	for _, b := range options {
		if a == b {
			return true
		}
	}
	return false
}
func (p *VBAParser) EqualsStringIgnoringCase(a string, options ...string) bool {
	for _, b := range options {
		if strings.EqualFold(a, b) {
			return true
		}
	}
	return false
}
func (p *VBAParser) IsTokenType(a int, options ...int) bool {
	for _, b := range options {
		if a == b {
			return true
		}
	}
	return false
}
func (l *VBALexer) CharAtRelativePosition(i int) int { return l.GetInputStream().LA(i) }
func (l *VBALexer) IsChar(a int, options ...rune) bool {
	for _, b := range options {
		if a == int(b) {
			return true
		}
	}
	return false
}
