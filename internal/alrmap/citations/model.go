// Package citations parses complete logical footnotes into explicit citation
// records under ten independently owned feature switches, and emits
// occurrence-addressed leaf patches. Every offset is a code-point index into
// the note text, as in the Python reference; nothing here mutates Word.
package citations

import (
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/eliziff/WordUp/internal/alrmap/kernels"
)

// AllSwitches lists the ten independent mutation switches.
var AllSwitches = []string{"labels", "range-dashes", "page-contraction", "pinpoint-and",
	"court-order", "journal-names", "case-names", "reporter-names", "dates", "terminal"}

var switchSet = kernels.NewSwitches(AllSwitches...)

// Courts is the enumerated neutral-citation court token set.
var Courts = func() map[string]bool {
	m := map[string]bool{}
	for _, c := range strings.Fields("SCC FC FCA TCC ABCA ABQB ABKB ABPC ABCJ BCCA BCSC BCPC MBCA MBQB MBKB " +
		"NBCA NBQB NBKB NLCA NLSC NSCA NSSC NUCJ NWTCA NWTSC ONCA ONSC ONCJ " +
		"PECA PESC QCCA QCCS QCCQ SKCA SKQB SKKB YKCA YKSC") {
		m[c] = true
	}
	return m
}()

// courtAlternation is the sorted court alternation used by the court regexes.
var courtAlternation = func() string {
	names := make([]string, 0, len(Courts))
	for c := range Courts {
		names = append(names, c)
	}
	sort.Strings(names)
	return strings.Join(names, "|")
}()

// Signals are the recognized introductory signals, longest variants first.
var Signals = []string{"See especially", "See generally", "See also", "See e.g.", "But see", "See", "Cf", "Contra"}

// Qualifications is the exact set of known bracketed qualification flags.
var Qualifications = map[string]bool{"emphasis added": true, "emphasis in original": true, "footnotes omitted": true,
	"citations omitted": true, "emphasis omitted": true, "translated by author": true}

// Ident is the opaque legal identifier grammar.
const Ident = kernels.LegalIdent

var (
	locatorRe = regexp.MustCompile(`^( at |, )(paras?|pp?|pages?|ss?|arts?|rules?)?(\.?) *(.+)$`)
	numberRe  = regexp.MustCompile(`^([1-9][0-9]*)(?: *([-–]) *([0-9]+))?$`)
	identRe   = regexp.MustCompile(`^` + Ident + `$`)
)

// Token is an owned syntax token inside the note text.
type Token struct {
	Kind       string
	Start, End int
	Value      string
}

// LegalLocator retains unit, opaque identifiers and the original separators,
// so a range and a list stay different.
type LegalLocator struct {
	Unit        string
	Identifiers []string
	Separators  []string
}

// Locator is a typed locator: numeric ranges for page/paragraph bases or a
// legal identifier list for sections. A nil *Locator means none.
type Locator struct {
	Ranges []kernels.Range
	Legal  *LegalLocator
}

// String renders a canonical, comparable form of the locator ("" for nil).
func (l *Locator) String() string {
	if l == nil {
		return ""
	}
	if l.Legal != nil {
		return "legal(" + l.Legal.Unit + ";" + strings.Join(l.Legal.Identifiers, "\x1f") + ";" + strings.Join(l.Legal.Separators, "\x1f") + ")"
	}
	parts := make([]string, len(l.Ranges))
	for i, r := range l.Ranges {
		parts[i] = strconv.Itoa(r.Lo) + "-" + strconv.Itoa(r.Hi)
	}
	return "ranges(" + strings.Join(parts, ",") + ")"
}

// SourceKey is the comparable identity of a parsed source tuple.
func SourceKey(source []string) string {
	if source == nil {
		return ""
	}
	return "(" + strings.Join(source, "\x1f") + ")"
}

// Citation is one fully parsed citation component.
type Citation struct {
	Kind           string
	Source         []string // nil for roots without a source identity
	Alias          string
	NoteLabel      string
	Basis          string
	Locator        *Locator
	Qualifications []string
	DeclaredAlias  string
	Tokens         []Token
}

// Meaning is the comparable parsed meaning (kind, source, alias, note label,
// basis, locator, qualifications, declared alias), independent of tokens.
func (c Citation) Meaning() string {
	return strings.Join([]string{c.Kind, SourceKey(c.Source), c.Alias, c.NoteLabel, c.Basis,
		c.Locator.String(), strings.Join(c.Qualifications, "\x1f"), c.DeclaredAlias}, "\x1e")
}

// Equal reports full record equality, tokens included.
func (c Citation) Equal(o Citation) bool {
	if c.Meaning() != o.Meaning() || len(c.Tokens) != len(o.Tokens) {
		return false
	}
	for i := range c.Tokens {
		if c.Tokens[i] != o.Tokens[i] {
			return false
		}
	}
	return true
}

// Result is the canonical text, its occurrence-addressed edits and citations.
type Result struct {
	Text      string
	Edits     []kernels.Edit
	Citations []Citation
}

// Registry holds exact supplied dictionaries. Keys/values are exact strings;
// a value list with more than one entry is ambiguous.
type Registry struct {
	Journals  map[string][]string
	Cases     map[string][]string
	Reporters map[string][]string
	Statutes  map[string]bool
	Aliases   map[string]bool
	// Books require an exact title boundary; italic-looking text is not enough.
	Books map[string]bool
	// SourceBasis maps SourceKey(source) to a locator basis.
	SourceBasis map[string]string
}

// root is the outcome of one root alternative.
type root struct {
	kind      string
	source    []string
	alias     string
	noteLabel string
	headEnd   int
	basisHint string
}
