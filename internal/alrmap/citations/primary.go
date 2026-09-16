package citations

import (
	"errors"
	"regexp"
	"strings"
	"sync"

	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
)

// The reference roots end in lookaheads ((?= |,|$), (?= |$), (?=,| |$)).
// RE2 has none, so each is replaced by a consuming group whose extent is
// ignored: with nothing after it in the pattern the chosen match is the same.
var (
	ibidRe            = regexp.MustCompile(`^(Ibid)( |,|$)`)
	supraRe           = regexp.MustCompile(`^(?:(.+), )?([Ss]upra) note ([1-9][0-9]*)( |,|$)`)
	caseRe            = regexp.MustCompile(`^([12][0-9]{3}) (CanLII|[A-Z][A-Z0-9]+) ([1-9][0-9]*)( |,|$)`)
	numericPeriodRe   = regexp.MustCompile(`\p{Nd}\.\p{Nd}|\.\.`)
	statuteRe         = regexp.MustCompile(`^((?:RSC|SC|RSA|SA|RSBC|SBC|RSO|SO|RSM|SM|RSNB|SNB|RSNS|SNS|RSPEI|SPEI|RSNL|SNL|RSS|SS)) ([12][0-9]{3}), c ([A-Za-z0-9]+(?:[-.][A-Za-z0-9]+)*)(,| |$)`)
	regulationRe      = regexp.MustCompile(`^(SOR|SI)/([0-9]{2}|[12][0-9]{3})-([1-9][0-9]*)(,| |$)`)
	reporterPatterns  = map[string]*regexp.Regexp{}
	reporterPatternMu sync.Mutex
)

func reporterPattern(reporter string) *regexp.Regexp {
	reporterPatternMu.Lock()
	defer reporterPatternMu.Unlock()
	if re, ok := reporterPatterns[reporter]; ok {
		return re
	}
	re := regexp.MustCompile(`^(?:([12][0-9]{3}), )?(?:\[([12][0-9]{3})\] )?(?:([1-9][0-9]*) )?` + regexp.QuoteMeta(reporter) +
		`(?: \(([1-9][0-9]*(?:d|st|nd|rd|th))\))? ([1-9][0-9]*)( |$)`)
	reporterPatterns[reporter] = re
	return re
}

// group returns the text of capture group i, or "" and false when unset.
func group(s string, m []int, i int) (string, bool) {
	if m[2*i] < 0 {
		return "", false
	}
	return s[m[2*i]:m[2*i+1]], true
}

// parsePrimary tries the short, case, reporter and legislation roots. A
// failed prefix owns no tokens or edits.
func parsePrimary(c *component) (root, error) {
	body := c.body
	off := c.bodyOffsets
	var r root
	// Short roots, including an unlabelled supra pointer. No fuzzy alias search.
	if m := ibidRe.FindStringSubmatchIndex(body); m != nil {
		r.kind = "ibid"
		r.headEnd = off[m[3]]
		c.token("ibid", 0, r.headEnd)
		return r, nil
	}
	if m := supraRe.FindStringSubmatchIndex(body); m != nil {
		alias, hasAlias := group(body, m, 1)
		r.alias, r.noteLabel = alias, body[m[6]:m[7]]
		if hasAlias && !c.registry.Aliases[alias] {
			return root{}, errors.New("supra alias is not declared")
		}
		r.kind = "supra"
		r.headEnd = off[m[7]]
		c.token("supra", off[m[4]], off[m[5]])
		c.token("note-label", off[m[6]], off[m[7]])
		return r, nil
	}

	// Bare or exact-name case citation. Unknown names are not guessed.
	names, err := Symbols(c.registry.Cases)
	if err != nil {
		return root{}, err
	}
	caseStart, matchedName, hasName := 0, "", false
	for _, name := range names {
		if strings.HasPrefix(body, name+", ") {
			caseStart = len([]rune(name)) + 2
			matchedName, hasName = name, true
			break
		}
	}
	rest := string(c.bodyRunes[caseStart:])
	if m := caseRe.FindStringSubmatchIndex(rest); m != nil && (rest[m[4]:m[5]] == "CanLII" || Courts[rest[m[4]:m[5]]]) {
		if hasName {
			canonical, err := Exact(matchedName, c.registry.Cases)
			if err != nil {
				return root{}, err
			}
			if canonical != matchedName && (canonical != strings.ReplaceAll(matchedName, ".", "") || numericPeriodRe.MatchString(matchedName)) {
				return root{}, errors.New("case-name formatter may only delete nonnumeric abbreviation periods")
			}
			c.token("case-name", 0, len([]rune(matchedName)))
			if c.enabled.Has("case-names") {
				c.change(c.offset, c.offset+len([]rune(matchedName)), canonical, "CIT-CASE-NAME")
			}
		}
		r.kind = "case"
		r.source = []string{"case", rest[m[2]:m[3]], rest[m[4]:m[5]], rest[m[6]:m[7]]}
		r.headEnd = caseStart + pytext.RuneOffsets(rest)[m[7]]
		c.token("case-identifier", caseStart, r.headEnd)
		r.basisHint = c.registry.SourceBasis[SourceKey(r.source)]
		// A neutral citation does not justify inventing a missing locator label.
		return r, nil
	}

	// Reporter identifiers are exact supplied tokens. Preserve both year
	// fields; never equate a reporter case with a neutral citation by name.
	reporters, err := Symbols(c.registry.Reporters)
	if err != nil {
		return root{}, err
	}
	for _, reporter := range reporters {
		m := reporterPattern(reporter).FindStringSubmatchIndex(body)
		if m == nil {
			continue
		}
		decision, _ := group(body, m, 1)
		year, _ := group(body, m, 2)
		if decision == "" && year == "" {
			return root{}, errors.New("reporter grammar requires an explicit year")
		}
		canonical, err := Exact(reporter, c.registry.Reporters)
		if err != nil {
			return root{}, err
		}
		volume, _ := group(body, m, 3)
		series, hasSeries := group(body, m, 4)
		page := body[m[10]:m[11]]
		r.kind = "reporter"
		r.headEnd = off[m[11]]
		r.source = []string{"reporter", decision, year, volume, canonical, series, page}
		// Locate from parsed prefix length, not a fresh document Find.
		endPrefix := off[m[10]] - 1
		if hasSeries {
			endPrefix = off[m[8]] - 2
		}
		startReporter := endPrefix - len([]rune(reporter))
		c.token("reporter", startReporter, endPrefix)
		if c.enabled.Has("reporter-names") {
			c.change(c.offset+startReporter, c.offset+endPrefix, canonical, "CIT-REPORTER")
		}
		return r, nil
	}

	// Narrow Canadian statute/regulation productions, named only through an
	// exact supplied title set. Chapter/decision hyphens are opaque identity.
	pos, matchedTitle, hasTitle := 0, "", false
	titles := make([]string, 0, len(c.registry.Statutes))
	for t := range c.registry.Statutes {
		titles = append(titles, t)
	}
	sortLongestFirst(titles)
	for _, title := range titles {
		if strings.HasPrefix(body, title+", ") {
			pos = len([]rune(title)) + 2
			matchedTitle, hasTitle = title, true
			break
		}
	}
	rest = string(c.bodyRunes[pos:])
	if m := statuteRe.FindStringSubmatchIndex(rest); m != nil {
		r.kind = "legislation"
		r.headEnd = pos + pytext.RuneOffsets(rest)[m[7]]
		r.source = []string{"legislation", rest[m[2]:m[3]], rest[m[4]:m[5]], rest[m[6]:m[7]]}
		c.token("statute-identifier", pos, r.headEnd)
	} else if m := regulationRe.FindStringSubmatchIndex(rest); m != nil {
		r.kind = "legislation"
		r.headEnd = pos + pytext.RuneOffsets(rest)[m[7]]
		r.source = []string{"regulation", rest[m[2]:m[3]], rest[m[4]:m[5]], rest[m[6]:m[7]]}
		c.token("statute-identifier", pos, r.headEnd)
	}
	if r.kind == "legislation" && hasTitle {
		c.token("source-title", 0, len([]rune(matchedTitle)))
	}
	return r, nil
}
