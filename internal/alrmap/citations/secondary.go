package citations

import (
	"fmt"
	"regexp"
	"sync"
)

var (
	journalPatterns   = map[string]*regexp.Regexp{}
	bookPatterns      = map[string]*regexp.Regexp{}
	secondaryPatterns sync.Mutex
	// Python's \s inside the wrapped URI is Unicode whitespace; spelled out.
	webRe = regexp.MustCompile(`^([^;“”]+), (“[^”]+”) \(([^()]*)\), online(?: \((blog|pdf|video|podcast|channel)\))?: ` +
		`(<https?://[^<>\t\n\x0B\x0C\r \x1C-\x1F\x{85}\p{Z}]+>|\[perma\.cc/[A-Z0-9]{4}-[A-Z0-9]{4}\])$`)
)

func journalPattern(journal string) *regexp.Regexp {
	secondaryPatterns.Lock()
	defer secondaryPatterns.Unlock()
	if re, ok := journalPatterns[journal]; ok {
		return re
	}
	re := regexp.MustCompile(`^([^“”";]+), (“[^”]+”|"[^"]+") \(([12][0-9]{3})\) ([1-9][0-9]*)` +
		`(?::([1-9][0-9]*(?:[&–-][1-9][0-9]*)?))? ` + regexp.QuoteMeta(journal) + ` ([1-9][0-9]*)( |$)`)
	journalPatterns[journal] = re
	return re
}

func bookPattern(title string) *regexp.Regexp {
	secondaryPatterns.Lock()
	defer secondaryPatterns.Unlock()
	if re, ok := bookPatterns[title]; ok {
		return re
	}
	re := regexp.MustCompile(`^([^;]+), (` + regexp.QuoteMeta(title) + `)(?:, ([1-9][0-9]*(?:st|nd|rd|th) ed))? ` +
		`\(([^():;]+): ([^();]+), ([12][0-9]{3})\)( |$)`)
	bookPatterns[title] = re
	return re
}

// parseSecondary tries the journal, book and web roots. A failed prefix owns
// no tokens or edits.
func parseSecondary(c *component) (root, error) {
	body := c.body
	off := c.bodyOffsets
	var r root
	// Full journal citation: arbitrary title words remain opaque. The journal
	// key must be present in a supplied exact dictionary (even identity entries).
	journals, err := Symbols(c.registry.Journals)
	if err != nil {
		return root{}, err
	}
	for _, journal := range journals {
		m := journalPattern(journal).FindStringSubmatchIndex(body)
		if m == nil {
			continue
		}
		canonical, err := Exact(journal, c.registry.Journals)
		if err != nil {
			return root{}, err
		}
		issue, _ := group(body, m, 5)
		r.source = []string{"journal", body[m[2]:m[3]], body[m[4]:m[5]], body[m[6]:m[7]], body[m[8]:m[9]], issue, canonical, body[m[12]:m[13]]}
		r.kind = "journal"
		r.headEnd = off[m[13]]
		r.basisHint = "page"
		c.token("source-title", off[m[4]], off[m[5]])
		c.token("authors", off[m[2]], off[m[3]])
		jl := off[m[12]] - len([]rune(journal)) - 1
		c.token("journal", jl, jl+len([]rune(journal)))
		if c.enabled.Has("journal-names") {
			c.change(c.offset+jl, c.offset+jl+len([]rune(journal)), canonical, "CIT-JOURNAL")
		}
		return r, nil
	}

	// Exact book-title boundary plus a completely parsed publication tuple.
	titles := make([]string, 0, len(c.registry.Books))
	for t := range c.registry.Books {
		titles = append(titles, t)
	}
	sortLongestFirst(titles)
	for _, title := range titles {
		m := bookPattern(title).FindStringSubmatchIndex(body)
		if m == nil {
			continue
		}
		edition, _ := group(body, m, 3)
		r.kind = "book"
		// The production ends at the closing parenthesis, i.e. where the
		// consumed lookahead group begins.
		r.headEnd = off[m[14]]
		r.basisHint = "page"
		r.source = []string{"book", body[m[2]:m[3]], body[m[4]:m[5]], edition, body[m[8]:m[9]], body[m[10]:m[11]], body[m[12]:m[13]]}
		c.token("source-title", off[m[4]], off[m[5]])
		return r, nil
	}

	// Web: title/date + explicit online component + one entire URI. No DOI,
	// access-date or source-type inference from the URL extension.
	if m := webRe.FindStringSubmatchIndex(body); m != nil {
		normalizedDate, when, err := ParseDate(body[m[6]:m[7]])
		if err != nil {
			return root{}, err
		}
		kind, _ := group(body, m, 4)
		r.kind = "web"
		r.headEnd = len(c.bodyRunes)
		// repr((year, month, day)) in the reference identity tuple.
		r.source = []string{"web", body[m[2]:m[3]], body[m[4]:m[5]], fmt.Sprintf("(%d, %d, %d)", when.Year, when.Month, when.Day), kind, body[m[10]:m[11]]}
		c.token("source-title", off[m[4]], off[m[5]])
		c.token("url", off[m[10]], off[m[11]])
		c.token("publication-date", off[m[6]], off[m[7]])
		if c.enabled.Has("dates") {
			c.change(c.offset+off[m[6]], c.offset+off[m[7]], normalizedDate, "DATE-DMY")
		}
	}
	return r, nil
}
