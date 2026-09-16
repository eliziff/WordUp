package ooxml

import (
	"errors"

	"github.com/eliziff/WordUp/internal/alrmap/citations"
	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
	"github.com/eliziff/WordUp/internal/alrmap/kernels"
)

// SourceTitle is one parsed journal/book/web title: the outer delimited span
// (nil when the title is not quote-wrapped) and its interior.
type SourceTitle struct {
	Outer    *kernels.Span
	Interior kernels.Span
}

// SourceTitles fully parses the enclosing note and selects only journal,
// book and web source-title interiors. Case and legislation titles are
// excluded from the general title policy.
func SourceTitles(c *Container, registry *citations.Registry) ([]SourceTitle, error) {
	start, end, err := noteBounds(c, "title grammar requires one complete native note",
		"title production intersects opaque structure", "title production is inside a protected paragraph role")
	if err != nil {
		return nil, err
	}
	runes := []rune(c.Text)
	result, err := citations.NormalizeNote(string(runes[start:end]), registry, citations.Options{Enabled: kernels.NewSwitches()})
	if err != nil {
		return nil, err
	}
	var out []SourceTitle
	for _, citation := range result.Citations {
		if citation.Kind != "journal" && citation.Kind != "book" && citation.Kind != "web" {
			continue
		}
		for _, token := range citation.Tokens {
			if token.Kind != "source-title" {
				continue
			}
			a, b := start+token.Start, start+token.End
			if (runes[a] == '“' || runes[a] == '"') && (runes[b-1] == '”' || runes[b-1] == '"') {
				out = append(out, SourceTitle{Outer: &kernels.Span{Start: a, End: b, Reason: "outer-title"}, Interior: kernels.Span{Start: a + 1, End: b - 1, Reason: "title-interior"}})
			} else {
				out = append(out, SourceTitle{Interior: kernels.Span{Start: a, End: b, Reason: "title-interior"}})
			}
		}
	}
	return out, nil
}

// TitleGuards re-derives a narrow capability from the full production at
// write time. Only capitalization of parsed journal/book/web title interiors
// is permitted; nested quotations, brackets, URLs and native structural
// masks remain intact. A forged rule name alone cannot grant a spelling
// operation access to a title.
func TitleGuards(scanned *Scan, c *Container, bound *BoundPlan) ([]kernels.Span, error) {
	if bound.Registry == nil {
		return nil, errors.New("title plan lacks parser inputs")
	}
	titles, err := SourceTitles(c, bound.Registry)
	if err != nil {
		return nil, err
	}
	outer := map[[2]int]bool{}
	for _, t := range titles {
		if t.Outer != nil {
			outer[[2]int{t.Outer.Start, t.Outer.End}] = true
		}
	}
	for _, e := range bound.Plan.Edits {
		if e.Rule != "TITLE-EN" || pytext.Casefold(e.Old) != pytext.Casefold(e.New) {
			return nil, errors.New("title capability permits case changes only")
		}
		inside := false
		for _, t := range titles {
			if t.Interior.Start <= e.ReadStart && e.ReadStart <= e.ReadEnd && e.ReadEnd <= t.Interior.End {
				inside = true
				break
			}
		}
		if !inside {
			return nil, errors.New("edit is outside a parsed title interior")
		}
	}
	guards, err := c.Guards("citation")
	if err != nil {
		return nil, err
	}
	var out []kernels.Span
	for _, g := range guards {
		if g.Reason == "delimited-source-or-title" && outer[[2]int{g.Start, g.End}] {
			continue
		}
		out = append(out, g)
	}
	return out, nil
}
