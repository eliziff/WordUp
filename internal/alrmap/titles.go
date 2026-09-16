package alrmap

import (
	"sort"

	"github.com/eliziff/WordUp/internal/alrmap/citations"
	"github.com/eliziff/WordUp/internal/alrmap/kernels"
	"github.com/eliziff/WordUp/internal/alrmap/ooxml"
)

// PlanSourceTitles derives a title-case plan for the parsed journal/book/web
// title interiors of one native note. Nested quotation/bracket/URL guards
// inside a title are honoured, and the plan is validated by ooxml.Prepare
// (title family) before it is returned.
//
// The reference keeps this function in ooxml/titles.py; it lives here because
// it needs rules.TitleCase while layout.go needs ooxml, which would be an
// import cycle in Go. ooxml.SourceTitles and ooxml.TitleGuards stay in ooxml.
func PlanSourceTitles(scanned *ooxml.Scan, container string, registry *citations.Registry, language, withAfterColon string) (*ooxml.BoundPlan, error) {
	c, err := scanned.Container(container)
	if err != nil {
		return nil, err
	}
	titles, err := ooxml.SourceTitles(c, registry)
	if err != nil {
		return nil, err
	}
	runes := []rune(c.Text)
	var edits []kernels.Edit
	for _, title := range titles {
		span := title.Interior
		text := string(runes[span.Start:span.End])
		result, err := TitleCase(text, TitleOptions{Language: language, WithAfterColon: withAfterColon})
		if err != nil {
			return nil, err
		}
		nested, err := kernels.Protection(text, nil)
		if err != nil {
			return nil, err
		}
		for _, e := range result.Edits {
			blocked := false
			for _, g := range nested {
				if g.Overlaps(e.ReadStart, e.ReadEnd) {
					blocked = true
					break
				}
			}
			if blocked {
				continue
			}
			e.Start += span.Start
			e.End += span.Start
			e.ReadStart += span.Start
			e.ReadEnd += span.Start
			edits = append(edits, e)
		}
	}
	sort.SliceStable(edits, func(i, j int) bool { return edits[i].Start < edits[j].Start })
	plan := kernels.Plan{Digest: kernels.Digest(c.Text), Edits: edits, Authorized: true, AuthorizationBasis: "parsed title; explicit language policy"}
	bound := &ooxml.BoundPlan{Fingerprint: scanned.Fingerprint(), Container: c.Key, Family: "title", Plan: plan, Registry: registry}
	if _, _, err := ooxml.Prepare(scanned, bound); err != nil {
		return nil, err
	}
	return bound, nil
}
