// Package grade classifies every text change between two Flat OPC snapshots
// of a manuscript (before and after an ALR Setup run) against the ported
// ALR code map. A change is expected when the oracle would have produced it
// (a citation normalization, a formatting-only rewrite such as curled quotes
// or moved punctuation, a case-only heading change, an emptied paragraph, or
// text inside protected field structure); everything else is reported with
// a bounded excerpt. Nothing here launches Word.
package grade

import (
	"errors"
	"fmt"
	"regexp"
	"sort"
	"strings"
	"unicode"

	"github.com/eliziff/WordUp/internal/alrmap/citations"
	"github.com/eliziff/WordUp/internal/alrmap/kernels"
	"github.com/eliziff/WordUp/internal/alrmap/ooxml"
	"github.com/eliziff/WordUp/internal/office"
)

// Options tune the oracle used for notes.
type Options struct {
	Registry    *citations.Registry
	Switches    kernels.Switches // nil means every switch, as the reference default
	MaxExamples int              // bounded example list; 0 means 40
}

// Change is one classified difference.
type Change struct {
	Container string `json:"container"`
	Index     int    `json:"index"`
	Class     string `json:"class"`
	Reason    string `json:"reason,omitempty"`
	Before    string `json:"before,omitempty"`
	After     string `json:"after,omitempty"`
}

// Report is the grading outcome. Expected changes are counted by class;
// unexpected ones are listed first in Examples.
type Report struct {
	Parts      []string       `json:"parts"`
	Paragraphs int            `json:"paragraphs"`
	Notes      int            `json:"notes"`
	Changed    int            `json:"changed"`
	Expected   map[string]int `json:"expected"`
	Unexpected int            `json:"unexpected"`
	Refusals   map[string]int `json:"refusals"`
	Deleted    int            `json:"deleted_paragraphs"`
	Inserted   int            `json:"inserted_paragraphs"`
	Examples   []Change       `json:"examples"`
}

const (
	classCitation  = "expected:citation"
	classFormat    = "expected:format"
	classCase      = "expected:case"
	classProtected = "expected:protected"
	classEmpty     = "expected:empty"
	classUnexp     = "unexpected"
	classDeleted   = "unexpected:deleted"
	classInserted  = "unexpected:inserted"
)

type paragraph struct {
	text      string
	protected bool
}

type story struct {
	body  []paragraph
	notes map[string]paragraph
	order []string
}

// Grade compares two Flat OPC documents.
func Grade(before, after []byte, opts Options) (*Report, error) {
	if opts.MaxExamples <= 0 {
		opts.MaxExamples = 40
	}
	b, partsB, err := load(before)
	if err != nil {
		return nil, fmt.Errorf("before: %w", err)
	}
	a, partsA, err := load(after)
	if err != nil {
		return nil, fmt.Errorf("after: %w", err)
	}
	parts := partsB
	for _, p := range partsA {
		found := false
		for _, q := range parts {
			if q == p {
				found = true
			}
		}
		if !found {
			parts = append(parts, p)
		}
	}
	r := &Report{Parts: parts, Paragraphs: len(b.body), Notes: len(b.notes), Expected: map[string]int{}, Refusals: map[string]int{}}
	var examples []Change
	record := func(c Change) {
		if strings.HasPrefix(c.Class, "expected") {
			r.Expected[c.Class]++
		} else {
			r.Unexpected++
		}
		if c.Class != classEmpty {
			r.Changed++
		}
		examples = append(examples, c)
	}
	// Body paragraphs: align by exact text, then pair the remaining runs by
	// similarity so a curled quote or a deleted empty paragraph does not shift
	// every later comparison.
	for _, pr := range alignBody(b.body, a.body) {
		switch {
		case pr.before == nil:
			r.Inserted++
			record(Change{Container: "body", Index: pr.afterIndex, Class: classify(paragraph{}, *pr.after, false, opts, r), After: excerpt(pr.after.text, 0)})
		case pr.after == nil:
			r.Deleted++
			cls := classDeleted
			if strings.TrimSpace(pr.before.text) == "" {
				cls = classEmpty
			} else if pr.before.protected {
				cls = classProtected
			}
			record(Change{Container: "body", Index: pr.beforeIndex, Class: cls, Before: excerpt(pr.before.text, 0)})
		default:
			if pr.before.text == pr.after.text {
				continue
			}
			c := Change{Container: "body", Index: pr.beforeIndex}
			c.Class = classify(*pr.before, *pr.after, false, opts, r)
			c.Before, c.After = diffExcerpt(pr.before.text, pr.after.text)
			record(c)
		}
	}
	for _, id := range b.order {
		nb := b.notes[id]
		na, ok := a.notes[id]
		if !ok {
			r.Deleted++
			record(Change{Container: id, Class: classDeleted, Before: excerpt(nb.text, 0)})
			continue
		}
		if nb.text == na.text {
			continue
		}
		c := Change{Container: id, Class: classify(nb, na, true, opts, r)}
		c.Before, c.After = diffExcerpt(nb.text, na.text)
		record(c)
	}
	for _, id := range a.order {
		if _, ok := b.notes[id]; !ok {
			r.Inserted++
			record(Change{Container: id, Class: classInserted, After: excerpt(a.notes[id].text, 0)})
		}
	}
	sort.SliceStable(examples, func(i, j int) bool {
		ui, uj := strings.HasPrefix(examples[i].Class, "unexpected"), strings.HasPrefix(examples[j].Class, "unexpected")
		return ui && !uj
	})
	if len(examples) > opts.MaxExamples {
		examples = examples[:opts.MaxExamples]
	}
	r.Examples = examples
	return r, nil
}

func load(data []byte) (story, []string, error) {
	var parts []string
	stylesRaw, err := office.XMLInput(data, "word/styles.xml")
	if err != nil {
		return story{}, nil, fmt.Errorf("word/styles.xml: %w", err)
	}
	styles, err := ooxml.NewStyles(stylesRaw)
	if err != nil {
		return story{}, nil, fmt.Errorf("styles: %w", err)
	}
	docRaw, err := office.XMLInput(data, "word/document.xml")
	if err != nil {
		return story{}, nil, fmt.Errorf("word/document.xml: %w", err)
	}
	parts = append(parts, "word/document.xml")
	scan, err := ooxml.ScanPart(docRaw, styles, "word/document.xml")
	if err != nil {
		return story{}, nil, fmt.Errorf("scan document: %w", err)
	}
	s := story{notes: map[string]paragraph{}}
	for _, c := range scan.Containers {
		if c.Key != "body" {
			continue
		}
		runes := []rune(c.Text)
		for _, span := range c.Paragraphs {
			lo, hi := span[0], span[1]
			if lo < 0 || hi > len(runes) || lo > hi {
				return story{}, nil, errors.New("paragraph span outside body text")
			}
			s.body = append(s.body, paragraph{text: string(runes[lo:hi]), protected: overlaps(c.Hard, lo, hi)})
		}
	}
	if notesRaw, err := office.XMLInput(data, "word/footnotes.xml"); err == nil {
		parts = append(parts, "word/footnotes.xml")
		notes, err := ooxml.ScanPart(notesRaw, styles, "word/footnotes.xml")
		if err != nil {
			return story{}, nil, fmt.Errorf("scan footnotes: %w", err)
		}
		for _, c := range notes.Containers {
			s.notes[c.Key] = paragraph{text: c.Text, protected: overlaps(c.Hard, 0, len([]rune(c.Text)))}
			s.order = append(s.order, c.Key)
		}
	}
	return s, parts, nil
}

func overlaps(spans []kernels.Span, lo, hi int) bool {
	for _, sp := range spans {
		if sp.Start < hi && sp.End > lo {
			return true
		}
	}
	return false
}

type pair struct {
	before, after           *paragraph
	beforeIndex, afterIndex int
}

// alignBody pairs paragraphs across the edit. Exact matches anchor the
// alignment (longest common subsequence on trimmed text); inside each gap,
// paragraphs are paired in order when their canonical forms agree or when
// they are similar enough, and the rest are deletions or insertions.
func alignBody(before, after []paragraph) []pair {
	n, m := len(before), len(after)
	keyB := make([]string, n)
	keyA := make([]string, m)
	for i := range before {
		keyB[i] = strings.TrimSpace(before[i].text)
	}
	for j := range after {
		keyA[j] = strings.TrimSpace(after[j].text)
	}
	// LCS table on exact keys (empty paragraphs never anchor).
	dp := make([][]int, n+1)
	for i := range dp {
		dp[i] = make([]int, m+1)
	}
	for i := n - 1; i >= 0; i-- {
		for j := m - 1; j >= 0; j-- {
			if keyB[i] != "" && keyB[i] == keyA[j] {
				dp[i][j] = dp[i+1][j+1] + 1
			} else if dp[i+1][j] >= dp[i][j+1] {
				dp[i][j] = dp[i+1][j]
			} else {
				dp[i][j] = dp[i][j+1]
			}
		}
	}
	var out []pair
	i, j := 0, 0
	gapB, gapA := []int{}, []int{}
	flush := func() {
		usedA := make([]bool, len(gapA))
		for _, bi := range gapB {
			matched := -1
			for k, aj := range gapA {
				if usedA[k] {
					continue
				}
				if canonical(before[bi].text) == canonical(after[aj].text) || similar(before[bi].text, after[aj].text) {
					matched = k
					break
				}
			}
			if matched >= 0 {
				usedA[matched] = true
				out = append(out, pair{before: &before[bi], after: &after[gapA[matched]], beforeIndex: bi, afterIndex: gapA[matched]})
			} else {
				out = append(out, pair{before: &before[bi], beforeIndex: bi, afterIndex: -1})
			}
		}
		for k, aj := range gapA {
			if !usedA[k] {
				out = append(out, pair{after: &after[aj], beforeIndex: -1, afterIndex: aj})
			}
		}
		gapB, gapA = gapB[:0], gapA[:0]
	}
	for i < n && j < m {
		if keyB[i] != "" && keyB[i] == keyA[j] {
			flush()
			i++
			j++
			continue
		}
		if dp[i+1][j] >= dp[i][j+1] {
			gapB = append(gapB, i)
			i++
		} else {
			gapA = append(gapA, j)
			j++
		}
	}
	for ; i < n; i++ {
		gapB = append(gapB, i)
	}
	for ; j < m; j++ {
		gapA = append(gapA, j)
	}
	flush()
	return out
}

// similar accepts paragraphs whose longer half is shared, measured on runes.
func similar(a, b string) bool {
	ra, rb := []rune(strings.TrimSpace(a)), []rune(strings.TrimSpace(b))
	if len(ra) == 0 || len(rb) == 0 {
		return false
	}
	prefix := 0
	for prefix < len(ra) && prefix < len(rb) && ra[prefix] == rb[prefix] {
		prefix++
	}
	suffix := 0
	for suffix < len(ra)-prefix && suffix < len(rb)-prefix && ra[len(ra)-1-suffix] == rb[len(rb)-1-suffix] {
		suffix++
	}
	longest := len(ra)
	if len(rb) > longest {
		longest = len(rb)
	}
	return float64(prefix+suffix) >= 0.6*float64(longest)
}

var (
	punctuationInsideQuotes = regexp.MustCompile(`(["'])([.,!?;:])`)
	headingPrefix           = regexp.MustCompile(`^(?:(?:PART|Part)\s+[IVXLC]+|[IVXLC]+\.)\s*[–—:.-]?\s*`)
	spaceRuns               = regexp.MustCompile(`[ \t\x{00A0}]+`)
)

// canonical removes the differences a Setup run is entitled to make in prose:
// straight or curly quotes, punctuation on either side of a closing quote,
// leading tabs, space runs and a removed "Part I –" heading prefix.
func canonical(text string) string {
	t := strings.Map(func(r rune) rune {
		switch r {
		case '“', '”', '„':
			return '"'
		case '‘', '’', '‚':
			return '\''
		case ' ':
			return ' '
		}
		return r
	}, text)
	t = punctuationInsideQuotes.ReplaceAllString(t, "$2$1")
	t = spaceRuns.ReplaceAllString(t, " ")
	t = strings.TrimSpace(t)
	t = headingPrefix.ReplaceAllString(t, "")
	return strings.TrimSpace(t)
}

func classify(before, after paragraph, note bool, opts Options, r *Report) string {
	if before.protected || after.protected {
		return classProtected
	}
	cb, ca := canonical(before.text), canonical(after.text)
	if cb == ca {
		return classFormat
	}
	if strings.EqualFold(cb, ca) {
		return classCase
	}
	if note {
		// The scanned note text carries its paragraph marks and the tab the
		// Setup adds after the number; the oracle reads the bare note.
		res, err := citations.NormalizeNote(strings.TrimSpace(before.text), opts.Registry, citations.Options{Enabled: opts.Switches})
		if err != nil {
			r.Refusals[refusalKey(err)]++
		} else if canonical(res.Text) == ca || strings.EqualFold(canonical(res.Text), ca) {
			return classCitation
		}
	}
	return classUnexp
}

func refusalKey(err error) string {
	msg := err.Error()
	if i := strings.IndexAny(msg, ":("); i > 0 {
		msg = msg[:i]
	}
	return strings.TrimSpace(msg)
}

// diffExcerpt shows the differing region with a little context on each side.
func diffExcerpt(before, after string) (string, string) {
	rb, ra := []rune(before), []rune(after)
	lo := 0
	for lo < len(rb) && lo < len(ra) && rb[lo] == ra[lo] {
		lo++
	}
	hb, ha := len(rb), len(ra)
	for hb > lo && ha > lo && rb[hb-1] == ra[ha-1] {
		hb--
		ha--
	}
	const ctx = 30
	start := lo - ctx
	if start < 0 {
		start = 0
	}
	cut := func(rs []rune, hi int) string {
		end := hi + ctx
		if end > len(rs) {
			end = len(rs)
		}
		return strings.Map(func(r rune) rune {
			if unicode.IsControl(r) && r != '\t' {
				return '¶'
			}
			return r
		}, string(rs[start:end]))
	}
	return cut(rb, hb), cut(ra, ha)
}

func excerpt(text string, from int) string {
	rs := []rune(text)
	if from > len(rs) {
		from = len(rs)
	}
	end := from + 80
	if end > len(rs) {
		end = len(rs)
	}
	return string(rs[from:end])
}
