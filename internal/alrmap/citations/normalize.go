package citations

import (
	"errors"
	"regexp"
	"strings"

	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
	"github.com/eliziff/WordUp/internal/alrmap/kernels"
)

var (
	bracketSuffixRe = regexp.MustCompile(` \[([^\[\]]+)\]$`)
	permaSuffixRe   = regexp.MustCompile(`^perma\.cc/[A-Z0-9]{4}-[A-Z0-9]{4}$`)
	courtFrontRe    = regexp.MustCompile(`^ \((` + courtAlternation + `)\)`)
	courtRearRe     = regexp.MustCompile(` \((` + courtAlternation + `)\)$`)
)

// component carries one top-level note component through the root parsers.
// s is the component after its signal was peeled; body is s without its
// bracket suffixes. All offsets are code points; offset is the absolute
// position of s (and body) inside the note text.
type component struct {
	text        string
	offset      int
	s           []rune
	body        string
	bodyRunes   []rune
	bodyOffsets []int
	registry    *Registry
	enabled     kernels.Switches
	tokens      []Token
	edits       *[]kernels.Edit
}

func (c *component) token(kind string, lo, hi int) {
	c.tokens = append(c.tokens, Token{Kind: kind, Start: c.offset + lo, End: c.offset + hi, Value: string(c.s[lo:hi])})
}

func (c *component) change(lo, hi int, new, rule string) {
	if e := kernels.MinimalEdit(c.text, lo, hi, new, rule); e != nil {
		*c.edits = append(*c.edits, *e)
	}
}

// Options are the keyword arguments of the reference normalize_note.
type Options struct {
	Enabled    kernels.Switches // nil means every switch (the reference default)
	BasisHints []string         // one per component; "" means no hint
}

// NormalizeNote reads one complete logical footnote, excluding its native
// reference/mark. Returned edits are advisory until bound to actual
// native/XML container protections. Title words, URLs, identifiers and
// qualifiers are not globally normalized. Text containing an unexplained
// component is rejected wholesale.
func NormalizeNote(text string, registry *Registry, opts Options) (Result, error) {
	if registry == nil {
		registry = &Registry{}
	}
	enabled := opts.Enabled
	if enabled == nil {
		enabled = switchSet
	}
	for name := range enabled {
		if enabled[name] && !switchSet[name] {
			return Result{}, errors.New("unknown feature switch")
		}
	}
	runes := []rune(text)
	if len(runes) == 0 || len(runes) > 100_000 || strings.ContainsAny(text, "\r\n\t￼\x02\x13\x14\x15") {
		return Result{}, errors.New("unsupported logical note")
	}
	// Require a terminal or explicit terminal-repair permission.
	terminal := strings.HasSuffix(text, ".")
	core := runes
	if terminal {
		core = runes[:len(runes)-1]
	}
	if !terminal && !enabled.Has("terminal") {
		return Result{}, errors.New("missing note terminal")
	}
	pieces, err := SplitTop(string(core), ';')
	if err != nil {
		return Result{}, err
	}
	if len(opts.BasisHints) > 0 && len(opts.BasisHints) != len(pieces) {
		return Result{}, errors.New("one basis hint per complete production required")
	}
	for _, h := range opts.BasisHints {
		if h != "" && h != "page" && h != "paragraph" && h != "section" {
			return Result{}, errors.New("unknown locator basis hint")
		}
	}
	var edits []kernels.Edit
	var citations []Citation
	for pieceIndex, p := range pieces {
		a, b := p.Start, p.End
		for a < b && core[a] == ' ' {
			a++
		}
		for b > a && core[b-1] == ' ' {
			b--
		}
		c := &component{text: text, offset: a, s: core[a:b], registry: registry, enabled: enabled, edits: &edits}
		for _, signal := range Signals {
			if strings.HasPrefix(string(c.s), signal+" ") {
				n := len([]rune(signal))
				c.tokens = append(c.tokens, Token{Kind: "signal", Start: c.offset, End: c.offset + n, Value: signal})
				c.offset += n + 1
				c.s = c.s[n+1:]
				break
			}
		}
		var qualifications []string
		declared := ""
		// Peel only recognized trailing qualifiers / one explicit alias. Unknown
		// bracket data is NOT dropped, rewritten, or assumed to be a short form.
		tailEnd := len(c.s)
		for {
			head := string(c.s[:tailEnd])
			m := bracketSuffixRe.FindStringSubmatchIndex(head)
			if m == nil {
				break
			}
			inner := head[m[2]:m[3]]
			if permaSuffixRe.MatchString(inner) {
				break
			}
			offsets := pytext.RuneOffsets(head)
			items := strings.Split(inner, ", ")
			known := true
			for _, x := range items {
				if !Qualifications[x] {
					known = false
					break
				}
			}
			switch {
			case known:
				qualifications = append(append([]string(nil), items...), qualifications...)
				c.token("qualification", offsets[m[0]], offsets[m[1]])
			case registry.Aliases[inner] && declared == "":
				declared = inner
				c.token("shortform-declaration", offsets[m[0]], offsets[m[1]])
			default:
				return Result{}, errors.New("unrecognized bracket suffix")
			}
			tailEnd = offsets[m[0]]
		}
		c.bodyRunes = c.s[:tailEnd]
		c.body = string(c.bodyRunes)
		c.bodyOffsets = pytext.RuneOffsets(c.body)

		r, err := parsePrimary(c)
		if err != nil {
			return Result{}, err
		}
		if r.kind == "" {
			if r, err = parseSecondary(c); err != nil {
				return Result{}, err
			}
		}
		if r.kind == "" {
			return Result{}, errors.New("no complete supported root")
		}
		basisHint := r.basisHint
		if len(opts.BasisHints) > 0 {
			hint := opts.BasisHints[pieceIndex]
			if basisHint != "" && hint != "" && basisHint != hint {
				return Result{}, errors.New("locator basis disagrees with source production")
			}
			if basisHint == "" {
				basisHint = hint
			}
		}
		rawTail := string(c.bodyRunes[r.headEnd:])
		mainTail := rawTail
		beforeCourt, afterCourt := "", ""
		if r.kind == "case" || r.kind == "reporter" {
			if front := courtFrontRe.FindStringIndex(mainTail); front != nil {
				beforeCourt = mainTail[:front[1]]
				mainTail = mainTail[front[1]:]
			}
			if rear := courtRearRe.FindStringIndex(mainTail); rear != nil {
				afterCourt = mainTail[rear[0]:]
				mainTail = mainTail[:rear[0]]
			}
			if beforeCourt != "" && afterCourt != "" {
				return Result{}, errors.New("duplicate court components")
			}
			if r.kind == "case" && r.source[2] != "CanLII" && (beforeCourt != "" || afterCourt != "") {
				return Result{}, errors.New("unneeded/unknown neutral-citation suffix")
			}
		}
		newLocator, basis, loc, err := locator(mainTail, basisHint, enabled)
		if err != nil {
			return Result{}, err
		}
		moveCourt := r.kind == "case" && beforeCourt != "" && mainTail != "" && enabled.Has("court-order")
		newTail := beforeCourt + newLocator + afterCourt
		if moveCourt {
			newTail = newLocator + beforeCourt
		}
		if rawTail != "" {
			c.token("locator-and-court", r.headEnd, len(c.bodyRunes))
			if moveCourt {
				c.change(c.offset+r.headEnd, c.offset+len(c.bodyRunes), newTail, "CIT-COURT-MOVE")
			} else {
				leaf, err := LocatorEdits(text, c.offset+r.headEnd+len([]rune(beforeCourt)), mainTail, basisHint, enabled)
				if err != nil {
					return Result{}, err
				}
				edits = append(edits, leaf...)
			}
		}
		source := r.source
		if r.kind == "case" && r.source[2] == "CanLII" {
			court := beforeCourt
			if court == "" {
				court = afterCourt
			}
			source = append(append([]string(nil), source...), strings.Trim(court, " ()"))
		}
		citations = append(citations, Citation{Kind: r.kind, Source: source, Alias: r.alias, NoteLabel: r.noteLabel, Basis: basis,
			Locator: loc, Qualifications: qualifications, DeclaredAlias: declared, Tokens: c.tokens})
	}
	if !terminal {
		edits = append(edits, kernels.Edit{Start: len(runes), End: len(runes), Old: "", New: ".", Rule: "CIT-TERMINAL", ReadStart: max(0, len(runes)-1), ReadEnd: len(runes)})
	}
	edits = sortEditsByStart(edits)
	for i := 1; i < len(edits); i++ {
		if edits[i].Start < edits[i-1].End {
			return Result{}, errors.New("overlapping citation edits")
		}
	}
	canonical := append([]rune(nil), runes...)
	for i := len(edits) - 1; i >= 0; i-- {
		e := edits[i]
		canonical = append(append(append([]rune(nil), canonical[:e.Start]...), []rune(e.New)...), canonical[e.End:]...)
	}
	return Result{Text: string(canonical), Edits: edits, Citations: citations}, nil
}

// sortEditsByStart is a stable sort by start (list.sort(key=start)).
func sortEditsByStart(edits []kernels.Edit) []kernels.Edit {
	out := append([]kernels.Edit(nil), edits...)
	for i := 1; i < len(out); i++ {
		for j := i; j > 0 && out[j].Start < out[j-1].Start; j-- {
			out[j], out[j-1] = out[j-1], out[j]
		}
	}
	return out
}
