package kernels

import (
	"errors"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
)

// Lexicon is the finite Canadian-spelling map of the base kernel.
var Lexicon = map[string]string{"towards": "toward", "judgement": "judgment", "judgements": "judgments", "caselaw": "case law"}

// MonthNames lists the English month names in calendar order.
var MonthNames = strings.Split("January February March April May June July August September October November December", " ")

// Months maps lower-cased month names to their 1-based number.
var Months = func() map[string]int {
	m := map[string]int{}
	for i, name := range MonthNames {
		m[strings.ToLower(name)] = i + 1
	}
	return m
}()

// MonthAlternation is the case-insensitive month alternation shared by the
// date grammars.
var MonthAlternation = strings.ToLower(strings.Join(MonthNames, "|"))

// The reference DATE, PERCENT, WORDS and SPACES patterns carry lookarounds
// ((?<!\w), (?!\w), (?<![\w.,+\-]), (?![%\w]), \b, (?<=\S), (?=\S)). RE2 has
// none, so the patterns below match the visible token only and the plan
// checks the neighbouring code points by hand. Because a rejected candidate
// can never contain a later valid candidate for these grammars, resuming
// after the rejected match is equivalent to Python's position-by-position scan.
var (
	dateRe    = regexp.MustCompile(`(?i)(` + MonthAlternation + `) +([1-9]|[12][0-9]|3[01]),? +([12][0-9]{3})`)
	percentRe = regexp.MustCompile(`([0-9]+(?:\.[0-9]+)?) *%`)
	wordsRe   = regexp.MustCompile(`(?i)\b(?:towards|judgements?|caselaw)\b`)
	spacesRe  = regexp.MustCompile(` {2,}`)
)

// ValidDate reports whether the proleptic Gregorian date exists.
func ValidDate(year, month, day int) bool {
	if month < 1 || month > 12 || day < 1 {
		return false
	}
	t := time.Date(year, time.Month(month), day, 0, 0, 0, 0, time.UTC)
	return t.Year() == year && int(t.Month()) == month && t.Day() == day
}

// CaseLike transfers the case pattern of old onto new: all lower, all upper,
// or initial capital. Other patterns yield ok=false.
func CaseLike(old, new string) (string, bool) {
	if pytext.IsLower(old) {
		return new, true
	}
	if pytext.IsUpper(old) {
		return strings.ToUpper(new), true
	}
	oldRunes := []rune(old)
	if len(oldRunes) > 0 && pytext.IsUpper(string(oldRunes[:1])) && pytext.IsLower(string(oldRunes[1:])) {
		return pytext.Capitalize(new), true
	}
	return "", false
}

// DefaultProseSwitches is the lexical switch set plan_prose enables by default.
var DefaultProseSwitches = []string{"spelling", "date", "percent", "spaces"}

// ProseOptions are the keyword arguments of the reference plan_prose.
type ProseOptions struct {
	Extra              []Span
	Authorized         bool
	AuthorizationBasis string
	Enabled            Switches // nil means DefaultProseSwitches
}

// PlanProse plans finite lexical edits. The default is a read-only advisory
// plan on raw input. Authorized means the caller supplied an explicit
// author-text scope; native "Normal", absence of quotation marks and existing
// italic state do NOT supply that scope. It is input, not an inference.
func PlanProse(text string, opts ProseOptions) (Plan, error) {
	if opts.Authorized && opts.AuthorizationBasis == "" {
		return Plan{}, errors.New("authorization needs explicit provenance")
	}
	enabled := opts.Enabled
	if enabled == nil {
		enabled = NewSwitches(DefaultProseSwitches...)
	}
	runes := []rune(text)
	guards, err := protection(runes, text, opts.Extra)
	if err != nil {
		return Plan{}, err
	}
	offsets := pytext.RuneOffsets(text)
	var edits []Edit
	add := func(lo, hi int, new, rule string) {
		// Test the complete recognized token (read set), not just its deletion.
		for _, g := range guards {
			if g.Overlaps(lo, hi) {
				return
			}
		}
		if e := minimalEdit(runes, lo, hi, new, rule); e != nil {
			edits = append(edits, *e)
		}
	}
	if enabled.Has("spelling") {
		for _, m := range wordsRe.FindAllStringIndex(text, -1) {
			lo, hi := offsets[m[0]], offsets[m[1]]
			left, right := pytext.RuneAt(runes, lo-1), pytext.RuneAt(runes, hi)
			if pytext.IsWord(left) || pytext.IsWord(right) || left == '-' || left == '/' || right == '-' || right == '/' {
				continue
			}
			word := text[m[0]:m[1]]
			if new, ok := CaseLike(word, Lexicon[strings.ToLower(word)]); ok {
				add(lo, hi, new, "LEX-"+strings.ToLower(word))
			}
		}
	}
	if enabled.Has("date") {
		for _, m := range dateRe.FindAllStringSubmatchIndex(text, -1) {
			lo, hi := offsets[m[0]], offsets[m[1]]
			if pytext.IsWord(pytext.RuneAt(runes, lo-1)) || pytext.IsWord(pytext.RuneAt(runes, hi)) {
				continue
			}
			monthText := text[m[2]:m[3]]
			day, _ := strconv.Atoi(text[m[4]:m[5]])
			year, _ := strconv.Atoi(text[m[6]:m[7]])
			if !ValidDate(year, Months[strings.ToLower(monthText)], day) {
				continue
			}
			add(lo, hi, strconv.Itoa(day)+" "+pytext.Capitalize(strings.ToLower(monthText))+" "+strconv.Itoa(year), "DATE-DMY")
		}
	}
	if enabled.Has("percent") {
		for _, m := range percentRe.FindAllStringSubmatchIndex(text, -1) {
			lo, hi := offsets[m[0]], offsets[m[1]]
			before, after := pytext.RuneAt(runes, lo-1), pytext.RuneAt(runes, hi)
			// (?<![\w.,+\-]) and (?![%\w]) from the reference pattern, plus its
			// explicit refusal of a preceding Unicode minus sign.
			if pytext.IsWord(before) || strings.ContainsRune(".,+-−", before) || pytext.IsWord(after) || after == '%' {
				continue
			}
			add(lo, hi, text[m[2]:m[3]]+" percent", "NUM-PERCENT")
		}
	}
	if enabled.Has("spaces") {
		for _, m := range spacesRe.FindAllStringIndex(text, -1) {
			lo, hi := offsets[m[0]], offsets[m[1]]
			// (?<=\S) and (?=\S): only interword runs, never a run at an edge
			// or beside other whitespace.
			before, after := pytext.RuneAt(runes, lo-1), pytext.RuneAt(runes, hi)
			if before < 0 || after < 0 || pytext.IsSpace(before) || pytext.IsSpace(after) {
				continue
			}
			// Never infer permission to alter whitespace at an opaque edge.
			blocked := false
			for _, g := range guards {
				if g.Overlaps(lo-1, hi+1) {
					blocked = true
					break
				}
			}
			if blocked {
				continue
			}
			add(lo, hi, " ", "TXT-ASCII-SPACES")
		}
	}
	// Semantic recognizers may overlap: an entire date with internal spaces,
	// for example. Do not silently select a winner; expose a separate phase.
	ordered := append([]Edit(nil), edits...)
	sort.SliceStable(ordered, func(i, j int) bool {
		a, b := ordered[i], ordered[j]
		if a.ReadStart != b.ReadStart {
			return a.ReadStart < b.ReadStart
		}
		if a.ReadEnd != b.ReadEnd {
			return a.ReadEnd < b.ReadEnd
		}
		return a.Rule < b.Rule
	})
	var kept []Edit
	for _, e := range ordered {
		conflict := false
		for _, p := range kept {
			if e.ReadStart < p.ReadEnd && p.ReadStart < e.ReadEnd {
				conflict = true
				break
			}
		}
		if conflict {
			// Date owns internal spacing; equivalent spaces edits are redundant.
			redundant := false
			if e.Rule == "TXT-ASCII-SPACES" {
				for _, p := range kept {
					if (p.Rule == "DATE-DMY" || p.Rule == "NUM-PERCENT") && p.ReadStart <= e.ReadStart && p.ReadEnd >= e.ReadEnd {
						redundant = true
						break
					}
				}
			}
			if redundant {
				continue
			}
			return Plan{}, errors.New("conflicting recognizer read sets")
		}
		kept = append(kept, e)
	}
	return Plan{Digest(text), sortedEdits(kept), guards, opts.Authorized, opts.AuthorizationBasis}, nil
}
