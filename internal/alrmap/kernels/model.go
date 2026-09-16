// Package kernels holds the offline ALR reference kernels: exact spans,
// snapshot-bound plans, conservative quotation/URI masks, advisory lexical
// plans, bounded locator grammars and exact source utilities. Nothing here
// touches Word; every offset is a Unicode code-point index into the string it
// was computed from, as in the Python reference.
package kernels

import (
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"sort"
	"unicode/utf16"

	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
)

// PermissionError reports a plan without attributed write authority. It is
// the only refusal that the reference implementation raised as a
// PermissionError rather than a ValueError, so callers can tell "advisory
// plan" apart from "invalid plan".
type PermissionError struct{ Reason string }

func (e *PermissionError) Error() string { return e.Reason }

// InvariantError reports a broken internal invariant (a Python
// AssertionError): a serializer and its leaf edits disagreeing, or an XML
// copy whose projection differs from the planned text.
type InvariantError struct{ Reason string }

func (e *InvariantError) Error() string { return e.Reason }

// Span is a protected or annotated code-point interval.
type Span struct {
	Start, End int
	Reason     string
}

// Overlaps reports whether [start,end) touches the span. An insertion exactly
// at a protected endpoint is also refused.
func (s Span) Overlaps(start, end int) bool {
	if start == end {
		return s.Start <= start && start <= s.End
	}
	return s.Start < end && start < s.End
}

func spanLess(a, b Span) bool {
	if a.Start != b.Start {
		return a.Start < b.Start
	}
	if a.End != b.End {
		return a.End < b.End
	}
	return a.Reason < b.Reason
}

// SortSpans orders spans by (start, end, reason), the reference dataclass order.
func SortSpans(spans []Span) []Span {
	out := append([]Span(nil), spans...)
	sort.SliceStable(out, func(i, j int) bool { return spanLess(out[i], out[j]) })
	return out
}

// Edit is one occurrence-addressed replacement. Start/End bound the written
// text, ReadStart/ReadEnd the complete recognized token that justified it.
type Edit struct {
	Start, End         int
	Old, New           string
	Rule               string
	ReadStart, ReadEnd int
}

// Plan is a snapshot-bound set of edits with the protections they were
// checked against and the caller-supplied write authority.
type Plan struct {
	Digest             string
	Edits              []Edit
	Protected          []Span
	Authorized         bool
	AuthorizationBasis string
}

// Digest is the SHA-256 of the UTF-8 text, the snapshot identity of a plan.
func Digest(text string) string {
	sum := sha256.Sum256([]byte(text))
	return hex.EncodeToString(sum[:])
}

// UTF16Offset converts a code-point offset into a UTF-16 code-unit offset.
func UTF16Offset(text string, offset int) (int, error) {
	runes := []rune(text)
	if offset < 0 || offset > len(runes) {
		return 0, errors.New("offset outside text")
	}
	return len(utf16.Encode(runes[:offset])), nil
}

// IsWord mirrors the reference _word predicate (alphanumeric, underscore or
// combining mark).
func IsWord(r rune) bool { return pytext.IsWord(r) }

// ValidateSpans sorts spans and refuses any lying outside the container.
func ValidateSpans(text string, spans []Span) ([]Span, error) {
	return validateSpans(len([]rune(text)), spans)
}

func validateSpans(length int, spans []Span) ([]Span, error) {
	result := SortSpans(spans)
	for _, s := range result {
		if !(0 <= s.Start && s.Start <= s.End && s.End <= length) {
			return nil, errors.New("protection span outside container")
		}
	}
	return result, nil
}

// MinimalEdit trims the common prefix/suffix of text[lo:hi] and replacement
// and returns the narrowest edit, or nil when nothing changes.
func MinimalEdit(text string, lo, hi int, replacement, rule string) *Edit {
	return minimalEdit([]rune(text), lo, hi, replacement, rule)
}

func minimalEdit(text []rune, lo, hi int, replacement, rule string) *Edit {
	old := text[lo:hi]
	repl := []rune(replacement)
	if string(old) == replacement {
		return nil
	}
	left := 0
	for left < min(len(old), len(repl)) && old[left] == repl[left] {
		left++
	}
	right := 0
	for right < min(len(old)-left, len(repl)-left) && old[len(old)-1-right] == repl[len(repl)-1-right] {
		right++
	}
	a, b := lo+left, hi-right
	newEnd := len(repl)
	if right != 0 {
		newEnd = len(repl) - right
	}
	return &Edit{Start: a, End: b, Old: string(text[a:b]), New: string(repl[left:newEnd]), Rule: rule, ReadStart: lo, ReadEnd: hi}
}

// ApplyPlan applies a plan to a string only. It refuses advisory, stale,
// overlapping or unsafe plans. Induction over disjoint edits establishes that
// every original interval outside the write set keeps its literal text; this
// is not a Word package proof.
func ApplyPlan(text string, plan Plan) (string, error) {
	if !plan.Authorized || plan.AuthorizationBasis == "" {
		return "", &PermissionError{"plan has no attributed write authority"}
	}
	runes := []rune(text)
	if _, err := validateSpans(len(runes), plan.Protected); err != nil {
		return "", err
	}
	if Digest(text) != plan.Digest {
		return "", errors.New("stale plan")
	}
	edits := sortedEdits(plan.Edits)
	for i, e := range edits {
		if !(0 <= e.ReadStart && e.ReadStart <= e.Start && e.Start <= e.End && e.End <= e.ReadEnd && e.ReadEnd <= len(runes)) {
			return "", errors.New("bad edit bounds")
		}
		if string(runes[e.Start:e.End]) != e.Old {
			return "", errors.New("old text differs")
		}
		for _, g := range plan.Protected {
			if g.Overlaps(e.ReadStart, e.ReadEnd) || g.Overlaps(e.Start, e.End) {
				return "", errors.New("protected overlap")
			}
		}
		if i > 0 {
			previous := edits[i-1]
			if e.Start < previous.End || (e.Start == previous.End && (e.Start == e.End || previous.Start == previous.End)) {
				return "", errors.New("overlapping or boundary-coincident edits")
			}
		}
	}
	var out []rune
	cursor := 0
	for _, e := range edits {
		out = append(out, runes[cursor:e.Start]...)
		out = append(out, []rune(e.New)...)
		cursor = e.End
	}
	out = append(out, runes[cursor:]...)
	return string(out), nil
}

func sortedEdits(edits []Edit) []Edit {
	out := append([]Edit(nil), edits...)
	sort.SliceStable(out, func(i, j int) bool { return out[i].Start < out[j].Start })
	return out
}

// Switches is an explicit feature-switch set. A nil set means "the default
// set" of the receiving operation; an empty non-nil set disables everything.
type Switches map[string]bool

// NewSwitches builds a non-nil switch set (an empty call disables every switch).
func NewSwitches(names ...string) Switches {
	s := make(Switches, len(names))
	for _, n := range names {
		s[n] = true
	}
	return s
}

// Has reports whether the switch is enabled.
func (s Switches) Has(name string) bool { return s[name] }

// Without returns a copy with the named switches removed.
func (s Switches) Without(names ...string) Switches {
	out := make(Switches, len(s))
	for k, v := range s {
		if v {
			out[k] = true
		}
	}
	for _, n := range names {
		delete(out, n)
	}
	return out
}

// Names returns the enabled switches in sorted order.
func (s Switches) Names() []string {
	var out []string
	for k, v := range s {
		if v {
			out = append(out, k)
		}
	}
	sort.Strings(out)
	return out
}
