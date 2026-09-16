package alrmap

import (
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"fmt"
	"sort"
	"strconv"
	"strings"

	"github.com/eliziff/WordUp/internal/alrmap/citations"
	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
)

// Note is one native note: its persistent occurrence identity (ID), its
// currently displayed number (Label), its numbering scope and its parsed
// citations. Native note IDs and displayed labels are separate inputs.
type Note struct {
	ID        string
	Label     string
	Scope     string
	Citations []citations.Citation
}

// Binding is one citation occurrence resolved to a source.
type Binding struct {
	NoteID         string
	Occurrence     int
	Source         []string
	Basis          string
	Locator        *citations.Locator
	Qualifications []string
	OriginNote     string
}

// Meaning is the comparable (source, basis, locator, qualifications) tuple.
func (b Binding) Meaning() string {
	return strings.Join([]string{citations.SourceKey(b.Source), b.Basis, b.Locator.String(), strings.Join(b.Qualifications, "\x1f")}, "\x1e")
}

// BindingKey addresses one citation occurrence inside a note.
type BindingKey struct {
	NoteID     string
	Occurrence int
}

// Graph is the compiled reference graph. This model never inserts a Word
// field: a pointer-preserving command returns the exact target ID, and a
// correction command requires a unique parsed source/alias association.
type Graph struct {
	Fingerprint string
	Notes       []Note
	Bindings    map[BindingKey]Binding
	Unresolved  map[BindingKey]string
	Aliases     map[string][][]string
}

// PreservePointer returns the one native note ID displaying label in scope.
func (g *Graph) PreservePointer(scope, label string) (string, error) {
	var targets []string
	for _, n := range g.Notes {
		if n.Scope == scope && n.Label == label {
			targets = append(targets, n.ID)
		}
	}
	if len(targets) != 1 {
		return "", errors.New("display label has no unique native target")
	}
	return targets[0], nil
}

// CorrectAliasTarget finds the first preceding full citation of one exact
// declared source. It corrects a source association only, not legal
// relevance; names/editions not connected by a parsed identity remain
// different, and there is no same-surname rule.
func (g *Graph) CorrectAliasTarget(noteID, alias string) (string, error) {
	sources := g.Aliases[alias]
	if len(sources) != 1 {
		return "", errors.New("unknown/colliding alias")
	}
	index := -1
	for i, n := range g.Notes {
		if n.ID == noteID {
			index = i
			break
		}
	}
	if index < 0 {
		return "", errors.New("unknown note ID")
	}
	want := citations.SourceKey(sources[0])
	for _, n := range g.Notes[:index] {
		for _, c := range n.Citations {
			if c.Source != nil && citations.SourceKey(c.Source) == want {
				return n.ID, nil
			}
		}
	}
	return "", errors.New("no preceding full-source declaration")
}

var sourceKinds = map[string]bool{"case": true, "reporter": true, "journal": true, "book": true, "web": true, "legislation": true}

func compareSource(a, b []string) int {
	for i := 0; i < len(a) && i < len(b); i++ {
		if c := strings.Compare(a[i], b[i]); c != 0 {
			return c
		}
	}
	return len(a) - len(b)
}

// notesFingerprint is the stable identity of the analyzed note sequence
// (the reference hashes repr(notes)); it covers IDs, labels, scopes and the
// complete parsed citations including their tokens.
func notesFingerprint(notes []Note) string {
	var b strings.Builder
	for _, n := range notes {
		fmt.Fprintf(&b, "%q\x1d%q\x1d%q\x1d", n.ID, n.Label, n.Scope)
		for _, c := range n.Citations {
			fmt.Fprintf(&b, "%q\x1c", c.Meaning())
			for _, t := range c.Tokens {
				fmt.Fprintf(&b, "%q:%d:%d:%q;", t.Kind, t.Start, t.End, t.Value)
			}
			b.WriteString("\x1c")
		}
		b.WriteString("\x1d\n")
	}
	sum := sha256.Sum256([]byte(b.String()))
	return hex.EncodeToString(sum[:])
}

// CompileGraph binds complete parsed sources, exact declarations, supported
// supra pointers and supported Ibid chains. Unresolvable occurrences are
// recorded with their reason rather than guessed.
func CompileGraph(notes []Note) (*Graph, error) {
	notes = append([]Note(nil), notes...)
	if len(notes) == 0 {
		return nil, errors.New("complete native note identities and numbering scopes required")
	}
	ids := map[string]bool{}
	for _, n := range notes {
		if n.ID == "" || n.Scope == "" || n.Label == "" {
			return nil, errors.New("complete native note identities and numbering scopes required")
		}
		ids[n.ID] = true
	}
	if len(ids) != len(notes) {
		return nil, errors.New("duplicate native note ID")
	}
	aliasSets := map[string]map[string][]string{}
	for _, n := range notes {
		for _, c := range n.Citations {
			if c.Source == nil {
				continue
			}
			if !sourceKinds[c.Kind] {
				return nil, errors.New("source key attached to an unsupported citation kind")
			}
			if len(c.Source) < 2 {
				return nil, errors.New("malformed parsed source key")
			}
			if c.DeclaredAlias != "" {
				if aliasSets[c.DeclaredAlias] == nil {
					aliasSets[c.DeclaredAlias] = map[string][]string{}
				}
				aliasSets[c.DeclaredAlias][citations.SourceKey(c.Source)] = c.Source
			}
		}
	}
	aliases := map[string][][]string{}
	for k, set := range aliasSets {
		var list [][]string
		for _, s := range set {
			list = append(list, s)
		}
		sort.Slice(list, func(i, j int) bool { return compareSource(list[i], list[j]) < 0 })
		aliases[k] = list
	}
	g := &Graph{Fingerprint: notesFingerprint(notes), Notes: notes, Bindings: map[BindingKey]Binding{}, Unresolved: map[BindingKey]string{}, Aliases: aliases}
	index := map[string]int{}
	for i, n := range notes {
		index[n.ID] = i
	}
	for i, n := range notes {
		for j, c := range n.Citations {
			key := BindingKey{n.ID, j}
			if err := g.bind(notes, index, i, j, c, key); err != nil {
				g.Unresolved[key] = err.Error()
			}
		}
	}
	return g, nil
}

func (g *Graph) bind(notes []Note, index map[string]int, i, j int, c citations.Citation, key BindingKey) error {
	n := notes[i]
	if c.Source != nil {
		identified := false
		for _, s := range c.Source[1:] {
			if s != "" {
				identified = true
				break
			}
		}
		if c.Source[0] == "" || !identified {
			return errors.New("empty source identity")
		}
		g.Bindings[key] = Binding{n.ID, j, c.Source, c.Basis, c.Locator, c.Qualifications, n.ID}
		return nil
	}
	switch c.Kind {
	case "supra":
		target, err := g.PreservePointer(n.Scope, c.NoteLabel)
		if err != nil {
			return err
		}
		if index[target] >= i {
			return errors.New("supra does not target a preceding note")
		}
		targetNote := notes[index[target]]
		var candidates []Binding
		for k := range targetNote.Citations {
			if b, ok := g.Bindings[BindingKey{target, k}]; ok {
				candidates = append(candidates, b)
			}
		}
		// An unresolved occurrence in the target makes it incomplete.
		if len(candidates) != len(targetNote.Citations) {
			return errors.New("target note has unresolved source components")
		}
		if c.Alias != "" {
			sourceSet := g.Aliases[c.Alias]
			if len(sourceSet) != 1 {
				return errors.New("unknown/colliding alias")
			}
			want := citations.SourceKey(sourceSet[0])
			var filtered []Binding
			for _, b := range candidates {
				if citations.SourceKey(b.Source) == want {
					filtered = append(filtered, b)
				}
			}
			candidates = filtered
		}
		unique := map[string]bool{}
		for _, b := range candidates {
			unique[citations.SourceKey(b.Source)] = true
		}
		if len(unique) != 1 {
			return errors.New("supra target source is ambiguous")
		}
		parent := candidates[0]
		// Supra does not silently inherit the old pinpoint.
		g.Bindings[key] = Binding{n.ID, j, parent.Source, c.Basis, c.Locator, c.Qualifications, parent.OriginNote}
	case "ibid":
		// Restrict to first citation in current note and exactly one fully
		// resolved citation in the immediately preceding note.
		if j != 0 || i == 0 || len(notes[i-1].Citations) != 1 {
			return errors.New("ibid governing citation not uniquely established")
		}
		parent, ok := g.Bindings[BindingKey{notes[i-1].ID, 0}]
		if !ok {
			return errors.New("ibid parent unresolved")
		}
		basis, locator := parent.Basis, parent.Locator
		if c.Locator != nil {
			basis, locator = c.Basis, c.Locator
		}
		g.Bindings[key] = Binding{n.ID, j, parent.Source, basis, locator, c.Qualifications, parent.OriginNote}
	default:
		return errors.New("unsupported unresolved root")
	}
	return nil
}

var legalLabels = map[string][2]string{"section": {"s", "ss"}, "article": {"art", "arts"}, "rule": {"rule", "rules"}}

// renderLocator serializes a binding's locator losslessly.
func renderLocator(b Binding) (string, error) {
	if b.Locator == nil {
		return "", nil
	}
	if b.Basis == "page" || b.Basis == "paragraph" {
		if b.Locator.Legal != nil {
			return "", errors.New("unsupported numeric binding")
		}
		var vals []string
		multiple := len(b.Locator.Ranges) > 1
		for _, r := range b.Locator.Ranges {
			if r.Lo <= 0 || r.Hi <= 0 {
				return "", errors.New("unsupported numeric binding")
			}
			if r.Hi < r.Lo {
				return "", errors.New("descending binding")
			}
			if r.Lo == r.Hi {
				vals = append(vals, strconv.Itoa(r.Lo))
			} else {
				multiple = true
				vals = append(vals, strconv.Itoa(r.Lo)+"–"+strconv.Itoa(r.Hi))
			}
		}
		label := ""
		if b.Basis != "page" {
			label = "para "
			if multiple {
				label = "paras "
			}
		}
		return " at " + label + strings.Join(vals, ", "), nil
	}
	if b.Basis == "section" && b.Locator.Legal != nil {
		loc := b.Locator.Legal
		if len(loc.Separators) != len(loc.Identifiers)-1 {
			return "", errors.New("incomplete legal locator")
		}
		labels, ok := legalLabels[loc.Unit]
		if !ok {
			return "", errors.New("locator needs a lossless renderer")
		}
		label := labels[0]
		if len(loc.Identifiers) > 1 {
			label = labels[1]
		}
		value := loc.Identifiers[0]
		for i, sep := range loc.Separators {
			value += sep + loc.Identifiers[i+1]
		}
		return ", " + label + " " + value, nil
	}
	return "", errors.New("locator needs a lossless renderer")
}

func citationsEqual(a, b []citations.Citation) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if !a[i].Equal(b[i]) {
			return false
		}
	}
	return true
}

// RepairIbidAfterReorder uses the old binding, never reinterpreting Ibid from
// its new neighbour. The caller supplies the validated new native-note
// ordering/labels and a fresh fingerprint of the old analyzed graph. Added
// unrelated notes are permitted; modifications to any retained citation
// invalidate the routine. ok=false means the existing Ibid still has the
// exact tested meaning. The text is a proposal, not a native mutation.
func RepairIbidAfterReorder(graph *Graph, current []Note, noteID, currentGraphFingerprint string) (string, bool, error) {
	if currentGraphFingerprint != graph.Fingerprint {
		return "", false, errors.New("stale reference graph")
	}
	newGraph, err := CompileGraph(current)
	if err != nil {
		return "", false, err
	}
	oldBy := map[string]Note{}
	for _, n := range graph.Notes {
		oldBy[n.ID] = n
	}
	for _, n := range current {
		if old, ok := oldBy[n.ID]; ok && !citationsEqual(n.Citations, old.Citations) {
			return "", false, errors.New("a retained citation changed; rebuild associations")
		}
	}
	old, ok := graph.Bindings[BindingKey{noteID, 0}]
	if !ok {
		return "", false, errors.New("old intended source is unresolved")
	}
	original := oldBy[noteID]
	if len(original.Citations) != 1 || original.Citations[0].Kind != "ibid" {
		return "", false, errors.New("not a single bound ibid note")
	}
	index := -1
	for i, n := range current {
		if n.ID == noteID {
			index = i
			break
		}
	}
	if index < 0 {
		return "", false, errors.New("note is absent from the current ordering")
	}
	if rebound, ok := newGraph.Bindings[BindingKey{noteID, 0}]; ok && rebound.Meaning() == old.Meaning() {
		return "", false, nil
	}
	oldKey := citations.SourceKey(old.Source)
	var target *Note
	for i := range current[:index] {
		n := current[i]
		for _, c := range n.Citations {
			if c.Source != nil && citations.SourceKey(c.Source) == oldKey {
				target = &n
				break
			}
		}
		if target != nil {
			break
		}
	}
	if target == nil {
		return "", false, errors.New("original source has no preceding full citation")
	}
	if !pytext.IsDecimalDigits(target.Label) {
		return "", false, errors.New("supra label is not ordinary Arabic numbering")
	}
	pointer, err := newGraph.PreservePointer(current[index].Scope, target.Label)
	if err != nil {
		return "", false, err
	}
	if pointer != target.ID {
		return "", false, errors.New("cross-scope numbering is ambiguous")
	}
	var aliases []string
	for k, v := range graph.Aliases {
		if len(v) == 1 && citations.SourceKey(v[0]) == oldKey {
			aliases = append(aliases, k)
		}
	}
	if len(aliases) == 0 {
		return "", false, errors.New("no unique declared short form for old source")
	}
	sort.Strings(aliases)
	locator, err := renderLocator(old)
	if err != nil {
		return "", false, err
	}
	suffix := ""
	if len(old.Qualifications) > 0 {
		suffix = " [" + strings.Join(old.Qualifications, ", ") + "]"
	}
	return aliases[0] + ", supra note " + target.Label + locator + suffix + ".", true, nil
}
