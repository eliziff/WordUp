package ooxml

import (
	"bytes"
	"errors"
	"sort"
	"strings"

	"github.com/eliziff/WordUp/internal/alrmap/citations"
	"github.com/eliziff/WordUp/internal/alrmap/internal/pytext"
	"github.com/eliziff/WordUp/internal/alrmap/kernels"
)

// BoundPlan is a plan bound to one container of one exact part snapshot.
// Registry is only needed by the title family, whose guards are re-derived
// from a full parse at write time.
type BoundPlan struct {
	Fingerprint string
	Container   string
	Family      string
	Plan        kernels.Plan
	Registry    *citations.Registry
}

// Bind validates the whole batch of edits against the container's derived
// guards before binding it to the snapshot.
func Bind(scanned *Scan, container string, edits []kernels.Edit, family, authority string) (*BoundPlan, error) {
	c, err := scanned.Container(container)
	if err != nil {
		return nil, err
	}
	guards, err := c.Guards(family)
	if err != nil {
		return nil, err
	}
	plan := kernels.Plan{Digest: kernels.Digest(c.Text), Edits: append([]kernels.Edit(nil), edits...), Protected: guards,
		Authorized: strings.TrimSpace(authority) != "", AuthorizationBasis: authority}
	if _, err := kernels.ApplyPlan(c.Text, plan); err != nil {
		return nil, err
	}
	return &BoundPlan{Fingerprint: scanned.Fingerprint(), Container: container, Family: family, Plan: plan}, nil
}

// ByteWrite is one exact original byte slice replaced in the output copy.
type ByteWrite struct {
	Start, End int
	Before     []byte
	After      []byte
}

// Prepare maps all text edits of a bound plan onto original text nodes
// without reparsing or mutating the part. It returns the expected container
// text and the replacement text per node.
func Prepare(scanned *Scan, bound *BoundPlan) (string, map[*Node]string, error) {
	if bound.Fingerprint != scanned.Fingerprint() {
		return "", nil, errors.New("stale XML/styles/part")
	}
	c, err := scanned.Container(bound.Container)
	if err != nil {
		return "", nil, err
	}
	var guards []kernels.Span
	if bound.Family == "title" {
		guards, err = TitleGuards(scanned, c, bound)
	} else {
		guards, err = c.Guards(bound.Family)
	}
	if err != nil {
		return "", nil, err
	}
	checked := kernels.Plan{Digest: bound.Plan.Digest, Edits: bound.Plan.Edits, Protected: guards,
		Authorized: bound.Plan.Authorized, AuthorizationBasis: bound.Plan.AuthorizationBasis}
	expected, err := kernels.ApplyPlan(c.Text, checked)
	if err != nil {
		return "", nil, err
	}
	replacements := map[*Node]string{}
	edits := append([]kernels.Edit(nil), checked.Edits...)
	sort.SliceStable(edits, func(i, j int) bool { return edits[i].Start > edits[j].Start })
	for _, e := range edits {
		var touched []Atom
		for _, a := range c.Atoms {
			if e.Start == e.End {
				if a.Start < e.Start && e.Start <= a.End {
					touched = append(touched, a)
				}
			} else if a.Start < e.End && e.Start < a.End {
				touched = append(touched, a)
			}
		}
		if len(touched) == 0 {
			return "", nil, errors.New("edit is not mapped to text atoms")
		}
		signatures := map[string]bool{}
		for _, a := range c.Atoms {
			if a.Start < e.ReadEnd && e.ReadStart < a.End {
				signatures[a.Signature] = true
			}
		}
		if e.New != "" && len(signatures) != 1 {
			return "", nil, errors.New("mixed run formatting requires a native move adapter")
		}
		covered := 0
		for _, a := range touched {
			covered += min(a.End, e.End) - max(a.Start, e.Start)
		}
		if covered != e.End-e.Start {
			return "", nil, errors.New("write crosses a structural text gap")
		}
		for i, atom := range touched {
			old, ok := replacements[atom.Node]
			if !ok {
				old = atom.Node.Text()
			}
			oldRunes := []rune(old)
			lo := max(e.Start, atom.Start) - atom.Start
			hi := min(e.End, atom.End) - atom.Start
			insert := ""
			if i == 0 {
				insert = e.New
			}
			replacements[atom.Node] = string(oldRunes[:lo]) + insert + string(oldRunes[hi:])
		}
	}
	for node, new := range replacements {
		if _, err := encodeText(node, new); err != nil {
			return "", nil, err
		}
	}
	return expected, replacements, nil
}

func encodeText(n *Node, new string) ([]byte, error) {
	if n.OpenEnd >= n.CloseStart {
		return nil, errors.New("self-closing/empty text node")
	}
	xmlSpace := ""
	for inherited := n; inherited != nil; inherited = inherited.Parent() {
		if v, ok := inherited.Attrs[XML+"}space"]; ok {
			xmlSpace = v
			break
		}
	}
	if new != "" && xmlSpace != "preserve" {
		runes := []rune(new)
		if pytext.IsSpace(runes[0]) || pytext.IsSpace(runes[len(runes)-1]) {
			return nil, errors.New("rewrite would require an xml:space attribute change")
		}
	}
	escaped := strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;").Replace(new)
	return []byte(strings.ReplaceAll(escaped, "\r", "&#13;")), nil
}

// ApplyBound applies one bound plan (a one-element batch).
func ApplyBound(scanned *Scan, bound *BoundPlan) ([]byte, []ByteWrite, error) {
	return ApplyBatch(scanned, []*BoundPlan{bound})
}

// ApplyBatch verifies disjoint original-snapshot plans and reconstructs and
// checks the XML copy just once. Bytes outside changed text-node values are
// retained exactly. This creates a reference XML copy; it never fabricates
// native tracked changes or Word passes.
func ApplyBatch(scanned *Scan, bounds []*BoundPlan) ([]byte, []ByteWrite, error) {
	keys := map[string]bool{}
	for _, b := range bounds {
		keys[b.Container] = true
	}
	if len(keys) != len(bounds) {
		return nil, nil, errors.New("merge same-container plans before committing")
	}
	containers := map[string]*Container{}
	for _, c := range scanned.Containers {
		containers[c.Key] = c
	}
	replacements := map[*Node]string{}
	expected := map[string]string{}
	var order []*Node
	for _, bound := range bounds {
		text, part, err := Prepare(scanned, bound)
		if err != nil {
			return nil, nil, err
		}
		expected[bound.Container] = text
		for n := range part {
			if _, dup := replacements[n]; dup {
				return nil, nil, errors.New("overlapping text-node ownership")
			}
		}
		for n, v := range part {
			replacements[n] = v
			order = append(order, n)
		}
	}
	var writes []ByteWrite
	for _, n := range order {
		new := replacements[n]
		if new == n.Text() {
			continue
		}
		encoded, err := encodeText(n, new)
		if err != nil {
			return nil, nil, err
		}
		writes = append(writes, ByteWrite{n.OpenEnd, n.CloseStart, scanned.Raw[n.OpenEnd:n.CloseStart], encoded})
	}
	sort.SliceStable(writes, func(i, j int) bool { return writes[i].Start < writes[j].Start })
	var out bytes.Buffer
	cursor := 0
	for _, w := range writes {
		if w.Start < cursor || !bytes.Equal(scanned.Raw[w.Start:w.End], w.Before) {
			return nil, nil, errors.New("stale/overlapping byte slice")
		}
		out.Write(scanned.Raw[cursor:w.Start])
		out.Write(w.After)
		cursor = w.End
	}
	out.Write(scanned.Raw[cursor:])
	result := out.Bytes()
	after, err := ScanPart(result, scanned.Styles, scanned.Part)
	if err != nil {
		return nil, nil, err
	}
	for _, ac := range after.Containers {
		want, ok := expected[ac.Key]
		if !ok {
			before, known := containers[ac.Key]
			if !known {
				return nil, nil, &kernels.InvariantError{Reason: "XML projection differs from planned text"}
			}
			want = before.Text
		}
		if ac.Text != want {
			return nil, nil, &kernels.InvariantError{Reason: "XML projection differs from planned text"}
		}
	}
	// Full tree equality with only the explicitly changed text nodes substituted.
	oldNodes := append([]*Node{scanned.Root}, scanned.Root.Descendants()...)
	newNodes := append([]*Node{after.Root}, after.Root.Descendants()...)
	if len(oldNodes) != len(newNodes) {
		return nil, nil, &kernels.InvariantError{Reason: "XML structure changed"}
	}
	byStart := map[int]string{}
	for n, v := range replacements {
		byStart[n.Start] = v
	}
	for i := range oldNodes {
		old, new := oldNodes[i], newNodes[i]
		if old.Tag != new.Tag || len(old.Children) != len(new.Children) || !attrsEqual(old.Attrs, new.Attrs) {
			return nil, nil, &kernels.InvariantError{Reason: "XML structure/attributes changed"}
		}
		want, ok := byStart[old.Start]
		if !ok {
			want = old.Text()
		}
		if new.Text() != want {
			return nil, nil, &kernels.InvariantError{Reason: "unowned XML text changed"}
		}
	}
	return result, writes, nil
}

func attrsEqual(a, b map[string]string) bool {
	if len(a) != len(b) {
		return false
	}
	for k, v := range a {
		if bv, ok := b[k]; !ok || bv != v {
			return false
		}
	}
	return true
}

// PlanNote bridges full citation parsing to the actual XML inventory, not
// fake masks. It reads the entire single-paragraph note, not a matched
// sentence inside prose, and permits only the native leading note mark and
// its initial tab as structural exceptions. Third-party fields anywhere in
// the production veto this bridge, even if the eventual punctuation edit
// would be outside the field result. A nil enabled set means every switch.
func PlanNote(scanned *Scan, container string, registry *citations.Registry, enabled kernels.Switches) (*BoundPlan, citations.Result, error) {
	c, err := scanned.Container(container)
	if err != nil {
		return nil, citations.Result{}, err
	}
	start, end, err := noteBounds(c, "one complete native note paragraph required",
		"citation production depends on opaque native structure", "note paragraph has a protected structural role")
	if err != nil {
		return nil, citations.Result{}, err
	}
	// Known quote/list/indented roles remain a veto here. Native FootnoteText
	// styles with indentation may therefore be conservatively excluded; the
	// scanner never promotes a quotation into authored citation by a regex.
	if registry == nil {
		registry = &citations.Registry{}
	}
	if enabled == nil {
		enabled = kernels.NewSwitches(citations.AllSwitches...)
	}
	result, err := citations.NormalizeNote(string([]rune(c.Text)[start:end]), registry, citations.Options{Enabled: enabled})
	if err != nil {
		return nil, citations.Result{}, err
	}
	edits := make([]kernels.Edit, len(result.Edits))
	for i, e := range result.Edits {
		e.Start += start
		e.End += start
		e.ReadStart += start
		e.ReadEnd += start
		edits[i] = e
	}
	bound, err := Bind(scanned, c.Key, edits, "citation", "complete native-note syntax production")
	if err != nil {
		return nil, citations.Result{}, err
	}
	return bound, result, nil
}

// noteBounds locates the trimmed single native note paragraph and applies
// the structural vetoes shared by the citation bridge and the title grammar.
func noteBounds(c *Container, shapeMsg, opaqueMsg, roleMsg string) (int, int, error) {
	if !(strings.HasPrefix(c.Key, "footnote:") || strings.HasPrefix(c.Key, "endnote:")) || len(c.Paragraphs) != 1 {
		return 0, 0, errors.New(shapeMsg)
	}
	runes := []rune(c.Text)
	start, end := c.Paragraphs[0][0], c.Paragraphs[0][1]
	for start < end && (runes[start] == ' ' || runes[start] == '\t') {
		start++
	}
	for end > start && runes[end-1] == ' ' {
		end--
	}
	for _, g := range c.Hard {
		if g.Overlaps(start, end) {
			return 0, 0, errors.New(opaqueMsg)
		}
	}
	for _, g := range c.Lexical {
		if g.Reason == "paragraph-role" && g.Overlaps(start, end) {
			return 0, 0, errors.New(roleMsg)
		}
	}
	return start, end, nil
}
