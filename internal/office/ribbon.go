package office

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"sort"
	"strings"
)

const (
	ribbon2007Namespace = "http://schemas.microsoft.com/office/2006/01/customui"
	ribbon2010Namespace = "http://schemas.microsoft.com/office/2009/07/customui"
)

type ribbonNode struct {
	children []int
}

func ribbonTree(spans []XMLSpan) []ribbonNode {
	nodes := make([]ribbonNode, len(spans))
	stack := []int{}
	for i, span := range spans {
		if len(stack) > span.Depth {
			stack = stack[:span.Depth]
		}
		if len(stack) > 0 {
			nodes[stack[len(stack)-1]].children = append(nodes[stack[len(stack)-1]].children, i)
		}
		stack = append(stack, i)
	}
	return nodes
}

func ribbonDocument(data []byte, label string) ([]XMLSpan, []ribbonNode, error) {
	spans, err := XMLSpans(data)
	if err != nil {
		return nil, nil, fmt.Errorf("%s: %w", label, err)
	}
	if len(spans) == 0 || spans[0].Name.Local != "customUI" {
		return nil, nil, fmt.Errorf("%s: root must be customUI", label)
	}
	if spans[0].Name.Space != ribbon2007Namespace && spans[0].Name.Space != ribbon2010Namespace {
		return nil, nil, fmt.Errorf("%s: unsupported RibbonX namespace", label)
	}
	ids := map[string]bool{}
	for _, span := range spans {
		for _, attr := range span.Attr {
			if attr.Name.Space != "" || (attr.Name.Local != "id" && attr.Name.Local != "idQ") || attr.Value == "" {
				continue
			}
			key := attr.Name.Local + "\x00" + attr.Value
			if ids[key] {
				return nil, nil, fmt.Errorf("%s: duplicate Ribbon %s %q", label, attr.Name.Local, attr.Value)
			}
			ids[key] = true
		}
	}
	return spans, ribbonTree(spans), nil
}

func ribbonAttrKey(space, local string) string {
	return space + "\x00" + local
}

func ribbonNamespace(attr xml.Attr) (string, bool) {
	if attr.Name.Space == "xmlns" {
		return attr.Name.Local, true
	}
	if attr.Name.Space == "" && attr.Name.Local == "xmlns" {
		return "", true
	}
	return "", false
}

func ribbonNamespaceMap(span XMLSpan) map[string]string {
	result := map[string]string{}
	for _, attr := range span.Attr {
		if prefix, ok := ribbonNamespace(attr); ok {
			result[prefix] = attr.Value
		}
	}
	return result
}

func ribbonRootAttributes(base []byte, baseRoot, fragmentRoot XMLSpan) ([]byte, error) {
	baseAttrs := map[string]xml.Attr{}
	for _, attr := range baseRoot.Attr {
		if _, ok := ribbonNamespace(attr); ok {
			continue
		}
		baseAttrs[ribbonAttrKey(attr.Name.Space, attr.Name.Local)] = attr
	}
	baseNamespaces := ribbonNamespaceMap(baseRoot)
	additions := []string{}
	for _, attr := range fragmentRoot.Attr {
		if prefix, ok := ribbonNamespace(attr); ok {
			if prior, exists := baseNamespaces[prefix]; exists {
				if prior != attr.Value {
					return nil, fmt.Errorf("Ribbon namespace prefix %q is bound to both %q and %q", prefix, prior, attr.Value)
				}
				continue
			}
			if prefix == "" {
				return nil, fmt.Errorf("Ribbon fragment changes the default namespace")
			}
			additions = append(additions, ` xmlns:`+prefix+`="`+Esc(attr.Value)+`"`)
			continue
		}
		key := ribbonAttrKey(attr.Name.Space, attr.Name.Local)
		if prior, exists := baseAttrs[key]; exists {
			if prior.Value != attr.Value {
				return nil, fmt.Errorf("Ribbon root attribute %s conflicts (%q versus %q)", attr.Name.Local, prior.Value, attr.Value)
			}
			continue
		}
		if attr.Name.Space != "" {
			return nil, fmt.Errorf("Ribbon fragment adds unsupported namespaced root attribute %s", attr.Name.Local)
		}
		additions = append(additions, ` `+attr.Name.Local+`="`+Esc(attr.Value)+`"`)
	}
	if len(additions) == 0 {
		return base, nil
	}
	insertAt := baseRoot.OpenEnd - 1
	if insertAt > baseRoot.Start && base[insertAt-1] == '/' {
		insertAt--
	}
	var out bytes.Buffer
	out.Grow(len(base) + len(strings.Join(additions, "")))
	out.Write(base[:insertAt])
	out.WriteString(strings.Join(additions, ""))
	out.Write(base[insertAt:])
	return out.Bytes(), nil
}

func ribbonIdentity(parent string, span XMLSpan) (string, bool) {
	// These elements are singletons in their parent according to the RibbonX
	// schema. Merging them by local name lets independent fragments contribute
	// children without rewriting the existing XML.
	switch parent + "\x00" + span.Name.Local {
	case "customUI\x00commands", "customUI\x00ribbon", "customUI\x00backstage", "customUI\x00contextMenus",
		"ribbon\x00tabs", "ribbon\x00qat", "ribbon\x00quickAccessToolbar":
		return "singleton\x00" + span.Name.Local, true
	}
	for _, attrName := range []string{"id", "idQ", "idMso"} {
		if value := span.Attribute("", attrName); value != "" {
			return attrName + "\x00" + value, true
		}
	}
	return "", false
}

var ribbonChildOrder = map[string]map[string]int{
	"customUI": {"commands": 0, "ribbon": 1, "backstage": 2, "contextMenus": 3},
	"ribbon":   {"qat": 0, "quickAccessToolbar": 0, "tabs": 1},
	"tabs":     {"tab": 0, "tabSet": 0},
}

func ribbonRank(parent, child string) (int, bool) {
	ranks, ok := ribbonChildOrder[parent]
	if !ok {
		return 0, false
	}
	rank, ok := ranks[child]
	return rank, ok
}

func ribbonIDs(spans []XMLSpan) map[string]bool {
	ids := map[string]bool{}
	for _, span := range spans {
		for _, attrName := range []string{"id", "idQ"} {
			if value := span.Attribute("", attrName); value != "" {
				ids[attrName+"\x00"+value] = true
			}
		}
	}
	return ids
}

func rejectRibbonIDs(spans []XMLSpan, nodes []ribbonNode, index int, baseIDs map[string]bool) error {
	span := spans[index]
	for _, attrName := range []string{"id", "idQ"} {
		if value := span.Attribute("", attrName); value != "" && baseIDs[attrName+"\x00"+value] {
			return fmt.Errorf("Ribbon %s %q collides with an existing control", attrName, value)
		}
	}
	for _, child := range nodes[index].children {
		if err := rejectRibbonIDs(spans, nodes, child, baseIDs); err != nil {
			return err
		}
	}
	return nil
}

func ribbonCompatibleAttributes(base, fragment XMLSpan) error {
	baseAttrs := map[string]xml.Attr{}
	for _, attr := range base.Attr {
		if _, ok := ribbonNamespace(attr); !ok {
			baseAttrs[ribbonAttrKey(attr.Name.Space, attr.Name.Local)] = attr
		}
	}
	for _, attr := range fragment.Attr {
		if _, ok := ribbonNamespace(attr); ok {
			continue
		}
		key := ribbonAttrKey(attr.Name.Space, attr.Name.Local)
		prior, exists := baseAttrs[key]
		if !exists {
			return fmt.Errorf("Ribbon merge would add %s=%q to existing %s", attr.Name.Local, attr.Value, base.Name.Local)
		}
		if prior.Value != attr.Value {
			return fmt.Errorf("Ribbon %s attribute %s conflicts (%q versus %q)", base.Name.Local, attr.Name.Local, prior.Value, attr.Value)
		}
	}
	return nil
}

func ribbonRawName(data []byte, span XMLSpan) string {
	start := span.Start + 1
	end := start
	for end < span.OpenEnd && !strings.ContainsRune(" \t\r\n/>", rune(data[end])) {
		end++
	}
	return string(data[start:end])
}

func mergeRibbonNode(base, fragment []byte, baseSpans []XMLSpan, baseNodes []ribbonNode, baseIndex int, fragmentSpans []XMLSpan, fragmentNodes []ribbonNode, fragmentIndex int, baseIDs map[string]bool) ([]byte, error) {
	baseSpan, fragmentSpan := baseSpans[baseIndex], fragmentSpans[fragmentIndex]
	if baseSpan.Name != fragmentSpan.Name {
		return nil, fmt.Errorf("Ribbon merge cannot combine %s with %s", baseSpan.Name.Local, fragmentSpan.Name.Local)
	}
	if err := ribbonCompatibleAttributes(baseSpan, fragmentSpan); err != nil {
		return nil, err
	}
	fragmentChildren := fragmentNodes[fragmentIndex].children
	if len(fragmentChildren) == 0 {
		return append([]byte(nil), base[baseSpan.Start:baseSpan.End]...), nil
	}
	baseChildren := baseNodes[baseIndex].children
	replacements := map[int][]byte{}
	insertions := map[int][][]byte{}
	for _, fragmentChild := range fragmentChildren {
		fragmentChildSpan := fragmentSpans[fragmentChild]
		fragmentKey, fragmentKeyed := ribbonIdentity(baseSpan.Name.Local, fragmentChildSpan)
		match := -1
		if fragmentKeyed {
			for _, baseChild := range baseChildren {
				baseKey, baseKeyed := ribbonIdentity(baseSpan.Name.Local, baseSpans[baseChild])
				if baseKeyed && baseKey == fragmentKey {
					match = baseChild
					break
				}
			}
		}
		if match >= 0 {
			merged, err := mergeRibbonNode(base, fragment, baseSpans, baseNodes, match, fragmentSpans, fragmentNodes, fragmentChild, baseIDs)
			if err != nil {
				return nil, err
			}
			replacements[match] = merged
			continue
		}
		if err := rejectRibbonIDs(fragmentSpans, fragmentNodes, fragmentChild, baseIDs); err != nil {
			return nil, err
		}
		position := len(baseChildren)
		if fragmentRank, ok := ribbonRank(baseSpan.Name.Local, fragmentChildSpan.Name.Local); ok {
			for i, baseChild := range baseChildren {
				if baseRank, ok := ribbonRank(baseSpan.Name.Local, baseSpans[baseChild].Name.Local); ok && baseRank > fragmentRank {
					position = i
					break
				}
			}
		}
		insertions[position] = append(insertions[position], append([]byte(nil), fragment[fragmentChildSpan.Start:fragmentChildSpan.End]...))
	}
	open := base[baseSpan.Start:baseSpan.OpenEnd]
	selfClosing := len(open) >= 2 && bytes.Equal(open[len(open)-2:], []byte("/>"))
	if selfClosing {
		var out bytes.Buffer
		out.Grow(baseSpan.End - baseSpan.Start + 64)
		out.Write(open[:len(open)-2])
		out.WriteByte('>')
		for _, addition := range insertions[len(baseChildren)] {
			out.Write(addition)
		}
		out.WriteString("</")
		out.WriteString(ribbonRawName(base, baseSpan))
		out.WriteByte('>')
		return out.Bytes(), nil
	}
	var out bytes.Buffer
	out.Grow(baseSpan.End - baseSpan.Start + 128)
	out.Write(open)
	cursor := baseSpan.OpenEnd
	for i, child := range baseChildren {
		childSpan := baseSpans[child]
		out.Write(base[cursor:childSpan.Start])
		for _, addition := range insertions[i] {
			out.Write(addition)
		}
		if replacement, ok := replacements[child]; ok {
			out.Write(replacement)
		} else {
			out.Write(base[childSpan.Start:childSpan.End])
		}
		cursor = childSpan.End
	}
	for _, addition := range insertions[len(baseChildren)] {
		out.Write(addition)
	}
	out.Write(base[cursor:baseSpan.CloseStart])
	out.Write(base[baseSpan.CloseStart:baseSpan.End])
	return out.Bytes(), nil
}

// MergeRibbonXML composes a RibbonX fragment into an existing customUI part.
// It edits only the existing byte ranges and inserts the fragment's elements;
// namespace prefixes, unknown attributes, comments and whitespace in the base
// remain untouched. Existing controls are never silently overwritten.
func MergeRibbonXML(base, fragment []byte) ([]byte, error) {
	baseSpans, _, err := ribbonDocument(base, "base Ribbon")
	if err != nil {
		return nil, err
	}
	fragmentSpans, fragmentNodes, err := ribbonDocument(fragment, "Ribbon fragment")
	if err != nil {
		return nil, err
	}
	if baseSpans[0].Name.Space != fragmentSpans[0].Name.Space {
		return nil, fmt.Errorf("Ribbon namespace mismatch: %s versus %s", baseSpans[0].Name.Space, fragmentSpans[0].Name.Space)
	}
	base, err = ribbonRootAttributes(base, baseSpans[0], fragmentSpans[0])
	if err != nil {
		return nil, err
	}
	baseSpans, baseNodes, err := ribbonDocument(base, "base Ribbon")
	if err != nil {
		return nil, err
	}
	merged, err := mergeRibbonNode(base, fragment, baseSpans, baseNodes, 0, fragmentSpans, fragmentNodes, 0, ribbonIDs(baseSpans))
	if err != nil {
		return nil, err
	}
	if _, err := RibbonCallbacks(merged); err != nil {
		return nil, fmt.Errorf("merged Ribbon callbacks: %w", err)
	}
	return merged, nil
}

// MergeRibbon adds a new Ribbon part or composes a fragment into an existing
// part, then connects it to the package relationships. The package is changed
// only after every validation step succeeds.
func (p *Package) MergeRibbon(part string, fragment []byte) error {
	if !SafePart(part) || !strings.HasSuffix(strings.ToLower(part), ".xml") {
		return fmt.Errorf("invalid Ribbon part path %q", part)
	}
	work := p.Clone()
	current, exists := work.Files[part]
	if !exists {
		if _, _, err := ribbonDocument(fragment, "Ribbon fragment"); err != nil {
			return err
		}
		if _, err := RibbonCallbacks(fragment); err != nil {
			return fmt.Errorf("Ribbon fragment callbacks: %w", err)
		}
		work.Files[part] = append([]byte(nil), fragment...)
	} else {
		merged, err := MergeRibbonXML(current, fragment)
		if err != nil {
			return err
		}
		work.Files[part] = merged
	}
	if err := work.ContentType(part, "application/xml"); err != nil {
		return err
	}
	if err := work.ConnectRibbons(); err != nil {
		return err
	}
	*p = *work
	return nil
}

// ConnectRibbons makes newly authored RibbonX parts discoverable by Office.
// Existing relationship IDs and arbitrary part paths are preserved.
func (p *Package) ConnectRibbons() error {
	work := p.Clone()
	if err := work.connectRibbons(); err != nil {
		return err
	}
	*p = *work
	return nil
}

func (p *Package) connectRibbons() error {
	types := map[string]string{
		ribbon2007Namespace: "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility",
		ribbon2010Namespace: "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility",
	}
	ids := map[string]bool{}
	linked := map[string]string{}
	if b := p.Files["_rels/.rels"]; len(b) > 0 {
		spans, err := XMLSpans(b)
		if err != nil {
			return err
		}
		for _, s := range spans {
			if s.Name.Local == "Relationship" {
				ids[s.Attribute("", "Id")] = true
				linked[s.Attribute("", "Type")] = strings.TrimPrefix(s.Attribute("", "Target"), "/")
			}
		}
	}
	names := []string{}
	for name := range p.Files {
		if strings.HasSuffix(strings.ToLower(name), ".xml") {
			names = append(names, name)
		}
	}
	sort.Strings(names)
	seen := map[string]string{}
	for _, name := range names {
		if !bytes.Contains(p.Files[name], []byte("customUI")) {
			continue
		}
		spans, err := XMLSpans(p.Files[name])
		if err != nil {
			return fmt.Errorf("%s: %w", name, err)
		}
		if len(spans) == 0 || spans[0].Name.Local != "customUI" {
			continue
		}
		typ := types[spans[0].Name.Space]
		if typ == "" {
			return fmt.Errorf("%s: unsupported RibbonX namespace", name)
		}
		if prior := seen[typ]; prior != "" {
			return fmt.Errorf("multiple RibbonX parts for the same Office version: %s and %s", prior, name)
		}
		seen[typ] = name
		if prior := linked[typ]; prior != "" {
			if prior != name {
				return fmt.Errorf("RibbonX relationship targets %s instead of %s", prior, name)
			}
			continue
		}
		id := "rIdWordUpRibbon"
		for i := 1; ids[id]; i++ {
			id = fmt.Sprintf("rIdWordUpRibbon%d", i)
		}
		ids[id] = true
		if err = p.Relationship("", id, typ, name, ""); err != nil {
			return err
		}
		if err = p.ContentType(name, "application/xml"); err != nil {
			return err
		}
	}
	return nil
}
