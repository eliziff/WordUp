package office

import (
	"bytes"
	"fmt"
	"sort"
	"strings"
)

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
		"http://schemas.microsoft.com/office/2006/01/customui": "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility",
		"http://schemas.microsoft.com/office/2009/07/customui": "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility",
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
