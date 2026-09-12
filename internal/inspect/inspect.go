// Package inspect returns factual source/package observations, not a compiler verdict.
package inspect

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"wordwright.local/internal/office"
	"wordwright.local/internal/project"
)

type Symbol struct {
	Module      string `json:"module"`
	Name        string `json:"name"`
	Kind        string `json:"kind"`
	Line        int    `json:"line"`
	Declaration string `json:"declaration"`
}

var procedure = regexp.MustCompile(`(?i)^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+([\pL_][\pL\pM\pN_]*)\b`)

func CommentFree(line string) string {
	quoted := false
	for i := 0; i < len(line); i++ {
		if line[i] == '"' {
			if quoted && i+1 < len(line) && line[i+1] == '"' {
				i++
				continue
			}
			quoted = !quoted
		}
		if line[i] == '\'' && !quoted {
			return line[:i]
		}
	}
	if strings.HasPrefix(strings.ToLower(strings.TrimSpace(line)), "rem ") {
		return ""
	}
	return line
}
func Symbols(name, source string) []Symbol {
	out := []Symbol{}
	for i, l := range strings.Split(strings.ReplaceAll(source, "\r\n", "\n"), "\n") {
		l = CommentFree(l)
		if m := procedure.FindStringSubmatch(l); m != nil {
			out = append(out, Symbol{name, m[2], m[1], i + 1, strings.TrimSpace(l)})
		}
	}
	return out
}
func Artifact(file string) (map[string]any, error) {
	abs, e := filepath.Abs(file)
	if e != nil {
		return nil, e
	}
	b, e := project.Read(filepath.Dir(abs), filepath.Base(abs))
	if e != nil {
		return nil, e
	}
	p, e := office.ReadPackage(b)
	if e != nil {
		return nil, e
	}
	if e = p.Validate(); e != nil {
		return nil, e
	}
	return Package(p), nil
}
func Package(p *office.Package) map[string]any {
	out := office.Catalog(p)
	out["sha256"] = office.Hash(p.Original)
	out["bytes"] = len(p.Original)
	out["signed"] = p.HasSignatures()
	out["package_validated"] = true
	out["word_executed"] = false
	parts := []map[string]any{}
	names := []string{}
	for n := range p.Files {
		names = append(names, n)
	}
	sort.Strings(names)
	for _, n := range names {
		parts = append(parts, map[string]any{"part": n, "bytes": len(p.Files[n]), "sha256": office.Hash(p.Files[n])})
	}
	out["parts"] = parts
	mods := []map[string]any{}
	forms := []any{}
	symbols := []Symbol{}
	lines := 0
	if b := p.Files["word/vbaProject.bin"]; b != nil {
		v, e := office.ReadVBA(b)
		if e != nil {
			out["vba_error"] = e.Error()
			return out
		}
		out["project_name"] = v.Name
		out["codepage"] = v.Codepage
		for _, m := range v.Modules {
			n := strings.Count(m.Source, "\n") + 1
			lines += n
			mods = append(mods, map[string]any{"name": m.Name, "kind": m.Kind, "lines": n, "source_sha256": office.Hash([]byte(m.Source))})
			symbols = append(symbols, Symbols(m.Name, m.Source)...)
			if m.Kind == "form" {
				f, e := office.ReadForm(v.CFB, m.Name, v.Codepage)
				if e != nil {
					forms = append(forms, map[string]any{"name": m.Name, "opaque": true, "error": e.Error()})
				} else {
					forms = append(forms, f.Design())
				}
			}
		}
	}
	out["modules"] = mods
	out["forms"] = forms
	out["symbols"] = symbols
	out["source_lines"] = lines
	return out
}

// TextObservations retains the actual property XML and named-style IDs. An
// agent may infer a recipe; unresolved inheritance is never labelled resolved.
func TextObservations(p *office.Package) ([]map[string]any, error) {
	b := p.Files["word/document.xml"]
	spans, e := office.XMLSpans(b)
	if e != nil {
		return nil, e
	}
	out := []map[string]any{}
	for index, s := range spans {
		if s.Name.Space != office.W || s.Name.Local != "p" {
			continue
		}
		text := strings.Builder{}
		style := ""
		props := ""
		runs := []map[string]any{}
		for _, x := range spans[index+1:] {
			if x.Start >= s.End {
				break
			}
			if x.Start <= s.Start || x.End >= s.End {
				continue
			}
			if x.Name.Local == "pStyle" {
				style = x.Attribute(office.W, "val")
			}
			if x.Name.Local == "pPr" {
				props = string(b[x.Start:x.End])
			}
			if x.Name.Local == "t" {
				var tmp struct {
					Text string `xml:",chardata"`
				}
				if e = xml.Unmarshal(b[x.Start:x.End], &tmp); e != nil {
					return nil, e
				}
				text.WriteString(tmp.Text)
			}
			if x.Name.Local == "rPr" {
				runs = append(runs, map[string]any{"xml": string(b[x.Start:x.End])})
			}
		}
		out = append(out, map[string]any{"text": text.String(), "style_id": style, "paragraph_properties_xml": props, "run_properties": runs})
		if len(out) > 10000 {
			return nil, fmt.Errorf("paragraph observation budget exceeded")
		}
	}
	return out, nil
}
func StyleReference(file string) (map[string]any, error) {
	b, e := project.Read(filepath.Dir(file), filepath.Base(file))
	if e != nil {
		return nil, e
	}
	p, e := office.ReadPackage(b)
	if e != nil {
		return nil, e
	}
	out := office.Catalog(p)
	text, e := TextObservations(p)
	if e != nil {
		return nil, e
	}
	out["paragraph_observations"] = text
	out["source_sha256"] = office.Hash(b)
	out["effective_layout_verified"] = false
	return out, nil
}
func Check(w *project.Workspace) (map[string]any, error) {
	files, e := w.SourceFiles()
	if e != nil {
		return nil, e
	}
	symbols := []Symbol{}
	byName := map[string]bool{}
	diagnostics := []map[string]any{}
	for n, b := range files {
		if strings.HasPrefix(n, "vba/") {
			name := strings.TrimSuffix(filepath.Base(n), filepath.Ext(n))
			s := Symbols(name, string(b))
			symbols = append(symbols, s...)
			for _, x := range s {
				byName[strings.ToLower(x.Name)] = true
			}
			if bytes.Contains(b, []byte{0}) {
				diagnostics = append(diagnostics, map[string]any{"severity": "error", "file": n, "message": "NUL byte in VBA source"})
			}
		}
	}
	for n, b := range files {
		if !strings.HasPrefix(n, "package/") || !strings.HasSuffix(n, ".xml") {
			continue
		}
		spans, e := office.XMLSpans(b)
		if e != nil {
			diagnostics = append(diagnostics, map[string]any{"severity": "error", "file": n, "message": e.Error()})
			continue
		}
		if !strings.Contains(strings.ToLower(n), "customui") {
			continue
		}
		ids := map[string]bool{}
		for _, s := range spans {
			for _, a := range s.Attr {
				if a.Name.Local == "id" {
					if ids[a.Value] {
						diagnostics = append(diagnostics, map[string]any{"severity": "error", "file": n, "message": "duplicate Ribbon id " + a.Value})
					}
					ids[a.Value] = true
				}
				if strings.HasPrefix(a.Name.Local, "get") || strings.HasPrefix(a.Name.Local, "on") {
					cb := strings.ToLower(a.Value)
					if i := strings.LastIndex(cb, "."); i >= 0 {
						cb = cb[i+1:]
					}
					if !byName[cb] {
						diagnostics = append(diagnostics, map[string]any{"severity": "warning", "file": n, "message": "callback not lexically found: " + a.Value + "; dynamic/external routing needs native verification"})
					}
				}
			}
		}
	}
	sort.Slice(symbols, func(i, j int) bool {
		if symbols[i].Module == symbols[j].Module {
			return symbols[i].Line < symbols[j].Line
		}
		return symbols[i].Module < symbols[j].Module
	})
	return map[string]any{"symbols": symbols, "diagnostics": diagnostics, "vba_compiled": false, "word_executed": false, "coverage": "lexical symbols, XML syntax, duplicate ribbon IDs and unresolved callback warnings; not full VBA or RibbonX validation"}, nil
}
