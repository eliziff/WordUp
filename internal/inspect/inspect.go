// Package inspect returns factual source/package observations, not a compiler verdict.
package inspect

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/vbaparse"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"unicode/utf16"
)

func textElement(b []byte) (string, error) {
	start, end := bytes.IndexByte(b, '>'), bytes.LastIndexByte(b, '<')
	if start < 0 || end <= start {
		return "", fmt.Errorf("invalid text element")
	}
	inner := b[start+1 : end]
	if !bytes.ContainsRune(inner, '&') {
		return string(inner), nil
	}
	var value struct {
		Text string `xml:",chardata"`
	}
	if err := xml.Unmarshal(b, &value); err != nil {
		return "", err
	}
	return value.Text, nil
}

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

// TextObservations retains the actual property XML and named-style IDs for
// the main document story. An agent may infer a recipe; unresolved inheritance
// is never labelled resolved.
func TextObservations(p *office.Package) ([]map[string]any, error) {
	return textObservationsPart(p, "word/document.xml")
}

// textObservationsPart reads any Word text-bearing story without rewriting its
// XML. Keeping the part name in every locator lets callers compare a header,
// footer, note, or glossary paragraph directly against the source package.
func textObservationsPart(p *office.Package, part string) ([]map[string]any, error) {
	b := p.Files[part]
	if len(b) == 0 {
		return []map[string]any{}, nil
	}
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
		paragraphMarkFormatting := map[string]bool{}
		runs := []map[string]any{}
		references := []map[string]any{}
		nestedEnd := 0
		propertiesEnd := 0
		paragraphMarkEnd := 0
		for _, x := range spans[index+1:] {
			if x.Start >= s.End {
				break
			}
			if x.Start <= s.Start || x.End >= s.End {
				continue
			}
			if x.Start < nestedEnd || x.Name.Space != office.W {
				continue
			}
			if x.Name.Local == "p" {
				nestedEnd = x.End
				continue
			}
			if x.Name.Local == "pStyle" {
				style = x.Attribute(office.W, "val")
			}
			if x.Name.Local == "pPr" {
				props = string(b[x.Start:x.End])
				propertiesEnd = x.End
			}
			if x.Name.Local == "rPr" && x.Depth == s.Depth+2 && x.Start < propertiesEnd {
				paragraphMarkEnd = x.End
			}
			if paragraphMarkEnd > 0 && x.Start < paragraphMarkEnd && x.Depth == s.Depth+3 {
				switch x.Name.Local {
				case "b", "i", "caps", "smallCaps", "u":
					value := x.Attribute(office.W, "val")
					key := map[string]string{"b": "bold", "i": "italic", "caps": "caps", "smallCaps": "small_caps", "u": "underline"}[x.Name.Local]
					enabled := value == "" || (value != "0" && !strings.EqualFold(value, "false") && !strings.EqualFold(value, "off"))
					if x.Name.Local == "u" && strings.EqualFold(value, "none") {
						enabled = false
					}
					paragraphMarkFormatting[key] = enabled
				}
			}
			if x.Name.Local == "t" {
				value, err := textElement(b[x.Start:x.End])
				if err != nil {
					return nil, err
				}
				text.WriteString(value)
			}
			if x.Start < propertiesEnd {
				continue
			}
			switch x.Name.Local {
			case "tab":
				text.WriteByte('\t')
			case "br", "cr":
				text.WriteByte('\n')
			case "footnoteReference", "endnoteReference", "commentReference":
				references = append(references, map[string]any{"kind": x.Name.Local, "id": x.Attribute(office.W, "id"), "xml_start": x.Start})
			}
		}
		weighted := map[string]int{"text_units": 0, "bold_units": 0, "italic_units": 0, "caps_units": 0, "small_caps_units": 0, "underline_units": 0}
		nestedEnd = 0
		for runIndex, run := range spans[index+1:] {
			if run.Start >= s.End {
				break
			}
			if run.Start < nestedEnd || run.Name.Space != office.W {
				continue
			}
			if run.Name.Local == "p" {
				nestedEnd = run.End
				continue
			}
			if run.Name.Local != "r" || run.Start < propertiesEnd {
				continue
			}
			item := map[string]any{"xml_start": run.Start, "xml_end": run.End, "text_units": 0}
			var runText strings.Builder
			nestedRunEnd := 0
			for _, x := range spans[index+1+runIndex+1:] {
				if x.Start >= run.End {
					break
				}
				if x.Start < nestedRunEnd || x.Name.Space != office.W {
					continue
				}
				if x.Name.Local == "p" {
					nestedRunEnd = x.End
					continue
				}
				switch x.Name.Local {
				case "rPr":
					item["properties_xml"] = string(b[x.Start:x.End])
				case "t":
					value, err := textElement(b[x.Start:x.End])
					if err != nil {
						return nil, err
					}
					runText.WriteString(value)
				case "tab", "br", "cr":
					runText.WriteRune('\t')
				case "b", "i", "caps", "smallCaps", "u":
					value := x.Attribute(office.W, "val")
					enabled := value == "" || (value != "0" && !strings.EqualFold(value, "false") && !strings.EqualFold(value, "off"))
					if x.Name.Local == "u" && strings.EqualFold(value, "none") {
						enabled = false
					}
					item[x.Name.Local] = enabled
				case "rFonts":
					fonts := map[string]string{}
					for _, name := range []string{"ascii", "hAnsi", "cs", "eastAsia", "asciiTheme", "hAnsiTheme", "cstheme", "csTheme", "eastAsiaTheme"} {
						if value := x.Attribute(office.W, name); value != "" {
							fonts[name] = value
						}
					}
					item["fonts"] = fonts
					for _, name := range []string{"ascii", "hAnsi", "cs", "eastAsia"} {
						if value := fonts[name]; value != "" {
							item["font"] = value
							break
						}
					}
				case "rStyle":
					item["style_id"] = x.Attribute(office.W, "val")
				case "sz":
					item["size_half_points"] = x.Attribute(office.W, "val")
				}
			}
			units := len(utf16.Encode([]rune(runText.String())))
			item["text_units"] = units
			weighted["text_units"] += units
			for _, name := range []string{"bold", "italic", "caps", "small_caps", "underline"} {
				xmlName := map[string]string{"bold": "b", "italic": "i", "caps": "caps", "small_caps": "smallCaps", "underline": "u"}[name]
				if item[xmlName] == true {
					weighted[name+"_units"] += units
				}
			}
			runs = append(runs, item)
		}
		paragraphID := s.Attribute("http://schemas.microsoft.com/office/word/2010/wordml", "paraId")
		xmlPath := fmt.Sprintf("(//w:p)[%d]", len(out)+1)
		sourceID := part + "#" + xmlPath
		if paragraphID != "" {
			sourceID = part + "#paraId=" + paragraphID
		}
		observation := map[string]any{
			"text": text.String(), "style_id": style, "paragraph_properties_xml": props, "run_properties": runs,
			"direct_formatting_evidence": weighted,
			"source_part":                part, "xml_start": s.Start, "xml_end": s.End,
			"xml_path":     xmlPath,
			"paragraph_id": paragraphID,
			"source_id":    sourceID,
			"references":   references,
		}
		if len(paragraphMarkFormatting) > 0 {
			observation["paragraph_mark_formatting"] = paragraphMarkFormatting
		}
		out = append(out, observation)
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
	// Headers, footers, notes, comments, and glossary entries carry real
	// journal/template formatting too. Keep them in separate, part-qualified
	// observations so the main-story contract remains stable and callers can
	// inspect only the stories they need.
	parts := make([]string, 0)
	for part := range p.Files {
		if part != "word/document.xml" && isTextStoryPart(part) {
			parts = append(parts, part)
		}
	}
	sort.Strings(parts)
	stories := make([]map[string]any, 0, len(parts))
	for _, part := range parts {
		observations, err := textObservationsPart(p, part)
		if err != nil {
			return nil, err
		}
		if len(observations) > 0 {
			stories = append(stories, map[string]any{"part": part, "paragraph_observations": observations})
		}
	}
	if len(stories) > 0 {
		out["story_observations"] = stories
	}
	out["xml_namespaces"] = map[string]string{"w": office.W}
	out["locator_units"] = "xml_start/xml_end are zero-based UTF-8 byte offsets in source_part, end exclusive; xml_path uses xml_namespaces and includes nested paragraphs, not Word document paragraph indexes"
	out["source_sha256"] = office.Hash(b)
	out["effective_layout_verified"] = false
	return out, nil
}

func isTextStoryPart(part string) bool {
	if !strings.HasPrefix(part, "word/") || !strings.HasSuffix(part, ".xml") {
		return false
	}
	base := strings.TrimPrefix(part, "word/")
	if strings.HasPrefix(base, "header") || strings.HasPrefix(base, "footer") {
		return true
	}
	switch base {
	case "footnotes.xml", "endnotes.xml", "comments.xml", "glossary/document.xml":
		return true
	default:
		return false
	}
}
func Check(w *project.Workspace) (map[string]any, error) {
	return CheckWithConstants(w, nil)
}

func CheckWithConstants(w *project.Workspace, constants map[string]any) (map[string]any, error) {
	files, e := w.SourceFiles()
	if e != nil {
		return nil, e
	}
	symbols := []Symbol{}
	byName := map[string]bool{}
	diagnostics := []map[string]any{}
	parsedModules, skippedModules := 0, 0
	for n, b := range files {
		if strings.HasPrefix(n, "vba/") {
			{
				result, err := vbaparse.ParseWithConstants(string(b), constants)
				if err != nil {
					skippedModules++
					diagnostics = append(diagnostics, map[string]any{"severity": "error", "file": n, "message": err.Error()})
				} else {
					parsedModules++
					for _, d := range result["diagnostics"].([]vbaparse.Diagnostic) {
						diagnostics = append(diagnostics, map[string]any{"severity": "error", "file": n, "line": d.Line, "column": d.Column, "message": d.Message, "engine": "Rubberduck"})
					}
				}
			}
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
		if len(spans) == 0 || spans[0].Name.Local != "customUI" ||
			(spans[0].Name.Space != "http://schemas.microsoft.com/office/2009/07/customui" && spans[0].Name.Space != "http://schemas.microsoft.com/office/2006/01/customui") {
			continue
		}
		if len(spans) > 0 && spans[0].Name.Local == "customUI" {
			validation, err := native.ValidateRibbon(b)
			if err != nil {
				diagnostics = append(diagnostics, map[string]any{"severity": "error", "file": n, "message": err.Error()})
			} else if validation["valid"] != true {
				severity := "error"
				if validation["available"] == false {
					severity = "warning"
				}
				diagnostics = append(diagnostics, map[string]any{"severity": severity, "file": n, "message": "Ribbon schema validation did not pass", "validation": validation})
			}
		}
		ids := map[string]bool{}
		for _, s := range spans {
			if s.Name.Space != spans[0].Name.Space {
				continue
			}
			for _, a := range s.Attr {
				if a.Name.Space != "" {
					continue
				}
				if a.Name.Local == "id" {
					if ids[a.Value] {
						diagnostics = append(diagnostics, map[string]any{"severity": "error", "file": n, "message": "duplicate Ribbon id " + a.Value})
					}
					ids[a.Value] = true
				}
				if strings.HasPrefix(a.Name.Local, "get") || strings.HasPrefix(a.Name.Local, "on") || a.Name.Local == "loadImage" {
					cb := strings.ToLower(a.Value)
					if i := strings.LastIndex(cb, "."); i >= 0 {
						cb = cb[i+1:]
					}
					if !byName[cb] {
						diagnostic := map[string]any{"severity": "warning", "file": n, "message": "callback not lexically found: " + a.Value + "; dynamic/external routing needs native verification", "callback": a.Value, "control": s.Name.Local, "control_id": s.Attribute("", "id"), "attribute": a.Name.Local, "xml_start": s.Start}
						if declaration, known := office.RibbonCallbackDeclaration(s.Name.Local, a.Name.Local, a.Value); known {
							diagnostic["expected_declaration"] = declaration
						}
						diagnostics = append(diagnostics, diagnostic)
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
	return map[string]any{"symbols": symbols, "diagnostics": diagnostics, "compilation_constants": constants, "syntax_modules_parsed": parsedModules, "syntax_modules_skipped": skippedModules, "vba_compiled": false, "word_executed": false, "coverage": "Rubberduck VBA syntax with conditional preprocessing, lexical symbols, XML syntax, Windows RibbonX XSD validation, duplicate IDs and unresolved callback warnings; not a VBA compiler or complete callback/idMso checker"}, nil
}
