// Package alrtest holds the synthetic adversarial fixtures shared by the
// alrmap test suites (the reference test_support module). No private
// manuscript text is embedded here.
package alrtest

import (
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap"
	"github.com/eliziff/WordUp/internal/alrmap/citations"
	"github.com/eliziff/WordUp/internal/alrmap/kernels"
	"github.com/eliziff/WordUp/internal/alrmap/ooxml"
)

// Escape mirrors xml.sax.saxutils.escape (&, <, > only).
func Escape(s string) string {
	return strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;").Replace(s)
}

// Styles builds a minimal styles part with optional extra style definitions
// and document defaults.
func Styles(t testing.TB, extra, defaults string) *ooxml.Styles {
	t.Helper()
	st, err := ooxml.NewStyles([]byte(`<w:styles xmlns:w="` + ooxml.W + `">` + defaults +
		`<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
		`<w:style w:type="character" w:default="1" w:styleId="Default"><w:name w:val="Default"/></w:style>` +
		extra + `</w:styles>`))
	if err != nil {
		t.Fatalf("styles fixture: %v", err)
	}
	return st
}

// Run renders one run with optional run properties.
func Run(text, props string, preserve bool) string {
	space := ""
	if preserve {
		space = ` xml:space="preserve"`
	}
	rpr := ""
	if props != "" {
		rpr = "<w:rPr>" + props + "</w:rPr>"
	}
	return "<w:r>" + rpr + "<w:t" + space + ">" + Escape(text) + "</w:t></w:r>"
}

// Para renders one paragraph with optional paragraph properties.
func Para(content, props string) string {
	ppr := ""
	if props != "" {
		ppr = "<w:pPr>" + props + "</w:pPr>"
	}
	return "<w:p>" + ppr + content + "</w:p>"
}

// Document wraps body content in a document part with the test namespaces.
func Document(content string) []byte {
	return []byte(`<w:document xmlns:w="` + ooxml.W + `" xmlns:m="urn:math" xmlns:x="urn:extra"><w:body>` + content + `</w:body></w:document>`)
}

// Project scans body content as word/document.xml (default styles when st is nil).
func Project(t testing.TB, content string, st *ooxml.Styles) *ooxml.Scan {
	t.Helper()
	if st == nil {
		st = Styles(t, "", "")
	}
	sc, err := ooxml.ScanPart(Document(content), st, "word/document.xml")
	if err != nil {
		t.Fatalf("project fixture: %v", err)
	}
	return sc
}

// Lexical plans authored-text edits over the first container and applies
// them through the XML-copy writer.
func Lexical(t testing.TB, sc *ooxml.Scan) ([]byte, []ooxml.ByteWrite, error) {
	t.Helper()
	c := sc.Containers[0]
	extra := append(append([]kernels.Span(nil), c.Hard...), c.Lexical...)
	p, err := alrmap.PlanAuthorText(c.Text, alrmap.AuthorTextOptions{Extra: extra, Authority: "synthetic authored scope"})
	if err != nil {
		return nil, nil, err
	}
	bound, err := ooxml.Bind(sc, c.Key, p.Edits, "lexical", "synthetic authored scope")
	if err != nil {
		return nil, nil, err
	}
	return ooxml.ApplyBound(sc, bound)
}

// Reg mirrors the shared REG registry of the reference suite.
var Reg = &citations.Registry{
	Journals: map[string][]string{"Example Law Review": {"Ex L Rev"}, "Ex L Rev": {"Ex L Rev"}},
	Cases:    map[string][]string{"R. v. Example Ltd.": {"R v Example Ltd"}, "R v Example Ltd": {"R v Example Ltd"}},
	Aliases:  map[string]bool{"Example": true, "Author, Treatise": true},
	Books:    map[string]bool{"Towards Judgement": true},
}

// Text rescans an output copy and returns the text of container i.
func Text(t testing.TB, out []byte, st *ooxml.Styles, part string, i int) string {
	t.Helper()
	sc, err := ooxml.ScanPart(out, st, part)
	if err != nil {
		t.Fatalf("rescan: %v", err)
	}
	return sc.Containers[i].Text
}
