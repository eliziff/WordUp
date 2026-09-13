package office

import (
	"bytes"
	"context"
	"fmt"
	"math"
	"unicode/utf8"

	"github.com/antchfx/xmlquery"
	"github.com/antchfx/xpath"
)

// QueryXMLInput accepts raw XML or an explicitly selected OPC part. Both the
// container and part hashes are retained; part selection is never inferred.
func QueryXMLInput(ctx context.Context, data []byte, part, query string, namespaces map[string]string, limit int) (map[string]any, error) {
	if part == "" {
		return QueryXML(ctx, data, query, namespaces, limit)
	}
	if !SafePart(part) {
		return nil, fmt.Errorf("unsafe XML package part %q", part)
	}
	pkg, err := ReadPackage(data)
	if err != nil {
		return nil, err
	}
	xml, ok := pkg.Files[part]
	if !ok {
		return nil, fmt.Errorf("XML package part %q not found", part)
	}
	result, err := QueryXML(ctx, xml, query, namespaces, limit)
	if err == nil {
		result["package_sha256"], result["part"] = Hash(data), part
	}
	return result, err
}

// QueryXML is read-only: the XPath tree is only an index into original UTF-8
// bytes. It is never serialized back into an editable package.
func QueryXML(ctx context.Context, data []byte, query string, namespaces map[string]string, limit int) (map[string]any, error) {
	if len(data) > 16<<20 || len(query) == 0 || len(query) > 4096 || len(namespaces) > 64 {
		return nil, fmt.Errorf("xml.query requires XML <=16 MiB, XPath 1..4096 bytes and <=64 namespaces")
	}
	if limit == 0 {
		limit = 20
	}
	if limit < 1 || limit > 100 {
		return nil, fmt.Errorf("xml.query limit must be 1..100")
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	// Reuse strict validation, DTD prohibition, depth bounds and exact offsets.
	spans, err := XMLSpans(data)
	if err != nil {
		return nil, err
	}
	if len(spans) > 200000 {
		return nil, fmt.Errorf("xml.query element budget exceeded (200000)")
	}
	expr, err := xpath.CompileWithNS(query, namespaces)
	if err != nil {
		return nil, fmt.Errorf("XPath: %w", err)
	}
	doc, err := xmlquery.Parse(bytes.NewReader(data))
	if err != nil {
		return nil, err
	}
	locations := make(map[*xmlquery.Node]XMLSpan, len(spans))
	index := 0
	var locate func(*xmlquery.Node) error
	locate = func(n *xmlquery.Node) error {
		if n.Type == xmlquery.ElementNode {
			if index >= len(spans) || spans[index].Name.Local != n.Data || spans[index].Name.Space != n.NamespaceURI {
				return fmt.Errorf("XPath parser and exact XML source disagree")
			}
			locations[n] = spans[index]
			index++
		}
		for child := n.FirstChild; child != nil; child = child.NextSibling {
			if err := locate(child); err != nil {
				return err
			}
		}
		return nil
	}
	if err := locate(doc); err != nil {
		return nil, err
	}
	if index != len(spans) {
		return nil, fmt.Errorf("XPath source location count mismatch")
	}
	result := map[string]any{"schema": 1, "source_sha256": Hash(data), "source_bytes": len(data), "query": query, "namespaces": namespaces}
	value := expr.Evaluate(xmlquery.CreateXPathNavigator(doc))
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	iter, nodes := value.(*xpath.NodeIterator)
	if !nodes {
		if number, ok := value.(float64); ok && (math.IsNaN(number) || math.IsInf(number, 0)) {
			return nil, fmt.Errorf("XPath returned a non-finite number")
		}
		if text, ok := value.(string); ok {
			result["value"] = queryPreview(text)
		} else {
			result["value"] = value
		}
		result["kind"] = "scalar"
		return result, nil
	}
	matches := []any{}
	truncated := false
	for iter.MoveNext() {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		if len(matches) == limit {
			truncated = true
			break
		}
		nav := iter.Current().(*xmlquery.NodeNavigator)
		n := nav.Current()
		match := map[string]any{"name": nav.LocalName(), "namespace": nav.NamespaceURL(), "node_type": nav.NodeType()}
		if nav.NodeType() != xpath.ElementNode {
			match["value"] = queryPreview(nav.Value())
		}
		for n != nil && n.Type != xmlquery.ElementNode {
			n = n.Parent
		}
		if span, ok := locations[n]; ok {
			match["element_start"] = span.Start
			match["element_end"] = span.End
			match["element_xml"] = queryPreview(string(data[span.Start:span.End]))
		}
		matches = append(matches, match)
	}
	result["kind"], result["matches"], result["returned"], result["truncated"] = "nodes", matches, len(matches), truncated
	return result, nil
}

func queryPreview(text string) any {
	const limit = 1024
	if len(text) <= limit {
		return text
	}
	end := limit
	for !utf8.ValidString(text[:end]) {
		end--
	}
	return map[string]any{"prefix": text[:end], "bytes": len(text), "sha256": Hash([]byte(text)), "truncated": true}
}
