package office

import (
	"bytes"
	"context"
	"encoding/xml"
	"fmt"
	"sort"
	"strings"
	"unicode/utf8"
)

// XMLReview is an inspection aid, not an editorial verdict or equivalence test.
// Paragraph ordinals deliberately remain visible: insertions can shift alignment.
func XMLReview(ctx context.Context, before, after []byte, offset, limit int) (map[string]any, error) {
	if offset < 0 || limit < 0 || limit > 100 {
		return nil, fmt.Errorf("xml.review requires offset >=0 and limit 0..100")
	}
	if limit == 0 {
		limit = 20
	}
	a, err := reviewParagraphs(ctx, before)
	if err != nil {
		return nil, err
	}
	b, err := reviewParagraphs(ctx, after)
	if err != nil {
		return nil, err
	}
	keys := map[string]bool{}
	for k := range a {
		keys[k] = true
	}
	for k := range b {
		keys[k] = true
	}
	ordered := make([]string, 0, len(keys))
	for k := range keys {
		ordered = append(ordered, k)
	}
	sort.Strings(ordered)
	rows := []any{}
	total := 0
	for _, key := range ordered {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		x, y := a[key], b[key]
		if x != nil && y != nil && x.SHA256 == y.SHA256 && x.EnclosingRevision == y.EnclosingRevision {
			continue
		}
		if total >= offset && len(rows) < limit {
			rows = append(rows, map[string]any{"location": key, "before": x, "after": y})
		}
		total++
	}
	return map[string]any{
		"schema": 1, "reference_sha256": Hash(before), "candidate_sha256": Hash(after),
		"byte_identical": bytes.Equal(before, after), "changed_paragraphs": total,
		"reference_paragraphs": len(a), "candidate_paragraphs": len(b),
		"offset": offset, "returned": len(rows), "more": offset+len(rows) < total, "rows": rows,
		"output_reviewed": false,
		"coverage":        "Transitional WordprocessingML paragraphs, displayed text with pending deletions excluded, explicit insertion/deletion/move spans, field markup, and serialized direct run/paragraph formatting. Scope and paragraph ordinal align rows, not semantic paragraph identity. Previews are bounded; full source hashes and byte offsets are retained. Does not resolve inherited formatting, prove serial behavior, inspect every package feature, or approve editorial output.",
	}, nil
}

type reviewParagraph struct {
	EnclosingRevision string          `json:"enclosing_revision,omitempty"`
	SHA256            string          `json:"xml_sha256"`
	Start             int             `json:"byte_start"`
	End               int             `json:"byte_end"`
	Text              string          `json:"text"`
	TextTruncated     bool            `json:"text_truncated"`
	Evidence          []reviewElement `json:"evidence"`
	EvidenceOmitted   int             `json:"evidence_omitted"`
}

type reviewElement struct {
	Kind      string `json:"kind"`
	XML       string `json:"xml"`
	Start     int    `json:"byte_start"`
	End       int    `json:"byte_end"`
	Truncated bool   `json:"truncated"`
}

func reviewParagraphs(ctx context.Context, raw []byte) (map[string]*reviewParagraph, error) {
	spans, err := XMLSpans(raw)
	if err != nil {
		return nil, err
	}
	if len(spans) > 500000 {
		return nil, fmt.Errorf("xml.review element budget exceeded")
	}
	out := map[string]*reviewParagraph{}
	counts := map[string]int{}
	parents := []int{}
	for i, p := range spans {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		for len(parents) > 0 && spans[parents[len(parents)-1]].End <= p.Start {
			parents = parents[:len(parents)-1]
		}
		if p.Name == (xml.Name{Space: wordXMLNamespace, Local: "p"}) {
			scope := "xml"
			enclosingRevision := ""
			for _, parent := range parents {
				s := spans[parent]
				if s.Name.Space == wordXMLNamespace {
					switch s.Name.Local {
					case "ins", "del", "moveFrom", "moveTo":
						enclosingRevision += s.Name.Local + "/"
					}
				}
				if s.Name.Space == "http://schemas.microsoft.com/office/2006/xmlPackage" && s.Name.Local == "part" {
					scope = s.Attribute("*", "name")
				}
				if s.Name.Space == wordXMLNamespace {
					switch s.Name.Local {
					case "body", "hdr", "ftr", "txbxContent":
						scope += "/" + s.Name.Local
					case "footnote", "endnote", "comment":
						scope += "/" + s.Name.Local + ":" + s.Attribute(wordXMLNamespace, "id")
					}
				}
			}
			counts[scope]++
			r := &reviewParagraph{EnclosingRevision: enclosingRevision, SHA256: Hash(raw[p.Start:p.End]), Start: p.Start, End: p.End, Evidence: []reviewElement{}}
			var text strings.Builder
			deletedUntil, nestedUntil := 0, 0
			for _, parent := range parents {
				s := spans[parent]
				if s.Name.Space == wordXMLNamespace && (s.Name.Local == "del" || s.Name.Local == "moveFrom") {
					deletedUntil = s.End
				}
			}
			for j := i + 1; j < len(spans) && spans[j].Start < p.End; j++ {
				s := spans[j]
				if s.Start < nestedUntil || s.Name.Space != wordXMLNamespace {
					continue
				}
				if s.Name.Local == "p" {
					nestedUntil = s.End
					continue
				}
				switch s.Name.Local {
				case "del", "moveFrom":
					if s.End > deletedUntil {
						deletedUntil = s.End
					}
				}
				switch s.Name.Local {
				case "ins", "del", "moveFrom", "moveTo", "rPr", "pPr", "instrText", "delInstrText", "fldChar", "fldSimple", "sdtPr":
					if len(r.Evidence) < 100 {
						v := string(raw[s.Start:s.End])
						preview := reviewPreview(v)
						r.Evidence = append(r.Evidence, reviewElement{s.Name.Local, preview, s.Start, s.End, preview != v})
					} else {
						r.EvidenceOmitted++
					}
				}
				if s.Start < deletedUntil {
					continue
				}
				switch s.Name.Local {
				case "t":
					var value struct {
						Text string `xml:",chardata"`
					}
					if err := xml.Unmarshal(raw[s.Start:s.End], &value); err != nil {
						return nil, err
					}
					text.WriteString(value.Text)
				case "tab":
					text.WriteByte('\t')
				case "br", "cr":
					text.WriteByte('\n')
				case "noBreakHyphen":
					text.WriteRune('\u2011')
				case "softHyphen":
					text.WriteRune('\u00ad')
				case "footnoteReference", "endnoteReference":
					text.WriteString("[note]")
				}
			}
			r.Text = reviewPreview(text.String())
			r.TextTruncated = r.Text != text.String()
			out[fmt.Sprintf("%s/p:%06d", scope, counts[scope])] = r
		}
		parents = append(parents, i)
	}
	return out, nil
}

func reviewPreview(text string) string {
	const limit = 1600
	if len(text) <= limit {
		return text
	}
	end := limit
	for !utf8.RuneStart(text[end]) {
		end--
	}
	return text[:end] + "…"
}
