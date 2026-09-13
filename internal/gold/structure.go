// Package gold validates independently authored expectations. Validation and
// annotator agreement are evidence hygiene, never automatic editorial approval.
package gold

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strings"

	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

type Alternative struct {
	Role   string `json:"role"`
	Level  int    `json:"level"`
	Parent int    `json:"parent"`
	Reason string `json:"reason"`
}
type Label struct {
	Paragraph    int           `json:"paragraph"`
	Text         string        `json:"text"`
	Role         string        `json:"role"`
	Level        int           `json:"level"`
	Parent       int           `json:"parent"`
	Alternatives []Alternative `json:"alternatives"`
	Rationale    string        `json:"rationale"`
}
type Document struct {
	Source   string `json:"source"`
	SHA256   string `json:"source_sha256"`
	Part     string `json:"part"`
	Coverage struct {
		First int `json:"first"`
		Last  int `json:"last"`
	} `json:"coverage"`
	Labels []Label `json:"labels"`
}
type Annotation struct {
	Schema    int        `json:"schema"`
	Annotator string     `json:"annotator"`
	Status    string     `json:"status"`
	Documents []Document `json:"documents"`
}

func read(root, path string, out any) ([]byte, error) {
	data, err := project.Read(root, path)
	if err != nil {
		return nil, err
	}
	if err := project.ReadJSON(data, out); err != nil {
		return nil, err
	}
	return data, nil
}

func validRole(role string, level int) bool {
	if role == "heading" {
		return level >= 1 && level <= 9
	}
	switch role {
	case "body", "title", "author", "metadata", "abstract", "toc", "table", "quotation", "blank", "uncertain":
		return level == 0
	default:
		return false
	}
}

func ValidateStructure(root, path string) (*Annotation, map[string]any, error) {
	var a Annotation
	data, err := read(root, path, &a)
	if err != nil {
		return nil, nil, err
	}
	if a.Schema != 1 || a.Annotator == "" || (a.Status != "proposed" && a.Status != "silver") || len(a.Documents) == 0 {
		return nil, nil, fmt.Errorf("expected schema 1 proposed or silver annotation with annotator and documents; validation does not certify labels")
	}
	seen := map[string]bool{}
	total, uncertain := 0, 0
	for _, d := range a.Documents {
		key := d.Source + "#" + d.Part
		if seen[key] || !office.SafePart(d.Part) {
			return nil, nil, fmt.Errorf("duplicate document or unsafe part: %s", key)
		}
		seen[key] = true
		source, err := project.Read(root, d.Source)
		if err != nil {
			return nil, nil, err
		}
		if office.Hash(source) != d.SHA256 {
			return nil, nil, fmt.Errorf("source hash mismatch: %s", d.Source)
		}
		pkg, err := office.ReadPackage(source)
		if err != nil {
			return nil, nil, err
		}
		text, err := paragraphTexts(pkg.Files[d.Part])
		if err != nil {
			return nil, nil, err
		}
		if d.Coverage.First < 1 || d.Coverage.Last > len(text) || d.Coverage.Last < d.Coverage.First || len(d.Labels) != d.Coverage.Last-d.Coverage.First+1 {
			return nil, nil, fmt.Errorf("incomplete/out-of-range coverage: %s", d.Source)
		}
		for i, l := range d.Labels {
			if l.Paragraph != d.Coverage.First+i || l.Text != text[l.Paragraph-1] {
				return nil, nil, fmt.Errorf("paragraph location/text mismatch: %s paragraph %d", d.Source, l.Paragraph)
			}
			if !validRole(l.Role, l.Level) || strings.TrimSpace(l.Rationale) == "" {
				return nil, nil, fmt.Errorf("invalid role/level or missing rationale: paragraph %d", l.Paragraph)
			}
			if l.Parent < 0 || l.Parent >= l.Paragraph {
				return nil, nil, fmt.Errorf("invalid parent: paragraph %d", l.Paragraph)
			}
			if l.Parent > 0 {
				j := l.Parent - d.Coverage.First
				if j < 0 || j >= i || d.Labels[j].Role != "heading" || (l.Role == "heading" && d.Labels[j].Level >= l.Level) {
					return nil, nil, fmt.Errorf("parent must be a covered earlier heading of lower level: paragraph %d", l.Paragraph)
				}
			}
			for _, alt := range l.Alternatives {
				if !validRole(alt.Role, alt.Level) || alt.Parent < 0 || alt.Parent >= l.Paragraph || strings.TrimSpace(alt.Reason) == "" {
					return nil, nil, fmt.Errorf("invalid alternative: paragraph %d", l.Paragraph)
				}
			}
			if l.Role == "uncertain" || len(l.Alternatives) > 0 {
				uncertain++
			}
			total++
		}
	}
	return &a, map[string]any{"valid": true, "status": a.Status, "annotation_sha256": office.Hash(data), "documents": len(a.Documents), "paragraphs": total, "uncertain": uncertain, "editorially_approved": false}, nil
}

func paragraphTexts(data []byte) ([]string, error) {
	if _, err := office.XMLSpans(data); err != nil {
		return nil, err
	}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	var texts []*strings.Builder
	var stack []int
	inText := false
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch v := token.(type) {
		case xml.StartElement:
			if v.Name.Space == office.W && v.Name.Local == "p" {
				texts = append(texts, &strings.Builder{})
				stack = append(stack, len(texts)-1)
			}
			if v.Name.Space == office.W && v.Name.Local == "t" {
				inText = true
			}
		case xml.EndElement:
			if v.Name.Space == office.W && v.Name.Local == "t" {
				inText = false
			}
			if v.Name.Space == office.W && v.Name.Local == "p" {
				stack = stack[:len(stack)-1]
			}
		case xml.CharData:
			if inText {
				for _, i := range stack {
					texts[i].Write(v)
				}
			}
		}
	}
	out := make([]string, len(texts))
	for i := range texts {
		out[i] = texts[i].String()
	}
	return out, nil
}

func CompareStructure(root, left, right string) (map[string]any, error) {
	a, ar, err := ValidateStructure(root, left)
	if err != nil {
		return nil, err
	}
	b, br, err := ValidateStructure(root, right)
	if err != nil {
		return nil, err
	}
	if a.Annotator == b.Annotator {
		return nil, fmt.Errorf("agreement requires distinct annotator identities (independence must also be reviewed)")
	}
	if len(a.Documents) != len(b.Documents) {
		return nil, fmt.Errorf("annotation scopes differ")
	}
	bySource := map[string]Document{}
	for _, d := range b.Documents {
		bySource[d.Source+"#"+d.Part] = d
	}
	total, agree, disputes := 0, 0, 0
	differences := []any{}
	for _, d := range a.Documents {
		other, ok := bySource[d.Source+"#"+d.Part]
		if !ok || other.SHA256 != d.SHA256 || other.Coverage != d.Coverage {
			return nil, fmt.Errorf("annotation scopes or source hashes differ")
		}
		for i, l := range d.Labels {
			r := other.Labels[i]
			total++
			if l.Role == r.Role && l.Level == r.Level && l.Parent == r.Parent {
				agree++
			} else {
				disputes++
				if len(differences) < 100 {
					differences = append(differences, map[string]any{"source": d.Source, "part": d.Part, "paragraph": l.Paragraph, "left": Alternative{l.Role, l.Level, l.Parent, l.Rationale}, "right": Alternative{r.Role, r.Level, r.Parent, r.Rationale}})
				}
			}
		}
	}
	return map[string]any{"schema": 1, "left": ar, "right": br, "paragraphs": total, "primary_label_agreements": agree, "primary_label_disagreements": disputes, "agreement_fraction": float64(agree) / float64(total), "differences": differences, "truncated": disputes > len(differences), "editorially_approved": false, "note": "Primary role/level/parent agreement only; alternatives and uncertainty still require adjudication."}, nil
}
