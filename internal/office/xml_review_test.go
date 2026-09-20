package office

import (
	"context"
	"strings"
	"testing"
)

func reviewDoc(body string) []byte {
	return []byte(`<w:document xmlns:w="` + W + `"><w:body>` + body + `</w:body></w:document>`)
}

func TestXMLReviewShowsNativeEvidenceWithoutAcceptingRevisions(t *testing.T) {
	a := reviewDoc(`<w:p><w:r><w:t>old</w:t></w:r></w:p><w:p><w:r><w:t>same</w:t></w:r></w:p>`)
	b := reviewDoc(`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:del><w:r><w:delText>old</w:delText></w:r></w:del><w:ins><w:r><w:rPr><w:i/></w:rPr><w:t>new &amp; é</w:t></w:r></w:ins><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> REF bookmark </w:instrText><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p><w:r><w:t>same</w:t></w:r></w:p>`)
	r, err := XMLReview(context.Background(), a, b, 0, 1)
	if err != nil {
		t.Fatal(err)
	}
	if r["changed_paragraphs"] != 1 || r["output_reviewed"] != false || r["byte_identical"] != false || r["candidate_sha256"] != Hash(b) {
		t.Fatal(r)
	}
	row := r["rows"].([]any)[0].(map[string]any)
	after := row["after"].(*reviewParagraph)
	if after.Text != "new & é" {
		t.Fatal(after.Text)
	}
	kinds := map[string]bool{}
	for _, e := range after.Evidence {
		kinds[e.Kind] = true
		if e.XML != string(b[e.Start:e.End]) {
			t.Fatal("evidence lost exact byte location")
		}
	}
	for _, kind := range []string{"del", "ins", "rPr", "pPr", "fldChar", "instrText"} {
		if !kinds[kind] {
			t.Fatal("missing evidence", kind)
		}
	}
}

func TestXMLReviewScopesNotesAndBoundsPreviews(t *testing.T) {
	a := []byte(`<w:footnotes xmlns:w="` + W + `"><w:footnote w:id="2"><w:p><w:r><w:t>A</w:t></w:r></w:p></w:footnote><w:footnote w:id="7"><w:p/></w:footnote></w:footnotes>`)
	b := []byte(strings.ReplaceAll(string(a), "<w:t>A</w:t>", "<w:t>"+strings.Repeat("中", 3000)+"</w:t>"))
	r, err := XMLReview(context.Background(), a, b, 0, 1)
	if err != nil {
		t.Fatal(err)
	}
	row := r["rows"].([]any)[0].(map[string]any)
	if row["location"] != "xml/footnote:2/p:000001" || !row["after"].(*reviewParagraph).TextTruncated {
		t.Fatal(row)
	}
	r, err = XMLReview(context.Background(), a, b, 1, 1)
	if err != nil || r["returned"] != 0 || r["more"] != false {
		t.Fatal(r, err)
	}
	for _, input := range []string{`<!DOCTYPE r><r/>`, `<r>`, `<r/><r/>`} {
		if _, err := XMLReview(context.Background(), []byte(input), b, 0, 1); err == nil {
			t.Fatal("accepted invalid XML")
		}
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if _, err := XMLReview(ctx, a, b, 0, 1); err == nil {
		t.Fatal("ignored cancellation")
	}
}

func TestXMLReviewParagraphLevelRevision(t *testing.T) {
	p := `<w:p><w:r><w:t>removed</w:t></w:r></w:p>`
	r, err := XMLReview(context.Background(), reviewDoc(p), reviewDoc(`<w:del>`+p+`</w:del>`), 0, 1)
	if err != nil || r["changed_paragraphs"] != 1 {
		t.Fatal(r, err)
	}
	after := r["rows"].([]any)[0].(map[string]any)["after"].(*reviewParagraph)
	if after.Text != "" || after.EnclosingRevision != "del/" {
		t.Fatal(after)
	}
}

func TestXMLReviewReportsFormattingOnlyAndRemovedParagraphs(t *testing.T) {
	a := reviewDoc(`<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>same</w:t></w:r></w:p><w:p/>`)
	b := reviewDoc(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>same</w:t></w:r></w:p>`)
	r, err := XMLReview(context.Background(), a, b, 0, 1)
	if err != nil || r["changed_paragraphs"] != 2 || r["more"] != true {
		t.Fatal(r, err)
	}
	r, err = XMLReview(context.Background(), a, b, 1, 1)
	if err != nil {
		t.Fatal(err)
	}
	if r["rows"].([]any)[0].(map[string]any)["after"].(*reviewParagraph) != nil {
		t.Fatal("removed paragraph not represented")
	}
}
