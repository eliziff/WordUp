package grade

import (
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap/ooxml"
)

const flatOPCHead = `<?xml version="1.0" standalone="yes"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">`

func part(name, xml string) string {
	return `<pkg:part pkg:name="` + name + `" pkg:contentType="application/xml"><pkg:xmlData>` + xml + `</pkg:xmlData></pkg:part>`
}

func styles() string {
	return `<w:styles xmlns:w="` + ooxml.W + `"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="character" w:default="1" w:styleId="Default"><w:name w:val="Default"/></w:style></w:styles>`
}

func para(text string) string {
	if text == "" {
		return "<w:p/>"
	}
	return `<w:p><w:r><w:t xml:space="preserve">` + text + `</w:t></w:r></w:p>`
}

func fieldPara(text string) string {
	return `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \o</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve">` + text + `</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
}

func note(id, text string) string {
	return `<w:footnote w:id="` + id + `"><w:p><w:r><w:t xml:space="preserve">` + text + `</w:t></w:r></w:p></w:footnote>`
}

func flatOPC(bodyParas []string, notes []string) []byte {
	body := `<w:document xmlns:w="` + ooxml.W + `"><w:body>` + strings.Join(bodyParas, "") + `</w:body></w:document>`
	fn := `<w:footnotes xmlns:w="` + ooxml.W + `">` + strings.Join(notes, "") + `</w:footnotes>`
	return []byte(flatOPCHead + part("/word/document.xml", body) + part("/word/styles.xml", styles()) + part("/word/footnotes.xml", fn) + `</pkg:package>`)
}

func TestGradeClassifiesSetupChanges(t *testing.T) {
	before := flatOPC([]string{
		para("Part I – The Commercial Lease"),
		para(""),
		para(`He said "law", but the priority between "rules" remains.`),
		para("Untouched paragraph."),
		fieldPara("Introduction 2"),
		para("A sentence that will be rewritten."),
	}, []string{
		note("1", "Ibid at pp. 400-408."),
		note("2", " Smith v Jones, [1999] 2 SCR 1 at 12."),
		note("3", "Some note that changes wording."),
	})
	after := flatOPC([]string{
		para("The Commercial Lease"),
		para("He said “law,” but the priority between “rules” remains."),
		para("Untouched paragraph."),
		fieldPara("Introduction 3"),
		para("A sentence that was rewritten."),
	}, []string{
		note("1", "\tIbid at 400–08."),
		note("2", "\tSmith v Jones, [1999] 2 SCR 1 at 12."),
		note("3", "\tSome note that changed wording."),
	})
	r, err := Grade(before, after, Options{})
	if err != nil {
		t.Fatal(err)
	}
	if r.Paragraphs != 6 || r.Notes != 3 {
		t.Fatalf("counts: %+v", r)
	}
	// note 2 loses a leading space for a tab: a formatting-only change.
	want := map[string]int{classCase: 0, classFormat: 3, classEmpty: 1, classProtected: 1, classCitation: 1}
	for cls, n := range want {
		if n > 0 && r.Expected[cls] != n {
			t.Errorf("%s: got %d want %d (%+v)", cls, r.Expected[cls], n, r.Expected)
		}
	}
	if r.Unexpected != 2 {
		t.Errorf("unexpected: got %d want 2: %+v", r.Unexpected, r.Examples)
	}
	if len(r.Examples) == 0 || !strings.HasPrefix(r.Examples[0].Class, "unexpected") {
		t.Fatalf("unexpected changes must lead the examples: %+v", r.Examples)
	}
	seen := map[string]bool{}
	for _, ex := range r.Examples {
		if strings.HasPrefix(ex.Class, "unexpected") {
			seen[ex.Container] = true
			if ex.Before == "" || ex.After == "" {
				t.Errorf("unexpected change lacks an excerpt: %+v", ex)
			}
		}
	}
	if !seen["body"] || !seen["footnote:3"] {
		t.Errorf("unexpected changes should name the body paragraph and note 3: %+v", r.Examples)
	}
	if r.Deleted != 1 {
		t.Errorf("deleted paragraphs: %d", r.Deleted)
	}
}

func TestGradeRejectsMissingParts(t *testing.T) {
	_, err := Grade([]byte(flatOPCHead+`</pkg:package>`), []byte(flatOPCHead+`</pkg:package>`), Options{})
	if err == nil {
		t.Fatal("expected an error for a package without styles")
	}
}

func TestCanonicalIsIdempotentAndOrderInsensitiveForQuotes(t *testing.T) {
	a := canonical(`Then "law", he said "yes."`)
	b := canonical("Then “law,” he said “yes”.")
	if a != b {
		t.Fatalf("canonical forms differ: %q vs %q", a, b)
	}
	if canonical(canonical(a)) != a {
		t.Fatalf("canonical is not idempotent")
	}
	if canonical("PART III – Overview") != "Overview" || canonical("III. Overview") != "Overview" {
		t.Fatalf("heading prefixes not removed: %q %q", canonical("PART III – Overview"), canonical("III. Overview"))
	}
}
