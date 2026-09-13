package inspect

import (
	"github.com/eliziff/WordUp/internal/office"
	"os"
	"path/filepath"
	"testing"
)

func TestStructureEvidenceInheritanceOverrideAndContainment(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/styles.xml"] = []byte(`<w:styles xmlns:w="` + office.W + `"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"/><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/></w:style><w:style w:type="paragraph" w:styleId="Cycle"><w:basedOn w:val="Cycle"/></w:style></w:styles>`)
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>Inherited</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>Another native heading</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Child"/><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>Explicit body</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>Table heading</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:pStyle w:val="Cycle"/></w:pPr></w:p></w:body></w:document>`)
	b, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "evidence.docx")
	if err = os.WriteFile(path, b, 0600); err != nil {
		t.Fatal(err)
	}
	result, err := StructureReference(path)
	if err != nil {
		t.Fatal(err)
	}
	rows := result["paragraphs"].([]map[string]any)
	if len(rows) != 5 {
		t.Fatal(rows)
	}
	inherited := rows[0]["outline_evidence"].(map[string]any)
	if inherited["level"] != 2 || inherited["origin"] != "style:Base" {
		t.Fatal(inherited)
	}
	direct := rows[2]["outline_evidence"].(map[string]any)
	if direct["body_text"] != true || direct["origin"] != "direct paragraph formatting" {
		t.Fatal(direct)
	}
	if _, exists := rows[3]["style_family_evidence"]; exists {
		t.Fatal("table contributed to body style families")
	}
	if rows[3]["context"] != "table" || rows[4]["style_error"] != "cyclic style inheritance at Cycle" {
		t.Fatal(rows)
	}
	family := rows[2]["style_family_evidence"].(map[string]any)
	if family["corroborated_level"] != 2 || family["requires_review"] != true {
		t.Fatal("lost override conflict", family)
	}
	if result["editorial_hierarchy_verified"] != false {
		t.Fatal("overclaimed hierarchy")
	}
}
