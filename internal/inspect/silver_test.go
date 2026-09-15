package inspect

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/office"
)

func TestCompareStructureSilverChecksOptionalAmbiguity(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="12345678"><w:r><w:t>I. Ambiguous</w:t></w:r></w:p></w:body></w:document>`)
	b, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	root := t.TempDir()
	document := filepath.Join(root, "fixture.docx")
	if err := os.WriteFile(document, b, 0600); err != nil {
		t.Fatal(err)
	}
	resolved, err := StructureResolved(document)
	if err != nil {
		t.Fatal(err)
	}
	rows := resolved["paragraphs"].([]map[string]any)
	if len(rows) != 1 {
		t.Fatalf("paragraphs=%d", len(rows))
	}
	structure := rows[0]["resolved_structure"].(map[string]any)
	if structure["ambiguous"] != true {
		t.Fatalf("fixture did not produce ambiguity: %#v", structure)
	}
	id := rows[0]["source_id"].(string)
	if !strings.HasPrefix(id, "word/document.xml#paraId=") {
		t.Fatalf("unexpected source id: %s", id)
	}
	paragraphID := strings.TrimPrefix(id, "word/document.xml#paraId=")
	silver := filepath.Join(root, "silver.xml")
	content := fmt.Sprintf(`<?xml version="1.0"?><structure-silver schema="1"><documents><document source="fixture.docx" source_sha256="%s"><p paraId="%s" role="%s" level="%d" parent="none" ambiguity="roman-marker"/></document></documents></structure-silver>`,
		resolved["source_sha256"], paragraphID, structure["role"], structure["level"])
	if err := os.WriteFile(silver, []byte(content), 0600); err != nil {
		t.Fatal(err)
	}
	report, err := CompareStructureSilver(root, silver, 10)
	if err != nil {
		t.Fatal(err)
	}
	if report["mismatch_count"] != 0 {
		t.Fatalf("optional ambiguity did not match: %#v", report)
	}
	content = strings.Replace(content, `ambiguity="roman-marker"`, `ambiguity="none"`, 1)
	if err := os.WriteFile(silver, []byte(content), 0600); err != nil {
		t.Fatal(err)
	}
	report, err = CompareStructureSilver(root, silver, 10)
	if err != nil {
		t.Fatal(err)
	}
	if report["mismatch_count"] != 1 || report["mismatch_fields"].(map[string]int)["ambiguous"] != 1 {
		t.Fatalf("ambiguity mismatch was not reported: %#v", report)
	}
	examples := report["uncertainty_examples"].(map[string]any)
	if examples["ambiguous"].(map[string]any)["source_id"] != id {
		t.Fatalf("ambiguity evidence example lost its source: %#v", examples)
	}
}

func TestCompareStructureSilverUsesXMLPathForDuplicateParagraphIDs(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="ABCDEF01"><w:r><w:t>First</w:t></w:r></w:p><w:p w14:paraId="ABCDEF01"><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>`)
	b, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	root := t.TempDir()
	document := filepath.Join(root, "duplicate.docx")
	if err := os.WriteFile(document, b, 0600); err != nil {
		t.Fatal(err)
	}
	resolved, err := StructureResolved(document)
	if err != nil {
		t.Fatal(err)
	}
	rows := resolved["paragraphs"].([]map[string]any)
	if len(rows) != 2 {
		t.Fatalf("paragraphs=%d", len(rows))
	}
	for _, row := range rows {
		if row["source_id"].(string) == "word/document.xml#paraId=ABCDEF01" || row["paragraph_id_ambiguous"] != true {
			t.Fatalf("duplicate paragraph identity was not downgraded: %#v", row)
		}
	}
	first := rows[0]["resolved_structure"].(map[string]any)
	second := rows[1]["resolved_structure"].(map[string]any)
	silver := filepath.Join(root, "silver.xml")
	content := fmt.Sprintf(`<?xml version="1.0"?><structure-silver schema="1"><documents><document source="duplicate.docx" source_sha256="%s"><p xml_path="(//w:p)[1]" role="%s" level="%d" parent="none"/><p xml_path="(//w:p)[2]" role="%s" level="%d" parent="none"/></document></documents></structure-silver>`,
		resolved["source_sha256"], first["role"], first["level"], second["role"], second["level"])
	if err := os.WriteFile(silver, []byte(content), 0600); err != nil {
		t.Fatal(err)
	}
	report, err := CompareStructureSilver(root, silver, 10)
	if err != nil {
		t.Fatal(err)
	}
	if report["mismatch_count"] != 0 {
		t.Fatalf("xml-path silver did not disambiguate duplicate IDs: %#v", report)
	}
}

func TestCompareStructureSilverAcceptsXMLPathWithUniqueParagraphID(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="ABCDEF01"><w:r><w:t>One</w:t></w:r></w:p></w:body></w:document>`)
	b, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	root := t.TempDir()
	document := filepath.Join(root, "unique.docx")
	if err := os.WriteFile(document, b, 0600); err != nil {
		t.Fatal(err)
	}
	resolved, err := StructureResolved(document)
	if err != nil {
		t.Fatal(err)
	}
	rows := resolved["paragraphs"].([]map[string]any)
	if len(rows) != 1 {
		t.Fatalf("paragraphs=%d", len(rows))
	}
	structure := rows[0]["resolved_structure"].(map[string]any)
	silver := filepath.Join(root, "silver.xml")
	content := fmt.Sprintf(`<?xml version="1.0"?><structure-silver schema="1"><documents><document source="unique.docx" source_sha256="%s"><p paraId="ABCDEF01" xml_path="(//w:p)[1]" role="%s" level="%d" parent="none"/></document></documents></structure-silver>`,
		resolved["source_sha256"], structure["role"], structure["level"])
	if err := os.WriteFile(silver, []byte(content), 0600); err != nil {
		t.Fatal(err)
	}
	report, err := CompareStructureSilver(root, silver, 10)
	if err != nil {
		t.Fatal(err)
	}
	if report["mismatch_count"] != 0 {
		t.Fatalf("xml path with unique ID did not match: %#v", report)
	}
	content = strings.Replace(content, `paraId="ABCDEF01"`, `paraId="ABCDEF02"`, 1)
	if err := os.WriteFile(silver, []byte(content), 0600); err != nil {
		t.Fatal(err)
	}
	report, err = CompareStructureSilver(root, silver, 10)
	if err != nil {
		t.Fatal(err)
	}
	if report["mismatch_count"] != 1 || report["mismatch_fields"].(map[string]int)["identity"] != 1 {
		t.Fatalf("xml path did not validate accompanying paragraph ID: %#v", report)
	}
}
