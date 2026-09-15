package inspect

import (
	"github.com/eliziff/WordUp/internal/office"
	"os"
	"path/filepath"
	"testing"
)

func TestStructureEvidenceInheritanceOverrideAndContainment(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/settings.xml"] = []byte(`<w:settings xmlns:w="` + office.W + `"><w:defaultTabStop w:val="720"/></w:settings>`)
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
	if result["styles_xml"] != string(p.Files["word/styles.xml"]) || result["styles_sha256"] != office.Hash(p.Files["word/styles.xml"]) || result["document_xml_sha256"] != office.Hash(p.Files["word/document.xml"]) {
		t.Fatalf("exact style/document inputs were not retained: %#v", result)
	}
	if result["settings_xml"] != string(p.Files["word/settings.xml"]) || result["settings_sha256"] != office.Hash(p.Files["word/settings.xml"]) {
		t.Fatalf("exact settings input was not retained: %#v", result)
	}
}

func TestStructureResolvesParagraphNumberingDefinition(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Numbered</w:t></w:r></w:p></w:body></w:document>`)
	p.Files["word/numbering.xml"] = []byte(`<w:numbering xmlns:w="` + office.W + `"><w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="1"><w:start w:val="2"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="3"/><w:lvlOverride w:ilvl="1"><w:startOverride w:val="4"/></w:lvlOverride></w:num></w:numbering>`)
	b, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "numbering.docx")
	if err = os.WriteFile(path, b, 0600); err != nil {
		t.Fatal(err)
	}
	result, err := StructureReference(path)
	if err != nil {
		t.Fatal(err)
	}
	evidence := result["paragraphs"].([]map[string]any)[0]["numbering_evidence"].(map[string]any)
	if evidence["level"] != 2 || evidence["family"] != "lowerLetter" || evidence["label_pattern"] != "%2." || evidence["start"] != 4 || evidence["start_override"] != true {
		t.Fatalf("unresolved numbering evidence: %v", evidence)
	}
	if result["numbering_xml"] != string(p.Files["word/numbering.xml"]) || result["numbering_sha256"] != office.Hash(p.Files["word/numbering.xml"]) {
		t.Fatalf("exact numbering input was not retained: %#v", result)
	}
}

func TestStructureRetainsDisabledNumberingContradiction(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:pPr><w:outlineLvl w:val="1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>Heading without numbering</w:t></w:r></w:p></w:body></w:document>`)
	b, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "disabled-numbering.docx")
	if err = os.WriteFile(path, b, 0600); err != nil {
		t.Fatal(err)
	}
	result, err := StructureResolved(path)
	if err != nil {
		t.Fatal(err)
	}
	row := result["paragraphs"].([]map[string]any)[0]
	numbering, ok := row["numbering_evidence"].(map[string]any)
	if !ok || numbering["disabled"] != true || numbering["num_id"] != "0" {
		t.Fatalf("explicit disabled numbering was dropped: %v", row)
	}
	resolved := row["resolved_structure"].(map[string]any)
	contradictions, ok := resolved["contradictions"].([]string)
	if !ok || len(contradictions) != 1 {
		t.Fatalf("heading/numbering contradiction was not retained: %v", resolved)
	}
}

func TestStructureReportsConflictingDuplicateStyleIDs(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/styles.xml"] = []byte(`<w:styles xmlns:w="` + office.W + `"><w:style w:type="paragraph" w:styleId="Duplicate"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Duplicate"><w:pPr><w:outlineLvl w:val="8"/></w:pPr></w:style></w:styles>`)
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:pPr><w:pStyle w:val="Duplicate"/></w:pPr><w:r><w:t>Ambiguous</w:t></w:r></w:p></w:body></w:document>`)
	b, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "duplicate-style.docx")
	if err = os.WriteFile(path, b, 0600); err != nil {
		t.Fatal(err)
	}
	result, err := StructureReference(path)
	if err != nil {
		t.Fatal(err)
	}
	rows := result["paragraphs"].([]map[string]any)
	ambiguity, ok := rows[0]["style_ambiguity"].(map[string]any)
	if !ok || ambiguity["kind"] != "conflicting_duplicate_style_id" {
		t.Fatalf("duplicate style ambiguity missing: %v", rows[0])
	}
	duplicates, ok := result["style_duplicates"].([]map[string]any)
	if !ok || len(duplicates) != 1 || duplicates[0]["style_id"] != "Duplicate" || duplicates[0]["conflicting"] != true {
		t.Fatalf("duplicate style evidence missing: %v", result["style_duplicates"])
	}
	definitions, ok := duplicates[0]["definitions"].([]map[string]any)
	if !ok || len(definitions) != 2 || definitions[0]["sha256"] == definitions[1]["sha256"] {
		t.Fatalf("duplicate definitions were not retained: %v", duplicates[0])
	}
}

func TestStructureReferenceAcceptsDocumentWithoutStylesPart(t *testing.T) {
	p := office.BlankPackage()
	delete(p.Files, "word/styles.xml")
	p.Files[office.RelPart("word/document.xml")] = []byte(`<Relationships xmlns="` + office.RelNS + `"/>`)
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>`)
	b, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "no-styles.docx")
	if err = os.WriteFile(path, b, 0600); err != nil {
		t.Fatal(err)
	}
	result, err := StructureReference(path)
	if err != nil {
		t.Fatal(err)
	}
	rows := result["paragraphs"].([]map[string]any)
	if len(rows) != 1 || rows[0]["text"] != "Body" {
		t.Fatalf("document without styles was not inspected: %v", rows)
	}
}

func TestNumberingLabelHonorsWordRestartRules(t *testing.T) {
	definition := numberingDefinition{Levels: map[int]numberingLevel{
		0: {Family: "decimal", Pattern: "%1.", Start: 1, RestartAfter: -1},
		1: {Family: "lowerLetter", Pattern: "%2.", Start: 1, RestartAfter: -1},
		2: {Family: "lowerRoman", Pattern: "%3.", Start: 1, RestartAfter: 0},
	}}
	state := &numberingState{}
	levels := []int{0, 1, 2, 0, 1, 2}
	want := []string{"1.", "a.", "i.", "2.", "a.", "ii."}
	for i, level := range levels {
		got, certain := numberingLabel(definition, level, state)
		if !certain || got != want[i] {
			t.Fatalf("step %d level %d: label=%q certain=%v want %q", i, level, got, certain, want[i])
		}
	}

	definition.Levels[2] = numberingLevel{Family: "lowerRoman", Pattern: "%3.", Start: 1, RestartAfter: 2}
	state = &numberingState{}
	for i, level := range []int{0, 1, 2, 0, 2} {
		got, certain := numberingLabel(definition, level, state)
		want := []string{"1.", "a.", "i.", "2.", "i."}[i]
		if !certain || got != want {
			t.Fatalf("explicit restart step %d level %d: label=%q certain=%v want %q", i, level, got, certain, want)
		}
	}
}

func TestStructureUsesNumberingLevelOverrideFormatting(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>`)
	p.Files["word/numbering.xml"] = []byte(`<w:numbering xmlns:w="` + office.W + `"><w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="3"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/><w:lvl w:ilvl="0"><w:start w:val="2"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1)"/></w:lvl></w:lvlOverride></w:num></w:numbering>`)
	b, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "numbering-override.docx")
	if err = os.WriteFile(path, b, 0600); err != nil {
		t.Fatal(err)
	}
	result, err := StructureReference(path)
	if err != nil {
		t.Fatal(err)
	}
	rows := result["paragraphs"].([]map[string]any)
	if len(rows) != 2 {
		t.Fatal(rows)
	}
	first := rows[0]["numbering_evidence"].(map[string]any)
	if first["family"] != "upperRoman" || first["label_pattern"] != "%1)" || first["start"] != 4 || first["start_override"] != true || first["level_override"] != true || first["displayed_label"] != "IV)" {
		t.Fatalf("level override formatting was not applied: %#v", first)
	}
	second := rows[1]["numbering_evidence"].(map[string]any)
	if second["displayed_label"] != "V)" {
		t.Fatalf("overridden numbering counter did not continue: %#v", second)
	}
}
