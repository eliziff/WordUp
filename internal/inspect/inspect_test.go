package inspect

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

func TestRibbonDiagnosticsUseXMLIdentity(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("RibbonCheck", root); err != nil {
		t.Fatal(err)
	}
	files := map[string]string{
		"tools.xml":         `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" loadImage="MissingImage"><ribbon><tabs><tab id="tools" label="Tools"><group id="tools" label="Group"><button id="action" onAction="MissingAction"/></group></tab></tabs></ribbon></customUI>`,
		"customUI-data.xml": `<data id="same" onAction="OrdinaryData"><item id="same"/></data>`,
	}
	for name, value := range files {
		if err := os.WriteFile(filepath.Join(root, "package", name), []byte(value), 0600); err != nil {
			t.Fatal(err)
		}
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	r, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	messages := ""
	for _, diagnostic := range r["diagnostics"].([]map[string]any) {
		if diagnostic["file"] == "package/customUI-data.xml" {
			t.Fatal("treated ordinary XML as Ribbon", diagnostic)
		}
		if diagnostic["file"] == "package/tools.xml" {
			messages += diagnostic["message"].(string) + "\n"
		}
		if diagnostic["callback"] == "MissingImage" && diagnostic["expected_declaration"] != "Public Sub MissingImage(imageID As String, ByRef returnedVal)" {
			t.Fatal("missing actionable image callback declaration", diagnostic)
		}
		if diagnostic["callback"] == "MissingAction" && (diagnostic["control_id"] != "action" || diagnostic["expected_declaration"] != "Public Sub MissingAction(control As IRibbonControl)") {
			t.Fatal("missing action callback location or declaration", diagnostic)
		}
	}
	for _, want := range []string{"duplicate Ribbon id tools", "MissingImage", "MissingAction"} {
		if !strings.Contains(messages, want) {
			t.Fatal("missing Ribbon diagnostic", want, messages)
		}
	}
}

func TestCheckDiagnosticsAreStableAcrossRuns(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("StableCheck", root); err != nil {
		t.Fatal(err)
	}
	for name, source := range map[string]string{
		"vba/Zed.bas":  "Attribute VB_Name = \"Zed\"\nOption Explicit\nPublic Sub Shared(\nEnd Sub\n",
		"vba/Alpha.bas": "Attribute VB_Name = \"Alpha\"\nOption Explicit\nPublic Sub Shared(\nEnd Sub\n",
	} {
		if err := project.Write(root, name, []byte(source), ""); err != nil {
			t.Fatal(err)
		}
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	first, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	second, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	firstDiagnostics, ok := first["diagnostics"].([]map[string]any)
	if !ok {
		t.Fatalf("first check diagnostics have wrong shape: %#v", first["diagnostics"])
	}
	secondDiagnostics, ok := second["diagnostics"].([]map[string]any)
	if !ok || len(firstDiagnostics) != len(secondDiagnostics) {
		t.Fatalf("check diagnostics changed shape or count: %#v %#v", first["diagnostics"], second["diagnostics"])
	}
	for i := range firstDiagnostics {
		if string(project.JSON(firstDiagnostics[i])) != string(project.JSON(secondDiagnostics[i])) {
			t.Fatalf("check diagnostics changed order at %d: %#v vs %#v", i, firstDiagnostics, secondDiagnostics)
		}
	}
	if len(firstDiagnostics) == 0 {
		t.Fatalf("check did not report the deliberate duplicate procedure: %#v", firstDiagnostics)
	}
	fileOrder := []string{}
	for _, diagnostic := range firstDiagnostics {
		if file, ok := diagnostic["file"].(string); ok && strings.HasPrefix(file, "vba/") {
			fileOrder = append(fileOrder, file)
		}
	}
	if len(fileOrder) < 2 || fileOrder[0] != "vba/Alpha.bas" || fileOrder[1] != "vba/Zed.bas" {
		t.Fatalf("diagnostics were not source-ordered: %v", fileOrder)
	}
}

func TestRibbonDiagnosticsReportIncompatibleCallbackReuse(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("RibbonCallbackCheck", root); err != nil {
		t.Fatal(err)
	}
	data := `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="tab"><group id="group"><button id="plain" onAction="Shared"/><toggleButton id="toggle" onAction="Shared"/></group></tab></tabs></ribbon></customUI>`
	if err := os.WriteFile(filepath.Join(root, "package", "callbacks.xml"), []byte(data), 0600); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	r, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	for _, diagnostic := range r["diagnostics"].([]map[string]any) {
		if diagnostic["engine"] == "Office RibbonX callback mapping" && strings.Contains(diagnostic["message"].(string), "incompatible signatures") {
			return
		}
	}
	t.Fatal("missing incompatible Ribbon callback diagnostic", r["diagnostics"])
}

func TestRibbonDiagnosticsCheckCallbackDeclarationShape(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("RibbonDeclarationCheck", root); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(root, "vba", "Callbacks.bas"), []byte("Attribute VB_Name = \"Callbacks\"\nOption Explicit\nPublic Function BadAction(control As String) As Boolean\n    BadAction = True\nEnd Function\nPublic Sub WrongType(control As String)\nEnd Sub\n"), 0600); err != nil {
		t.Fatal(err)
	}
	data := `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="tab"><group id="group"><button id="bad" onAction="BadAction"/><button id="wrong-type" onAction="WrongType"/></group></tab></tabs></ribbon></customUI>`
	if err := os.WriteFile(filepath.Join(root, "package", "callbacks.xml"), []byte(data), 0600); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	r, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	kindMismatch, typeMismatch := false, false
	for _, diagnostic := range r["diagnostics"].([]map[string]any) {
		if diagnostic["engine"] != "Ribbon callback declaration shape" {
			continue
		}
		message := diagnostic["message"].(string)
		if diagnostic["callback"] == "BadAction" && strings.Contains(message, "must be Public Sub") {
			kindMismatch = true
		}
		if diagnostic["callback"] == "WrongType" && strings.Contains(message, "expects IRibbonControl") {
			typeMismatch = true
		}
	}
	if kindMismatch && typeMismatch {
		return
	}
	t.Fatalf("missing callback declaration shape diagnostics (kind=%v type=%v): %v", kindMismatch, typeMismatch, r["diagnostics"])
}

func TestCheckValidatesOPCRelationshipTargets(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	if _, err := project.New("PackageCheck", root); err != nil {
		t.Fatal(err)
	}
	relPath := "package/_rels/.rels"
	rels, err := project.Read(root, relPath)
	if err != nil {
		t.Fatal(err)
	}
	rels = []byte(strings.Replace(string(rels), `Target="word/document.xml"`, `Target="word/missing.xml"`, 1))
	if err := project.Write(root, relPath, rels, ""); err != nil {
		t.Fatal(err)
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	result, err := Check(w)
	if err != nil {
		t.Fatal(err)
	}
	validation := result["package_validation"].(map[string]any)
	if validation["checked"] != true || validation["valid"] != false || !strings.Contains(validation["error"].(string), "missing relationship target") {
		t.Fatalf("unexpected package validation: %#v", validation)
	}
	for _, diagnostic := range result["diagnostics"].([]map[string]any) {
		if diagnostic["engine"] == "Open Packaging Conventions" && strings.Contains(diagnostic["message"].(string), "missing relationship target") {
			return
		}
	}
	t.Fatal("missing OPC package diagnostic", result["diagnostics"])
}

func TestParagraphSourceLocationsAndNestedText(t *testing.T) {
	xml := `<w:document xmlns:w="` + office.W + `" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:x="urn:foreign"><w:body><w:p w14:paraId="1234ABCD"><w:pPr><w:pStyle w:val="Title"/><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs><w:rPr><w:i/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Author</w:t><w:tab/><w:t>Name</w:t><w:br/><w:footnoteReference w:id="7"/><x:t>not Word text</x:t><w:drawing><w:txbxContent><w:p><w:pPr><w:pStyle w:val="Textbox"/></w:pPr><w:r><w:t>Nested</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r></w:p><w:p><w:r><w:t>Body &amp; text</w:t></w:r></w:p></w:body></w:document>`
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(xml)
	rows, err := TextObservations(p)
	if err != nil || len(rows) != 3 {
		t.Fatalf("paragraph inventory: %v %v", rows, err)
	}
	if rows[0]["text"] != "Author\tName\n" || rows[0]["style_id"] != "Title" || rows[1]["text"] != "Nested" || rows[2]["text"] != "Body & text" {
		t.Fatalf("mixed nested/foreign observations: %v", rows)
	}
	if rows[0]["paragraph_id"] != "1234ABCD" || rows[2]["xml_path"] != "(//w:p)[3]" {
		t.Fatalf("lost source identity: %v", rows)
	}
	start, end := rows[1]["xml_start"].(int), rows[1]["xml_end"].(int)
	if !strings.HasPrefix(xml[start:end], "<w:p>") || !strings.Contains(xml[start:end], ">Nested<") {
		t.Fatal("invalid XML byte range")
	}
	refs := rows[0]["references"].([]map[string]any)
	if len(refs) != 1 || refs[0]["id"] != "7" || refs[0]["kind"] != "footnoteReference" {
		t.Fatalf("lost author note reference: %v", refs)
	}
	evidence := rows[0]["direct_formatting_evidence"].(map[string]int)
	if evidence["text_units"] != 12 || evidence["bold_units"] != 12 {
		t.Fatalf("character-weighted formatting includes nested text or loses controls: %v", evidence)
	}
	mark := rows[0]["paragraph_mark_formatting"].(map[string]bool)
	if !mark["italic"] {
		t.Fatalf("lost paragraph-mark formatting evidence: %v", mark)
	}
}

func TestTextObservationsAcceptEmptyTextElements(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:r><w:t/><w:t>kept</w:t><w:instrText/></w:r></w:p></w:body></w:document>`)
	rows, err := TextObservations(p)
	if err != nil || len(rows) != 1 {
		t.Fatalf("empty Word text element was rejected: %v %v", rows, err)
	}
	if rows[0]["text"] != "kept" {
		t.Fatalf("empty text element changed displayed text: %#v", rows[0]["text"])
	}
}

func TestRunEvidencePreservesBreakCharacters(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:r><w:t>first</w:t><w:br/><w:t>second</w:t><w:tab/><w:t>third</w:t></w:r></w:p></w:body></w:document>`)
	rows, err := TextObservations(p)
	if err != nil || len(rows) != 1 {
		t.Fatalf("break observations: %v %v", rows, err)
	}
	runs, ok := rows[0]["run_properties"].([]map[string]any)
	if !ok || len(runs) != 1 {
		t.Fatalf("run evidence missing: %#v", rows[0]["run_properties"])
	}
	if rows[0]["text"] != "first\nsecond\tthird" {
		t.Fatalf("paragraph text lost break semantics: %q", rows[0]["text"])
	}
	if runs[0]["text_units"] != 18 {
		t.Fatalf("run text units lost break semantics: %#v", runs[0])
	}
}

func TestTrackedRevisionEvidenceDoesNotDuplicateText(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:ins w:id="4" w:author="Editor" w:date="2026-01-02T03:04:05Z"><w:r><w:t>new</w:t></w:r></w:ins><w:del w:id="5" w:author="Editor"><w:r><w:delText>old</w:delText></w:r></w:del></w:p></w:body></w:document>`)
	rows, err := TextObservations(p)
	if err != nil || len(rows) != 1 {
		t.Fatalf("revision observations: %v %v", rows, err)
	}
	if rows[0]["text"] != "new" {
		t.Fatalf("deleted text leaked into displayed text: %v", rows[0]["text"])
	}
	revisions, ok := rows[0]["revision_evidence"].([]map[string]any)
	if !ok || len(revisions) != 2 {
		t.Fatalf("missing revision evidence: %v", rows[0]["revision_evidence"])
	}
	if revisions[0]["kind"] != "ins" || revisions[0]["author"] != "Editor" || revisions[0]["text_units"] != 3 || revisions[1]["kind"] != "del" || revisions[1]["text_units"] != 3 {
		t.Fatalf("incorrect revision evidence: %v", revisions)
	}
}

func TestTextObservationsKeepFieldLinkAndBookmarkEvidenceCompact(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `" xmlns:r="` + office.R + `"><w:body><w:p><w:bookmarkStart w:id="1" w:name="Cite"/><w:hyperlink r:id="rId5" w:anchor="Source"><w:r><w:t>citation</w:t></w:r></w:hyperlink><w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> CITATION Source </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>Source</w:t><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="1"/></w:p></w:body></w:document>`)
	p.Files["word/_rels/document.xml.rels"] = []byte(`<Relationships xmlns="` + office.RelNS + `"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/source" TargetMode="External"/></Relationships>`)
	rows, err := TextObservations(p)
	if err != nil || len(rows) != 1 {
		t.Fatalf("field/link observations: %v %v", rows, err)
	}
	if rows[0]["text"] != "citationSource" {
		t.Fatalf("displayed text changed: %#v", rows[0]["text"])
	}
	links, ok := rows[0]["hyperlink_evidence"].([]map[string]any)
	if !ok || len(links) != 1 || links[0]["relationship_id"] != "rId5" || links[0]["anchor"] != "Source" || links[0]["relationship_target"] != "https://example.test/source" || links[0]["relationship_target_mode"] != "External" {
		t.Fatalf("link evidence missing: %#v", rows[0]["hyperlink_evidence"])
	}
	fields, ok := rows[0]["field_evidence"].([]map[string]any)
	if !ok || len(fields) != 4 || fields[1]["instruction"] != " CITATION Source " || fields[0]["marker"] != "begin" || fields[3]["marker"] != "end" {
		t.Fatalf("field evidence missing: %#v", rows[0]["field_evidence"])
	}
	bookmarks, ok := rows[0]["bookmark_evidence"].([]map[string]any)
	if !ok || len(bookmarks) != 2 || bookmarks[0]["name"] != "Cite" || bookmarks[1]["kind"] != "bookmarkEnd" {
		t.Fatalf("bookmark evidence missing: %#v", rows[0]["bookmark_evidence"])
	}
	for _, field := range fields {
		if _, duplicated := field["text"]; duplicated {
			t.Fatal("field evidence duplicated displayed text")
		}
	}
}

func TestTextObservationsKeepArtworkEvidenceSourceLocated(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="` + office.R + `"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="123" cy="456"/><wp:docPr id="4" name="Figure 4" descr="Figure alt"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`)
	p.Files["word/_rels/document.xml.rels"] = []byte(`<Relationships xmlns="` + office.RelNS + `"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/figure.png"/></Relationships>`)
	rows, err := TextObservations(p)
	if err != nil || len(rows) != 1 {
		t.Fatalf("artwork observations: %v %v", rows, err)
	}
	artwork, ok := rows[0]["artwork_evidence"].([]map[string]any)
	if !ok || len(artwork) != 1 || artwork[0]["kind"] != "drawing" || artwork[0]["relationship_id"] != "rId9" || artwork[0]["relationship_target"] != "media/figure.png" || artwork[0]["extent_cx"] != "123" || artwork[0]["extent_cy"] != "456" || artwork[0]["descr"] != "Figure alt" {
		t.Fatalf("artwork evidence missing: %#v", rows[0]["artwork_evidence"])
	}
}

func TestTextObservationsKeepContentControlEvidenceCompact(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:sdt><w:sdtPr><w:alias w:val="Author"/><w:tag w:val="author"/><w:id w:val="42"/><w:lock w:val="sdtContentLocked"/><w:dataBinding w:xpath="/author" w:storeItemID="{abc}"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Author</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>`)
	rows, err := TextObservations(p)
	if err != nil || len(rows) != 1 {
		t.Fatalf("content-control observations: %v %v", rows, err)
	}
	controls, ok := rows[0]["content_control_evidence"].([]map[string]any)
	if !ok || len(controls) != 1 {
		t.Fatalf("content-control evidence missing: %#v", rows[0]["content_control_evidence"])
	}
	control := controls[0]
	if control["alias"] != "Author" || control["tag"] != "author" || control["id"] != "42" || control["lock"] != "sdtContentLocked" || control["type"] != "text" {
		t.Fatalf("content-control metadata missing: %#v", control)
	}
	binding, ok := control["data_binding"].(map[string]string)
	if !ok || binding["xpath"] != "/author" || binding["storeItemID"] != "{abc}" {
		t.Fatalf("content-control binding missing: %#v", control["data_binding"])
	}
	if _, duplicated := control["text"]; duplicated {
		t.Fatal("content-control evidence duplicated paragraph text")
	}
}

func TestStoryObservationsKeepPartQualifiedLocations(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/header1.xml"] = []byte(`<w:hdr xmlns:w="` + office.W + `"><w:p w14:paraId="ABC" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Aptos"/><w:sz w:val="18"/></w:rPr><w:t>Running head</w:t></w:r></w:p></w:hdr>`)
	rows, err := textObservationsPart(p, "word/header1.xml")
	if err != nil || len(rows) != 1 {
		t.Fatalf("header observations: %v %v", rows, err)
	}
	if rows[0]["source_part"] != "word/header1.xml" || rows[0]["source_id"] != "word/header1.xml#paraId=ABC" || rows[0]["text"] != "Running head" {
		t.Fatalf("lost story identity: %v", rows[0])
	}
	if rows[0]["xml_path"] != "(//w:p)[1]" {
		t.Fatalf("unexpected story path: %v", rows[0]["xml_path"])
	}
	data, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	file := filepath.Join(t.TempDir(), "reference.docx")
	if err := os.WriteFile(file, data, 0600); err != nil {
		t.Fatal(err)
	}
	reference, err := StyleReference(file)
	if err != nil {
		t.Fatal(err)
	}
	stories, ok := reference["story_observations"].([]map[string]any)
	if !ok || len(stories) != 1 || stories[0]["part"] != "word/header1.xml" {
		t.Fatalf("public style reference omitted header story: %v", reference["story_observations"])
	}
}

func TestStyleReferenceIncludesThemeAndNumberingInputs(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`)
	p.Files["word/theme/theme1.xml"] = []byte(`<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>`)
	p.Files["word/numbering.xml"] = []byte(`<w:numbering xmlns:w="` + office.W + `"><w:abstractNum w:abstractNumId="1"/></w:numbering>`)
	p.Files["word/fontTable.xml"] = []byte(`<w:fonts xmlns:w="` + office.W + `"><w:font w:name="Aptos"><w:charset w:val="00"/></w:font></w:fonts>`)
	p.Files["customXml/item1.xml"] = []byte(`<item><value>opaque</value></item>`)
	data, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	file := filepath.Join(t.TempDir(), "reference.docx")
	if err := os.WriteFile(file, data, 0600); err != nil {
		t.Fatal(err)
	}
	reference, err := StyleReference(file)
	if err != nil {
		t.Fatal(err)
	}
	theme := reference["theme_fonts"].(map[string]string)
	if !strings.Contains(reference["font_table_xml"].(string), "Aptos") || reference["font_table_sha256"] != office.Hash(p.Files["word/fontTable.xml"]) {
		t.Fatal("lost font table input", reference)
	}
	if theme["majorHAnsi"] != "Aptos Display" || !strings.Contains(reference["theme_xml"].(string), "Aptos Display") || reference["theme_sha256"] != office.Hash(p.Files["word/theme/theme1.xml"]) || !strings.Contains(reference["styles_xml"].(string), "styleId=\"Normal\"") || !strings.Contains(reference["numbering_xml"].(string), "abstractNumId") || reference["styles_sha256"] != office.Hash(p.Files["word/styles.xml"]) || reference["numbering_sha256"] != office.Hash(p.Files["word/numbering.xml"]) {
		t.Fatalf("lost style cascade inputs: %#v", reference)
	}
	parts, ok := reference["part_inventory"].([]map[string]any)
	if !ok {
		t.Fatalf("missing part inventory: %#v", reference["part_inventory"])
	}
	found := false
	for _, part := range parts {
		if part["part"] == "customXml/item1.xml" && part["bytes"] == len(p.Files["customXml/item1.xml"]) && part["sha256"] == office.Hash(p.Files["customXml/item1.xml"]) {
			found = true
		}
	}
	if !found {
		t.Fatalf("opaque part missing from inventory: %#v", parts)
	}
	relationships, ok := reference["relationship_inventory"].([]map[string]any)
	if !ok || len(relationships) == 0 {
		t.Fatalf("missing relationship inventory: %#v", reference["relationship_inventory"])
	}
	foundDocumentStyles := false
	for _, group := range relationships {
		if group["source_part"] != "word/document.xml" {
			continue
		}
		for _, relation := range group["relationships"].([]map[string]any) {
			if relation["target_part"] == "word/styles.xml" {
				foundDocumentStyles = true
			}
		}
	}
	if !foundDocumentStyles {
		t.Fatalf("document relationship target was not resolved: %#v", relationships)
	}
	contentTypes, ok := reference["content_types"].([]map[string]any)
	if !ok || len(contentTypes) == 0 {
		t.Fatalf("missing content-type inventory: %#v", reference["content_types"])
	}
	foundDocumentType := false
	for _, item := range contentTypes {
		if item["part"] == "word/document.xml" && item["content_type"] == office.MainDOCX {
			foundDocumentType = true
		}
	}
	if !foundDocumentType {
		t.Fatalf("document content type was not retained: %#v", contentTypes)
	}
}

func TestStyleReferenceIncludesSectionAndDocumentSettings(t *testing.T) {
	p := office.BlankPackage()
	p.Files["word/document.xml"] = []byte(`<w:document xmlns:w="` + office.W + `"><w:body><w:p><w:r><w:t>Page one</w:t></w:r></w:p><w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="2"/></w:sectPr></w:pPr><w:r><w:t>Page two</w:t></w:r></w:p></w:body></w:document>`)
	p.Files["word/settings.xml"] = []byte(`<w:settings xmlns:w="` + office.W + `"><w:updateFields w:val="true"/></w:settings>`)
	p.Files["word/_rels/document.xml.rels"] = []byte(`<Relationships xmlns="` + office.RelNS + `"><Relationship Id="rIdHeader" Type="header" Target="header1.xml"/></Relationships>`)
	data, err := p.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	file := filepath.Join(t.TempDir(), "reference.docx")
	if err := os.WriteFile(file, data, 0600); err != nil {
		t.Fatal(err)
	}
	reference, err := StyleReference(file)
	if err != nil {
		t.Fatal(err)
	}
	sections, ok := reference["section_observations"].([]map[string]any)
	if !ok || len(sections) != 1 || sections[0]["index"] != 1 || !strings.Contains(sections[0]["xml"].(string), `w:num="2"`) {
		t.Fatalf("lost exact section geometry: %#v", reference["section_observations"])
	}
	if reference["document_xml_sha256"] != office.Hash(p.Files["word/document.xml"]) || reference["settings_part_sha256"] != office.Hash(p.Files["word/settings.xml"]) || reference["document_relationships_part_sha256"] != office.Hash(p.Files["word/_rels/document.xml.rels"]) {
		t.Fatalf("lost document-level source hashes: %#v", reference)
	}
}
