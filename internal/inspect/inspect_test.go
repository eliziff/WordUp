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

func TestParagraphSourceLocationsAndNestedText(t *testing.T) {
	xml := `<w:document xmlns:w="` + office.W + `" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:x="urn:foreign"><w:body><w:p w14:paraId="1234ABCD"><w:pPr><w:pStyle w:val="Title"/><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Author</w:t><w:tab/><w:t>Name</w:t><w:br/><w:footnoteReference w:id="7"/><x:t>not Word text</x:t><w:drawing><w:txbxContent><w:p><w:pPr><w:pStyle w:val="Textbox"/></w:pPr><w:r><w:t>Nested</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r></w:p><w:p><w:r><w:t>Body &amp; text</w:t></w:r></w:p></w:body></w:document>`
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
}
