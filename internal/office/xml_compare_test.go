package office

import (
	"encoding/json"
	"encoding/xml"
	"strings"
	"testing"
)

func TestComparisonContextLocatesFormattingAndBoundsLargeText(t *testing.T) {
	a := []byte(`<w:p xmlns:w="word"><w:r><w:rPr><w:b/></w:rPr><w:t> </w:t></w:r></w:p>`)
	b := []byte(strings.Replace(string(a), "w:b", "w:i", 1))
	r, err := CompareXML(a, b, XMLComparePolicy{})
	if err != nil {
		t.Fatal(err)
	}
	c := r["difference_context"].(map[string]any)["reference"].(map[string]any)
	start, end := c["byte_start"].(int64), c["byte_end"].(int64)
	if string(a[start:end]) != "<w:b/>" || len(c["ancestors"].([]map[string]any)) < 3 {
		t.Fatalf("missing exact formatting location: %v", c)
	}
	r, err = CompareXML([]byte("<p>"+strings.Repeat("x", 10000)+"</p>"), []byte("<p>y</p>"), XMLComparePolicy{})
	if err != nil {
		t.Fatal(err)
	}
	c = r["difference_context"].(map[string]any)["reference"].(map[string]any)
	if len(c["xml"].(string)) > 2048 {
		t.Fatal("unbounded context")
	}
	r, err = CompareXML(a, a, XMLComparePolicy{})
	if err != nil || r["difference_context"] != nil {
		t.Fatal("context generated for identical XML", r, err)
	}
}

func TestFlatOPCPartEvidenceIncludesLaterAndMissingParts(t *testing.T) {
	a := []byte(`<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/body"><x id="1"/></pkg:part><pkg:part pkg:name="/notes"><x>old</x></pkg:part></pkg:package>`)
	b := []byte(`<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/body"><x id="2"/></pkg:part><pkg:part pkg:name="/notes"><x>new</x></pkg:part><pkg:part pkg:name="/added"><x/></pkg:part></pkg:package>`)
	r, err := CompareXML(a, b, XMLComparePolicy{})
	if err != nil {
		t.Fatal(err)
	}
	rows := r["parts"].([]map[string]any)
	if len(rows) != 3 || rows[0]["name"] != "/added" || rows[0]["reference_sha256"] != "" || rows[2]["name"] != "/notes" || rows[2]["byte_identical"] != false {
		t.Fatalf("lost changes after first difference: %v", rows)
	}
	change := rows[2]["first_difference"].(map[string]any)
	before := change["reference"].(map[string]any)
	after := change["candidate"].(map[string]any)
	if before["token"] != `text "old"` || after["token"] != `text "new"` {
		t.Fatalf("metadata mismatch hid later part text: %v", change)
	}
	start, end := before["byte_start"].(int64), before["byte_end"].(int64)
	if string(a[start:end]) != "old" {
		t.Fatal("part offset is not relative to original XML")
	}
	selected, err := CompareXML(a, b, XMLComparePolicy{ContextParts: []string{"/notes"}})
	if err != nil || selected["equal"] != false || selected["byte_identical"] != false {
		t.Fatal(selected, err)
	}
	filtered := selected["parts"].([]map[string]any)
	if filtered[1]["context_omitted"] != true || filtered[2]["first_difference"] == nil {
		t.Fatal("context selection changed evidence scope", filtered)
	}
}

func TestXMLComparisonNeverIgnoresUnrequestedChanges(t *testing.T) {
	a := []byte(`<w:p xmlns:w="word" xmlns:x="metadata" x:id="1"><w:b/><w:t>keep</w:t></w:p>`)
	b := []byte(`<z:p xmlns:z="word" xmlns:x="metadata" x:id="2"><z:b/><z:t>keep</z:t></z:p>`)
	strict, e := CompareXML(a, b, XMLComparePolicy{})
	if e != nil || strict["equal"] != false {
		t.Fatal(strict, e)
	}
	policy := XMLComparePolicy{Attributes: []xml.Name{{Space: "metadata", Local: "id"}}}
	same, e := CompareXML(a, b, policy)
	if e != nil || same["equal"] != true || same["byte_identical"] != false {
		t.Fatal(same, e)
	}
	for _, changed := range []string{`<p xmlns="word"><i/><t>keep</t></p>`, `<p xmlns="word"><b/><t>changed</t></p>`} {
		result, e := CompareXML(a, []byte(changed), policy)
		if e != nil || result["equal"] != false {
			t.Fatal(result, e)
		}
	}
	if _, e := CompareXML([]byte(`<!DOCTYPE p><p/>`), b, policy); e == nil {
		t.Fatal("accepted DTD")
	}
}

func TestLargeTokenDiagnosticsDoNotHideLateDifference(t *testing.T) {
	prefix := `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/body"><p>`
	suffix := `</p></pkg:part></pkg:package>`
	shared := strings.Repeat("\u00e9", 50000)
	result, err := CompareXML([]byte(prefix+shared+"a"+suffix), []byte(prefix+shared+"b"+suffix), XMLComparePolicy{})
	if err != nil || result["equal"] != false {
		t.Fatal(result, err)
	}
	encoded, err := json.Marshal(result)
	if err != nil || len(encoded) > 20000 {
		t.Fatalf("unbounded failure report: %d bytes, %v", len(encoded), err)
	}
	difference := result["first_difference"].(map[string]any)
	if difference["reference_token_bytes"].(int) < 100000 || !strings.HasPrefix(difference["reference"].(string), "[truncated]...") {
		t.Fatal("lost truncation evidence")
	}
	if !strings.HasSuffix(difference["reference"].(string), `a"`) || !strings.HasSuffix(difference["candidate"].(string), `b"`) || difference["canonical_token_byte"].(int) < 100000 {
		t.Fatal("preview hid the actual mismatch", difference)
	}
}

func TestGeneratedTOCBookmarkPolicyPreservesTargets(t *testing.T) {
	a := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:hyperlink w:anchor="_Toc123"/><w:instrText> PAGEREF _Toc456 \h </w:instrText><w:bookmarkStart w:name="_Toc123" w:id="1"/><w:t>First</w:t><w:bookmarkEnd w:id="1"/><w:bookmarkStart w:name="_Toc456" w:id="2"/><w:t>Second</w:t><w:bookmarkEnd w:id="2"/></w:document>`
	b := strings.NewReplacer("_Toc123", "_Toc900", "_Toc456", "_Toc901").Replace(a)
	policy := XMLComparePolicy{GeneratedTOCBookmarks: true}
	result, err := CompareXML([]byte(a), []byte(b), policy)
	if err != nil || result["equal"] != true || result["byte_identical"] != false {
		t.Fatal(result, err)
	}
	for _, wrong := range []string{
		strings.Replace(b, `w:anchor="_Toc900"`, `w:anchor="_Toc901"`, 1),
		strings.Replace(b, `PAGEREF _Toc901`, `PAGEREF _Toc900`, 1),
		strings.Replace(b, `First`, `Changed`, 1),
		strings.Replace(b, `w:bookmarkEnd w:id="1"`, `w:bookmarkEnd w:id="2"`, 1),
	} {
		result, err = CompareXML([]byte(a), []byte(wrong), policy)
		if err != nil || result["equal"] != false {
			t.Fatal("hid changed target/content/boundary", result, err)
		}
	}
	result, err = CompareXML([]byte(a), []byte(b), XMLComparePolicy{})
	if err != nil || result["equal"] != false {
		t.Fatal("normalization was not opt-in", result, err)
	}
	if _, err = CompareXML([]byte(strings.Replace(a, "_Toc456", "_Toc123", -1)), []byte(b), policy); err == nil {
		t.Fatal("accepted duplicate bookmark names")
	}
}

func TestAttributeEquivalenceRequiresExactElementAndValue(t *testing.T) {
	a := []byte(`<root><relationship target="session-a/template.dotm"/><link target="session-a/template.dotm"/></root>`)
	b := []byte(`<root><relationship target="session-b/template.dotm"/><link target="session-a/template.dotm"/></root>`)
	policy := XMLComparePolicy{AttributeEquivalences: []XMLAttributeEquivalence{{Element: xml.Name{Local: "relationship"}, Attribute: xml.Name{Local: "target"}, Values: []string{"session-a/template.dotm", "session-b/template.dotm"}}}}
	result, err := CompareXML(a, b, policy)
	if err != nil || result["equal"] != true || result["byte_identical"] != false {
		t.Fatal(result, err)
	}
	for _, candidate := range [][]byte{
		[]byte(strings.Replace(string(b), "session-b/template.dotm", "unexpected/template.dotm", 1)),
		[]byte(strings.Replace(string(b), `link target="session-a/template.dotm"`, `link target="session-b/template.dotm"`, 1)),
	} {
		result, err = CompareXML(a, candidate, policy)
		if err != nil || result["equal"] != false {
			t.Fatal("equivalence hid an unexpected change", result, err)
		}
	}
}
