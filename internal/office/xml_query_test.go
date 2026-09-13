package office

import (
	"context"
	"encoding/json"
	"strings"
	"testing"
)

func TestQueryXMLExactLocationsAndNamespaces(t *testing.T) {
	data := []byte(`<x:document xmlns:x="urn:word"><!--keep--><x:p><x:r x:font="Garamond">é中</x:r></x:p><x:p/></x:document>`)
	ns := map[string]string{"w": "urn:word"}
	for _, query := range []string{"//w:r", "//w:r/@w:font", "//w:r/text()"} {
		got, err := QueryXML(context.Background(), data, query, ns, 20)
		if err != nil {
			t.Fatal(err)
		}
		matches := got["matches"].([]any)
		if len(matches) != 1 || got["source_sha256"] != Hash(data) {
			t.Fatalf("wrong result: %+v", got)
		}
		m := matches[0].(map[string]any)
		raw := string(data[m["element_start"].(int):m["element_end"].(int)])
		if raw != `<x:r x:font="Garamond">é中</x:r>` || m["element_xml"] != raw {
			t.Fatalf("lost source bytes: %v", m)
		}
	}
	for query, expected := range map[string]any{"count(//w:p)": float64(2), "boolean(//w:r[@w:font='Garamond'])": true, "string(//w:r)": "é中"} {
		got, err := QueryXML(context.Background(), data, query, ns, 20)
		if err != nil || got["value"] != expected {
			t.Fatalf("%s: %v %v", query, got, err)
		}
	}
}

func TestQueryXMLBudgetsAndRejections(t *testing.T) {
	data := []byte("<r>" + strings.Repeat("<p>"+strings.Repeat("中", 1000)+"</p>", 3) + "</r>")
	got, err := QueryXML(context.Background(), data, "//p", nil, 2)
	if err != nil || got["truncated"] != true || got["returned"] != 2 {
		t.Fatalf("missing result bound: %v %v", got, err)
	}
	encoded, err := json.Marshal(got)
	if err != nil || len(encoded) > 4000 {
		t.Fatal("unbounded match previews")
	}
	for _, input := range []string{`<!DOCTYPE r SYSTEM "file:///secret"><r/>`, `<r>`, `<r/><r/>`, "<r>" + strings.Repeat("<p>", 260) + strings.Repeat("</p>", 260) + "</r>"} {
		if _, err := QueryXML(context.Background(), []byte(input), "//*", nil, 1); err == nil {
			t.Fatal("accepted malformed, DTD, or over-depth XML")
		}
	}
	for _, query := range []string{"", "//*[", "1 div 0"} {
		if _, err := QueryXML(context.Background(), []byte("<r/>"), query, nil, 1); err == nil {
			t.Fatalf("accepted invalid/non-JSON result: %q", query)
		}
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if _, err := QueryXML(ctx, data, "//*", nil, 1); err == nil {
		t.Fatal("ignored cancellation")
	}
}

func BenchmarkQueryXMLManuscript(b *testing.B) {
	data := []byte(`<w:document xmlns:w="urn:word">` + strings.Repeat(`<w:p><w:r><w:t>Sixty thousand characters of manuscript text.</w:t></w:r></w:p>`, 1400) + `</w:document>`)
	for i := 0; i < b.N; i++ {
		if _, err := QueryXML(context.Background(), data, "count(//w:p)", map[string]string{"w": "urn:word"}, 20); err != nil {
			b.Fatal(err)
		}
	}
}
