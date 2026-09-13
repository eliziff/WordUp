package office

import (
	"bytes"
	"testing"
)

func TestUTF16ValidationPreservesCustomPart(t *testing.T) {
	p := BlankPackage()
	original := append([]byte{255, 254}, utf16bytes(`<?xml version="1.0" encoding="utf-16"?><properties><name>Journal</name></properties>`)...)
	p.Files["customXml/item2.xml"] = original
	if err := p.Validate(); err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(p.Files["customXml/item2.xml"], original) {
		t.Fatal("rewrote original XML")
	}
	if _, err := validationXMLSpans([]byte{255, 254, 0, 216}); err == nil {
		t.Fatal("accepted unpaired surrogate")
	}
}

func TestUTF8BOMRetainsEditOffsets(t *testing.T) {
	b := append([]byte{0xef, 0xbb, 0xbf}, []byte(`<?xml version="1.0"?><root><child/></root>`)...)
	spans, err := XMLSpans(b)
	if err != nil {
		t.Fatal(err)
	}
	if string(b[spans[1].Start:spans[1].End]) != "<child/>" {
		t.Fatal("BOM changed source offsets", spans)
	}
	if _, err = XMLSpans([]byte("<root/>\xef\xbb\xbf")); err == nil {
		t.Fatal("accepted trailing BOM")
	}
}
