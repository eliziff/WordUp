package office

import (
	"bytes"
	"testing"
)

// Word refuses freshly written forms and modules that omit records its own
// designer always emits. These offline checks pin the interoperable bytes; the
// native selftest proves Word actually loads the result.
func TestNativeFormRootRecords(t *testing.T) {
	f, e := NewForm("Fresh", 1252)
	if e != nil {
		t.Fatal(e)
	}
	if e = f.Apply(Design{Name: "Fresh", Mode: "replace", Properties: map[string]any{"StartUpPosition": 1.}, Controls: []ControlDesign{{Name: "Run", Type: "CommandButton"}}}); e != nil {
		t.Fatal(e)
	}
	if got := f.root.record.values["BooleanProperties"]; got != 0x4004 {
		t.Fatalf("root BooleanProperties = %#v, want 0x4004", got)
	}
	streams, e := f.Streams()
	if e != nil {
		t.Fatal(e)
	}
	if !bytes.Contains(streams["f"], nativeRootFontAsset()) {
		t.Fatal("root f stream lacks the StdFont Tahoma asset")
	}
	if !bytes.HasSuffix(streams["f"], formDesignExtenderRoot) {
		t.Fatal("root f stream lacks the FormDesignExtender trailer")
	}
	if !bytes.Contains(streams["\x03VBFrame"], []byte("StartUpPosition = 1\r")) && !bytes.Contains(streams["\x03VBFrame"], []byte("StartUpPosition = 1\n")) {
		t.Fatalf("StartUpPosition must serialize as an integer: %q", streams["\x03VBFrame"])
	}
}

func TestModuleRecordsCookieAndPrivateForm(t *testing.T) {
	form, e := moduleRecords(Module{Name: "frmOne", Kind: "form", Source: "Attribute VB_Name = \"frmOne\"\n"}, 1252)
	if e != nil {
		t.Fatal(e)
	}
	std, e := moduleRecords(Module{Name: "Main", Kind: "standard", Source: "Attribute VB_Name = \"Main\"\n"}, 1252)
	if e != nil {
		t.Fatal(e)
	}
	if !bytes.Contains(form, tlv(0x28, nil)) {
		t.Fatal("form module lacks MODULEPRIVATE")
	}
	if bytes.Contains(std, tlv(0x28, nil)) {
		t.Fatal("standard module must not be private")
	}
	for _, m := range []Module{{Name: "frmOne", Kind: "form"}, {Name: "Main", Kind: "standard"}} {
		c := moduleCookie(m)
		if c == 0 || c == 0xFFFF {
			t.Fatalf("MODULECOOKIE for %s = %#x; MS-OVBA requires 0 < cookie < 0xFFFF", m.Name, c)
		}
	}
	if bytes.Contains(std, tlv(0x2c, word(0))) {
		t.Fatal("standard module still writes a zero MODULECOOKIE")
	}
}
