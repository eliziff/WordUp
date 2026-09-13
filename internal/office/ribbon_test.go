package office

import (
	"bytes"
	"testing"
)

func TestConnectRibbons(t *testing.T) {
	p := &Package{Files: map[string][]byte{"ui/tools.xml": []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"/>`)}}
	if err := p.ConnectRibbons(); err != nil {
		t.Fatal(err)
	}
	before := append([]byte(nil), p.Files["_rels/.rels"]...)
	if !bytes.Contains(before, []byte(`Target="ui/tools.xml"`)) {
		t.Fatal("Ribbon is disconnected")
	}
	if err := p.ConnectRibbons(); err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(before, p.Files["_rels/.rels"]) {
		t.Fatal("existing relationships changed")
	}
	p.Files["ui/duplicate.xml"] = p.Files["ui/tools.xml"]
	if err := p.ConnectRibbons(); err == nil {
		t.Fatal("ambiguous Ribbon accepted")
	}
}
