package office

import (
	"bytes"
	"strings"
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

func TestMergeRibbonXMLPreservesBaseAndAddsFragment(t *testing.T) {
	base := []byte(`<?xml version="1.0"?><customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="base" label="Base"><group id="baseGroup"><button id="baseButton" label="Keep"/></group></tab></tabs></ribbon></customUI>`)
	fragment := []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Load"><ribbon><tabs><tab id="added" label="Added"><group id="addedGroup"><button id="addedButton" label="Add" onAction="Run"/></group></tab></tabs></ribbon></customUI>`)
	merged, err := MergeRibbonXML(base, fragment)
	if err != nil {
		t.Fatal(err)
	}
	text := string(merged)
	for _, want := range []string{`label="Base"`, `id="baseButton" label="Keep"`, `onLoad="Load"`, `id="addedButton"`, `onAction="Run"`} {
		if !strings.Contains(text, want) {
			t.Fatalf("merged Ribbon lost %q: %s", want, text)
		}
	}
	if strings.Index(text, `id="base"`) > strings.Index(text, `id="added"`) {
		t.Fatalf("fragment changed existing child order: %s", text)
	}
	if _, err := XMLSpans(merged); err != nil {
		t.Fatal(err)
	}
	repeated, err := MergeRibbonXML(merged, fragment)
	if err != nil {
		t.Fatal("repeated merge:", err)
	}
	if !bytes.Equal(merged, repeated) {
		t.Fatal("unchanged Ribbon fragment was not idempotent")
	}
}

func TestMergeRibbonXMLRejectsCollisionsAndKeepsBase(t *testing.T) {
	base := []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="base" label="Old"/></tabs></ribbon></customUI>`)
	conflict := []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="base" label="New"/></tabs></ribbon></customUI>`)
	if _, err := MergeRibbonXML(base, conflict); err == nil || !strings.Contains(err.Error(), "attribute label conflicts") {
		t.Fatalf("attribute collision was not rejected: %v", err)
	}
	duplicate := []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="new"><group id="g"><button id="base"/></group></tab></tabs></ribbon></customUI>`)
	if _, err := MergeRibbonXML(base, duplicate); err == nil || !strings.Contains(err.Error(), `id "base" collides`) {
		t.Fatalf("control collision was not rejected: %v", err)
	}
	if !bytes.Equal(base, []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="base" label="Old"/></tabs></ribbon></customUI>`)) {
		t.Fatal("failed merge changed the base input")
	}
}

func TestMergeRibbonXMLHonorsSchemaOrderAndSelfClosingParents(t *testing.T) {
	base := []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon/></customUI>`)
	fragment := []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><commands><command idMso="FileSave"/></commands><ribbon><tabs><tab id="tab"><group id="group"><button id="button"/></group></tab></tabs></ribbon></customUI>`)
	merged, err := MergeRibbonXML(base, fragment)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Index(string(merged), "<commands>") > strings.Index(string(merged), "<ribbon>") {
		t.Fatalf("Ribbon sequence order was not preserved: %s", merged)
	}
	if !strings.Contains(string(merged), `<ribbon><tabs>`) || !strings.Contains(string(merged), `id="button"`) {
		t.Fatalf("self-closing parent was not expanded: %s", merged)
	}
}

func TestPackageMergeRibbonConnectsAndIsAtomic(t *testing.T) {
	p := BlankPackage()
	first := []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="one"/></tabs></ribbon></customUI>`)
	if err := p.MergeRibbon("customUI/customUI14.xml", first); err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(p.Files["_rels/.rels"], []byte(`Target="customUI/customUI14.xml"`)) {
		t.Fatal("merged Ribbon was not connected")
	}
	before := append([]byte(nil), p.Files["customUI/customUI14.xml"]...)
	bad := []byte(`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="one" label="different"/></tabs></ribbon></customUI>`)
	if err := p.MergeRibbon("customUI/customUI14.xml", bad); err == nil {
		t.Fatal("conflicting package Ribbon was accepted")
	}
	if !bytes.Equal(before, p.Files["customUI/customUI14.xml"]) {
		t.Fatal("failed package merge changed the package")
	}
}
