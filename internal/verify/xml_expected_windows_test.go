//go:build windows && (amd64 || arm64)

package verify_test

import (
	"context"
	"encoding/xml"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestNativeExpectedXML(t *testing.T) {
	if os.Getenv("WORDUP_NATIVE_TEST") != "1" {
		t.Skip("set WORDUP_NATIVE_TEST=1")
	}
	ctx := context.Background()
	h, err := native.Start(ctx, native.Options{Execute: true})
	if err != nil {
		t.Fatal(err)
	}
	defer h.Close()
	call := func(op native.Operation) any {
		t.Helper()
		value, err := h.Call(ctx, op)
		if err != nil {
			t.Fatal(err)
		}
		return value
	}
	call(native.Operation{Op: "new", As: "doc"})
	call(native.Operation{Op: "get", Target: "doc", Member: "Content", As: "content"})
	call(native.Operation{Op: "put", Target: "content", Member: "Text", Value: "Before"})
	call(native.Operation{Op: "put", Target: "content", Member: "Italic", Value: true})
	// Word removes the run's formatting-history rsidRPr on text assignment.
	// Establish a repeated-edit baseline using that same operation; do not
	// normalize either XML file or exempt history attributes from comparison.
	call(native.Operation{Op: "put", Target: "content", Member: "Text", Value: "Before"})
	snapshot := func() []byte {
		return []byte(call(native.Operation{Op: "xml.snapshot", Target: "doc"}).(map[string]any)["xml"].(string))
	}
	// Start from a saved document, as a real input fixture would. A brand-new
	// unsaved Word document lazily acquires package metadata on first export.
	call(native.Operation{Op: "invoke", Target: "doc", Member: "SaveAs2", Args: []any{filepath.Join(t.TempDir(), "input.docx")}})
	source := snapshot()
	// The expected XML is a byte-preserving copy, with just the requested text
	// replacement. No reconstruction of Word's styles, runs or package parts.
	if strings.Count(string(source), ">Before<") != 1 {
		t.Fatal("source text not uniquely anchored")
	}
	expected := []byte(strings.Replace(string(source), ">Before<", ">After!<", 1))
	if _, err := office.VerifyXML(expected, source, ""); err == nil {
		t.Fatal("no-op passed")
	}
	call(native.Operation{Op: "put", Target: "content", Member: "Text", Value: "After!"})
	actual := snapshot()
	// Editing-history identifiers can change independently of document content.
	// This fixture has no comments or other paragraph-ID consumers. Keep this
	// policy local, explicit and visible in the result, never a global default.
	policy := office.XMLComparePolicy{Elements: []xml.Name{{Space: "http://schemas.microsoft.com/office/word/2010/wordml", Local: "docId"}}, Attributes: []xml.Name{
		{Space: "http://schemas.microsoft.com/office/word/2010/wordml", Local: "textId"},
		{Space: "http://schemas.microsoft.com/office/word/2010/wordml", Local: "paraId"},
		{Space: office.W, Local: "rsidRPr"},
		{Space: office.W, Local: "rsidRDefault"},
	}}
	noOp, err := office.CompareXML(expected, source, policy)
	if err != nil || noOp["equal"] == true {
		t.Fatalf("history policy accepted no-op: %v", err)
	}
	report, err := office.CompareXML(expected, actual, policy)
	if err != nil || report["equal"] != true {
		t.Fatalf("Word output differs beyond history: %v; difference=%v", err, report["first_difference"])
	}
	if raw, err := office.VerifyXML(expected, actual, ""); err != nil && raw["matches_expected"] != false {
		t.Fatal("exact mismatch was hidden")
	}
	call(native.Operation{Op: "put", Target: "content", Member: "Italic", Value: false})
	broken, err := office.CompareXML(expected, snapshot(), policy)
	if err != nil || broken["equal"] == true {
		t.Fatalf("lost italics passed or comparison failed: %v", err)
	}
	t.Log("Word matches copied-and-edited XML with explicit fixture-local history policy; lost italics rejected under the same policy")
}
