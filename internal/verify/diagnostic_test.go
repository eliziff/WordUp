package verify

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"unicode/utf8"

	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
)

func TestXPathObservationWithoutNativeHost(t *testing.T) {
	path := filepath.Join(t.TempDir(), "evidence.xml")
	if err := os.WriteFile(path, []byte(`<w:p xmlns:w="urn:word"><w:spacing w:after="0"/></w:p>`), 0600); err != nil {
		t.Fatal(err)
	}
	op := native.Operation{Op: "xml.query", File: path, Member: "count(//w:spacing[@w:after='0'])", Named: map[string]any{"namespaces": map[string]string{"w": "urn:word"}}}
	value, err := observedCall(context.Background(), nil, op)
	if err != nil {
		t.Fatal(err)
	}
	if err := Check(value, Assertion{Path: "/value", Kind: "equals", Expected: 1}); err != nil {
		t.Fatal(err)
	}
	op.Named["typo"] = true
	if _, err := observedCall(context.Background(), nil, op); err == nil {
		t.Fatal("silently accepted unknown XPath options")
	}
}

func TestExpectedXMLObservationWithoutNativeHost(t *testing.T) {
	root := t.TempDir()
	expected := filepath.Join(root, "expected.xml")
	actual := filepath.Join(root, "actual.xml")
	for _, path := range []string{expected, actual} {
		if err := os.WriteFile(path, []byte(`<p><i>kept</i></p>`), 0600); err != nil {
			t.Fatal(err)
		}
	}
	op := native.Operation{Op: "xml.verify", Target: expected, File: actual}
	if _, err := observedCall(context.Background(), nil, op); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(actual, []byte(`<p>kept</p>`), 0600); err != nil {
		t.Fatal(err)
	}
	value, err := observedCall(context.Background(), nil, op)
	if err == nil || value.(map[string]any)["matches_expected"] != false {
		t.Fatalf("lost formatting passed: %v %v", value, err)
	}
	op.Named = map[string]any{"ignore": "i"}
	if _, err := observedCall(context.Background(), nil, op); err == nil {
		t.Fatal("unknown verification option accepted")
	}
}

func TestLargeAssertionDiagnostics(t *testing.T) {
	value := strings.Repeat("é中<xml>\n", 100000)
	for _, kind := range []string{"equals", "absent"} {
		a := Assertion{Path: "/xml", Kind: kind, Expected: "unchanged"}
		input := map[string]any{"xml": value}
		err := Check(input, a)
		if err == nil || len(err.Error()) > 3000 || !utf8.ValidString(err.Error()) {
			t.Fatalf("unbounded or invalid diagnostic for %s", kind)
		}
		fault := assertionFailure(input, a, 0, err).(*native.Fault)
		details := fault.Details.(map[string]any)
		summary := details["actual"].(map[string]any)
		data, _ := json.Marshal(value)
		if summary["json_sha256"] != office.Hash(data) || summary["json_bytes"] != len(data) || summary["truncated"] != true {
			t.Fatal("lost evidence identity")
		}
		if input["xml"] != value {
			t.Fatal("modified full observation")
		}
	}
	if diagnosticValue(nil) != nil || diagnosticValue(42) != 42 || diagnosticValue("short") != "short" {
		t.Fatal("small values must retain their original types")
	}
}
