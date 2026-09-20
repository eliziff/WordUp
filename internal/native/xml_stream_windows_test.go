//go:build windows && (amd64 || arm64)

package native

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
	"unsafe"

	"github.com/eliziff/WordUp/internal/office"
)

func TestXMLFileBeyondResponseLimit(t *testing.T) {
	// Place a surrogate pair directly across the decoder's chunk boundary.
	prefix := `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/test.xml"><pkg:xmlData><text>`
	source := prefix + strings.Repeat("x", 32767-len(prefix)) + "😀" + strings.Repeat("y", 16<<20) + `</text></pkg:xmlData></pkg:part></pkg:package>`
	p, err := bstr(source)
	if err != nil {
		t.Fatal(err)
	}
	defer sysFree.Call(p)
	if _, err := bstrText(p); err == nil {
		t.Fatal("inline response limit bypassed")
	}
	path := filepath.Join(t.TempDir(), "snapshot.xml")
	n, _, _ := sysLen.Call(p)
	result, err := exportXMLFile(&xmlBSTRReader{remaining: unsafe.Slice((*uint16)(unsafe.Pointer(p)), int(n))}, path)
	if err != nil {
		t.Fatal(err)
	}
	raw, err := os.ReadFile(path)
	if err != nil || string(raw) != source {
		t.Fatal("file changed XML bytes", err)
	}
	report := result.(map[string]any)
	if report["sha256"] != office.Hash(raw) || report["bytes"] != int64(len(raw)) || len(report["parts"].([]string)) != 1 || report["parts"].([]string)[0] != "/test.xml" {
		t.Fatal(report)
	}
	if _, err := exportXMLFile(strings.NewReader("<broken>"), path); err == nil {
		t.Fatal("malformed XML accepted")
	}
	after, err := os.ReadFile(path)
	if err != nil || string(after) != source {
		t.Fatal("failed export overwrote previous evidence", err)
	}
}
