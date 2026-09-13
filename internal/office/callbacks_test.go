package office

import (
	"strings"
	"testing"
)

func TestRibbonCallbackSignaturesAndConflicts(t *testing.T) {
	prefix := `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Load">`
	r, e := RibbonCallbacks([]byte(prefix + `<toggleButton onAction="Toggle"/><gallery onAction="Pick" getItemLabel="Label"/></customUI>`))
	if e != nil {
		t.Fatal(e)
	}
	s := r["source"].(string)
	for _, want := range []string{"Load(ribbon As IRibbonUI)", "pressed As Boolean", "id As String, index As Integer", "index As Integer, ByRef returnedVal"} {
		if !strings.Contains(s, want) {
			t.Fatal(s)
		}
	}
	if _, e = RibbonCallbacks([]byte(prefix + `<button onAction="Shared"/><toggleButton onAction="shared"/></customUI>`)); e == nil {
		t.Fatal("conflicting callback silently reused")
	}
}
