//go:build windows && (amd64 || arm64)

package native

import "testing"

func TestRibbonSchemas(t *testing.T) {
	for _, ns := range []string{"http://schemas.microsoft.com/office/2006/01/customui", "http://schemas.microsoft.com/office/2009/07/customui"} {
		for _, test := range []struct {
			body  string
			valid bool
		}{
			{`<ribbon><tabs><tab id="test" label="Test"><group id="group" label="Tools"><button id="run" label="Run" onAction="Run"/></group></tab></tabs></ribbon>`, true},
			{`<ribbon><tabs><button id="wrongParent"/></tabs></ribbon>`, false},
			{`<ribbon nonsense="true"/>`, false},
		} {
			r, e := ValidateRibbon([]byte(`<customUI xmlns="` + ns + `">` + test.body + `</customUI>`))
			if e != nil {
				t.Fatal(e)
			}
			if r["valid"] != test.valid {
				t.Fatalf("expected valid=%v: %v", test.valid, r)
			}
		}
	}
}
