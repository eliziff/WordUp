//go:build windows && (amd64 || arm64)

package native

import "testing"

func TestAutomationFailureDistinguishesBusyFromMissingMember(t *testing.T) {
	for _, tc := range []struct {
		hr   uintptr
		code string
	}{
		{0x80010001, "word_busy"},
		{0x8001010A, "word_busy"},
		{0x80020006, "unknown_member"},
		{0x80020003, "unknown_member"},
		{0x80020009, "word_automation_error"},
		{0x80070005, "word_automation_error"},
	} {
		f := automationFailure("Run", "failed", "VBA", "", tc.hr, exceptInfo{Scode: -1}, 0)
		d := f.Details.(map[string]any)
		if f.Code != tc.code || d["member"] != "Run" || d["source"] != "VBA" || d["word_error"] != int32(-1) {
			t.Fatalf("HRESULT %08X: %#v", tc.hr, f)
		}
		if tc.code == "word_busy" && d["hint"] == nil {
			t.Fatal("busy calls need a diagnostic next step")
		}
	}
}

func TestAutomationFailureReportsOnlyValidArgumentIndex(t *testing.T) {
	for _, hr := range []uintptr{0x80020009, 0x80020005, 0x80020004} {
		f := automationFailure("Run", "failed", "Macro", "macro.chm", hr, exceptInfo{Code: 513, HelpContext: 7}, 2)
		d := f.Details.(map[string]any)
		_, indexed := d["argument_index_reversed"]
		if indexed != (hr != 0x80020009) || d["exception_code"] != uint16(513) || d["help_context"] != uint32(7) || d["help_file"] != "macro.chm" {
			t.Fatal(d)
		}
	}
}
