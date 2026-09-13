package office

import "testing"

func TestControlTextAlignmentPersistsWithFont(t *testing.T) {
	for _, kind := range []string{"Label", "TextBox", "CommandButton"} {
		for align, stored := range map[int]int64{1: 1, 2: 3, 3: 2} {
			r, err := defaultControl(kind, "Sample")
			if err != nil {
				t.Fatal(err)
			}
			if err = applyRecord(r, map[string]any{"TextAlign": align, "Font": map[string]any{"Name": "Segoe UI", "Size": 12}}, "Size"); err != nil {
				t.Fatal(err)
			}
			font, err := readRecord(r.tail, textSpec, 1252)
			if err != nil || font.values["ParagraphAlign"] != stored || font.strings["FontName"] != "Segoe UI" {
				t.Fatalf("%s alignment=%d: %v %v", kind, align, font, err)
			}
			if err = applyRecord(r, map[string]any{"TextAlign": 4}, "Size"); err == nil {
				t.Fatal("accepted invalid alignment")
			}
		}
	}
}
