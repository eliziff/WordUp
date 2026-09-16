package grade

import (
	"os"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap/ooxml"
	"github.com/eliziff/WordUp/internal/office"
)

func TestDebugHardSpans(t *testing.T) {
	path := os.Getenv("WORDUP_GRADE_DEBUG")
	if path == "" {
		t.Skip("set WORDUP_GRADE_DEBUG to a Flat OPC file")
	}
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	stylesRaw, _ := office.XMLInput(data, "word/styles.xml")
	styles, err := ooxml.NewStyles(stylesRaw)
	if err != nil {
		t.Fatal(err)
	}
	for _, part := range []string{"word/document.xml", "word/footnotes.xml"} {
		raw, err := office.XMLInput(data, part)
		if err != nil {
			t.Logf("%s: %v", part, err)
			continue
		}
		scan, err := ooxml.ScanPart(raw, styles, part)
		if err != nil {
			t.Fatalf("%s: %v", part, err)
		}
		for _, c := range scan.Containers {
			reasons := map[string]int{}
			covered := 0
			for _, h := range c.Hard {
				reasons[h.Reason]++
				covered += h.End - h.Start
			}
			if len(c.Hard) > 0 || c.Key == "body" {
				t.Logf("%s %s: text %d runes, paragraphs %d, hard spans %d covering %d, reasons %v", part, c.Key, len([]rune(c.Text)), len(c.Paragraphs), len(c.Hard), covered, reasons)
			}
			if c.Key == "body" {
				for i, h := range c.Hard {
					if i < 8 {
						txt := []rune(c.Text)
						lo, hi := h.Start, h.End
						if hi > lo+60 {
							hi = lo + 60
						}
						t.Logf("   hard[%d] %d-%d %s: %q", i, h.Start, h.End, h.Reason, strings.ReplaceAll(string(txt[lo:hi]), "", "|"))
					}
				}
			}
		}
	}
}
