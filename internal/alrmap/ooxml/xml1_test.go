package ooxml_test

import (
	"bytes"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap/alrtest"
	"github.com/eliziff/WordUp/internal/alrmap/ooxml"
)

func lexicalOut(t *testing.T, sc *ooxml.Scan) ([]byte, []ooxml.ByteWrite) {
	t.Helper()
	out, writes, err := alrtest.Lexical(t, sc)
	if err != nil {
		t.Fatalf("lexical: %v", err)
	}
	return out, writes
}

func mustText(t *testing.T, got, want string) {
	t.Helper()
	if got != want {
		t.Fatalf("got %q, want %q", got, want)
	}
}

func unchanged(t *testing.T, sc *ooxml.Scan) {
	t.Helper()
	out, _ := lexicalOut(t, sc)
	if !bytes.Equal(out, sc.Raw) {
		t.Fatalf("output changed:\n%s", out)
	}
}

func TestXmlProtection1(t *testing.T) {
	t.Run("test_split_run_exact_address", func(t *testing.T) {
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run("A towar", "", true)+alrtest.Run("ds result.", "", true), ""), nil)
		output, writes := lexicalOut(t, sc)
		mustText(t, alrtest.Text(t, output, sc.Styles, sc.Part, 0), "A toward result.\r")
		if len(writes) != 1 {
			t.Fatalf("writes = %d", len(writes))
		}
		if !bytes.Contains(output, []byte(`<w:t xml:space="preserve">d result.</w:t>`)) {
			t.Fatalf("output %s", output)
		}
	})
	t.Run("test_bytes_outside_owned_text_values", func(t *testing.T) {
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run("A judgement & towards 25%.", "", true), ""), nil)
		output, writes := lexicalOut(t, sc)
		oldCursor, newCursor := 0, 0
		for _, patch := range writes {
			unchangedSlice := sc.Raw[oldCursor:patch.Start]
			if !bytes.Equal(output[newCursor:newCursor+len(unchangedSlice)], unchangedSlice) {
				t.Fatal("unchanged slice differs")
			}
			newCursor += len(unchangedSlice)
			if !bytes.Equal(output[newCursor:newCursor+len(patch.After)], patch.After) {
				t.Fatal("patch bytes differ")
			}
			newCursor += len(patch.After)
			oldCursor = patch.End
		}
		if !bytes.Equal(output[newCursor:], sc.Raw[oldCursor:]) {
			t.Fatal("tail differs")
		}
	})
	t.Run("test_quotes_cross_paragraphs", func(t *testing.T) {
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run("We move towards “a judgement", "", true), "")+alrtest.Para(alrtest.Run("towards 20%”.", "", true), ""), nil)
		output, _ := lexicalOut(t, sc)
		mustText(t, alrtest.Text(t, output, sc.Styles, sc.Part, 0), "We move toward “a judgement\rtowards 20%”.\r")
	})
	t.Run("test_unclosed_quote_quarantines_body", func(t *testing.T) {
		unchanged(t, alrtest.Project(t, alrtest.Para(alrtest.Run("“towards", "", true), "")+alrtest.Para(alrtest.Run("judgement", "", true), ""), nil))
	})
	t.Run("test_inherited_italic_is_protected", func(t *testing.T) {
		st := alrtest.Styles(t, `<w:style w:type="character" w:styleId="Book"><w:basedOn w:val="Default"/><w:rPr><w:i/></w:rPr></w:style>`, "")
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run("towards ", "", true)+alrtest.Run("judgement", `<w:rStyle w:val="Book"/>`, true), ""), st)
		output, _ := lexicalOut(t, sc)
		mustText(t, alrtest.Text(t, output, st, sc.Part, 0), "toward judgement\r")
	})
	t.Run("test_style_toggle_false_does_not_erase_parent", func(t *testing.T) {
		st := alrtest.Styles(t, `<w:style w:type="character" w:styleId="A"><w:rPr><w:i/></w:rPr></w:style>`+
			`<w:style w:type="character" w:styleId="B"><w:basedOn w:val="A"/><w:rPr><w:i w:val="false"/></w:rPr></w:style>`, "")
		unchanged(t, alrtest.Project(t, alrtest.Para(alrtest.Run("towards", `<w:rStyle w:val="B"/>`, true), ""), st))
	})
	t.Run("test_true_style_toggle_and_direct_false", func(t *testing.T) {
		st := alrtest.Styles(t, `<w:style w:type="character" w:styleId="A"><w:rPr><w:i/></w:rPr></w:style>`+
			`<w:style w:type="character" w:styleId="B"><w:basedOn w:val="A"/><w:rPr><w:i/></w:rPr></w:style>`, "")
		for _, props := range []string{`<w:rStyle w:val="B"/>`, `<w:rStyle w:val="A"/><w:i w:val="0"/>`} {
			t.Run(props, func(t *testing.T) {
				sc := alrtest.Project(t, alrtest.Para(alrtest.Run("towards", props, true), ""), st)
				output, _ := lexicalOut(t, sc)
				mustText(t, alrtest.Text(t, output, st, sc.Part, 0), "toward\r")
			})
		}
	})
	t.Run("test_web_hidden_is_not_a_toggle", func(t *testing.T) {
		st := alrtest.Styles(t, `<w:style w:type="character" w:styleId="A"><w:rPr><w:webHidden/></w:rPr></w:style>`+
			`<w:style w:type="character" w:styleId="B"><w:basedOn w:val="A"/><w:rPr><w:webHidden/></w:rPr></w:style>`, "")
		unchanged(t, alrtest.Project(t, alrtest.Para(alrtest.Run("towards", `<w:rStyle w:val="B"/>`, true), ""), st))
	})
	t.Run("test_default_italic_and_language", func(t *testing.T) {
		for _, props := range []string{`<w:i/>`, `<w:lang w:val="fr-CA"/>`, `<w:vanish/>`} {
			st := alrtest.Styles(t, "", `<w:docDefaults><w:rPrDefault><w:rPr>`+props+`</w:rPr></w:rPrDefault></w:docDefaults>`)
			unchanged(t, alrtest.Project(t, alrtest.Para(alrtest.Run("towards", "", true), ""), st))
		}
	})
	t.Run("test_cyclic_and_missing_styles_quarantine", func(t *testing.T) {
		for _, definition := range []string{"", `<w:style w:type="character" w:styleId="A"><w:basedOn w:val="A"/></w:style>`} {
			st := alrtest.Styles(t, definition, "")
			unchanged(t, alrtest.Project(t, alrtest.Para(alrtest.Run("towards", `<w:rStyle w:val="A"/>`, true), ""), st))
		}
	})
	t.Run("test_inherited_indentation", func(t *testing.T) {
		st := alrtest.Styles(t, `<w:style w:type="paragraph" w:styleId="Indented"><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="288"/></w:pPr></w:style>`, "")
		unchanged(t, alrtest.Project(t, alrtest.Para(alrtest.Run("towards", "", true), `<w:pStyle w:val="Indented"/>`), st))
	})
	t.Run("test_direct_indent_override", func(t *testing.T) {
		st := alrtest.Styles(t, `<w:style w:type="paragraph" w:styleId="Indented"><w:pPr><w:ind w:left="288"/></w:pPr></w:style>`, "")
		sc := alrtest.Project(t, alrtest.Para(alrtest.Run("towards", "", true), `<w:pStyle w:val="Indented"/><w:ind w:left="0"/>`), st)
		out, _ := lexicalOut(t, sc)
		if bytes.Equal(out, sc.Raw) {
			t.Fatal("direct indent override did not permit the edit")
		}
	})
	t.Run("test_structural_wrappers_protect_descendants", func(t *testing.T) {
		for _, tag := range []string{"w:hyperlink", "w:sdt", "w:ins", "w:del", "w:moveFrom", "w:fldSimple", "m:oMath", "w:customXml", "w:drawing", "x:unknown"} {
			t.Run(tag, func(t *testing.T) {
				sc := alrtest.Project(t, alrtest.Para("<"+tag+">"+alrtest.Run("towards", "", true)+"</"+tag+">"+alrtest.Run(" judgement", "", true), ""), nil)
				output, _ := lexicalOut(t, sc)
				mustText(t, alrtest.Text(t, output, sc.Styles, sc.Part, 0), "towards judgment\r")
			})
		}
	})
	t.Run("test_tables_protect_contents", func(t *testing.T) {
		table := "<w:tbl><w:tr><w:tc>" + alrtest.Para(alrtest.Run("towards 25%", "", true), "") + "</w:tc></w:tr></w:tbl>"
		sc := alrtest.Project(t, table+alrtest.Para(alrtest.Run("A judgement.", "", true), ""), nil)
		output, _ := lexicalOut(t, sc)
		mustText(t, alrtest.Text(t, output, sc.Styles, sc.Part, 0), "towards 25%\rA judgment.\r")
	})
	t.Run("test_complex_fields_across_paragraphs", func(t *testing.T) {
		sc := alrtest.Project(t, alrtest.Para(`<w:r><w:fldChar w:fldCharType="begin"/><w:instrText>REF a</w:instrText>`+
			`<w:fldChar w:fldCharType="separate"/></w:r>`+alrtest.Run("towards", "", true), "")+
			alrtest.Para(alrtest.Run("judgement", "", true)+`<w:r><w:fldChar w:fldCharType="end"/></w:r>`+alrtest.Run(" towards", "", true), ""), nil)
		output, _ := lexicalOut(t, sc)
		mustText(t, alrtest.Text(t, output, sc.Styles, sc.Part, 0), "towards\rjudgement toward\r")
	})
}
