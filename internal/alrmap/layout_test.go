package alrmap_test

import (
	"reflect"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap"
	"github.com/eliziff/WordUp/internal/alrmap/alrtest"
)

func TestLayoutInputs(t *testing.T) {
	t.Run("test_ambiguous_unit_is_not_a_number", func(t *testing.T) {
		if _, err := alrmap.Twips("1.5", ""); err == nil {
			t.Fatal("unitless value accepted")
		}
		for _, c := range []struct {
			value, unit string
			want        int
		}{{"0.2", "in", 288}, {"0.51", "cm", 289}, {"12", "pt", 240}} {
			got, err := alrmap.Twips(c.value, c.unit)
			if err != nil || got != c.want {
				t.Fatalf("Twips(%q,%q) = %d, %v; want %d", c.value, c.unit, got, err, c.want)
			}
		}
	})
	t.Run("test_shared_normal_cannot_change_protected_quote", func(t *testing.T) {
		st := alrtest.Styles(t, `<w:style w:type="paragraph" w:styleId="Quote"><w:basedOn w:val="Normal"/></w:style>`, "")
		parts := map[string][]byte{"word/document.xml": alrtest.Document(alrtest.Para(alrtest.Run("Author text", "", true), "") +
			alrtest.Para(alrtest.Run("Quoted text", "", true), `<w:pStyle w:val="Quote"/>`))}
		uses, err := alrmap.StyleUses(parts, st)
		if err != nil {
			t.Fatal(err)
		}
		owned := map[alrmap.ElementRef]bool{}
		for _, u := range uses {
			if u.Style == "Normal" {
				owned[alrmap.ElementRef{Part: u.Part, ElementIndex: u.ElementIndex}] = true
			}
		}
		if _, err := alrmap.RequireStyleOwnership(parts, st, "Normal", owned); err == nil {
			t.Fatal("shared style change accepted")
		}
		allOwned := map[alrmap.ElementRef]bool{}
		for _, u := range uses {
			allOwned[alrmap.ElementRef{Part: u.Part, ElementIndex: u.ElementIndex}] = true
		}
		hits, err := alrmap.RequireStyleOwnership(parts, st, "Normal", allOwned)
		if err != nil || len(hits) != 2 {
			t.Fatalf("hits = %v, %v", hits, err)
		}
	})
	t.Run("test_linked_style_partner_is_in_effect_closure", func(t *testing.T) {
		st := alrtest.Styles(t, `<w:style w:type="paragraph" w:styleId="Head"><w:link w:val="HeadChar"/></w:style>`+
			`<w:style w:type="character" w:styleId="HeadChar"><w:link w:val="Head"/></w:style>`, "")
		got, err := alrmap.AffectedStyles(st, "Head")
		if err != nil || !reflect.DeepEqual(got, map[string]bool{"Head": true, "HeadChar": true}) {
			t.Fatalf("got %v, %v", got, err)
		}
	})
	t.Run("test_heading_ids_survive_different_intro_policy", func(t *testing.T) {
		hs := []alrmap.Heading{{ID: "intro", Level: 1, Introduction: true}, {ID: "a", Level: 1}, {ID: "b", Level: 2}, {ID: "c", Level: 3}, {ID: "d", Level: 1}}
		got, err := alrmap.HeadingLabels(hs, true)
		if err != nil || !reflect.DeepEqual(got, map[string]string{"intro": "I", "a": "II", "b": "II.A", "c": "II.A.1", "d": "III"}) {
			t.Fatalf("got %v, %v", got, err)
		}
		got, err = alrmap.HeadingLabels(hs, false)
		if err != nil || !reflect.DeepEqual(got, map[string]string{"intro": "", "a": "I", "b": "I.A", "c": "I.A.1", "d": "II"}) {
			t.Fatalf("got %v, %v", got, err)
		}
		if _, err := alrmap.HeadingLabels([]alrmap.Heading{{ID: "bad", Level: 2}}, true); err == nil {
			t.Fatal("missing parent accepted")
		}
	})
}
