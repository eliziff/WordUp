package alrmap_test

import (
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/alrmap"
	"github.com/eliziff/WordUp/internal/alrmap/alrtest"
	"github.com/eliziff/WordUp/internal/alrmap/citations"
	"github.com/eliziff/WordUp/internal/alrmap/ooxml"
)

// note mirrors ReferenceBindings.note (scope defaults to "article").
func note(t *testing.T, id, label, text string, scope ...string) alrmap.Note {
	t.Helper()
	result, err := citations.NormalizeNote(text, alrtest.Reg, citations.Options{})
	if err != nil {
		t.Fatalf("NormalizeNote(%q): %v", text, err)
	}
	s := "article"
	if len(scope) > 0 {
		s = scope[0]
	}
	return alrmap.Note{ID: id, Label: label, Scope: s, Citations: result.Citations}
}

func graph(t *testing.T, notes ...alrmap.Note) *alrmap.Graph {
	t.Helper()
	g, err := alrmap.CompileGraph(notes)
	if err != nil {
		t.Fatal(err)
	}
	return g
}

func pointer(t *testing.T, g *alrmap.Graph, scope, label string) string {
	t.Helper()
	id, err := g.PreservePointer(scope, label)
	if err != nil {
		t.Fatal(err)
	}
	return id
}

func TestReferenceBindings(t *testing.T) {
	t.Run("test_native_id_is_not_printed_label", func(t *testing.T) {
		g := graph(t, note(t, "xml:105", "*", "2024 SCC 1."), note(t, "xml:999", "1", "2024 SCC 7 [Example]."))
		if pointer(t, g, "article", "1") != "xml:999" {
			t.Fatal("wrong pointer")
		}
	})
	t.Run("test_restarted_labels_require_scope", func(t *testing.T) {
		g := graph(t, note(t, "a", "1", "2024 SCC 1.", "partA"), note(t, "b", "1", "2024 SCC 2.", "partB"))
		if pointer(t, g, "partA", "1") != "a" || pointer(t, g, "partB", "1") != "b" {
			t.Fatal("scoped pointers wrong")
		}
		if _, err := g.PreservePointer("unknown", "1"); err == nil {
			t.Fatal("unknown scope resolved")
		}
	})
	t.Run("test_duplicate_display_labels_within_scope_refused", func(t *testing.T) {
		g := graph(t, note(t, "a", "1", "2024 SCC 1."), note(t, "b", "1", "2024 SCC 2."))
		if _, err := g.PreservePointer("article", "1"); err == nil {
			t.Fatal("duplicate label resolved")
		}
	})
	t.Run("test_supra_resolves_source_not_old_pinpoint", func(t *testing.T) {
		g := graph(t, note(t, "a", "1", "2024 SCC 7 at para 9 [Example]."), note(t, "b", "2", "Example, supra note 1."))
		b := g.Bindings[alrmap.BindingKey{NoteID: "b", Occurrence: 0}]
		if b.Locator != nil {
			t.Fatalf("supra inherited a pinpoint: %v", b.Locator)
		}
		if citations.SourceKey(b.Source) != citations.SourceKey(g.Bindings[alrmap.BindingKey{NoteID: "a"}].Source) {
			t.Fatal("supra source differs")
		}
	})
	t.Run("test_ibid_inherits_old_pinpoint", func(t *testing.T) {
		g := graph(t, note(t, "a", "1", "2024 SCC 7 at para 9 [Example]."), note(t, "b", "2", "Ibid."))
		if g.Bindings[alrmap.BindingKey{NoteID: "a"}].Meaning() != g.Bindings[alrmap.BindingKey{NoteID: "b"}].Meaning() {
			t.Fatal("ibid meaning differs")
		}
	})
	t.Run("test_correction_separate_from_preservation", func(t *testing.T) {
		g := graph(t, note(t, "a", "1", "2024 SCC 7 [Example]."), note(t, "b", "2", "2024 SCC 8."), note(t, "c", "3", "Example, supra note 2."))
		if pointer(t, g, "article", "2") != "b" {
			t.Fatal("pointer not preserved")
		}
		if _, ok := g.Unresolved[alrmap.BindingKey{NoteID: "c"}]; !ok {
			t.Fatal("c should be unresolved")
		}
		if id, err := g.CorrectAliasTarget("c", "Example"); err != nil || id != "a" {
			t.Fatalf("got %q, %v", id, err)
		}
	})
	t.Run("test_colliding_alias_does_not_select_first", func(t *testing.T) {
		g := graph(t, note(t, "a", "1", "2024 SCC 7 [Example]."), note(t, "b", "2", "2024 SCC 8 [Example]."), note(t, "c", "3", "Example, supra note 1."))
		if _, err := g.CorrectAliasTarget("c", "Example"); err == nil {
			t.Fatal("colliding alias resolved")
		}
		if _, ok := g.Unresolved[alrmap.BindingKey{NoteID: "c"}]; !ok {
			t.Fatal("c should be unresolved")
		}
	})
	t.Run("test_ibid_after_multi_source_note_unresolved", func(t *testing.T) {
		g := graph(t, note(t, "a", "1", "2024 SCC 7; 2024 SCC 8."), note(t, "b", "2", "Ibid."))
		if _, ok := g.Unresolved[alrmap.BindingKey{NoteID: "b"}]; !ok {
			t.Fatal("b should be unresolved")
		}
	})
	t.Run("test_inserted_note_does_not_retarget_ibid", func(t *testing.T) {
		a, b := note(t, "a", "1", "2024 SCC 7 at para 9 [Example]."), note(t, "b", "2", "Ibid.")
		g := graph(t, a, b)
		x := note(t, "x", "2", "2024 SCC 8.")
		b3 := b
		b3.Label = "3"
		got, ok, err := alrmap.RepairIbidAfterReorder(g, []alrmap.Note{a, x, b3}, "b", g.Fingerprint)
		if err != nil || !ok || got != "Example, supra note 1 at para 9." {
			t.Fatalf("got %q, %v, %v", got, ok, err)
		}
	})
	t.Run("test_explicit_ibid_locator_does_not_need_repair", func(t *testing.T) {
		a, b := note(t, "a", "1", "2024 SCC 7 at para 9 [Example]."), note(t, "b", "2", "Ibid at para 12.")
		g := graph(t, a, b)
		x := note(t, "x", "2", "2024 SCC 7 at para 4.")
		b3 := b
		b3.Label = "3"
		got, ok, err := alrmap.RepairIbidAfterReorder(g, []alrmap.Note{a, x, b3}, "b", g.Fingerprint)
		if err != nil || ok {
			t.Fatalf("got %q, %v, %v", got, ok, err)
		}
	})
	t.Run("test_source_edition_identity_not_merged", func(t *testing.T) {
		a := note(t, "a", "1", "A Author, Towards Judgement, 1st ed (Town: Press, 2020) [Example].")
		b := note(t, "b", "2", "A Author, Towards Judgement, 2nd ed (Town: Press, 2024) [Example].")
		g := graph(t, a, b)
		if len(g.Aliases["Example"]) != 2 {
			t.Fatalf("aliases = %v", g.Aliases["Example"])
		}
	})
	t.Run("test_stale_reference_graph_refused", func(t *testing.T) {
		a, b := note(t, "a", "1", "2024 SCC 7 [Example]."), note(t, "b", "2", "Ibid.")
		g := graph(t, a, b)
		if _, _, err := alrmap.RepairIbidAfterReorder(g, []alrmap.Note{a, b}, "b", "stale"); err == nil {
			t.Fatal("stale fingerprint accepted")
		}
		changed := note(t, "a", "1", "2024 SCC 8 [Example].")
		if _, _, err := alrmap.RepairIbidAfterReorder(g, []alrmap.Note{changed, b}, "b", g.Fingerprint); err == nil {
			t.Fatal("changed citation accepted")
		}
	})
	t.Run("test_article_section_and_range_list_are_distinct", func(t *testing.T) {
		seen := map[string]bool{}
		for _, s := range []string{"Ibid, s 7.", "Ibid, art 7.", "Ibid, ss 7 to 9.", "Ibid, ss 7, 9."} {
			result, err := citations.NormalizeNote(s, nil, citations.Options{})
			if err != nil {
				t.Fatal(err)
			}
			seen[result.Citations[0].Locator.String()] = true
		}
		if len(seen) != 4 {
			t.Fatalf("locators = %v", seen)
		}
	})
	t.Run("test_legal_ibid_repair_retains_article_and_separator", func(t *testing.T) {
		a, b := note(t, "a", "1", "2024 SCC 7, arts 3-2 to 3-5 [Example]."), note(t, "b", "2", "Ibid.")
		g := graph(t, a, b)
		x := note(t, "x", "2", "2024 SCC 8.")
		b3 := b
		b3.Label = "3"
		got, ok, err := alrmap.RepairIbidAfterReorder(g, []alrmap.Note{a, x, b3}, "b", g.Fingerprint)
		if err != nil || !ok || got != "Example, supra note 1, arts 3-2 to 3-5." {
			t.Fatalf("got %q, %v, %v", got, ok, err)
		}
	})
	t.Run("test_whole_native_note_dependency_veto", func(t *testing.T) {
		raw := []byte(`<w:footnotes xmlns:w="` + ooxml.W + `"><w:footnote w:id="1">` +
			alrtest.Para(`<w:fldSimple w:instr="REF source">`+alrtest.Run("Ibid", "", true)+`</w:fldSimple>`+alrtest.Run(" at para. 5.", "", true), "") +
			`</w:footnote></w:footnotes>`)
		sc, err := ooxml.ScanPart(raw, alrtest.Styles(t, "", ""), "word/footnotes.xml")
		if err != nil {
			t.Fatal(err)
		}
		if _, _, err := ooxml.PlanNote(sc, "footnote:1", nil, nil); err == nil || !strings.Contains(err.Error(), "opaque") {
			t.Fatalf("got %v", err)
		}
	})
}
