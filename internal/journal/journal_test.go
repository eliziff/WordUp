package journal

import (
	"encoding/json"
	"math"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/eliziff/WordUp/internal/inspect"
	"github.com/eliziff/WordUp/internal/project"
)

func TestProfilesAreStableAndComplete(t *testing.T) {
	profiles := Profiles()
	if len(profiles) != 34 {
		t.Fatalf("profile count = %d, want 34", len(profiles))
	}
	seen := map[string]bool{}
	for _, profile := range profiles {
		if profile.ID == "" || seen[strings.ToLower(profile.ID)] {
			t.Fatalf("duplicate or empty profile id: %#v", profile)
		}
		seen[strings.ToLower(profile.ID)] = true
		if len(profile.Features) < 8 || profile.BodyFont == "" || profile.BodySizePT <= 0 {
			t.Fatalf("incomplete profile: %#v", profile)
		}
	}
}

func TestCatalogEvidenceOverlayPreservesBundledSurface(t *testing.T) {
	base, err := Get("UBC-L-REV")
	if err != nil {
		t.Fatal(err)
	}
	overlay := Profile{ID: base.ID, BodyFont: "Cambria", BodySizePT: 10.5, EvidenceStatus: "corpus-observed", Observed: ObservedEvidence{Articles: 2}}
	got := WithCatalogEvidence(base, Catalog{Journals: []Profile{overlay}})
	if got.Name != base.Name || len(got.Features) != len(base.Features) || got.BodySizePT != 10.5 || got.Observed.Articles != 2 || got.EvidenceStatus != "corpus-observed" {
		t.Fatalf("catalog overlay lost profile surface: got=%#v base=%#v", got, base)
	}
}

func TestProfileWithoutPermalinksOmitsPermaWiring(t *testing.T) {
	profile := Profile{ID: "NO-PERMA", Name: "No Permalink Journal", PermalinkPolicy: "none"}
	design := formDesign(profile)
	for _, control := range design.Controls {
		if control.Name == "cmdPerma" {
			t.Fatal("permalink-disabled profile still exposes a Perma form control")
		}
	}
	if source := formSource(profile); strings.Contains(source, "cmdPerma") || strings.Contains(source, "WU_JournalPermaAssistant") {
		t.Fatal("permalink-disabled form still contains Perma event wiring")
	}
	if ribbon := ribbonSource(profile, "NoPerma"); strings.Contains(ribbon, "_Perma") || strings.Contains(ribbon, "Perma assistant") {
		t.Fatal("permalink-disabled Ribbon still exposes Perma wiring")
	}
}

func TestProfileFeaturesControlGeneratedSurface(t *testing.T) {
	profile := Profile{
		ID:              "MINIMAL",
		Name:            "Minimal Journal",
		PermalinkPolicy: "assistant",
		Features:        []string{"journal setup", "preflight"},
	}
	design := formDesign(profile)
	for _, control := range design.Controls {
		if control.Name == "cmdStyles" || control.Name == "cmdFields" || control.Name == "cmdCitations" || control.Name == "cmdPerma" || control.Name == "cmdSupra" || control.Name == "cmdTracking" || control.Name == "cmdQuality" {
			t.Fatalf("feature-disabled form still exposes %s", control.Name)
		}
	}
	form := formSource(profile)
	for _, omitted := range []string{"cmdStyles", "cmdFields", "cmdCitations", "cmdPerma", "cmdSupra", "cmdTracking", "cmdQuality"} {
		if strings.Contains(form, omitted) {
			t.Fatalf("feature-disabled form still contains %s wiring", omitted)
		}
	}
	for _, retained := range []string{"cmdReview", "cmdCommands", "cmdRemoveCommands", "cmdPreflight", "cmdCancel"} {
		if !strings.Contains(form, retained) {
			t.Fatalf("feature-enabled form lost %s wiring", retained)
		}
	}
	ribbon := ribbonSource(profile, "Minimal")
	for _, omitted := range []string{"_Styles", "_Fields", "_Citations", "_Perma", "_Supra", "_Tracking", "_Quality"} {
		if strings.Contains(ribbon, omitted) {
			t.Fatalf("feature-disabled Ribbon still exposes %s", omitted)
		}
	}
	for _, retained := range []string{"_Setup", "_Review", "_Commands", "_RemoveCommands", "_Preflight"} {
		if !strings.Contains(ribbon, retained) {
			t.Fatalf("feature-enabled Ribbon lost %s", retained)
		}
	}
}

func TestCreateAllPreflightsDestinations(t *testing.T) {
	root := filepath.Join(t.TempDir(), "all")
	if err := os.MkdirAll(filepath.Join(root, "ALTA_L_REV"), 0700); err != nil {
		t.Fatal(err)
	}
	if _, err := CreateAll(root); err == nil || !strings.Contains(err.Error(), "ALTA-L-REV") {
		t.Fatalf("existing destination was not rejected before batch creation: %v", err)
	}
	entries, err := os.ReadDir(root)
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 1 || entries[0].Name() != "ALTA_L_REV" {
		t.Fatalf("batch created partial output after preflight failure: %#v", entries)
	}
}

func TestScanUsesCurrentEvidenceAndKeepsOutputCompact(t *testing.T) {
	root := filepath.Join(t.TempDir(), "final_contracts")
	oldArticle := filepath.Join(root, "UBC-L-REV", "58", "4", "article-old")
	if err := os.MkdirAll(oldArticle, 0700); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(oldArticle, "provenance.json"), []byte(`{"band_year":2000}`), 0600); err != nil {
		t.Fatal(err)
	}
	// The date prefix is sufficient to exclude an old record; malformed
	// telemetry after it must not make a current-year catalog fail.
	if err := os.WriteFile(filepath.Join(oldArticle, "digitalborn_native_summary.json"), []byte(`{"document_date_en":"2000",`), 0600); err != nil {
		t.Fatal(err)
	}
	article := filepath.Join(root, "UBC-L-REV", "59", "1", "article-1")
	if err := os.MkdirAll(article, 0700); err != nil {
		t.Fatal(err)
	}
	provenance := []byte(`{"schema_version":"test","heading_grammar_demotions":{"profile":{"band_year":2026}}}`)
	if err := os.WriteFile(filepath.Join(article, "provenance.json"), provenance, 0600); err != nil {
		t.Fatal(err)
	}
	summary := map[string]any{
		"document_date_en":         "2026-01-01",
		"font_role_body":           map[string]any{"active": true, "font": "Cambria", "size": 11},
		"font_role_note":           map[string]any{"active": true, "font": "Cambria", "size": 9},
		"font_role_counts":         map[string]any{"note": 3},
		"block_quote_line_count":   2,
		"list_item_line_count":     1,
		"running_furniture_counts": map[string]any{"header": 1},
		"drop_cap_merged_count":    1,
		"two_column_ladder_fired":  1,
		"toc_outline":              map[string]any{"entries": 2},
	}
	b, err := json.Marshal(summary)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(article, "digitalborn_native_summary.json"), b, 0600); err != nil {
		t.Fatal(err)
	}
	catalog, err := Scan(root, []int{2026})
	if err != nil {
		t.Fatal(err)
	}
	if catalog.JournalCount != 1 || catalog.ArticleCount != 1 {
		t.Fatalf("unexpected catalog counts: %#v", catalog)
	}
	profile := catalog.Journals[0]
	if profile.ID != "UBC-L-REV" || profile.CurrentYears[0] != 2026 || profile.StyleEvidence.BodyFont != "Cambria" {
		t.Fatalf("unexpected profile: %#v", profile)
	}
	if !profile.Observed.HasFootnotes || !profile.Observed.HasContents || !profile.Observed.HasTwoColumnPages {
		t.Fatalf("feature evidence was not retained: %#v", profile.Observed)
	}
	encoded, err := json.Marshal(catalog)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(encoded), "article-1") {
		t.Fatalf("catalog leaked an article path: %s", encoded)
	}
}

func TestCreateBuildsIndependentJournalWorkspace(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	profile, err := Get("UBC-L-REV")
	if err != nil {
		t.Fatal(err)
	}
	report, err := Create(root, profile)
	if err != nil {
		t.Fatal(err)
	}
	if report.Build == nil || !report.Build.PackageValidated || report.Artifact == "" {
		t.Fatalf("build report is incomplete: %#v", report)
	}
	if report.Build.Artifact != report.Artifact {
		t.Fatalf("build report retained staging artifact path: build=%q report=%q", report.Build.Artifact, report.Artifact)
	}
	if report.Build.Modules != 13 || report.Build.Forms != 1 {
		t.Fatalf("generated workspace did not vendor the expected source surface: modules=%d forms=%d", report.Build.Modules, report.Build.Forms)
	}
	if _, err := os.Stat(report.Artifact); err != nil {
		t.Fatal(err)
	}
	lock, err := os.ReadFile(filepath.Join(root, ".wordwright", "components.json"))
	if err != nil || !strings.Contains(string(lock), "operation.safe-edit") || !strings.Contains(string(lock), "command.hotkey") || !strings.Contains(string(lock), "structure.detect") || !strings.Contains(string(lock), "document.field-refresh") || !strings.Contains(string(lock), "document.text-operations") {
		t.Fatalf("generated workspace did not record reusable components: err=%v lock=%s", err, lock)
	}
	core, err := os.ReadFile(filepath.Join(root, "vba", "WUJournalCore.bas"))
	if err != nil {
		t.Fatal(err)
	}
	for _, marker := range []string{"WU_JournalOpenSetupFromRibbon", "WU_JournalApplyStyles", "WU_JournalPermaAssistant", "WU_DetectStructure", "WU_BeginSafeEdit", "Style definitions are shared document state", "If StrComp(CStr(.Font.Name), fontName, vbTextCompare) <> 0 Then .Font.Name = fontName", "Set value = doc.Styles(styleName)\n    Err.Clear\n    On Error GoTo 0", "Outline levels are shared style state too", "If heading1.ParagraphFormat.OutlineLevel <> wdOutlineLevel1 Then heading1.ParagraphFormat.OutlineLevel = wdOutlineLevel1", "StoryRanges already contains every header/footer story", "structureColumns < WU_COLUMNS", "startPosition <> priorEnd", "endPosition <= startPosition", "priorEnd <> storyEnd", "WU_ApplyResolvedParagraphStyles", "headingNames(1) = heading1.NameLocal", "role = vbNullString", "StrComp(role, \"body\", vbTextCompare)", "StrComp(styleName, desiredName", "paragraphRange.Style", "WU_JOURNAL_STYLE_H9", "wdOutlineLevel9", "paragraphRange As Range", "WU_FlushParagraphStyleBatch", "WU_ApplyNoteStoryChain", "WU_ApplyNoteStoryStyle", "note story boundaries are unavailable", "main story boundaries are unavailable", "WU_JOURNAL_MAX_STORY_CHAIN As Long = 32768", "story chain exceeds 32768 linked stories", "WU_RefreshFields(ActiveDocument, \"all\", True)", "could not refresh fields in \" & CStr(failures) & \" story/table(s)", "WU_CountStoryItems", "WU_JournalNextStory", "If nextStory Is story Then failed = True: Exit Function", "Unavailable story scans", "revision review cancelled", "Application.StatusBar = priorStatus", "is not a paragraph style", "WU_EndSafeEdit updating, undoStarted, captured", "Toggle tracked changes", "supraCount = WU_CountLiteral(doc, \"supra\", \"notes\", False, False)", "preserve every note's rich runs"} {
		if !strings.Contains(string(core), marker) {
			t.Fatalf("generated source lacks %s", marker)
		}
	}
	if strings.Contains(string(core), "noteText = note.Range.Text") {
		t.Fatal("citation audit materialized every note's text instead of using the shared bounded count")
	}
	if strings.Contains(string(core), "paragraphRange.Style =") {
		t.Fatal("generated style pass reverted to one COM style setter per paragraph")
	}
	w, err := project.Open(root)
	if err != nil {
		t.Fatal(err)
	}
	check, err := inspect.CheckWithConstants(w, nil)
	if err != nil {
		t.Fatal(err)
	}
	if diagnostics, ok := check["diagnostics"].([]map[string]any); !ok || len(diagnostics) != 0 {
		t.Fatalf("generated workspace diagnostics: %#v", check)
	}
}

func TestCreateAllRejectsSanitizedDestinationCollisions(t *testing.T) {
	root := filepath.Join(t.TempDir(), "all")
	profiles := []Profile{
		{ID: "A-B", Name: "First", BodyFont: "Times New Roman", NoteFont: "Times New Roman", BodySizePT: 11, NoteSizePT: 9},
		{ID: "A_B", Name: "Second", BodyFont: "Times New Roman", NoteFont: "Times New Roman", BodySizePT: 11, NoteSizePT: 9},
	}
	if _, err := createAll(root, profiles); err == nil || !strings.Contains(err.Error(), "collides") {
		t.Fatalf("sanitized destination collision was accepted: %v", err)
	}
	if _, err := os.Stat(root); !os.IsNotExist(err) {
		t.Fatalf("collision preflight created output: %v", err)
	}
}

func TestCreateRejectsUnsafeProfileValuesBeforeWriting(t *testing.T) {
	root := filepath.Join(t.TempDir(), "workspace")
	profile := Profile{ID: "BAD\nID", Name: "Unsafe", BodySizePT: 11, NoteSizePT: 9}
	if _, err := Create(root, profile); err == nil || !strings.Contains(err.Error(), "control character") {
		t.Fatalf("unsafe profile was accepted: %v", err)
	}
	if _, err := os.Stat(root); !os.IsNotExist(err) {
		t.Fatalf("unsafe profile created output: %v", err)
	}
	profile = Profile{ID: "BAD-SIZE", Name: "Unsafe", BodySizePT: math.NaN(), NoteSizePT: 9}
	if _, err := Create(root, profile); err == nil || !strings.Contains(err.Error(), "finite") {
		t.Fatalf("non-finite profile size was accepted: %v", err)
	}
	for name, profile := range map[string]Profile{
		"zero size":    {ID: "BAD-ZERO", Name: "Unsafe", BodyFont: "Times New Roman", NoteFont: "Times New Roman", BodySizePT: 0, NoteSizePT: 9},
		"missing font": {ID: "BAD-FONT", Name: "Unsafe", BodyFont: "", NoteFont: "Times New Roman", BodySizePT: 11, NoteSizePT: 9},
	} {
		t.Run(name, func(t *testing.T) {
			if _, err := Create(filepath.Join(t.TempDir(), "workspace"), profile); err == nil {
				t.Fatal("invalid style profile was accepted")
			}
		})
	}
}
