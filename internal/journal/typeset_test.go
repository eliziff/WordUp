package journal

import (
	"strings"
	"testing"
)

func TestInferredProfilesLoadForEveryBundledJournal(t *testing.T) {
	ids := InferredIDs()
	if len(ids) != 30 {
		t.Fatalf("expected 30 bundled inferred profiles, found %d: %v", len(ids), ids)
	}
	withLayout := 0
	for _, id := range ids {
		inf, ok, err := LoadInferred(id)
		if err != nil || !ok {
			t.Fatalf("%s: load failed: ok=%v err=%v", id, ok, err)
		}
		if inf.Status != "ok" {
			t.Fatalf("%s: status %q", id, inf.Status)
		}
		layout, _ := layoutFromInferred(inf)
		if layout.PageWidthPT > 0 {
			withLayout++
		}
		if _, err := Get(id); err != nil {
			// Every inferred dataset must be a seeded journal, otherwise the
			// evidence can never reach a generated template.
			t.Fatalf("%s: no seed profile: %v", id, err)
		}
	}
	if withLayout < 25 {
		t.Fatalf("only %d of %d inferred profiles carry a page size", withLayout, len(ids))
	}
}

func TestWithInferredOverlaysMeasuredEvidence(t *testing.T) {
	profile, err := Get("UBC-L-REV")
	if err != nil {
		t.Fatal(err)
	}
	if profile.Layout == nil || profile.Conventions == nil {
		t.Fatalf("UBC-L-REV carries no inferred layout/conventions: %+v", profile)
	}
	l := profile.Layout
	if l.PageWidthPT != 432 || l.PageHeightPT != 648 {
		t.Fatalf("UBC-L-REV page size: %v x %v", l.PageWidthPT, l.PageHeightPT)
	}
	if profile.BodyFont != "Cambria" || profile.BodySizePT != 11 || profile.NoteSizePT != 9 {
		t.Fatalf("UBC-L-REV fonts: %s %v / %v", profile.BodyFont, profile.BodySizePT, profile.NoteSizePT)
	}
	if l.HeadingCase != "upper" || l.BodyLeadPT != 13 {
		t.Fatalf("UBC-L-REV heading case %q leading %v", l.HeadingCase, l.BodyLeadPT)
	}
	if profile.Conventions.IbidCase != "Ibid" || profile.Conventions.SupraForm != "supra note N" || profile.Conventions.RangeDash != "en dash" {
		t.Fatalf("UBC-L-REV conventions: %+v", profile.Conventions)
	}
	if !strings.Contains(profile.EvidenceStatus, "published 2025/2026 articles") {
		t.Fatalf("evidence status not updated: %q", profile.EvidenceStatus)
	}
}

func TestTypesetSourceReflectsLayoutAndStages(t *testing.T) {
	profile, err := Get("UBC-L-REV")
	if err != nil {
		t.Fatal(err)
	}
	source := typesetSource(profile)
	for _, want := range []string{
		`Attribute VB_Name = "WUJournalTypesetting"`,
		"Public Const WU_TS_PAGE_WIDTH_PT As Single = 432.00",
		"Public Const WU_TS_PAGE_HEIGHT_PT As Single = 648.00",
		`Public Const WU_TS_HEADING_CASE As String = "upper"`,
		"Public Sub WU_Typeset()",
		"Public Sub WU_TypesetPageSetup()",
		"Public Sub WU_TypesetRunningHeads()",
		"Public Sub WU_TypesetCitationAudit()",
		"lower-case ibid",
		"hyphen in pinpoint range",
	} {
		if !strings.Contains(source, want) {
			t.Fatalf("typeset source lacks %q", want)
		}
	}
	form := typesetFormSource(profile)
	for _, want := range []string{"WU_BatchStart total", "WU_BeginSafeEdit updating, opened, captured", "WU_TypesetPageSetup", "WU_TypesetCitationFix", "WUJournalVolumeLabel"} {
		if !strings.Contains(form, want) {
			t.Fatalf("typeset form lacks %q", want)
		}
	}
	design := typesetFormDesign(profile)
	if design.Name != "WUJournalTypeset" {
		t.Fatalf("form name %q", design.Name)
	}
	boxes := 0
	for _, c := range design.Controls {
		if c.Type == "CheckBox" {
			boxes++
		}
	}
	if boxes != len(typesetStages(profile)) {
		t.Fatalf("form has %d checkboxes for %d stages", boxes, len(typesetStages(profile)))
	}
}

func TestTypesetStagesWithoutEvidenceAreDisabled(t *testing.T) {
	profile := Profile{ID: "NONE", Name: "No Evidence Review", BodyFont: "Times New Roman", BodySizePT: 11, NoteFont: "Times New Roman", NoteSizePT: 9}
	enabled := 0
	for _, s := range typesetStages(profile) {
		if s.Enabled {
			enabled++
			if s.Macro != "WU_TypesetStyles" {
				t.Fatalf("stage %s enabled without evidence", s.Macro)
			}
		}
	}
	if enabled != 1 {
		t.Fatalf("%d stages enabled without evidence", enabled)
	}
	source := typesetSource(profile)
	if !strings.Contains(source, `Private Const WU_TS_AUDIT_RULES As String = ""`) {
		t.Fatalf("audit rules should be empty without evidence")
	}
	design := typesetFormDesign(profile)
	for _, c := range design.Controls {
		if c.Type == "CheckBox" && c.Name != "chkStyles" {
			if v, _ := c.Properties["Enabled"].(bool); v {
				t.Fatalf("%s should be disabled without evidence", c.Name)
			}
		}
	}
}

func TestConventionRulesFlagOnlyTheUnusedForm(t *testing.T) {
	rows := conventionRules(Conventions{IbidCase: "ibid", EtAl: "et al.", Eg: "e.g.", EmphasisNote: "(emphasis added)"})
	for _, want := range []string{"capitalised Ibid||Ibid||False||True||True", "et al without period||et al ||False||False||False", "eg without periods||eg||False||True||True", "[emphasis added] in brackets"} {
		if !strings.Contains(rows, want) {
			t.Fatalf("rules lack %q in %q", want, rows)
		}
	}
	if strings.Contains(rows, "lower-case ibid") {
		t.Fatalf("rules flag the journal's own form: %q", rows)
	}
	fixes := conventionFixes(Conventions{IbidCase: "Ibid", Ie: "ie"})
	if !strings.Contains(fixes, "ibid to Ibid||ibid||Ibid||True||True") || !strings.Contains(fixes, "i.e. to ie||i.e.||ie||False||False") {
		t.Fatalf("fixes: %q", fixes)
	}
}
