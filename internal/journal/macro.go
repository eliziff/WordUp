package journal

import (
	"fmt"
	"math"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"unicode/utf8"

	"github.com/eliziff/WordUp/internal/component"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
)

type CreateReport struct {
	Journal   Profile              `json:"journal"`
	Workspace string               `json:"workspace"`
	Artifact  string               `json:"artifact"`
	Build     *project.BuildReport `json:"build,omitempty"`
}

func safeIdentifier(value string) string {
	var b strings.Builder
	for _, r := range value {
		if (r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') || (r >= '0' && r <= '9') || r == '_' {
			b.WriteRune(r)
		} else {
			b.WriteByte('_')
		}
	}
	result := b.String()
	if result == "" || (result[0] >= '0' && result[0] <= '9') {
		result = "Journal_" + result
	}
	return result
}

func vbaString(value string) string {
	return `"` + strings.ReplaceAll(value, `"`, `""`) + `"`
}

func projectName(profile Profile) string {
	return "WU_" + safeIdentifier(profile.ID)
}

// Create builds an ordinary source workspace and an uncompiled DOTM. Native
// compilation/signing remains an explicit outer-loop operation. The workspace
// is assembled out of sight and published only after the complete offline
// build succeeds, so callers never observe a half-created journal tree.
func Create(root string, profile Profile) (CreateReport, error) {
	if root == "" {
		return CreateReport{}, fmt.Errorf("journal workspace path required")
	}
	if err := validateProfile(profile); err != nil {
		return CreateReport{}, err
	}
	finalRoot := filepath.Clean(root)
	if _, err := os.Lstat(finalRoot); err == nil {
		return CreateReport{}, fmt.Errorf("journal workspace destination already exists")
	} else if !os.IsNotExist(err) {
		return CreateReport{}, fmt.Errorf("inspect journal workspace destination: %w", err)
	}
	if profile.BodyFont == "" {
		profile.BodyFont = "Times New Roman"
	}
	if profile.BodySizePT <= 0 {
		profile.BodySizePT = 11
	}
	if profile.NoteFont == "" {
		profile.NoteFont = profile.BodyFont
	}
	if profile.NoteSizePT <= 0 {
		profile.NoteSizePT = 9
	}
	if len(profile.Features) == 0 {
		profile.Features = defaultFeatures(profile.PermalinkPolicy)
	}
	name := projectName(profile)
	parent, err := filepath.Abs(filepath.Dir(finalRoot))
	if err != nil {
		return CreateReport{}, err
	}
	if err := os.MkdirAll(parent, 0700); err != nil {
		return CreateReport{}, err
	}
	stageParent, err := os.MkdirTemp(parent, ".wordwright-journal-*")
	if err != nil {
		return CreateReport{}, err
	}
	defer os.RemoveAll(stageParent)
	stageRoot := filepath.Join(stageParent, "workspace")
	if _, err := project.New(name, stageRoot); err != nil {
		return CreateReport{}, err
	}
	// Keep the generated workspace useful as a real agent starting point:
	// vendor the small, dependency-free primitives that its source may call.
	// These are copied source, not runtime references, and the component lock
	// makes their provenance and starting hashes visible to Git.
	for _, id := range []string{
		"structure.detect",
		"operation.safe-edit",
		"ui.form-shell",
		"ui.progress-cancel",
		"ui.ribbon-command",
		"command.hotkey",
		"command.context-menu",
		"document.style-converter",
		"document.text-operations",
	} {
		if _, err := component.AddWith(stageRoot, id, nil); err != nil {
			return CreateReport{}, fmt.Errorf("install %s: %w", id, err)
		}
	}
	module := safeIdentifier(profile.ID)
	writes := map[string][]byte{
		"journal/profile.json":            project.JSON(profile),
		"vba/WUJournalCore.bas":           []byte(coreSource(profile, module)),
		"vba/WUJournalSetup.vba":          []byte(formSource(profile)),
		"forms/WUJournalSetup.json":       project.JSON(formDesign(profile)),
		"package/customUI/customUI14.xml": []byte(ribbonSource(profile, module)),
	}
	for path, data := range writes {
		if err := project.Write(stageRoot, path, data, ""); err != nil {
			return CreateReport{}, err
		}
	}
	artifact := filepath.Join(stageRoot, "dist", name+".dotm")
	w, err := project.Open(stageRoot)
	if err != nil {
		return CreateReport{}, err
	}
	build, err := w.Build(artifact)
	if err != nil {
		return CreateReport{}, err
	}
	finalArtifact := filepath.Join(finalRoot, "dist", name+".dotm")
	build.Artifact = finalArtifact
	if err := project.Write(stageRoot, "reports/build.json", project.JSON(build), ""); err != nil {
		return CreateReport{}, err
	}
	// The destination was checked before work began, but check again just
	// before publication so a concurrent creator is never overwritten.
	if _, err := os.Lstat(finalRoot); err == nil {
		return CreateReport{}, fmt.Errorf("journal workspace destination appeared during build")
	} else if !os.IsNotExist(err) {
		return CreateReport{}, fmt.Errorf("inspect journal workspace destination: %w", err)
	}
	if err := os.Rename(stageRoot, finalRoot); err != nil {
		return CreateReport{}, fmt.Errorf("publish journal workspace: %w", err)
	}
	return CreateReport{Journal: profile, Workspace: root, Artifact: finalArtifact, Build: build}, nil
}

func validateProfile(profile Profile) error {
	if profile.ID == "" {
		return fmt.Errorf("journal profile id required")
	}
	for label, value := range map[string]string{
		"id": profile.ID, "name": profile.Name, "abbreviation": profile.Abbreviation,
		"scope": profile.Scope, "permalink_policy": profile.PermalinkPolicy,
		"body_font": profile.BodyFont, "note_font": profile.NoteFont, "evidence_status": profile.EvidenceStatus,
	} {
		if !utf8.ValidString(value) {
			return fmt.Errorf("journal profile %s is not valid UTF-8", label)
		}
		if strings.ContainsAny(value, "\x00\r\n") {
			return fmt.Errorf("journal profile %s contains a prohibited control character", label)
		}
	}
	for i, feature := range profile.Features {
		if !utf8.ValidString(feature) || strings.ContainsAny(feature, "\x00\r\n") {
			return fmt.Errorf("journal profile feature %d contains a prohibited character", i)
		}
	}
	if math.IsNaN(profile.BodySizePT) || math.IsInf(profile.BodySizePT, 0) || math.IsNaN(profile.NoteSizePT) || math.IsInf(profile.NoteSizePT, 0) {
		return fmt.Errorf("journal profile font sizes must be finite")
	}
	return nil
}

func CreateAll(root string) ([]CreateReport, error) {
	if root == "" {
		return nil, fmt.Errorf("journal output path required")
	}
	return createAll(root, Profiles())
}

// WithCatalogEvidence overlays the mechanically observed fields for a
// bundled profile while retaining the stable profile name, feature order and
// safe defaults. It is intentionally pure so an agent can review the result
// before asking Create to write anything.
func WithCatalogEvidence(profile Profile, catalog Catalog) Profile {
	for _, observed := range catalog.Journals {
		if strings.EqualFold(observed.ID, profile.ID) {
			// Keep the bundled identity and ordered tool surface authoritative;
			// only overlay fields that a catalog actually measures.
			if observed.BodyFont != "" {
				profile.BodyFont = observed.BodyFont
			}
			if observed.BodySizePT > 0 {
				profile.BodySizePT = observed.BodySizePT
			}
			if observed.NoteFont != "" {
				profile.NoteFont = observed.NoteFont
			}
			if observed.NoteSizePT > 0 {
				profile.NoteSizePT = observed.NoteSizePT
			}
			profile.CurrentYears = append([]int(nil), observed.CurrentYears...)
			profile.Volumes = append([]VolumeEvidence(nil), observed.Volumes...)
			profile.StyleEvidence = observed.StyleEvidence
			profile.Observed = observed.Observed
			if observed.EvidenceStatus != "" {
				profile.EvidenceStatus = observed.EvidenceStatus
			}
			return profile
		}
	}
	return profile
}

// CreateFromCatalog uses one catalog snapshot for a single generated
// workspace. The snapshot is copied into journal/profile.json; the corpus is
// not retained as a runtime input.
func CreateFromCatalog(root string, profile Profile, catalog Catalog) (CreateReport, error) {
	return Create(root, WithCatalogEvidence(profile, catalog))
}

// CreateAllWithCatalog applies one already-scanned catalog to every bundled
// profile, avoiding a repeated corpus scan for a batch operation.
func CreateAllWithCatalog(root string, catalog Catalog) ([]CreateReport, error) {
	profiles := Profiles()
	for i := range profiles {
		profiles[i] = WithCatalogEvidence(profiles[i], catalog)
	}
	return createAll(root, profiles)
}

func createAll(root string, profiles []Profile) ([]CreateReport, error) {
	if root == "" {
		return nil, fmt.Errorf("journal output path required")
	}
	// Preflight every destination before starting workers. A batch must not
	// leave half a release generated merely because a later profile already
	// exists.
	destinations := map[string]string{}
	for _, profile := range profiles {
		if err := validateProfile(profile); err != nil {
			return nil, fmt.Errorf("%s: %w", profile.ID, err)
		}
		dir := filepath.Join(root, safeIdentifier(profile.ID))
		key := strings.ToLower(filepath.ToSlash(filepath.Clean(dir)))
		if prior, exists := destinations[key]; exists {
			return nil, fmt.Errorf("%s: destination collides with %s", profile.ID, prior)
		}
		destinations[key] = profile.ID
		if _, err := os.Lstat(dir); err == nil {
			return nil, fmt.Errorf("%s: destination already exists", profile.ID)
		} else if !os.IsNotExist(err) {
			return nil, fmt.Errorf("%s: inspect destination: %w", profile.ID, err)
		}
	}
	reports := make([]CreateReport, len(profiles))
	errs := make([]error, len(profiles))
	jobs := make(chan int)
	workers := runtime.GOMAXPROCS(0)
	if workers < 1 {
		workers = 1
	}
	if workers > 4 {
		workers = 4
	}
	var group sync.WaitGroup
	group.Add(workers)
	for worker := 0; worker < workers; worker++ {
		go func() {
			defer group.Done()
			for index := range jobs {
				profile := profiles[index]
				dir := filepath.Join(root, safeIdentifier(profile.ID))
				reports[index], errs[index] = Create(dir, profile)
				if errs[index] != nil {
					errs[index] = fmt.Errorf("%s: %w", profile.ID, errs[index])
				}
			}
		}()
	}
	for index := range profiles {
		jobs <- index
	}
	close(jobs)
	group.Wait()
	for index, err := range errs {
		if err != nil {
			out := make([]CreateReport, 0, index)
			for _, report := range reports[:index] {
				if report.Workspace != "" {
					out = append(out, report)
				}
			}
			return out, err
		}
	}
	return reports, nil
}

func profileHasFeature(profile Profile, feature string) bool {
	if len(profile.Features) == 0 {
		return true
	}
	for _, candidate := range profile.Features {
		if strings.EqualFold(strings.TrimSpace(candidate), feature) {
			return true
		}
	}
	return false
}

func formDesign(profile Profile) office.Design {
	styles := profileHasFeature(profile, "style conversion") || profileHasFeature(profile, "heading and contents")
	fields := profileHasFeature(profile, "field refresh")
	citations := profileHasFeature(profile, "footnote and citation tools")
	permalink := !strings.EqualFold(profile.PermalinkPolicy, "none") && profileHasFeature(profile, "permalink assistant")
	supra := profileHasFeature(profile, "supra tools")
	tracking := profileHasFeature(profile, "tracked changes")
	quality := profileHasFeature(profile, "quality report")
	preflight := profileHasFeature(profile, "preflight")
	controls := []office.ControlDesign{
		{Name: "lblTitle", Type: "Label", Properties: map[string]any{"Left": 14.0, "Top": 12.0, "Width": 348.0, "Height": 28.0, "Caption": profile.Name}},
		{Name: "lblEvidence", Type: "Label", Properties: map[string]any{"Left": 14.0, "Top": 42.0, "Width": 348.0, "Height": 32.0, "Caption": "Profile-driven tools preserve source text and inline formatting."}},
	}
	if styles {
		controls = append(controls, office.ControlDesign{Name: "cmdStyles", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": 86.0, "Width": 164.0, "Height": 28.0, "Caption": "Apply house styles"}})
	}
	controls = append(controls, office.ControlDesign{Name: "cmdReview", Type: "CommandButton", Properties: map[string]any{"Left": 194.0, "Top": 86.0, "Width": 164.0, "Height": 28.0, "Caption": "Review changes"}})
	if fields {
		controls = append(controls, office.ControlDesign{Name: "cmdFields", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": 124.0, "Width": 164.0, "Height": 28.0, "Caption": "Refresh fields"}})
	}
	if citations {
		controls = append(controls, office.ControlDesign{Name: "cmdCitations", Type: "CommandButton", Properties: map[string]any{"Left": 194.0, "Top": 124.0, "Width": 164.0, "Height": 28.0, "Caption": "Audit citations"}})
	}
	if permalink {
		controls = append(controls, office.ControlDesign{Name: "cmdPerma", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": 162.0, "Width": 164.0, "Height": 28.0, "Caption": "Perma assistant"}})
	}
	if supra {
		controls = append(controls, office.ControlDesign{Name: "cmdSupra", Type: "CommandButton", Properties: map[string]any{"Left": 194.0, "Top": 162.0, "Width": 164.0, "Height": 28.0, "Caption": "Supra audit"}})
	}
	controls = append(controls,
		office.ControlDesign{Name: "cmdCommands", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": 200.0, "Width": 164.0, "Height": 28.0, "Caption": "Install commands"}},
		office.ControlDesign{Name: "cmdRemoveCommands", Type: "CommandButton", Properties: map[string]any{"Left": 194.0, "Top": 200.0, "Width": 164.0, "Height": 28.0, "Caption": "Remove commands"}},
	)
	if tracking {
		controls = append(controls, office.ControlDesign{Name: "cmdTracking", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": 238.0, "Width": 164.0, "Height": 28.0, "Caption": "Toggle tracked changes"}})
	}
	if quality {
		controls = append(controls, office.ControlDesign{Name: "cmdQuality", Type: "CommandButton", Properties: map[string]any{"Left": 194.0, "Top": 238.0, "Width": 164.0, "Height": 28.0, "Caption": "Quality report"}})
	}
	if preflight {
		controls = append(controls, office.ControlDesign{Name: "cmdPreflight", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": 276.0, "Width": 344.0, "Height": 28.0, "Caption": "Run preflight"}})
	}
	controls = append(controls, office.ControlDesign{Name: "cmdClose", Type: "CommandButton", Properties: map[string]any{"Left": 274.0, "Top": 314.0, "Width": 84.0, "Height": 28.0, "Caption": "Close"}})
	return office.Design{
		Name: "WUJournalSetup",
		Mode: "replace",
		Properties: map[string]any{
			"Caption": "Journal setup - " + profile.Name,
			"Width":   380.0,
			"Height":  350.0,
		},
		Controls: controls,
	}
}

func formHandler(control, macro, title string) string {
	return fmt.Sprintf(`Private Sub %s_Click()
    On Error GoTo Failed
    %s
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, %s
End Sub`, control, macro, vbaString(title))
}

func formSource(profile Profile) string {
	styles := profileHasFeature(profile, "style conversion") || profileHasFeature(profile, "heading and contents")
	fields := profileHasFeature(profile, "field refresh")
	citations := profileHasFeature(profile, "footnote and citation tools")
	permalink := !strings.EqualFold(profile.PermalinkPolicy, "none") && profileHasFeature(profile, "permalink assistant")
	supra := profileHasFeature(profile, "supra tools")
	tracking := profileHasFeature(profile, "tracked changes")
	quality := profileHasFeature(profile, "quality report")
	preflight := profileHasFeature(profile, "preflight")
	parts := []string{`Attribute VB_Name = "WUJournalSetup"
Option Explicit`}
	if styles {
		parts = append(parts, formHandler("cmdStyles", "WU_JournalApplyStyles", "Journal setup"))
	}
	parts = append(parts, formHandler("cmdReview", "WU_JournalReviewNext", "Journal review"))
	if fields {
		parts = append(parts, formHandler("cmdFields", "WU_JournalRefreshFields", "Journal fields"))
	}
	if citations {
		parts = append(parts, formHandler("cmdCitations", "WU_JournalCitationAudit", "Citation audit"))
	}
	if permalink {
		parts = append(parts, formHandler("cmdPerma", "WU_JournalPermaAssistant", "Perma assistant"))
	}
	if supra {
		parts = append(parts, formHandler("cmdSupra", "WU_JournalSupraAudit", "Supra audit"))
	}
	parts = append(parts,
		formHandler("cmdCommands", "WU_JournalInstallCommands", "Journal commands"),
		formHandler("cmdRemoveCommands", "WU_JournalRemoveCommands", "Journal commands"),
	)
	if tracking {
		parts = append(parts, formHandler("cmdTracking", "WU_JournalToggleTracking", "Tracked changes"))
	}
	if quality {
		parts = append(parts, formHandler("cmdQuality", "WU_JournalQualityReport", "Quality report"))
	}
	if preflight {
		parts = append(parts, formHandler("cmdPreflight", "WU_JournalPreflight", "Journal preflight"))
	}
	parts = append(parts, `Private Sub cmdClose_Click()
    Unload Me
End Sub`)
	return strings.Join(parts, "\n\n") + "\n"
}

func ribbonSource(profile Profile, module string) string {
	prefix := "WU_" + module
	button := func(id, label, action string) string {
		return fmt.Sprintf(`          <button id="%s_%s" label="%s" onAction="%s"/>`, prefix, id, EscapeXML(label), action)
	}
	styles := profileHasFeature(profile, "style conversion") || profileHasFeature(profile, "heading and contents")
	fields := profileHasFeature(profile, "field refresh")
	citations := profileHasFeature(profile, "footnote and citation tools")
	permalink := !strings.EqualFold(profile.PermalinkPolicy, "none") && profileHasFeature(profile, "permalink assistant")
	supra := profileHasFeature(profile, "supra tools")
	tracking := profileHasFeature(profile, "tracked changes")
	quality := profileHasFeature(profile, "quality report")
	preflight := profileHasFeature(profile, "preflight")
	buttons := []string{fmt.Sprintf(`          <button id="%s_Setup" label="Journal setup" size="large" onAction="WU_JournalOpenSetupFromRibbon"/>`, prefix)}
	if styles {
		buttons = append(buttons, button("Styles", "Apply house styles", "WU_JournalApplyStylesFromRibbon"))
	}
	buttons = append(buttons, button("Review", "Review changes", "WU_JournalReviewNextFromRibbon"))
	if fields {
		buttons = append(buttons, button("Fields", "Refresh fields", "WU_JournalRefreshFieldsFromRibbon"))
	}
	if citations {
		buttons = append(buttons, button("Citations", "Audit citations", "WU_JournalCitationAuditFromRibbon"))
	}
	if permalink {
		buttons = append(buttons, button("Perma", "Perma assistant", "WU_JournalPermaAssistantFromRibbon"))
	}
	if supra {
		buttons = append(buttons, button("Supra", "Supra audit", "WU_JournalSupraAuditFromRibbon"))
	}
	buttons = append(buttons,
		button("Commands", "Install commands", "WU_JournalInstallCommandsFromRibbon"),
		button("RemoveCommands", "Remove commands", "WU_JournalRemoveCommandsFromRibbon"),
	)
	if tracking {
		buttons = append(buttons, button("Tracking", "Toggle tracked changes", "WU_JournalToggleTrackingFromRibbon"))
	}
	if quality {
		buttons = append(buttons, button("Quality", "Quality report", "WU_JournalQualityReportFromRibbon"))
	}
	if preflight {
		buttons = append(buttons, button("Preflight", "Run preflight", "WU_JournalPreflightFromRibbon"))
	}
	return fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="WU_JournalRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="%s_Tab" label="%s">
        <group id="%s_Group" label="Journal tools">
%s
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
`, prefix, EscapeXML(profile.Name), prefix, strings.Join(buttons, "\n"))
}

func coreSource(profile Profile, module string) string {
	styleBase := "WU " + profile.ID
	return fmt.Sprintf(`Attribute VB_Name = "WUJournalCore"
Option Explicit

Public Const WU_JOURNAL_ID As String = %s
Public Const WU_JOURNAL_NAME As String = %s
Public Const WU_JOURNAL_BODY_FONT As String = %s
Public Const WU_JOURNAL_NOTE_FONT As String = %s
Public Const WU_JOURNAL_BODY_SIZE As Single = %.2f
Public Const WU_JOURNAL_NOTE_SIZE As Single = %.2f
Public Const WU_JOURNAL_PERMALINK_POLICY As String = %s
Public Const WU_JOURNAL_STYLE_BODY As String = %s
Public Const WU_JOURNAL_STYLE_NOTE As String = %s
Public Const WU_JOURNAL_STYLE_H1 As String = %s
Public Const WU_JOURNAL_STYLE_H2 As String = %s
Public Const WU_JOURNAL_STYLE_H3 As String = %s
Public Const WU_JOURNAL_STYLE_H4 As String = %s
Public Const WU_JOURNAL_STYLE_H5 As String = %s
Public Const WU_JOURNAL_STYLE_H6 As String = %s
Public Const WU_JOURNAL_STYLE_H7 As String = %s
Public Const WU_JOURNAL_STYLE_H8 As String = %s
Public Const WU_JOURNAL_STYLE_H9 As String = %s

Public Sub WU_JournalRibbonLoad(ByVal ribbon As IRibbonUI)
    ' The callback is intentionally a no-op. Keeping it public makes the
    ' Ribbon fragment independently valid and leaves refresh policy to Word.
End Sub

Public Sub WU_JournalOpenSetupFromRibbon(ByVal control As IRibbonControl)
    WU_JournalOpenSetup
End Sub

Public Sub WU_JournalApplyStylesFromRibbon(ByVal control As IRibbonControl)
    WU_JournalApplyStyles
End Sub

Public Sub WU_JournalReviewNextFromRibbon(ByVal control As IRibbonControl)
    WU_JournalReviewNext
End Sub

Public Sub WU_JournalRefreshFieldsFromRibbon(ByVal control As IRibbonControl)
    WU_JournalRefreshFields
End Sub

Public Sub WU_JournalCitationAuditFromRibbon(ByVal control As IRibbonControl)
    WU_JournalCitationAudit
End Sub

Public Sub WU_JournalPermaAssistantFromRibbon(ByVal control As IRibbonControl)
    WU_JournalPermaAssistant
End Sub

Public Sub WU_JournalSupraAuditFromRibbon(ByVal control As IRibbonControl)
    WU_JournalSupraAudit
End Sub

Public Sub WU_JournalInstallCommandsFromRibbon(ByVal control As IRibbonControl)
    WU_JournalInstallCommands
End Sub

Public Sub WU_JournalRemoveCommandsFromRibbon(ByVal control As IRibbonControl)
    WU_JournalRemoveCommands
End Sub

Public Sub WU_JournalToggleTrackingFromRibbon(ByVal control As IRibbonControl)
    WU_JournalToggleTracking
End Sub

Public Sub WU_JournalQualityReportFromRibbon(ByVal control As IRibbonControl)
    WU_JournalQualityReport
End Sub

Public Sub WU_JournalPreflightFromRibbon(ByVal control As IRibbonControl)
    WU_JournalPreflight
End Sub

Public Sub WU_JournalOpenSetup()
    WUJournalSetup.Show vbModeless
End Sub

Public Sub WU_JournalApplyStyles()
    Dim updating As Boolean, undoStarted As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim priorStatus As Variant
    Dim doc As Document, bodyStyle As Style, noteStyle As Style
    Dim heading1 As Style, heading2 As Style, heading3 As Style
    Dim heading4 As Style, heading5 As Style, heading6 As Style
    Dim heading7 As Style, heading8 As Style, heading9 As Style
    On Error GoTo Failed
    Set doc = ActiveDocument
    priorStatus = Application.StatusBar
    WU_BeginSafeEdit updating, undoStarted, captured, "Apply %s styles"
    WU_ResetProgress
    Set bodyStyle = WU_EnsureStyle(doc, WU_JOURNAL_STYLE_BODY, WU_JOURNAL_BODY_FONT, WU_JOURNAL_BODY_SIZE, False)
    Set noteStyle = WU_EnsureStyle(doc, WU_JOURNAL_STYLE_NOTE, WU_JOURNAL_NOTE_FONT, WU_JOURNAL_NOTE_SIZE, False)
    Set heading1 = WU_EnsureStyle(doc, WU_JOURNAL_STYLE_H1, WU_JOURNAL_BODY_FONT, WU_JOURNAL_BODY_SIZE + 1, True)
    Set heading2 = WU_EnsureStyle(doc, WU_JOURNAL_STYLE_H2, WU_JOURNAL_BODY_FONT, WU_JOURNAL_BODY_SIZE, True)
    Set heading3 = WU_EnsureStyle(doc, WU_JOURNAL_STYLE_H3, WU_JOURNAL_BODY_FONT, WU_JOURNAL_BODY_SIZE, True)
    Set heading4 = WU_EnsureStyle(doc, WU_JOURNAL_STYLE_H4, WU_JOURNAL_BODY_FONT, WU_JOURNAL_BODY_SIZE, True)
    Set heading5 = WU_EnsureStyle(doc, WU_JOURNAL_STYLE_H5, WU_JOURNAL_BODY_FONT, WU_JOURNAL_BODY_SIZE, True)
    Set heading6 = WU_EnsureStyle(doc, WU_JOURNAL_STYLE_H6, WU_JOURNAL_BODY_FONT, WU_JOURNAL_BODY_SIZE, True)
    Set heading7 = WU_EnsureStyle(doc, WU_JOURNAL_STYLE_H7, WU_JOURNAL_BODY_FONT, WU_JOURNAL_BODY_SIZE, True)
    Set heading8 = WU_EnsureStyle(doc, WU_JOURNAL_STYLE_H8, WU_JOURNAL_BODY_FONT, WU_JOURNAL_BODY_SIZE, True)
    Set heading9 = WU_EnsureStyle(doc, WU_JOURNAL_STYLE_H9, WU_JOURNAL_BODY_FONT, WU_JOURNAL_BODY_SIZE, True)
    heading1.ParagraphFormat.OutlineLevel = wdOutlineLevel1
    heading2.ParagraphFormat.OutlineLevel = wdOutlineLevel2
    heading3.ParagraphFormat.OutlineLevel = wdOutlineLevel3
    heading4.ParagraphFormat.OutlineLevel = wdOutlineLevel4
    heading5.ParagraphFormat.OutlineLevel = wdOutlineLevel5
    heading6.ParagraphFormat.OutlineLevel = wdOutlineLevel6
    heading7.ParagraphFormat.OutlineLevel = wdOutlineLevel7
    heading8.ParagraphFormat.OutlineLevel = wdOutlineLevel8
    heading9.ParagraphFormat.OutlineLevel = wdOutlineLevel9
    WU_ApplyParagraphStyles doc, bodyStyle, heading1, heading2, heading3, heading4, heading5, heading6, heading7, heading8, heading9
    WU_ApplyFootnoteStyle doc, noteStyle
Cleanup:
    On Error Resume Next
    WU_EndSafeEdit updating, undoStarted, captured
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Application.StatusBar = priorStatus
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Sub
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume Cleanup
End Sub

Private Function WU_EnsureStyle(ByVal doc As Document, ByVal styleName As String, ByVal fontName As String, ByVal fontSize As Single, ByVal keepNext As Boolean) As Style
    Dim value As Style
    On Error Resume Next
    Set value = doc.Styles(styleName)
    On Error GoTo 0
    If value Is Nothing Then Set value = doc.Styles.Add(Name:=styleName, Type:=wdStyleTypeParagraph)
    If value.Type <> wdStyleTypeParagraph Then Err.Raise 5, "WU_EnsureStyle", "style " & styleName & " is not a paragraph style"
    With value
        .Font.Name = fontName
        .Font.NameAscii = fontName
        .Font.NameOther = fontName
        .Font.NameFarEast = fontName
        .Font.NameBi = fontName
        .Font.Size = fontSize
        .ParagraphFormat.SpaceAfter = 6
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .ParagraphFormat.KeepWithNext = keepNext
    End With
    Set WU_EnsureStyle = value
End Function

Private Sub WU_ApplyParagraphStyles(ByVal doc As Document, ByVal bodyStyle As Style, ByVal heading1 As Style, ByVal heading2 As Style, ByVal heading3 As Style, ByVal heading4 As Style, ByVal heading5 As Style, ByVal heading6 As Style, ByVal heading7 As Style, ByVal heading8 As Style, ByVal heading9 As Style)
    Dim story As Range, paragraph As Paragraph, paragraphRange As Range, batch As Range
    Dim batchStyle As Style, desiredStyle As Style, currentStyle As String, detectedRole As String
    Dim structure As Variant, haveStructure As Boolean, paragraphIndex As Long, level As Long
    On Error Resume Next
    Set story = doc.StoryRanges(wdMainTextStory)
    On Error GoTo 0
    If story Is Nothing Then Exit Sub
    ' Run the detector once. It retains native outline, style-family and
    ' marker evidence while keeping the formatting pass linear. If an input
    ' is too large or cannot expose its main story, native outline evidence
    ' remains a safe fallback rather than making style application fail.
    On Error Resume Next
    structure = WU_DetectStructure(doc)
    haveStructure = (Err.Number = 0 And IsArray(structure))
    Err.Clear
    On Error GoTo 0
    paragraphIndex = 0
    For Each paragraph In story.Paragraphs
        paragraphIndex = paragraphIndex + 1
        If paragraphIndex Mod 256 = 0 Then
            Application.StatusBar = "Applying " & WU_JOURNAL_NAME & " styles (paragraph " & CStr(paragraphIndex) & ")"
            If WU_CancelRequested() Then Err.Raise 18, "Apply styles", "style application cancelled"
        End If
        Set paragraphRange = paragraph.Range
        Set desiredStyle = Nothing
        If Not paragraphRange.Information(wdWithInTable) Then
            level = paragraph.OutlineLevel
            ' Reset before the guarded array read. A malformed or stale
            ' detector result must fall back to Word's native outline level,
            ' never reuse the previous paragraph's role.
            detectedRole = vbNullString
            If haveStructure Then
                On Error Resume Next
                detectedRole = CStr(structure(paragraphIndex - 1, WU_ROLE))
                If StrComp(detectedRole, "heading", vbTextCompare) = 0 Then level = CLng(structure(paragraphIndex - 1, WU_LEVEL))
                Err.Clear
                On Error GoTo 0
            End If
            If level = wdOutlineLevel1 Then
                Set desiredStyle = heading1
            ElseIf level = wdOutlineLevel2 Then
                Set desiredStyle = heading2
            ElseIf level = wdOutlineLevel3 Then
                Set desiredStyle = heading3
            ElseIf level = wdOutlineLevel4 Then
                Set desiredStyle = heading4
            ElseIf level = wdOutlineLevel5 Then
                Set desiredStyle = heading5
            ElseIf level = wdOutlineLevel6 Then
                Set desiredStyle = heading6
            ElseIf level = wdOutlineLevel7 Then
                Set desiredStyle = heading7
            ElseIf level = wdOutlineLevel8 Then
                Set desiredStyle = heading8
            ElseIf level = wdOutlineLevel9 Then
                Set desiredStyle = heading9
            ElseIf StrComp(detectedRole, "body", vbTextCompare) = 0 Then
                ' The neutral detector has enough evidence to distinguish an
                ' ordinary body paragraph from a candidate, quotation, or
                ' front-matter role. Apply the house body style even when the
                ' source used a custom paragraph style; direct run formatting
                ' remains direct formatting on the existing range.
                Set desiredStyle = bodyStyle
            Else
                currentStyle = ""
                On Error Resume Next
                currentStyle = CStr(paragraphRange.Style)
                On Error GoTo 0
                If StrComp(currentStyle, "Normal", vbTextCompare) = 0 Or StrComp(currentStyle, "Body Text", vbTextCompare) = 0 Then Set desiredStyle = bodyStyle
            End If
        End If
        If desiredStyle Is Nothing Then
            WU_FlushParagraphStyleBatch batch, batchStyle
        ElseIf batch Is Nothing Then
            Set batch = paragraphRange.Duplicate
            Set batchStyle = desiredStyle
        ElseIf desiredStyle Is batchStyle And paragraphRange.Start <= batch.End Then
            batch.End = paragraphRange.End
        Else
            WU_FlushParagraphStyleBatch batch, batchStyle
            Set batch = paragraphRange.Duplicate
            Set batchStyle = desiredStyle
        End If
    Next paragraph
    WU_FlushParagraphStyleBatch batch, batchStyle
End Sub

Private Sub WU_FlushParagraphStyleBatch(ByRef batch As Range, ByRef style As Style)
    If batch Is Nothing Then Exit Sub
    batch.Style = style
    Set batch = Nothing
    Set style = Nothing
End Sub

Private Sub WU_ApplyFootnoteStyle(ByVal doc As Document, ByVal noteStyle As Style)
    Dim story As Range
    ' A note story is already a contiguous Range. Applying its paragraph
    ' style once avoids one COM round-trip per note while leaving direct run
    ' formatting (italic case names, emphasis, and fields) untouched.
    On Error Resume Next
    Set story = doc.StoryRanges(wdFootnotesStory)
    On Error GoTo 0
    If Not story Is Nothing Then story.Style = noteStyle
    Set story = Nothing
    On Error Resume Next
    Set story = doc.StoryRanges(wdEndnotesStory)
    On Error GoTo 0
    If Not story Is Nothing Then story.Style = noteStyle
End Sub

Public Sub WU_JournalRefreshFields()
    Dim updating As Boolean, undoStarted As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim doc As Document, story As Range, linked As Range, contents As TableOfContents
    Dim fieldFailures As Long, contentsFailures As Long
    On Error GoTo Failed
    Set doc = ActiveDocument
    WU_BeginSafeEdit updating, undoStarted, captured, "Refresh %s fields"
    For Each story In doc.StoryRanges
        Set linked = story
        Do While Not linked Is Nothing
            WU_UpdateFieldsInStory linked, fieldFailures
            Set linked = linked.NextStoryRange
        Loop
    Next story
    ' StoryRanges already contains every header/footer story and its linked
    ' sections. Walking Sections as well updates those fields twice and can
    ' make a large manuscript needlessly repaginate.
    For Each contents In doc.TablesOfContents
        WU_UpdateContents contents, contentsFailures
    Next contents
    If fieldFailures + contentsFailures > 0 Then Err.Raise 5, "WU_JournalRefreshFields", "could not refresh fields in " & CStr(fieldFailures) & " story(s) and " & CStr(contentsFailures) & " table(s)"
Cleanup:
    On Error Resume Next
    WU_EndSafeEdit updating, undoStarted, captured
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Sub
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume Cleanup
End Sub

Private Sub WU_UpdateFieldsInStory(ByVal story As Range, ByRef failures As Long)
    On Error Resume Next
    story.Fields.Update
    If Err.Number <> 0 Then failures = failures + 1
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub WU_UpdateContents(ByVal contents As TableOfContents, ByRef failures As Long)
    On Error Resume Next
    contents.Update
    If Err.Number <> 0 Then failures = failures + 1
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub WU_JournalReviewNext()
    Dim firstStory As Range, story As Range, revision As Revision
    On Error GoTo Failed
    For Each firstStory In ActiveDocument.StoryRanges
        Set story = firstStory
        Do While Not story Is Nothing
            Set revision = Nothing
            On Error Resume Next
            If story.Revisions.Count > 0 Then
                Set revision = story.Revisions(1)
            Else
                Set revision = Nothing
            End If
            Err.Clear
            On Error GoTo Failed
            If Not revision Is Nothing Then
                revision.Range.Select
                Application.StatusBar = "Selected the next revision for review."
                Exit Sub
            End If
            Set story = story.NextStoryRange
        Loop
    Next firstStory
    MsgBox "No tracked changes were found in the document stories.", vbInformation, "Journal review"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Journal review"
End Sub

Public Sub WU_JournalCitationAudit()
    Dim noteCount As Long, supraCount As Long
    On Error GoTo Failed
    WU_ScanNotes ActiveDocument, noteCount, supraCount
    MsgBox "Footnotes and endnotes: " & CStr(noteCount) & vbCrLf & "Notes containing 'supra': " & CStr(supraCount), vbInformation, "Citation audit"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Citation audit"
End Sub

Public Sub WU_JournalSupraAudit()
    Dim noteCount As Long, supraCount As Long
    On Error GoTo Failed
    WU_ScanNotes ActiveDocument, noteCount, supraCount
    MsgBox "Notes containing 'supra': " & CStr(supraCount), vbInformation, "Supra audit"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Supra audit"
End Sub

Private Sub WU_ScanNotes(ByVal doc As Document, ByRef noteCount As Long, ByRef supraCount As Long)
    Dim note As Footnote, endnote As Endnote, noteText As String
    For Each note In doc.Footnotes
        noteCount = noteCount + 1
        noteText = note.Range.Text
        If InStr(1, noteText, "supra", vbTextCompare) > 0 Then supraCount = supraCount + 1
    Next note
    For Each endnote In doc.Endnotes
        noteCount = noteCount + 1
        noteText = endnote.Range.Text
        If InStr(1, noteText, "supra", vbTextCompare) > 0 Then supraCount = supraCount + 1
    Next endnote
End Sub

Public Sub WU_JournalInstallCommands()
    Dim hotkey As Long, hotkeyInstalled As Boolean
    Dim failure As Long, failureText As String
    On Error GoTo Failed
    hotkey = BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyJ)
    WU_RegisterHotkey hotkey, "WU_JournalOpenSetup"
    hotkeyInstalled = True
    WU_RegisterContextMenu "Journal setup", "WU_JournalOpenSetup"
    Application.StatusBar = "Journal commands installed (Ctrl+Alt+J and the text context menu)."
    Exit Sub
Failed:
    failure = Err.Number: failureText = Err.Description
    On Error Resume Next
    If hotkeyInstalled Then WU_RemoveHotkey hotkey
    On Error GoTo 0
    If failure = 0 Then failureText = "Command installation failed."
    MsgBox failureText, vbExclamation, "Journal commands"
End Sub

Public Sub WU_JournalRemoveCommands()
    On Error GoTo Failed
    WU_RemoveHotkey BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyJ)
    WU_RemoveContextMenu
    Application.StatusBar = "Journal commands removed."
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Journal commands"
End Sub

Public Sub WU_JournalToggleTracking()
    Dim doc As Document
    On Error GoTo Failed
    Set doc = ActiveDocument
    doc.TrackRevisions = Not doc.TrackRevisions
    MsgBox "Track changes is now " & IIf(doc.TrackRevisions, "on", "off") & " for this document.", vbInformation, "Tracked changes"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Tracked changes"
End Sub

Public Sub WU_JournalQualityReport()
    Dim doc As Document, report As String
    Dim fieldCount As Long, tableCount As Long, revisionCount As Long, hyperlinkCount As Long
    On Error GoTo Failed
    Set doc = ActiveDocument
    WU_CountStoryItems doc, fieldCount, tableCount, revisionCount, hyperlinkCount
    report = "Paragraphs: " & CStr(doc.Paragraphs.Count) & vbCrLf
    report = report & "Sections: " & CStr(doc.Sections.Count) & vbCrLf
    report = report & "Footnotes: " & CStr(doc.Footnotes.Count) & vbCrLf
    report = report & "Endnotes: " & CStr(doc.Endnotes.Count) & vbCrLf
    report = report & "Fields (all stories): " & CStr(fieldCount) & vbCrLf
    report = report & "Tables (all stories): " & CStr(tableCount) & vbCrLf
    report = report & "Revisions (all stories): " & CStr(revisionCount) & vbCrLf
    report = report & "Hyperlinks (all stories): " & CStr(hyperlinkCount)
    MsgBox report, vbInformation, "Quality report"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Quality report"
End Sub

Private Sub WU_CountStoryItems(ByVal doc As Document, ByRef fieldCount As Long, ByRef tableCount As Long, ByRef revisionCount As Long, ByRef hyperlinkCount As Long)
    Dim firstStory As Range, story As Range
    For Each firstStory In doc.StoryRanges
        Set story = firstStory
        Do While Not story Is Nothing
            On Error Resume Next
            fieldCount = fieldCount + story.Fields.Count
            tableCount = tableCount + story.Tables.Count
            revisionCount = revisionCount + story.Revisions.Count
            hyperlinkCount = hyperlinkCount + story.Hyperlinks.Count
            Err.Clear
            On Error GoTo 0
            Set story = story.NextStoryRange
        Loop
    Next firstStory
End Sub

Public Sub WU_JournalPreflight()
    Dim doc As Document, missing As String, issues As String, report As String
    On Error GoTo Failed
    Set doc = ActiveDocument
    If doc.ProtectionType <> wdNoProtection Then issues = issues & "Document protection is enabled." & vbCrLf
    If Not WU_HasStyle(doc, WU_JOURNAL_STYLE_BODY) Then missing = missing & WU_JOURNAL_STYLE_BODY & ", "
    If Not WU_HasStyle(doc, WU_JOURNAL_STYLE_NOTE) Then missing = missing & WU_JOURNAL_STYLE_NOTE & ", "
    If Not WU_HasStyle(doc, WU_JOURNAL_STYLE_H1) Then missing = missing & WU_JOURNAL_STYLE_H1 & ", "
    If Not WU_HasStyle(doc, WU_JOURNAL_STYLE_H2) Then missing = missing & WU_JOURNAL_STYLE_H2 & ", "
    If Not WU_HasStyle(doc, WU_JOURNAL_STYLE_H3) Then missing = missing & WU_JOURNAL_STYLE_H3 & ", "
    If Not WU_HasStyle(doc, WU_JOURNAL_STYLE_H4) Then missing = missing & WU_JOURNAL_STYLE_H4 & ", "
    If Not WU_HasStyle(doc, WU_JOURNAL_STYLE_H5) Then missing = missing & WU_JOURNAL_STYLE_H5 & ", "
    If Not WU_HasStyle(doc, WU_JOURNAL_STYLE_H6) Then missing = missing & WU_JOURNAL_STYLE_H6 & ", "
    If Not WU_HasStyle(doc, WU_JOURNAL_STYLE_H7) Then missing = missing & WU_JOURNAL_STYLE_H7 & ", "
    If Not WU_HasStyle(doc, WU_JOURNAL_STYLE_H8) Then missing = missing & WU_JOURNAL_STYLE_H8 & ", "
    If Not WU_HasStyle(doc, WU_JOURNAL_STYLE_H9) Then missing = missing & WU_JOURNAL_STYLE_H9 & ", "
    If Len(missing) > 0 Then issues = issues & "Missing generated styles: " & Left$(missing, Len(missing) - 2) & vbCrLf
    report = "Main-story paragraphs: " & CStr(doc.Paragraphs.Count) & vbCrLf
    report = report & "Sections: " & CStr(doc.Sections.Count)
    If Len(issues) = 0 Then
        MsgBox report & vbCrLf & "No preflight issues were found.", vbInformation, "Journal preflight"
    Else
        MsgBox issues & vbCrLf & report, vbExclamation, "Journal preflight"
    End If
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Journal preflight"
End Sub

Private Function WU_HasStyle(ByVal doc As Document, ByVal styleName As String) As Boolean
    Dim value As Style
    On Error Resume Next
    Set value = doc.Styles(styleName)
    WU_HasStyle = Not value Is Nothing
    On Error GoTo 0
End Function

Public Sub WU_JournalPermaAssistant()
    Dim target As Range, address As String
    Dim updating As Boolean, undoStarted As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo FailedPerma
    If StrComp(WU_JOURNAL_PERMALINK_POLICY, "none", vbTextCompare) = 0 Then
        MsgBox "This journal profile does not use permalinks.", vbInformation, "Perma assistant"
        Exit Sub
    End If
    Set target = Selection.Range.Duplicate
    If target.Start = target.End Then
        MsgBox "Select the citation text to link first.", vbExclamation, "Perma assistant"
        Exit Sub
    End If
    address = InputBox("Perma.cc URL", "Perma assistant", "https://perma.cc/")
    If Len(Trim$(address)) = 0 Then Exit Sub
    If LCase$(Left$(Trim$(address), 8)) <> "https://" Then Err.Raise 5, "Perma assistant", "Use an HTTPS permalink."
    WU_TrimAnchorParagraphMark target
    WU_BeginSafeEdit updating, undoStarted, captured, "Add permalink"
    ' Adding a hyperlink to the existing range preserves its text and direct
    ' character formatting; TextToDisplay would replace rich inline content.
    ActiveDocument.Hyperlinks.Add Anchor:=target, Address:=Trim$(address)
CleanupPerma:
    On Error Resume Next
    WU_EndSafeEdit updating, undoStarted, captured
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Sub
FailedPerma:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanupPerma
End Sub

Private Sub WU_TrimAnchorParagraphMark(ByVal target As Range)
    Dim tail As String
    If target Is Nothing Then Exit Sub
    If target.End <= target.Start Then Exit Sub
    tail = target.Characters.Last.Text
    If tail = Chr$(13) Or tail = Chr$(7) Then target.End = target.End - 1
End Sub
`, vbaString(profile.ID), vbaString(profile.Name), vbaString(profile.BodyFont), vbaString(profile.NoteFont), profile.BodySizePT, profile.NoteSizePT, vbaString(profile.PermalinkPolicy), vbaString(styleBase+" Body"), vbaString(styleBase+" Note"), vbaString(styleBase+" Heading 1"), vbaString(styleBase+" Heading 2"), vbaString(styleBase+" Heading 3"), vbaString(styleBase+" Heading 4"), vbaString(styleBase+" Heading 5"), vbaString(styleBase+" Heading 6"), vbaString(styleBase+" Heading 7"), vbaString(styleBase+" Heading 8"), vbaString(styleBase+" Heading 9"), profile.Name, profile.Name)
}
