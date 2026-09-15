package journal

import (
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"

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
// compilation/signing remains an explicit outer-loop operation.
func Create(root string, profile Profile) (CreateReport, error) {
	if root == "" {
		return CreateReport{}, fmt.Errorf("journal workspace path required")
	}
	if profile.ID == "" {
		return CreateReport{}, fmt.Errorf("journal profile id required")
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
	if _, err := project.New(name, root); err != nil {
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
	} {
		if _, err := component.AddWith(root, id, nil); err != nil {
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
		if err := project.Write(root, path, data, ""); err != nil {
			return CreateReport{}, err
		}
	}
	artifact := filepath.Join(root, "dist", name+".dotm")
	w, err := project.Open(root)
	if err != nil {
		return CreateReport{}, err
	}
	build, err := w.Build(artifact)
	if err != nil {
		return CreateReport{}, err
	}
	return CreateReport{Journal: profile, Workspace: root, Artifact: artifact, Build: build}, nil
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
	for _, profile := range profiles {
		dir := filepath.Join(root, safeIdentifier(profile.ID))
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

func formDesign(profile Profile) office.Design {
	controls := []office.ControlDesign{
		{Name: "lblTitle", Type: "Label", Properties: map[string]any{"Left": 14.0, "Top": 12.0, "Width": 348.0, "Height": 28.0, "Caption": profile.Name}},
		{Name: "lblEvidence", Type: "Label", Properties: map[string]any{"Left": 14.0, "Top": 42.0, "Width": 348.0, "Height": 32.0, "Caption": "Profile-driven tools preserve source text and inline formatting."}},
		{Name: "cmdStyles", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": 86.0, "Width": 164.0, "Height": 28.0, "Caption": "Apply house styles"}},
		{Name: "cmdReview", Type: "CommandButton", Properties: map[string]any{"Left": 194.0, "Top": 86.0, "Width": 164.0, "Height": 28.0, "Caption": "Review changes"}},
		{Name: "cmdFields", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": 124.0, "Width": 164.0, "Height": 28.0, "Caption": "Refresh fields"}},
		{Name: "cmdCitations", Type: "CommandButton", Properties: map[string]any{"Left": 194.0, "Top": 124.0, "Width": 164.0, "Height": 28.0, "Caption": "Audit citations"}},
		{Name: "cmdSupra", Type: "CommandButton", Properties: map[string]any{"Left": 194.0, "Top": 162.0, "Width": 164.0, "Height": 28.0, "Caption": "Supra audit"}},
		{Name: "cmdCommands", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": 200.0, "Width": 164.0, "Height": 28.0, "Caption": "Install commands"}},
		{Name: "cmdRemoveCommands", Type: "CommandButton", Properties: map[string]any{"Left": 194.0, "Top": 200.0, "Width": 164.0, "Height": 28.0, "Caption": "Remove commands"}},
		{Name: "cmdTracking", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": 238.0, "Width": 164.0, "Height": 28.0, "Caption": "Toggle tracked changes"}},
		{Name: "cmdQuality", Type: "CommandButton", Properties: map[string]any{"Left": 194.0, "Top": 238.0, "Width": 164.0, "Height": 28.0, "Caption": "Quality report"}},
		{Name: "cmdPreflight", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": 276.0, "Width": 344.0, "Height": 28.0, "Caption": "Run preflight"}},
		{Name: "cmdClose", Type: "CommandButton", Properties: map[string]any{"Left": 274.0, "Top": 314.0, "Width": 84.0, "Height": 28.0, "Caption": "Close"}},
	}
	if !strings.EqualFold(profile.PermalinkPolicy, "none") {
		controls = append(controls[:6], append([]office.ControlDesign{{Name: "cmdPerma", Type: "CommandButton", Properties: map[string]any{"Left": 14.0, "Top": 162.0, "Width": 164.0, "Height": 28.0, "Caption": "Perma assistant"}}}, controls[6:]...)...)
	}
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

func formSource(profile Profile) string {
	permalinkHandler := `Private Sub cmdPerma_Click()
    On Error GoTo Failed
    WU_JournalPermaAssistant
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Perma assistant"
End Sub
`
	if strings.EqualFold(profile.PermalinkPolicy, "none") {
		permalinkHandler = ""
	}
	return fmt.Sprintf(`Attribute VB_Name = "WUJournalSetup"
Option Explicit

Private Sub cmdStyles_Click()
    On Error GoTo Failed
    WU_JournalApplyStyles
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Journal setup"
End Sub

Private Sub cmdReview_Click()
    On Error GoTo Failed
    WU_JournalReviewNext
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Journal review"
End Sub

Private Sub cmdFields_Click()
    On Error GoTo Failed
    WU_JournalRefreshFields
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Journal fields"
End Sub

Private Sub cmdCitations_Click()
    On Error GoTo Failed
    WU_JournalCitationAudit
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Citation audit"
End Sub

%s

Private Sub cmdSupra_Click()
    On Error GoTo Failed
    WU_JournalSupraAudit
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Supra audit"
End Sub

Private Sub cmdCommands_Click()
    On Error GoTo Failed
    WU_JournalInstallCommands
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Journal commands"
End Sub

Private Sub cmdRemoveCommands_Click()
    On Error GoTo Failed
    WU_JournalRemoveCommands
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Journal commands"
End Sub

Private Sub cmdTracking_Click()
    On Error GoTo Failed
    WU_JournalToggleTracking
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Tracked changes"
End Sub

Private Sub cmdQuality_Click()
    On Error GoTo Failed
    WU_JournalQualityReport
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Quality report"
End Sub

Private Sub cmdPreflight_Click()
    On Error GoTo Failed
    WU_JournalPreflight
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Journal preflight"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
`, permalinkHandler)
}

func ribbonSource(profile Profile, module string) string {
	prefix := "WU_" + module
	permalinkButton := ""
	if !strings.EqualFold(profile.PermalinkPolicy, "none") {
		permalinkButton = fmt.Sprintf(`          <button id="%s_Perma" label="Perma assistant" onAction="WU_JournalPermaAssistantFromRibbon"/>
`, prefix)
	}
	return fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="WU_JournalRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="%s_Tab" label="%s">
        <group id="%s_Group" label="Journal tools">
          <button id="%s_Setup" label="Journal setup" size="large" onAction="WU_JournalOpenSetupFromRibbon"/>
          <button id="%s_Styles" label="Apply house styles" onAction="WU_JournalApplyStylesFromRibbon"/>
          <button id="%s_Review" label="Review changes" onAction="WU_JournalReviewNextFromRibbon"/>
          <button id="%s_Fields" label="Refresh fields" onAction="WU_JournalRefreshFieldsFromRibbon"/>
          <button id="%s_Citations" label="Audit citations" onAction="WU_JournalCitationAuditFromRibbon"/>
%s
          <button id="%s_Supra" label="Supra audit" onAction="WU_JournalSupraAuditFromRibbon"/>
          <button id="%s_Commands" label="Install commands" onAction="WU_JournalInstallCommandsFromRibbon"/>
          <button id="%s_RemoveCommands" label="Remove commands" onAction="WU_JournalRemoveCommandsFromRibbon"/>
          <button id="%s_Tracking" label="Toggle tracked changes" onAction="WU_JournalToggleTrackingFromRibbon"/>
          <button id="%s_Quality" label="Quality report" onAction="WU_JournalQualityReportFromRibbon"/>
          <button id="%s_Preflight" label="Run preflight" onAction="WU_JournalPreflightFromRibbon"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
`, prefix, EscapeXML(profile.Name), prefix, prefix, prefix, prefix, prefix, prefix, permalinkButton, prefix, prefix, prefix, prefix, prefix, prefix)
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
    heading1.ParagraphFormat.OutlineLevel = wdOutlineLevel1
    heading2.ParagraphFormat.OutlineLevel = wdOutlineLevel2
    heading3.ParagraphFormat.OutlineLevel = wdOutlineLevel3
    WU_ApplyParagraphStyles doc, bodyStyle, heading1, heading2, heading3
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

Private Sub WU_ApplyParagraphStyles(ByVal doc As Document, ByVal bodyStyle As Style, ByVal heading1 As Style, ByVal heading2 As Style, ByVal heading3 As Style)
    Dim story As Range, paragraph As Paragraph, currentStyle As String, detectedRole As String
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
        If Not paragraph.Range.Information(wdWithInTable) Then
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
                paragraph.Range.Style = heading1
            ElseIf level = wdOutlineLevel2 Then
                paragraph.Range.Style = heading2
            ElseIf level = wdOutlineLevel3 Then
                paragraph.Range.Style = heading3
            Else
                currentStyle = ""
                On Error Resume Next
                currentStyle = CStr(paragraph.Style)
                On Error GoTo 0
                If StrComp(currentStyle, "Normal", vbTextCompare) = 0 Or StrComp(currentStyle, "Body Text", vbTextCompare) = 0 Then
                    paragraph.Range.Style = bodyStyle
                End If
            End If
        End If
    Next paragraph
End Sub

Private Sub WU_ApplyFootnoteStyle(ByVal doc As Document, ByVal noteStyle As Style)
    Dim note As Footnote, endnote As Endnote
    For Each note In doc.Footnotes
        note.Range.Style = noteStyle
    Next note
    For Each endnote In doc.Endnotes
        endnote.Range.Style = noteStyle
    Next endnote
End Sub

Public Sub WU_JournalRefreshFields()
    Dim updating As Boolean, undoStarted As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim doc As Document, story As Range, linked As Range, contents As TableOfContents
    On Error GoTo Failed
    Set doc = ActiveDocument
    WU_BeginSafeEdit updating, undoStarted, captured, "Refresh %s fields"
    For Each story In doc.StoryRanges
        Set linked = story
        Do While Not linked Is Nothing
            WU_UpdateFieldsInStory linked
            Set linked = linked.NextStoryRange
        Loop
    Next story
    ' StoryRanges already contains every header/footer story and its linked
    ' sections. Walking Sections as well updates those fields twice and can
    ' make a large manuscript needlessly repaginate.
    For Each contents In doc.TablesOfContents
        WU_UpdateContents contents
    Next contents
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

Private Sub WU_UpdateFieldsInStory(ByVal story As Range)
    On Error Resume Next
    story.Fields.Update
    On Error GoTo 0
End Sub

Private Sub WU_UpdateContents(ByVal contents As TableOfContents)
    On Error Resume Next
    contents.Update
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
    Dim note As Footnote, endnote As Endnote, noteText As String, supraCount As Long, noteCount As Long
    For Each note In ActiveDocument.Footnotes
        noteCount = noteCount + 1
        noteText = note.Range.Text
        If InStr(1, noteText, "supra", vbTextCompare) > 0 Then supraCount = supraCount + 1
    Next note
    For Each endnote In ActiveDocument.Endnotes
        noteCount = noteCount + 1
        noteText = endnote.Range.Text
        If InStr(1, noteText, "supra", vbTextCompare) > 0 Then supraCount = supraCount + 1
    Next endnote
    MsgBox "Footnotes and endnotes: " & CStr(noteCount) & vbCrLf & "Notes containing 'supra': " & CStr(supraCount), vbInformation, "Citation audit"
End Sub

Public Sub WU_JournalSupraAudit()
    Dim note As Footnote, endnote As Endnote, count As Long
    For Each note In ActiveDocument.Footnotes
        If InStr(1, note.Range.Text, "supra", vbTextCompare) > 0 Then count = count + 1
    Next note
    For Each endnote In ActiveDocument.Endnotes
        If InStr(1, endnote.Range.Text, "supra", vbTextCompare) > 0 Then count = count + 1
    Next endnote
    MsgBox "Notes containing 'supra': " & CStr(count), vbInformation, "Supra audit"
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
    On Error GoTo Failed
    Set doc = ActiveDocument
    report = "Paragraphs: " & CStr(doc.Paragraphs.Count) & vbCrLf
    report = report & "Sections: " & CStr(doc.Sections.Count) & vbCrLf
    report = report & "Footnotes: " & CStr(doc.Footnotes.Count) & vbCrLf
    report = report & "Endnotes: " & CStr(doc.Endnotes.Count) & vbCrLf
    report = report & "Fields (all stories): " & CStr(WU_CountFields(doc)) & vbCrLf
    report = report & "Tables (all stories): " & CStr(WU_CountTables(doc)) & vbCrLf
    report = report & "Revisions (all stories): " & CStr(WU_CountRevisions(doc)) & vbCrLf
    report = report & "Hyperlinks (all stories): " & CStr(WU_CountHyperlinks(doc))
    MsgBox report, vbInformation, "Quality report"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Quality report"
End Sub

Private Function WU_CountFields(ByVal doc As Document) As Long
    Dim firstStory As Range, story As Range
    For Each firstStory In doc.StoryRanges
        Set story = firstStory
        Do While Not story Is Nothing
            On Error Resume Next
            WU_CountFields = WU_CountFields + story.Fields.Count
            Err.Clear
            On Error GoTo 0
            Set story = story.NextStoryRange
        Loop
    Next firstStory
End Function

Private Function WU_CountTables(ByVal doc As Document) As Long
    Dim firstStory As Range, story As Range
    For Each firstStory In doc.StoryRanges
        Set story = firstStory
        Do While Not story Is Nothing
            On Error Resume Next
            WU_CountTables = WU_CountTables + story.Tables.Count
            Err.Clear
            On Error GoTo 0
            Set story = story.NextStoryRange
        Loop
    Next firstStory
End Function

Private Function WU_CountRevisions(ByVal doc As Document) As Long
    Dim firstStory As Range, story As Range
    For Each firstStory In doc.StoryRanges
        Set story = firstStory
        Do While Not story Is Nothing
            On Error Resume Next
            WU_CountRevisions = WU_CountRevisions + story.Revisions.Count
            Err.Clear
            On Error GoTo 0
            Set story = story.NextStoryRange
        Loop
    Next firstStory
End Function

Private Function WU_CountHyperlinks(ByVal doc As Document) As Long
    Dim firstStory As Range, story As Range
    For Each firstStory In doc.StoryRanges
        Set story = firstStory
        Do While Not story Is Nothing
            On Error Resume Next
            WU_CountHyperlinks = WU_CountHyperlinks + story.Hyperlinks.Count
            Err.Clear
            On Error GoTo 0
            Set story = story.NextStoryRange
        Loop
    Next firstStory
End Function

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
    On Error GoTo FailedPerma
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
`, vbaString(profile.ID), vbaString(profile.Name), vbaString(profile.BodyFont), vbaString(profile.NoteFont), profile.BodySizePT, profile.NoteSizePT, vbaString(profile.PermalinkPolicy), vbaString(styleBase+" Body"), vbaString(styleBase+" Note"), vbaString(styleBase+" Heading 1"), vbaString(styleBase+" Heading 2"), vbaString(styleBase+" Heading 3"), profile.Name, profile.Name)
}
