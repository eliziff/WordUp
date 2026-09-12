// Package example creates reproducible native integration fixtures, not mock UI.
package example

import (
	"path/filepath"
	"wordwright.local/internal/compat"
	"wordwright.local/internal/native"
	"wordwright.local/internal/office"
	"wordwright.local/internal/project"
	"wordwright.local/internal/verify"
)

func Studio(root string) (*project.BuildReport, error) {
	if _, e := project.New("WordwrightStudio", root); e != nil {
		return nil, e
	}
	w, e := project.Open(root)
	if e != nil {
		return nil, e
	}
	w.Manifest.Components["StudioForm"] = "form"
	writes := map[string][]byte{"project.json": project.JSON(w.Manifest), "vba/Studio.bas": []byte(source), "vba/PlatformProbe.bas": []byte(compat.ProbeSource()), "vba/StudioForm.vba": []byte(formSource)}
	design := office.Design{Name: "StudioForm", Mode: "replace", Properties: map[string]any{"Caption": "Wordwright Studio", "Width": 360, "Height": 240}, Controls: []office.ControlDesign{{Name: "Pages", Type: "MultiPage", Properties: map[string]any{"Left": 12, "Top": 12, "Width": 330, "Height": 170}, Pages: []office.ControlDesign{{Name: "Operations", Type: "Page", Properties: map[string]any{"Caption": "Text operations"}, Controls: []office.ControlDesign{{Name: "txtInput", Type: "TextBox", Properties: map[string]any{"Left": 12, "Top": 12, "Width": 270, "Height": 28, "Value": "Editable native control"}}, {Name: "cmdNormalize", Type: "CommandButton", Properties: map[string]any{"Left": 12, "Top": 52, "Width": 120, "Height": 24, "Caption": "Normalize selection"}}}}, {Name: "Metadata", Type: "Page", Properties: map[string]any{"Caption": "Metadata"}, Controls: []office.ControlDesign{{Name: "lblNative", Type: "Label", Properties: map[string]any{"Left": 12, "Top": 12, "Width": 270, "Height": 36, "Caption": "This form is stored in the template."}}}}}}, {Name: "cmdClose", Type: "CommandButton", Properties: map[string]any{"Left": 258, "Top": 192, "Width": 84, "Height": 24, "Caption": "Close"}}}}
	writes["forms/StudioForm.json"] = project.JSON(design)
	writes["styles/recipe.json"] = project.JSON(office.StyleRecipe{Styles: []office.StyleSpec{{ID: "StudioTitle", Name: "Studio Title", BasedOn: "Normal", Next: "StudioBody", Run: map[string]any{"font": "Times New Roman", "size_pt": 22, "bold": true}, Paragraph: map[string]any{"after_pt": 12, "keep_next": true}}, {ID: "StudioBody", Name: "Studio Body", BasedOn: "Normal", Run: map[string]any{"font": "Times New Roman", "size_pt": 11}, Paragraph: map[string]any{"after_pt": 8}}, {ID: "StudioNote", Name: "Studio Note", BasedOn: "StudioBody", Run: map[string]any{"size_pt": 10, "italic": true}, Paragraph: map[string]any{"left_pt": 18, "right_pt": 18, "after_pt": 8}}}})
	content := office.ContentRecipe{Page: &office.PageSpec{WidthPT: 612, HeightPT: 792, MarginsPT: map[string]float64{"top": 54, "bottom": 54, "left": 54, "right": 54, "header": 24, "footer": 24}, Header: []office.Block{{Text: "WORDWRIGHT / NATIVE TEMPLATE", Run: map[string]any{"size_pt": 8}}}, Footer: []office.Block{{Inlines: []office.Inline{{Text: "Page "}, {Field: "PAGE", Text: "1"}}, Paragraph: map[string]any{"alignment": "right"}}}}, Blocks: []office.Block{{Text: "Editorial Studio", Style: "StudioTitle"}, {Type: "content_control", Tag: "author", Title: "Author", Blocks: []office.Block{{Text: "Author name", Style: "StudioBody"}}}, {Style: "StudioBody", Inlines: []office.Inline{{Text: "A native template, not a screenshot. Edit the styles, form, Ribbon and VBA as source."}, {Footnote: []office.Block{{Text: "This is a native Word footnote with its own reference marker.", Run: map[string]any{"size_pt": 9}}}}}}, {Type: "table", ColumnsPT: []float64{144, 360}, HeaderRows: 1, Rows: [][]office.Cell{{{Text: "Component"}, {Text: "Editable source"}}, {{Text: "VBA and forms"}, {Text: "Text modules and persistent native MSForms storage"}}, {{Text: "Document design"}, {Text: "Named styles, content controls, footnotes and saved parts"}}}}, {Text: "A reusable editorial note is also saved in Quick Parts.", Style: "StudioNote"}}}
	writes["content/recipe.json"] = project.JSON(content)
	writes["building_blocks/recipe.json"] = project.JSON([]office.BuildingBlock{{Name: "Studio Editorial Note", Category: "Wordwright", Gallery: "autoTxt", Blocks: []office.Block{{Text: "Editorial note: replace with a reusable observation.", Style: "StudioNote"}}}})
	writes["tests/suite.json"] = project.JSON(WindowsSuite())
	writes["tests/mac-suite.json"] = project.JSON(MacSuite())
	for path, b := range writes {
		if e = project.Write(root, path, b, ""); e != nil {
			return nil, e
		}
	}
	// RibbonX is a real package part and relationship, not a mocked web toolbar.
	p, e := office.ReadPackage(w.Baseline.Original)
	if e != nil {
		return nil, e
	}
	p.Files["customUI/customUI14.xml"] = []byte(ribbon)
	if e = p.ContentType("customUI/customUI14.xml", "application/xml"); e != nil {
		return nil, e
	}
	if e = p.Relationship("", "rIdStudioRibbon", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility", "customUI/customUI14.xml", ""); e != nil {
		return nil, e
	}
	for _, path := range []string{"customUI/customUI14.xml", "[Content_Types].xml", "_rels/.rels"} {
		if e = project.Write(root, "package/"+path, p.Files[path], ""); e != nil {
			return nil, e
		}
	}
	w, e = project.Open(root)
	if e != nil {
		return nil, e
	}
	return w.Build(filepath.Join(root, "dist", "WordwrightStudio.dotm"))
}
func WindowsSuite() verify.Suite {
	return verify.Suite{Schema: 1, Name: "Native persisted template acceptance", Platforms: []string{"windows"}, Steps: []verify.Step{
		{Name: "Open exact artifact without repair", Operation: native.Operation{Op: "open", File: "$artifact", As: "template"}, Assert: []verify.Assertion{{Path: "/open_and_repair", Kind: "equals", Expected: false}}},
		{Name: "Run actual VBA and instantiate persistent nested form", Operation: native.Operation{Op: "run", Macro: "$project.Studio.NativeAcceptance"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "WORDWRIGHT_NATIVE_OK"}}},
		{Name: "Verify real Ribbon initialization", Operation: native.Operation{Op: "run", Macro: "$project.Studio.RibbonAcceptance"}, Assert: []verify.Assertion{{Kind: "equals", Expected: true}}},
		{Name: "Observe native document contents", Operation: native.Operation{Op: "get", Target: "template", Member: "Content", As: "body"}},
		{Name: "Check real document text", Operation: native.Operation{Op: "get", Target: "body", Member: "Text"}, Assert: []verify.Assertion{{Kind: "contains", Expected: "Editorial Studio"}}},
		{Name: "Verify actual selection behavior and formatting preservation", Operation: native.Operation{Op: "run", Macro: "$project.Studio.NormalizationAcceptance"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "SELECTION_OK"}}},
		{Name: "Create a document from the finished template", Operation: native.Operation{Op: "new", File: "$artifact", As: "created"}},
		{Name: "Get inherited content controls", Operation: native.Operation{Op: "get", Target: "created", Member: "ContentControls", As: "controls"}},
		{Name: "Assert author content control survived", Operation: native.Operation{Op: "get", Target: "controls", Member: "Count"}, Assert: []verify.Assertion{{Kind: "greater_than", Expected: 0}}},
		{Name: "Insert actual saved building block", Operation: native.Operation{Op: "run", Macro: "$project.Studio.BuildingBlockAcceptance"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "BUILDING_BLOCK_OK"}}},
		{Name: "Close generated document without altering the artifact", Operation: native.Operation{Op: "unload", Target: "created"}},
		{Name: "Release template revision for the next warm iteration", Operation: native.Operation{Op: "unload", Target: "template", File: "$artifact"}},
	}}
}
func MacSuite() verify.Suite {
	return verify.Suite{Schema: 1, Name: "Mac native template acceptance", Platforms: []string{"darwin"}, Steps: []verify.Step{{Name: "Open actual template", Operation: native.Operation{Op: "open", File: "$artifact", As: "template"}}, {Name: "Execute portable native VBA acceptance", Operation: native.Operation{Op: "run", Macro: "$project.Studio.NativeAcceptance"}}, {Name: "Inspect saved observable marker", Operation: native.Operation{Op: "get", Target: "template", Member: "content"}, Assert: []verify.Assertion{{Kind: "contains", Expected: "WORDWRIGHT_NATIVE_OK"}}}}}
}

const formSource = `Attribute VB_Name = "StudioForm"
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdNormalize_Click()
    Studio.NormalizeSelection
End Sub
`
const source = `Attribute VB_Name = "Studio"
Option Explicit
Public RibbonInitialized As Boolean
Public Sub OnRibbonLoad(ByVal ribbon As IRibbonUI)
    RibbonInitialized = True
End Sub
Public Sub ShowStudio(ByVal control As IRibbonControl)
    StudioForm.Show vbModeless
End Sub
Public Sub NormalizeButton(ByVal control As IRibbonControl)
    NormalizeSelection
End Sub
Public Sub NormalizeSelection()
    Dim scope As Range, beforeLength As Long, afterLength As Long
    Set scope = Selection.Range.Duplicate
    If scope.Start = scope.End Then Err.Raise vbObjectError + 700, "Wordwright", "Select the text to normalize."
    Do
        beforeLength = scope.End - scope.Start
        With scope.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "  "
            .Replacement.Text = " "
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
        afterLength = scope.End - scope.Start
    Loop While afterLength < beforeLength
End Sub
Public Function NativeAcceptance() As String
    Dim f As StudioForm, d As Document
    Set d = ActiveDocument
    Set f = New StudioForm
    If f.Pages.Pages.Count <> 2 Then Err.Raise vbObjectError + 701, "Wordwright", "Persisted MultiPage count differs."
    If f.Pages.Pages(0).Controls("txtInput").Value <> "Editable native control" Then Err.Raise vbObjectError + 702, "Wordwright", "Native control property did not persist."
    If Abs(f.Pages.Pages(0).Controls("txtInput").Width - 270) > 1 Then Err.Raise vbObjectError + 703, "Wordwright", "Native geometry differs."
    Unload f
    If d.Styles("Studio Title").Font.Size <> 22 Then Err.Raise vbObjectError + 704, "Wordwright", "Named style differs."
    If d.Footnotes.Count <> 1 Then Err.Raise vbObjectError + 705, "Wordwright", "Footnote was not serialized."
    If d.Tables.Count <> 1 Then Err.Raise vbObjectError + 706, "Wordwright", "Table was not serialized."
#If Mac Then
    ' The native Mac scripting command need not return a function value.
    ' Write an independently observable marker into this disposable test copy.
    d.Content.InsertAfter vbCr & "WORDWRIGHT_NATIVE_OK"
#End If
    NativeAcceptance = "WORDWRIGHT_NATIVE_OK"
End Function
Public Function RibbonAcceptance() As Boolean
    Dim started As Single
    started = Timer
    Do While Not RibbonInitialized
        DoEvents
        If (Timer - started + 86400!) Mod 86400! > 3 Then Exit Do
    Loop
    RibbonAcceptance = RibbonInitialized
End Function
Public Function BuildingBlockAcceptance() As String
    Dim target As Range, block As BuildingBlock
    Set block = ActiveDocument.AttachedTemplate.BuildingBlockEntries("Studio Editorial Note")
    Set target = ActiveDocument.Range(ActiveDocument.Content.End - 1, ActiveDocument.Content.End - 1)
    block.Insert Where:=target, RichText:=True
    If InStr(ActiveDocument.Content.Text, "Editorial note: replace with a reusable observation.") = 0 Then Err.Raise vbObjectError + 713, "Wordwright", "Saved part insertion failed."
    BuildingBlockAcceptance = "BUILDING_BLOCK_OK"
End Function
Public Function NormalizationAcceptance() As String
    Dim d As Document, r As Range
    Set d = Documents.Add
    On Error GoTo Failed
    d.Content.Text = "LEFT  untouched|alpha  beta   gamma|RIGHT  untouched"
    Set r = d.Range(16, 35)
    r.Font.Bold = True
    r.Select
    NormalizeSelection
    If d.Content.Text <> "LEFT  untouched|alpha beta gamma|RIGHT  untouched" & vbCr Then Err.Raise vbObjectError + 710, "Wordwright", "Selection scope or normalization is wrong."
    If d.Range(16, 21).Font.Bold <> True Then Err.Raise vbObjectError + 711, "Wordwright", "Formatting was lost."
    d.Close SaveChanges:=wdDoNotSaveChanges
    NormalizationAcceptance = "SELECTION_OK"
    Exit Function
Failed:
    Dim n As Long, msg As String
    n = Err.Number: msg = Err.Description
    d.Close SaveChanges:=wdDoNotSaveChanges
    Err.Raise n, "Wordwright", msg
End Function
`
const ribbon = `<?xml version="1.0" encoding="UTF-8"?><customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad"><ribbon><tabs><tab id="wwStudio" label="Wordwright"><group id="wwEditing" label="Editorial tools"><button id="wwOpen" label="Editorial Studio" size="large" imageMso="FormControlEditBox" onAction="ShowStudio"/><button id="wwNormalize" label="Normalize selection" imageMso="ReplaceDialog" onAction="NormalizeButton"/></group></tab></tabs></ribbon></customUI>`
