// Package example creates reproducible native integration fixtures, not mock UI.
package example

import (
	"github.com/eliziff/WordUp/internal/compat"
	"github.com/eliziff/WordUp/internal/native"
	"github.com/eliziff/WordUp/internal/office"
	"github.com/eliziff/WordUp/internal/project"
	"github.com/eliziff/WordUp/internal/verify"
	"path/filepath"
)

func Studio(root string) (*project.BuildReport, error) {
	if _, e := project.New("WordUpStudio", root); e != nil {
		return nil, e
	}
	w, e := project.Open(root)
	if e != nil {
		return nil, e
	}
	writes := map[string][]byte{"project.json": project.JSON(w.Manifest), "vba/Studio.bas": []byte(source), "vba/PlatformProbe.bas": []byte(compat.ProbeSource()), "vba/StudioForm.vba": []byte(formSource)}
	design := office.Design{Name: "StudioForm", Mode: "replace", Properties: map[string]any{"Caption": "WordUp Studio", "Width": 360, "Height": 240}, Controls: []office.ControlDesign{{Name: "Pages", Type: "MultiPage", Properties: map[string]any{"Left": 12, "Top": 12, "Width": 330, "Height": 170}, Pages: []office.ControlDesign{{Name: "Operations", Type: "Page", Properties: map[string]any{"Caption": "Text operations"}, Controls: []office.ControlDesign{{Name: "txtInput", Type: "TextBox", Properties: map[string]any{"Left": 12, "Top": 12, "Width": 270, "Height": 28, "Value": "Editable native control — café"}}, {Name: "cmdNormalize", Type: "CommandButton", Properties: map[string]any{"Left": 12, "Top": 52, "Width": 120, "Height": 24, "Caption": "Normalize selection"}}}}, {Name: "Metadata", Type: "Page", Properties: map[string]any{"Caption": "Metadata"}, Controls: []office.ControlDesign{{Name: "lblNative", Type: "Label", Properties: map[string]any{"Left": 12, "Top": 12, "Width": 270, "Height": 36, "Caption": "This form is stored in the template."}}}}}}, {Name: "cmdClose", Type: "CommandButton", Properties: map[string]any{"Left": 258, "Top": 192, "Width": 84, "Height": 24, "Caption": "Close"}}}}
	writes["forms/StudioForm.json"] = project.JSON(design)
	styles := office.StyleRecipe{Styles: []office.StyleSpec{{ID: "StudioTitle", Name: "Studio Title", BasedOn: "Normal", Next: "StudioBody", Run: map[string]any{"font": "Times New Roman", "size_pt": 22, "bold": true}, Paragraph: map[string]any{"after_pt": 12, "keep_next": true}}, {ID: "StudioBody", Name: "Studio Body", BasedOn: "Normal", Run: map[string]any{"font": "Times New Roman", "size_pt": 11}, Paragraph: map[string]any{"after_pt": 8}}, {ID: "StudioNote", Name: "Studio Note", BasedOn: "StudioBody", Run: map[string]any{"size_pt": 10, "italic": true}, Paragraph: map[string]any{"left_pt": 18, "right_pt": 18, "after_pt": 8}}}}
	content := office.ContentRecipe{Page: &office.PageSpec{WidthPT: 612, HeightPT: 792, MarginsPT: map[string]float64{"top": 54, "bottom": 54, "left": 54, "right": 54, "header": 24, "footer": 24}, Header: []office.Block{{Text: "WORDUP / NATIVE TEMPLATE", Run: map[string]any{"size_pt": 8}}}, Footer: []office.Block{{Inlines: []office.Inline{{Text: "Page "}, {Field: "PAGE", Text: "1"}}, Paragraph: map[string]any{"alignment": "right"}}}}, Blocks: []office.Block{{Text: "Editorial Studio", Style: "StudioTitle"}, {Type: "content_control", Tag: "author", Title: "Author", Blocks: []office.Block{{Text: "Author name", Style: "StudioBody"}}}, {Style: "StudioBody", Inlines: []office.Inline{{Text: "A native template, not a screenshot. Edit the styles, form, Ribbon and VBA as source."}, {Footnote: []office.Block{{Text: "This is a native Word footnote with its own reference marker.", Run: map[string]any{"size_pt": 9}}}}}}, {Type: "table", ColumnsPT: []float64{144, 360}, HeaderRows: 1, Rows: [][]office.Cell{{{Text: "Component"}, {Text: "Editable source"}}, {{Text: "VBA and forms"}, {Text: "Text modules and persistent native MSForms storage"}}, {{Text: "Document design"}, {Text: "Named styles, content controls, footnotes and saved parts"}}}}, {Text: "A reusable editorial note is also saved in Quick Parts.", Style: "StudioNote"}}}
	blocks := []office.BuildingBlock{{Name: "Studio Editorial Note", Category: "WordUp", Gallery: "autoTxt", Blocks: []office.Block{{Text: "Editorial note: replace with a reusable observation.", Style: "StudioNote"}}}}
	writes["tests/suite.json"] = project.JSON(WindowsSuite())
	writes["tests/mac-suite.json"] = project.JSON(MacSuite())
	for path, b := range writes {
		if e = project.Write(root, path, b, ""); e != nil {
			return nil, e
		}
	}
	// Materialize the fixture as ordinary package source. The workspace build has
	// no recipe format or second interpretation of the desired document.
	p, e := office.ReadPackage(w.Baseline.Original)
	if e != nil {
		return nil, e
	}
	if e = office.ApplyStyles(p, styles); e != nil {
		return nil, e
	}
	if e = office.Compose(p, content, nil); e != nil {
		return nil, e
	}
	if e = office.AddBuildingBlocks(p, blocks, nil); e != nil {
		return nil, e
	}
	if e = p.MergeRibbon("customUI/customUI14.xml", []byte(ribbon)); e != nil {
		return nil, e
	}
	for path, data := range p.Files {
		if office.Hash(data) == office.Hash(w.Baseline.Files[path]) {
			continue
		}
		if e = project.Write(root, "package/"+path, data, ""); e != nil {
			return nil, e
		}
	}
	w, e = project.Open(root)
	if e != nil {
		return nil, e
	}
	return w.Build(filepath.Join(root, "dist", "WordUpStudio.dotm"))
}
func WindowsSuite() verify.Suite {
	return verify.Suite{Schema: 1, Name: "Native persisted template acceptance", Platforms: []string{"windows"}, RequireCompile: true, Steps: []verify.Step{
		{Name: "Open exact artifact without repair", Operation: native.Operation{Op: "open", File: "$artifact", As: "template"}, Assert: []verify.Assertion{{Path: "/open_and_repair", Kind: "equals", Expected: false}}},
		{Name: "Compile the complete generated project", Operation: native.Operation{Op: "compile", Target: "template", Member: "$project"}, Assert: []verify.Assertion{{Path: "/vba_compiled", Kind: "equals", Expected: true}}},
		{Name: "Return from compiler to Word", Operation: native.Operation{Op: "put", Target: "app", Member: "ShowVisualBasicEditor", Value: false}},
		{Name: "Run actual VBA and instantiate persistent nested form", Operation: native.Operation{Op: "run", Macro: "Studio.NativeAcceptance"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "WORDUP_NATIVE_OK"}}},
		{Name: "Verify real Ribbon initialization", Operation: native.Operation{Op: "run", Macro: "Studio.RibbonAcceptance"}, Assert: []verify.Assertion{{Kind: "equals", Expected: true}}},
		{Name: "Execute arbitrary scratch VBA", Operation: native.Operation{Op: "eval", Value: "Evaluate = 6 * 7"}, Assert: []verify.Assertion{{Path: "/result", Kind: "equals", Expected: 42}}},
		{Name: "Render the actual Word document", Operation: native.Operation{Op: "render", Target: "template", File: "$output/pages"}, Assert: []verify.Assertion{{Path: "/page_count", Kind: "greater_than", Expected: 0}}},
		{Name: "Close the template editor before testing an attached template", Operation: native.Operation{Op: "unload", Target: "template"}},
		{Name: "Create a disposable document for the modeless form", Operation: native.Operation{Op: "new", File: "$artifact", As: "formDocument"}},
		{Name: "Click the actual custom Ribbon tab", Operation: native.Operation{Op: "ui.invoke", Target: "formDocument", Named: map[string]any{"name": "WordUp", "scope": "ribbon", "role": 37, "wait_ms": 2000}}},
		{Name: "Click the real Ribbon button", Operation: native.Operation{Op: "ui.invoke", Target: "formDocument", Named: map[string]any{"name": "Editorial Studio", "scope": "ribbon", "role": 43}}, Assert: []verify.Assertion{{Path: "/executed", Kind: "equals", Expected: true}}},
		{Name: "Wait for the UserForm before capturing it", Operation: native.Operation{Op: "ui.find", Named: map[string]any{"name": "Close", "window": "WordUp Studio", "scope": "form", "role": 43, "wait_ms": 2000}}},
		{Name: "Capture Word and the persistent UserForm", Operation: native.Operation{Op: "ui.diagnostics", File: "$output/interface", Named: map[string]any{"trees": false}}},
		{Name: "Change the selection after the modeless form opens", Operation: native.Operation{Op: "eval", Value: `ActiveDocument.Content.Text = "LEFT  untouched|alpha  beta   gamma|RIGHT  untouched"
ActiveDocument.Range(16, 35).Font.Bold = True
ActiveDocument.Range(16, 35).Select
Evaluate = True`}},
		{Name: "Click the nested form editing button", Operation: native.Operation{Op: "ui.invoke", Named: map[string]any{"name": "Normalize selection", "window": "WordUp Studio", "scope": "form", "role": 43}}},
		{Name: "Assert callback output, formatting and caller state", Operation: native.Operation{Op: "eval", Value: `Evaluate = Array(ActiveDocument.Content.Text, ActiveDocument.Range(16, 21).Font.Bold, Application.ScreenUpdating, Application.UndoRecord.CustomRecordLevel)`}, Assert: []verify.Assertion{{Path: "/result/array/0", Kind: "equals", Expected: "LEFT  untouched|alpha beta gamma|RIGHT  untouched\r"}, {Path: "/result/array/1", Kind: "equals", Expected: -1}, {Path: "/result/array/2", Kind: "equals", Expected: true}, {Path: "/result/array/3", Kind: "equals", Expected: 0}}},
		{Name: "Retain exact XML after the form edit", Operation: native.Operation{Op: "xml.snapshot", Target: "formDocument", File: "$output/form-edit.xml"}, Assert: []verify.Assertion{{Path: "/normalized", Kind: "equals", Expected: false}}},
		{Name: "Undo the complete form action once", Operation: native.Operation{Op: "eval", Value: `Dim restored As Boolean
restored = ActiveDocument.Undo
Evaluate = restored And ActiveDocument.Content.Text = "LEFT  untouched|alpha  beta   gamma|RIGHT  untouched" & vbCr And ActiveDocument.Range(16, 35).Font.Bold = True
Application.ScreenUpdating = True`}, Assert: []verify.Assertion{{Path: "/result", Kind: "equals", Expected: true}}},
		{Name: "Click the real form Close button", Operation: native.Operation{Op: "ui.invoke", Named: map[string]any{"name": "Close", "window": "WordUp Studio", "scope": "form", "role": 43}}, Assert: []verify.Assertion{{Path: "/executed", Kind: "equals", Expected: true}}},
		{Name: "Observe the form event callback", Operation: native.Operation{Op: "run", Macro: "Studio.FormEventAcceptance"}, Assert: []verify.Assertion{{Kind: "equals", Expected: true}}},
		{Name: "Discard the form test document", Operation: native.Operation{Op: "unload", Target: "formDocument"}},
		{Name: "Reopen the artifact for independent content checks", Operation: native.Operation{Op: "open", File: "$artifact", As: "template"}},
		{Name: "Register and execute a template hotkey", Operation: native.Operation{Op: "run", Macro: "Studio.HotkeyAcceptance"}, Assert: []verify.Assertion{{Kind: "equals", Expected: true}}},
		{Name: "Prepare the context-menu selection", Operation: native.Operation{Op: "run", Macro: "Studio.ShowContextMenu"}},
		{Name: "Open the actual text context menu", Operation: native.Operation{Op: "context_menu", Target: "template"}},
		{Name: "Click the custom context-menu command", Operation: native.Operation{Op: "ui.invoke", Named: map[string]any{"name": "WordUp test command", "scope": "menu", "wait_ms": 2000}}, Assert: []verify.Assertion{{Path: "/executed", Kind: "equals", Expected: true}}},
		{Name: "Observe context-menu callback", Operation: native.Operation{Op: "run", Macro: "Studio.ContextAcceptance"}, Assert: []verify.Assertion{{Kind: "equals", Expected: true}}},
		{Name: "Observe native document contents", Operation: native.Operation{Op: "get", Target: "template", Member: "Content", As: "body"}},
		{Name: "Check real document text", Operation: native.Operation{Op: "get", Target: "body", Member: "Text"}, Assert: []verify.Assertion{{Kind: "contains", Expected: "Editorial Studio"}}},
		{Name: "Create a document from the finished template", Operation: native.Operation{Op: "new", File: "$artifact", As: "created"}},
		{Name: "Get inherited content controls", Operation: native.Operation{Op: "get", Target: "created", Member: "ContentControls", As: "controls"}},
		{Name: "Assert author content control survived", Operation: native.Operation{Op: "get", Target: "controls", Member: "Count"}, Assert: []verify.Assertion{{Kind: "greater_than", Expected: 0}}},
		{Name: "Insert actual saved building block", Operation: native.Operation{Op: "run", Macro: "Studio.BuildingBlockAcceptance"}, Assert: []verify.Assertion{{Kind: "equals", Expected: "BUILDING_BLOCK_OK"}}},
		{Name: "Verify actual selection behavior and formatting preservation", Operation: native.Operation{Op: "run", Macro: "Studio.NormalizationAcceptance", Args: []any{true}}, Assert: []verify.Assertion{{Kind: "equals", Expected: "SELECTION_OK"}}},
		{Name: "Close generated document without altering the artifact", Operation: native.Operation{Op: "unload", Target: "created"}},
		{Name: "Release template revision for the next warm iteration", Operation: native.Operation{Op: "unload", Target: "template", File: "$artifact"}},
	}}
}
func MacSuite() verify.Suite {
	return verify.Suite{Schema: 1, Name: "Mac native template acceptance", Platforms: []string{"darwin"}, Steps: []verify.Step{{Name: "Open actual template", Operation: native.Operation{Op: "open", File: "$artifact", As: "template"}}, {Name: "Execute portable native VBA acceptance", Operation: native.Operation{Op: "run", Macro: "Studio.NativeAcceptance"}}, {Name: "Inspect saved observable marker", Operation: native.Operation{Op: "get", Target: "template", Member: "content"}, Assert: []verify.Assertion{{Kind: "contains", Expected: "WORDUP_NATIVE_OK"}}}}}
}

const formSource = `Attribute VB_Name = "StudioForm"
Option Explicit
Private Sub cmdClose_Click()
    Studio.FormClosed = True
    Unload Me
End Sub
Private Sub cmdNormalize_Click()
    Studio.NormalizeSelection
End Sub
`
const source = `Attribute VB_Name = "Studio"
Option Explicit
Public RibbonInitialized As Boolean
Public FormClosed As Boolean
Public CommandInvocations As Long
Public Sub OnRibbonLoad(ByVal ribbon As IRibbonUI)
    RibbonInitialized = True
End Sub
Public Function FormEventAcceptance() As Boolean
    FormEventAcceptance = FormClosed
End Function
Public Sub Ping()
    CommandInvocations = CommandInvocations + 1
End Sub
Public Sub ContextPing(ByVal control As IRibbonControl)
    Ping
End Sub
Public Function HotkeyAcceptance() As Boolean
    Dim binding As KeyBinding, previousCount As Long
    CustomizationContext = ThisDocument
    Set binding = KeyBindings.Add(wdKeyCategoryMacro, "Studio.Ping", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyU))
    previousCount = CommandInvocations
    FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyU)).Execute
    HotkeyAcceptance = CommandInvocations = previousCount + 1
    binding.Clear
End Function
Public Sub ShowContextMenu()
    CommandInvocations = 0
    ' RibbonX ContextMenuText is exposed only for a text selection. Keep the
    ' fixture deterministic while exercising the same menu a user sees.
    ActiveDocument.Range(0, 1).Select
End Sub
Public Function ContextAcceptance() As Boolean
    ContextAcceptance = CommandInvocations = 1
End Function
Public Sub ShowStudio(ByVal control As IRibbonControl)
    StudioForm.Show vbModeless
End Sub
Public Sub NormalizeButton(ByVal control As IRibbonControl)
    NormalizeSelection
End Sub
Public Sub NormalizeSelection()
    Dim scope As Range, beforeLength As Long, afterLength As Long
    Dim updating As Boolean, undoStarted As Boolean
    Dim failure As Long, failureSource As String, failureDescription As String
    Set scope = Selection.Range.Duplicate
    If scope.Start = scope.End Then Err.Raise vbObjectError + 700, "WordUp", "Select the text to normalize."
    updating = Application.ScreenUpdating
    On Error GoTo Failed
    Application.UndoRecord.StartCustomRecord "Normalize selection"
    undoStarted = True
    Application.ScreenUpdating = False
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
Cleanup:
    On Error Resume Next
    If undoStarted Then Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = updating
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureDescription
    Exit Sub
Failed:
    failure = Err.Number: failureSource = Err.Source: failureDescription = Err.Description
    Resume Cleanup
End Sub
Public Function NativeAcceptance() As String
    Dim f As StudioForm, d As Document
    Dim stage As String
    On Error GoTo Failed
    Set d = ActiveDocument
    stage = "instantiate form"
    Set f = New StudioForm
    stage = "page count"
    If f.Pages.Pages.Count <> 2 Then Err.Raise vbObjectError + 701, "WordUp", "Persisted MultiPage count differs."
    Dim ctl As Object, pg As Object
    stage = "page textbox value; persisted controls: "
    For Each pg In f.Pages.Pages
        stage = stage & "[" & pg.Name & ":"
        For Each ctl In pg.Controls
            stage = stage & ctl.Name & ","
        Next ctl
        stage = stage & "]"
    Next pg
    If f.Pages.Pages(0).Controls("txtInput").Value <> "Editable native control — café" Then Err.Raise vbObjectError + 702, "WordUp", "Native control property did not persist."
    If Abs(f.Pages.Pages(0).Controls("txtInput").Width - 270) > 1 Then Err.Raise vbObjectError + 703, "WordUp", "Native geometry differs."
    Unload f
    If d.Styles("Studio Title").Font.Size <> 22 Then Err.Raise vbObjectError + 704, "WordUp", "Named style differs."
    If d.Footnotes.Count <> 1 Then Err.Raise vbObjectError + 705, "WordUp", "Footnote was not serialized."
    If d.Tables.Count <> 1 Then Err.Raise vbObjectError + 706, "WordUp", "Table was not serialized."
#If Mac Then
    ' The native Mac scripting command need not return a function value.
    ' Write an independently observable marker into this disposable test copy.
    d.Content.InsertAfter vbCr & "WORDUP_NATIVE_OK"
#End If
    NativeAcceptance = "WORDUP_NATIVE_OK"
    Exit Function
Failed:
    NativeAcceptance = stage & ": " & Err.Description
    On Error Resume Next
    Unload f
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
    On Error GoTo Failed
    Dim target As Range, block As BuildingBlock, context As String
    context = ActiveDocument.AttachedTemplate.FullName & "; entries=" & ActiveDocument.AttachedTemplate.BuildingBlockEntries.Count & "; "
    Set block = ActiveDocument.AttachedTemplate.BuildingBlockEntries("Studio Editorial Note")
    Set target = ActiveDocument.Range(ActiveDocument.Content.End - 1, ActiveDocument.Content.End - 1)
    block.Insert Where:=target, RichText:=True
    If InStr(ActiveDocument.Content.Text, "Editorial note: replace with a reusable observation.") = 0 Then Err.Raise vbObjectError + 713, "WordUp", "Saved part insertion failed."
    BuildingBlockAcceptance = "BUILDING_BLOCK_OK"
    Exit Function
Failed:
    BuildingBlockAcceptance = context & Err.Description
End Function
Public Function NormalizationAcceptance(Optional ByVal reuseDocument As Boolean = False) As String
    Dim d As Document, r As Range, updating As Boolean
    updating = Application.ScreenUpdating
    If reuseDocument Then
        Set d = ActiveDocument
    Else
        Set d = Documents.Add
    End If
    On Error GoTo Failed
    d.Content.Text = "LEFT  untouched|alpha  beta   gamma|RIGHT  untouched"
    Set r = d.Range(16, 35)
    r.Font.Bold = True
    r.Select
    Application.ScreenUpdating = False
    NormalizeSelection
    If Application.ScreenUpdating Then Err.Raise vbObjectError + 714, "WordUp", "Caller screen-updating state was not preserved."
    Application.ScreenUpdating = updating
    If d.Content.Text <> "LEFT  untouched|alpha beta gamma|RIGHT  untouched" & vbCr Then Err.Raise vbObjectError + 710, "WordUp", "Selection scope or normalization is wrong."
    If d.Range(16, 21).Font.Bold <> True Then Err.Raise vbObjectError + 711, "WordUp", "Formatting was lost."
    ' A rejected edit must also close its undo record and restore caller state.
    Dim rejected As Long
    d.Content.Text = "protected  text"
    d.Content.Select
    d.Protect Type:=wdAllowOnlyReading, NoReset:=True
    Application.ScreenUpdating = False
    On Error Resume Next
    NormalizeSelection
    rejected = Err.Number
    Err.Clear
    On Error GoTo Failed
    d.Unprotect
    If rejected = 0 Or rejected = vbObjectError + 700 Then Err.Raise vbObjectError + 715, "WordUp", "Protected edit did not exercise the editing failure path."
    If d.Content.Text <> "protected  text" & vbCr Then Err.Raise vbObjectError + 717, "WordUp", "Rejected edit changed protected text."
    If Application.ScreenUpdating Or Application.UndoRecord.CustomRecordLevel <> 0 Then Err.Raise vbObjectError + 716, "WordUp", "Failed edit left application state changed."
    Application.ScreenUpdating = updating
    If Not reuseDocument Then d.Close SaveChanges:=wdDoNotSaveChanges
    NormalizationAcceptance = "SELECTION_OK"
    Exit Function
Failed:
    Dim n As Long, msg As String
    n = Err.Number: msg = Err.Description
    Application.ScreenUpdating = updating
    If Not reuseDocument Then d.Close SaveChanges:=wdDoNotSaveChanges
    Err.Raise n, "WordUp", msg
End Function
`
const ribbon = `<?xml version="1.0" encoding="UTF-8"?><customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad"><ribbon><tabs><tab id="wwStudio" label="WordUp"><group id="wwEditing" label="Editorial tools"><button id="wwOpen" label="Editorial Studio" size="large" imageMso="FormControlEditBox" onAction="ShowStudio"/><button id="wwNormalize" label="Normalize selection" imageMso="ReplaceDialog" onAction="NormalizeButton"/></group></tab></tabs></ribbon><contextMenus><contextMenu idMso="ContextMenuText"><button id="wwContextPing" label="WordUp test command" onAction="ContextPing"/></contextMenu></contextMenus></customUI>`
