Attribute VB_Name = "Studio"
Option Explicit
Public RibbonInitialized As Boolean
Public FormClosed As Boolean
Public CommandInvocations As Long
Private StudioRibbon As IRibbonUI
Public Sub OnRibbonLoad(ByVal ribbon As IRibbonUI)
    RibbonInitialized = True
    Set StudioRibbon = ribbon
End Sub
Public Sub ActivateStudioRibbon()
    StudioRibbon.ActivateTab "wwStudio"
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
    Set scope = Selection.Range.Duplicate
    If scope.Start = scope.End Then Err.Raise vbObjectError + 700, "WordUp", "Select the text to normalize."
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
    Dim d As Document, r As Range
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
    NormalizeSelection
    If d.Content.Text <> "LEFT  untouched|alpha beta gamma|RIGHT  untouched" & vbCr Then Err.Raise vbObjectError + 710, "WordUp", "Selection scope or normalization is wrong."
    If d.Range(16, 21).Font.Bold <> True Then Err.Raise vbObjectError + 711, "WordUp", "Formatting was lost."
    If Not reuseDocument Then d.Close SaveChanges:=wdDoNotSaveChanges
    NormalizationAcceptance = "SELECTION_OK"
    Exit Function
Failed:
    Dim n As Long, msg As String
    n = Err.Number: msg = Err.Description
    If Not reuseDocument Then d.Close SaveChanges:=wdDoNotSaveChanges
    Err.Raise n, "WordUp", msg
End Function
