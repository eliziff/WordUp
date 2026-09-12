Attribute VB_Name = "Studio"
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
