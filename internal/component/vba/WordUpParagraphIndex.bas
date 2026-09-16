Attribute VB_Name = "WordUpParagraphIndex"
Option Explicit
' One-shot paragraph index for the main story.
'
' Batch typesetting macros that ask Word for Range.Tables.Count, Range.Style
' and Range.Text on every paragraph spend most of their time in those calls.
' WU_IndexParagraphs enumerates Paragraphs once (start, end, style name),
' reads the body text once and derives table membership from Document.Tables,
' so a stage can decide on arrays and touch the object model only where it
' writes. Positions are story offsets valid until the document is edited;
' apply edits from the end of the document backwards, or re-index.
'
'   count = WU_IndexParagraphs(doc, starts, ends, styles, inTable, texts)
'
' Every array is redimensioned 1 To count. texts hold each paragraph's text
' including its paragraph mark. A paragraph whose style Word cannot read has
' an empty style name.

Public Function WU_IndexParagraphs(ByVal document As Document, ByRef starts() As Long, ByRef ends() As Long, ByRef styles() As String, ByRef inTable() As Boolean, ByRef texts() As String) As Long
    Dim para As Paragraph, i As Long, bodyText As String, aligned As Boolean, base As Long, count As Long
    Dim tbl As Table, t As Long, tableCount As Long, tStarts() As Long, tEnds() As Long
    If document Is Nothing Then Err.Raise 91, "WU_IndexParagraphs", "document is required"
    count = document.Paragraphs.Count
    If count = 0 Then Exit Function
    ReDim starts(1 To count)
    ReDim ends(1 To count)
    ReDim styles(1 To count)
    ReDim inTable(1 To count)
    ReDim texts(1 To count)
    base = document.Content.Start
    bodyText = document.Content.Text
    ' Field codes and other hidden structure can make story positions and the
    ' visible text disagree; only slice the bulk text when the lengths match.
    aligned = (Len(bodyText) = document.Content.End - base)
    For Each para In document.Paragraphs
        i = i + 1
        If i > count Then Exit For
        starts(i) = para.Range.Start
        ends(i) = para.Range.End
        On Error Resume Next
        styles(i) = para.Style.NameLocal
        If Err.Number <> 0 Then styles(i) = "": Err.Clear
        On Error GoTo 0
        If aligned Then
            texts(i) = Mid$(bodyText, starts(i) - base + 1, ends(i) - starts(i))
        Else
            texts(i) = para.Range.Text
        End If
    Next para
    count = i
    tableCount = document.Tables.Count
    If tableCount > 0 Then
        ReDim tStarts(1 To tableCount)
        ReDim tEnds(1 To tableCount)
        For Each tbl In document.Tables
            t = t + 1
            tStarts(t) = tbl.Range.Start
            tEnds(t) = tbl.Range.End
        Next tbl
        For i = 1 To count
            For t = 1 To tableCount
                If starts(i) >= tStarts(t) And starts(i) < tEnds(t) Then
                    inTable(i) = True
                    Exit For
                End If
            Next t
        Next i
    End If
    WU_IndexParagraphs = count
End Function
