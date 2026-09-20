Attribute VB_Name = "WordUp_Trace"
Option Explicit
Option Private Module

' Opt-in observation. Never exports XML, accepts revisions or changes ranges.
' Export after the action's undo record closes. Failures invalidate evidence,
' not the editing action; Finish raises them outside that action.
Private enabled As Boolean
Private rows As Collection, documents As Collection, stories As Collection
Private pending As Object
Private edits As Collection
Private coverageJSON As String, traceFault As String
Private size As Double

Public Sub WordUp_TraceStart(ByVal coverage As String)
    Set rows = New Collection
    Set documents = New Collection
    Set stories = New Collection
    Set pending = CreateObject("Scripting.Dictionary")
    Set edits = New Collection
    coverageJSON = coverage
    traceFault = ""
    size = 0
    enabled = True
End Sub

Public Function WordUp_TraceQuote(ByVal value As String) As String
    Dim i As Long
    value = Replace(value, "\", "\\")
    value = Replace(value, ChrW(34), "\" & ChrW(34))
    For i = 0 To 31
        value = Replace(value, ChrW(i), "\u" & Right$("0000" & Hex$(i), 4))
    Next
    WordUp_TraceQuote = ChrW(34) & value & ChrW(34)
End Function

Private Function DocumentID(ByVal d As Document) As Long
    Dim i As Long
    For i = 1 To documents.Count
        If documents(i) Is d Then DocumentID = i: Exit Function
    Next
    documents.Add d
    DocumentID = documents.Count
End Function

Private Function StoryID(ByVal r As Range) As Long
    Dim i As Long, anchor As Range, d As Document, other As Document
    Set d = r.Document
    For i = 1 To stories.Count
        Set anchor = stories(i)
        Set other = anchor.Document
        If d Is other Then
            If r.InStory(anchor) Then StoryID = i: Exit Function
        End If
    Next
    Set anchor = r.Duplicate
    stories.Add anchor
    StoryID = stories.Count
End Function

' Word maintains these duplicated anchors as subsequent edits occur. A touch
' is diagnostic, not proof of causality. Collapsed anchors remain explicit.
' ponytail: linear scan of observed writes; index by story only if profiling
' shows this opt-in diagnostic is the bottleneck. Never infer Word shifts.
Private Function TouchesJSON(ByVal stage As String, ByVal r As Range) As String
    Dim item As Variant, anchor As Range, d As Document, other As Document
    Dim value As String
    value = "["
    Set d = r.Document
    For Each item In edits
        If item(1) <> stage Then
            Set anchor = item(2)
            Set other = anchor.Document
            If d Is other Then
                If r.InStory(anchor) Then
                    If anchor.Start <= r.End And anchor.End >= r.Start Then
                        If Len(value) > 1 Then value = value & ","
                        value = value & "{""operation"":" & WordUp_TraceQuote(CStr(item(0))) & _
                            ",""stage"":" & WordUp_TraceQuote(CStr(item(1))) & ",""start"":" & anchor.Start & _
                            ",""end"":" & anchor.End & ",""collapsed"":" & LCase$(CStr(anchor.Start = anchor.End)) & "}"
                    End If
                End If
            End If
        End If
    Next
    TouchesJSON = value & "]"
End Function

' dataMembers is a JSON object body, not a complete object. It holds the exact
' project-specific payload; the shared recorder supplies identity and order.
Public Sub WordUp_TraceEvent(ByVal kind As String, ByVal stage As String, ByVal r As Range, ByVal dataMembers As String, Optional ByVal operation As String = "")
    Dim row As String
    If Not enabled Then Exit Sub
    On Error GoTo Failed
    row = "{""sequence"":" & (rows.Count + 1) & ",""kind"":" & WordUp_TraceQuote(kind) & _
        ",""stage"":" & WordUp_TraceQuote(stage) & ",""document"":" & WordUp_TraceQuote("d" & DocumentID(r.Document)) & _
        ",""story"":" & WordUp_TraceQuote("s" & StoryID(r)) & ",""operation"":" & WordUp_TraceQuote(operation) & _
        ",""data"":{""start"":" & r.Start & ",""end"":" & r.End & ",""story_type"":" & r.StoryType & _
        ",""story_length"":" & r.StoryLength & ",""text"":" & WordUp_TraceQuote(r.Text)
    If Len(dataMembers) > 0 Then row = row & "," & dataMembers
    If kind = "read" Or kind = "plan" Or kind = "write" Then row = row & ",""touches"":" & TouchesJSON(stage, r)
    row = row & "}}"
    size = size + Len(row)
    If size > 30000000 Then Err.Raise vbObjectError + 771, , "Trace buffer exceeds 30 million UTF-16 units"
    rows.Add row
    Exit Sub
Failed:
    traceFault = "Trace observation failed: " & Err.Number & " " & Err.Description
    enabled = False
End Sub

Public Function WordUp_TraceWrite(ByVal stage As String, ByVal r As Range, ByVal property As String, ByVal requestedJSON As String) As String
    Dim key As String
    If Not enabled Then Exit Function
    On Error GoTo Failed
    key = "w" & (rows.Count + 1)
    pending.Add key, stage
    WordUp_TraceEvent "write", stage, r, """property"":" & WordUp_TraceQuote(property) & _
        ",""requested"":" & requestedJSON & ",""italic"":" & r.Italic & ",""highlight"":" & r.HighlightColorIndex & _
        ",""tracking"":" & LCase$(CStr(r.Document.TrackRevisions)) & _
        ",""track_formatting"":" & LCase$(CStr(r.Document.TrackFormatting)), key
    WordUp_TraceWrite = key
    Exit Function
Failed:
    traceFault = "Trace write observation failed: " & Err.Number & " " & Err.Description
    enabled = False
End Function

Public Sub WordUp_TraceResult(ByVal operation As String, ByVal r As Range, ByVal errorNumber As Long)
    Dim anchor As Range
    If Not enabled Then Exit Sub
    On Error GoTo Failed
    WordUp_TraceEvent "result", CStr(pending(operation)), r, """error"":" & errorNumber & _
        ",""italic"":" & r.Italic & ",""highlight"":" & r.HighlightColorIndex, operation
    If errorNumber = 0 Then
        Set anchor = r.Duplicate
        edits.Add Array(operation, CStr(pending(operation)), anchor)
    End If
    pending.Remove operation
    Exit Sub
Failed:
    traceFault = "Trace result observation failed: " & Err.Number & " " & Err.Description
    enabled = False
End Sub

Public Sub WordUp_TraceFinish(ByVal path As String)
    Dim stream As Object, row As Variant, first As Boolean
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    enabled = False
    If rows Is Nothing Then Err.Raise vbObjectError + 772, , "Trace was not started"
    If Len(traceFault) > 0 Then Err.Raise vbObjectError + 773, , traceFault
    If pending.Count <> 0 Then Err.Raise vbObjectError + 774, , "Trace has unfinished writes"
    If Application.UndoRecord.CustomRecordLevel <> 0 Then Err.Raise vbObjectError + 775, , "Export trace after closing undo record"
    Set stream = CreateObject("Scripting.FileSystemObject").CreateTextFile(path, False, True)
    stream.Write "{""schema"":1,""complete"":true,""coverage"":" & coverageJSON & ",""events"":["
    first = True
    For Each row In rows
        If Not first Then stream.Write ","
        stream.Write vbCrLf & CStr(row)
        first = False
    Next
    stream.Write vbCrLf & "]}"
    stream.Close
    WordUp_TraceReset
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    WordUp_TraceReset
    On Error GoTo 0
    Err.Raise errorNumber, "WordUp_TraceFinish", errorText
End Sub

Public Sub WordUp_TraceReset()
    enabled = False
    Set rows = Nothing
    Set documents = Nothing
    Set stories = Nothing
    Set pending = Nothing
    Set edits = Nothing
End Sub
