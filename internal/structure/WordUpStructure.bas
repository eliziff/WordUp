Attribute VB_Name = "WordUpStructure"
Option Explicit

' WordUp structure contract 1.0.0. MIT licensed; editable and dependency-free.
' Detection is separate from publication-specific style mapping.
Public Const WU_ROLE As Long = 0
Public Const WU_LEVEL As Long = 1
Public Const WU_PARENT As Long = 2
Public Const WU_CONFIDENCE As Long = 3
Public Const WU_AMBIGUOUS As Long = 4
Public Const WU_EVIDENCE As Long = 5
Public Const WU_STYLE As Long = 6
Public Const WU_TEXT As Long = 7
Public Const WU_START As Long = 8
Public Const WU_END As Long = 9
Public Const WU_MARKER As Long = 10
Public Const WU_LIST_LABEL As Long = 11
Public Const WU_CONTEXT As Long = 12
Public Const WU_COLUMNS As Long = 13

' Returns result(paragraphIndex, WU_*). Positions are Word UTF-16 story offsets.
' This procedure reads the document and never edits it.
Public Function WU_DetectStructure(ByVal document As Document) As Variant
    Dim paragraphs As Paragraphs, count As Long, result() As Variant
    Dim i As Long, paragraph As Paragraph, scope As Range, text As String
    Dim marker As String, label As String, style As String, context As String
    Dim score As Long, level As Long, outline As Long, evidence As String, parentAt(1 To 9) As Long
    Set paragraphs = document.StoryRanges(wdMainTextStory).Paragraphs
    count = paragraphs.count
    If count > 20000 Then Err.Raise 5, "WU_DetectStructure", "paragraph limit exceeds 20000"
    If count = 0 Then WU_DetectStructure = Array(): Exit Function
    ReDim result(0 To count - 1, 0 To WU_COLUMNS - 1)
    i = 0
    For Each paragraph In paragraphs
        i = i + 1
        Set scope = paragraph.Range
        text = WU_CleanText(scope.text): style = CStr(paragraph.Style.NameLocal)
        marker = WU_ParseMarker(text): label = vbNullString
        If scope.ListFormat.ListType <> wdListNoNumbering Then label = scope.ListFormat.ListString
        context = "body": If scope.Information(wdWithInTable) Then context = "table"
        score = 0: level = 0: evidence = vbNullString
        If context = "body" And Len(Trim$(text)) > 0 Then
            outline = paragraph.OutlineLevel
            If outline >= 1 And outline <= 9 Then
                level = outline: score = score + 100: WU_AddEvidence evidence, "native-outline"
            End If
            If WU_HeadingStyle(style) Then score = score + 35: WU_AddEvidence evidence, "heading-style"
            If Len(marker) > 0 Then score = score + 30: WU_AddEvidence evidence, "marker:" & marker
            If Len(text) <= 160 Then score = score + 8 Else score = score - 25
            If paragraph.KeepWithNext <> 0 Then score = score + 8: WU_AddEvidence evidence, "keep-next"
            If Len(text) <= 160 Or level > 0 Or Len(marker) > 0 Or score >= 35 Then
                ' wdUndefined is mixed formatting, not affirmative evidence.
                If scope.Bold = True Then score = score + 8: WU_AddEvidence evidence, "aggregate-bold"
                If scope.Font.SmallCaps = True Or scope.Font.AllCaps = True Then score = score + 8: WU_AddEvidence evidence, "aggregate-caps"
                If WU_UpperShare(text) >= 0.8 Then score = score + 8: WU_AddEvidence evidence, "uppercase-text"
            End If
            If Len(label) > 0 And Len(marker) = 0 And level = 0 Then score = score - 20: WU_AddEvidence evidence, "ordinary-list-risk"
            If WU_SentenceEnding(text) And level = 0 And Len(marker) = 0 Then score = score - 18: WU_AddEvidence evidence, "sentence-ending"
            WU_CustomizeCandidate paragraph, score, level, evidence
        End If
        result(i - 1, WU_STYLE) = style: result(i - 1, WU_TEXT) = text
        result(i - 1, WU_START) = scope.Start: result(i - 1, WU_END) = scope.End
        result(i - 1, WU_MARKER) = marker: result(i - 1, WU_LIST_LABEL) = label
        result(i - 1, WU_CONTEXT) = context: result(i - 1, WU_EVIDENCE) = evidence
        result(i - 1, WU_CONFIDENCE) = score: result(i - 1, WU_LEVEL) = level
        result(i - 1, WU_PARENT) = 0: result(i - 1, WU_AMBIGUOUS) = False
    Next paragraph
    WU_ResolveStyleFamilies result, count
    WU_ResolveMarkerLadder result, count
    For i = 0 To count - 1
        score = CLng(result(i, WU_CONFIDENCE)): level = CLng(result(i, WU_LEVEL))
        If result(i, WU_CONTEXT) <> "body" Then
            result(i, WU_ROLE) = result(i, WU_CONTEXT)
        ElseIf level > 0 And score >= 35 Then
            result(i, WU_ROLE) = "heading": result(i, WU_PARENT) = WU_FindParent(parentAt, level)
            parentAt(level) = i + 1: WU_ClearDeeper parentAt, level
        ElseIf score >= 25 Then
            result(i, WU_ROLE) = "candidate": result(i, WU_AMBIGUOUS) = True
        Else
            result(i, WU_ROLE) = "body"
        End If
        If score > 100 Then result(i, WU_CONFIDENCE) = 100
        If score < 0 Then result(i, WU_CONFIDENCE) = 0
    Next i
    WU_DetectStructure = result
End Function

' Publication-specific copies may adjust evidence here. Do not apply styles here.
Private Sub WU_CustomizeCandidate(ByVal paragraph As Paragraph, ByRef score As Long, ByRef level As Long, ByRef evidence As String)
End Sub

Private Sub WU_ResolveStyleFamilies(ByRef result As Variant, ByVal count As Long)
    Dim i As Long, slot As Long, level As Long, capacity As Long
    Dim keys() As String, levels() As Long, votes() As Long, conflict() As Boolean
    capacity = count * 2 + 1
    ReDim keys(0 To capacity - 1): ReDim levels(0 To capacity - 1)
    ReDim votes(0 To capacity - 1): ReDim conflict(0 To capacity - 1)
    ' Snapshot only original evidence. Inferred paragraphs never cast votes.
    For i = 0 To count - 1
        level = CLng(result(i, WU_LEVEL))
        If level >= 1 And level <= 9 And result(i, WU_CONTEXT) = "body" Then
            slot = WU_StyleSlot(keys, capacity, CStr(result(i, WU_STYLE)))
            If votes(slot) > 0 And levels(slot) <> level Then conflict(slot) = True
            levels(slot) = level: votes(slot) = votes(slot) + 1
        End If
    Next i
    For i = 0 To count - 1
        If CLng(result(i, WU_LEVEL)) = 0 And result(i, WU_CONTEXT) = "body" And CLng(result(i, WU_CONFIDENCE)) >= 25 Then
            slot = WU_StyleSlot(keys, capacity, CStr(result(i, WU_STYLE)))
            If votes(slot) >= 2 And Not conflict(slot) Then
                result(i, WU_LEVEL) = levels(slot): result(i, WU_CONFIDENCE) = CLng(result(i, WU_CONFIDENCE)) + 30
                result(i, WU_EVIDENCE) = WU_Appended(CStr(result(i, WU_EVIDENCE)), "coherent-style-family")
            End If
        End If
    Next i
End Sub

' Native arrays only: no Scripting.Dictionary, added reference or runtime DLL.
Private Function WU_StyleSlot(ByRef keys() As String, ByVal capacity As Long, ByVal name As String) As Long
    Dim slot As Long, i As Long
    name = LCase$(name)
    For i = 1 To Len(name)
        slot = (slot * 33 + CLng(AscW(Mid$(name, i, 1))) + 65536) Mod capacity
    Next i
    Do While Len(keys(slot)) > 0 And keys(slot) <> name
        slot = (slot + 1) Mod capacity
    Loop
    keys(slot) = name: WU_StyleSlot = slot
End Function

Private Sub WU_ResolveMarkerLadder(ByRef result As Variant, ByVal count As Long)
    Dim family(1 To 9) As String, value(1 To 9) As Long, depth As Long
    Dim i As Long, at As Long, currentFamily As String, currentValue As Long
    For i = 0 To count - 1
        If result(i, WU_CONTEXT) = "body" And Len(CStr(result(i, WU_MARKER))) > 0 Then
            currentFamily = WU_MarkerFamily(CStr(result(i, WU_MARKER)))
            currentValue = WU_MarkerValue(CStr(result(i, WU_MARKER)), currentFamily): at = 0
            For depth = 1 To 9
                If family(depth) = currentFamily And value(depth) + 1 = currentValue Then at = depth: Exit For
            Next depth
            If at = 0 And currentValue = 1 Then
                For depth = 1 To 9: If Len(family(depth)) = 0 Then at = depth: Exit For
                Next depth
            End If
            If at > 0 Then
                family(at) = currentFamily: value(at) = currentValue
                For depth = at + 1 To 9: family(depth) = vbNullString: value(depth) = 0: Next depth
                If CLng(result(i, WU_LEVEL)) = 0 Then result(i, WU_LEVEL) = at
                result(i, WU_CONFIDENCE) = CLng(result(i, WU_CONFIDENCE)) + 15
                result(i, WU_EVIDENCE) = WU_Appended(CStr(result(i, WU_EVIDENCE)), "sequence-ladder")
            Else
                result(i, WU_AMBIGUOUS) = True
                result(i, WU_EVIDENCE) = WU_Appended(CStr(result(i, WU_EVIDENCE)), "sequence-violation")
            End If
        End If
    Next i
End Sub

Public Function WU_ParseMarker(ByVal text As String) As String
    Dim at As Long, prefix As String, rest As String, i As Long, c As String
    text = Trim$(Replace(Replace(text, vbTab, " "), ChrW(160), " "))
    at = InStr(text, "."): If at < 2 Or at > 8 Then Exit Function
    prefix = Left$(text, at - 1): rest = Mid$(text, at + 1)
    If Left$(rest, 1) <> " " Or Len(Trim$(rest)) = 0 Then Exit Function
    If IsNumeric(prefix) Then WU_ParseMarker = prefix: Exit Function
    If Len(prefix) = 1 And prefix >= "A" And prefix <= "Z" Then WU_ParseMarker = prefix: Exit Function
    For i = 1 To Len(prefix): c = Mid$(prefix, i, 1): If InStr(1, "IVXLCDM", c, vbBinaryCompare) = 0 Then Exit Function
    Next i
    WU_ParseMarker = prefix
End Function

Private Function WU_MarkerFamily(ByVal marker As String) As String
    If IsNumeric(marker) Then
        WU_MarkerFamily = "decimal"
    ElseIf Len(marker) = 1 And InStr(1, "IVXLCDM", marker, vbBinaryCompare) = 0 Then
        WU_MarkerFamily = "alpha"
    Else
        WU_MarkerFamily = "roman"
    End If
End Function

Private Function WU_MarkerValue(ByVal marker As String, ByVal family As String) As Long
    If family = "decimal" Then WU_MarkerValue = CLng(marker): Exit Function
    If family = "alpha" Then WU_MarkerValue = AscW(marker) - 64: Exit Function
    Dim i As Long, n As Long, last As Long, current As Long
    For i = Len(marker) To 1 Step -1
        current = InStr(1, "IVXLCDM", Mid$(marker, i, 1), vbBinaryCompare)
        current = Choose(current, 1, 5, 10, 50, 100, 500, 1000)
        If current < last Then n = n - current Else n = n + current: last = current
    Next i
    WU_MarkerValue = n
End Function

Private Function WU_HeadingStyle(ByVal name As String) As Boolean
    name = LCase$(name): WU_HeadingStyle = InStr(name, "heading") > 0 Or InStr(name, "title") > 0 Or InStr(name, "head") > 0
End Function

Private Function WU_CleanText(ByVal text As String) As String
    If Right$(text, 1) = vbCr Then text = Left$(text, Len(text) - 1)
    WU_CleanText = text
End Function

Private Function WU_UpperShare(ByVal text As String) As Double
    Dim i As Long, letters As Long, upper As Long, c As String
    For i = 1 To Len(text)
        c = Mid$(text, i, 1)
        If LCase$(c) <> UCase$(c) Then letters = letters + 1: If c = UCase$(c) Then upper = upper + 1
    Next i
    If letters > 0 Then WU_UpperShare = upper / letters
End Function

Private Function WU_SentenceEnding(ByVal text As String) As Boolean
    text = Trim$(text): If Len(text) = 0 Then Exit Function
    WU_SentenceEnding = InStr(1, ".,;:!?", Right$(text, 1), vbBinaryCompare) > 0
End Function

Private Sub WU_AddEvidence(ByRef evidence As String, ByVal item As String): evidence = WU_Appended(evidence, item): End Sub
Private Function WU_Appended(ByVal evidence As String, ByVal item As String) As String
    If Len(evidence) = 0 Then WU_Appended = item Else WU_Appended = evidence & ";" & item
End Function
Private Function WU_FindParent(ByRef parents() As Long, ByVal level As Long) As Long
    Dim i As Long: For i = level - 1 To 1 Step -1: If parents(i) > 0 Then WU_FindParent = parents(i): Exit Function
    Next i
End Function
Private Sub WU_ClearDeeper(ByRef parents() As Long, ByVal level As Long)
    Dim i As Long: For i = level + 1 To 9: parents(i) = 0: Next i
End Sub
