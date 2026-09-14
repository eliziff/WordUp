Attribute VB_Name = "WordUpStructure"
Option Explicit

' WordUp structure contract 1.2.1. MIT licensed; editable and dependency-free.
' Detection is separate from publication-specific style mapping.
Public Const WU_ROLE As Long = 0
Public Const WU_LEVEL As Long = 1
Public Const WU_PARENT As Long = 2
Public Const WU_CANDIDATE_SCORE As Long = 3
Public Const WU_CONFIDENCE As Long = 4
Public Const WU_AMBIGUOUS As Long = 5
Public Const WU_EVIDENCE As Long = 6
Public Const WU_STYLE As Long = 7
Public Const WU_TEXT As Long = 8
Public Const WU_START As Long = 9
Public Const WU_END As Long = 10
Public Const WU_MARKER As Long = 11
Public Const WU_LIST_LABEL As Long = 12
Public Const WU_CONTEXT As Long = 13
Public Const WU_SOURCE_ID As Long = 14
Public Const WU_PARENT_SOURCE_ID As Long = 15
Public Const WU_PROVENANCE As Long = 16
Public Const WU_CONTRADICTION As Long = 17
Public Const WU_ALTERNATIVES As Long = 18
Public Const WU_COLUMNS As Long = 19

' Returns result(paragraphIndex, WU_*). Positions are Word UTF-16 story offsets.
' This procedure reads the document and never edits it.
Public Function WU_DetectStructure(ByVal document As Document) As Variant
    Dim paragraphs As Paragraphs, count As Long, result() As Variant
    Dim i As Long, paragraph As Paragraph, scope As Range, text As String, rawText As String
    Dim marker As String, label As String, style As String, context As String, alternatives As String
    Dim score As Long, confidence As Long, level As Long, outline As Long, evidence As String, semanticRole As String, parentAt(1 To 9) As Long
    Dim hasLists As Boolean, hasTables As Boolean, plausible As Boolean
    Dim headingStyle As Boolean, keepNext As Boolean, upperText As Boolean
    Dim story As Range, pieces As Variant, position As Long, startPosition As Long, endPosition As Long
    Dim starts() As Long, ends() As Long, texts() As String
    Dim centered() As Boolean, frontEmphasis() As Boolean, frontRoles() As String
    Set story = document.StoryRanges(wdMainTextStory): Set paragraphs = story.Paragraphs
    count = paragraphs.count
    If count > 20000 Then Err.Raise 5, "WU_DetectStructure", "paragraph limit exceeds 20000"
    If count = 0 Then WU_DetectStructure = Array(): Exit Function
    ReDim result(0 To count - 1, 0 To WU_COLUMNS - 1)
    ReDim centered(0 To count - 1): ReDim frontEmphasis(0 To count - 1): ReDim frontRoles(0 To count - 1)
    hasLists = (document.Lists.count > 0): hasTables = (document.Tables.count > 0)
    If Not hasTables Then
        pieces = Split(story.text, vbCr): position = story.Start
        ReDim starts(0 To count - 1): ReDim ends(0 To count - 1)
        ReDim texts(0 To count - 1)
        For i = 0 To count - 1
            rawText = CStr(pieces(i)): starts(i) = position
            ends(i) = position + Len(rawText) + 1: position = ends(i)
            texts(i) = WU_CleanText(rawText)
        Next i
    End If
    i = 0
    For Each paragraph In paragraphs
        i = i + 1
        Set scope = Nothing
        If hasTables Then
            Set scope = paragraph.Range: rawText = scope.text
            startPosition = scope.Start: endPosition = scope.End
            text = WU_CleanText(rawText): style = CStr(paragraph.Style)
        Else
            text = texts(i - 1): style = CStr(paragraph.Style)
            startPosition = starts(i - 1): endPosition = ends(i - 1)
        End If
        context = "body": If hasTables Then If scope.Information(wdWithInTable) Then context = "table"
        marker = WU_ParseMarker(text): alternatives = WU_MarkerAlternatives(marker): label = vbNullString
        If i <= 64 And context = "body" And Len(Trim$(text)) > 0 Then
            If scope Is Nothing Then Set scope = paragraph.Range
            centered(i - 1) = (paragraph.Alignment = wdAlignParagraphCenter)
            frontEmphasis(i - 1) = (scope.Bold = True Or scope.Font.SmallCaps = True Or scope.Font.AllCaps = True)
        End If
        If hasLists Then
            If scope Is Nothing Then Set scope = paragraph.Range
            If scope.ListFormat.ListType <> wdListNoNumbering Then label = scope.ListFormat.ListString
        End If
        score = 0: level = 0: evidence = vbNullString
        If context = "body" And Len(Trim$(text)) > 0 Then
            headingStyle = WU_HeadingStyle(style)
            upperText = (Len(text) <= 160 And WU_UpperShare(text) >= 0.8)
            plausible = (Len(marker) > 0 Or headingStyle Or upperText Or (Len(text) <= 160 And Not WU_SentenceEnding(text)))
            keepNext = False
            If plausible Then
                outline = paragraph.OutlineLevel
                If outline >= 1 And outline <= 9 Then
                    level = outline: score = score + 55: WU_AddEvidence evidence, "native-outline"
                End If
                If scope Is Nothing Then Set scope = paragraph.Range
                keepNext = (paragraph.KeepWithNext <> 0)
                If keepNext Then score = score + 8: WU_AddEvidence evidence, "keep-next"
                If upperText Then score = score + 8: WU_AddEvidence evidence, "uppercase-text"
                ' wdUndefined is mixed formatting, not affirmative evidence.
                If scope.Bold = True Then score = score + 8: WU_AddEvidence evidence, "aggregate-bold"
                If scope.Font.SmallCaps = True Or scope.Font.AllCaps = True Then score = score + 8: WU_AddEvidence evidence, "aggregate-caps"
            End If
            If headingStyle Then score = score + 35: WU_AddEvidence evidence, "heading-style"
            If Len(marker) > 0 Then score = score + 30: WU_AddEvidence evidence, "marker:" & marker
            If Len(text) <= 160 Then score = score + 8 Else score = score - 25
            ' A numbered list is not a heading merely because its text looks
            ' like a marker. Preserve bold/keep-with-next heading evidence,
            ' but demote ordinary numbered list paragraphs before resolution.
            If Len(label) > 0 And level = 0 And Not headingStyle And Not keepNext Then
                If scope.Bold <> True And scope.Font.SmallCaps <> True And scope.Font.AllCaps <> True Then score = score - 30: WU_AddEvidence evidence, "ordinary-list-risk"
            End If
            If WU_SentenceEnding(text) And level = 0 And Len(marker) = 0 Then score = score - 18: WU_AddEvidence evidence, "sentence-ending"
            WU_CustomizeCandidate paragraph, score, level, evidence
        End If
        result(i - 1, WU_STYLE) = style: result(i - 1, WU_TEXT) = text
        result(i - 1, WU_START) = startPosition: result(i - 1, WU_END) = endPosition
        result(i - 1, WU_MARKER) = marker: result(i - 1, WU_LIST_LABEL) = label
        result(i - 1, WU_CONTEXT) = context: result(i - 1, WU_EVIDENCE) = evidence
        result(i - 1, WU_CANDIDATE_SCORE) = score: result(i - 1, WU_CONFIDENCE) = score: result(i - 1, WU_LEVEL) = level
        result(i - 1, WU_PARENT) = 0: result(i - 1, WU_AMBIGUOUS) = False
        result(i - 1, WU_SOURCE_ID) = "main:" & CStr(startPosition) & ":" & CStr(endPosition)
        result(i - 1, WU_PARENT_SOURCE_ID) = vbNullString
        result(i - 1, WU_PROVENANCE) = "Word object model"
        result(i - 1, WU_CONTRADICTION) = vbNullString
        result(i - 1, WU_ALTERNATIVES) = alternatives
        result(i - 1, WU_AMBIGUOUS) = (InStr(1, alternatives, "|", vbBinaryCompare) > 0)
        If result(i - 1, WU_AMBIGUOUS) Then WU_AddEvidence evidence, "ambiguous-marker": result(i - 1, WU_EVIDENCE) = evidence
    Next paragraph
    WU_ClassifyFrontMatter result, count, centered, frontEmphasis, frontRoles
    WU_ResolveStyleFamilies result, count
    WU_ResolveMarkerLadder result, count
    For i = 0 To count - 1
        score = CLng(result(i, WU_CONFIDENCE)): confidence = score: level = CLng(result(i, WU_LEVEL)): semanticRole = vbNullString
        If result(i, WU_CONTEXT) <> "body" Then
            result(i, WU_ROLE) = result(i, WU_CONTEXT)
            score = 0: confidence = 100
        ElseIf Len(Trim$(CStr(result(i, WU_TEXT)))) = 0 Then
            result(i, WU_ROLE) = "blank"
            score = 0: confidence = 100
        Else
            semanticRole = WU_SemanticRole(CStr(result(i, WU_TEXT)), CStr(result(i, WU_STYLE)), CStr(result(i, WU_CONTEXT)), level)
            If Len(semanticRole) > 0 Then
                result(i, WU_ROLE) = semanticRole: result(i, WU_LEVEL) = 0
                score = 100: confidence = 100: result(i, WU_AMBIGUOUS) = False
                result(i, WU_EVIDENCE) = "semantic-style-or-context"
                result(i, WU_CONTRADICTION) = vbNullString: result(i, WU_ALTERNATIVES) = vbNullString
                If semanticRole = "quotation" Or semanticRole = "abstract" Then result(i, WU_PARENT) = WU_FindDeepestParent(parentAt)
            ElseIf Len(frontRoles(i)) > 0 Then
                result(i, WU_ROLE) = frontRoles(i): result(i, WU_LEVEL) = 0: result(i, WU_PARENT) = 0
                score = 85: confidence = 85: result(i, WU_AMBIGUOUS) = False
                result(i, WU_EVIDENCE) = "front-matter-layout"
                result(i, WU_CONTRADICTION) = vbNullString: result(i, WU_ALTERNATIVES) = vbNullString
            ElseIf level > 0 And score >= 35 Then
                result(i, WU_ROLE) = "heading": result(i, WU_PARENT) = WU_FindParent(parentAt, level)
                parentAt(level) = i + 1: WU_ClearDeeper parentAt, level
            ElseIf score >= 25 Then
                result(i, WU_ROLE) = "candidate": result(i, WU_AMBIGUOUS) = True
            Else
                result(i, WU_ROLE) = "body"
            End If
        End If
        If result(i, WU_ROLE) <> "heading" And result(i, WU_ROLE) <> "blank" And Len(semanticRole) = 0 And Len(frontRoles(i)) = 0 And CLng(result(i, WU_PARENT)) = 0 Then result(i, WU_PARENT) = WU_FindDeepestParent(parentAt)
        If CLng(result(i, WU_PARENT)) > 0 Then result(i, WU_PARENT_SOURCE_ID) = result(CLng(result(i, WU_PARENT)) - 1, WU_SOURCE_ID)
        result(i, WU_CANDIDATE_SCORE) = score
        If confidence > 100 Then confidence = 100
        If confidence < 0 Then confidence = 0
        result(i, WU_CONFIDENCE) = confidence
    Next i
    WU_DetectStructure = result
End Function

Private Sub WU_ClassifyFrontMatter(ByRef result As Variant, ByVal count As Long, ByRef centered() As Boolean, ByRef frontEmphasis() As Boolean, ByRef frontRoles() As String)
    Dim i As Long, firstContent As Long, limit As Long
    Dim text As String, style As String, titleSeen As Boolean, authorSeen As Boolean
    Dim isCentered As Boolean, titleStyle As Boolean
    firstContent = -1: limit = count
    For i = 0 To count - 1
        If result(i, WU_CONTEXT) = "body" And Len(Trim$(CStr(result(i, WU_TEXT)))) > 0 Then
            If firstContent < 0 Then firstContent = i
            style = LCase$(CStr(result(i, WU_STYLE)))
            If (CLng(result(i, WU_LEVEL)) > 0 Or WU_HeadingStyle(style)) And InStr(style, "title") = 0 Then limit = i: Exit For
        End If
    Next i
    If limit > 64 Then limit = 64
    If firstContent < 0 Then Exit Sub
    For i = 0 To limit - 1
        If result(i, WU_CONTEXT) <> "body" Then GoTo NextFrontMatter
        text = Trim$(CStr(result(i, WU_TEXT))): If Len(text) = 0 Or Len(text) > 240 Then GoTo NextFrontMatter
        style = LCase$(CStr(result(i, WU_STYLE)))
        isCentered = centered(i) Or InStr(style, "centred") > 0 Or InStr(style, "centered") > 0
        titleStyle = ((WU_StyleToken(style, "document") And WU_StyleToken(style, "title")) Or (WU_StyleToken(style, "heading") And WU_StyleToken(style, "title")) Or Trim$(style) = "title")
        If Not titleSeen And (titleStyle Or (i = firstContent And (isCentered Or frontEmphasis(i)))) Then
            frontRoles(i) = "title": titleSeen = True
        ElseIf titleSeen And Not authorSeen And titleStyle Then
            frontRoles(i) = "title"
        ElseIf titleSeen And Not authorSeen And isCentered And WU_LikelyAuthorLine(text) Then
            frontRoles(i) = "author": authorSeen = True
        ElseIf titleSeen And Not authorSeen And isCentered And frontEmphasis(i) Then
            frontRoles(i) = "title"
        End If
NextFrontMatter:
    Next i
End Sub

Private Function WU_LikelyAuthorLine(ByVal text As String) As Boolean
    Dim lower As String, item As Variant
    lower = LCase$(Trim$(text))
    If Left$(lower, 3) = "by " Then WU_LikelyAuthorLine = True: Exit Function
    For Each item In Array("draft", "revised", "january", "february", "march", "april", "may ", "june", "july", "august", "september", "october", "november", "december")
        If InStr(lower, CStr(item)) > 0 Then Exit Function
    Next item
    WU_LikelyAuthorLine = (InStr(text, ",") > 0 Or (InStr(text, "*") > 0 And InStr(lower, " and ") > 0))
End Function

Private Function WU_SemanticRole(ByVal text As String, ByVal style As String, ByVal context As String, ByVal level As Long) As String
    Dim plain As String
    plain = LCase$(Trim$(text))
    Do While Len(plain) > 0 And Right$(plain, 1) = ":": plain = Trim$(Left$(plain, Len(plain) - 1)): Loop
    If plain = "" Then Exit Function
    If plain = "abstract" And level = 0 Then WU_SemanticRole = "abstract": Exit Function
    If plain = "contents" Or plain = "table of contents" Or context = "contents" Then WU_SemanticRole = "toc": Exit Function
    style = LCase$(style)
	If (InStr(style, "tocheading") > 0 Or (WU_StyleToken(style, "toc") And WU_StyleToken(style, "heading"))) And WU_LikelyTOCText(text, context) Then WU_SemanticRole = "toc": Exit Function
    If WU_StyleToken(style, "quotation") Or WU_StyleToken(style, "quote") Or (WU_StyleToken(style, "block") And WU_StyleToken(style, "text")) Then WU_SemanticRole = "quotation": Exit Function
    If WU_StyleToken(style, "abstract") Then WU_SemanticRole = "abstract": Exit Function
    If WU_StyleToken(style, "author") Or WU_StyleToken(style, "byline") Then WU_SemanticRole = "author": Exit Function
	If Not WU_StyleToken(style, "heading") And ((WU_StyleToken(style, "document") And WU_StyleToken(style, "title")) Or Trim$(style) = "title" Or Trim$(style) = "title normal") Then WU_SemanticRole = "title"
End Function

' A built-in TOC style may be repurposed for body/front-matter text. Require
' short TOC-label or entry evidence before allowing the style name to win.
Private Function WU_LikelyTOCText(ByVal text As String, ByVal context As String) As Boolean
    Dim plain As String
    If context = "contents" Then WU_LikelyTOCText = True: Exit Function
    plain = LCase$(Trim$(text))
    If plain = "contents" Or plain = "table of contents" Or plain = "table of contents:" Then WU_LikelyTOCText = True: Exit Function
    If InStr(text, vbTab) > 0 Then WU_LikelyTOCText = True: Exit Function
    If Len(plain) <= 96 And WU_TrailingPageNumber(plain) Then WU_LikelyTOCText = True
End Function

Private Function WU_TrailingPageNumber(ByVal text As String) As Boolean
    Dim i As Long, c As String
    text = Trim$(text)
    If Len(text) = 0 Then Exit Function
    For i = Len(text) To 1 Step -1
        c = Mid$(text, i, 1)
        If c < "0" Or c > "9" Then Exit For
    Next i
    WU_TrailingPageNumber = (i < Len(text))
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
        If level >= 1 And level <= 9 And result(i, WU_CONTEXT) = "body" And Len(Trim$(CStr(result(i, WU_STYLE)))) > 0 Then
            slot = WU_StyleSlot(keys, capacity, CStr(result(i, WU_STYLE)))
            If votes(slot) > 0 And levels(slot) <> level Then conflict(slot) = True
            levels(slot) = level: votes(slot) = votes(slot) + 1
        End If
    Next i
    For i = 0 To count - 1
        If CLng(result(i, WU_LEVEL)) = 0 And result(i, WU_CONTEXT) = "body" And Len(Trim$(CStr(result(i, WU_STYLE)))) > 0 And CLng(result(i, WU_CONFIDENCE)) >= 25 Then
            slot = WU_StyleSlot(keys, capacity, CStr(result(i, WU_STYLE)))
            If votes(slot) >= 2 And Not conflict(slot) Then
                result(i, WU_LEVEL) = levels(slot): result(i, WU_CONFIDENCE) = CLng(result(i, WU_CONFIDENCE)) + 15
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
    Dim at As Long, dashAt As Long, prefix As String, namedKind As String, rest As String, i As Long, c As String, lowerText As String, candidate As Variant
    text = Trim$(Replace(Replace(text, vbTab, " "), ChrW(160), " "))
    lowerText = LCase$(text)
    If WU_IsNamedHeading(lowerText) Then
        at = InStr(text, ":"): If at = 0 Then at = InStr(text, " - ")
        If at = 0 Then at = InStr(text, ChrW(&H2013))
        If at = 0 Then at = InStr(text, ChrW(&H2014))
        If at > 1 Then
            namedKind = Left$(text, InStr(text, " ") - 1)
            prefix = Trim$(Mid$(text, InStr(text, " ") + 1, at - InStr(text, " ") - 1))
            If IsNumeric(prefix) Or WU_WordNumber(prefix) > 0 Or WU_IsRoman(prefix) Or (Len(prefix) = 1 And LCase$(prefix) >= "a" And LCase$(prefix) <= "z") Then WU_ParseMarker = "named:" & LCase$(namedKind) & ":" & prefix: Exit Function
        End If
    End If
    at = InStr(text, ".")
    If at >= 2 And at <= 8 And Left$(Mid$(text, at + 1), 1) = " " Then
        prefix = Left$(text, at - 1): rest = Mid$(text, at + 1)
    Else
        at = InStr(text, ")")
        If at >= 2 And at <= 8 And Left$(Mid$(text, at + 1), 1) = " " Then
            prefix = Left$(text, at - 1): rest = Mid$(text, at + 1)
        Else
            at = 0
            For Each candidate In Array("-", ChrW(&H2013), ChrW(&H2014))
                dashAt = InStr(text, CStr(candidate))
                If dashAt >= 2 And dashAt <= 8 And Left$(Mid$(text, dashAt + 1), 1) = " " Then
                    at = dashAt: prefix = Trim$(Left$(text, at - 1)): rest = Mid$(text, at + 1): Exit For
                End If
            Next candidate
            If at = 0 Then Exit Function
        End If
    End If
    If Len(Trim$(rest)) = 0 Or InStr(prefix, " ") > 0 Then Exit Function
    If IsNumeric(prefix) Then WU_ParseMarker = prefix: Exit Function
    If Len(prefix) = 1 And LCase$(prefix) >= "a" And LCase$(prefix) <= "z" Then WU_ParseMarker = prefix: Exit Function
    For i = 1 To Len(prefix): c = UCase$(Mid$(prefix, i, 1)): If InStr(1, "IVXLCDM", c, vbBinaryCompare) = 0 Then Exit Function
    Next i
    WU_ParseMarker = prefix
End Function

' Keep the primary marker for the fast ladder, but expose competing readings
' for one-letter Roman/alpha prefixes instead of silently choosing one.
Private Function WU_MarkerAlternatives(ByVal marker As String) As String
    Dim value As Long, alphaValue As Long, c As String
    If Len(marker) = 0 Then Exit Function
    If Left$(marker, 6) = "named:" Then
        WU_MarkerAlternatives = marker
        Exit Function
    End If
    If Len(marker) = 1 Then
        c = UCase$(marker)
        If InStr(1, "IVXLCDM", c, vbBinaryCompare) > 0 Then
            value = WU_MarkerValue(marker, "roman")
            alphaValue = AscW(c) - 64
            If marker = LCase$(marker) Then
                WU_MarkerAlternatives = marker & "|roman:" & CStr(value) & "|lower_alpha:" & CStr(alphaValue)
            Else
                WU_MarkerAlternatives = marker & "|roman:" & CStr(value) & "|upper_alpha:" & CStr(alphaValue)
            End If
            Exit Function
        End If
        If c >= "A" And c <= "Z" Then
            WU_MarkerAlternatives = marker
            Exit Function
        End If
    End If
    WU_MarkerAlternatives = marker
End Function

Private Function WU_MarkerFamily(ByVal marker As String) As String
    Dim lowerMarker As String, namedKind As String, at As Long
    lowerMarker = LCase$(marker)
    If Left$(lowerMarker, 5) = "part:" Then
        WU_MarkerFamily = "named_section": Exit Function
    End If
    If Left$(lowerMarker, 6) = "named:" Then
        at = InStr(7, lowerMarker, ":")
        If at > 0 Then namedKind = Mid$(lowerMarker, 7, at - 7)
        If namedKind = "part" Or namedKind = "chapter" Or Len(namedKind) = 0 Then
            WU_MarkerFamily = "named_section"
        Else
            WU_MarkerFamily = "named_" & namedKind & "_section"
        End If
        Exit Function
    ElseIf IsNumeric(marker) Then
        WU_MarkerFamily = "decimal"
    ElseIf Len(marker) = 1 And InStr(1, "IVXLCDM", UCase$(marker), vbBinaryCompare) = 0 Then
        If marker = LCase$(marker) Then WU_MarkerFamily = "lower_alpha" Else WU_MarkerFamily = "upper_alpha"
    Else
        WU_MarkerFamily = "roman"
    End If
End Function

Private Function WU_MarkerValue(ByVal marker As String, ByVal family As String) As Long
    If family = "named" Or Left$(family, 6) = "named_" Then
        marker = Mid$(marker, InStrRev(marker, ":", -1, vbBinaryCompare) + 1)
        If WU_WordNumber(marker) > 0 Then WU_MarkerValue = WU_WordNumber(marker): Exit Function
        If IsNumeric(marker) Then WU_MarkerValue = CLng(marker): Exit Function
        If Len(marker) = 1 And LCase$(marker) >= "a" And LCase$(marker) <= "z" Then WU_MarkerValue = AscW(UCase$(marker)) - 64: Exit Function
        marker = UCase$(marker): family = "roman"
    End If
    If family = "decimal" Then WU_MarkerValue = CLng(marker): Exit Function
    If family = "lower_alpha" Or family = "upper_alpha" Then WU_MarkerValue = AscW(UCase$(marker)) - 64: Exit Function
    marker = UCase$(marker)
    Dim i As Long, n As Long, last As Long, current As Long
    For i = Len(marker) To 1 Step -1
        current = InStr(1, "IVXLCDM", Mid$(marker, i, 1), vbBinaryCompare)
        current = Choose(current, 1, 5, 10, 50, 100, 500, 1000)
        If current < last Then n = n - current Else n = n + current: last = current
    Next i
    WU_MarkerValue = n
End Function

Private Function WU_IsNamedHeading(ByVal lowerText As String) As Boolean
    Dim prefix As Variant
    For Each prefix In Array("part ", "chapter ", "theme ", "section ", "article ", "appendix ", "schedule ", "division ", "book ", "title ")
        If Left$(lowerText, Len(CStr(prefix))) = CStr(prefix) Then WU_IsNamedHeading = True: Exit Function
    Next prefix
End Function

Private Function WU_WordNumber(ByVal value As String) As Long
    Select Case LCase$(Trim$(value))
        Case "one": WU_WordNumber = 1
        Case "two": WU_WordNumber = 2
        Case "three": WU_WordNumber = 3
        Case "four": WU_WordNumber = 4
        Case "five": WU_WordNumber = 5
        Case "six": WU_WordNumber = 6
        Case "seven": WU_WordNumber = 7
        Case "eight": WU_WordNumber = 8
        Case "nine": WU_WordNumber = 9
        Case "ten": WU_WordNumber = 10
    End Select
End Function

Private Function WU_IsRoman(ByVal value As String) As Boolean
    Dim i As Long
    value = UCase$(Trim$(value)): If Len(value) = 0 Then Exit Function
    For i = 1 To Len(value): If InStr(1, "IVXLCDM", Mid$(value, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    WU_IsRoman = True
End Function

Private Function WU_HeadingStyle(ByVal name As String) As Boolean
    WU_HeadingStyle = WU_StyleToken(name, "heading") Or WU_StyleToken(name, "title") Or WU_StyleToken(name, "head")
End Function

' Match complete style-name tokens, including common camel-case and digit
' boundaries. This keeps Header/AheadBody/Headnote from becoming headings.
Private Function WU_StyleToken(ByVal name As String, ByVal wanted As String) As Boolean
    Dim at As Long, before As String, after As String, original As String, current As String
    original = name: name = LCase$(name): wanted = LCase$(wanted)
    at = InStr(1, name, wanted, vbBinaryCompare)
    Do While at > 0
        before = vbNullString: after = vbNullString: current = vbNullString
        If at > 1 Then before = Mid$(name, at - 1, 1)
        If at + Len(wanted) <= Len(name) Then after = Mid$(name, at + Len(wanted), 1)
        current = Mid$(original, at, 1)
        If (Not WU_StyleIdentifier(before) Or (before >= "a" And before <= "z" And current >= "A" And current <= "Z") Or after >= "0" And after <= "9") And _
           (Not WU_StyleIdentifier(after) Or after >= "0" And after <= "9" Or (current >= "a" And current <= "z" And after >= "A" And after <= "Z")) Then
            WU_StyleToken = True: Exit Function
        End If
        at = InStr(at + 1, name, wanted, vbBinaryCompare)
    Loop
End Function

Private Function WU_StyleIdentifier(ByVal value As String) As Boolean
    If Len(value) = 0 Then Exit Function
    WU_StyleIdentifier = value Like "[A-Za-z0-9_]"
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
Private Function WU_FindDeepestParent(ByRef parents() As Long) As Long
    Dim i As Long: For i = 9 To 1 Step -1: If parents(i) > 0 Then WU_FindDeepestParent = parents(i): Exit Function
    Next i
End Function
Private Sub WU_ClearDeeper(ByRef parents() As Long, ByVal level As Long)
    Dim i As Long: For i = level + 1 To 9: parents(i) = 0: Next i
End Sub
