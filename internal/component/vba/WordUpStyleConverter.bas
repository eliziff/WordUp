Attribute VB_Name = "WordUpStyleConverter"
Option Explicit
' Style conversion and application over explicitly selected Word stories or
' an exact Range. Requires WordUpStories and WordUpSafeEdit.
'
'   WU_ConvertParagraphStyle / WU_ConvertCharacterStyle   restyle every use
'   WU_ConvertStyleBatch / WU_ConvertStyleBatchInRange    ordered from/to rows
'   WU_ConvertStyleInRange                                one mapping, exact Range
'   WU_ApplyParagraphStyleInRange / WU_ApplyCharacterStyleInRange
'   WU_ApplyParagraphStyleRuns                            [start, end, style] rows
'
' Conversions use Word's own formatted Find/Replace in one native pass per
' story, inside one undo record, restoring Application.ScreenUpdating. Word
' resets direct character formatting inside a run when a character style is
' applied, exactly as it does in its user interface; text and inline structure
' are never rewritten. Missing styles and mismatched style kinds are rejected
' before any state changes. A same-style mapping is a no-op.

Public Function WU_ConvertParagraphStyle(ByVal document As Document, ByVal fromStyle As String, ByVal toStyle As String, Optional ByVal scope As String = "all") As Boolean
    WU_ConvertParagraphStyle = WU_ConvertOne(document, fromStyle, toStyle, scope, wdStyleTypeParagraph, "WU_ConvertParagraphStyle")
End Function

Public Function WU_ConvertCharacterStyle(ByVal document As Document, ByVal fromStyle As String, ByVal toStyle As String, Optional ByVal scope As String = "all") As Boolean
    WU_ConvertCharacterStyle = WU_ConvertOne(document, fromStyle, toStyle, scope, wdStyleTypeCharacter, "WU_ConvertCharacterStyle")
End Function

' mappings is a two-dimensional array whose rows hold a source style name and
' a target style name. The kind (paragraph or character) follows each source
' style. Rows apply in order inside one undo record; the result is the number
' of rows that changed something.
Public Function WU_ConvertStyleBatch(ByVal document As Document, ByVal mappings As Variant, Optional ByVal scope As String = "all") As Long
    Dim stories As Collection, story As Range, row As Long, firstRow As Long, lastRow As Long, firstColumn As Long, s As Long
    Dim sources() As Style, targets() As Style, kinds() As Long, matched As Boolean
    Dim storyStarts() As Long, storyEnds() As Long
    Dim updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ConvertStyleBatch", "document is required"
    WU_ValidateMappings document, mappings, "WU_ConvertStyleBatch", sources, targets, kinds, firstRow, lastRow, firstColumn
    Set stories = WU_Stories(document, WU_NormalizeScope(scope, "WU_ConvertStyleBatch"))
    If stories.Count > 0 Then
        ' Word shrinks every Range spanning a replaced paragraph mark, so each
        ' row re-derives the full story bounds captured before any change.
        ReDim storyStarts(1 To stories.Count)
        ReDim storyEnds(1 To stories.Count)
        For s = 1 To stories.Count
            storyStarts(s) = stories(s).Start
            storyEnds(s) = stories(s).End
        Next s
    End If
    For row = firstRow To lastRow
        matched = False
        If StrComp(sources(row).NameLocal, targets(row).NameLocal, vbTextCompare) <> 0 Then
            For s = 1 To stories.Count
                Set story = WU_Rebound(stories(s), storyStarts(s), storyEnds(s))
                If WU_StyleExists(story, sources(row), kinds(row)) Then
                    If Not opened Then WU_BeginSafeEdit updating, opened, captured, "Convert styles"
                    If WU_ConvertIn(story, sources(row), targets(row), kinds(row)) Then matched = True
                End If
            Next s
        End If
        If matched Then WU_ConvertStyleBatch = WU_ConvertStyleBatch + 1
    Next row
CleanUp:
    On Error Resume Next
    WU_EndSafeEdit updating, opened, captured
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

Public Function WU_ConvertStyleBatchInRange(ByVal target As Range, ByVal mappings As Variant) As Long
    Dim row As Long, firstRow As Long, lastRow As Long, firstColumn As Long, targetStart As Long, targetEnd As Long, scope As Range
    Dim sources() As Style, targets() As Style, kinds() As Long
    Dim updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ConvertStyleBatchInRange", "target range is required"
    WU_ValidateMappings target.Document, mappings, "WU_ConvertStyleBatchInRange", sources, targets, kinds, firstRow, lastRow, firstColumn
    targetStart = target.Start: targetEnd = target.End
    If targetEnd <= targetStart Then Exit Function
    For row = firstRow To lastRow
        If StrComp(sources(row).NameLocal, targets(row).NameLocal, vbTextCompare) <> 0 Then
            ' Re-derive the caller's bounds for every row: a replaced paragraph
            ' mark shrinks every Range that spanned it, including target.
            Set scope = WU_Rebound(target, targetStart, targetEnd)
            If WU_StyleExists(scope, sources(row), kinds(row)) Then
                If Not opened Then WU_BeginSafeEdit updating, opened, captured, "Convert styles"
                If WU_ConvertIn(scope, sources(row), targets(row), kinds(row)) Then WU_ConvertStyleBatchInRange = WU_ConvertStyleBatchInRange + 1
            End If
        End If
    Next row
CleanUp:
    On Error Resume Next
    WU_EndSafeEdit updating, opened, captured
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

Public Function WU_ConvertStyleInRange(ByVal target As Range, ByVal fromStyle As String, ByVal toStyle As String) As Boolean
    Dim mappings(0 To 0, 0 To 1) As Variant
    mappings(0, 0) = fromStyle: mappings(0, 1) = toStyle
    WU_ConvertStyleInRange = (WU_ConvertStyleBatchInRange(target, mappings) > 0)
End Function

' Apply a paragraph style to every paragraph the exact Range touches. The
' range is never widened.
Public Function WU_ApplyParagraphStyleInRange(ByVal target As Range, ByVal styleName As String) As Boolean
    Dim st As Style, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ApplyParagraphStyleInRange", "target range is required"
    Set st = WU_ResolveStyle(target.Document, styleName, wdStyleTypeParagraph, "WU_ApplyParagraphStyleInRange")
    If target.End <= target.Start Then Exit Function
    If WU_RangeHasStyle(target, st) Then Exit Function
    WU_BeginSafeEdit updating, opened, captured, "Apply paragraph style"
    target.Duplicate.Style = st
    WU_ApplyParagraphStyleInRange = True
CleanUp:
    On Error Resume Next
    WU_EndSafeEdit updating, opened, captured
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

' Apply a character style to the exact Range. Word resets direct character
' formatting inside the range, as it does in its user interface.
Public Function WU_ApplyCharacterStyleInRange(ByVal target As Range, ByVal styleName As String) As Boolean
    Dim st As Style, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ApplyCharacterStyleInRange", "target range is required"
    Set st = WU_ResolveStyle(target.Document, styleName, wdStyleTypeCharacter, "WU_ApplyCharacterStyleInRange")
    If target.End <= target.Start Then Exit Function
    If WU_RangeHasStyle(target, st) Then Exit Function
    WU_BeginSafeEdit updating, opened, captured, "Apply character style"
    target.Duplicate.Style = st
    WU_ApplyCharacterStyleInRange = True
CleanUp:
    On Error Resume Next
    WU_EndSafeEdit updating, opened, captured
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

' runs is a two-dimensional array whose rows hold an absolute start position,
' an absolute end position and a paragraph style name, in ascending order
' inside target. Every row is validated before any style changes; the result
' is the number of runs whose style changed. Direct character formatting is
' left in place, as Word does when styling a range.
Public Function WU_ApplyParagraphStyleRuns(ByVal target As Range, ByVal runs As Variant) As Long
    Dim scope As Range, row As Long, firstRow As Long, lastRow As Long, firstColumn As Long
    Dim styleCache() As Style, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ApplyParagraphStyleRuns", "target range is required"
    If Not WU_ValidateRuns(target, runs, styleCache, firstRow, lastRow, firstColumn) Then Exit Function
    WU_BeginSafeEdit updating, opened, captured, "Apply paragraph style runs"
    Set scope = target.Duplicate
    For row = firstRow To lastRow
        ' Extend first, then move the start; assigning a later start to a
        ' reused Range before its end can make Word reject the transient span.
        scope.End = CLng(runs(row, firstColumn + 1))
        scope.Start = CLng(runs(row, firstColumn))
        If Not WU_RangeHasStyle(scope, styleCache(row)) Then
            scope.Style = styleCache(row)
            WU_ApplyParagraphStyleRuns = WU_ApplyParagraphStyleRuns + 1
        End If
    Next row
CleanUp:
    On Error Resume Next
    WU_EndSafeEdit updating, opened, captured
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

' ---- internals -------------------------------------------------------------------

' A fresh Range over the given bounds of scope, staying in scope's story.
' Extend first, then move the start, so a transient inverted span is never
' handed to Word.
Private Function WU_Rebound(ByVal scope As Range, ByVal startPosition As Long, ByVal endPosition As Long) As Range
    Set WU_Rebound = scope.Duplicate
    WU_Rebound.End = endPosition
    WU_Rebound.Start = startPosition
End Function

Private Function WU_ConvertOne(ByVal document As Document, ByVal fromStyle As String, ByVal toStyle As String, ByVal scope As String, ByVal kind As Long, ByVal sourceName As String) As Boolean
    Dim stories As Collection, story As Range, source As Style, target As Style
    Dim updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, sourceName, "document is required"
    Set source = WU_ResolveStyle(document, fromStyle, kind, sourceName)
    Set target = WU_ResolveStyle(document, toStyle, kind, sourceName)
    Set stories = WU_Stories(document, WU_NormalizeScope(scope, sourceName))
    If StrComp(source.NameLocal, target.NameLocal, vbTextCompare) = 0 Then Exit Function
    For Each story In stories
        If WU_StyleExists(story, source, kind) Then
            If Not opened Then WU_BeginSafeEdit updating, opened, captured, "Convert style"
            If WU_ConvertIn(story, source, target, kind) Then WU_ConvertOne = True
        End If
    Next story
CleanUp:
    On Error Resume Next
    WU_EndSafeEdit updating, opened, captured
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

Private Sub WU_ValidateMappings(ByVal document As Document, ByVal mappings As Variant, ByVal sourceName As String, ByRef sources() As Style, ByRef targets() As Style, ByRef kinds() As Long, ByRef firstRow As Long, ByRef lastRow As Long, ByRef firstColumn As Long)
    Dim row As Long, kind As Long
    If Not IsArray(mappings) Then Err.Raise 5, sourceName, "mappings must be a two-dimensional array of source and target style names"
    On Error Resume Next
    firstRow = LBound(mappings, 1): lastRow = UBound(mappings, 1): firstColumn = LBound(mappings, 2)
    If Err.Number <> 0 Or UBound(mappings, 2) - firstColumn < 1 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 5, sourceName, "mappings must be a two-dimensional array with source and target columns"
    End If
    On Error GoTo 0
    If lastRow < firstRow Then Exit Sub
    ReDim sources(firstRow To lastRow)
    ReDim targets(firstRow To lastRow)
    ReDim kinds(firstRow To lastRow)
    For row = firstRow To lastRow
        Set sources(row) = WU_ResolveStyle(document, CStr(mappings(row, firstColumn)), 0, sourceName)
        kind = WU_StyleKind(sources(row))
        Set targets(row) = WU_ResolveStyle(document, CStr(mappings(row, firstColumn + 1)), kind, sourceName)
        kinds(row) = kind
    Next row
End Sub

' Resolve a style by name; kind 0 accepts paragraph or character styles,
' otherwise the style must be usable as that kind (linked styles are both).
Private Function WU_ResolveStyle(ByVal document As Document, ByVal styleName As String, ByVal kind As Long, ByVal sourceName As String) As Style
    Dim st As Style, styleError As Long
    If document Is Nothing Then Err.Raise 91, sourceName, "document is required"
    If Len(Trim$(styleName)) = 0 Then Err.Raise 5, sourceName, "style name is required"
    On Error Resume Next
    Set st = document.Styles(styleName)
    styleError = Err.Number
    Err.Clear
    On Error GoTo 0
    If styleError <> 0 Or st Is Nothing Then Err.Raise 5, sourceName, "style " & styleName & " was not found"
    Select Case kind
        Case wdStyleTypeParagraph
            If st.Type <> wdStyleTypeParagraph And st.Type <> wdStyleTypeParagraphOnly And st.Type <> wdStyleTypeLinked Then Err.Raise 5, sourceName, "style " & styleName & " is not a paragraph style"
        Case wdStyleTypeCharacter
            If st.Type <> wdStyleTypeCharacter And st.Type <> wdStyleTypeLinked Then Err.Raise 5, sourceName, "style " & styleName & " is not a character style"
        Case Else
            If WU_StyleKind(st) = 0 Then Err.Raise 5, sourceName, "style " & styleName & " is not a paragraph or character style"
    End Select
    Set WU_ResolveStyle = st
End Function

Private Function WU_StyleKind(ByVal st As Style) As Long
    Select Case st.Type
        Case wdStyleTypeCharacter
            WU_StyleKind = wdStyleTypeCharacter
        Case wdStyleTypeParagraph, wdStyleTypeParagraphOnly, wdStyleTypeLinked
            WU_StyleKind = wdStyleTypeParagraph
    End Select
End Function

Private Function WU_RangeHasStyle(ByVal scope As Range, ByVal st As Style) As Boolean
    Dim current As String
    On Error Resume Next
    current = scope.Style.NameLocal
    Err.Clear
    On Error GoTo 0
    WU_RangeHasStyle = (StrComp(current, st.NameLocal, vbTextCompare) = 0)
End Function

Private Sub WU_PrepareStyleFind(ByVal criteria As Find, ByVal source As Style, ByVal target As Style, ByVal kind As Long)
    With criteria
        .ClearFormatting
        .Replacement.ClearFormatting
        If kind = wdStyleTypeParagraph Then
            .Text = "^p"
            .Replacement.Text = "^&"
        Else
            .Text = vbNullString
            .Replacement.Text = vbNullString
        End If
        .Style = source
        If Not target Is Nothing Then .Replacement.Style = target
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    WU_PinFindOptions criteria
End Sub

Private Function WU_StyleExists(ByVal scope As Range, ByVal source As Style, ByVal kind As Long) As Boolean
    Dim probe As Range
    Set probe = scope.Duplicate
    WU_PrepareStyleFind probe.Find, source, Nothing, kind
    WU_StyleExists = probe.Find.Execute
End Function

Private Function WU_ConvertIn(ByVal scope As Range, ByVal source As Style, ByVal target As Style, ByVal kind As Long) As Boolean
    Dim probe As Range
    Set probe = scope.Duplicate
    WU_PrepareStyleFind probe.Find, source, target, kind
    WU_ConvertIn = probe.Find.Execute(Replace:=wdReplaceAll)
End Function

Private Function WU_ValidateRuns(ByVal target As Range, ByVal runs As Variant, ByRef styleCache() As Style, ByRef firstRow As Long, ByRef lastRow As Long, ByRef firstColumn As Long) As Boolean
    Dim row As Long, startPosition As Long, endPosition As Long, previousEnd As Long
    If Not IsArray(runs) Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "runs must be a two-dimensional array of start, end and style name"
    On Error Resume Next
    firstRow = LBound(runs, 1): lastRow = UBound(runs, 1): firstColumn = LBound(runs, 2)
    If Err.Number <> 0 Or UBound(runs, 2) - firstColumn < 2 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 5, "WU_ApplyParagraphStyleRuns", "runs must be a two-dimensional array with start, end and style columns"
    End If
    On Error GoTo 0
    If lastRow < firstRow Then Exit Function
    ReDim styleCache(firstRow To lastRow)
    previousEnd = target.Start
    For row = firstRow To lastRow
        If Not IsNumeric(runs(row, firstColumn)) Or Not IsNumeric(runs(row, firstColumn + 1)) Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "run " & CStr(row) & " positions must be numeric"
        startPosition = CLng(runs(row, firstColumn)): endPosition = CLng(runs(row, firstColumn + 1))
        If startPosition < previousEnd Or endPosition <= startPosition Or endPosition > target.End Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "run " & CStr(row) & " must be an ascending, non-overlapping span inside the target range"
        Set styleCache(row) = WU_ResolveStyle(target.Document, CStr(runs(row, firstColumn + 2)), wdStyleTypeParagraph, "WU_ApplyParagraphStyleRuns")
        previousEnd = endPosition
    Next row
    WU_ValidateRuns = True
End Function

' Word keeps several Find flags sticky across interactive searches; pin them so
' results are deterministic. An unpinned formatted Find with empty text can
' also loop on some hosts. Options missing on older hosts are ignored.
Private Sub WU_PinFindOptions(ByVal criteria As Find)
    On Error Resume Next
    criteria.MatchFuzzy = False
    criteria.MatchPhrase = False
    criteria.MatchByte = False
    criteria.MatchKashida = False
    criteria.MatchDiacritics = False
    criteria.MatchAlefHamza = False
    criteria.MatchControl = False
    criteria.MatchPrefix = False
    criteria.MatchSuffix = False
    Err.Clear
    On Error GoTo 0
End Sub
