Attribute VB_Name = "WordUpTextOperations"
Option Explicit
' Bounded text operations over explicitly selected Word stories or an exact
' Range, never through Selection. Requires WordUpStories and WordUpSafeEdit.
'
'   WU_ReplaceText / WU_ReplaceTextInRange     replace every occurrence
'   WU_CountText / WU_CountTextInRange         read-only occurrence count
'   WU_ReplaceBatch / WU_ReplaceBatchInRange   ordered find/replace rules
'   WU_StyleMatches / WU_StyleMatchesInRange   apply a character style to matches
'   WU_ApplyCharacterStyleRuns                 character styles on exact
'                                              [start, end, style] offset runs
'
' findText is literal unless wildcards is True, in which case it is Word's own
' wildcard grammar and replacement text may use \1 style backreferences.
' Literal text is escaped so carets stay literal. Word's 255-character Find
' limit is enforced before any state changes. Mutating operations run inside
' one undo record, restore Application.ScreenUpdating, and return without
' touching Word state when nothing matches.
'
' Exact ranges are honoured to the character. Word's own Find cannot be
' trusted at a range edge: a whole word is missed when the range equals the
' word, ReplaceAll runs past the range end, and ReplaceOne on a range equal to
' its match replaces a later occurrence instead. Every search therefore runs
' on a probe one character wider than the bounds and filters matches to the
' bounds; exact-range replacement replaces each located match on its own.

' ---- replace ---------------------------------------------------------------

Public Function WU_ReplaceText(ByVal document As Document, ByVal findText As String, ByVal replaceText As String, Optional ByVal scope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False, Optional ByVal wildcards As Boolean = False) As Boolean
    Dim stories As Collection, story As Range, storyEnd As Long, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ReplaceText", "document is required"
    WU_ValidateFind findText, replaceText, wildcards, "WU_ReplaceText"
    If matchCase And Not wildcards And findText = replaceText Then Exit Function
    Set stories = WU_Stories(document, WU_NormalizeScope(scope, "WU_ReplaceText"))
    For Each story In stories
        If WU_Exists(story, findText, matchCase, wholeWord, wildcards) Then
            If Not opened Then WU_BeginSafeEdit updating, opened, captured, "Replace text"
            storyEnd = story.End
            If WU_ReplaceIn(story, findText, replaceText, matchCase, wholeWord, wildcards, storyEnd) Then WU_ReplaceText = True
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

Public Function WU_ReplaceTextInRange(ByVal target As Range, ByVal findText As String, ByVal replaceText As String, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False, Optional ByVal wildcards As Boolean = False) As Boolean
    Dim targetEnd As Long, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ReplaceTextInRange", "target range is required"
    WU_ValidateFind findText, replaceText, wildcards, "WU_ReplaceTextInRange"
    If target.End <= target.Start Then Exit Function
    If matchCase And Not wildcards And findText = replaceText Then Exit Function
    If Not WU_Exists(target, findText, matchCase, wholeWord, wildcards) Then Exit Function
    WU_BeginSafeEdit updating, opened, captured, "Replace text"
    targetEnd = target.End
    WU_ReplaceTextInRange = WU_ReplaceIn(target, findText, replaceText, matchCase, wholeWord, wildcards, targetEnd)
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

' ---- count -----------------------------------------------------------------

Public Function WU_CountText(ByVal document As Document, ByVal findText As String, Optional ByVal scope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False, Optional ByVal wildcards As Boolean = False) As Long
    Dim story As Range
    If document Is Nothing Then Err.Raise 91, "WU_CountText", "document is required"
    WU_ValidateFind findText, "", wildcards, "WU_CountText"
    For Each story In WU_Stories(document, WU_NormalizeScope(scope, "WU_CountText"))
        WU_CountText = WU_CountText + WU_CountIn(story, findText, matchCase, wholeWord, wildcards)
    Next story
End Function

Public Function WU_CountTextInRange(ByVal target As Range, ByVal findText As String, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False, Optional ByVal wildcards As Boolean = False) As Long
    If target Is Nothing Then Err.Raise 91, "WU_CountTextInRange", "target range is required"
    WU_ValidateFind findText, "", wildcards, "WU_CountTextInRange"
    If target.End <= target.Start Then Exit Function
    WU_CountTextInRange = WU_CountIn(target, findText, matchCase, wholeWord, wildcards)
End Function

' ---- batch -----------------------------------------------------------------

' rules is a two-dimensional array whose rows hold find text and replacement
' text. Rules apply in order inside one undo record; the result is the number
' of rules that matched somewhere in the scope.
Public Function WU_ReplaceBatch(ByVal document As Document, ByVal rules As Variant, Optional ByVal scope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False, Optional ByVal wildcards As Boolean = False) As Long
    Dim stories As Collection, story As Range, storyEnd As Long, row As Long, firstRow As Long, lastRow As Long, firstColumn As Long, matched As Boolean, s As Long
    Dim updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ReplaceBatch", "document is required"
    WU_ValidateRules rules, wildcards, "WU_ReplaceBatch", firstRow, lastRow, firstColumn
    Set stories = WU_Stories(document, WU_NormalizeScope(scope, "WU_ReplaceBatch"))
    For row = firstRow To lastRow
        matched = False
        For s = 1 To stories.Count
            ' A replacement that touches a paragraph mark shrinks every Range
            ' that spanned it, so each rule re-derives the whole story.
            Set story = WU_Rebound(stories(s), 0, WU_StoryEnd(stories(s)))
            If WU_Exists(story, CStr(rules(row, firstColumn)), matchCase, wholeWord, wildcards) Then
                If Not opened Then WU_BeginSafeEdit updating, opened, captured, "Replace text batch"
                storyEnd = story.End
                If WU_ReplaceIn(story, CStr(rules(row, firstColumn)), CStr(rules(row, firstColumn + 1)), matchCase, wholeWord, wildcards, storyEnd) Then matched = True
            End If
        Next s
        If matched Then WU_ReplaceBatch = WU_ReplaceBatch + 1
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

Public Function WU_ReplaceBatchInRange(ByVal target As Range, ByVal rules As Variant, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False, Optional ByVal wildcards As Boolean = False) As Long
    Dim row As Long, firstRow As Long, lastRow As Long, firstColumn As Long, targetStart As Long, targetEnd As Long, scope As Range
    Dim updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ReplaceBatchInRange", "target range is required"
    WU_ValidateRules rules, wildcards, "WU_ReplaceBatchInRange", firstRow, lastRow, firstColumn
    targetStart = target.Start: targetEnd = target.End
    If targetEnd <= targetStart Then Exit Function
    For row = firstRow To lastRow
        ' The caller's bounds are re-derived per rule: replacement text of a
        ' different length moves the end, and a replaced paragraph mark
        ' shrinks every Range that spanned it.
        Set scope = WU_Rebound(target, targetStart, targetEnd)
        If WU_Exists(scope, CStr(rules(row, firstColumn)), matchCase, wholeWord, wildcards) Then
            If Not opened Then WU_BeginSafeEdit updating, opened, captured, "Replace text batch"
            If WU_ReplaceIn(scope, CStr(rules(row, firstColumn)), CStr(rules(row, firstColumn + 1)), matchCase, wholeWord, wildcards, targetEnd) Then WU_ReplaceBatchInRange = WU_ReplaceBatchInRange + 1
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

' ---- character styles on matches ------------------------------------------

' Apply a character style to every match; the result is the number of matches
' styled. Word resets direct character formatting inside a range when a
' character style is applied, exactly as it does in its own user interface.
Public Function WU_StyleMatches(ByVal document As Document, ByVal findText As String, ByVal styleName As String, Optional ByVal scope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False, Optional ByVal wildcards As Boolean = False) As Long
    Dim stories As Collection, story As Range, st As Style, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_StyleMatches", "document is required"
    WU_ValidateFind findText, "", wildcards, "WU_StyleMatches"
    Set st = WU_CharacterStyle(document, styleName, "WU_StyleMatches")
    Set stories = WU_Stories(document, WU_NormalizeScope(scope, "WU_StyleMatches"))
    For Each story In stories
        If WU_Exists(story, findText, matchCase, wholeWord, wildcards) Then
            If Not opened Then WU_BeginSafeEdit updating, opened, captured, "Style matches"
            WU_StyleMatches = WU_StyleMatches + WU_StyleIn(story, findText, st, matchCase, wholeWord, wildcards)
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

Public Function WU_StyleMatchesInRange(ByVal target As Range, ByVal findText As String, ByVal styleName As String, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False, Optional ByVal wildcards As Boolean = False) As Long
    Dim st As Style, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_StyleMatchesInRange", "target range is required"
    WU_ValidateFind findText, "", wildcards, "WU_StyleMatchesInRange"
    Set st = WU_CharacterStyle(target.Document, styleName, "WU_StyleMatchesInRange")
    If target.End <= target.Start Then Exit Function
    If Not WU_Exists(target, findText, matchCase, wholeWord, wildcards) Then Exit Function
    WU_BeginSafeEdit updating, opened, captured, "Style matches"
    WU_StyleMatchesInRange = WU_StyleIn(target, findText, st, matchCase, wholeWord, wildcards)
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

' ---- offset runs -------------------------------------------------------------

' runs is a two-dimensional array whose rows hold an absolute start position,
' an absolute end position and a character style name. Rows must lie inside
' target in ascending, non-overlapping order; every row is validated before
' any formatting changes. The result is the number of runs whose style changed.
Public Function WU_ApplyCharacterStyleRuns(ByVal target As Range, ByVal runs As Variant) As Long
    Dim document As Document, scope As Range, row As Long, firstRow As Long, lastRow As Long, firstColumn As Long
    Dim styleCache() As Style, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ApplyCharacterStyleRuns", "target range is required"
    Set document = target.Document
    If Not WU_ValidateRuns(document, target, runs, styleCache, firstRow, lastRow, firstColumn) Then Exit Function
    WU_BeginSafeEdit updating, opened, captured, "Apply character style runs"
    Set scope = target.Duplicate
    For row = firstRow To lastRow
        ' Extend first, then move the start; assigning a later start to a
        ' reused Range before its end can make Word reject the transient span.
        scope.End = CLng(runs(row, firstColumn + 1))
        scope.Start = CLng(runs(row, firstColumn))
        If StrComp(WU_RangeStyleName(scope), styleCache(row).NameLocal, vbTextCompare) <> 0 Then
            scope.Style = styleCache(row)
            WU_ApplyCharacterStyleRuns = WU_ApplyCharacterStyleRuns + 1
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

' ---- internals -----------------------------------------------------------------

' A fresh Range over the given bounds of scope, staying in scope's story.
' Extend first, then move the start, so a transient inverted span is never
' handed to Word.
Private Function WU_Rebound(ByVal scope As Range, ByVal startPosition As Long, ByVal endPosition As Long) As Range
    Set WU_Rebound = scope.Duplicate
    WU_Rebound.End = endPosition
    WU_Rebound.Start = startPosition
End Function

Private Sub WU_ValidateFind(ByVal findText As String, ByVal replaceText As String, ByVal wildcards As Boolean, ByVal sourceName As String)
    If Len(findText) = 0 Then Err.Raise 5, sourceName, "find text is required"
    If Len(findText) > 255 Then Err.Raise 5, sourceName, "find text exceeds Word's 255-character limit"
    If Len(replaceText) > 255 Then Err.Raise 5, sourceName, "replacement text exceeds Word's 255-character limit"
    If Not wildcards Then
        If Len(WU_EscapeLiteral(findText)) > 255 Then Err.Raise 5, sourceName, "find text exceeds Word's escaped 255-character limit"
        If Len(WU_EscapeLiteral(replaceText)) > 255 Then Err.Raise 5, sourceName, "replacement text exceeds Word's escaped 255-character limit"
    End If
End Sub

Private Sub WU_ValidateRules(ByVal rules As Variant, ByVal wildcards As Boolean, ByVal sourceName As String, ByRef firstRow As Long, ByRef lastRow As Long, ByRef firstColumn As Long)
    Dim row As Long
    If Not IsArray(rules) Then Err.Raise 5, sourceName, "rules must be a two-dimensional array of find and replacement text"
    On Error Resume Next
    firstRow = LBound(rules, 1): lastRow = UBound(rules, 1): firstColumn = LBound(rules, 2)
    If Err.Number <> 0 Or UBound(rules, 2) - firstColumn < 1 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 5, sourceName, "rules must be a two-dimensional array with find and replacement columns"
    End If
    On Error GoTo 0
    For row = firstRow To lastRow
        WU_ValidateFind CStr(rules(row, firstColumn)), CStr(rules(row, firstColumn + 1)), wildcards, sourceName
    Next row
End Sub

Private Function WU_ValidateRuns(ByVal document As Document, ByVal target As Range, ByVal runs As Variant, ByRef styleCache() As Style, ByRef firstRow As Long, ByRef lastRow As Long, ByRef firstColumn As Long) As Boolean
    Dim row As Long, startPosition As Long, endPosition As Long, previousEnd As Long
    If Not IsArray(runs) Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "runs must be a two-dimensional array of start, end and style name"
    On Error Resume Next
    firstRow = LBound(runs, 1): lastRow = UBound(runs, 1): firstColumn = LBound(runs, 2)
    If Err.Number <> 0 Or UBound(runs, 2) - firstColumn < 2 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 5, "WU_ApplyCharacterStyleRuns", "runs must be a two-dimensional array with start, end and style columns"
    End If
    On Error GoTo 0
    If lastRow < firstRow Then Exit Function
    ReDim styleCache(firstRow To lastRow)
    previousEnd = target.Start
    For row = firstRow To lastRow
        If Not IsNumeric(runs(row, firstColumn)) Or Not IsNumeric(runs(row, firstColumn + 1)) Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "run " & CStr(row) & " positions must be numeric"
        startPosition = CLng(runs(row, firstColumn)): endPosition = CLng(runs(row, firstColumn + 1))
        If startPosition < previousEnd Or endPosition <= startPosition Or endPosition > target.End Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "run " & CStr(row) & " must be an ascending, non-overlapping span inside the target range"
        Set styleCache(row) = WU_CharacterStyle(document, CStr(runs(row, firstColumn + 2)), "WU_ApplyCharacterStyleRuns")
        previousEnd = endPosition
    Next row
    WU_ValidateRuns = True
End Function

Private Function WU_CharacterStyle(ByVal document As Document, ByVal styleName As String, ByVal sourceName As String) As Style
    Dim st As Style, styleError As Long
    If Len(Trim$(styleName)) = 0 Then Err.Raise 5, sourceName, "style name is required"
    On Error Resume Next
    Set st = document.Styles(styleName)
    styleError = Err.Number
    Err.Clear
    On Error GoTo 0
    If styleError <> 0 Or st Is Nothing Then Err.Raise 5, sourceName, "style " & styleName & " was not found"
    If st.Type <> wdStyleTypeCharacter And st.Type <> wdStyleTypeLinked Then Err.Raise 5, sourceName, "style " & styleName & " is not a character style"
    Set WU_CharacterStyle = st
End Function

Private Function WU_RangeStyleName(ByVal scope As Range) As String
    On Error Resume Next
    WU_RangeStyleName = scope.Style.NameLocal
    Err.Clear
    On Error GoTo 0
End Function

Private Sub WU_PrepareFind(ByVal criteria As Find, ByVal findText As String, ByVal replaceText As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, ByVal wildcards As Boolean)
    With criteria
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = WU_Pattern(findText, wildcards)
        .Replacement.Text = WU_Pattern(replaceText, wildcards)
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = matchCase
        .MatchWholeWord = wholeWord
        .MatchWildcards = wildcards
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    WU_PinFindOptions criteria
End Sub

' The number of characters in scope's story; positions run from 0 to this.
Private Function WU_StoryEnd(ByVal scope As Range) As Long
    On Error Resume Next
    WU_StoryEnd = scope.End
    WU_StoryEnd = scope.StoryLength
    Err.Clear
    On Error GoTo 0
End Function

Private Function WU_IsWholeStory(ByVal scope As Range) As Boolean
    WU_IsWholeStory = (scope.Start = 0 And scope.End >= WU_StoryEnd(scope))
End Function

' A search range one character wider than scope on each side, clamped to the
' story. Word only recognises a word boundary it can see, so a whole-word
' search on a range that exactly equals the word finds nothing; matches are
' filtered back to the caller's bounds by WU_NextMatch.
Private Function WU_Probe(ByVal scope As Range) As Range
    Dim probe As Range
    Set probe = scope.Duplicate
    If probe.End < WU_StoryEnd(scope) Then probe.End = probe.End + 1
    If probe.Start > 0 Then probe.Start = probe.Start - 1
    Set WU_Probe = probe
End Function

' Move probe onto the next non-empty match lying inside [lowLimit, highLimit).
' Matches that start before the bounds are skipped; a match starting at or
' after highLimit ends the search.
Private Function WU_NextMatch(ByVal probe As Range, ByVal lowLimit As Long, ByVal highLimit As Long, ByVal searchEnd As Long) As Boolean
    Do While probe.Find.Execute
        If probe.Start >= highLimit Then Exit Function
        If probe.Start >= lowLimit And probe.End <= highLimit And probe.End > probe.Start Then
            WU_NextMatch = True
            Exit Function
        End If
        If Not WU_Advance(probe, searchEnd) Then Exit Function
    Loop
End Function

' Read-only existence probe so a mutating operation opens no undo record and
' leaves the document unmodified when nothing matches.
Private Function WU_Exists(ByVal scope As Range, ByVal findText As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, ByVal wildcards As Boolean) As Boolean
    Dim probe As Range
    Set probe = WU_Probe(scope)
    WU_PrepareFind probe.Find, findText, "", matchCase, wholeWord, wildcards
    WU_Exists = WU_NextMatch(probe, scope.Start, scope.End, probe.End)
End Function

' Replace every match inside scope. endPosition carries the caller's end
' bound in and the adjusted bound out, so a caller can keep an exact range
' aligned across replacements of a different length. A whole story is handed
' to ReplaceAll, which cannot overrun there. On an exact range every match is
' located first and then replaced by ReplaceOne on a probe one character
' wider than the match, the only form Word applies to that match alone.
Private Function WU_ReplaceIn(ByVal scope As Range, ByVal findText As String, ByVal replaceText As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, ByVal wildcards As Boolean, ByRef endPosition As Long) As Boolean
    Dim probe As Range, match As Range, lowLimit As Long, highLimit As Long, searchEnd As Long
    Dim matchStart As Long, matchLength As Long, storyBefore As Long, delta As Long, nextStart As Long
    If WU_IsWholeStory(scope) Then
        Set probe = scope.Duplicate
        WU_PrepareFind probe.Find, findText, replaceText, matchCase, wholeWord, wildcards
        WU_ReplaceIn = probe.Find.Execute(Replace:=wdReplaceAll)
        endPosition = WU_StoryEnd(scope)
        Exit Function
    End If
    lowLimit = scope.Start: highLimit = endPosition
    Set probe = WU_Probe(WU_Rebound(scope, lowLimit, highLimit))
    searchEnd = probe.End
    WU_PrepareFind probe.Find, findText, "", matchCase, wholeWord, wildcards
    Do While WU_NextMatch(probe, lowLimit, highLimit, searchEnd)
        matchStart = probe.Start: matchLength = probe.End - probe.Start
        storyBefore = WU_StoryEnd(probe)
        Set match = WU_Probe(probe)
        WU_PrepareFind match.Find, findText, replaceText, matchCase, wholeWord, wildcards
        If Not match.Find.Execute(Replace:=wdReplaceOne) Then Exit Do
        WU_ReplaceIn = True
        delta = WU_StoryEnd(probe) - storyBefore
        highLimit = highLimit + delta: searchEnd = searchEnd + delta
        nextStart = matchStart + matchLength + delta
        If nextStart >= highLimit Or nextStart >= searchEnd Then Exit Do
        probe.End = searchEnd
        probe.Start = nextStart
    Loop
    endPosition = highLimit
End Function

Private Function WU_CountIn(ByVal scope As Range, ByVal findText As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, ByVal wildcards As Boolean) As Long
    Dim probe As Range, searchEnd As Long
    Set probe = WU_Probe(scope)
    searchEnd = probe.End
    WU_PrepareFind probe.Find, findText, "", matchCase, wholeWord, wildcards
    Do While WU_NextMatch(probe, scope.Start, scope.End, searchEnd)
        WU_CountIn = WU_CountIn + 1
        If Not WU_Advance(probe, searchEnd) Then Exit Do
    Loop
End Function

Private Function WU_StyleIn(ByVal scope As Range, ByVal findText As String, ByVal st As Style, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, ByVal wildcards As Boolean) As Long
    Dim probe As Range, searchEnd As Long
    Set probe = WU_Probe(scope)
    searchEnd = probe.End
    WU_PrepareFind probe.Find, findText, "", matchCase, wholeWord, wildcards
    Do While WU_NextMatch(probe, scope.Start, scope.End, searchEnd)
        probe.Style = st
        WU_StyleIn = WU_StyleIn + 1
        If Not WU_Advance(probe, searchEnd) Then Exit Do
    Loop
End Function

' Move a match range past itself and restore the search boundary. A zero-width
' wildcard match steps one character so a permissive pattern cannot loop.
Private Function WU_Advance(ByVal probe As Range, ByVal limit As Long) As Boolean
    Dim position As Long
    position = probe.End
    If probe.End = probe.Start Then position = position + 1
    If position >= limit Then Exit Function
    probe.End = limit
    probe.Start = position
    WU_Advance = True
End Function

' Word keeps several Find flags sticky across interactive searches; pin them so
' results are deterministic. Options missing on older hosts are ignored.
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

Private Function WU_EscapeLiteral(ByVal value As String) As String
    WU_EscapeLiteral = Replace(value, "^", "^^")
End Function

Private Function WU_Pattern(ByVal value As String, ByVal wildcards As Boolean) As String
    If wildcards Then WU_Pattern = value Else WU_Pattern = WU_EscapeLiteral(value)
End Function
