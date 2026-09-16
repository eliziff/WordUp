Attribute VB_Name = "WordUpStyleConverter"
Option Explicit
Private Const WU_MAX_STYLE_BATCH_RULES As Long = 256
Private Const WU_MAX_STYLE_STORY_CHAIN As Long = 32768
Private Const WU_WORD_STORY_MISSING As Long = 5941
Public Function WU_ConvertStyle(ByVal document As Document, ByVal fromStyle As String, ByVal toStyle As String, Optional ByVal storyScope As String = "all") As Boolean
    Dim firstStory As Range, story As Range, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String, sourceError As Long, targetError As Long
    Dim sourceStyle As Style, targetStyle As Style
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ConvertStyle", "document is required"
    If Len(Trim$(fromStyle)) = 0 Then Err.Raise 5, "WU_ConvertStyle", "source style is required"
    If Len(Trim$(toStyle)) = 0 Then Err.Raise 5, "WU_ConvertStyle", "target style is required"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_ConvertStyle", "story scope must be main, notes, headers, footers, or all"
    updating = Application.ScreenUpdating
    captured = True
    On Error Resume Next
    Set sourceStyle = document.Styles(fromStyle)
    sourceError = Err.Number
    Err.Clear
    Set targetStyle = document.Styles(toStyle)
    targetError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If sourceError <> 0 Or sourceStyle Is Nothing Then Err.Raise 5, "WU_ConvertStyle", "source style " & fromStyle & " was not found"
    If targetError <> 0 Or targetStyle Is Nothing Then Err.Raise 5, "WU_ConvertStyle", "target style " & toStyle & " was not found"
    If sourceStyle.Type <> wdStyleTypeParagraph Then Err.Raise 5, "WU_ConvertStyle", "source style is not a paragraph style"
    If targetStyle.Type <> wdStyleTypeParagraph Then Err.Raise 5, "WU_ConvertStyle", "target style is not a paragraph style"
    If StrComp(sourceStyle.NameLocal, targetStyle.NameLocal, vbTextCompare) = 0 Then Exit Function
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Convert style": opened = True
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If WU_StyleStoryHasContent(story) Then WU_ConvertStyle = WU_ConvertStyleInStory(story, sourceStyle, targetStyle)
    ElseIf storyScope = "notes" Then
        Set firstStory = WU_StyleStory(document, wdFootnotesStory)
        If Not firstStory Is Nothing Then WU_ConvertStyle = WU_ConvertStyleInStoryChain(firstStory, sourceStyle, targetStyle)
        Set firstStory = Nothing
        Set firstStory = WU_StyleStory(document, wdEndnotesStory)
        If Not firstStory Is Nothing Then If WU_ConvertStyleInStoryChain(firstStory, sourceStyle, targetStyle) Then WU_ConvertStyle = True
    ElseIf storyScope = "headers" Then
        If WU_ConvertStyleInStoryType(document, wdPrimaryHeaderStory, sourceStyle, targetStyle) Then WU_ConvertStyle = True
        If WU_ConvertStyleInStoryType(document, wdFirstPageHeaderStory, sourceStyle, targetStyle) Then WU_ConvertStyle = True
        If WU_ConvertStyleInStoryType(document, wdEvenPagesHeaderStory, sourceStyle, targetStyle) Then WU_ConvertStyle = True
    ElseIf storyScope = "footers" Then
        If WU_ConvertStyleInStoryType(document, wdPrimaryFooterStory, sourceStyle, targetStyle) Then WU_ConvertStyle = True
        If WU_ConvertStyleInStoryType(document, wdFirstPageFooterStory, sourceStyle, targetStyle) Then WU_ConvertStyle = True
        If WU_ConvertStyleInStoryType(document, wdEvenPagesFooterStory, sourceStyle, targetStyle) Then WU_ConvertStyle = True
    Else
        For Each firstStory In document.StoryRanges
            If WU_ConvertStyleInStoryChain(firstStory, sourceStyle, targetStyle) Then WU_ConvertStyle = True
        Next firstStory
    End If
CleanUp:
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
        Err.Clear
    End If
    If captured Then Application.ScreenUpdating = updating
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

' Convert a bounded list of paragraph-style pairs in one safe edit. Rows run
' in order, so an intentional chain such as A -> B followed by B -> C is
' deterministic. Word remains the source of truth for paragraphs and styles;
' the two-column Variant array is only a compact command list.
Public Function WU_ConvertStyleBatch(ByVal document As Document, ByVal mappings As Variant, Optional ByVal storyScope As String = "all") As Long
    Dim firstStory As Range, story As Range, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, activeRows As Long, changed As Long, row As Long
    Dim sourceCache() As Style, targetCache() As Style, enabled() As Boolean, matched() As Boolean
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ConvertStyleBatch", "document is required"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_ConvertStyleBatch", "story scope must be main, notes, headers, footers, or all"
    activeRows = WU_ValidateStyleBatch(document, mappings, sourceCache, targetCache, enabled)
    If activeRows = 0 Then Exit Function
    firstRow = LBound(mappings, 1): lastRow = UBound(mappings, 1): firstColumn = LBound(mappings, 2)
    ReDim matched(firstRow To lastRow)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Convert style batch": opened = True
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If WU_StyleStoryHasContent(story) Then WU_ConvertStyleBatchInStory story, sourceCache, targetCache, enabled, firstRow, lastRow, matched
    ElseIf storyScope = "notes" Then
        Set firstStory = WU_StyleStory(document, wdFootnotesStory)
        If Not firstStory Is Nothing Then WU_ConvertStyleBatchInStoryChain firstStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        Set firstStory = Nothing
        Set firstStory = WU_StyleStory(document, wdEndnotesStory)
        If Not firstStory Is Nothing Then WU_ConvertStyleBatchInStoryChain firstStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
    ElseIf storyScope = "headers" Then
        WU_ConvertStyleBatchInStoryType document, wdPrimaryHeaderStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        WU_ConvertStyleBatchInStoryType document, wdFirstPageHeaderStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        WU_ConvertStyleBatchInStoryType document, wdEvenPagesHeaderStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
    ElseIf storyScope = "footers" Then
        WU_ConvertStyleBatchInStoryType document, wdPrimaryFooterStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        WU_ConvertStyleBatchInStoryType document, wdFirstPageFooterStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        WU_ConvertStyleBatchInStoryType document, wdEvenPagesFooterStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
    Else
        For Each firstStory In document.StoryRanges
            WU_ConvertStyleBatchInStoryChain firstStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        Next firstStory
    End If
    ' Re-derive each row scope from the saved span: a row that replaces the
    ' paragraph mark at the range end makes Word adjust the caller Range to
    ' exclude the replacement mark, so reusing the Range object would hide
    ' it from later rows.
    For row = firstRow To lastRow
        If matched(row) Then changed = changed + 1
    Next row
    WU_ConvertStyleBatch = changed
CleanUp:
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
        Err.Clear
    End If
    If captured Then Application.ScreenUpdating = updating
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

' Convert only paragraphs inside the exact caller-supplied Range. A caller
' can therefore limit a large manuscript pass to a section, table cell, or
' generated opening without re-enumerating every Word story.
Public Function WU_ConvertStyleInRange(ByVal target As Range, ByVal fromStyle As String, ByVal toStyle As String) As Boolean
    Dim updating As Boolean, opened As Boolean, captured As Boolean, targetStart As Long, targetEnd As Long
    Dim failure As Long, failureSource As String, failureText As String, sourceError As Long, targetError As Long
    Dim document As Document, sourceStyle As Style, targetStyle As Style
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ConvertStyleInRange", "target range is required"
    If Len(Trim$(fromStyle)) = 0 Then Err.Raise 5, "WU_ConvertStyleInRange", "source style is required"
    If Len(Trim$(toStyle)) = 0 Then Err.Raise 5, "WU_ConvertStyleInRange", "target style is required"
    Set document = target.Document
    On Error Resume Next
    Set sourceStyle = document.Styles(fromStyle)
    sourceError = Err.Number
    Err.Clear
    Set targetStyle = document.Styles(toStyle)
    targetError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If sourceError <> 0 Or sourceStyle Is Nothing Then Err.Raise 5, "WU_ConvertStyleInRange", "source style " & fromStyle & " was not found"
    If targetError <> 0 Or targetStyle Is Nothing Then Err.Raise 5, "WU_ConvertStyleInRange", "target style " & toStyle & " was not found"
    If sourceStyle.Type <> wdStyleTypeParagraph Then Err.Raise 5, "WU_ConvertStyleInRange", "source style is not a paragraph style"
    If targetStyle.Type <> wdStyleTypeParagraph Then Err.Raise 5, "WU_ConvertStyleInRange", "target style is not a paragraph style"
    targetStart = target.Start: targetEnd = target.End
    If targetEnd <= targetStart Then Exit Function
    If StrComp(sourceStyle.NameLocal, targetStyle.NameLocal, vbTextCompare) = 0 Then Exit Function
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Convert style": opened = True
    WU_ConvertStyleInRange = WU_ConvertStyleInStory(target, sourceStyle, targetStyle)
CleanUp:
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
        Err.Clear
    End If
    If captured Then Application.ScreenUpdating = updating
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

' Convert a character or linked style across explicitly selected stories.
' Word's formatting Find engine applies the destination style without
' rewriting the matched text, so existing italic, language, hyperlinks,
' fields, and other inline structure remain owned by Word.
Public Function WU_ConvertCharacterStyle(ByVal document As Document, ByVal fromStyle As String, ByVal toStyle As String, Optional ByVal storyScope As String = "all") As Boolean
    Dim firstStory As Range, story As Range, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String, sourceError As Long, targetError As Long
    Dim sourceStyle As Style, targetStyle As Style
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ConvertCharacterStyle", "document is required"
    If Len(Trim$(fromStyle)) = 0 Then Err.Raise 5, "WU_ConvertCharacterStyle", "source style is required"
    If Len(Trim$(toStyle)) = 0 Then Err.Raise 5, "WU_ConvertCharacterStyle", "target style is required"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_ConvertCharacterStyle", "story scope must be main, notes, headers, footers, or all"
    updating = Application.ScreenUpdating
    captured = True
    On Error Resume Next
    Set sourceStyle = document.Styles(fromStyle)
    sourceError = Err.Number
    Err.Clear
    Set targetStyle = document.Styles(toStyle)
    targetError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If sourceError <> 0 Or sourceStyle Is Nothing Then Err.Raise 5, "WU_ConvertCharacterStyle", "source style " & fromStyle & " was not found"
    If targetError <> 0 Or targetStyle Is Nothing Then Err.Raise 5, "WU_ConvertCharacterStyle", "target style " & toStyle & " was not found"
    If Not WU_IsCharacterStyle(sourceStyle) Then Err.Raise 5, "WU_ConvertCharacterStyle", "source style is not a character style"
    If Not WU_IsCharacterStyle(targetStyle) Then Err.Raise 5, "WU_ConvertCharacterStyle", "target style is not a character style"
    If StrComp(sourceStyle.NameLocal, targetStyle.NameLocal, vbTextCompare) = 0 Then Exit Function
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Convert character style": opened = True
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If WU_StyleStoryHasContent(story) Then WU_ConvertCharacterStyle = WU_ConvertCharacterStyleInStory(story, sourceStyle, targetStyle)
    ElseIf storyScope = "notes" Then
        Set firstStory = WU_StyleStory(document, wdFootnotesStory)
        If Not firstStory Is Nothing Then WU_ConvertCharacterStyle = WU_ConvertCharacterStyleInStoryChain(firstStory, sourceStyle, targetStyle)
        Set firstStory = Nothing
        Set firstStory = WU_StyleStory(document, wdEndnotesStory)
        If Not firstStory Is Nothing Then If WU_ConvertCharacterStyleInStoryChain(firstStory, sourceStyle, targetStyle) Then WU_ConvertCharacterStyle = True
    ElseIf storyScope = "headers" Then
        If WU_ConvertCharacterStyleInStoryType(document, wdPrimaryHeaderStory, sourceStyle, targetStyle) Then WU_ConvertCharacterStyle = True
        If WU_ConvertCharacterStyleInStoryType(document, wdFirstPageHeaderStory, sourceStyle, targetStyle) Then WU_ConvertCharacterStyle = True
        If WU_ConvertCharacterStyleInStoryType(document, wdEvenPagesHeaderStory, sourceStyle, targetStyle) Then WU_ConvertCharacterStyle = True
    ElseIf storyScope = "footers" Then
        If WU_ConvertCharacterStyleInStoryType(document, wdPrimaryFooterStory, sourceStyle, targetStyle) Then WU_ConvertCharacterStyle = True
        If WU_ConvertCharacterStyleInStoryType(document, wdFirstPageFooterStory, sourceStyle, targetStyle) Then WU_ConvertCharacterStyle = True
        If WU_ConvertCharacterStyleInStoryType(document, wdEvenPagesFooterStory, sourceStyle, targetStyle) Then WU_ConvertCharacterStyle = True
    Else
        For Each firstStory In document.StoryRanges
            If WU_ConvertCharacterStyleInStoryChain(firstStory, sourceStyle, targetStyle) Then WU_ConvertCharacterStyle = True
        Next firstStory
    End If
CleanUp:
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
        Err.Clear
    End If
    If captured Then Application.ScreenUpdating = updating
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

' Convert a character or linked style only inside the exact caller-supplied
' Range. Empty ranges and same-style requests are true no-ops: they do not
' open an undo record or dirty the document.
Public Function WU_ConvertCharacterStyleInRange(ByVal target As Range, ByVal fromStyle As String, ByVal toStyle As String) As Boolean
    Dim updating As Boolean, opened As Boolean, captured As Boolean, targetStart As Long, targetEnd As Long
    Dim failure As Long, failureSource As String, failureText As String, sourceError As Long, targetError As Long
    Dim document As Document, sourceStyle As Style, targetStyle As Style
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ConvertCharacterStyleInRange", "target range is required"
    If Len(Trim$(fromStyle)) = 0 Then Err.Raise 5, "WU_ConvertCharacterStyleInRange", "source style is required"
    If Len(Trim$(toStyle)) = 0 Then Err.Raise 5, "WU_ConvertCharacterStyleInRange", "target style is required"
    Set document = target.Document
    On Error Resume Next
    Set sourceStyle = document.Styles(fromStyle)
    sourceError = Err.Number
    Err.Clear
    Set targetStyle = document.Styles(toStyle)
    targetError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If sourceError <> 0 Or sourceStyle Is Nothing Then Err.Raise 5, "WU_ConvertCharacterStyleInRange", "source style " & fromStyle & " was not found"
    If targetError <> 0 Or targetStyle Is Nothing Then Err.Raise 5, "WU_ConvertCharacterStyleInRange", "target style " & toStyle & " was not found"
    If Not WU_IsCharacterStyle(sourceStyle) Then Err.Raise 5, "WU_ConvertCharacterStyleInRange", "source style is not a character style"
    If Not WU_IsCharacterStyle(targetStyle) Then Err.Raise 5, "WU_ConvertCharacterStyleInRange", "target style is not a character style"
    targetStart = target.Start: targetEnd = target.End
    If targetEnd <= targetStart Then Exit Function
    If StrComp(sourceStyle.NameLocal, targetStyle.NameLocal, vbTextCompare) = 0 Then Exit Function
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Convert character style": opened = True
    WU_ConvertCharacterStyleInRange = WU_ConvertCharacterStyleInStory(target, sourceStyle, targetStyle)
CleanUp:
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
        Err.Clear
    End If
    If captured Then Application.ScreenUpdating = updating
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

' Convert a bounded map of character or linked styles in row order. The
' validated Style handles are reused across stories, and every mapping shares
' one undo record. Row order is intentional when a source style is itself a
' destination of an earlier mapping.
Public Function WU_ConvertCharacterStyleBatch(ByVal document As Document, ByVal mappings As Variant, Optional ByVal storyScope As String = "all") As Long
    Dim firstStory As Range, story As Range, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, activeRows As Long, changed As Long, row As Long
    Dim sourceCache() As Style, targetCache() As Style, enabled() As Boolean, matched() As Boolean
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ConvertCharacterStyleBatch", "document is required"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_ConvertCharacterStyleBatch", "story scope must be main, notes, headers, footers, or all"
    activeRows = WU_ValidateStyleBatch(document, mappings, sourceCache, targetCache, enabled, True, "WU_ConvertCharacterStyleBatch")
    If activeRows = 0 Then Exit Function
    firstRow = LBound(mappings, 1): lastRow = UBound(mappings, 1): firstColumn = LBound(mappings, 2)
    ReDim matched(firstRow To lastRow)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Convert character style batch": opened = True
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If WU_StyleStoryHasContent(story) Then WU_ConvertCharacterStyleBatchInStory story, sourceCache, targetCache, enabled, firstRow, lastRow, matched
    ElseIf storyScope = "notes" Then
        Set firstStory = WU_StyleStory(document, wdFootnotesStory)
        If Not firstStory Is Nothing Then WU_ConvertCharacterStyleBatchInStoryChain firstStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        Set firstStory = Nothing
        Set firstStory = WU_StyleStory(document, wdEndnotesStory)
        If Not firstStory Is Nothing Then WU_ConvertCharacterStyleBatchInStoryChain firstStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
    ElseIf storyScope = "headers" Then
        WU_ConvertCharacterStyleBatchInStoryType document, wdPrimaryHeaderStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        WU_ConvertCharacterStyleBatchInStoryType document, wdFirstPageHeaderStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        WU_ConvertCharacterStyleBatchInStoryType document, wdEvenPagesHeaderStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
    ElseIf storyScope = "footers" Then
        WU_ConvertCharacterStyleBatchInStoryType document, wdPrimaryFooterStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        WU_ConvertCharacterStyleBatchInStoryType document, wdFirstPageFooterStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        WU_ConvertCharacterStyleBatchInStoryType document, wdEvenPagesFooterStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
    Else
        For Each firstStory In document.StoryRanges
            WU_ConvertCharacterStyleBatchInStoryChain firstStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        Next firstStory
    End If
    For row = firstRow To lastRow
        If matched(row) Then changed = changed + 1
    Next row
    WU_ConvertCharacterStyleBatch = changed
CleanUp:
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
        Err.Clear
    End If
    If captured Then Application.ScreenUpdating = updating
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

' Convert the same character-style map only inside an exact Range. Validation
' completes before state changes or the undo record are opened.
Public Function WU_ConvertCharacterStyleBatchInRange(ByVal target As Range, ByVal mappings As Variant) As Long
    Dim document As Document, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim targetStart As Long, targetEnd As Long, firstRow As Long, lastRow As Long, firstColumn As Long, activeRows As Long, row As Long, changed As Long
    Dim sourceCache() As Style, targetCache() As Style, enabled() As Boolean, matched() As Boolean
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ConvertCharacterStyleBatchInRange", "target range is required"
    Set document = target.Document
    activeRows = WU_ValidateStyleBatch(document, mappings, sourceCache, targetCache, enabled, True, "WU_ConvertCharacterStyleBatchInRange")
    If activeRows = 0 Then Exit Function
    targetStart = target.Start: targetEnd = target.End
    If targetEnd <= targetStart Then Exit Function
    firstRow = LBound(mappings, 1): lastRow = UBound(mappings, 1): firstColumn = LBound(mappings, 2)
    ReDim matched(firstRow To lastRow)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Convert character style batch": opened = True
    WU_ConvertCharacterStyleBatchInStory target, sourceCache, targetCache, enabled, firstRow, lastRow, matched
    For row = firstRow To lastRow
        If matched(row) Then changed = changed + 1
    Next row
    WU_ConvertCharacterStyleBatchInRange = changed
CleanUp:
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
        Err.Clear
    End If
    If captured Then Application.ScreenUpdating = updating
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

' Convert the same bounded style map only inside the exact caller-supplied
' Range. The range is never widened to a story or Selection.
Public Function WU_ConvertStyleBatchInRange(ByVal target As Range, ByVal mappings As Variant) As Long
    Dim document As Document, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim targetStart As Long, targetEnd As Long, firstRow As Long, lastRow As Long, firstColumn As Long, activeRows As Long, row As Long, changed As Long
    Dim rowScope As Range
    Dim sourceCache() As Style, targetCache() As Style, enabled() As Boolean, matched() As Boolean
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ConvertStyleBatchInRange", "target range is required"
    Set document = target.Document
    activeRows = WU_ValidateStyleBatch(document, mappings, sourceCache, targetCache, enabled)
    If activeRows = 0 Then Exit Function
    targetStart = target.Start: targetEnd = target.End
    If targetEnd <= targetStart Then Exit Function
    firstRow = LBound(mappings, 1): lastRow = UBound(mappings, 1): firstColumn = LBound(mappings, 2)
    ReDim matched(firstRow To lastRow)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Convert style batch": opened = True
    For row = firstRow To lastRow
        If enabled(row) Then
            ' Re-derive from the caller's Range so a note, header, footer, or
            ' text-frame story never gets rebuilt as a main-story Document
            ' range. Extend first, then move the start to avoid Word rejecting
            ' a transient inverted span after a prior row changed state.
            Set rowScope = target.Duplicate
            rowScope.End = targetEnd: rowScope.Start = targetStart
            If WU_ConvertStyleInStory(rowScope, sourceCache(row), targetCache(row)) Then matched(row) = True
        End If
    Next row
    For row = firstRow To lastRow
        If matched(row) Then changed = changed + 1
    Next row
    WU_ConvertStyleBatchInRange = changed
CleanUp:
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
        Err.Clear
    End If
    If captured Then Application.ScreenUpdating = updating
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

Private Function WU_ValidateStyleBatch(ByVal document As Document, ByVal mappings As Variant, ByRef sourceCache() As Style, ByRef targetCache() As Style, ByRef enabled() As Boolean, Optional ByVal characterStyles As Boolean = False, Optional ByVal sourceName As String = "WU_ConvertStyleBatch") As Long
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, lastColumn As Long, row As Long, dimensionError As Long
    Dim fromStyle As String, toStyle As String, sourceStyle As Style, targetStyle As Style, sourceError As Long, targetError As Long, activeRows As Long
    Dim cachedSourceNames() As String, cachedTargetNames() As String, cachedSourceStyles() As Style, cachedTargetStyles() As Style
    Dim sourceCacheCount As Long, targetCacheCount As Long, sourceCacheIndex As Long, targetCacheIndex As Long, cacheRow As Long, cacheCapacity As Long
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If Not IsArray(mappings) Then Err.Raise 5, sourceName, "mappings must be a two-dimensional array"
    On Error Resume Next
    firstRow = LBound(mappings, 1): lastRow = UBound(mappings, 1)
    firstColumn = LBound(mappings, 2): lastColumn = UBound(mappings, 2)
    dimensionError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If dimensionError <> 0 Then Err.Raise 5, sourceName, "mappings must be a two-dimensional array"
    If lastColumn - firstColumn + 1 <> 2 Then Err.Raise 5, sourceName, "mappings must have exactly two columns"
    If lastRow - firstRow + 1 > WU_MAX_STYLE_BATCH_RULES Then Err.Raise 5, sourceName, "style mapping count exceeds 256"
    cacheCapacity = lastRow - firstRow + 1
    ReDim sourceCache(firstRow To lastRow): ReDim targetCache(firstRow To lastRow): ReDim enabled(firstRow To lastRow)
    ReDim cachedSourceNames(1 To cacheCapacity): ReDim cachedTargetNames(1 To cacheCapacity)
    ReDim cachedSourceStyles(1 To cacheCapacity): ReDim cachedTargetStyles(1 To cacheCapacity)
    For row = firstRow To lastRow
        If IsError(mappings(row, firstColumn)) Or IsNull(mappings(row, firstColumn)) Or IsObject(mappings(row, firstColumn)) Or IsArray(mappings(row, firstColumn)) Then Err.Raise 5, sourceName, "style mapping " & CStr(row) & " source must be scalar"
        If IsError(mappings(row, firstColumn + 1)) Or IsNull(mappings(row, firstColumn + 1)) Or IsObject(mappings(row, firstColumn + 1)) Or IsArray(mappings(row, firstColumn + 1)) Then Err.Raise 5, sourceName, "style mapping " & CStr(row) & " target must be scalar"
        fromStyle = CStr(mappings(row, firstColumn))
        toStyle = CStr(mappings(row, firstColumn + 1))
        If Len(Trim$(fromStyle)) = 0 Then Err.Raise 5, sourceName, "style mapping " & CStr(row) & " source is required"
        If Len(Trim$(toStyle)) = 0 Then Err.Raise 5, sourceName, "style mapping " & CStr(row) & " target is required"
        sourceCacheIndex = 0
        For cacheRow = 1 To sourceCacheCount
            If StrComp(fromStyle, cachedSourceNames(cacheRow), vbTextCompare) = 0 Then sourceCacheIndex = cacheRow: Exit For
        Next cacheRow
        If sourceCacheIndex = 0 Then
            Set sourceStyle = Nothing: sourceError = 0
            On Error Resume Next
            Set sourceStyle = document.Styles(fromStyle)
            sourceError = Err.Number
            Err.Clear
            On Error GoTo Failed
            If sourceError <> 0 Or sourceStyle Is Nothing Then Err.Raise 5, sourceName, "style mapping " & CStr(row) & " names a missing source style"
            sourceCacheCount = sourceCacheCount + 1: sourceCacheIndex = sourceCacheCount
            cachedSourceNames(sourceCacheIndex) = fromStyle: Set cachedSourceStyles(sourceCacheIndex) = sourceStyle
        Else
            Set sourceStyle = cachedSourceStyles(sourceCacheIndex)
        End If
        If characterStyles Then
            If Not WU_IsCharacterStyle(sourceStyle) Then Err.Raise 5, sourceName, "style mapping " & CStr(row) & " source is not a character style"
        Else
            If sourceStyle.Type <> wdStyleTypeParagraph Then Err.Raise 5, sourceName, "style mapping " & CStr(row) & " source is not a paragraph style"
        End If
        targetCacheIndex = 0
        For cacheRow = 1 To targetCacheCount
            If StrComp(toStyle, cachedTargetNames(cacheRow), vbTextCompare) = 0 Then targetCacheIndex = cacheRow: Exit For
        Next cacheRow
        If targetCacheIndex = 0 Then
            Set targetStyle = Nothing: targetError = 0
            On Error Resume Next
            Set targetStyle = document.Styles(toStyle)
            targetError = Err.Number
            Err.Clear
            On Error GoTo Failed
            If targetError <> 0 Or targetStyle Is Nothing Then Err.Raise 5, sourceName, "style mapping " & CStr(row) & " names a missing target style"
            targetCacheCount = targetCacheCount + 1: targetCacheIndex = targetCacheCount
            cachedTargetNames(targetCacheIndex) = toStyle: Set cachedTargetStyles(targetCacheIndex) = targetStyle
        Else
            Set targetStyle = cachedTargetStyles(targetCacheIndex)
        End If
        If characterStyles Then
            If Not WU_IsCharacterStyle(targetStyle) Then Err.Raise 5, sourceName, "style mapping " & CStr(row) & " target is not a character style"
        Else
            If targetStyle.Type <> wdStyleTypeParagraph Then Err.Raise 5, sourceName, "style mapping " & CStr(row) & " target is not a paragraph style"
        End If
        Set sourceCache(row) = sourceStyle: Set targetCache(row) = targetStyle
        enabled(row) = (StrComp(sourceStyle.NameLocal, targetStyle.NameLocal, vbTextCompare) <> 0)
        If enabled(row) Then activeRows = activeRows + 1
    Next row
    WU_ValidateStyleBatch = activeRows
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
End Function

Private Sub WU_ConvertStyleBatchInStory(ByVal story As Range, ByRef sourceCache() As Style, ByRef targetCache() As Style, ByRef enabled() As Boolean, ByVal firstRow As Long, ByVal lastRow As Long, ByRef matched() As Boolean)
    Dim row As Long
    For row = firstRow To lastRow
        If enabled(row) Then If WU_ConvertStyleInStory(story, sourceCache(row), targetCache(row)) Then matched(row) = True
    Next row
End Sub

Private Function WU_ConvertStyleInStoryChain(ByVal firstStory As Range, ByVal sourceStyle As Style, ByVal targetStyle As Style) As Boolean
    Dim story As Range, changed As Boolean, chainLength As Long
    Set story = firstStory
    Do While Not story Is Nothing
        chainLength = chainLength + 1
        If chainLength > WU_MAX_STYLE_STORY_CHAIN Then Err.Raise 5, "WU_ConvertStyleInStoryChain", "story chain exceeds 32768 linked stories"
        If WU_StyleStoryHasContent(story) Then If WU_ConvertStyleInStory(story, sourceStyle, targetStyle) Then changed = True
        Set story = WU_StyleNextStory(story)
    Loop
    WU_ConvertStyleInStoryChain = changed
End Function

' Return False when a document has no story of the requested header/footer
' type. Word raises for absent first/even-page stories, so the optional story
' is probed once and the caller's error handler is restored before editing.
Private Function WU_ConvertStyleInStoryType(ByVal document As Document, ByVal storyType As Long, ByVal sourceStyle As Style, ByVal targetStyle As Style) As Boolean
    Dim firstStory As Range
    Set firstStory = WU_StyleStory(document, storyType)
    If Not firstStory Is Nothing Then WU_ConvertStyleInStoryType = WU_ConvertStyleInStoryChain(firstStory, sourceStyle, targetStyle)
End Function

Private Sub WU_ConvertStyleBatchInStoryType(ByVal document As Document, ByVal storyType As Long, ByRef sourceCache() As Style, ByRef targetCache() As Style, ByRef enabled() As Boolean, ByVal firstRow As Long, ByVal lastRow As Long, ByRef matched() As Boolean)
    Dim firstStory As Range
    Set firstStory = WU_StyleStory(document, storyType)
    If Not firstStory Is Nothing Then WU_ConvertStyleBatchInStoryChain firstStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
End Sub

Private Sub WU_ConvertStyleBatchInStoryChain(ByVal firstStory As Range, ByRef sourceCache() As Style, ByRef targetCache() As Style, ByRef enabled() As Boolean, ByVal firstRow As Long, ByVal lastRow As Long, ByRef matched() As Boolean)
    Dim story As Range, chainLength As Long
    Set story = firstStory
    Do While Not story Is Nothing
        chainLength = chainLength + 1
        If chainLength > WU_MAX_STYLE_STORY_CHAIN Then Err.Raise 5, "WU_ConvertStyleBatchInStoryChain", "story chain exceeds 32768 linked stories"
        If WU_StyleStoryHasContent(story) Then WU_ConvertStyleBatchInStory story, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        Set story = WU_StyleNextStory(story)
    Loop
End Sub

Private Sub WU_ConvertCharacterStyleBatchInStory(ByVal story As Range, ByRef sourceCache() As Style, ByRef targetCache() As Style, ByRef enabled() As Boolean, ByVal firstRow As Long, ByVal lastRow As Long, ByRef matched() As Boolean)
    Dim row As Long
    For row = firstRow To lastRow
        If enabled(row) Then If WU_ConvertCharacterStyleInStory(story, sourceCache(row), targetCache(row)) Then matched(row) = True
    Next row
End Sub

Private Sub WU_ConvertCharacterStyleBatchInStoryChain(ByVal firstStory As Range, ByRef sourceCache() As Style, ByRef targetCache() As Style, ByRef enabled() As Boolean, ByVal firstRow As Long, ByVal lastRow As Long, ByRef matched() As Boolean)
    Dim story As Range, chainLength As Long
    Set story = firstStory
    Do While Not story Is Nothing
        chainLength = chainLength + 1
        If chainLength > WU_MAX_STYLE_STORY_CHAIN Then Err.Raise 5, "WU_ConvertCharacterStyleBatchInStoryChain", "story chain exceeds 32768 linked stories"
        If WU_StyleStoryHasContent(story) Then WU_ConvertCharacterStyleBatchInStory story, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        Set story = WU_StyleNextStory(story)
    Loop
End Sub

Private Sub WU_ConvertCharacterStyleBatchInStoryType(ByVal document As Document, ByVal storyType As Long, ByRef sourceCache() As Style, ByRef targetCache() As Style, ByRef enabled() As Boolean, ByVal firstRow As Long, ByVal lastRow As Long, ByRef matched() As Boolean)
    Dim firstStory As Range
    Set firstStory = WU_StyleStory(document, storyType)
    If Not firstStory Is Nothing Then WU_ConvertCharacterStyleBatchInStoryChain firstStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
End Sub

Private Function WU_ConvertStyleInStory(ByVal story As Range, ByVal sourceStyle As Style, ByVal targetStyle As Style) As Boolean
    Dim scope As Range
    If story Is Nothing Then Exit Function
    If story.End <= story.Start Then Exit Function
    Set scope = story.Duplicate
    With scope.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^p"
        .Style = sourceStyle
        .Replacement.Text = "^&"
        .Replacement.Style = targetStyle
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Call WU_PinFindOptions(scope.Find)
WU_ConvertStyleInStory = scope.Find.Execute(Replace:=wdReplaceAll)
End Function

' Find's formatting-only mode is the native fast path for character styles:
' an empty text criterion plus Style selects the existing styled spans, while
' Replacement.Style changes only that formatting and leaves text/runs intact.
Private Function WU_ConvertCharacterStyleInStory(ByVal story As Range, ByVal sourceStyle As Style, ByVal targetStyle As Style) As Boolean
    Dim scope As Range
    If story Is Nothing Then Exit Function
    If story.End <= story.Start Then Exit Function
    Set scope = story.Duplicate
    With scope.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = vbNullString
        .Replacement.Text = vbNullString
        .Style = sourceStyle
        .Replacement.Style = targetStyle
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Call WU_PinFindOptions(scope.Find)
    WU_ConvertCharacterStyleInStory = scope.Find.Execute(Replace:=wdReplaceAll)
End Function

Private Function WU_ConvertCharacterStyleInStoryChain(ByVal firstStory As Range, ByVal sourceStyle As Style, ByVal targetStyle As Style) As Boolean
    Dim story As Range, changed As Boolean, chainLength As Long
    Set story = firstStory
    Do While Not story Is Nothing
        chainLength = chainLength + 1
        If chainLength > WU_MAX_STYLE_STORY_CHAIN Then Err.Raise 5, "WU_ConvertCharacterStyleInStoryChain", "story chain exceeds 32768 linked stories"
        If WU_StyleStoryHasContent(story) Then If WU_ConvertCharacterStyleInStory(story, sourceStyle, targetStyle) Then changed = True
        Set story = WU_StyleNextStory(story)
    Loop
    WU_ConvertCharacterStyleInStoryChain = changed
End Function

Private Function WU_ConvertCharacterStyleInStoryType(ByVal document As Document, ByVal storyType As Long, ByVal sourceStyle As Style, ByVal targetStyle As Style) As Boolean
    Dim firstStory As Range
    Set firstStory = WU_StyleStory(document, storyType)
    If Not firstStory Is Nothing Then WU_ConvertCharacterStyleInStoryType = WU_ConvertCharacterStyleInStoryChain(firstStory, sourceStyle, targetStyle)
End Function

Private Function WU_IsCharacterStyle(ByVal style As Style) As Boolean
    If style Is Nothing Then Exit Function
    WU_IsCharacterStyle = (style.Type = wdStyleTypeCharacter)
    If WU_IsCharacterStyle Then Exit Function
    ' Reading Linked raises 5891 on plain character styles, and the property is
    ' only needed to recognize linked paragraph styles. Guard the read so an
    ' unavailable property reports "not a character style" instead of failing.
    On Error Resume Next
    WU_IsCharacterStyle = style.Linked
    Err.Clear
    On Error GoTo 0
End Function

' Apply a paragraph style to one exact Range. Word applies a paragraph style
' to every paragraph touched by the range; the range itself is never widened
' and direct character formatting remains in place.
Public Function WU_ApplyParagraphStyleInRange(ByVal target As Range, ByVal styleName As String) As Boolean
    Dim document As Document, style As Style, scope As Range, expectedName As String
    Dim updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String, styleError As Long
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ApplyParagraphStyleInRange", "target range is required"
    If Len(Trim$(styleName)) = 0 Then Err.Raise 5, "WU_ApplyParagraphStyleInRange", "style name is required"
    Set document = target.Document
    On Error Resume Next
    Set style = document.Styles(styleName)
    styleError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If styleError <> 0 Or style Is Nothing Then Err.Raise 5, "WU_ApplyParagraphStyleInRange", "style " & styleName & " was not found"
    If style.Type <> wdStyleTypeParagraph Then Err.Raise 5, "WU_ApplyParagraphStyleInRange", "style is not a paragraph style"
    If target.End <= target.Start Then Exit Function
    expectedName = style.NameLocal
    If WU_ParagraphStyleMatches(target, style, expectedName) Then Exit Function
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Apply paragraph style": opened = True
    Set scope = target.Duplicate
    scope.Style = style
    WU_ApplyParagraphStyleInRange = True
CleanUp:
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
        Err.Clear
    End If
    If captured Then Application.ScreenUpdating = updating
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

' Apply paragraph styles to bounded, ordered source ranges in one edit. Each
' row is [absoluteStart, absoluteEnd, styleName] and must fall inside target;
' offsets are Word story positions, so detector output can be passed without
' copying paragraph text through a second representation.
Public Function WU_ApplyParagraphStyleRuns(ByVal target As Range, ByVal runs As Variant) As Long
    Dim document As Document, scope As Range, style As Style
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, row As Long, activeRows As Long, changed As Long
    Dim startPosition As Long, endPosition As Long, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim styleCache() As Style, styleNames() As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ApplyParagraphStyleRuns", "target range is required"
    Set document = target.Document
    activeRows = WU_ValidateParagraphStyleRuns(document, target, runs, styleCache, styleNames)
    If activeRows = 0 Then Exit Function
    firstRow = LBound(runs, 1): lastRow = UBound(runs, 1): firstColumn = LBound(runs, 2)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Apply paragraph style runs": opened = True
    Set scope = target.Duplicate
    For row = firstRow To lastRow
        startPosition = CLng(runs(row, firstColumn))
        endPosition = CLng(runs(row, firstColumn + 1))
        ' Extend first, then move the start; assigning a later start to a
        ' reused Range before its end can make Word reject the transient span.
        scope.End = endPosition: scope.Start = startPosition
        Set style = styleCache(row)
        If Not WU_ParagraphStyleMatches(scope, style, styleNames(row)) Then
            scope.Style = style
            changed = changed + 1
        End If
    Next row
    WU_ApplyParagraphStyleRuns = changed
CleanUp:
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
        Err.Clear
    End If
    If captured Then Application.ScreenUpdating = updating
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function

Private Function WU_ValidateParagraphStyleRuns(ByVal document As Document, ByVal target As Range, ByVal runs As Variant, ByRef styleCache() As Style, ByRef styleNames() As String) As Long
    Const WU_MAX_STYLE_RUNS As Long = 4096
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, lastColumn As Long, row As Long, dimensionError As Long
    Dim targetStart As Long, targetEnd As Long, startPosition As Long, endPosition As Long, previousEnd As Long
    Dim styleName As String, style As Style, styleError As Long
    Dim cachedNames() As String, cachedStyles() As Style, cachedCount As Long, cacheIndex As Long, cacheRow As Long, cacheCapacity As Long
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If Not IsArray(runs) Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "runs must be a two-dimensional array"
    On Error Resume Next
    firstRow = LBound(runs, 1): lastRow = UBound(runs, 1)
    firstColumn = LBound(runs, 2): lastColumn = UBound(runs, 2)
    dimensionError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If dimensionError <> 0 Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "runs must be a two-dimensional array"
    If lastColumn - firstColumn + 1 <> 3 Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "runs must have exactly three columns"
    If lastRow - firstRow + 1 > WU_MAX_STYLE_RUNS Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "style run count exceeds 4096"
    targetStart = target.Start: targetEnd = target.End: previousEnd = targetStart
    cacheCapacity = lastRow - firstRow + 1
    ReDim styleCache(firstRow To lastRow): ReDim styleNames(firstRow To lastRow)
    ReDim cachedNames(1 To cacheCapacity): ReDim cachedStyles(1 To cacheCapacity)
    For row = firstRow To lastRow
        If IsError(runs(row, firstColumn)) Or IsNull(runs(row, firstColumn)) Or IsEmpty(runs(row, firstColumn)) Or IsObject(runs(row, firstColumn)) Or IsArray(runs(row, firstColumn)) Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "style run " & CStr(row) & " start must be scalar"
        If IsError(runs(row, firstColumn + 1)) Or IsNull(runs(row, firstColumn + 1)) Or IsEmpty(runs(row, firstColumn + 1)) Or IsObject(runs(row, firstColumn + 1)) Or IsArray(runs(row, firstColumn + 1)) Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "style run " & CStr(row) & " end must be scalar"
        If IsError(runs(row, firstColumn + 2)) Or IsNull(runs(row, firstColumn + 2)) Or IsEmpty(runs(row, firstColumn + 2)) Or IsObject(runs(row, firstColumn + 2)) Or IsArray(runs(row, firstColumn + 2)) Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "style run " & CStr(row) & " style name must be scalar"
        If Not WU_ReadStylePosition(runs(row, firstColumn), startPosition) Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "style run " & CStr(row) & " start must be an integer position"
        If Not WU_ReadStylePosition(runs(row, firstColumn + 1), endPosition) Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "style run " & CStr(row) & " end must be an integer position"
        If startPosition < targetStart Or endPosition > targetEnd Or endPosition <= startPosition Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "style run " & CStr(row) & " is outside the target range"
        If startPosition < previousEnd Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "style runs must be ordered and non-overlapping"
        styleName = CStr(runs(row, firstColumn + 2))
        If Len(Trim$(styleName)) = 0 Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "style run " & CStr(row) & " style name is required"
        cacheIndex = 0
        For cacheRow = 1 To cachedCount
            If StrComp(styleName, cachedNames(cacheRow), vbTextCompare) = 0 Then cacheIndex = cacheRow: Exit For
        Next cacheRow
        If cacheIndex = 0 Then
            Set style = Nothing
            styleError = 0
            On Error Resume Next
            Set style = document.Styles(styleName)
            styleError = Err.Number
            Err.Clear
            On Error GoTo Failed
            If styleError <> 0 Or style Is Nothing Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "style run " & CStr(row) & " names a missing style"
            cachedCount = cachedCount + 1: cacheIndex = cachedCount
            cachedNames(cacheIndex) = styleName: Set cachedStyles(cacheIndex) = style
        Else
            Set style = cachedStyles(cacheIndex)
        End If
        If style.Type <> wdStyleTypeParagraph Then Err.Raise 5, "WU_ApplyParagraphStyleRuns", "style run " & CStr(row) & " style is not a paragraph style"
        Set styleCache(row) = style: styleNames(row) = cachedNames(cacheIndex)
        previousEnd = endPosition
    Next row
    WU_ValidateParagraphStyleRuns = lastRow - firstRow + 1
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
End Function

Private Function WU_ReadStylePosition(ByVal value As Variant, ByRef position As Long) As Boolean
    Dim numericValue As Double, conversionError As Long
    If IsEmpty(value) Or IsObject(value) Or IsArray(value) Then Exit Function
    If VarType(value) = vbBoolean Or VarType(value) = vbDate Then Exit Function
    If Not IsNumeric(value) Then Exit Function
    On Error Resume Next
    numericValue = CDbl(value)
    conversionError = Err.Number
    Err.Clear
    On Error GoTo 0
    If conversionError <> 0 Or numericValue <> Fix(numericValue) Then Exit Function
    If numericValue < -2147483647# - 1# Or numericValue > 2147483647# Then Exit Function
    position = CLng(numericValue)
    WU_ReadStylePosition = True
End Function

Private Function WU_ParagraphStyleMatches(ByVal target As Range, ByVal style As Style, Optional ByVal expectedName As String = "") As Boolean
    Dim currentStyle As String, readError As Long
    On Error Resume Next
    currentStyle = CStr(target.Style)
    readError = Err.Number
    Err.Clear
    On Error GoTo 0
    If Len(expectedName) = 0 Then expectedName = style.NameLocal
    If readError = 0 Then WU_ParagraphStyleMatches = (StrComp(currentStyle, expectedName, vbTextCompare) = 0)
End Function

Private Function WU_StyleStory(ByVal document As Document, ByVal storyType As Long) As Range
    Dim readError As Long, readDescription As String
    On Error Resume Next
    Err.Clear
    Set WU_StyleStory = document.StoryRanges(storyType)
    readError = Err.Number
    readDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    ' 5941 is the normal missing-member result for an optional Word story.
    ' Do not hide any other retrieval failure from the caller.
    If readError <> 0 And readError <> WU_WORD_STORY_MISSING Then
        If Len(readDescription) = 0 Then readDescription = "Word could not retrieve the requested story."
        Err.Raise readError, "WU_StyleStory", "story " & CStr(storyType) & " is unavailable: " & readDescription
    End If
End Function

Private Function WU_StyleStoryHasContent(ByVal story As Range) As Boolean
    ' Empty stories are harmless; an unavailable boundary is not. Raise it
    ' through the caller so a style conversion cannot report success after
    ' silently omitting a linked story.
    Dim storyStart As Long, storyEnd As Long, readError As Long
    If story Is Nothing Then Exit Function
    On Error Resume Next
    storyStart = story.Start
    storyEnd = story.End
    readError = Err.Number
    Err.Clear
    On Error GoTo 0
    If readError <> 0 Then Err.Raise readError, "WU_StyleStoryHasContent", "story boundary is unavailable"
    WU_StyleStoryHasContent = (storyEnd > storyStart)
End Function

Private Function WU_StyleNextStory(ByVal story As Range) As Range
    Dim nextStory As Range, readError As Long
    If story Is Nothing Then Exit Function
    On Error Resume Next
    Set nextStory = story.NextStoryRange
    readError = Err.Number
    Err.Clear
    On Error GoTo 0
    If readError <> 0 Then Err.Raise readError, "WU_StyleNextStory", "linked story traversal is unavailable"
    If nextStory Is Nothing Then Exit Function
    If nextStory Is story Then Err.Raise 5, "WU_StyleNextStory", "self-referential story chain"
    Set WU_StyleNextStory = nextStory
End Function

Private Sub WU_PinFindOptions(ByVal criteria As Find)
    ' These options exist in current Word object libraries but are optional
    ' for some language packs/older hosts. Ignore only an unavailable option;
    ' the required style and replacement settings remain fail-fast.
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
