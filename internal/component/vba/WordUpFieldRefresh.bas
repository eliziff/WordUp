Attribute VB_Name = "WordUpFieldRefresh"
Option Explicit
Private Const WU_MAX_FIELD_STORY_CHAIN As Long = 32768
Private Const WU_WORD_STORY_MISSING As Long = 5941

' Refresh fields in explicitly selected Word stories. The return value is the
' number of stories/tables that reported an update failure; Word's Fields.Update
' return value is checked because it can report a failing field without raising.
Public Function WU_RefreshFields(ByVal document As Document, Optional ByVal storyScope As String = "all", Optional ByVal updateContents As Boolean = True) As Long
    Dim firstStory As Range, story As Range, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String, failures As Long
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_RefreshFields", "document is required"
    storyScope = WU_NormalizeFieldScope(storyScope, "WU_RefreshFields")
    ' Probe the selected stories before changing any application state. Most
    ' documents have no fields; returning here avoids an empty undo record,
    ' screen flicker, and a needless state round-trip on that hot no-op path.
    ' The probe is bounded and read-only, and raises unexpected story failures
    ' instead of treating a broken scope as field-free.
    If Not WU_FieldScopeHasWork(document, storyScope, updateContents) Then Exit Function
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Refresh fields": opened = True
    If storyScope = "main" Then
        Set story = WU_FieldStory(document, wdMainTextStory)
        If Not story Is Nothing Then WU_RefreshFieldStory story, failures
    ElseIf storyScope = "notes" Then
        Set firstStory = WU_FieldStory(document, wdFootnotesStory)
        If Not firstStory Is Nothing Then WU_RefreshFieldStoryChain firstStory, failures
        Set firstStory = Nothing
        Set firstStory = WU_FieldStory(document, wdEndnotesStory)
        If Not firstStory Is Nothing Then WU_RefreshFieldStoryChain firstStory, failures
    ElseIf storyScope = "headers" Then
        WU_RefreshFieldStoryType document, wdPrimaryHeaderStory, failures
        WU_RefreshFieldStoryType document, wdFirstPageHeaderStory, failures
        WU_RefreshFieldStoryType document, wdEvenPagesHeaderStory, failures
    ElseIf storyScope = "footers" Then
        WU_RefreshFieldStoryType document, wdPrimaryFooterStory, failures
        WU_RefreshFieldStoryType document, wdFirstPageFooterStory, failures
        WU_RefreshFieldStoryType document, wdEvenPagesFooterStory, failures
    Else
        For Each firstStory In document.StoryRanges
            WU_RefreshFieldStoryChain firstStory, failures
        Next firstStory
    End If
    If updateContents And (storyScope = "main" Or storyScope = "all") Then WU_RefreshFieldContents document, failures
    WU_RefreshFields = failures
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

' Refresh only fields inside the exact caller-supplied Range. The range is
' never widened and no table of contents is implicitly updated.
Public Function WU_RefreshFieldsInRange(ByVal target As Range) As Long
    Dim updating As Boolean, opened As Boolean, captured As Boolean, fieldResult As Long, updateError As Long, fieldCount As Long, readError As Long
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_RefreshFieldsInRange", "target range is required"
    If target.End <= target.Start Then Exit Function
    On Error Resume Next
    fieldCount = target.Fields.Count
    readError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If readError <> 0 Then Err.Raise readError, "WU_RefreshFieldsInRange", "fields are unavailable"
    If fieldCount = 0 Then Exit Function
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Refresh fields": opened = True
    On Error Resume Next
    Err.Clear
    fieldResult = target.Fields.Update
    updateError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If updateError <> 0 Then
        WU_RefreshFieldsInRange = 1
    ElseIf fieldResult <> 0 Then
        WU_RefreshFieldsInRange = 1
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

Private Function WU_NormalizeFieldScope(ByVal value As String, ByVal sourceName As String) As String
    value = LCase$(Trim$(value))
    If value <> "main" And value <> "notes" And value <> "headers" And value <> "footers" And value <> "all" Then Err.Raise 5, sourceName, "story scope must be main, notes, headers, footers, or all"
    WU_NormalizeFieldScope = value
End Function

Private Function WU_FieldScopeHasWork(ByVal document As Document, ByVal storyScope As String, ByVal updateContents As Boolean) As Boolean
    Dim firstStory As Range, story As Range, chainLength As Long
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If storyScope = "main" Then
        Set story = WU_FieldStory(document, wdMainTextStory)
        If Not story Is Nothing Then WU_FieldScopeHasWork = WU_FieldStoryHasFields(story)
    ElseIf storyScope = "notes" Then
        Set firstStory = WU_FieldStory(document, wdFootnotesStory)
        If Not firstStory Is Nothing Then WU_FieldScopeHasWork = WU_FieldChainHasFields(firstStory)
        If Not WU_FieldScopeHasWork Then
            Set firstStory = Nothing
            Set firstStory = WU_FieldStory(document, wdEndnotesStory)
            If Not firstStory Is Nothing Then WU_FieldScopeHasWork = WU_FieldChainHasFields(firstStory)
        End If
    ElseIf storyScope = "headers" Then
        WU_FieldScopeHasWork = WU_FieldStoryTypeHasFields(document, wdPrimaryHeaderStory)
        If Not WU_FieldScopeHasWork Then WU_FieldScopeHasWork = WU_FieldStoryTypeHasFields(document, wdFirstPageHeaderStory)
        If Not WU_FieldScopeHasWork Then WU_FieldScopeHasWork = WU_FieldStoryTypeHasFields(document, wdEvenPagesHeaderStory)
    ElseIf storyScope = "footers" Then
        WU_FieldScopeHasWork = WU_FieldStoryTypeHasFields(document, wdPrimaryFooterStory)
        If Not WU_FieldScopeHasWork Then WU_FieldScopeHasWork = WU_FieldStoryTypeHasFields(document, wdFirstPageFooterStory)
        If Not WU_FieldScopeHasWork Then WU_FieldScopeHasWork = WU_FieldStoryTypeHasFields(document, wdEvenPagesFooterStory)
    Else
        For Each firstStory In document.StoryRanges
            Set story = firstStory
            chainLength = 0
            Do While Not story Is Nothing
                chainLength = chainLength + 1
                If chainLength > WU_MAX_FIELD_STORY_CHAIN Then Err.Raise 5, "WU_FieldScopeHasWork", "story chain exceeds 32768 linked stories"
                If WU_FieldStoryHasFields(story) Then WU_FieldScopeHasWork = True: Exit Function
                Set story = WU_FieldNextStory(story)
            Loop
        Next firstStory
    End If
    If updateContents And (storyScope = "main" Or storyScope = "all") Then
        If document.TablesOfContents.Count > 0 Then WU_FieldScopeHasWork = True
    End If
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
End Function

Private Function WU_FieldStoryTypeHasFields(ByVal document As Document, ByVal storyType As Long) As Boolean
    Dim firstStory As Range
    Set firstStory = WU_FieldStory(document, storyType)
    If Not firstStory Is Nothing Then WU_FieldStoryTypeHasFields = WU_FieldChainHasFields(firstStory)
End Function

Private Function WU_FieldChainHasFields(ByVal firstStory As Range) As Boolean
    Dim story As Range, chainLength As Long
    Set story = firstStory
    Do While Not story Is Nothing
        chainLength = chainLength + 1
        If chainLength > WU_MAX_FIELD_STORY_CHAIN Then Err.Raise 5, "WU_FieldChainHasFields", "story chain exceeds 32768 linked stories"
        If WU_FieldStoryHasFields(story) Then WU_FieldChainHasFields = True: Exit Function
        Set story = WU_FieldNextStory(story)
    Loop
End Function

Private Function WU_FieldStoryHasFields(ByVal story As Range) As Boolean
    Dim fieldCount As Long, readError As Long, readDescription As String
    If story Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear
    fieldCount = story.Fields.Count
    readError = Err.Number
    readDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If readError <> 0 Then
        If Len(readDescription) = 0 Then readDescription = "Word could not inspect fields in the requested story."
        Err.Raise readError, "WU_FieldStoryHasFields", readDescription
    End If
    WU_FieldStoryHasFields = (fieldCount > 0)
End Function

Private Function WU_FieldNextStory(ByVal story As Range) As Range
    Dim nextStory As Range, readError As Long, readDescription As String
    If story Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear
    Set nextStory = story.NextStoryRange
    readError = Err.Number
    readDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If readError <> 0 Then
        If Len(readDescription) = 0 Then readDescription = "Word could not traverse the linked story."
        Err.Raise readError, "WU_FieldNextStory", readDescription
    End If
    If nextStory Is Nothing Then Exit Function
    If nextStory Is story Then Err.Raise 5, "WU_FieldNextStory", "self-referential story chain"
    Set WU_FieldNextStory = nextStory
End Function

Private Function WU_FieldStory(ByVal document As Document, ByVal storyType As Long) As Range
    Dim readError As Long, readDescription As String
    On Error Resume Next
    Err.Clear
    Set WU_FieldStory = document.StoryRanges(storyType)
    readError = Err.Number
    readDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    ' 5941 is Word's normal response when an optional story (for example a
    ' first-page header or a footnote story) does not exist. Every other
    ' retrieval failure is real evidence that the requested scope is broken;
    ' do not silently turn it into a successful partial refresh.
    If readError <> 0 And readError <> WU_WORD_STORY_MISSING Then
        If Len(readDescription) = 0 Then readDescription = "Word could not retrieve the requested story."
        Err.Raise readError, "WU_FieldStory", "story " & CStr(storyType) & " is unavailable: " & readDescription
    End If
End Function

Private Sub WU_RefreshFieldStoryType(ByVal document As Document, ByVal storyType As Long, ByRef failures As Long)
    Dim firstStory As Range
    Set firstStory = WU_FieldStory(document, storyType)
    If Not firstStory Is Nothing Then WU_RefreshFieldStoryChain firstStory, failures
End Sub

Private Sub WU_RefreshFieldStoryChain(ByVal firstStory As Range, ByRef failures As Long)
    Dim story As Range, nextStory As Range, nextError As Long, chainLength As Long
    Set story = firstStory
    Do While Not story Is Nothing
        chainLength = chainLength + 1
        If chainLength > WU_MAX_FIELD_STORY_CHAIN Then failures = failures + 1: Exit Do
        WU_RefreshFieldStory story, failures
        On Error Resume Next
        Set nextStory = story.NextStoryRange
        nextError = Err.Number
        Err.Clear
        On Error GoTo 0
        If nextError <> 0 Then failures = failures + 1: Exit Do
        If nextStory Is story Then failures = failures + 1: Exit Do
        Set story = nextStory
    Loop
End Sub

Private Sub WU_RefreshFieldStory(ByVal story As Range, ByRef failures As Long)
    Dim fieldResult As Long, fieldCount As Long, readError As Long, storyStart As Long, storyEnd As Long
    If story Is Nothing Then Exit Sub
    On Error Resume Next
    Err.Clear
    storyStart = story.Start
    storyEnd = story.End
    readError = Err.Number
    Err.Clear
    On Error GoTo 0
    If readError <> 0 Then failures = failures + 1: Exit Sub
    If storyEnd <= storyStart Then Exit Sub
    On Error Resume Next
    fieldCount = story.Fields.Count
    readError = Err.Number
    Err.Clear
    On Error GoTo 0
    If readError <> 0 Then failures = failures + 1: Exit Sub
    If fieldCount = 0 Then Exit Sub
    On Error Resume Next
    Err.Clear
    fieldResult = story.Fields.Update
    readError = Err.Number
    Err.Clear
    On Error GoTo 0
    If readError <> 0 Or fieldResult <> 0 Then failures = failures + 1
End Sub

Private Sub WU_RefreshFieldContents(ByVal document As Document, ByRef failures As Long)
    Dim contents As TableOfContents
    For Each contents In document.TablesOfContents
        On Error Resume Next
        Err.Clear
        contents.Update
        If Err.Number <> 0 Then failures = failures + 1
        Err.Clear
        On Error GoTo 0
    Next contents
End Sub
