Attribute VB_Name = "WordUpFieldRefresh"
Option Explicit
' Field refresh over explicitly selected Word stories or an exact Range,
' never through Selection. Requires WordUpStories and WordUpSafeEdit.
'
'   WU_RefreshFields(document, scope, updateContents)   refresh fields in a
'       scope; when updateContents is True and the scope includes the main
'       story, tables of contents are updated too
'   WU_RefreshFieldsInRange(target)                      refresh only the
'       fields inside the exact Range
'
' The result is the number of stories, ranges or tables of contents whose
' update failed. Word's Fields.Update return value is checked because it can
' report a failing field without raising. Scopes without fields return before
' any application state changes. Unavailable stories raise instead of being
' skipped silently.

Public Function WU_RefreshFields(ByVal document As Document, Optional ByVal scope As String = "all", Optional ByVal updateContents As Boolean = True) As Long
    Dim stories As Collection, story As Range, contents As TableOfContents, failures As Long, work As Boolean
    Dim updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_RefreshFields", "document is required"
    scope = WU_NormalizeScope(scope, "WU_RefreshFields")
    Set stories = WU_Stories(document, scope)
    For Each story In stories
        If WU_FieldCount(story, "WU_RefreshFields") > 0 Then work = True: Exit For
    Next story
    If updateContents And (scope = "main" Or scope = "all") Then
        If document.TablesOfContents.Count > 0 Then work = True
    End If
    If Not work Then Exit Function
    WU_BeginSafeEdit updating, opened, captured, "Refresh fields"
    For Each story In stories
        If WU_FieldCount(story, "WU_RefreshFields") > 0 Then
            If Not WU_UpdateFields(story) Then failures = failures + 1
        End If
    Next story
    If updateContents And (scope = "main" Or scope = "all") Then
        For Each contents In document.TablesOfContents
            On Error Resume Next
            Err.Clear
            contents.Update
            If Err.Number <> 0 Then failures = failures + 1
            Err.Clear
            On Error GoTo Failed
        Next contents
    End If
    WU_RefreshFields = failures
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

Public Function WU_RefreshFieldsInRange(ByVal target As Range) As Long
    Dim updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_RefreshFieldsInRange", "target range is required"
    If target.End <= target.Start Then Exit Function
    If WU_FieldCount(target, "WU_RefreshFieldsInRange") = 0 Then Exit Function
    WU_BeginSafeEdit updating, opened, captured, "Refresh fields"
    If Not WU_UpdateFields(target) Then WU_RefreshFieldsInRange = 1
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

Private Function WU_FieldCount(ByVal scope As Range, ByVal sourceName As String) As Long
    Dim readError As Long, readDescription As String
    On Error Resume Next
    Err.Clear
    WU_FieldCount = scope.Fields.Count
    readError = Err.Number
    readDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If readError <> 0 Then
        If Len(readDescription) = 0 Then readDescription = "Word could not inspect fields in the requested story."
        Err.Raise readError, sourceName, readDescription
    End If
End Function

' True when every field in the range updated cleanly.
Private Function WU_UpdateFields(ByVal scope As Range) As Boolean
    Dim result As Long, updateError As Long
    On Error Resume Next
    Err.Clear
    result = scope.Fields.Update
    updateError = Err.Number
    Err.Clear
    On Error GoTo 0
    WU_UpdateFields = (updateError = 0 And result = 0)
End Function
