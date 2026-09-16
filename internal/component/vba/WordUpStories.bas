Attribute VB_Name = "WordUpStories"
Option Explicit
' Shared story enumeration for the WordUp components.
'
' A scope names which Word stories an operation may touch:
'   main     the main text story
'   notes    footnotes and endnotes, including their linked stories
'   headers  primary, first-page and even-page headers of every section
'   footers  primary, first-page and even-page footers of every section
'   all      every story Word exposes for the document
'
' WU_Stories returns the stories of a scope as a Collection of Range objects.
' Word reports a story that does not exist with error 5941; that is treated
' as absent. Every other retrieval failure is raised so a caller cannot report
' success after silently skipping a story, and linked story chains are bounded
' so a malformed package cannot loop forever.
Private Const WU_WORD_STORY_MISSING As Long = 5941
Public Const WU_MAX_STORY_CHAIN As Long = 32768

Public Function WU_NormalizeScope(ByVal scope As String, Optional ByVal sourceName As String = "WU_NormalizeScope") As String
    scope = LCase$(Trim$(scope))
    If scope <> "main" And scope <> "notes" And scope <> "headers" And scope <> "footers" And scope <> "all" Then Err.Raise 5, sourceName, "story scope must be main, notes, headers, footers, or all"
    WU_NormalizeScope = scope
End Function

Public Function WU_Stories(ByVal document As Document, Optional ByVal scope As String = "main") As Collection
    Dim stories As Collection, story As Range
    If document Is Nothing Then Err.Raise 91, "WU_Stories", "document is required"
    scope = WU_NormalizeScope(scope, "WU_Stories")
    Set stories = New Collection
    Select Case scope
        Case "main"
            WU_AddStoryChain document, wdMainTextStory, stories
        Case "notes"
            WU_AddStoryChain document, wdFootnotesStory, stories
            WU_AddStoryChain document, wdEndnotesStory, stories
        Case "headers"
            WU_AddStoryChain document, wdPrimaryHeaderStory, stories
            WU_AddStoryChain document, wdFirstPageHeaderStory, stories
            WU_AddStoryChain document, wdEvenPagesHeaderStory, stories
        Case "footers"
            WU_AddStoryChain document, wdPrimaryFooterStory, stories
            WU_AddStoryChain document, wdFirstPageFooterStory, stories
            WU_AddStoryChain document, wdEvenPagesFooterStory, stories
        Case Else
            For Each story In document.StoryRanges
                WU_AddChain story, stories
            Next story
    End Select
    Set WU_Stories = stories
End Function

' The first story of a type, or Nothing when Word reports that it does not exist.
Public Function WU_Story(ByVal document As Document, ByVal storyType As Long) As Range
    Dim readError As Long, readDescription As String
    If document Is Nothing Then Err.Raise 91, "WU_Story", "document is required"
    On Error Resume Next
    Err.Clear
    Set WU_Story = document.StoryRanges(storyType)
    readError = Err.Number
    readDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If readError <> 0 And readError <> WU_WORD_STORY_MISSING Then
        If Len(readDescription) = 0 Then readDescription = "Word could not retrieve the requested story."
        Err.Raise readError, "WU_Story", "story " & CStr(storyType) & " is unavailable: " & readDescription
    End If
End Function

Public Function WU_StoryHasContent(ByVal story As Range) As Boolean
    Dim storyStart As Long, storyEnd As Long, readError As Long
    If story Is Nothing Then Exit Function
    On Error Resume Next
    storyStart = story.Start
    storyEnd = story.End
    readError = Err.Number
    Err.Clear
    On Error GoTo 0
    If readError <> 0 Then Err.Raise readError, "WU_StoryHasContent", "story boundary is unavailable"
    WU_StoryHasContent = (storyEnd > storyStart)
End Function

Public Function WU_NextStory(ByVal story As Range) As Range
    Dim nextStory As Range, readError As Long
    If story Is Nothing Then Exit Function
    On Error Resume Next
    Set nextStory = story.NextStoryRange
    readError = Err.Number
    Err.Clear
    On Error GoTo 0
    If readError <> 0 Then Err.Raise readError, "WU_NextStory", "linked story traversal is unavailable"
    If nextStory Is Nothing Then Exit Function
    If nextStory Is story Then Err.Raise 5, "WU_NextStory", "self-referential story chain"
    Set WU_NextStory = nextStory
End Function

Private Sub WU_AddStoryChain(ByVal document As Document, ByVal storyType As Long, ByVal stories As Collection)
    WU_AddChain WU_Story(document, storyType), stories
End Sub

Private Sub WU_AddChain(ByVal first As Range, ByVal stories As Collection)
    Dim story As Range, chainLength As Long
    Set story = first
    Do While Not story Is Nothing
        chainLength = chainLength + 1
        If chainLength > WU_MAX_STORY_CHAIN Then Err.Raise 5, "WU_Stories", "story chain exceeds " & CStr(WU_MAX_STORY_CHAIN) & " linked stories"
        If WU_StoryHasContent(story) Then stories.Add story
        Set story = WU_NextStory(story)
    Loop
End Sub
