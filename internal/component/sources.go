package component

const safeEditSource = `Attribute VB_Name = "WordUpSafeEdit"
Option Explicit
' Call Begin and End from the editing procedure itself. Word Application.Run
' cannot reliably propagate an unhandled macro error across its COM boundary.
' The caller's handler must save Err before calling End, then re-raise it.
' captured is explicit because a failed setup must never restore an
' uninitialized Boolean over the caller's ScreenUpdating state.
Public Sub WU_BeginSafeEdit(ByRef updating As Boolean, ByRef opened As Boolean, ByRef captured As Boolean, Optional ByVal label As String = "WordUp edit")
    Dim failure As Long, failureSource As String, failureText As String
    opened = False
    captured = False
    On Error GoTo Failed
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord label: opened = True
    Exit Sub
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error Resume Next
    If captured Then Application.ScreenUpdating = updating
    captured = False
    opened = False
    On Error GoTo 0
    Err.Raise failure, failureSource, failureText
End Sub
Public Sub WU_EndSafeEdit(ByVal updating As Boolean, ByRef opened As Boolean, ByRef captured As Boolean)
    Dim failure As Long, failureSource As String, failureText As String
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
        If failure = 0 Then opened = False
    End If
    Err.Clear
    If captured Then Application.ScreenUpdating = updating
    captured = False
    If failure = 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
End Sub
`

const formShellSource = `Attribute VB_Name = "WordUpFormShell"
Option Explicit
Public Sub WU_ShowForm(ByVal formName As String)
    Dim instance As Object, failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If Len(Trim$(formName)) = 0 Then Err.Raise 5, "WU_ShowForm", "form name is required"
    Set instance = VBA.UserForms.Add(formName)
    instance.Show
CleanUp:
    On Error Resume Next
    If Not instance Is Nothing Then Unload instance
    If failure = 0 And Err.Number <> 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Err.Clear
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Sub
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Sub
`

const progressSource = `Attribute VB_Name = "WordUpProgress"
Option Explicit
Private WU_Cancelled As Boolean
Private WU_LastYield As Single
Public Sub WU_RequestCancel(): WU_Cancelled = True: End Sub
Public Sub WU_ResetProgress(): WU_Cancelled = False: WU_LastYield = 0: End Sub
Public Function WU_CancelRequested() As Boolean
    Dim current As Single
    If WU_Cancelled Then WU_CancelRequested = True: Exit Function
    current = Timer
    If WU_LastYield = 0 Or current < WU_LastYield Or current - WU_LastYield >= 0.05 Then
        DoEvents
        WU_LastYield = Timer
    End If
    WU_CancelRequested = WU_Cancelled
End Function
`

const ribbonSource = `Attribute VB_Name = "WordUpRibbon"
Option Explicit
Public Sub WU_RibbonCommand(ByVal control As Object)
    Dim controlID As String, macroName As String
    If control Is Nothing Then Err.Raise 91, "WU_RibbonCommand", "control is required"
    controlID = CStr(control.Id)
    If Len(controlID) = 0 Then Err.Raise 5, "WU_RibbonCommand", "control id is required"
    macroName = WU_RibbonMacroName(controlID)
    Application.Run WU_QualifiedMacro(macroName)
End Sub
Private Function WU_RibbonMacroName(ByVal controlID As String) As String
    Dim i As Long, character As String, value As String, code As Long
    value = "WU_Command_"
    For i = 1 To Len(controlID)
        character = Mid$(controlID, i, 1)
        code = AscW(character)
        If code < 0 Then code = code + 65536
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Then
            value = value & character
        Else
            value = value & "_x" & Right$("0000" & Hex$(code), 4) & "_"
        End If
    Next i
    If Len(value) > 240 Then Err.Raise 5, "WU_RibbonMacroName", "control id is too long for a VBA callback name"
    WU_RibbonMacroName = value
End Function
Private Function WU_QualifiedMacro(ByVal macroName As String) As String
    WU_QualifiedMacro = "'" & Replace(ThisDocument.Name, "'", "''") & "'!" & macroName
End Function
`

const hotkeySource = `Attribute VB_Name = "WordUpHotkey"
Option Explicit
Public Sub WU_RegisterHotkey(ByVal keyCode As Long, ByVal macroName As String)
    If Len(Trim$(macroName)) = 0 Then Err.Raise 5, "WU_RegisterHotkey", "macro name is required"
    WU_ChangeHotkey keyCode, macroName, False
End Sub
Public Sub WU_RemoveHotkey(ByVal keyCode As Long)
    WU_ChangeHotkey keyCode, "", True
End Sub
Public Function WU_HotkeyRegistered(ByVal keyCode As Long, Optional ByVal expectedMacro As String = "") As Boolean
    Dim prior As Object, binding As KeyBinding
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    Set prior = Application.CustomizationContext
    Application.CustomizationContext = ThisDocument
    Set binding = WU_OwnedHotkey(keyCode)
    If binding Is Nothing Then GoTo CleanUp
    If expectedMacro <> "" Then If StrComp(WU_MacroMember(binding.Command), WU_MacroMember(expectedMacro), vbTextCompare) <> 0 Then GoTo CleanUp
    WU_HotkeyRegistered = True
CleanUp:
    On Error Resume Next
    Err.Clear
    Application.CustomizationContext = prior
    If failure = 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function
Private Sub WU_ChangeHotkey(ByVal keyCode As Long, ByVal macroName As String, ByVal remove As Boolean)
    Dim prior As Object, binding As KeyBinding, owner As Object
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    Set prior = Application.CustomizationContext
    Application.CustomizationContext = ThisDocument
    If remove Then
        Set binding = WU_OwnedHotkey(keyCode)
        If Not binding Is Nothing Then binding.Clear
    Else
        ' Replace only a binding owned by this template. Re-registering a
        ' shortcut must not accumulate duplicate entries or disturb Word's
        ' Normal template or another add-in.
        Set binding = FindKey(keyCode)
        If Not binding Is Nothing Then
            Set owner = binding.Context
            If owner Is Nothing Then Err.Raise 5, "WU_RegisterHotkey", "refusing to replace an unowned key binding"
            If Not owner Is ThisDocument Then Err.Raise 5, "WU_RegisterHotkey", "key binding belongs to another template"
            binding.Clear
        End If
        KeyBindings.Add wdKeyCategoryMacro, WU_QualifiedMacro(macroName), keyCode
    End If
CleanUp:
    On Error Resume Next
    Err.Clear
    Application.CustomizationContext = prior
    If failure = 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Sub
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Sub
Private Function WU_OwnedHotkey(ByVal keyCode As Long) As KeyBinding
    Dim binding As KeyBinding, owner As Object
    On Error Resume Next
    Set binding = FindKey(keyCode)
    Set owner = binding.Context
    If owner Is ThisDocument Then Set WU_OwnedHotkey = binding
    On Error GoTo 0
End Function
Private Function WU_QualifiedMacro(ByVal macroName As String) As String
    If InStr(1, macroName, "!", vbBinaryCompare) > 0 Then WU_QualifiedMacro = macroName Else WU_QualifiedMacro = "'" & Replace(ThisDocument.Name, "'", "''") & "'!" & macroName
End Function
Private Function WU_MacroMember(ByVal macroName As String) As String
    Dim separator As Long
    separator = InStrRev(macroName, "!", -1, vbBinaryCompare)
    If separator = 0 Then WU_MacroMember = macroName Else WU_MacroMember = Mid$(macroName, separator + 1)
End Function
`

const contextMenuSource = `Attribute VB_Name = "WordUpContextMenu"
Option Explicit
Public Sub WU_RegisterContextMenu(ByVal caption As String, ByVal macroName As String)
    If Len(Trim$(caption)) = 0 Then Err.Raise 5, "WU_RegisterContextMenu", "caption is required"
    If Len(Trim$(macroName)) = 0 Then Err.Raise 5, "WU_RegisterContextMenu", "macro name is required"
    WU_ChangeContextMenu caption, macroName, False
End Sub
Public Sub WU_RemoveContextMenu()
    WU_ChangeContextMenu "", "", True
End Sub
Public Function WU_ContextMenuRegistered(Optional ByVal expectedCaption As String = "", Optional ByVal expectedMacro As String = "") As Boolean
    Dim prior As Object, item As CommandBarControl
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    Set prior = Application.CustomizationContext
    Application.CustomizationContext = ThisDocument
    Set item = WU_FindContextMenu()
    If item Is Nothing Then GoTo CleanUp
    If expectedCaption <> "" Then If item.Caption <> expectedCaption Then GoTo CleanUp
    If expectedMacro <> "" Then If StrComp(WU_MacroMember(item.OnAction), WU_MacroMember(expectedMacro), vbTextCompare) <> 0 Then GoTo CleanUp
    WU_ContextMenuRegistered = True
CleanUp:
    On Error Resume Next
    Err.Clear
    Application.CustomizationContext = prior
    If failure = 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function
Private Sub WU_ChangeContextMenu(ByVal caption As String, ByVal macroName As String, ByVal removeOnly As Boolean)
    Dim prior As Object, item As CommandBarControl
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    Set prior = Application.CustomizationContext
    Application.CustomizationContext = ThisDocument
    WU_DeleteContextMenus
    If Not removeOnly Then
        Set item = CommandBars("Text").Controls.Add(Type:=msoControlButton, Temporary:=True)
        item.Caption = caption: item.OnAction = WU_QualifiedMacro(macroName): item.Tag = WU_ContextMenuTag()
    End If
CleanUp:
    On Error Resume Next
    Err.Clear
    Application.CustomizationContext = prior
    If failure = 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Sub
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Sub
Private Function WU_FindContextMenu() As CommandBarControl
    Dim control As CommandBarControl
    On Error Resume Next
    For Each control In CommandBars("Text").Controls
        If control.Tag = WU_ContextMenuTag() And control.Type = msoControlButton Then Set WU_FindContextMenu = control: Exit Function
    Next control
    On Error GoTo 0
End Function
Private Sub WU_DeleteContextMenus()
    Dim controls As CommandBarControls, i As Long
    Set controls = CommandBars("Text").Controls
    For i = controls.Count To 1 Step -1
        If controls(i).Tag = WU_ContextMenuTag() Then controls(i).Delete
    Next i
End Sub
Private Function WU_ContextMenuTag() As String
    ' The full path keeps two templates with the same file name independent.
    ' FullName is still stable for an unsaved document (Word returns its name).
    WU_ContextMenuTag = "WordUp.ContextMenu." & ThisDocument.FullName
End Function
Private Function WU_QualifiedMacro(ByVal macroName As String) As String
    If InStr(1, macroName, "!", vbBinaryCompare) > 0 Then WU_QualifiedMacro = macroName Else WU_QualifiedMacro = "'" & Replace(ThisDocument.Name, "'", "''") & "'!" & macroName
End Function
Private Function WU_MacroMember(ByVal macroName As String) As String
    Dim separator As Long
    separator = InStrRev(macroName, "!", -1, vbBinaryCompare)
    If separator = 0 Then WU_MacroMember = macroName Else WU_MacroMember = Mid$(macroName, separator + 1)
End Function
`

const styleConverterSource = `Attribute VB_Name = "WordUpStyleConverter"
Option Explicit
Public Function WU_ConvertStyle(ByVal document As Document, ByVal fromStyle As String, ByVal toStyle As String) As Boolean
    Dim firstStory As Range, story As Range, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim sourceStyle As Style, targetStyle As Style
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ConvertStyle", "document is required"
    If Len(Trim$(fromStyle)) = 0 Then Err.Raise 5, "WU_ConvertStyle", "source style is required"
    If Len(Trim$(toStyle)) = 0 Then Err.Raise 5, "WU_ConvertStyle", "target style is required"
    updating = Application.ScreenUpdating
    captured = True
    Set sourceStyle = document.Styles(fromStyle)
    Set targetStyle = document.Styles(toStyle)
    If sourceStyle.Type <> wdStyleTypeParagraph Then Err.Raise 5, "WU_ConvertStyle", "source style is not a paragraph style"
    If targetStyle.Type <> wdStyleTypeParagraph Then Err.Raise 5, "WU_ConvertStyle", "target style is not a paragraph style"
    If StrComp(sourceStyle.NameLocal, targetStyle.NameLocal, vbTextCompare) = 0 Then Exit Function
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Convert style": opened = True
    For Each firstStory In document.StoryRanges
        Set story = firstStory
        Do
            If WU_ConvertStyleInRange(story, sourceStyle, targetStyle) Then WU_ConvertStyle = True
            Set story = story.NextStoryRange
        Loop Until story Is Nothing
    Next firstStory
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
Private Function WU_ConvertStyleInRange(ByVal story As Range, ByVal sourceStyle As Style, ByVal targetStyle As Style) As Boolean
    Dim scope As Range
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
    End With
WU_ConvertStyleInRange = scope.Find.Execute(Replace:=wdReplaceAll)
End Function
`

const textOperationsSource = `Attribute VB_Name = "WordUpTextOperations"
Option Explicit

' Replace visible literal text with one bounded Word Find pass per story.
' The operation never uses Selection and keeps Word's existing formatting on
' the found range. Use raw XML when the intended edit is a field instruction,
' relationship, or other package markup rather than visible text.
Public Function WU_ReplaceLiteral(ByVal document As Document, ByVal findText As String, ByVal replaceText As String, Optional ByVal storyScope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False) As Boolean
    Dim firstStory As Range, story As Range, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String, changed As Boolean
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ReplaceLiteral", "document is required"
    WU_ValidateLiteral findText, replaceText, "WU_ReplaceLiteral"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "all" Then Err.Raise 5, "WU_ReplaceLiteral", "story scope must be main, notes, or all"
    ' With MatchCase on, identical find/replacement text is an exact no-op.
    ' When MatchCase is off, the same spelling could intentionally normalize
    ' the case of a differently-cased match, so it still runs.
    If matchCase And StrComp(findText, replaceText, vbBinaryCompare) = 0 Then Exit Function
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Replace literal text": opened = True
    If storyScope = "main" Then
        ' The main story is a single range. Avoid enumerating every empty
        ' header, footer, text frame and note when the caller requested the
        ' overwhelmingly common body-only operation.
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then
            If story.End > story.Start Then changed = WU_ReplaceLiteralInStory(story, findText, replaceText, matchCase, wholeWord)
        End If
    ElseIf storyScope = "notes" Then
        ' Notes are the only non-main stories most journal operations need.
        ' Address their two roots directly instead of enumerating unrelated
        ' headers, footers, text boxes, comments, and text frames.
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then changed = WU_ReplaceLiteralInStoryChain(firstStory, findText, replaceText, matchCase, wholeWord)
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then If WU_ReplaceLiteralInStoryChain(firstStory, findText, replaceText, matchCase, wholeWord) Then changed = True
    Else
        For Each firstStory In document.StoryRanges
            If WU_ReplaceLiteralInStoryChain(firstStory, findText, replaceText, matchCase, wholeWord) Then changed = True
        Next firstStory
    End If
    WU_ReplaceLiteral = changed
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

' Replace only inside an already-bounded Range. This is the fast path for
' callers that have an exact paragraph, content control, table cell, or other
' Word range and must not touch any other story. The range's direct formatting
' is retained by Word's formatting-neutral replacement.
Public Function WU_ReplaceLiteralInRange(ByVal target As Range, ByVal findText As String, ByVal replaceText As String, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False) As Boolean
    Dim updating As Boolean, opened As Boolean, captured As Boolean, targetStart As Long, targetEnd As Long
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ReplaceLiteralInRange", "target range is required"
    WU_ValidateLiteral findText, replaceText, "WU_ReplaceLiteralInRange"
    targetStart = target.Start: targetEnd = target.End
    If targetEnd <= targetStart Then Exit Function
    If matchCase And StrComp(findText, replaceText, vbBinaryCompare) = 0 Then Exit Function
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Replace literal text": opened = True
    WU_ReplaceLiteralInRange = WU_ReplaceLiteralInStory(target, findText, replaceText, matchCase, wholeWord)
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

Private Sub WU_ValidateLiteral(ByVal findText As String, ByVal replaceText As String, ByVal sourceName As String)
    If Len(findText) = 0 Then Err.Raise 5, sourceName, "find text is required"
    If Len(findText) > 255 Then Err.Raise 5, sourceName, "find text exceeds Word's 255-character limit"
    If Len(replaceText) > 255 Then Err.Raise 5, sourceName, "replacement text exceeds Word's 255-character limit"
    If Len(WU_EscapeFindLiteral(findText)) > 255 Then Err.Raise 5, sourceName, "find text exceeds Word's escaped 255-character limit"
    If Len(WU_EscapeFindLiteral(replaceText)) > 255 Then Err.Raise 5, sourceName, "replacement text exceeds Word's escaped 255-character limit"
End Sub

Private Function WU_ReplaceLiteralInStoryChain(ByVal firstStory As Range, ByVal findText As String, ByVal replaceText As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean) As Boolean
    Dim story As Range, changed As Boolean
    Set story = firstStory
    Do While Not story Is Nothing
        If story.End > story.Start Then
            If WU_ReplaceLiteralInStory(story, findText, replaceText, matchCase, wholeWord) Then changed = True
        End If
        Set story = story.NextStoryRange
    Loop
    WU_ReplaceLiteralInStoryChain = changed
End Function

Private Function WU_ReplaceLiteralInStory(ByVal story As Range, ByVal findText As String, ByVal replaceText As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean) As Boolean
    Dim search As Range
    Set search = story.Duplicate
    With search.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        ' Find/Replace treats ^p, ^t, and similar sequences as structural
        ' tokens. Doubling the marker keeps this helper literal by default.
        .Text = WU_EscapeFindLiteral(findText)
        .Replacement.Text = WU_EscapeFindLiteral(replaceText)
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = matchCase
        .MatchWholeWord = wholeWord
        .MatchWildcards = False
    End With
    WU_ReplaceLiteralInStory = search.Find.Execute(Replace:=wdReplaceAll)
End Function

Private Function WU_EscapeFindLiteral(ByVal value As String) As String
    WU_EscapeFindLiteral = Replace(value, "^", "^^")
End Function
`
