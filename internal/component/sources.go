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
Private Const WU_MAX_STYLE_BATCH_RULES As Long = 256
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
        If Not story Is Nothing Then If story.End > story.Start Then WU_ConvertStyle = WU_ConvertStyleInStory(story, sourceStyle, targetStyle)
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then WU_ConvertStyle = WU_ConvertStyleInStoryChain(firstStory, sourceStyle, targetStyle)
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo Failed
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
        If Not story Is Nothing Then If story.End > story.Start Then WU_ConvertStyleBatchInStory story, sourceCache, targetCache, enabled, firstRow, lastRow, matched
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then WU_ConvertStyleBatchInStoryChain firstStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo Failed
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
        If Not story Is Nothing Then If story.End > story.Start Then WU_ConvertCharacterStyle = WU_ConvertCharacterStyleInStory(story, sourceStyle, targetStyle)
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then WU_ConvertCharacterStyle = WU_ConvertCharacterStyleInStoryChain(firstStory, sourceStyle, targetStyle)
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo Failed
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
        If Not story Is Nothing Then If story.End > story.Start Then WU_ConvertCharacterStyleBatchInStory story, sourceCache, targetCache, enabled, firstRow, lastRow, matched
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then WU_ConvertCharacterStyleBatchInStoryChain firstStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo Failed
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
    WU_ConvertStyleBatchInStory target, sourceCache, targetCache, enabled, firstRow, lastRow, matched
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
    Dim story As Range, changed As Boolean
    Set story = firstStory
    Do While Not story Is Nothing
        If story.End > story.Start Then If WU_ConvertStyleInStory(story, sourceStyle, targetStyle) Then changed = True
        Set story = story.NextStoryRange
    Loop
    WU_ConvertStyleInStoryChain = changed
End Function

' Return False when a document has no story of the requested header/footer
' type. Word raises for absent first/even-page stories, so the optional story
' is probed once and the caller's error handler is restored before editing.
Private Function WU_ConvertStyleInStoryType(ByVal document As Document, ByVal storyType As Long, ByVal sourceStyle As Style, ByVal targetStyle As Style) As Boolean
    Dim firstStory As Range
    On Error Resume Next
    Set firstStory = document.StoryRanges(storyType)
    Err.Clear
    On Error GoTo 0
    If Not firstStory Is Nothing Then WU_ConvertStyleInStoryType = WU_ConvertStyleInStoryChain(firstStory, sourceStyle, targetStyle)
End Function

Private Sub WU_ConvertStyleBatchInStoryType(ByVal document As Document, ByVal storyType As Long, ByRef sourceCache() As Style, ByRef targetCache() As Style, ByRef enabled() As Boolean, ByVal firstRow As Long, ByVal lastRow As Long, ByRef matched() As Boolean)
    Dim firstStory As Range
    On Error Resume Next
    Set firstStory = document.StoryRanges(storyType)
    Err.Clear
    On Error GoTo 0
    If Not firstStory Is Nothing Then WU_ConvertStyleBatchInStoryChain firstStory, sourceCache, targetCache, enabled, firstRow, lastRow, matched
End Sub

Private Sub WU_ConvertStyleBatchInStoryChain(ByVal firstStory As Range, ByRef sourceCache() As Style, ByRef targetCache() As Style, ByRef enabled() As Boolean, ByVal firstRow As Long, ByVal lastRow As Long, ByRef matched() As Boolean)
    Dim story As Range
    Set story = firstStory
    Do While Not story Is Nothing
        If story.End > story.Start Then WU_ConvertStyleBatchInStory story, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        Set story = story.NextStoryRange
    Loop
End Sub

Private Sub WU_ConvertCharacterStyleBatchInStory(ByVal story As Range, ByRef sourceCache() As Style, ByRef targetCache() As Style, ByRef enabled() As Boolean, ByVal firstRow As Long, ByVal lastRow As Long, ByRef matched() As Boolean)
    Dim row As Long
    For row = firstRow To lastRow
        If enabled(row) Then If WU_ConvertCharacterStyleInStory(story, sourceCache(row), targetCache(row)) Then matched(row) = True
    Next row
End Sub

Private Sub WU_ConvertCharacterStyleBatchInStoryChain(ByVal firstStory As Range, ByRef sourceCache() As Style, ByRef targetCache() As Style, ByRef enabled() As Boolean, ByVal firstRow As Long, ByVal lastRow As Long, ByRef matched() As Boolean)
    Dim story As Range
    Set story = firstStory
    Do While Not story Is Nothing
        If story.End > story.Start Then WU_ConvertCharacterStyleBatchInStory story, sourceCache, targetCache, enabled, firstRow, lastRow, matched
        Set story = story.NextStoryRange
    Loop
End Sub

Private Sub WU_ConvertCharacterStyleBatchInStoryType(ByVal document As Document, ByVal storyType As Long, ByRef sourceCache() As Style, ByRef targetCache() As Style, ByRef enabled() As Boolean, ByVal firstRow As Long, ByVal lastRow As Long, ByRef matched() As Boolean)
    Dim firstStory As Range
    On Error Resume Next
    Set firstStory = document.StoryRanges(storyType)
    Err.Clear
    On Error GoTo 0
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
    Dim story As Range, changed As Boolean
    Set story = firstStory
    Do While Not story Is Nothing
        If story.End > story.Start Then If WU_ConvertCharacterStyleInStory(story, sourceStyle, targetStyle) Then changed = True
        Set story = story.NextStoryRange
    Loop
    WU_ConvertCharacterStyleInStoryChain = changed
End Function

Private Function WU_ConvertCharacterStyleInStoryType(ByVal document As Document, ByVal storyType As Long, ByVal sourceStyle As Style, ByVal targetStyle As Style) As Boolean
    Dim firstStory As Range
    On Error Resume Next
    Set firstStory = document.StoryRanges(storyType)
    Err.Clear
    On Error GoTo 0
    If Not firstStory Is Nothing Then WU_ConvertCharacterStyleInStoryType = WU_ConvertCharacterStyleInStoryChain(firstStory, sourceStyle, targetStyle)
End Function

Private Function WU_IsCharacterStyle(ByVal style As Style) As Boolean
    If style Is Nothing Then Exit Function
    WU_IsCharacterStyle = (style.Type = wdStyleTypeCharacter Or style.Linked)
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
`

const textOperationsSource = `Attribute VB_Name = "WordUpTextOperations"
Option Explicit
Private Const WU_MAX_BATCH_RULES As Long = 1024

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
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_ReplaceLiteral", "story scope must be main, notes, headers, footers, or all"
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
    ElseIf storyScope = "headers" Then
        If WU_ReplaceLiteralInStoryType(document, wdPrimaryHeaderStory, findText, replaceText, matchCase, wholeWord) Then changed = True
        If WU_ReplaceLiteralInStoryType(document, wdFirstPageHeaderStory, findText, replaceText, matchCase, wholeWord) Then changed = True
        If WU_ReplaceLiteralInStoryType(document, wdEvenPagesHeaderStory, findText, replaceText, matchCase, wholeWord) Then changed = True
    ElseIf storyScope = "footers" Then
        If WU_ReplaceLiteralInStoryType(document, wdPrimaryFooterStory, findText, replaceText, matchCase, wholeWord) Then changed = True
        If WU_ReplaceLiteralInStoryType(document, wdFirstPageFooterStory, findText, replaceText, matchCase, wholeWord) Then changed = True
        If WU_ReplaceLiteralInStoryType(document, wdEvenPagesFooterStory, findText, replaceText, matchCase, wholeWord) Then changed = True
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

' Count literal matches without changing document state. This is the
' read-only companion for citation and quality audits: it uses the same
' bounded story selection and explicit Find flags as the edit paths, but
' never opens an undo record or toggles ScreenUpdating.
Public Function WU_CountLiteral(ByVal document As Document, ByVal findText As String, Optional ByVal storyScope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False) As Long
    Dim firstStory As Range, story As Range, count As Long
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_CountLiteral", "document is required"
    WU_ValidateLiteral findText, "", "WU_CountLiteral"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_CountLiteral", "story scope must be main, notes, headers, footers, or all"
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If story.End > story.Start Then count = WU_CountLiteralInStory(story, findText, matchCase, wholeWord)
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then count = WU_CountLiteralInStoryChain(firstStory, findText, matchCase, wholeWord)
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then count = count + WU_CountLiteralInStoryChain(firstStory, findText, matchCase, wholeWord)
    ElseIf storyScope = "headers" Then
        count = count + WU_CountLiteralInStoryType(document, wdPrimaryHeaderStory, findText, matchCase, wholeWord)
        count = count + WU_CountLiteralInStoryType(document, wdFirstPageHeaderStory, findText, matchCase, wholeWord)
        count = count + WU_CountLiteralInStoryType(document, wdEvenPagesHeaderStory, findText, matchCase, wholeWord)
    ElseIf storyScope = "footers" Then
        count = count + WU_CountLiteralInStoryType(document, wdPrimaryFooterStory, findText, matchCase, wholeWord)
        count = count + WU_CountLiteralInStoryType(document, wdFirstPageFooterStory, findText, matchCase, wholeWord)
        count = count + WU_CountLiteralInStoryType(document, wdEvenPagesFooterStory, findText, matchCase, wholeWord)
    Else
        For Each firstStory In document.StoryRanges
            count = count + WU_CountLiteralInStoryChain(firstStory, findText, matchCase, wholeWord)
        Next firstStory
    End If
    WU_CountLiteral = count
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
End Function

' Count only inside the exact caller-supplied Range. The range is never
' widened to a story, which keeps audits safe for a selected section or cell.
Public Function WU_CountLiteralInRange(ByVal target As Range, ByVal findText As String, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False) As Long
    Dim targetStart As Long, targetEnd As Long
    If target Is Nothing Then Err.Raise 91, "WU_CountLiteralInRange", "target range is required"
    WU_ValidateLiteral findText, "", "WU_CountLiteralInRange"
    targetStart = target.Start: targetEnd = target.End
    If targetEnd <= targetStart Then Exit Function
    WU_CountLiteralInRange = WU_CountLiteralInStory(target, findText, matchCase, wholeWord)
End Function

' Replace a two-column Variant array of find/replacement pairs in one safe edit.
' The first column is the literal find text and the second is its replacement.
' The array is only a compact command list: Word remains the source of truth
' for document text and formatting. The return value is the number of pairs
' that matched at least once, not a guessed character count. Pairs run in row
' order, so an intentional replacement chain is explicit and deterministic.
Public Function WU_ReplaceLiteralBatch(ByVal document As Document, ByVal replacements As Variant, Optional ByVal storyScope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False) As Long
    Dim firstStory As Range, story As Range, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, activeRows As Long, changed As Long, row As Long
    Dim matched() As Boolean
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ReplaceLiteralBatch", "document is required"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_ReplaceLiteralBatch", "story scope must be main, notes, headers, footers, or all"
    activeRows = WU_ValidateLiteralBatch(replacements, matchCase, "WU_ReplaceLiteralBatch")
    If activeRows = 0 Then Exit Function
    firstRow = LBound(replacements, 1): lastRow = UBound(replacements, 1): firstColumn = LBound(replacements, 2)
    ReDim matched(firstRow To lastRow)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Replace literal text batch": opened = True
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If story.End > story.Start Then WU_ReplaceLiteralBatchInStory story, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then WU_ReplaceLiteralBatchInStoryChain firstStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then WU_ReplaceLiteralBatchInStoryChain firstStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched
    ElseIf storyScope = "headers" Then
        WU_ReplaceLiteralBatchInStoryType document, wdPrimaryHeaderStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched
        WU_ReplaceLiteralBatchInStoryType document, wdFirstPageHeaderStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched
        WU_ReplaceLiteralBatchInStoryType document, wdEvenPagesHeaderStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched
    ElseIf storyScope = "footers" Then
        WU_ReplaceLiteralBatchInStoryType document, wdPrimaryFooterStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched
        WU_ReplaceLiteralBatchInStoryType document, wdFirstPageFooterStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched
        WU_ReplaceLiteralBatchInStoryType document, wdEvenPagesFooterStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched
    Else
        For Each firstStory In document.StoryRanges
            WU_ReplaceLiteralBatchInStoryChain firstStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched
        Next firstStory
    End If
    For row = firstRow To lastRow
        If matched(row) Then changed = changed + 1
    Next row
    WU_ReplaceLiteralBatch = changed
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

' Replace a two-column Variant array only inside the exact caller-supplied
' Range. Validation happens before the one undo record is opened, and the
' return value counts pairs that matched at least once.
Public Function WU_ReplaceLiteralBatchInRange(ByVal target As Range, ByVal replacements As Variant, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False) As Long
    Dim document As Document, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim targetStart As Long, targetEnd As Long, firstRow As Long, lastRow As Long, firstColumn As Long, row As Long, changed As Long, activeRows As Long
    Dim matched() As Boolean
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ReplaceLiteralBatchInRange", "target range is required"
    Set document = target.Document
    activeRows = WU_ValidateLiteralBatch(replacements, matchCase, "WU_ReplaceLiteralBatchInRange")
    If activeRows = 0 Then Exit Function
    targetStart = target.Start: targetEnd = target.End
    If targetEnd <= targetStart Then Exit Function
    firstRow = LBound(replacements, 1): lastRow = UBound(replacements, 1): firstColumn = LBound(replacements, 2)
    ReDim matched(firstRow To lastRow)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Replace literal text batch": opened = True
    WU_ReplaceLiteralBatchInStory target, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched
    For row = firstRow To lastRow
        If matched(row) Then changed = changed + 1
    Next row
    WU_ReplaceLiteralBatchInRange = changed
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

' Apply a character or linked style to literal matches without replacing the
' matched text. This is the neutral primitive for citation, case-name, and
' short-form emphasis rules; callers decide which terms are eligible.
Public Function WU_ApplyCharacterStyleToMatches(ByVal document As Document, ByVal findText As String, ByVal styleName As String, Optional ByVal storyScope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = True) As Long
    Dim firstStory As Range, story As Range, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String, style As Style, changed As Long, styleError As Long
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ApplyCharacterStyleToMatches", "document is required"
    WU_ValidateLiteral findText, "", "WU_ApplyCharacterStyleToMatches"
    If Len(Trim$(styleName)) = 0 Then Err.Raise 5, "WU_ApplyCharacterStyleToMatches", "style name is required"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_ApplyCharacterStyleToMatches", "story scope must be main, notes, headers, footers, or all"
    On Error Resume Next
    Set style = document.Styles(styleName)
    styleError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If styleError <> 0 Or style Is Nothing Then Err.Raise 5, "WU_ApplyCharacterStyleToMatches", "style " & styleName & " was not found"
    If style.Type <> wdStyleTypeCharacter And Not style.Linked Then Err.Raise 5, "WU_ApplyCharacterStyleToMatches", "style is not a character style"
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Style literal matches": opened = True
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If story.End > story.Start Then changed = WU_ApplyCharacterStyleInStory(story, findText, style, matchCase, wholeWord)
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then changed = WU_ApplyCharacterStyleInStoryChain(firstStory, findText, style, matchCase, wholeWord)
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then changed = changed + WU_ApplyCharacterStyleInStoryChain(firstStory, findText, style, matchCase, wholeWord)
    ElseIf storyScope = "headers" Then
        changed = changed + WU_ApplyCharacterStyleInStoryType(document, wdPrimaryHeaderStory, findText, style, matchCase, wholeWord)
        changed = changed + WU_ApplyCharacterStyleInStoryType(document, wdFirstPageHeaderStory, findText, style, matchCase, wholeWord)
        changed = changed + WU_ApplyCharacterStyleInStoryType(document, wdEvenPagesHeaderStory, findText, style, matchCase, wholeWord)
    ElseIf storyScope = "footers" Then
        changed = changed + WU_ApplyCharacterStyleInStoryType(document, wdPrimaryFooterStory, findText, style, matchCase, wholeWord)
        changed = changed + WU_ApplyCharacterStyleInStoryType(document, wdFirstPageFooterStory, findText, style, matchCase, wholeWord)
        changed = changed + WU_ApplyCharacterStyleInStoryType(document, wdEvenPagesFooterStory, findText, style, matchCase, wholeWord)
    Else
        For Each firstStory In document.StoryRanges
            changed = changed + WU_ApplyCharacterStyleInStoryChain(firstStory, findText, style, matchCase, wholeWord)
        Next firstStory
    End If
    WU_ApplyCharacterStyleToMatches = changed
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

' Apply a character style inside the exact caller-supplied Range. The public
' range boundary is never widened to a story or Selection.
Public Function WU_ApplyCharacterStyleToRange(ByVal target As Range, ByVal findText As String, ByVal styleName As String, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = True) As Long
    Dim updating As Boolean, opened As Boolean, captured As Boolean, failure As Long, failureSource As String, failureText As String, styleError As Long
    Dim style As Style, document As Document, targetStart As Long, targetEnd As Long
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ApplyCharacterStyleToRange", "target range is required"
    WU_ValidateLiteral findText, "", "WU_ApplyCharacterStyleToRange"
    If Len(Trim$(styleName)) = 0 Then Err.Raise 5, "WU_ApplyCharacterStyleToRange", "style name is required"
    Set document = target.Document
    On Error Resume Next
    Set style = document.Styles(styleName)
    styleError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If styleError <> 0 Or style Is Nothing Then Err.Raise 5, "WU_ApplyCharacterStyleToRange", "style " & styleName & " was not found"
    If style.Type <> wdStyleTypeCharacter And Not style.Linked Then Err.Raise 5, "WU_ApplyCharacterStyleToRange", "style is not a character style"
    targetStart = target.Start: targetEnd = target.End
    If targetEnd <= targetStart Then Exit Function
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Style literal matches": opened = True
    WU_ApplyCharacterStyleToRange = WU_ApplyCharacterStyleInStory(target, findText, style, matchCase, wholeWord)
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

' Apply a character or linked style to the exact caller-supplied Range. This is
' the zero-search path for an inspector or detector that already owns the
' source span; text, fields, hyperlinks, and direct run formatting stay in
' Word's Range while one style assignment changes only that span.
Public Function WU_ApplyCharacterStyleInRange(ByVal target As Range, ByVal styleName As String) As Boolean
    Dim document As Document, style As Style, scope As Range, expectedName As String
    Dim updating As Boolean, opened As Boolean, captured As Boolean, styleError As Long
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ApplyCharacterStyleInRange", "target range is required"
    If Len(Trim$(styleName)) = 0 Then Err.Raise 5, "WU_ApplyCharacterStyleInRange", "style name is required"
    Set document = target.Document
    On Error Resume Next
    Set style = document.Styles(styleName)
    styleError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If styleError <> 0 Or style Is Nothing Then Err.Raise 5, "WU_ApplyCharacterStyleInRange", "style " & styleName & " was not found"
    If style.Type <> wdStyleTypeCharacter And Not style.Linked Then Err.Raise 5, "WU_ApplyCharacterStyleInRange", "style is not a character style"
    If target.End <= target.Start Then Exit Function
    expectedName = style.NameLocal
    If WU_CharacterStyleMatches(target, expectedName) Then Exit Function
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Apply character style": opened = True
    Set scope = target.Duplicate
    scope.Style = style
    WU_ApplyCharacterStyleInRange = True
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

' Apply a character or linked style to exact Word story offsets in one edit.
' Each row is [absoluteStart, absoluteEnd, styleName]. This is the fast,
' deterministic path when an inspector or detector already located citation,
' case-name, or short-form spans: it never re-searches text and never copies
' the matched text through an intermediate representation.
Public Function WU_ApplyCharacterStyleRuns(ByVal target As Range, ByVal runs As Variant) As Long
    Dim document As Document, scope As Range, style As Style
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, row As Long, activeRows As Long, changed As Long
    Dim startPosition As Long, endPosition As Long, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim styleCache() As Style, styleNames() As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ApplyCharacterStyleRuns", "target range is required"
    Set document = target.Document
    activeRows = WU_ValidateCharacterStyleRuns(document, target, runs, styleCache, styleNames)
    If activeRows = 0 Then Exit Function
    firstRow = LBound(runs, 1): lastRow = UBound(runs, 1): firstColumn = LBound(runs, 2)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Apply character style runs": opened = True
    Set scope = target.Duplicate
    For row = firstRow To lastRow
        startPosition = CLng(runs(row, firstColumn))
        endPosition = CLng(runs(row, firstColumn + 1))
        ' Extend first, then move the start; assigning a later start to a
        ' reused Range before its end can make Word reject the transient span.
        scope.End = endPosition: scope.Start = startPosition
        Set style = styleCache(row)
        If Not WU_CharacterStyleMatches(scope, styleNames(row)) Then
            scope.Style = style
            changed = changed + 1
        End If
    Next row
    WU_ApplyCharacterStyleRuns = changed
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

' Apply a two-column Variant array of literal/style pairs in one safe edit.
' The first column is the literal to find and the second is an existing
' character or linked style name. The return value is the number of matching
' ranges whose style actually changed; repeated rules are therefore safe and
' do not inflate the result after the first application.
Public Function WU_ApplyCharacterStyleBatch(ByVal document As Document, ByVal matches As Variant, Optional ByVal storyScope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = True) As Long
    Dim firstStory As Range, story As Range, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, activeRows As Long, changed As Long
    Dim styleCache() As Style
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ApplyCharacterStyleBatch", "document is required"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_ApplyCharacterStyleBatch", "story scope must be main, notes, headers, footers, or all"
    activeRows = WU_ValidateCharacterStyleBatch(document, matches, styleCache)
    If activeRows = 0 Then Exit Function
    firstRow = LBound(matches, 1): lastRow = UBound(matches, 1): firstColumn = LBound(matches, 2)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Style literal matches batch": opened = True
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If story.End > story.Start Then changed = WU_ApplyCharacterStyleBatchInStory(story, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord)
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then changed = WU_ApplyCharacterStyleBatchInStoryChain(firstStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord)
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then changed = changed + WU_ApplyCharacterStyleBatchInStoryChain(firstStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord)
    ElseIf storyScope = "headers" Then
        changed = changed + WU_ApplyCharacterStyleBatchInStoryType(document, wdPrimaryHeaderStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord)
        changed = changed + WU_ApplyCharacterStyleBatchInStoryType(document, wdFirstPageHeaderStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord)
        changed = changed + WU_ApplyCharacterStyleBatchInStoryType(document, wdEvenPagesHeaderStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord)
    ElseIf storyScope = "footers" Then
        changed = changed + WU_ApplyCharacterStyleBatchInStoryType(document, wdPrimaryFooterStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord)
        changed = changed + WU_ApplyCharacterStyleBatchInStoryType(document, wdFirstPageFooterStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord)
        changed = changed + WU_ApplyCharacterStyleBatchInStoryType(document, wdEvenPagesFooterStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord)
    Else
        For Each firstStory In document.StoryRanges
            changed = changed + WU_ApplyCharacterStyleBatchInStoryChain(firstStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord)
        Next firstStory
    End If
    WU_ApplyCharacterStyleBatch = changed
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

' Apply the same literal/style pairs only inside an exact caller-supplied
' Range. This keeps a multi-rule citation pass from leaking into adjacent
' paragraphs, notes, or other stories while retaining one undo record.
Public Function WU_ApplyCharacterStyleBatchInRange(ByVal target As Range, ByVal matches As Variant, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = True) As Long
    Dim document As Document, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim targetStart As Long, targetEnd As Long, firstRow As Long, lastRow As Long, firstColumn As Long, activeRows As Long
    Dim styleCache() As Style
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ApplyCharacterStyleBatchInRange", "target range is required"
    Set document = target.Document
    activeRows = WU_ValidateCharacterStyleBatch(document, matches, styleCache)
    If activeRows = 0 Then Exit Function
    targetStart = target.Start: targetEnd = target.End
    If targetEnd <= targetStart Then Exit Function
    firstRow = LBound(matches, 1): lastRow = UBound(matches, 1): firstColumn = LBound(matches, 2)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Style literal matches batch": opened = True
    WU_ApplyCharacterStyleBatchInRange = WU_ApplyCharacterStyleBatchInStory(target, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord)
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

' Explicit wildcard counterparts for journal rules that need Word's native
' pattern language (for example, citation variants with a variable year or
' pin). Literal APIs above always escape wildcard syntax; callers must opt in
' here when they want pattern matching or replacement backreferences such as
' \1. The same story scoping, one-record undo, and exact-range guarantees
' apply.
Public Function WU_ReplaceWildcard(ByVal document As Document, ByVal pattern As String, ByVal replacement As String, Optional ByVal storyScope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False) As Boolean
    Dim updating As Boolean, opened As Boolean, captured As Boolean, changed As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ReplaceWildcard", "document is required"
    WU_ValidateWildcard pattern, replacement, "WU_ReplaceWildcard"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_ReplaceWildcard", "story scope must be main, notes, headers, footers, or all"
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Replace wildcard text": opened = True
    changed = WU_ReplaceWildcardInScope(document, pattern, replacement, storyScope, matchCase, wholeWord)
    WU_ReplaceWildcard = changed
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

' Replace only inside the exact caller-supplied Range using Word wildcards.
Public Function WU_ReplaceWildcardInRange(ByVal target As Range, ByVal pattern As String, ByVal replacement As String, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False) As Boolean
    Dim updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ReplaceWildcardInRange", "target range is required"
    WU_ValidateWildcard pattern, replacement, "WU_ReplaceWildcardInRange"
    If target.End <= target.Start Then Exit Function
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Replace wildcard text": opened = True
    WU_ReplaceWildcardInRange = WU_ReplaceLiteralInStory(target, pattern, replacement, matchCase, wholeWord, True)
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

' Apply an ordered two-column wildcard replacement table in one edit. Rows
' remain the compact command list; Word owns matching, text, and formatting.
Public Function WU_ReplaceWildcardBatch(ByVal document As Document, ByVal replacements As Variant, Optional ByVal storyScope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False) As Long
    Dim updating As Boolean, opened As Boolean, captured As Boolean, activeRows As Long
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, row As Long, changed As Long
    Dim matched() As Boolean, failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ReplaceWildcardBatch", "document is required"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_ReplaceWildcardBatch", "story scope must be main, notes, headers, footers, or all"
    activeRows = WU_ValidateLiteralBatch(replacements, matchCase, "WU_ReplaceWildcardBatch", True)
    If activeRows = 0 Then Exit Function
    firstRow = LBound(replacements, 1): lastRow = UBound(replacements, 1): firstColumn = LBound(replacements, 2)
    ReDim matched(firstRow To lastRow)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Replace wildcard text batch": opened = True
    WU_ReplaceWildcardBatchInScope document, replacements, firstRow, lastRow, firstColumn, storyScope, matchCase, wholeWord, matched
    For row = firstRow To lastRow
        If matched(row) Then changed = changed + 1
    Next row
    WU_ReplaceWildcardBatch = changed
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

Public Function WU_ReplaceWildcardBatchInRange(ByVal target As Range, ByVal replacements As Variant, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False) As Long
    Dim updating As Boolean, opened As Boolean, captured As Boolean, activeRows As Long
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, row As Long, changed As Long
    Dim matched() As Boolean, failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ReplaceWildcardBatchInRange", "target range is required"
    activeRows = WU_ValidateLiteralBatch(replacements, matchCase, "WU_ReplaceWildcardBatchInRange", True)
    If activeRows = 0 Or target.End <= target.Start Then Exit Function
    firstRow = LBound(replacements, 1): lastRow = UBound(replacements, 1): firstColumn = LBound(replacements, 2)
    ReDim matched(firstRow To lastRow)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Replace wildcard text batch": opened = True
    WU_ReplaceLiteralBatchInStory target, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, True
    For row = firstRow To lastRow
        If matched(row) Then changed = changed + 1
    Next row
    WU_ReplaceWildcardBatchInRange = changed
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

' Read-only wildcard count. It never opens an undo record or changes
' ScreenUpdating, so agents can cheaply preflight a rule before editing.
Public Function WU_CountWildcard(ByVal document As Document, ByVal pattern As String, Optional ByVal storyScope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False) As Long
    If document Is Nothing Then Err.Raise 91, "WU_CountWildcard", "document is required"
    WU_ValidateWildcard pattern, "", "WU_CountWildcard"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_CountWildcard", "story scope must be main, notes, headers, footers, or all"
    WU_CountWildcard = WU_CountWildcardInScope(document, pattern, storyScope, matchCase, wholeWord)
End Function

Public Function WU_CountWildcardInRange(ByVal target As Range, ByVal pattern As String, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False) As Long
    If target Is Nothing Then Err.Raise 91, "WU_CountWildcardInRange", "target range is required"
    WU_ValidateWildcard pattern, "", "WU_CountWildcardInRange"
    If target.End <= target.Start Then Exit Function
    WU_CountWildcardInRange = WU_CountLiteralInStory(target, pattern, matchCase, wholeWord, True)
End Function

' Apply a character or linked style to wildcard matches without rewriting
' their text. This is useful for citation variants where literal matching is
' too narrow but the intended styling is still a named, editable style.
Public Function WU_ApplyCharacterStyleToWildcardMatches(ByVal document As Document, ByVal pattern As String, ByVal styleName As String, Optional ByVal storyScope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = True) As Long
    Dim updating As Boolean, opened As Boolean, captured As Boolean, changed As Long, styleError As Long
    Dim failure As Long, failureSource As String, failureText As String, style As Style
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ApplyCharacterStyleToWildcardMatches", "document is required"
    WU_ValidateWildcard pattern, "", "WU_ApplyCharacterStyleToWildcardMatches"
    If Len(Trim$(styleName)) = 0 Then Err.Raise 5, "WU_ApplyCharacterStyleToWildcardMatches", "style name is required"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_ApplyCharacterStyleToWildcardMatches", "story scope must be main, notes, headers, footers, or all"
    On Error Resume Next
    Set style = document.Styles(styleName)
    styleError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If styleError <> 0 Or style Is Nothing Then Err.Raise 5, "WU_ApplyCharacterStyleToWildcardMatches", "style " & styleName & " was not found"
    If style.Type <> wdStyleTypeCharacter And Not style.Linked Then Err.Raise 5, "WU_ApplyCharacterStyleToWildcardMatches", "style is not a character style"
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Style wildcard matches": opened = True
    changed = WU_ApplyWildcardStyleInScope(document, pattern, style, storyScope, matchCase, wholeWord)
    WU_ApplyCharacterStyleToWildcardMatches = changed
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

Public Function WU_ApplyCharacterStyleToWildcardRange(ByVal target As Range, ByVal pattern As String, ByVal styleName As String, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = True) As Long
    Dim updating As Boolean, opened As Boolean, captured As Boolean, styleError As Long
    Dim failure As Long, failureSource As String, failureText As String, style As Style, document As Document
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ApplyCharacterStyleToWildcardRange", "target range is required"
    WU_ValidateWildcard pattern, "", "WU_ApplyCharacterStyleToWildcardRange"
    If Len(Trim$(styleName)) = 0 Then Err.Raise 5, "WU_ApplyCharacterStyleToWildcardRange", "style name is required"
    Set document = target.Document
    On Error Resume Next
    Set style = document.Styles(styleName)
    styleError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If styleError <> 0 Or style Is Nothing Then Err.Raise 5, "WU_ApplyCharacterStyleToWildcardRange", "style " & styleName & " was not found"
    If style.Type <> wdStyleTypeCharacter And Not style.Linked Then Err.Raise 5, "WU_ApplyCharacterStyleToWildcardRange", "style is not a character style"
    If target.End <= target.Start Then Exit Function
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Style wildcard matches": opened = True
    WU_ApplyCharacterStyleToWildcardRange = WU_ApplyCharacterStyleInStory(target, pattern, style, matchCase, wholeWord, True)
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

' Apply a bounded two-column [wildcard pattern, character-style] table in one
' edit. This is the multi-rule citation path: each style is resolved once,
' Word performs the native wildcard search, and all matches share one undo
' record. Literal batch styling remains separate and keeps its escaping rules.
Public Function WU_ApplyCharacterStyleToWildcardBatch(ByVal document As Document, ByVal matches As Variant, Optional ByVal storyScope As String = "main", Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = True) As Long
    Dim firstStory As Range, story As Range, updating As Boolean, opened As Boolean, captured As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, activeRows As Long, changed As Long
    Dim styleCache() As Style
    On Error GoTo Failed
    If document Is Nothing Then Err.Raise 91, "WU_ApplyCharacterStyleToWildcardBatch", "document is required"
    storyScope = LCase$(Trim$(storyScope))
    If storyScope <> "main" And storyScope <> "notes" And storyScope <> "headers" And storyScope <> "footers" And storyScope <> "all" Then Err.Raise 5, "WU_ApplyCharacterStyleToWildcardBatch", "story scope must be main, notes, headers, footers, or all"
    activeRows = WU_ValidateCharacterStyleBatch(document, matches, styleCache, True, "WU_ApplyCharacterStyleToWildcardBatch")
    If activeRows = 0 Then Exit Function
    firstRow = LBound(matches, 1): lastRow = UBound(matches, 1): firstColumn = LBound(matches, 2)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Style wildcard matches batch": opened = True
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If story.End > story.Start Then changed = WU_ApplyCharacterStyleBatchInStory(story, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, True)
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then changed = WU_ApplyCharacterStyleBatchInStoryChain(firstStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, True)
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo Failed
        If Not firstStory Is Nothing Then changed = changed + WU_ApplyCharacterStyleBatchInStoryChain(firstStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, True)
    ElseIf storyScope = "headers" Then
        changed = changed + WU_ApplyCharacterStyleBatchInStoryType(document, wdPrimaryHeaderStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, True)
        changed = changed + WU_ApplyCharacterStyleBatchInStoryType(document, wdFirstPageHeaderStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, True)
        changed = changed + WU_ApplyCharacterStyleBatchInStoryType(document, wdEvenPagesHeaderStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, True)
    ElseIf storyScope = "footers" Then
        changed = changed + WU_ApplyCharacterStyleBatchInStoryType(document, wdPrimaryFooterStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, True)
        changed = changed + WU_ApplyCharacterStyleBatchInStoryType(document, wdFirstPageFooterStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, True)
        changed = changed + WU_ApplyCharacterStyleBatchInStoryType(document, wdEvenPagesFooterStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, True)
    Else
        For Each firstStory In document.StoryRanges
            changed = changed + WU_ApplyCharacterStyleBatchInStoryChain(firstStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, True)
        Next firstStory
    End If
    WU_ApplyCharacterStyleToWildcardBatch = changed
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

' Apply the wildcard/style table only inside the exact caller-supplied Range.
Public Function WU_ApplyCharacterStyleToWildcardBatchInRange(ByVal target As Range, ByVal matches As Variant, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = True) As Long
    Dim document As Document, updating As Boolean, opened As Boolean, captured As Boolean, activeRows As Long
    Dim failure As Long, failureSource As String, failureText As String
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, changed As Long
    Dim styleCache() As Style
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 91, "WU_ApplyCharacterStyleToWildcardBatchInRange", "target range is required"
    Set document = target.Document
    activeRows = WU_ValidateCharacterStyleBatch(document, matches, styleCache, True, "WU_ApplyCharacterStyleToWildcardBatchInRange")
    If activeRows = 0 Or target.End <= target.Start Then Exit Function
    firstRow = LBound(matches, 1): lastRow = UBound(matches, 1): firstColumn = LBound(matches, 2)
    updating = Application.ScreenUpdating
    captured = True
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Style wildcard matches batch": opened = True
    WU_ApplyCharacterStyleToWildcardBatchInRange = WU_ApplyCharacterStyleBatchInStory(target, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, True)
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

Private Function WU_ReplaceWildcardInScope(ByVal document As Document, ByVal pattern As String, ByVal replacement As String, ByVal storyScope As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean) As Boolean
    Dim firstStory As Range, story As Range, changed As Boolean
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If story.End > story.Start Then changed = WU_ReplaceLiteralInStory(story, pattern, replacement, matchCase, wholeWord, True)
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo 0
        If Not firstStory Is Nothing Then changed = WU_ReplaceLiteralInStoryChain(firstStory, pattern, replacement, matchCase, wholeWord, True)
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo 0
        If Not firstStory Is Nothing Then If WU_ReplaceLiteralInStoryChain(firstStory, pattern, replacement, matchCase, wholeWord, True) Then changed = True
    ElseIf storyScope = "headers" Then
        If WU_ReplaceLiteralInStoryType(document, wdPrimaryHeaderStory, pattern, replacement, matchCase, wholeWord, True) Then changed = True
        If WU_ReplaceLiteralInStoryType(document, wdFirstPageHeaderStory, pattern, replacement, matchCase, wholeWord, True) Then changed = True
        If WU_ReplaceLiteralInStoryType(document, wdEvenPagesHeaderStory, pattern, replacement, matchCase, wholeWord, True) Then changed = True
    ElseIf storyScope = "footers" Then
        If WU_ReplaceLiteralInStoryType(document, wdPrimaryFooterStory, pattern, replacement, matchCase, wholeWord, True) Then changed = True
        If WU_ReplaceLiteralInStoryType(document, wdFirstPageFooterStory, pattern, replacement, matchCase, wholeWord, True) Then changed = True
        If WU_ReplaceLiteralInStoryType(document, wdEvenPagesFooterStory, pattern, replacement, matchCase, wholeWord, True) Then changed = True
    Else
        For Each firstStory In document.StoryRanges
            If WU_ReplaceLiteralInStoryChain(firstStory, pattern, replacement, matchCase, wholeWord, True) Then changed = True
        Next firstStory
    End If
    WU_ReplaceWildcardInScope = changed
End Function

Private Sub WU_ReplaceWildcardBatchInScope(ByVal document As Document, ByVal replacements As Variant, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstColumn As Long, ByVal storyScope As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, ByRef matched() As Boolean)
    Dim firstStory As Range, story As Range
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If story.End > story.Start Then WU_ReplaceLiteralBatchInStory story, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, True
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo 0
        If Not firstStory Is Nothing Then WU_ReplaceLiteralBatchInStoryChain firstStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, True
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo 0
        If Not firstStory Is Nothing Then WU_ReplaceLiteralBatchInStoryChain firstStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, True
    ElseIf storyScope = "headers" Then
        WU_ReplaceLiteralBatchInStoryType document, wdPrimaryHeaderStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, True
        WU_ReplaceLiteralBatchInStoryType document, wdFirstPageHeaderStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, True
        WU_ReplaceLiteralBatchInStoryType document, wdEvenPagesHeaderStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, True
    ElseIf storyScope = "footers" Then
        WU_ReplaceLiteralBatchInStoryType document, wdPrimaryFooterStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, True
        WU_ReplaceLiteralBatchInStoryType document, wdFirstPageFooterStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, True
        WU_ReplaceLiteralBatchInStoryType document, wdEvenPagesFooterStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, True
    Else
        For Each firstStory In document.StoryRanges
            WU_ReplaceLiteralBatchInStoryChain firstStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, True
        Next firstStory
    End If
End Sub

Private Function WU_CountWildcardInScope(ByVal document As Document, ByVal pattern As String, ByVal storyScope As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean) As Long
    Dim firstStory As Range, story As Range, count As Long
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If story.End > story.Start Then count = WU_CountLiteralInStory(story, pattern, matchCase, wholeWord, True)
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo 0
        If Not firstStory Is Nothing Then count = WU_CountLiteralInStoryChain(firstStory, pattern, matchCase, wholeWord, True)
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo 0
        If Not firstStory Is Nothing Then count = count + WU_CountLiteralInStoryChain(firstStory, pattern, matchCase, wholeWord, True)
    ElseIf storyScope = "headers" Then
        count = count + WU_CountLiteralInStoryType(document, wdPrimaryHeaderStory, pattern, matchCase, wholeWord, True)
        count = count + WU_CountLiteralInStoryType(document, wdFirstPageHeaderStory, pattern, matchCase, wholeWord, True)
        count = count + WU_CountLiteralInStoryType(document, wdEvenPagesHeaderStory, pattern, matchCase, wholeWord, True)
    ElseIf storyScope = "footers" Then
        count = count + WU_CountLiteralInStoryType(document, wdPrimaryFooterStory, pattern, matchCase, wholeWord, True)
        count = count + WU_CountLiteralInStoryType(document, wdFirstPageFooterStory, pattern, matchCase, wholeWord, True)
        count = count + WU_CountLiteralInStoryType(document, wdEvenPagesFooterStory, pattern, matchCase, wholeWord, True)
    Else
        For Each firstStory In document.StoryRanges
            count = count + WU_CountLiteralInStoryChain(firstStory, pattern, matchCase, wholeWord, True)
        Next firstStory
    End If
    WU_CountWildcardInScope = count
End Function

Private Function WU_ApplyWildcardStyleInScope(ByVal document As Document, ByVal pattern As String, ByVal style As Style, ByVal storyScope As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean) As Long
    Dim firstStory As Range, story As Range, changed As Long
    If storyScope = "main" Then
        Set story = document.StoryRanges(wdMainTextStory)
        If Not story Is Nothing Then If story.End > story.Start Then changed = WU_ApplyCharacterStyleInStory(story, pattern, style, matchCase, wholeWord, True)
    ElseIf storyScope = "notes" Then
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdFootnotesStory)
        Err.Clear
        On Error GoTo 0
        If Not firstStory Is Nothing Then changed = WU_ApplyCharacterStyleInStoryChain(firstStory, pattern, style, matchCase, wholeWord, True)
        Set firstStory = Nothing
        On Error Resume Next
        Set firstStory = document.StoryRanges(wdEndnotesStory)
        Err.Clear
        On Error GoTo 0
        If Not firstStory Is Nothing Then changed = changed + WU_ApplyCharacterStyleInStoryChain(firstStory, pattern, style, matchCase, wholeWord, True)
    ElseIf storyScope = "headers" Then
        changed = changed + WU_ApplyCharacterStyleInStoryType(document, wdPrimaryHeaderStory, pattern, style, matchCase, wholeWord, True)
        changed = changed + WU_ApplyCharacterStyleInStoryType(document, wdFirstPageHeaderStory, pattern, style, matchCase, wholeWord, True)
        changed = changed + WU_ApplyCharacterStyleInStoryType(document, wdEvenPagesHeaderStory, pattern, style, matchCase, wholeWord, True)
    ElseIf storyScope = "footers" Then
        changed = changed + WU_ApplyCharacterStyleInStoryType(document, wdPrimaryFooterStory, pattern, style, matchCase, wholeWord, True)
        changed = changed + WU_ApplyCharacterStyleInStoryType(document, wdFirstPageFooterStory, pattern, style, matchCase, wholeWord, True)
        changed = changed + WU_ApplyCharacterStyleInStoryType(document, wdEvenPagesFooterStory, pattern, style, matchCase, wholeWord, True)
    Else
        For Each firstStory In document.StoryRanges
            changed = changed + WU_ApplyCharacterStyleInStoryChain(firstStory, pattern, style, matchCase, wholeWord, True)
        Next firstStory
    End If
    WU_ApplyWildcardStyleInScope = changed
End Function

Private Sub WU_ValidateLiteral(ByVal findText As String, ByVal replaceText As String, ByVal sourceName As String)
    If Len(findText) = 0 Then Err.Raise 5, sourceName, "find text is required"
    If Len(findText) > 255 Then Err.Raise 5, sourceName, "find text exceeds Word's 255-character limit"
    If Len(replaceText) > 255 Then Err.Raise 5, sourceName, "replacement text exceeds Word's 255-character limit"
    If Len(WU_EscapeFindLiteral(findText)) > 255 Then Err.Raise 5, sourceName, "find text exceeds Word's escaped 255-character limit"
    If Len(WU_EscapeFindLiteral(replaceText)) > 255 Then Err.Raise 5, sourceName, "replacement text exceeds Word's escaped 255-character limit"
End Sub

Private Sub WU_ValidateWildcard(ByVal pattern As String, ByVal replacement As String, ByVal sourceName As String)
    If Len(pattern) = 0 Then Err.Raise 5, sourceName, "wildcard pattern is required"
    If Len(pattern) > 255 Then Err.Raise 5, sourceName, "wildcard pattern exceeds Word's 255-character limit"
    If Len(replacement) > 255 Then Err.Raise 5, sourceName, "wildcard replacement exceeds Word's 255-character limit"
End Sub

Private Function WU_ValidateLiteralBatch(ByVal replacements As Variant, ByVal matchCase As Boolean, ByVal sourceName As String, Optional ByVal useWildcards As Boolean = False) As Long
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, lastColumn As Long, row As Long
    Dim findText As String, replaceText As String, activeRows As Long, dimensionError As Long
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If Not IsArray(replacements) Then Err.Raise 5, sourceName, "replacements must be a two-dimensional array"
    On Error Resume Next
    firstRow = LBound(replacements, 1): lastRow = UBound(replacements, 1)
    firstColumn = LBound(replacements, 2): lastColumn = UBound(replacements, 2)
    dimensionError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If dimensionError <> 0 Then Err.Raise 5, sourceName, "replacements must be a two-dimensional array"
    If lastColumn - firstColumn + 1 <> 2 Then Err.Raise 5, sourceName, "replacements must have exactly two columns"
    If lastRow - firstRow + 1 > WU_MAX_BATCH_RULES Then Err.Raise 5, sourceName, "replacement rule count exceeds 1024"
    For row = firstRow To lastRow
        If IsError(replacements(row, firstColumn)) Or IsNull(replacements(row, firstColumn)) Or IsObject(replacements(row, firstColumn)) Or IsArray(replacements(row, firstColumn)) Then Err.Raise 5, sourceName, "replacement rule " & CStr(row) & " find text must be scalar"
        If IsError(replacements(row, firstColumn + 1)) Or IsNull(replacements(row, firstColumn + 1)) Or IsObject(replacements(row, firstColumn + 1)) Or IsArray(replacements(row, firstColumn + 1)) Then Err.Raise 5, sourceName, "replacement rule " & CStr(row) & " replacement text must be scalar"
        findText = CStr(replacements(row, firstColumn))
        replaceText = CStr(replacements(row, firstColumn + 1))
        If useWildcards Then
            WU_ValidateWildcard findText, replaceText, sourceName
            activeRows = activeRows + 1
        Else
            WU_ValidateLiteral findText, replaceText, sourceName
            If Not (matchCase And StrComp(findText, replaceText, vbBinaryCompare) = 0) Then activeRows = activeRows + 1
        End If
    Next row
    WU_ValidateLiteralBatch = activeRows
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
End Function

Private Function WU_ValidateCharacterStyleBatch(ByVal document As Document, ByVal matches As Variant, ByRef styleCache() As Style, Optional ByVal useWildcards As Boolean = False, Optional ByVal sourceName As String = "WU_ApplyCharacterStyleBatch") As Long
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, lastColumn As Long, row As Long, dimensionError As Long
    Dim findText As String, styleName As String, style As Style, activeRows As Long
    Dim cachedNames() As String, cachedStyles() As Style, cachedCount As Long, cacheIndex As Long, cacheRow As Long, cacheCapacity As Long
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If Not IsArray(matches) Then Err.Raise 5, sourceName, "matches must be a two-dimensional array"
    On Error Resume Next
    firstRow = LBound(matches, 1): lastRow = UBound(matches, 1)
    firstColumn = LBound(matches, 2): lastColumn = UBound(matches, 2)
    dimensionError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If dimensionError <> 0 Then Err.Raise 5, sourceName, "matches must be a two-dimensional array"
    If lastColumn - firstColumn + 1 <> 2 Then Err.Raise 5, sourceName, "matches must have exactly two columns"
    If lastRow - firstRow + 1 > WU_MAX_BATCH_RULES Then Err.Raise 5, sourceName, "style rule count exceeds 1024"
    cacheCapacity = lastRow - firstRow + 1
    ReDim styleCache(firstRow To lastRow)
    ReDim cachedNames(1 To cacheCapacity): ReDim cachedStyles(1 To cacheCapacity)
    For row = firstRow To lastRow
        If IsError(matches(row, firstColumn)) Or IsNull(matches(row, firstColumn)) Or IsObject(matches(row, firstColumn)) Or IsArray(matches(row, firstColumn)) Then Err.Raise 5, sourceName, "style rule " & CStr(row) & " pattern must be scalar"
        If IsError(matches(row, firstColumn + 1)) Or IsNull(matches(row, firstColumn + 1)) Or IsObject(matches(row, firstColumn + 1)) Or IsArray(matches(row, firstColumn + 1)) Then Err.Raise 5, sourceName, "style rule " & CStr(row) & " style name must be scalar"
        findText = CStr(matches(row, firstColumn))
        styleName = CStr(matches(row, firstColumn + 1))
        If useWildcards Then
            WU_ValidateWildcard findText, "", sourceName
        Else
            WU_ValidateLiteral findText, "", sourceName
        End If
        If Len(Trim$(styleName)) = 0 Then Err.Raise 5, sourceName, "style rule " & CStr(row) & " style name is required"
        cacheIndex = 0
        For cacheRow = 1 To cachedCount
            If StrComp(styleName, cachedNames(cacheRow), vbTextCompare) = 0 Then cacheIndex = cacheRow: Exit For
        Next cacheRow
        If cacheIndex = 0 Then
            Set style = Nothing
            On Error Resume Next
            Set style = document.Styles(styleName)
            If Err.Number <> 0 Or style Is Nothing Then
                Err.Clear
                On Error GoTo Failed
                Err.Raise 5, sourceName, "style rule " & CStr(row) & " names a missing style"
            End If
            Err.Clear
            On Error GoTo Failed
            cachedCount = cachedCount + 1: cacheIndex = cachedCount
            cachedNames(cacheIndex) = styleName: Set cachedStyles(cacheIndex) = style
        Else
            Set style = cachedStyles(cacheIndex)
        End If
        If style.Type <> wdStyleTypeCharacter And Not style.Linked Then Err.Raise 5, sourceName, "style rule " & CStr(row) & " style is not a character style"
        Set styleCache(row) = style
        activeRows = activeRows + 1
    Next row
    WU_ValidateCharacterStyleBatch = activeRows
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
End Function

Private Function WU_ValidateCharacterStyleRuns(ByVal document As Document, ByVal target As Range, ByVal runs As Variant, ByRef styleCache() As Style, ByRef styleNames() As String) As Long
    Const WU_MAX_CHARACTER_STYLE_RUNS As Long = 4096
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, lastColumn As Long, row As Long, dimensionError As Long
    Dim targetStart As Long, targetEnd As Long, startPosition As Long, endPosition As Long, previousEnd As Long
    Dim styleName As String, style As Style, styleError As Long
    Dim cachedNames() As String, cachedStyles() As Style, cachedCount As Long, cacheIndex As Long, cacheRow As Long, cacheCapacity As Long
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    If Not IsArray(runs) Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "runs must be a two-dimensional array"
    On Error Resume Next
    firstRow = LBound(runs, 1): lastRow = UBound(runs, 1)
    firstColumn = LBound(runs, 2): lastColumn = UBound(runs, 2)
    dimensionError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If dimensionError <> 0 Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "runs must be a two-dimensional array"
    If lastColumn - firstColumn + 1 <> 3 Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "runs must have exactly three columns"
    If lastRow - firstRow + 1 > WU_MAX_CHARACTER_STYLE_RUNS Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "character style run count exceeds 4096"
    targetStart = target.Start: targetEnd = target.End: previousEnd = targetStart
    cacheCapacity = lastRow - firstRow + 1
    ReDim styleCache(firstRow To lastRow): ReDim styleNames(firstRow To lastRow)
    ReDim cachedNames(1 To cacheCapacity): ReDim cachedStyles(1 To cacheCapacity)
    For row = firstRow To lastRow
        If IsError(runs(row, firstColumn)) Or IsNull(runs(row, firstColumn)) Or IsEmpty(runs(row, firstColumn)) Or IsObject(runs(row, firstColumn)) Or IsArray(runs(row, firstColumn)) Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "character style run " & CStr(row) & " start must be scalar"
        If IsError(runs(row, firstColumn + 1)) Or IsNull(runs(row, firstColumn + 1)) Or IsEmpty(runs(row, firstColumn + 1)) Or IsObject(runs(row, firstColumn + 1)) Or IsArray(runs(row, firstColumn + 1)) Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "character style run " & CStr(row) & " end must be scalar"
        If IsError(runs(row, firstColumn + 2)) Or IsNull(runs(row, firstColumn + 2)) Or IsEmpty(runs(row, firstColumn + 2)) Or IsObject(runs(row, firstColumn + 2)) Or IsArray(runs(row, firstColumn + 2)) Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "character style run " & CStr(row) & " style name must be scalar"
        If Not WU_ReadCharacterStylePosition(runs(row, firstColumn), startPosition) Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "character style run " & CStr(row) & " start must be an integer position"
        If Not WU_ReadCharacterStylePosition(runs(row, firstColumn + 1), endPosition) Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "character style run " & CStr(row) & " end must be an integer position"
        If startPosition < targetStart Or endPosition > targetEnd Or endPosition <= startPosition Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "character style run " & CStr(row) & " is outside the target range"
        If startPosition < previousEnd Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "character style runs must be ordered and non-overlapping"
        styleName = CStr(runs(row, firstColumn + 2))
        If Len(Trim$(styleName)) = 0 Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "character style run " & CStr(row) & " style name is required"
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
            If styleError <> 0 Or style Is Nothing Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "character style run " & CStr(row) & " names a missing style"
            cachedCount = cachedCount + 1: cacheIndex = cachedCount
            cachedNames(cacheIndex) = styleName: Set cachedStyles(cacheIndex) = style
        Else
            Set style = cachedStyles(cacheIndex)
        End If
        If style.Type <> wdStyleTypeCharacter And Not style.Linked Then Err.Raise 5, "WU_ApplyCharacterStyleRuns", "character style run " & CStr(row) & " style is not a character style"
        Set styleCache(row) = style: styleNames(row) = cachedNames(cacheIndex)
        previousEnd = endPosition
    Next row
    WU_ValidateCharacterStyleRuns = lastRow - firstRow + 1
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
End Function

Private Function WU_ReadCharacterStylePosition(ByVal value As Variant, ByRef position As Long) As Boolean
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
    WU_ReadCharacterStylePosition = True
End Function

Private Function WU_CharacterStyleMatches(ByVal target As Range, ByVal expectedName As String) As Boolean
    Dim currentStyle As String, readError As Long
    On Error Resume Next
    currentStyle = CStr(target.Style)
    readError = Err.Number
    Err.Clear
    On Error GoTo 0
    If readError = 0 Then WU_CharacterStyleMatches = (StrComp(currentStyle, expectedName, vbTextCompare) = 0)
End Function

Private Function WU_ApplyCharacterStyleBatchInStoryChain(ByVal firstStory As Range, ByVal matches As Variant, ByRef styleCache() As Style, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstColumn As Long, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, Optional ByVal useWildcards As Boolean = False) As Long
    Dim story As Range, changed As Long
    Set story = firstStory
    Do While Not story Is Nothing
        If story.End > story.Start Then changed = changed + WU_ApplyCharacterStyleBatchInStory(story, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, useWildcards)
        Set story = story.NextStoryRange
    Loop
    WU_ApplyCharacterStyleBatchInStoryChain = changed
End Function

Private Function WU_ApplyCharacterStyleBatchInStoryType(ByVal document As Document, ByVal storyType As Long, ByVal matches As Variant, ByRef styleCache() As Style, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstColumn As Long, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, Optional ByVal useWildcards As Boolean = False) As Long
    Dim firstStory As Range
    On Error Resume Next
    Set firstStory = document.StoryRanges(storyType)
    Err.Clear
    On Error GoTo 0
    If Not firstStory Is Nothing Then WU_ApplyCharacterStyleBatchInStoryType = WU_ApplyCharacterStyleBatchInStoryChain(firstStory, matches, styleCache, firstRow, lastRow, firstColumn, matchCase, wholeWord, useWildcards)
End Function

Private Function WU_ApplyCharacterStyleBatchInStory(ByVal story As Range, ByVal matches As Variant, ByRef styleCache() As Style, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstColumn As Long, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, Optional ByVal useWildcards As Boolean = False) As Long
    Dim row As Long, findText As String, style As Style, changed As Long
    For row = firstRow To lastRow
        findText = CStr(matches(row, firstColumn))
        Set style = styleCache(row)
        changed = changed + WU_ApplyCharacterStyleInStory(story, findText, style, matchCase, wholeWord, useWildcards)
    Next row
    WU_ApplyCharacterStyleBatchInStory = changed
End Function

Private Function WU_CountLiteralInStoryChain(ByVal firstStory As Range, ByVal findText As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, Optional ByVal useWildcards As Boolean = False) As Long
    Dim story As Range, count As Long
    Set story = firstStory
    Do While Not story Is Nothing
        If story.End > story.Start Then count = count + WU_CountLiteralInStory(story, findText, matchCase, wholeWord, useWildcards)
        Set story = story.NextStoryRange
    Loop
    WU_CountLiteralInStoryChain = count
End Function

Private Function WU_CountLiteralInStoryType(ByVal document As Document, ByVal storyType As Long, ByVal findText As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, Optional ByVal useWildcards As Boolean = False) As Long
    Dim firstStory As Range
    On Error Resume Next
    Set firstStory = document.StoryRanges(storyType)
    Err.Clear
    On Error GoTo 0
    If Not firstStory Is Nothing Then WU_CountLiteralInStoryType = WU_CountLiteralInStoryChain(firstStory, findText, matchCase, wholeWord, useWildcards)
End Function

Private Function WU_CountLiteralInStory(ByVal story As Range, ByVal findText As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, Optional ByVal useWildcards As Boolean = False) As Long
    Dim search As Range, nextStart As Long, storyEnd As Long, count As Long
    Set search = story.Duplicate
    ' Keep the fixed story boundary local; this loop may execute once per
    ' citation and should not ask Word for the same End property repeatedly.
    storyEnd = story.End
    With search.Find
        .ClearFormatting
        .Text = WU_FindPattern(findText, useWildcards)
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = matchCase
        .MatchWholeWord = wholeWord
        .MatchWildcards = useWildcards
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Call WU_PinFindOptions(search.Find)
    Do While search.Find.Execute
        count = count + 1
        nextStart = search.End
        ' A wildcard can legally match an empty span. Always advance a
        ' zero-width result so a permissive pattern cannot loop forever.
        If nextStart <= search.Start Then nextStart = search.Start + 1
        If nextStart >= storyEnd Then Exit Do
        search.SetRange Start:=nextStart, End:=storyEnd
    Loop
    WU_CountLiteralInStory = count
End Function

Private Sub WU_ReplaceLiteralBatchInStoryChain(ByVal firstStory As Range, ByVal replacements As Variant, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstColumn As Long, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, ByRef matched() As Boolean, Optional ByVal useWildcards As Boolean = False)
    Dim story As Range
    Set story = firstStory
    Do While Not story Is Nothing
        If story.End > story.Start Then WU_ReplaceLiteralBatchInStory story, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, useWildcards
        Set story = story.NextStoryRange
    Loop
End Sub

Private Sub WU_ReplaceLiteralBatchInStoryType(ByVal document As Document, ByVal storyType As Long, ByVal replacements As Variant, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstColumn As Long, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, ByRef matched() As Boolean, Optional ByVal useWildcards As Boolean = False)
    Dim firstStory As Range
    On Error Resume Next
    Set firstStory = document.StoryRanges(storyType)
    Err.Clear
    On Error GoTo 0
    If Not firstStory Is Nothing Then WU_ReplaceLiteralBatchInStoryChain firstStory, replacements, firstRow, lastRow, firstColumn, matchCase, wholeWord, matched, useWildcards
End Sub

Private Sub WU_ReplaceLiteralBatchInStory(ByVal story As Range, ByVal replacements As Variant, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstColumn As Long, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, ByRef matched() As Boolean, Optional ByVal useWildcards As Boolean = False)
    Dim row As Long, findText As String, replaceText As String
    For row = firstRow To lastRow
        findText = CStr(replacements(row, firstColumn))
        replaceText = CStr(replacements(row, firstColumn + 1))
        If useWildcards Or Not (matchCase And StrComp(findText, replaceText, vbBinaryCompare) = 0) Then
            If WU_ReplaceLiteralInStory(story, findText, replaceText, matchCase, wholeWord, useWildcards) Then matched(row) = True
        End If
    Next row
End Sub

Private Function WU_ApplyCharacterStyleInStoryChain(ByVal firstStory As Range, ByVal findText As String, ByVal style As Style, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, Optional ByVal useWildcards As Boolean = False) As Long
    Dim story As Range, changed As Long
    Set story = firstStory
    Do While Not story Is Nothing
        If story.End > story.Start Then changed = changed + WU_ApplyCharacterStyleInStory(story, findText, style, matchCase, wholeWord, useWildcards)
        Set story = story.NextStoryRange
    Loop
    WU_ApplyCharacterStyleInStoryChain = changed
End Function

Private Function WU_ApplyCharacterStyleInStoryType(ByVal document As Document, ByVal storyType As Long, ByVal findText As String, ByVal style As Style, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, Optional ByVal useWildcards As Boolean = False) As Long
    Dim firstStory As Range
    On Error Resume Next
    Set firstStory = document.StoryRanges(storyType)
    Err.Clear
    On Error GoTo 0
    If Not firstStory Is Nothing Then WU_ApplyCharacterStyleInStoryType = WU_ApplyCharacterStyleInStoryChain(firstStory, findText, style, matchCase, wholeWord, useWildcards)
End Function

Private Function WU_ApplyCharacterStyleInStory(ByVal story As Range, ByVal findText As String, ByVal style As Style, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, Optional ByVal useWildcards As Boolean = False) As Long
    Dim search As Range, nextStart As Long, storyEnd As Long, changed As Long, currentStyle As String, targetStyleName As String
    Set search = story.Duplicate
    ' Range.End is a COM property. Cache the fixed story boundary once rather
    ' than crossing the host boundary for every match in a long note story.
    storyEnd = story.End
    ' NameLocal is a COM property; resolve it once per story/rule rather
    ' than once for every match in a long citation-heavy document.
    targetStyleName = style.NameLocal
    With search.Find
        .ClearFormatting
        .Text = WU_FindPattern(findText, useWildcards)
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = matchCase
        .MatchWholeWord = wholeWord
        .MatchWildcards = useWildcards
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Call WU_PinFindOptions(search.Find)
    Do While search.Find.Execute
        currentStyle = vbNullString
        On Error Resume Next
        currentStyle = CStr(search.Style)
        Err.Clear
        On Error GoTo 0
        If StrComp(currentStyle, targetStyleName, vbTextCompare) <> 0 Then
            search.Style = style
            changed = changed + 1
        End If
        nextStart = search.End
        ' Keep wildcard patterns that match an empty span from re-finding the
        ' same position forever. Literal searches are unaffected by the guard.
        If nextStart <= search.Start Then nextStart = search.Start + 1
        If nextStart >= storyEnd Then Exit Do
        search.SetRange Start:=nextStart, End:=storyEnd
    Loop
    WU_ApplyCharacterStyleInStory = changed
End Function

Private Function WU_ReplaceLiteralInStoryChain(ByVal firstStory As Range, ByVal findText As String, ByVal replaceText As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, Optional ByVal useWildcards As Boolean = False) As Boolean
    Dim story As Range, changed As Boolean
    Set story = firstStory
    Do While Not story Is Nothing
        If story.End > story.Start Then
            If WU_ReplaceLiteralInStory(story, findText, replaceText, matchCase, wholeWord, useWildcards) Then changed = True
        End If
        Set story = story.NextStoryRange
    Loop
    WU_ReplaceLiteralInStoryChain = changed
End Function

Private Function WU_ReplaceLiteralInStoryType(ByVal document As Document, ByVal storyType As Long, ByVal findText As String, ByVal replaceText As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, Optional ByVal useWildcards As Boolean = False) As Boolean
    Dim firstStory As Range
    On Error Resume Next
    Set firstStory = document.StoryRanges(storyType)
    Err.Clear
    On Error GoTo 0
    If Not firstStory Is Nothing Then WU_ReplaceLiteralInStoryType = WU_ReplaceLiteralInStoryChain(firstStory, findText, replaceText, matchCase, wholeWord, useWildcards)
End Function

Private Function WU_ReplaceLiteralInStory(ByVal story As Range, ByVal findText As String, ByVal replaceText As String, ByVal matchCase As Boolean, ByVal wholeWord As Boolean, Optional ByVal useWildcards As Boolean = False) As Boolean
    Dim search As Range
    Set search = story.Duplicate
    With search.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        ' Find/Replace treats ^p, ^t, and similar sequences as structural
        ' tokens. Doubling the marker keeps this helper literal by default.
        .Text = WU_FindPattern(findText, useWildcards)
        .Replacement.Text = WU_ReplacementPattern(replaceText, useWildcards)
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = matchCase
        .MatchWholeWord = wholeWord
        .MatchWildcards = useWildcards
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Call WU_PinFindOptions(search.Find)
    WU_ReplaceLiteralInStory = search.Find.Execute(Replace:=wdReplaceAll)
End Function

Private Sub WU_PinFindOptions(ByVal criteria As Find)
    ' These options exist in current Word object libraries but are optional
    ' for some language packs/older hosts. Ignore only an unavailable option;
    ' the required text, scope, and replacement settings remain fail-fast.
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

Private Function WU_EscapeFindLiteral(ByVal value As String) As String
    WU_EscapeFindLiteral = Replace(value, "^", "^^")
End Function

Private Function WU_FindPattern(ByVal value As String, ByVal useWildcards As Boolean) As String
    If useWildcards Then WU_FindPattern = value Else WU_FindPattern = WU_EscapeFindLiteral(value)
End Function

Private Function WU_ReplacementPattern(ByVal value As String, ByVal useWildcards As Boolean) As String
    ' Wildcard callers may intentionally use Word replacement tokens such as
    ' \1. Literal callers retain the structural-token escaping contract.
    If useWildcards Then WU_ReplacementPattern = value Else WU_ReplacementPattern = WU_EscapeFindLiteral(value)
End Function
`
