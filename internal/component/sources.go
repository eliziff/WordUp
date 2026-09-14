package component

const safeEditSource = `Attribute VB_Name = "WordUpSafeEdit"
Option Explicit
' Call Begin and End from the editing procedure itself. Word Application.Run
' cannot reliably propagate an unhandled macro error across its COM boundary.
' The caller's handler must save Err before calling End, then re-raise it.
Public Sub WU_BeginSafeEdit(ByRef updating As Boolean, ByRef opened As Boolean, Optional ByVal label As String = "WordUp edit")
    Dim failure As Long, failureSource As String, failureText As String
    On Error GoTo Failed
    updating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    opened = False
    Application.UndoRecord.StartCustomRecord label: opened = True
    Exit Sub
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = updating
    On Error GoTo 0
    Err.Raise failure, failureSource, failureText
End Sub
Public Sub WU_EndSafeEdit(ByVal updating As Boolean, ByRef opened As Boolean)
    Dim failure As Long, failureSource As String, failureText As String
    On Error Resume Next
    If opened Then
        Application.UndoRecord.EndCustomRecord
        failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
        If failure = 0 Then opened = False
    End If
    Err.Clear
    Application.ScreenUpdating = updating
    If failure = 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
End Sub
`

const formShellSource = `Attribute VB_Name = "WordUpFormShell"
Option Explicit
Public Sub WU_ShowForm(ByVal formName As String)
    VBA.UserForms.Add(formName).Show
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
    Application.Run "WU_Command_" & Replace(control.Id, "-", "_")
End Sub
`

const hotkeySource = `Attribute VB_Name = "WordUpHotkey"
Option Explicit
Public Sub WU_RegisterHotkey(ByVal keyCode As Long, ByVal macroName As String)
    WU_ChangeHotkey keyCode, macroName, False
End Sub
Public Sub WU_RemoveHotkey(ByVal keyCode As Long)
    WU_ChangeHotkey keyCode, "", True
End Sub
Private Sub WU_ChangeHotkey(ByVal keyCode As Long, ByVal macroName As String, ByVal remove As Boolean)
    Dim prior As Object, failure As Long, failureSource As String, failureText As String
    Set prior = Application.CustomizationContext
    On Error GoTo Failed
    Application.CustomizationContext = ThisDocument
    If remove Then
        FindKey(keyCode).Clear
    Else
        KeyBindings.Add wdKeyCategoryMacro, macroName, keyCode
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
`

const contextMenuSource = `Attribute VB_Name = "WordUpContextMenu"
Option Explicit
Private Const WU_TAG As String = "WordUp.Component.ContextMenu"
Private WU_OwnedContextMenu As CommandBarButton
Public Sub WU_RegisterContextMenu(ByVal caption As String, ByVal macroName As String)
    WU_ChangeContextMenu caption, macroName, False
End Sub
Public Sub WU_RemoveContextMenu()
    WU_ChangeContextMenu "", "", True
End Sub
Public Function WU_ContextMenuRegistered(Optional ByVal expectedCaption As String = "", Optional ByVal expectedMacro As String = "") As Boolean
    On Error Resume Next
    WU_ContextMenuRegistered = Not (WU_OwnedContextMenu Is Nothing)
    If WU_ContextMenuRegistered Then WU_ContextMenuRegistered = (WU_OwnedContextMenu.Tag = WU_TAG)
    If WU_ContextMenuRegistered And expectedCaption <> "" Then WU_ContextMenuRegistered = (WU_OwnedContextMenu.Caption = expectedCaption)
    If WU_ContextMenuRegistered And expectedMacro <> "" Then WU_ContextMenuRegistered = (WU_OwnedContextMenu.OnAction = expectedMacro)
    On Error GoTo 0
End Function
Private Sub WU_ChangeContextMenu(ByVal caption As String, ByVal macroName As String, ByVal removeOnly As Boolean)
    Dim prior As Object, item As CommandBarButton
    Dim failure As Long, failureSource As String, failureText As String
    Set prior = Application.CustomizationContext
    On Error GoTo Failed
    Application.CustomizationContext = ThisDocument
    On Error Resume Next
    WU_OwnedContextMenu.Delete
    Set WU_OwnedContextMenu = Nothing
    Err.Clear
    On Error GoTo Failed
    If Not removeOnly Then
        Set item = CommandBars("Text").Controls.Add(Type:=msoControlButton, Temporary:=True)
        item.Caption = caption: item.OnAction = macroName: item.Tag = WU_TAG
        Set WU_OwnedContextMenu = item
    End If
CleanUp:
    On Error Resume Next
    If failure <> 0 Then Set WU_OwnedContextMenu = Nothing
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
`

const styleConverterSource = `Attribute VB_Name = "WordUpStyleConverter"
Option Explicit
Public Function WU_ConvertStyle(ByVal document As Document, ByVal fromStyle As String, ByVal toStyle As String) As Boolean
    Dim scope As Range, updating As Boolean, opened As Boolean
    Dim failure As Long, failureSource As String, failureText As String
    Dim sourceStyle As Style, targetStyle As Style
    On Error GoTo Failed
    updating = Application.ScreenUpdating
    Set sourceStyle = document.Styles(fromStyle)
    Set targetStyle = document.Styles(toStyle)
    If StrComp(sourceStyle.NameLocal, targetStyle.NameLocal, vbTextCompare) = 0 Then Exit Function
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Convert style": opened = True
    Set scope = document.Content.Duplicate
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
    WU_ConvertStyle = scope.Find.Execute(Replace:=wdReplaceAll)
CleanUp:
    On Error Resume Next
    If opened Then Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = updating
    If failure = 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function
`
