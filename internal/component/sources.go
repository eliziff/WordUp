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
Public Sub WU_RequestCancel(): WU_Cancelled = True: End Sub
Public Sub WU_ResetProgress(): WU_Cancelled = False: End Sub
Public Function WU_CancelRequested() As Boolean
    DoEvents
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
    CustomizationContext = ThisDocument
    KeyBindings.Add wdKeyCategoryMacro, macroName, keyCode
End Sub
Public Sub WU_RemoveHotkey(ByVal keyCode As Long)
    CustomizationContext = ThisDocument
    FindKey(keyCode).Disable
End Sub
`

const contextMenuSource = `Attribute VB_Name = "WordUpContextMenu"
Option Explicit
Private Const WU_TAG As String = "WordUp.Component.ContextMenu"
Public Sub WU_RegisterContextMenu(ByVal caption As String, ByVal macroName As String)
    Dim item As CommandBarButton
    WU_RemoveContextMenu
    Set item = CommandBars("Text").Controls.Add(Type:=msoControlButton, Temporary:=True)
    item.Caption = caption: item.OnAction = macroName: item.Tag = WU_TAG
End Sub
Public Sub WU_RemoveContextMenu()
    Dim item As CommandBarControl
    For Each item In CommandBars("Text").Controls
        If item.Tag = WU_TAG Then item.Delete
    Next item
End Sub
`

const styleConverterSource = `Attribute VB_Name = "WordUpStyleConverter"
Option Explicit
Public Function WU_ConvertStyle(ByVal document As Document, ByVal fromStyle As String, ByVal toStyle As String) As Long
    Dim paragraph As Paragraph, updating As Boolean, opened As Boolean, changed As Long
    Dim failure As Long, failureSource As String, failureText As String
    Dim sourceStyle As Style, targetStyle As Style
    On Error GoTo Failed
    updating = Application.ScreenUpdating
    Set sourceStyle = document.Styles(fromStyle)
    Set targetStyle = document.Styles(toStyle)
    If StrComp(sourceStyle.NameLocal, targetStyle.NameLocal, vbTextCompare) = 0 Then Exit Function
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Convert style": opened = True
    For Each paragraph In document.Paragraphs
        If CStr(paragraph.Style.NameLocal) = sourceStyle.NameLocal Then paragraph.Style = targetStyle: changed = changed + 1
    Next paragraph
CleanUp:
    On Error Resume Next
    If opened Then Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = updating
    If failure = 0 Then failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, failureSource, failureText
    WU_ConvertStyle = changed
    Exit Function
Failed:
    failure = Err.Number: failureSource = Err.Source: failureText = Err.Description
    Resume CleanUp
End Function
`
