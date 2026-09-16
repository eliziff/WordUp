Attribute VB_Name = "WordUpContextMenu"
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
