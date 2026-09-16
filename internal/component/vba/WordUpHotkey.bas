Attribute VB_Name = "WordUpHotkey"
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
