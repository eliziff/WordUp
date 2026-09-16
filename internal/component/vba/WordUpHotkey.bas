Attribute VB_Name = "WordUpHotkey"
Option Explicit
' Template-owned key bindings.
'
'   WU_RegisterHotkey keyCode, macroName   bind a key to a macro in this
'                                          template; idempotent
'   WU_RemoveHotkey keyCode                remove only a binding this template owns
'   WU_HotkeyRegistered(keyCode, macro)    whether this template owns the binding
'
' Bindings live in the template's customization context; the caller's context
' is restored on every exit. FindKey returns a KeyBinding object even for an
' unbound key (empty Command, KeyCategory wdKeyCategoryNil), so existence is
' judged by the category, never by the object being present. A binding owned
' by Word itself or by another template is never replaced or cleared.
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
    Dim prior As Object, binding As KeyBinding
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
        ' Normal template, a built-in command or another add-in.
        Set binding = WU_ExistingBinding(keyCode)
        If Not binding Is Nothing Then
            If Not WU_OwnedByThisTemplate(binding) Then Err.Raise 5, "WU_RegisterHotkey", "key binding belongs to Word or another template"
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
' The binding for keyCode in the current customization context, or Nothing
' when the key is unbound there.
Private Function WU_ExistingBinding(ByVal keyCode As Long) As KeyBinding
    Dim binding As KeyBinding, category As Long
    On Error Resume Next
    Set binding = FindKey(keyCode)
    If binding Is Nothing Then Exit Function
    category = wdKeyCategoryNil
    category = binding.KeyCategory
    On Error GoTo 0
    If category <> wdKeyCategoryNil Then Set WU_ExistingBinding = binding
End Function
Private Function WU_OwnedHotkey(ByVal keyCode As Long) As KeyBinding
    Dim binding As KeyBinding
    Set binding = WU_ExistingBinding(keyCode)
    If binding Is Nothing Then Exit Function
    If WU_OwnedByThisTemplate(binding) Then Set WU_OwnedHotkey = binding
End Function
' Only macro bindings can be owned; a built-in command binding is never
' cleared. Word reports the owner as a Template object even when ThisDocument
' is the opened Document, so ownership falls back to comparing full names.
Private Function WU_OwnedByThisTemplate(ByVal binding As KeyBinding) As Boolean
    Dim owner As Object, ownerName As String, thisName As String, category As Long
    On Error Resume Next
    category = wdKeyCategoryNil
    category = binding.KeyCategory
    If category <> wdKeyCategoryMacro Then Exit Function
    Set owner = binding.Context
    If owner Is Nothing Then Exit Function
    If owner Is ThisDocument Then WU_OwnedByThisTemplate = True: Exit Function
    ownerName = owner.FullName
    thisName = ThisDocument.FullName
    On Error GoTo 0
    If Len(ownerName) > 0 And Len(thisName) > 0 Then WU_OwnedByThisTemplate = (StrComp(ownerName, thisName, vbTextCompare) = 0)
End Function
' KeyBindings.Add resolves a macro name inside the customization context and
' stores it fully qualified (Project.Module.Macro). Application.Run's
' 'Template'!Macro syntax is not valid here and makes Word raise 5346.
Private Function WU_QualifiedMacro(ByVal macroName As String) As String
    WU_QualifiedMacro = Trim$(macroName)
End Function
' The procedure name without project, module or template qualification, so a
' stored Project.Module.Macro command compares with a caller's Module.Macro.
Private Function WU_MacroMember(ByVal macroName As String) As String
    Dim separator As Long
    separator = InStrRev(macroName, "!", -1, vbBinaryCompare)
    If separator > 0 Then macroName = Mid$(macroName, separator + 1)
    separator = InStrRev(macroName, ".", -1, vbBinaryCompare)
    If separator = 0 Then WU_MacroMember = macroName Else WU_MacroMember = Mid$(macroName, separator + 1)
End Function
