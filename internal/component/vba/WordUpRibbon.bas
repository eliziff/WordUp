Attribute VB_Name = "WordUpRibbon"
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
' Word's Application.Run resolves a plain procedure name inside the calling
' template's project first. The Excel-style 'Template'!Macro qualification is
' not valid in Word and raises 438 (verified natively).
Private Function WU_QualifiedMacro(ByVal macroName As String) As String
    WU_QualifiedMacro = macroName
End Function
