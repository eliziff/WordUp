Attribute VB_Name = "WordUpFormShell"
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
