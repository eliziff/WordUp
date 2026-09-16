Attribute VB_Name = "WordUpSafeEdit"
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
