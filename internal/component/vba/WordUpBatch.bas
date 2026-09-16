Attribute VB_Name = "WordUpBatch"
Option Explicit
' Batch progress, per-stage timing, notification capture and cooperative
' cancellation for long-running template macros.
'
'   WU_BatchStart steps        begin a batch: clears the log and cancellation
'   WU_BatchStage name         start a named stage; the previous stage's
'                              elapsed time is logged as "stage_ms<TAB>name<TAB>ms"
'   WU_BatchEnd                close the batch and clear the status bar
'   WU_BatchReport             Array(failed, log, step, total) for the caller
'                              or a test harness
'   WU_Notify message, buttons MsgBox outside a batch; inside a batch the
'                              message is logged and shown on the status bar,
'                              and critical/exclamation buttons mark failure
'   WU_CancelRequested         cooperative checkpoint; yields to Word at most
'                              every 50 ms and reports WU_RequestCancel
'   WU_ResetProgress           clear cancellation (also done by WU_BatchStart)
Private WU_TotalSteps As Long
Private WU_CurrentStep As Long
Private WU_Active As Boolean
Private WU_Failed As Boolean
Private WU_Log As String
Private WU_StageStarted As Single
Private WU_StageName As String
Private WU_Cancelled As Boolean
Private WU_LastYield As Single

Public Sub WU_BatchStart(ByVal steps As Long)
    WU_Active = True
    WU_Failed = False
    WU_Log = ""
    WU_StageName = ""
    WU_TotalSteps = steps
    WU_CurrentStep = 0
    WU_ResetProgress
    Application.StatusBar = "Starting..."
End Sub

Public Sub WU_BatchStage(Optional ByVal description As String = "")
    Dim percent As Long
    WU_FinishStage
    WU_StageName = description
    WU_StageStarted = Timer
    WU_CurrentStep = WU_CurrentStep + 1
    If WU_TotalSteps > 0 Then percent = CLng((WU_CurrentStep / WU_TotalSteps) * 100)
    Application.StatusBar = "Progress: " & CStr(percent) & "% " & description
    DoEvents
End Sub

Public Sub WU_BatchEnd()
    WU_FinishStage
    WU_Active = False
    Application.StatusBar = False
End Sub

Public Function WU_BatchActive() As Boolean
    WU_BatchActive = WU_Active
End Function

Public Function WU_BatchReport() As Variant
    WU_BatchReport = Array(WU_Failed, WU_Log, WU_CurrentStep, WU_TotalSteps)
End Function

Public Sub WU_Notify(ByVal message As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly)
    If WU_Active Then
        WU_Log = WU_Log & "message" & vbTab & message & vbCrLf
        If (buttons And vbCritical) <> 0 Or (buttons And vbExclamation) <> 0 Then WU_Failed = True
        Application.StatusBar = message
    Else
        MsgBox message, buttons
    End If
End Sub

Public Sub WU_RequestCancel()
    WU_Cancelled = True
End Sub

Public Sub WU_ResetProgress()
    WU_Cancelled = False
    WU_LastYield = 0
End Sub

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

Private Sub WU_FinishStage()
    Dim elapsed As Double
    If Len(WU_StageName) = 0 Then Exit Sub
    elapsed = Timer - WU_StageStarted
    If elapsed < 0 Then elapsed = elapsed + 86400
    WU_Log = WU_Log & "stage_ms" & vbTab & WU_StageName & vbTab & CStr(elapsed * 1000) & vbCrLf
    WU_StageName = ""
End Sub
