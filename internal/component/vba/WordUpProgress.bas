Attribute VB_Name = "WordUpProgress"
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
