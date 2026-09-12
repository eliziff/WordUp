Attribute VB_Name = "StudioForm"
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdNormalize_Click()
    Studio.NormalizeSelection
End Sub
