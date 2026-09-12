Attribute VB_Name = "PlatformProbe"
Option Explicit
Public Function PlatformReport() As String
    Dim s As String
#If Mac Then
    s = "Mac=1;"
#Else
    s = "Mac=0;"
#End If
#If Win32 Then
    s = s & "Win32=1;"
#Else
    s = s & "Win32=0;"
#End If
#If Win64 Then
    s = s & "Win64=1;"
#Else
    s = s & "Win64=0;"
#End If
#If VBA7 Then
    s = s & "VBA7=1;"
#Else
    s = s & "VBA7=0;"
#End If
    PlatformReport = s & "Version=" & Application.Version
End Function
