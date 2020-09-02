Attribute VB_Name = "modVB5AddInTool"
Option Explicit

' **********************************************************
' API call to determin where the user's Windows directory is
' **********************************************************
#If Win32 Then
  Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#Else
  Declare Function GetWindowsDirectory Lib "kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
#End If


' -------------------------------------------------
' Calls the windows API to get the windows directory
' -------------------------------------------------
Public Function sGetWindowsDir$()
    Dim x As Integer
    Dim sT As String

    sT = String$(145, 0)              ' Size Buffer
    x = GetWindowsDirectory(sT, 145)  ' Make API Call
    sT = Left$(sT, x)                 ' Trim Buffer

    If Right$(sT, 1) <> "\" Then      ' Add \ if necessary
       sGetWindowsDir = sT + "\"
    Else
       sGetWindowsDir = sT
    End If
End Function


