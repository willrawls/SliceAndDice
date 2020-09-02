Attribute VB_Name = "modGeneral"
Option Explicit

' ********************************************************************************
' Class Module      modGeneral
'
' Filename          cls
'
' Author            William M. Rawls
'
' Created On        9/3/1997 8:00 pm
'
' Description
'
' General functions
'
' ********************************************************************************

' ************************
' Publicly available stuff
' ************************
' True if processing is occurring that should cause any cascading events to exit immediately (search for gbProcessing to see impact)
  'Public gbProcessing As Boolean
  
' True if the user cancel processing while doing an insertion
 'Public CancelInsertion As Boolean
  
'
  Public gbEvaluationHasExpired As Boolean

' ***************************************************
' Publicly available constant strings
'   Call InitPublic() to set at beginning of program
' Why ? These strings are very common in VB
'   and using the Publicly available
' ***************************************************

  Public Const gsB As String = vbNullString
  Public Const gsQ As String = """"
  Public Const gsE As String = "="
  Public Const gsA As String = "'"
  Public Const gsBO As String = "{"
  Public Const gsBC As String = "}"
  Public Const gsC As String = ","
  
  Public Const gcPC As String = ")"
  Public Const gcPO As String = "("
  Public Const gsS As String = " "
  Public Const gsSC As String = ";"
  Public Const gsFindBO = "Find{"
  Public Const gsSelectFrom As String = "SELECT * FROM "
  Public Const gsWhere As String = " WHERE "

  Public Const gsSoftVarDelimiter As String = "%%"
  Public Const gsSoftCmdDelimiter As String = "~~"

' ***********************************
' ****** BrowseForFolder stuff ******
' ***********************************
  Private Type BrowseInfo
          hWndOwner As Long
          pIDLRoot As Long
          pszDisplayName As String
          lpszTitle As String
          ulFlags As Long
          lpfnCallback As Long
          lParam As Long
          iImage As Long
  End Type
  
  Private Const BIF_RETURNONLYFSDIRS = 1
  Private Const MAX_PATH = 260
  Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
  Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
  Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

' **********************************************************
' API call to determin where the user's Windows directory is
' **********************************************************
  Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function sChoose(sChoices As String, Optional ByVal sDelimiter As String = ";", Optional ByVal sDefault As String)
On Error GoTo EH_Wizard_sChoose
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Function
    bInHereAlready = True

    If Len(sDelimiter) = 0 Then sDelimiter = ";"

    Dim frmX As SandySupport.ISandyWindowSelect
    Set frmX = New frmListSelect
    With frmX
         .Initialize sChoices, sDelimiter, sDefault
         .ZOrder
         .Show vbModal
         sChoose = .Choice
    End With

EH_Wizard_sChoose_Continue:
    bInHereAlready = False
    Exit Function

EH_Wizard_sChoose:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: Wizard" & vbCr & vbTab & "Procedure: sChoose" & vbCr & vbCr & Err.Description
    Resume EH_Wizard_sChoose_Continue

    Resume
End Function

Public Sub BrowseTo(sURL As String)
    Static WinVer As String
    Static WebBrowserCommand As String

    If Len(WinVer) = 0 Then
       WinVer = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion", "Version")
       If WinVer = "Windows 98" Then
          WebBrowserCommand = "start "
       Else
          WebBrowserCommand = "explorer "
       End If
    End If

    Shell WebBrowserCommand & sURL, vbNormalFocus
End Sub

Public Function sGetGUID(ByVal sProgID As String) As String
On Error Resume Next
    sGetGUID = GetStringValue("HKEY_CLASSES_ROOT\" & sProgID & "\CLSID", vbNullString)
End Function



