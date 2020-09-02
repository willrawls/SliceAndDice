Attribute VB_Name = "modGeneral"
Option Explicit

Public Const SRCCOPY = &HCC0020
Public winSysDir As String
Public winDir As String
Public gbWin31 As Long, gbWin95 As Long

Declare Function GetFileVersionInfoSize Lib "ver.dll" (ByVal lpszFileName As String, lpdwHandle As Long) As Long
Declare Function GetFileVersionInfo Lib "ver.dll" (ByVal lpszFileName As String, ByVal lpdwHandle As Long, ByVal cbbuf As Long, ByVal lpvdata As String) As Long
Declare Function VerQueryValue Lib "ver.dll" (ByVal lpvBlock As String, ByVal lpszSubBlock As String, lplpBuffer As Long, lpcb As Long) As Long

Type OFStruct
    cBytes As String * 1
    fFixedDisk As String * 1
    nErrCode As Long
    Reserved As String * 4
    szPathName As String * 128
End Type

' ---------------------------
' Start additions by WMR
' ---------------------------

Type UT_APP
     sTitle As String
     sDesc As String
     sDir As String
     bNecessary As Long
     bShareware As Long
     nFirstFile As Long
     nLastFile As Long
     nMainProg As Long
End Type
Public gaApp() As UT_APP, gnAppCount As Long

Type UT_FILE
     nApp As Long
     sName As String
     lSize As Long
     sAddPath As String
     bINI As Long
     bDriver As Long
End Type
Public gaFile() As UT_FILE, gnFileCount As Long

Type UT_Assoc
     sKey As String
     sValue As String
End Type
Public gaAssoc() As UT_Assoc, gnAssocCount As Long

Public gbDoInstallation As Long
Public glSpaceNeeded&, glSysSpaceNeeded&
Public gnCurrDisk As Long

Public gnRegEditCount As Long, gnCities As Long

Public gaLabel() As String
Public gnLabelCount As Long

Public iNextCellAcross As Long

' ************************
' Globally available stuff
' ************************
' True if processing is occurring that should cause any cascading events to exit immediately (search for gbProcessing to see impact)
  Public gbProcessing As Boolean

' *************************
' Set window position stuff
' *************************
  Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
  Private Const HWND_TOPMOST = -1
  Private Const SWP_NOMOVE = &H2
  Private Const SWP_NOSIZE = &H1
  Private Const TOPFLAGS = SWP_NOMOVE Or SWP_NOSIZE

' ***********************************
' ****** Extend ListView stuff ******
' ***********************************
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Const LVM_FIRST = &H1000
  Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
  Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
  Private Const LVS_EX_FULLROWSELECT = &H20

' ***********************************
' ****** Extend ListView stuff ******
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
  

' ************
' Memory stuff
' ************
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal numBytes As Long)

' **********************************************************
' API call to determin where the user's Windows directory is
' **********************************************************
  Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
  Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' -Title(Testing 123)                                           Sets the page's title
' -Dir(W:\Graphics\backgrounds)                                 Sets the path to walk
' -Quiet                                                        No visual interface 8)
' -Output(W:\Graphics\backgrounds\Folder.htm)                   Send results of walk to the named file
' -Build                                                        Use current settings to build the page
' -Put                                                          Place the page on HTTP site
' -ExploreLocal                                                 Cause browser to appear with outputed file's contents.

Public gsEOLTab              As String
Public gs2EOL                As String
Public gsLinePrefix          As String
Public gsLineSuffix          As String
Public gsOut                 As String
Public gsDefaultHTMLFilename As String
Public gsStartPath           As String
Public gsFileURLPath         As String

Public gbQuiet               As Boolean
Public gbPageGenerated       As Boolean
Public glInterruptBuild      As Long

Public Const gsTimeOut                  As String = "10"
Public Const gsHTMLFileWalkerHomePage   As String = "http://www.firmsolutions.com/htmlfilewalker.html"

Public Sub Main()
    With frmMain
         .Show
         If Len(.txtTitle) = 0 And Len(.txtOutFilename) = 0 And Len(.txtStartPath) = 0 Then
            .cmdReset_Click
         End If
    End With
End Sub

Public Function sCapitalize(ByVal sValue As String)
    Static i As Long
    i = Len(sValue)
    If i > 1 Then
       sCapitalize = UCase$(Left$(sValue, 1)) & LCase$(Mid$(sValue, 2))
    ElseIf i = 1 Then
       sCapitalize = UCase$(sValue)
    Else
       sCapitalize = vbNullString
    End If
End Function

Public Function sGetExtension(ByVal sPath As String)
    Static sT As String
    sT = sGetFilename(sPath)
    If InStr(sT, ".") > 0 Then
       sGetExtension = sGetToken(sT, nTokens(sT, "."), ".")
      'sGetExtension = Mid$(sT, InStr(sT, ".") + 1)
    Else
       sGetExtension = vbNullString
    End If
End Function

Public Function sGetFilename(ByVal sPath As String)
    
    If Right$(sPath, 1) = "\" Then sPath = Left$(sPath, Len(sPath) - 1)
    Do While InStr(sPath, "\")
       sPath = Mid$(sPath, InStr(sPath, "\") + 1)
    Loop
    sGetFilename = sPath
End Function

Public Function sGetPath(ByVal sPath As String)
    Static i As Long
    Static j As Long

    i = 1
    j = InStr(sPath, "\")
    Do While j
       i = i + j
       j = InStr(Mid$(sPath, i), "\")
    Loop
    If i = 1 Then
       sGetPath = vbNullString
    Else
       sGetPath = Left$(sPath, i - 1)
    End If
End Function




' -------------------------------------------------
' Calls the windows API to get the windows directory
' -------------------------------------------------
Public Function sGetWindowsDir() As String
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




' ********************************************************************************
' Name              Functions_BrowseForFolder
'
' Parameters
'      hWndOwner                     (I)  Window handle of owner
'      sPrompt                       (I)  Browse window caption
'
' Description
'
' Allows the user to "browse" for a directory (32 bit only!!!)
'
' ********************************************************************************
Public Function BrowseForFolder(ByVal hWndOwner As Long, ByVal sPrompt As String) As String
    Dim iNull As Long
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo

    With udtBI
         .hWndOwner = hWndOwner
         .lpszTitle = sPrompt
         .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
       sPath = String(MAX_PATH, 0)
       lResult = SHGetPathFromIDList(lpIDList, sPath)
       Call CoTaskMemFree(lpIDList)
       iNull = InStr(sPath, vbNullChar)
       If iNull Then
          sPath = Left$(sPath, iNull - 1)
       End If
    End If

    BrowseForFolder = sPath

End Function


' ********************************************************************************
' Name              Functions_bUserSure
'
' Parameters
'       sPrompt                      (I)  Opt. Question to ask the user
'                                         Default = "Are you sure this is what you want to do ?"
' Description
'
' Returns true if the user selects "Yes" from the MsgBox displayed
'
' ********************************************************************************
Public Function bUserSure(Optional ByVal sPrompt As String = "Are you sure this is what you want to do ?", Optional ByVal sTitle As String = "ARE YOU SURE ?") As Boolean
    bUserSure = (MsgBox(sPrompt, vbYesNo, sTitle) = vbYes)
End Function

' ********************************************************************************
' Name              Functions_ExtendListView
'
' Parameters
'      lvwIn                         (O)  The ListView to set line selection for
'
' Description
'
' Sets a ListView object to select an entire line when clicked (instead of just
' the first column)
'
' ********************************************************************************
Public Sub ExtendListView(hWndListView As Long)
    Dim style As Long

    style = SendMessage(hWndListView, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    style = style Or LVS_EX_FULLROWSELECT
    Call SendMessage(hWndListView, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, style)
End Sub

' ***********************************************************************************
' Synopsis          Returns the Nth Token from siAllTokens delimited by sDelim
'
' Parameters
'
'   siAllTokens                 (I) Required. The string containing all the tokens
'   iToken                      (I) Optional. The index of the token to return
'                                   DEFAULT = 1
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = " "
' Description
'  For the following:
'    siAllTokens         iToken   sDelim  Returns       Notes
'   "William M Rawls"    1       " "     "William"      First word
'   "William M Rawls"    2       " "     "M"            Second word
'   "William M Rawls"    3       " "     "Rawls"        Third word
'   "William M Rawls"    4       " "     vbNullString             No forth word
'   "William M Rawls"    0       " "     vbNullString             Zeroth token is always empty
'   "William M Rawls"   -1       " "     vbNullString             Negative tokesn always empty
'   "William M Rawls"    1       vbNullString      vbNullString             No delimiter ? Token empty
' ***********************************************************************************
Public Function sGetToken(ByVal siAllTokens As String, Optional ByVal iToken As Long = 1, Optional ByVal sDelim As String = " ") As String
    Static iCurTokenLocation As Long ' Character position of the first delimiter string
    Static nDelim As Long            ' Length of the delimiter string
    nDelim = Len(sDelim)

    If iToken < 1 Or nDelim < 1 Then
     ' Negative or zeroth token or empty delimiter strings mean an empty token
       Exit Function
    ElseIf iToken = 1 Then
     ' Quickly extract the first token
       iCurTokenLocation = InStr(siAllTokens, sDelim)
       If iCurTokenLocation > 1 Then
          sGetToken = Left$(siAllTokens, iCurTokenLocation - 1)
       ElseIf iCurTokenLocation = 1 Then
          sGetToken = vbNullString
       Else
          sGetToken = siAllTokens
       End If
       Exit Function
    Else
     ' Find the Nth token
       Do
          iCurTokenLocation = InStr(siAllTokens, sDelim)
          If iCurTokenLocation = 0 Then
             Exit Function
          Else
             siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
          End If
          iToken = iToken - 1
       Loop Until iToken = 1

     ' Extract the Nth token (Which is the next token at this point)
       iCurTokenLocation = InStr(siAllTokens, sDelim)
       If iCurTokenLocation > 0 Then
          sGetToken = Left$(siAllTokens, iCurTokenLocation - 1)
          Exit Function
       Else
          sGetToken = siAllTokens
          Exit Function
       End If
    End If
End Function
' *********************************************************************************************
' Synopsis          Returns everything AFTER the Nth Token from siAllTokens delimited by sDelim
'
' Parameters
'
'   siAllTokens                 (I) Required. The string containing all the tokens
'   iToken                      (I) Optional. The index of the token to use as an "after" ref
'                                   DEFAULT = 1
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = " "
' Description
'  For the following:
'    siAllTokens         iToken   sDelim  Returns           Notes
'   "William M Rawls"    1       " "     "M Rawls"          After the first word
'   "William M Rawls"    2       " "     "Rawls"            After the second word
'   "William M Rawls"    3       " "     vbNullString                 After the third word (nothing)
'   "William M Rawls"    0       " "     "William M Rawls"  After zeroth token is always the input string
'   "William M Rawls"   -1       " "     "William M Rawls"  Negative tokens act same as zero
'   "William M Rawls"    1       vbNullString      "William M Rawls"  Same as one
' *********************************************************************************************
Public Function sAfter(ByVal siAllTokens As String, Optional ByVal iToken As Long = 1, Optional ByVal sDelim As String = " ") As String
    Static iCurTokenLocation As Long ' Character position of the first delimiter string
    Static nDelim As Long            ' Length of the delimiter string
    
    nDelim = Len(sDelim)
    If iToken < 1 Or nDelim < 1 Then
     ' Negative or zeroth token or empty delimiter strings mean an empty token
       sAfter = siAllTokens
       Exit Function
    ElseIf iToken = 1 Then
     ' Quickly extract the first token
       iCurTokenLocation = InStr(siAllTokens, sDelim)
       If iCurTokenLocation > 1 Then
          sAfter = Mid$(siAllTokens, iCurTokenLocation + nDelim)
          Exit Function
       ElseIf iCurTokenLocation = 0 Then
          sAfter = siAllTokens
          Exit Function
       Else
          sAfter = Mid$(siAllTokens, nDelim + 1)
          Exit Function
       End If
    Else
     ' Find the Nth token
       Do
          iCurTokenLocation = InStr(siAllTokens, sDelim)
          If iCurTokenLocation = 0 Then
             Exit Function
          Else
             siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
          End If
          iToken = iToken - 1
       Loop Until iToken = 1

     ' Extract the Nth token (Which is the next token at this point)
       iCurTokenLocation = InStr(siAllTokens, sDelim)
       If iCurTokenLocation > 0 Then
          sAfter = Mid$(siAllTokens, iCurTokenLocation + nDelim)
          Exit Function
       Else
          Exit Function
       End If
    End If
End Function

' **********************************************************************************************
' Synopsis          Returns everything BEFORE the Nth Token from siAllTokens delimited by sDelim
'
' Parameters
'
'   siAllTokens                 (I) Required. The string containing all the tokens
'   iToken                      (I) Optional. The index of the token to use as a "before" ref
'                                   DEFAULT = 2
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = " " (Space)
' Description
'  For the following:
'    siAllTokens         iToken   sDelim  Returns           Notes
'   "William M Rawls"    1       " "     vbNullString                 Before the first word (nothing)
'   "William M Rawls"    2       " "     "William"          Before the second word
'   "William M Rawls"    3       " "     "William M"        Before the third word
'   "William M Rawls"    0       " "     vbNullString                 Before zeroth token (nothing)
'   "William M Rawls"   -1       " "     vbNullString                 Negative tokens act same as zero
'   "William M Rawls"    1       vbNullString      vbNullString                 Same as one
' *********************************************************************************************
Public Function sBefore(ByVal siAllTokens As String, Optional ByVal iToken As Long = 2, Optional ByVal sDelim As String = " ") As String
    Static iCurTokenLocation As Long ' Character position of the first delimiter string
    Static nDelim As Long            ' Length of the delimiter string
    Static sReturned As String

    nDelim = Len(sDelim)
    If iToken < 2 Or nDelim < 1 Then
     ' First, Zeroth, or Negative tokens or empty delimiter strings mean an empty string returned
       sBefore = vbNullString
       Exit Function
    ElseIf iToken = 2 Then
     ' Quickly extract the first token
       sBefore = sGetToken(siAllTokens, 1, sDelim)
       Exit Function
    Else
     ' Find the Nth token
       Do
          iCurTokenLocation = InStr(siAllTokens, sDelim)
          If iCurTokenLocation = 0 Or iToken = 1 Then
             sBefore = sReturned
             sReturned = vbNullString
             Exit Function
          ElseIf Len(sReturned) = 0 Then
             sReturned = Left$(siAllTokens, iCurTokenLocation - 1)
          Else
             sReturned = sReturned & sDelim & Left$(siAllTokens, iCurTokenLocation - 1)
          End If
          siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
          iToken = iToken - 1
       Loop
    End If
End Function
' ---------------------------------------------------------
' Check for the existence of a file by attempting an OPEN.
' ---------------------------------------------------------
Function FileExists(path As String) As Long
    FileExists = (Len(Dir(path)) > 0)
End Function

' ---------------------------
' By: William "Miller" Rawls
' ---------------------------
Function nTokens(ByVal strIn As String, sDelim As String)
    Dim i As Long, nCount As Long
    i = InStr(strIn, sDelim)
    If i < 1 Then
       nTokens = Abs((Len(strIn) > 0))
    Else
       nCount = 2
       Do Until i < 1
          strIn = Mid$(strIn, i + Len(sDelim))
          i = InStr(strIn, sDelim)
          If i > 0 Then nCount = nCount + 1
       Loop
       nTokens = nCount
    End If
End Function


' --------------------------------------------------------
' Calls the windows API to get the windows\SYSTEM directory
' --------------------------------------------------------
Public Function sGetWindowsSysDir() As String
    Dim x As Integer
    Dim sT As String

    sT = String$(145, 0)                 ' Size Buffer
    x = GetSystemDirectory(sT, 145)      ' Make API Call
    sT = Left$(sT, x)                 ' Trim Buffer

    If Right$(sT, 1) <> "\" Then         ' Add \ if necessary
       sGetWindowsSysDir = sT + "\"
    Else
       sGetWindowsSysDir = sT
    End If
End Function


Public Function LogError(ByVal sModuleName As String, sProcName As String, lError As Long, sErrorMsg As String) As Boolean
    Dim fh As Long
    Dim sMessage As String

    fh = FreeFile
    Open "ERRORLOG.TXT" For Append As #fh
         sMessage = "***** Error " & Format(lError, "00000") & " at: " & Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM")
         sMessage = sMessage & vbNewLine & "  *** Module:         " & sModuleName
         sMessage = sMessage & vbNewLine & "  *** Procedure:      " & sProcName
         sMessage = sMessage & vbNewLine & "  *** Description:    " & sErrorMsg
         Print #fh, sMessage
         sMessage = sMessage & vbNewLine & vbNewLine & vbTab & "Continue after error ? (No to exit program)"
         If MsgBox(sMessage, vbYesNo) = vbNo Then
            Print #fh, "  *** Program shut down by user after error."
            ShutDownNicely
         Else
            Print #fh, "  *** Program continued by user after error."
         End If
    Close #fh

End Function

Public Sub ShutDownNicely()
  ' Close all objects, forms, handles, etc. here
    End
End Sub
' ***********************************************************************************
' Synopsis          Returns the number of tokens as delimited by siDelim
'
' Parameters
'
'   sAllTokens                 (I) Required. The string containing all the tokens
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = " "
' Description
'  For the following:
'    sAllTokens         sDelim  Returns       Notes
'   "William M Rawls"    " "     3             "William", "M", and "Rawls"
'   "William M Rawls"    "iam"   2             "Will" and " M Rawls"
'   "William M Rawls"    vbNullString      1             No delimiter? String has one token,
'                                              "William M Rawls"
'   "1.00.05"            "."     3             "1", "00", and "05"
' ***********************************************************************************
Public Function lTokenCount(ByVal sAllTokens As String, Optional ByVal siDelim As String = " ") As Long
    Static iCurTokenLocation As Long ' Character position of the first delimiter string
    Static iTokensSoFar As Long      ' Used to keep track of how many tokens we've counted so far
    Static iDelim As Long            ' Length of the delimiter string

    iDelim = Len(siDelim)
    If iDelim < 1 Then
     ' Empty delimiter strings means only one token equal to the string
       lTokenCount = 1
       Exit Function
    ElseIf Len(sAllTokens) = 0 Then
     ' Empty input string means no tokens
       Exit Function
    Else
     ' Count the number of tokens
       iTokensSoFar = 0
       Do
          iCurTokenLocation = InStr(sAllTokens, siDelim)
          If iCurTokenLocation = 0 Then
             lTokenCount = iTokensSoFar + 1 'Abs(Len(sAllTokens) > 0)
             Exit Function
          End If
          iTokensSoFar = iTokensSoFar + 1
          sAllTokens = Mid$(sAllTokens, iCurTokenLocation + iDelim)
       Loop
    End If
End Function
Public Function sDenormalize(sLine As String) As String
    sDenormalize = Replace$(Replace$(sLine, "%$%EOL%$%", vbNewLine), "%$%TAB%$%", vbTab)
End Function
Public Function sNormalize(sLine As String) As String
    sNormalize = Replace$(Replace$(sLine, vbNewLine, "%$%EOL%$%"), vbTab, "%$%TAB%$%")
End Function
Public Function LoadFormPosition(frmToActOn As Form, Optional ByVal bAutoCenter = True)
    Dim ProductName As String
    
    If Len(App.ProductName) = 0 Then
       ProductName = "Your Product Name Here"
    Else
       ProductName = App.ProductName
    End If
    
    If GetSetting(ProductName, frmToActOn.Name, "Position Saved", False) Then
       frmToActOn.Left = GetSetting(ProductName, frmToActOn.Name, "Form Position Left", frmToActOn.Left)
       frmToActOn.Top = GetSetting(ProductName, frmToActOn.Name, "Form Position Top", frmToActOn.Top)
       frmToActOn.Width = GetSetting(ProductName, frmToActOn.Name, "Form Position Width", frmToActOn.Width)
       frmToActOn.Height = GetSetting(ProductName, frmToActOn.Name, "Form Position Height", frmToActOn.Height)
    ElseIf bAutoCenter Then
       frmToActOn.Left = (Screen.Width - frmToActOn.Width) / 2
       frmToActOn.Top = (Screen.Height - frmToActOn.Height) / 2
    End If
End Function

Public Function SaveFormPosition(frmToActOn As Form)
    Dim ProductName As String
    
    If frmToActOn.WindowState <> vbNormal Then Exit Function

    If Len(App.ProductName) = 0 Then
       ProductName = "Your Product Name Here"
    Else
       ProductName = App.ProductName
    End If
    
    SaveSetting ProductName, frmToActOn.Name, "Position Saved", True
    SaveSetting ProductName, frmToActOn.Name, "Form Position Left", frmToActOn.Left
    SaveSetting ProductName, frmToActOn.Name, "Form Position Top", frmToActOn.Top
    SaveSetting ProductName, frmToActOn.Name, "Form Position Width", frmToActOn.Width
    SaveSetting ProductName, frmToActOn.Name, "Form Position Height", frmToActOn.Height
End Function

