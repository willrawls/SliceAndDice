Attribute VB_Name = "WMR"
'Option Explicit

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
    reserved As String * 4
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
          sPath = Left(sPath, iNull - 1)
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

' ********************************************************************************
' Name              Functions_GetFiles
'
' Parameters
'      sSpec                         (I)  The path/directory to retrieve
'      sFiles()                      (O)  Array of strings to contain the list of
'                                         files in the directory
' Description
'
' Returns an array of strings containing the names of the files in the directory
' supplied
'
' ********************************************************************************
Public Function GetFiles(ByVal sSpec As String, sFiles() As String) As Long
    Dim nFiles As Long
    Dim sFile As String
    ReDim sFiles(1 To 10) As String

    sFile = Dir(sSpec)
    Do While Len(sFile)
       nFiles = nFiles + 1
       sFiles(nFiles) = sFile
       If nFiles = UBound(sFiles) Then
          ReDim Preserve sFiles(1 To UBound(sFiles) + 10) As String
        End If
        sFile = Dir()
    Loop

    If nFiles Then
       ReDim Preserve sFiles(1 To nFiles) As String
    End If

    GetFiles = nFiles

End Function

' ********************************************************************************
' Name              Functions_KeepFormOnTop
'
' Parameters
'      hWnd                          (I)  Window handle of form to keep on top
'
' Description
'
' Keeps a form on top of all others.
'
' ********************************************************************************
Public Sub KeepFormOnTop(ByVal hWnd As Long)
On Error GoTo EH_KeepFormOnTop
       SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPFLAGS
       Exit Sub

EH_KeepFormOnTop:
    Err.Raise vbObjectError + 0, "Functions_KeepFormOnTop"
End Sub

' ********************************************************************************
' Name              Functions_NextNegativeUnique
'
' Parameters
'      None
'
' Description
'
' Used to return a unique negative number. Numbers are unique to the current
' program session only.
'
' ********************************************************************************
Public Function NextNegativeUnique() As Long
    Static lNextSerial As Long
    lNextSerial = lNextSerial - 1
    NextNegativeUnique = lNextSerial
End Function

' ***********************************************************************************
' Synopsis          Returns the number of tokens as delimited by siDelim
'
' Parameters
'
'   siAllTokens                 (I) Required. The string containing all the tokens
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = " "
' Description
'  For the following:
'    siAllTokens         sDelim  Returns       Notes
'   "William M Rawls"    " "     3             "William", "M", and "Rawls"
'   "William M Rawls"    "iam"   2             "Will" and " M Rawls"
'   "William M Rawls"    ""      1             No delimiter? String has one token,
'                                              "William M Rawls"
'   "1.00.05"            "."     3             "1", "00", and "05"
' ***********************************************************************************
Public Function iTokenCount(ByVal siAllTokens As String, Optional ByVal siDelim As String = " ") As Long
    Static iCurTokenLocation As Long ' Character position of the first delimiter string
    Static iTokensSoFar As Long      ' Used to keep track of how many tokens we've counted so far
    Static iDelim As Long            ' Length of the delimiter string

    iDelim = Len(siDelim)
    If iDelim < 1 Then
     ' Empty delimiter strings means only one token equal to the string
       iTokenCount = 1
       Exit Function
    ElseIf Len(siAllTokens) = 0 Then
     ' Empty input string means no tokens
       Exit Function
    Else
     ' Count the number of tokens
       iTokensSoFar = 0
       Do
          iCurTokenLocation = InStr(siAllTokens, siDelim)
          If iCurTokenLocation = 0 Then
             iTokenCount = iTokensSoFar + 1 'Abs(Len(siAllTokens) > 0)
             Exit Function
          End If
          iTokensSoFar = iTokensSoFar + 1
          siAllTokens = Mid(siAllTokens, iCurTokenLocation + iDelim)
       Loop
    End If
End Function
' ********************************************************************************
' Name              Functions_nz
'
' Parameters
'      vData                         (O)  Variant to test for NULL
'       sDefault                     (O)  Opt. On NULL this string is returned
'                                         Default = ""
' Description
'
' Returns sDefault if the variant is NULL, otherwise it returns the Variant
'
' ********************************************************************************
Public Function nz(vData As Variant, Optional sDefault As String = "") As String
    If IsNull(vData) Then nz = sDefault Else nz = vData
End Function

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
'   "William M Rawls"    4       " "     ""             No forth word
'   "William M Rawls"    0       " "     ""             Zeroth token is always empty
'   "William M Rawls"   -1       " "     ""             Negative tokesn always empty
'   "William M Rawls"    1       ""      ""             No delimiter ? Token empty
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
          sGetToken = Left(siAllTokens, iCurTokenLocation - 1)
       ElseIf iCurTokenLocation = 1 Then
          sGetToken = ""
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
             siAllTokens = Mid(siAllTokens, iCurTokenLocation + nDelim)
          End If
          iToken = iToken - 1
       Loop Until iToken = 1

     ' Extract the Nth token (Which is the next token at this point)
       iCurTokenLocation = InStr(siAllTokens, sDelim)
       If iCurTokenLocation > 0 Then
          sGetToken = Left(siAllTokens, iCurTokenLocation - 1)
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
'   "William M Rawls"    3       " "     ""                 After the third word (nothing)
'   "William M Rawls"    0       " "     "William M Rawls"  After zeroth token is always the input string
'   "William M Rawls"   -1       " "     "William M Rawls"  Negative tokens act same as zero
'   "William M Rawls"    1       ""      "William M Rawls"  Same as one
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
          sAfter = Mid(siAllTokens, iCurTokenLocation + nDelim)
          Exit Function
       ElseIf iCurTokenLocation = 0 Then
          sAfter = siAllTokens
          Exit Function
       Else
          sAfter = Mid(siAllTokens, nDelim + 1)
          Exit Function
       End If
    Else
     ' Find the Nth token
       Do
          iCurTokenLocation = InStr(siAllTokens, sDelim)
          If iCurTokenLocation = 0 Then
             Exit Function
          Else
             siAllTokens = Mid(siAllTokens, iCurTokenLocation + nDelim)
          End If
          iToken = iToken - 1
       Loop Until iToken = 1

     ' Extract the Nth token (Which is the next token at this point)
       iCurTokenLocation = InStr(siAllTokens, sDelim)
       If iCurTokenLocation > 0 Then
          sAfter = Mid(siAllTokens, iCurTokenLocation + nDelim)
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
'   "William M Rawls"    1       " "     ""                 Before the first word (nothing)
'   "William M Rawls"    2       " "     "William"          Before the second word
'   "William M Rawls"    3       " "     "William M"        Before the third word
'   "William M Rawls"    0       " "     ""                 Before zeroth token (nothing)
'   "William M Rawls"   -1       " "     ""                 Negative tokens act same as zero
'   "William M Rawls"    1       ""      ""                 Same as one
' *********************************************************************************************
Public Function sBefore(ByVal siAllTokens As String, Optional ByVal iToken As Long = 2, Optional ByVal sDelim As String = " ") As String
    Static iCurTokenLocation As Long ' Character position of the first delimiter string
    Static nDelim As Long            ' Length of the delimiter string
    Static sReturned As String

    nDelim = Len(sDelim)
    If iToken < 2 Or nDelim < 1 Then
     ' First, Zeroth, or Negative tokens or empty delimiter strings mean an empty string returned
       sBefore = ""
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
             sReturned = ""
             Exit Function
          ElseIf Len(sReturned) = 0 Then
             sReturned = Left(siAllTokens, iCurTokenLocation - 1)
          Else
             sReturned = sReturned & sDelim & Left(siAllTokens, iCurTokenLocation - 1)
          End If
          siAllTokens = Mid(siAllTokens, iCurTokenLocation + nDelim)
          iToken = iToken - 1
       Loop
    End If
End Function
' **********************************************************************************************
' Synopsis          Returns everything EXCEPT the Nth Token from siAllTokens delimited by sDelim
'
' Parameters
'
'   siAllTokens                 (I) Required. The string containing all the tokens
'   iToken                      (I) Optional. The index of the token to exclude
'                                   DEFAULT = 1
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = " " (Space)
' Description
'  For the following:
'    siAllTokens         iToken   sDelim  Returns           Notes
'   "William M Rawls"    1       " "     "M Rawls"          After 1st token
'   "William M Rawls"    2       " "     "William Rawls"    1st and 3rd token
'   "William M Rawls"    3       " "     "William M"        Before the third word
'   "William M Rawls"    0       " "     "William M Rawls"  Everything except 0th token (everything)
'   "William M Rawls"   -1       " "     ""                 Negative tokens act same as zero
'   "William M Rawls"    1       ""      "William M Rawls"  Same as zero
' *********************************************************************************************
Public Function sExcept(ByVal siAllTokens As String, Optional ByVal iToken As Long = 1, Optional ByVal sDelim As String = " ") As String
    Static iCurTokenLocation As Long ' Character position of the first delimiter string
    Static nDelim As Long            ' Length of the delimiter string
    Static sReturned As String

    nDelim = Len(sDelim)
    If iToken < 1 Or nDelim < 1 Then
     ' First, Zeroth, or Negative tokens or empty delimiter strings mean an empty string returned
       sExcept = siAllTokens
       Exit Function
    ElseIf iToken = 1 Then
     ' Quickly Return after token 1
       iCurTokenLocation = InStr(siAllTokens, sDelim)
       If iCurTokenLocation = 0 Then
          sExcept = siAllTokens
          Exit Function
       Else
          sExcept = Mid(siAllTokens, iCurTokenLocation + nDelim)
          Exit Function
       End If
    Else
     ' Find the Nth token
       Do
          iCurTokenLocation = InStr(siAllTokens, sDelim)
          If iToken = 1 Then
             If iCurTokenLocation > 0 Then
                sExcept = sReturned & sDelim & Mid(siAllTokens, iCurTokenLocation + nDelim)
             Else
                sExcept = sReturned
             End If
             sReturned = ""
             Exit Function
          ElseIf iCurTokenLocation = 0 Then
             sExcept = sReturned & sDelim & siAllTokens
             sReturned = ""
             Exit Function
          ElseIf Len(sReturned) = 0 Then
             sReturned = Left(siAllTokens, iCurTokenLocation - 1)
          Else
             sReturned = sReturned & sDelim & Left(siAllTokens, iCurTokenLocation - 1)
          End If
          siAllTokens = Mid(siAllTokens, iCurTokenLocation + nDelim)
          iToken = iToken - 1
       Loop
    End If
End Function
' ********************************************************************************
' Name              Functions_sReplace
'
' Parameters
'      sAll                          (I)  String
'      sFind                         (I)  String
'      sReplaceWith                  (I)  String
'
' Description
'
' Replaces all occurances of one string with another.
' NOTE: This function does NOT test for re-entrant replacements and could cause
'       an infinite loop.
' ********************************************************************************
Public Function sReplace(ByVal sAll As String, ByVal sFind As String, ByVal sReplaceWith As String) As String
    Dim iCurFindPos As Long
    Dim iFind As Long

    iFind = Len(sFind)
    iCurFindPos = InStr(sAll, sFind)
    Do While iCurFindPos > 0
       If iCurFindPos > 1 Then
          sAll = Left(sAll, iCurFindPos - 1) & sReplaceWith & Mid(sAll, iCurFindPos + iFind)
       Else
          sAll = sReplaceWith & Mid(sAll, iCurFindPos + iFind)
       End If
       iCurFindPos = InStr(sAll, sFind)
    Loop
    sReplace = sAll
End Function

' ********************************************************************************
' Name              Functions_zn
'
' Parameters
'      sData                         (I)  String to test for empty
'
' Description
'
' If the Variant passed in is an empty string, this procedure returns NULL,
' otherwise it returns the string passed.
'
' ********************************************************************************
Public Function zn(sData As String) As Variant
    If Len(sData) = 0 Then zn = Null Else zn = sData
End Function


' ------------------------------------------------------
' Centers the passed form just above center on the screen
' ------------------------------------------------------
Sub CenterForm(x As Form)
  
    Screen.MousePointer = 11
    x.Top = (Screen.Height * 0.85) / 2 - x.Height / 2
    x.Left = Screen.Width / 2 - x.Width / 2
    Screen.MousePointer = 0

End Sub

' --------------------------------------------
' Create the path contained in DestPath
' First char must be drive letter, followed by
' a ":\" followed by the path, if any.
' --------------------------------------------
Function CreatePath(ByVal destPath As String) As Long
    Screen.MousePointer = 11

    ' --------------------------------------------
    ' Add slash to end of path if not there already
    ' --------------------------------------------
    If Right(destPath, 1) <> "\" Then
        destPath = destPath + "\"
    End If
          

    ' ----------------------------------
    ' Change to the root dir of the drive
    ' ----------------------------------
    On Error Resume Next
    ChDrive destPath
    If Err <> 0 Then GoTo errorOut
    ChDir "\"

    ' ------------------------------------------------
    ' Attempt to make each directory, then change to it
    ' ------------------------------------------------
    BackPos = 3
    forePos = InStr(4, destPath, "\")
    Do While forePos <> 0
        temp = Mid(destPath, BackPos + 1, forePos - BackPos - 1)

        Err = 0
        MkDir temp
        If Err <> 0 And Err <> 75 Then GoTo errorOut

        Err = 0
        ChDir temp
        If Err <> 0 Then GoTo errorOut

        BackPos = forePos
        forePos = InStr(BackPos + 1, destPath, "\")
    Loop
                 
    CreatePath = True
    Screen.MousePointer = 0
    Exit Function
                 
errorOut:
    MsgBox "Error While Attempting to Create Directories on Destination Drive.", 48, "SETUP"
    CreatePath = False
    Screen.MousePointer = 0

End Function

' ---------------------------------------------------------
' Check for the existence of a file by attempting an OPEN.
' ---------------------------------------------------------
Function FileExists(path As String) As Long

    x = FreeFile

    On Error Resume Next
    Open path For Input As x
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close x

End Function

' -----------------------
' Get the size of the file
' -----------------------
Function GetFileSize(source As String) As Long
    x = FreeFile
    Open source For Binary Access Read As x
    GetFileSize = LOF(x)
    Close x
End Function


' -----------------------------------------------------
' Function:   IsValidPath as Long
' arguments:  DestPath         a string that is a full path
'             DefaultDrive     the default drive.  eg.  "C:"
'
'  If DestPath does not include a drive specification,
'  IsValidPath uses Default Drive
'
'  When IsValidPath is finished, DestPath is reformated
'  to the format "X:\dir\dir\dir\"
'
' Result:  True (-1) if path is valid.
'          False (0) if path is invalid
' ------------------------------------------------------
Function IsValidPath(destPath As String, ByVal DefaultDrive As String) As Long

    ' ---------------------------
    ' Remove left and right spaces
    ' ---------------------------
    destPath = RTrim(LTrim(destPath))
    

    ' ----------------------------
    ' Check Default Drive Parameter
    ' ----------------------------
    If Right(DefaultDrive, 1) <> ":" Or Len(DefaultDrive) <> 2 Then
        MsgBox "Bad default drive parameter specified in IsValidPath Function.  You passed,  """ + DefaultDrive + """.  Must be one drive letter and "":"".  For example, ""C:"", ""D:""...", 64, "Setup Kit Error"
        GoTo parseErr
    End If
    

    ' ------------------------------------------------------
    ' Insert default drive if path begins with root backslash
    ' ------------------------------------------------------
    If Left(destPath, 1) = "\" Then
        destPath = DefaultDrive + destPath
    End If
    
    ' ----------------------------
    ' check for invalid characters
    ' ----------------------------
    On Error Resume Next
    tmp = Dir(destPath)
    If Err <> 0 Then
        GoTo parseErr
    End If
    

    ' ----------------------------------------
    ' Check for wildcard characters and spaces
    ' ----------------------------------------
    If (InStr(destPath, "*") <> 0) Then GoTo parseErr
    If (InStr(destPath, "?") <> 0) Then GoTo parseErr
    If (InStr(destPath, " ") <> 0) Then GoTo parseErr
         
    
    ' -----------------------------------------
    ' Make Sure colon is in second char position
    ' -----------------------------------------
    If Mid(destPath, 2, 1) <> Chr(58) Then GoTo parseErr
    

    ' ------------------------------
    ' Insert root backslash if needed
    ' ------------------------------
    If Len(destPath) > 2 Then
      If Right(Left(destPath, 3), 1) <> "\" Then
        destPath = Left(destPath, 2) + "\" + Right(destPath, Len(destPath) - 2)
      End If
    End If

    ' ------------------------
    ' Check drive to install on
    ' ------------------------
    Drive = Left(destPath, 1)
    ChDrive (Drive)                                                        ' Try to change to the dest drive
    If Err <> 0 Then GoTo parseErr
    
    ' ----------
    ' Add final \
    ' ----------
    If Right(destPath, 1) <> "\" Then
        destPath = destPath + "\"
    End If
    

    ' ------------------------------------
    ' Root dir is a valid dir
    ' ------------------------------------
    If Len(destPath) = 3 Then
        If Right(destPath, 2) = ":\" Then
            GoTo ParseOK
        End If
    End If

    ' -----------------------
    ' Check for repeated Slash
    ' -----------------------
    If InStr(destPath, "\\") <> 0 Then GoTo parseErr
        
    ' -------------------------------------
    ' Check for illegal directory names
    ' -------------------------------------
    legalChar = "!#$%&'()-0123456789@ABCDEFGHIJKLMNOPQRSTUVWXYZ^_`{}~.üäöÄÖÜß"
    BackPos = 3
    forePos = InStr(4, destPath, "\")
    Do
        temp = Mid(destPath, BackPos + 1, forePos - BackPos - 1)
        
        ' ---------------------------
        ' Test for illegal characters
        ' ---------------------------
        For i = 1 To Len(temp)
            If InStr(legalChar, UCase(Mid(temp, i, 1))) = 0 Then GoTo parseErr
        Next i

        ' ------------------------------------------
        ' Check combinations of periods and lengths
        ' ------------------------------------------
        periodPos = InStr(temp, ".")
        length = Len(temp)
        If periodPos = 0 Then
            If length > 8 Then GoTo parseErr                         ' Base too long
        Else
            If periodPos > 9 Then GoTo parseErr                      ' Base too long
            If length > periodPos + 3 Then GoTo parseErr             ' Extension too long
            If InStr(periodPos + 1, temp, ".") <> 0 Then GoTo parseErr ' Two periods not allowed
        End If

        BackPos = forePos
        forePos = InStr(BackPos + 1, destPath, "\")
    Loop Until forePos = 0

ParseOK:
    IsValidPath = True
    Exit Function

parseErr:
    IsValidPath = False
End Function

Sub ModifySystemIni(sWinDir As String, ByVal sSection As String, sToAdd As String)
    Dim fileIn As Long, fileOut As Long, sStr As String

    sSection = "[" & sSection & "]"

    On Error GoTo EH_ModifySystemIni_Copy
    
    fileIn = FreeFile
    Open sWinDir & "SYSTEM.INI" For Input Access Read As #fileIn
         fileOut = FreeFile
         Open sWinDir & "SYSTEMP.XXX" For Output Access Write As #fileOut
              Do Until EOF(fileIn)
                 Line Input #fileIn, sStr
                 Print #fileOut, sStr
                 If sStr = sSection Then Print #fileOut, sToAdd
              Loop
         Close #fileOut
    Close #fileIn

    If FileExists(sWinDir & "SYSTEM.OLD") Then Kill sWinDir & "SYSTEM.OLD"
    Name sWinDir & "SYSTEM.INI" As sWinDir & "SYSTEM.OLD"
    Name sWinDir & "SYSTEMP.XXX" As sWinDir & "SYSTEM.INI"
Exit Sub

EH_ModifySystemIni_Copy:
    MsgBox "Error modifying SYSTEM.INI file, exiting."
    End
End Sub

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
          strIn = Mid(strIn, i + Len(sDelim))
          i = InStr(strIn, sDelim)
          If i > 0 Then nCount = nCount + 1
       Loop
       nTokens = nCount
    End If
End Function

