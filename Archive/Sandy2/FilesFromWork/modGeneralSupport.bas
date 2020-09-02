Attribute VB_Name = "modGeneral"
Option Explicit

' ********************************************************************************
' Class Module      modGeneral
'
' Filename          modGeneral.cls
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
  Public gbProcessing As Boolean
  
' True if the user cancel processing while doing an insertion
 'Public gbCancelInsertion As Boolean

'
  Public gbEvaluationHasExpired As Boolean

' ***************************************************
' Publicly available constant strings
'   Call InitPublic() to set at beginning of program
' Why ? These strings are very common in VB
'   and using the Publicly available
' ***************************************************
 'Public vbNewLine As String            ' vbCrLf
 'Public vbTab As String            ' vbTab
  Public gsEolTab As String         ' vbNewLine & vbTab
  Public gs2EOL As String           ' vbNewLine & vbNewLine
  Public gs2EOLTab As String        ' gs2EOL & vbTab

 'Public Const gsB As String = vbNullString
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

Private Const msHaxValues As String = "000102030405060708090A0B0C0D0E0F101112131415161718191A1B1C1D1E1F202122232425262728292A2B2C2D2E2F303132333435363738393A3B3C3D3E3F404142434445464748494A4B4C4D4E4F505152535455565758595A5B5C5D5E5F606162636465666768696A6B6C6D6E6F707172737475767778797A7B7C7D7E7F808182838485868788898A8B8C8D8E8F909192939495969798999A9B9C9D9E9FA0A1A2A3A4A5A6A7A8A9AAABACADAEAFB0B1B2B3B4B5B6B7B8B9BABBBCBDBEBFC0C1C2C3C4C5C6C7C8C9CACBCCCDCECFD0D1D2D3D4D5D6D7D8D9DADBDCDDDEDFE0E1E2E3E4E5E6E7E8E9EAEBECEDEEEFF0F1F2F3F4F5F6F7F8F9FAFBFCFDFEFF"

' ***********************************
' ****** BrowseForFolder stuff ******
' ***********************************
  Public Type BrowseInfo
          hWndOwner As Long
          pIDLRoot As Long
          pszDisplayName As String
          lpszTitle As String
          ulFlags As Long
          lpfnCallback As Long
          lParam As Long
          iImage As Long
  End Type
  
  Public Const BIF_RETURNONLYFSDIRS = 1
  Public Const MAX_PATH = 260
  Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
  Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
  Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

' **********************************************************
' API call to determin where the user's Windows directory is
' **********************************************************
  Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function FindListIndex(ctrlToSearch As Object, sToFind As String) As Long
On Error Resume Next
    Dim i As Long

    If Len(sToFind) > 0 Then
       For i = 0 To ctrlToSearch.ListCount - 1
           If StrComp(ctrlToSearch.List(i), sToFind, vbTextCompare) = 0 Then
              FindListIndex = i
              Exit Function
           End If
       Next i
    End If
    FindListIndex = -1
End Function

Public Function SaveFormPosition(frmToActOn As Object, Optional ByVal sSectionName As String, Optional ByVal ProductName As String = "SliceAndDice")
    Dim SectionName As String
    
    If Len(ProductName) = 0 Then ProductName = "SliceAndDice"

    With frmToActOn
         If Len(sSectionName) Then
            SectionName = sSectionName
         Else
            SectionName = .Name
         End If
         SaveSetting ProductName, SectionName, "Position Saved", True
         SaveSetting ProductName, SectionName, "Form Position Left", .Left
         SaveSetting ProductName, SectionName, "Form Position Top", .Top
         SaveSetting ProductName, SectionName, "Form Position Width", .Width
         SaveSetting ProductName, SectionName, "Form Position Height", .Height
    End With
End Function


Public Function LoadFormPosition(frmToActOn As Object, Optional ByVal bAutoCenter = True, Optional ByVal bRemeberWidth As Boolean = True, Optional ByVal sSectionName As String, Optional ByVal ProductName As String = "SliceAndDice")
    Dim SectionName As String
    
    If Len(ProductName) = 0 Then
       ProductName = "Slice and Dice"
    End If
    
    With frmToActOn
         If Len(sSectionName) Then
            SectionName = sSectionName
         Else
            SectionName = .Name
         End If
         If GetSetting(ProductName, SectionName, "Position Saved", False) Then
            .Left = GetSetting(ProductName, SectionName, "Form Position Left", .Left)
            .Top = GetSetting(ProductName, SectionName, "Form Position Top", .Top)
            If bRemeberWidth Then .Width = GetSetting(ProductName, SectionName, "Form Position Width", .Width)
            .Height = GetSetting(ProductName, SectionName, "Form Position Height", .Height)
         ElseIf bAutoCenter Then
            .Left = (Screen.Width - .Width) / 2
            .Top = (Screen.Height - .Height) / 2
         End If

       ' Ensure it'll fit on the screen (screen resolution change ?)
         If .Left > Screen.Width Then .Left = 0
         If .Top > Screen.Height Then .Top = 0
         If .Left + .Width > Screen.Width Then .Width = Screen.Width - .Left
         If bRemeberWidth Then
            If .Top + .Height > Screen.Height Then .Height = Screen.Height - .Top
         End If
    End With
End Function

 

Public Function CreateSandyDatabase(ByVal hWnd As Long) As String
    Dim sDatabasePath    As String
    Dim sNewDatabaseName As String

    Dim db               As Database
    Dim tblTemplates     As TableDef
    Dim fldTemplates     As Field
    Dim ndxTemplates     As Index
    Dim rstCategory      As Recordset

    sDatabasePath = Trim$(BrowseForFolder(hWnd, "Where should database go ?"))
    If Len(sDatabasePath) = 0 Then Exit Function

    sNewDatabaseName = Trim$(InputBox("What should the name of the new Template database be ?", "CREATE TEMPLATE DATABASE"))
    If Len(sNewDatabaseName) = 0 Then Exit Function

    If Right$(sDatabasePath, 1) <> "\" Then sDatabasePath = sDatabasePath & "\"
    If Right$(LCase$(sNewDatabaseName), 4) <> ".mdb" Then sNewDatabaseName = sDatabasePath & sNewDatabaseName & ".mdb"

On Error GoTo mnuSpecialNewDatabase_Click
    Err.Clear
    Set db = CreateDatabase(sNewDatabaseName, dbLangGeneral, dbVersion30)
        If Err.Number <> 0 Then
           MsgBox "Error creating template database. Aborting."
           Exit Function
        End If

        Set tblTemplates = db.CreateTableDef("Category")
        With tblTemplates
             Set fldTemplates = .CreateField("CategoryID", dbLong)
             fldTemplates.Attributes = dbAutoIncrField
             .Fields.Append fldTemplates
            .Fields.Append .CreateField("CategoryName", dbText, 255)
            .Fields.Append .CreateField("CategoryType", dbLong)
            .Fields.Append .CreateField("ColumnWidth", dbSingle)
            .Fields.Append .CreateField("View", dbInteger)
            .Fields.Append .CreateField("Arrange", dbInteger)
            .Fields.Append .CreateField("DateCreated", dbDate)
            .Fields.Append .CreateField("DateModified", dbDate)
            .Fields.Append .CreateField("memoAttributes", dbMemo)

            Set ndxTemplates = .CreateIndex("PrimaryKey")
            With ndxTemplates
                 .Fields.Append .CreateField("CategoryID")
                 .Primary = True
                 .Unique = True
                 .Required = True
            End With
            .Indexes.Append ndxTemplates

            Set ndxTemplates = .CreateIndex("CategoryName")
            With ndxTemplates
                 .Fields.Append .CreateField("CategoryName")
                 .Primary = False
                 .Unique = True
                 .Required = True
            End With
            .Indexes.Append ndxTemplates
            
            Set ndxTemplates = Nothing
            
            db.TableDefs.Append tblTemplates
        End With
        
        Set tblTemplates = db.CreateTableDef("Template")
        With tblTemplates
             Set fldTemplates = .CreateField("TemplateID", dbLong)
             fldTemplates.Attributes = dbAutoIncrField
             .Fields.Append fldTemplates
             .Fields.Append .CreateField("CategoryID", dbLong)
             .Fields.Append .CreateField("TemplateName", dbText, 255)
             .Fields.Append .CreateField("ShortTemplateName", dbText, 255)
             .Fields.Append .CreateField("Filename", dbText, 255)
             .Fields.Append .CreateField("Undeletable", dbBoolean)
             .Fields.Append .CreateField("Locked", dbBoolean)
             .Fields.Append .CreateField("IncludeInMenu", dbBoolean)
             .Fields.Append .CreateField("memoCodeAtCursor", dbMemo)
             .Fields.Append .CreateField("memoCodeAtTop", dbMemo)
             .Fields.Append .CreateField("memoCodeAtBottom", dbMemo)
             .Fields.Append .CreateField("memoCodeToFile", dbMemo)
             .Fields.Append .CreateField("DateCreated", dbDate)
             .Fields.Append .CreateField("DateModified", dbDate)
             .Fields.Append .CreateField("memoAttributes", dbMemo)
             .Fields.Append .CreateField("Favorite", dbBoolean)
             .Fields.Append .CreateField("RevisionCount", dbLong)
             .Fields.Append .CreateField("TimerInsertion", dbText, 255)

            Set ndxTemplates = .CreateIndex("PrimaryKey")
            With ndxTemplates
                 .Fields.Append .CreateField("TemplateID")
                 .Primary = True
                 .Unique = True
                 .Required = True
            End With
            .Indexes.Append ndxTemplates

            Set ndxTemplates = .CreateIndex("CategoryID")
            With ndxTemplates
                 .Fields.Append .CreateField("CategoryID")
                 .Primary = False
                 .Unique = False
                 .Required = True
            End With
            .Indexes.Append ndxTemplates

            Set ndxTemplates = .CreateIndex("ShortTemplateName")
            With ndxTemplates
                 .Fields.Append .CreateField("ShortTemplateName")
                 .Primary = False
                 .Unique = False
                 .Required = True
            End With
            .Indexes.Append ndxTemplates

            Set ndxTemplates = .CreateIndex("TemplateName")
            With ndxTemplates
                 .Fields.Append .CreateField("TemplateName")
                 .Primary = False
                 .Unique = False
                 .Required = True
            End With
            .Indexes.Append ndxTemplates
            Set ndxTemplates = Nothing
            
            db.TableDefs.Append tblTemplates
        End With
    
        Set tblTemplates = db.CreateTableDef("SystemInfo")
        With tblTemplates
             Set fldTemplates = .CreateField("SystemInfoID", dbLong)
             fldTemplates.Attributes = dbAutoIncrField
             .Fields.Append fldTemplates
            .Fields.Append .CreateField("SystemInfoName", dbText, 255)
            .Fields.Append .CreateField("DateCreated", dbDate)
            .Fields.Append .CreateField("DateModified", dbDate)
            .Fields.Append .CreateField("memoAttributes", dbMemo)

            Set ndxTemplates = .CreateIndex("PrimaryKey")
            With ndxTemplates
                 .Fields.Append .CreateField("SystemInfoID")
                 .Primary = True
                 .Unique = True
                 .Required = True
            End With
            .Indexes.Append ndxTemplates

            Set ndxTemplates = .CreateIndex("SystemInfoName")
            With ndxTemplates
                 .Fields.Append .CreateField("SystemInfoName")
                 .Primary = False
                 .Unique = True
                 .Required = True
            End With
            .Indexes.Append ndxTemplates
            
            Set ndxTemplates = Nothing
            
            db.TableDefs.Append tblTemplates
        End With

    Set rstCategory = db.OpenRecordset("Category")
    With rstCategory
         .AddNew
             !CategoryName = "Basic"
             !CategoryType = 0
             !View = 3
         .Update
         .AddNew
             !CategoryName = "Change from"
             !CategoryType = 0
             !View = 3
         .Update
         .AddNew
             !CategoryName = "From the Internet"
             !CategoryType = 0
             !View = 3
         .Update
    End With
    
    db.Close

    CreateSandyDatabase = sNewDatabaseName

mnuSpecialNewDatabase_Click_Continue:
    Exit Function
    
mnuSpecialNewDatabase_Click:
    LogError "SandySupport", "CreateSandyDatabase", Err.Number, Err.Description
    Resume mnuSpecialNewDatabase_Click_Continue:

    Resume
End Function

' -------------------------------------------------
' Calls the windows API to get the windows directory
' -------------------------------------------------
Public Function sGetWindowsDir$()
    Dim X As Integer
    Dim sT As String

    sT = String$(145, 0)              ' Size Buffer
    X = GetWindowsDirectory(sT, 145)  ' Make API Call
    sT = Left$(sT, X)                 ' Trim Buffer

    If Right$(sT, 1) <> "\" Then      ' Add \ if necessary
       sGetWindowsDir = sT + "\"
    Else
       sGetWindowsDir = sT
    End If
End Function
Public Function FileExists(sFilename As String) As Boolean
On Error Resume Next
    Err.Clear
       FileExists = Len(Dir$(sFilename)) > 0
    Err.Clear
End Function

Public Function GetListIndex(cboToSearch As Object, ByVal sItemToFind As String) As Integer
On Error Resume Next
    Static nCurItem As Integer

    If cboToSearch Is Nothing Then Exit Function

    If Len(sItemToFind) = 0 Or cboToSearch.ListCount = 0 Then
       GetListIndex = -1
       Exit Function
    End If

    sItemToFind = UCase$(sItemToFind)

    For nCurItem = 0 To cboToSearch.ListCount - 1
        If StrComp(UCase$(cboToSearch.List(nCurItem)), sItemToFind) = 0 Then
           GetListIndex = nCurItem
           Exit Function
        End If
    Next nCurItem
End Function

Public Function EnumFiles(ByVal sPath As String, Optional ByVal sMask As String = "SAD*.dll", Optional ByVal Attr As VbFileAttribute = vbNormal) As String
    Dim CurrFile As String
    Dim sFileList As String
    
    If Len(sPath) = 0 Then sPath = CurDir
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    CurrFile = Dir$(sPath & sMask, Attr)
    sFileList = vbNullString

    Do While Len(CurrFile)
       sFileList = sFileList & CurrFile & ";"
       CurrFile = Dir
    Loop

    If sMask = "SAD*.dll" Then
       EnumFiles = Replace(sFileList, ".dll;", ".NewCommands=Load" & vbNewLine)
    Else
       EnumFiles = sFileList
    End If
End Function


Public Function LogError(ByVal sModuleName As String, sProcName As String, lError As Long, sErrorMsg As String) As Boolean
    Dim fh As Long
    Dim sMessage As String

    fh = FreeFile
    Open "SADERRS.LOG" For Append As #fh
         sMessage = "***** Error " & Format(lError, "00000") & " at: " & Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM")
         sMessage = sMessage & vbNewLine & "  *** Module:         " & sModuleName
         sMessage = sMessage & vbNewLine & "  *** Procedure:      " & sProcName
         sMessage = sMessage & vbNewLine & "  *** Description:    " & sErrorMsg
         Print #fh, sMessage
         MsgBox sMessage
         Print #fh, "  *** Program continued by user after error."
    Close #fh
End Function

Public Sub LogEvent(ByVal sMessage As String)
    Dim fh As Long

    fh = FreeFile
    Open "sadDebug.log" For Append As #fh
         Print #fh, sMessage
    Close #fh
End Sub

Public Function FindInCollection(colToFindIn As Object, sToFind As String) As Object
    Dim CurItem As Object
    For Each CurItem In colToFindIn
        If StrComp(CurItem.Name, sToFind, vbTextCompare) = 0 Then
           Set FindInCollection = CurItem
           Exit Function
        End If
    Next CurItem
    Set FindInCollection = Nothing
End Function


Public Function sFileContents(ByVal sPathAndFilename As String) As String
On Error Resume Next
    Dim fh As Long
    If Len(Dir$(sPathAndFilename)) Then
       fh = FreeFile
       Open sPathAndFilename For Input Access Read As #fh
            sFileContents = Input(LOF(fh), fh)
       Close #fh
    End If
End Function

Public Function sChoose(sChoices As String, Optional ByVal sDelimiter As String = ";", Optional ByVal sDefault As String)
On Error GoTo EH_Wizard_sChoose
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Function
    bInHereAlready = True

    If Len(sDelimiter) = 0 Then sDelimiter = ";"

    Dim frmX As ISandyWindowSelect
    Set frmX = CreateObject("SandyInstance.CSandyWindows").CreateForm("LIST")
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

Public Sub SetListIndex(cboToSearch As Object, sToFind As String)
    Dim nIndex As Long
    
    nIndex = FindListIndex(cboToSearch, sToFind)
    If nIndex > -1 Then cboToSearch.ListIndex = nIndex
End Sub

Public Sub SetListViewIndex(lvwToSearch As Object, sToFind As String)
   'Dim nIndex As Long
    
    FindListViewIndex lvwToSearch, sToFind, True
    
   'nIndex = FindListViewIndex(lvwToSearch, sToFind, True)
   'If nIndex > -1 Then lvwToSearch.ListIndex = nIndex
End Sub

Public Function FindListViewIndex(lvwToSearch As Object, sToFind As String, Optional ByVal bSelectOnFind As Boolean) As Long
    Dim i As Long
    Dim CurrItem As ListItem

    If Len(sToFind) > 0 Then
       i = 0
       For Each CurrItem In lvwToSearch.ListItems
           i = i + 1
           If StrComp(CurrItem.Text, sToFind, vbTextCompare) = 0 Then
              FindListViewIndex = i
              If bSelectOnFind Then
                 CurrItem.Selected = True
              End If
              Exit Function
           End If
       Next CurrItem
    End If
    FindListViewIndex = -1
End Function

' On Return:
'    sExtractToken =
'       Token # nToken
'    sOrigStr =
'       Token # (1 through nToken-1) & (nToken+1 through nTokenCount(sOrigStr))
'
' Notes:
'    Yes there are a lot of "Exit Function" statements...
'       That's to make the function return as quickly as possible
'
Public Function sExtractToken(ByRef sOrigStr As String, Optional ByVal nToken As Integer = 1, Optional ByVal strDelim As String = " ")
    Static strIn As String
    Static strOut As String
    Static nCurrTokenStart As Long
    Static nNextTokenStart As Long
    Static nLenDelim As Long

  ' Handle the "simple" cases (No delimiter, or token # less than 2)
    nLenDelim = Len(strDelim)
    If nToken < 1 Or nLenDelim = 0 Then
     ' Nothing to extract, return nothing
       Exit Function
    ElseIf nToken = 1 Then
       nCurrTokenStart = InStr(sOrigStr, strDelim)
       If nCurrTokenStart > 0 Then
          sExtractToken = Left$(sOrigStr, nCurrTokenStart - 1)
          sOrigStr = Trim$(Mid$(sOrigStr, nCurrTokenStart + nLenDelim))
          Exit Function
       Else
          sExtractToken = sOrigStr
          sOrigStr = vbNullString
          Exit Function
       End If
    End If

  ' Find the start of then nToken'th Token
    strIn = sOrigStr: strOut = vbNullString
    nToken = nToken - 1
    Do Until nToken = 0
       nCurrTokenStart = InStr(strIn, strDelim)
       If nCurrTokenStart = 0 Or Len(strIn) = 0 Then Exit Function
       strOut = strOut & Left$(strIn, nCurrTokenStart - 1)
       strIn = Mid$(strIn, nCurrTokenStart + nLenDelim)

     ' Check to see if this is the one the calling function is looking for
       nToken = nToken - 1
    Loop

  ' Now we're at the point" & gsWhere & "the token sought for resides
    nCurrTokenStart = InStr(strIn, strDelim)
    If nCurrTokenStart > 0 Then
       If nCurrTokenStart > 1 Then
          sExtractToken = Left$(strIn, nCurrTokenStart - 1)
       Else
          sExtractToken = vbNullString
       End If
     ' Rewrite the original string without the last token
       sOrigStr = Trim$(strOut & Mid$(strIn, nCurrTokenStart))
       Exit Function
    Else
       sExtractToken = strIn
       sOrigStr = Trim$(strOut)
       Exit Function
    End If
End Function

' ********************************************************************************
' Name              BrowseForFolder
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
       sPath = String$(MAX_PATH, 0)
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
' Name              bUserSure
'
' Parameters
'       sPrompt                      (I)  Opt. Question to ask the user
'                                         Default = "Are you sure this is what you want to do ?"
' Description
'
' Returns true if the user selects "Yes" from the MsgBox displayed
'
' ********************************************************************************
Public Function bUserSure(Optional ByVal sPrompt As String = "Are you sure this is what you want to do ?") As Boolean
    bUserSure = (MsgBox(sPrompt, vbYesNo, "ARE YOU SURE ?") = vbYes)
End Function

' ********************************************************************************
' Name              NextNegativeUnique
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
Public Function NextNegativeUnique(Optional ByVal InitOnly As Boolean) As Long
    Static lNextSerial As Long
    If Not InitOnly Then
       lNextSerial = lNextSerial - 1
    Else
       lNextSerial = -1
    End If
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
'   "William M Rawls"    vbNullString      1             No delimiter? String has one token,
'                                              "William M Rawls"
'   "1.00.05"            "."     3             "1", "00", and "05"
' ***********************************************************************************
Public Function lTokenCount(ByVal siAllTokens As String, Optional ByVal siDelim As String = " ") As Long
    Static iCurTokenLocation As Long ' Character position of the first delimiter string
    Static iTokensSoFar As Long      ' Used to keep track of how many tokens we've counted so far
    Static iDelim As Long            ' Length of the delimiter string

    iDelim = Len(siDelim)
    If iDelim < 1 Then
     ' Empty delimiter strings means only one token equal to the string
       lTokenCount = 1
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
             lTokenCount = iTokensSoFar + 1 'Abs(Len(siAllTokens) > 0)
             Exit Function
          End If
          iTokensSoFar = iTokensSoFar + 1
          siAllTokens = Mid$(siAllTokens, iCurTokenLocation + iDelim)
       Loop
    End If
End Function
' ********************************************************************************
' Name              nz
'
' Parameters
'      vData                         (O)  Variant to test for NULL
'       sDefault                     (O)  Opt. On NULL this string is returned
'                                         Default = vbNullString
' Description
'
' Returns sDefault if the variant is NULL, otherwise it returns the Variant
'
' ********************************************************************************
Public Function nZ(ByRef vData As Variant, Optional sDefault As String = vbNullString) As String
    If IsNull(vData) Then
       nZ = sDefault
    Else
       nZ = vData
    End If
End Function

Public Function sDenormalize(sLine As String) As String
    sDenormalize = Replace(Replace(sLine, "%$%EOL%$%", vbCrLf), "%$%TAB%$%", vbTab)
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
          sAfter = vbNullString
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
'   "William M Rawls"   -1       " "     vbNullString                 Negative tokens act same as zero
'   "William M Rawls"    1       vbNullString      "William M Rawls"  Same as zero
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
          sExcept = Mid$(siAllTokens, iCurTokenLocation + nDelim)
          Exit Function
       End If
    Else
     ' Find the Nth token
       Do
          iCurTokenLocation = InStr(siAllTokens, sDelim)
          If iToken = 1 Then
             If iCurTokenLocation > 0 Then
                sExcept = sReturned & sDelim & Mid$(siAllTokens, iCurTokenLocation + nDelim)
             Else
                sExcept = sReturned
             End If
             sReturned = vbNullString
             Exit Function
          ElseIf iCurTokenLocation = 0 Then
             sExcept = sReturned & sDelim & siAllTokens
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
' ******************************************************************************************
' Name              SandySupport.modGeneral_sInsertSpaces
'
' Parameters
'      sToInsertInto                 (I)  String.
'
' Returns
'      String                        .
'
' Description
'
' Inserts spaces into a string. Spaces are inserted at each capital letter.
' Common words not to be capitalized are handled automatically (a, an, the, or,
' of). The sub string "ID" is handled specially and is removed. This function is
' excellent for labels.
'
' ******************************************************************************************
Public Function sInsertSpaces(ByVal sToInsertInto As String) As String
    Dim bytOriginal() As Byte
    Dim sWithSpaces As String
    Dim nUpper As Long
    Dim nCurrent As Long
    Dim nA As Byte
    Dim nZ As Byte

    bytOriginal = StrConv(sToInsertInto, vbFromUnicode)
    nUpper = UBound(bytOriginal)
    
    For nCurrent = 0 To nUpper
        If bytOriginal(nCurrent) >= 65 And bytOriginal(nCurrent) <= 90 And nCurrent <> 0 Then
           sWithSpaces = sWithSpaces & " " & Chr$(bytOriginal(nCurrent))
        Else
           sWithSpaces = sWithSpaces & Chr$(bytOriginal(nCurrent))
        End If
    Next nCurrent
    
    sInsertSpaces = Replace(Replace(Replace(Replace(Replace(Replace(sWithSpaces, " Of ", " of "), " The ", " the "), " A ", " a "), " An ", " an "), " I D", vbNullString), "  ", " ")
End Function

Public Function sNormalize(sLine As String) As String
    sNormalize = Replace(Replace(sLine, vbNewLine, "%$%EOL%$%"), vbTab, "%$%TAB%$%")
End Function

Public Function sTableToPropertyName(ByVal sTableName As String) As String
    sTableToPropertyName = Replace(Replace(Replace(Replace(sTableName, " ", "_"), "*", "_"), "-", "_"), ".", "__")
End Function

Public Function zn(sData As String) As Variant
    If Len(sData) = 0 Then zn = Null Else zn = sData
End Function

Public Sub Main()
    gsEolTab = vbNewLine & vbTab
    gs2EOL = vbNewLine & vbNewLine
    gs2EOLTab = gs2EOL & vbTab

    Call NextNegativeUnique(True)
End Sub

Public Function lFindToken(ByVal sAllTokens As String, ByVal sTokenToFind As String, Optional ByVal sDelimiter As String = " ") As Long
    Dim lTokens As Long
    Dim l As Long

    lTokens = lTokenCount(sAllTokens, sDelimiter)

    For l = 1 To lTokens
        If StrComp(UCase$(sGetToken(sAllTokens, l, sDelimiter)), UCase$(sTokenToFind)) = 0 Then
           lFindToken = l
           Exit Function
        End If
    Next l

    lFindToken = 0
End Function

Public Function LocalGenerated(Length As Long) As String
    Dim CurrByte As Long
    Dim sOut As String
    
    For CurrByte = 1 To Length
        sOut = sOut & (CLng(Rnd() * 10) Mod 10)
    Next CurrByte
    
    LocalGenerated = sOut
End Function

Public Function sadDecrypt(strIn As String) As String
    Dim strOut As String
    If Len(strIn) = 0 Then Exit Function
    If Left$(strIn, 3) <> "EN*" Then Exit Function

    strIn = Scramble(strIn)
    Do While Len(strIn)
       strOut = strOut & Chr$((255 - Val("&H" & Left$(strIn, 2) & "&")) Mod 255)
       strIn = Mid$(strIn, 3)
    Loop
    sadDecrypt = strOut
End Function

Public Function sadEncrypt(ByVal strIn As String) As String
    Dim strOut As String
    Dim bytArray() As Byte
    Dim CurrByte As Long

    bytArray = StrConv(strIn, vbFromUnicode)
    For CurrByte = 0 To UBound(bytArray)
        If bytArray(CurrByte) < 240 Then
           strOut = strOut & Hex$(255 - bytArray(CurrByte))
        Else
           strOut = strOut & "0" & Hex$(255 - bytArray(CurrByte))
        End If
    Next CurrByte

    sadEncrypt = "EN* " & Scramble(strOut)
End Function


Public Function Scramble(ByVal strIn As String) As String
    Dim strOut As String
    Dim bytArray() As Byte
    Dim bytStack As Byte
    Dim CurrByte As Long
    Dim MaxCount As Long

    If Left$(strIn, 4) = "EN* " Then
       strIn = Mid$(strIn, 5)
    End If

    bytArray = strIn
    MaxCount = UBound(bytArray)
    MaxCount = MaxCount - (MaxCount Mod 2) - 8
       For CurrByte = 0 To MaxCount Step 8
           bytStack = bytArray(CurrByte + 0)
           bytArray(CurrByte + 0) = bytArray(CurrByte + 6)
           bytArray(CurrByte + 6) = bytStack
           bytStack = bytArray(CurrByte + 2)
           bytArray(CurrByte + 2) = bytArray(CurrByte + 4)
           bytArray(CurrByte + 4) = bytStack
       Next CurrByte
    strOut = bytArray
    Scramble = strOut
End Function



