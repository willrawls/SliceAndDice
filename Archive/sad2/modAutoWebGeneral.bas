Attribute VB_Name = "modGeneral"
Option Explicit


Public Function CollectFormData(wb As WebBrowser) As String
    Dim lCurrForm As Long
    Dim lCurrField As Long
    Dim CurrForm As Object
    Dim CurrField As Object
    Dim sOut As String
On Error Resume Next
    For lCurrForm = 0 To wb.Document.Forms.length - 1
        Set CurrForm = wb.Document.Forms(lCurrForm)
        For lCurrField = 0 To CurrForm.length - 1
            Set CurrField = CurrForm(lCurrField)
            sOut = sOut & lCurrForm & "." & CurrField.Name & "=" & CurrField.Value & "&"
        Next lCurrField
    Next lCurrForm
    Set CurrField = Nothing
    Set CurrForm = Nothing
    
    CollectFormData = sOut
End Function

Public Function ListToString(ByRef lstToRead As Control, Optional ByVal bMoveBackward = True, Optional ByVal sDelimiter As String = vbCrLf) As String
    Dim CurrItem As Long
    Dim sOut As String

    If bMoveBackward Then
       For CurrItem = lstToRead.ListCount - 1 To 0 Step -1
           sOut = sOut & lstToRead.List(CurrItem) & sDelimiter
       Next CurrItem
    Else
       For CurrItem = 0 To lstToRead.ListCount - 1
           sOut = sOut & lstToRead.List(CurrItem) & sDelimiter
       Next CurrItem
    End If

    ListToString = sOut
End Function

Public Function LoadFormPosition(frmToActOn As Form, Optional ByVal bAutoCenter = True)
    Dim ProductName As String
    
    If Len(ProductName) = 0 Then
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

Public Sub Main()

End Sub

Public Function SaveFormPosition(frmToActOn As Form)
    Dim ProductName As String
    
    If Len(ProductName) = 0 Then
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

Public Sub StringToList(ByVal sContents As String, ByRef lstToFill As Control, Optional ByVal bClearFirst As Boolean = True, Optional ByVal sDelimiter As String = vbCrLf)
    Dim CurrItem As Long
    Dim sEntry As String

    If bClearFirst Then lstToFill.Clear
    Do While InStr(1, sContents, sDelimiter, vbTextCompare)
       sEntry = Left$(sContents, InStr(1, sContents, sDelimiter, vbTextCompare) - 1)
       lstToFill.AddItem sEntry
       sContents = Mid$(sContents, InStr(1, sContents, sDelimiter, vbTextCompare) + Len(sDelimiter))
    Loop
End Sub


' ***********************************************************************************
' Synopsis          Returns the Nth Token from sAllTokens delimited by sDelim
'
' Parameters
Public Function LogError(ByVal sModuleName As String, sProcName As String, lError As Long, sErrorMsg As String) As Boolean
    Dim fh As Long
    Dim sMessage As String

    fh = FreeFile
    Open "ERRORLOG.TXT" For Append As #fh
         sMessage = "***** Error " & Format(lError, "00000") & " at: " & Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM")
         sMessage = sMessage & Chr(13) & "  *** Module:         " & sModuleName
         sMessage = sMessage & Chr(13) & "  *** Procedure:      " & sProcName
         sMessage = sMessage & Chr(13) & "  *** Description:    " & sErrorMsg
         Print #fh, sMessage
         sMessage = sMessage & Chr(13) & Chr(13) & Chr(9) & "Continue after error ? (No to exit program)"
         If MsgBox(sMessage, vbYesNo) = vbNo Then
            Print #fh, "  *** Program shut down by user after error."
            ShutDownNicely
         Else
            Print #fh, "  *** Program continued by user after error."
         End If
    Close #fh

End Function

Public Sub ShutDownNicely()
On Error Resume Next
  ' Close all objects, forms, handles, etc. here
    If frmBrowser.Visible Then frmBrowser.Hide
   'If frmMain.Visible Then frmMain.Hide
   'frmMain.IEBrowsers.Clear
    Err.Clear
    'End
End Sub
'
'   sAllTokens                 (I) Required. The string containing all the tokens
'   iToken                      (I) Optional. The index of the token to return
'                                   DEFAULT = 1
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = " "
' Description
'  For the following:
'    sAllTokens         iToken   sDelim  Returns       Notes
'   "William M Rawls"    1       " "     "William"      First word
'   "William M Rawls"    2       " "     "M"            Second word
'   "William M Rawls"    3       " "     "Rawls"        Third word
'   "William M Rawls"    4       " "     ""             No forth word
'   "William M Rawls"    0       " "     ""             Zeroth token is always empty
'   "William M Rawls"   -1       " "     ""             Negative tokesn always empty
'   "William M Rawls"    1       ""      ""             No delimiter ? Token empty
' ***********************************************************************************
Public Function sGetToken(ByVal sAllTokens As String, Optional ByVal iToken As Integer = 1, Optional ByVal sDelim As String = " ") As String
    Dim iCurTokenLocation As Long ' Character position of the first delimiter string
    Dim nDelim As Integer            ' Length of the delimiter string
    nDelim = Len(sDelim)

    If iToken < 1 Or nDelim < 1 Then
     ' Negative or zeroth token or empty delimiter strings mean an empty token
       Exit Function
    ElseIf iToken = 1 Then
     ' Quickly extract the first token
       iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare)
       If iCurTokenLocation > 1 Then
          sGetToken = Left$(sAllTokens, iCurTokenLocation - 1)
       ElseIf iCurTokenLocation = 1 Then
          sGetToken = ""
       Else
          sGetToken = sAllTokens
       End If
       Exit Function
    Else
     ' Find the Nth token
       Do
          iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare)
          If iCurTokenLocation = 0 Then
             Exit Function
          Else
             sAllTokens = Mid$(sAllTokens, iCurTokenLocation + nDelim)
          End If
          iToken = iToken - 1
       Loop Until iToken = 1

     ' Extract the Nth token (Which is the next token at this point)
       iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare)
       If iCurTokenLocation > 0 Then
          sGetToken = Left$(sAllTokens, iCurTokenLocation - 1)
          Exit Function
       Else
          sGetToken = sAllTokens
          Exit Function
       End If
    End If
End Function
' *********************************************************************************************
' Synopsis          Returns everything AFTER the Nth Token from sAllTokens delimited by sDelim
'
' Parameters
'
'   sAllTokens                 (I) Required. The string containing all the tokens
'   iToken                      (I) Optional. The index of the token to use as an "after" ref
'                                   DEFAULT = 1
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = " "
' Description
'  For the following:
'    sAllTokens         iToken   sDelim  Returns           Notes
'   "William M Rawls"    1       " "     "M Rawls"          After the first word
'   "William M Rawls"    2       " "     "Rawls"            After the second word
'   "William M Rawls"    3       " "     ""                 After the third word (nothing)
'   "William M Rawls"    0       " "     "William M Rawls"  After zeroth token is always the input string
'   "William M Rawls"   -1       " "     "William M Rawls"  Negative tokens act same as zero
'   "William M Rawls"    1       ""      "William M Rawls"  Same as one
' *********************************************************************************************
Public Function sAfter(ByVal sAllTokens As String, Optional ByVal iToken As Integer = 1, Optional ByVal sDelim As String = " ") As String
    Dim iCurTokenLocation As Long ' Character position of the first delimiter string
    Dim nDelim As Integer            ' Length of the delimiter string
    
    nDelim = Len(sDelim)
    If iToken < 1 Or nDelim < 1 Then
     ' Negative or zeroth token or empty delimiter strings mean an empty token
       sAfter = sAllTokens
       Exit Function
    ElseIf iToken = 1 Then
     ' Quickly extract the first token
       iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare)
       If iCurTokenLocation > 1 Then
          sAfter = Mid$(sAllTokens, iCurTokenLocation + nDelim)
          Exit Function
       ElseIf iCurTokenLocation = 0 Then
          sAfter = ""
          Exit Function
       Else
          sAfter = Mid$(sAllTokens, nDelim + 1)
          Exit Function
       End If
    Else
     ' Find the Nth token
       Do
          iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare)
          If iCurTokenLocation = 0 Then
             Exit Function
          Else
             sAllTokens = Mid$(sAllTokens, iCurTokenLocation + nDelim)
          End If
          iToken = iToken - 1
       Loop Until iToken = 1

     ' Extract the Nth token (Which is the next token at this point)
       iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare)
       If iCurTokenLocation > 0 Then
          sAfter = Mid$(sAllTokens, iCurTokenLocation + nDelim)
          Exit Function
       Else
          Exit Function
       End If
    End If
End Function
' **********************************************************************************************
' Synopsis          Returns everything BEFORE the Nth Token from sAllTokens delimited by sDelim
'
' Parameters
'
'   sAllTokens                 (I) Required. The string containing all the tokens
'   iToken                      (I) Optional. The index of the token to use as a "before" ref
'                                   DEFAULT = 2
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = " " (Space)
' Description
'  For the following:
'    sAllTokens         iToken   sDelim  Returns           Notes
'   "William M Rawls"    1       " "     ""                 Before the first word (nothing)
'   "William M Rawls"    2       " "     "William"          Before the second word
'   "William M Rawls"    3       " "     "William M"        Before the third word
'   "William M Rawls"    0       " "     ""                 Before zeroth token (nothing)
'   "William M Rawls"   -1       " "     ""                 Negative tokens act same as zero
'   "William M Rawls"    1       ""      ""                 Same as one
' *********************************************************************************************
Public Function sBefore(ByVal sAllTokens As String, Optional ByVal iToken As Integer = 2, Optional ByVal sDelim As String = " ") As String
    Dim iCurTokenLocation As Long ' Character position of the first delimiter string
    Dim nDelim As Integer            ' Length of the delimiter string
    Dim sReturned As String

    nDelim = Len(sDelim)
    If iToken < 2 Or nDelim < 1 Then
     ' First, Zeroth, or Negative tokens or empty delimiter strings mean an empty string returned
       sBefore = ""
       Exit Function
    ElseIf iToken = 2 Then
     ' Quickly extract the first token
       sBefore = sGetToken(sAllTokens, 1, sDelim)
       Exit Function
    Else
     ' Find the Nth token
       Do
          iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare)
          If iCurTokenLocation = 0 Or iToken = 1 Then
             sBefore = sReturned
             sReturned = ""
             Exit Function
          ElseIf Len(sReturned) = 0 Then
             sReturned = Left$(sAllTokens, iCurTokenLocation - 1)
          Else
             sReturned = sReturned & sDelim & Left$(sAllTokens, iCurTokenLocation - 1)
          End If
          sAllTokens = Mid$(sAllTokens, iCurTokenLocation + nDelim)
          iToken = iToken - 1
       Loop
    End If
End Function
Public Function lFindToken(ByVal sAllTokens As String, ByVal sTokenToFind As String, Optional ByVal sDelimiter As String = " ") As Long
    Dim lTokens As Long
    Dim l As Long

    lTokens = lTokenCount(sAllTokens, sDelimiter)

    For l = 1 To lTokens
        If StrComp(UCase$(Left$(sGetToken(sAllTokens, l, sDelimiter), Len(sTokenToFind))), UCase$(sTokenToFind)) = 0 Then
           lFindToken = l
           Exit Function
        End If
    Next

    lFindToken = 0
End Function

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
'   "William M Rawls"    ""      1             No delimiter? String has one token,
'                                              "William M Rawls"
'   "1.00.05"            "."     3             "1", "00", and "05"
' ***********************************************************************************
Public Function lTokenCount(ByVal sAllTokens As String, Optional ByVal siDelim As String = " ") As Long
    Dim iCurTokenLocation As Long ' Character position of the first delimiter string
    Dim iTokensSoFar As Integer      ' Used to keep track of how many tokens we've counted so far
    Dim iDelim As Integer            ' Length of the delimiter string

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
          iCurTokenLocation = InStr(1, sAllTokens, siDelim, vbTextCompare)
          If iCurTokenLocation = 0 Then
             lTokenCount = iTokensSoFar + 1 'Abs(Len(sAllTokens) > 0)
             Exit Function
          End If
          iTokensSoFar = iTokensSoFar + 1
          sAllTokens = Mid$(sAllTokens, iCurTokenLocation + iDelim)
       Loop
    End If
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
Public Function bUserSure(Optional ByVal sPrompt As String = "Are you sure this is what you want to do ?", Optional ByVal sTitle As String = "ARE YOU SURE ?") As Boolean
    bUserSure = (MsgBox(sPrompt, vbYesNo, "ARE YOU SURE ?") = vbYes)
End Function
