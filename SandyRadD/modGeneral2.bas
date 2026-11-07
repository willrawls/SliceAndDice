Attribute VB_Name = "modGeneral"
Option Explicit
' ***********************************
' ****** Extend ListView stuff ******
' ***********************************
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Const LVM_FIRST = &H1000
  Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
  Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
  Private Const LVS_EX_FULLROWSELECT = &H20

Public Const ChildDelimiter As String = "<CHILD>"
Public Const EndChildDelimiter As String = "<ENDCHILD>"
Public Const IconDelimiter As String = "<ICON>"
Public Const TagDelimiter As String = "<TAG>"



' ********************************************************************************
' Name              ExtendListView
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
    Dim lReturned As Long

    style = SendMessage(hWndListView, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    style = style Or LVS_EX_FULLROWSELECT
    lReturned = SendMessage(hWndListView, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, style)
End Sub

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
Public Function iTokenCount(ByVal siAllTokens As String, Optional ByVal siDelim As String = " ") As Integer
    Static iCurTokenLocation As Long ' Character position of the first delimiter string
    Static iTokensSoFar As Integer      ' Used to keep track of how many tokens we've counted so far
    Static iDelim As Integer            ' Length of the delimiter string

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
          siAllTokens = Mid$(siAllTokens, iCurTokenLocation + iDelim)
       Loop
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
Public Function sAfter(ByVal siAllTokens As String, Optional ByVal iToken As Integer = 1, Optional ByVal sDelim As String = " ") As String
    Static iCurTokenLocation As Long ' Character position of the first delimiter string
    Static nDelim As Integer            ' Length of the delimiter string
    
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
          sAfter = ""
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
Public Function sGetToken(ByVal siAllTokens As String, Optional ByVal iToken As Integer = 1, Optional ByVal sDelim As String = " ") As String
    Static iCurTokenLocation As Long ' Character position of the first delimiter string
    Static nDelim As Integer            ' Length of the delimiter string
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

'****************************************************************************
' SetListIndex:
'
'   Given a list or combobox and a string to search for, the list entry matching
'   the string to find will be selected.
'
' Parameters:
'      cboToSearch                         (I) List or Combobox to search
'      sItemToFind                         (I) Item in list to select
'
' Returns
'      Nothing
'****************************************************************************
Public Sub SetListIndex(cboToSearch As Control, ByVal sItemToFind As String)
    cboToSearch.ListIndex = GetListIndex(cboToSearch, sItemToFind)
End Sub

Public Function GetListIndex(cboToSearch As Control, ByVal sItemToFind As String) As Integer
    Static nCurItem As Integer

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

Public Function sNormalize(sLine As String) As String
    sNormalize = sReplace(sReplace(sLine, Chr$(13) + Chr$(10), "%$%EOL%$%"), Chr$(9), "%$%TAB%$%")
End Function

' ********************************************************************************
' Name              sReplace
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
    Dim iFind As Integer

    iFind = Len(sFind)
    iCurFindPos = InStr(sAll, sFind)
    Do While iCurFindPos > 0
       If iCurFindPos > 1 Then
          sAll = Left$(sAll, iCurFindPos - 1) & sReplaceWith & Mid$(sAll, iCurFindPos + iFind)
       Else
          sAll = sReplaceWith & Mid$(sAll, iCurFindPos + iFind)
       End If
       iCurFindPos = InStr(sAll, sFind)
    Loop
    sReplace = sAll
End Function

Public Function sDenormalize(sLine As String) As String
    sDenormalize = sReplace(sReplace(sLine, "%$%EOL%$%", Chr$(13) + Chr$(10)), "%$%TAB%$%", Chr$(9))
End Function

'   "William M Rawls"    0       " "     ""                 Before zeroth token (nothing)
'   "William M Rawls"   -1       " "     ""                 Negative tokens act same as zero
'   "William M Rawls"    1       ""      ""                 Same as one
' *********************************************************************************************
Public Function sBefore(ByVal siAllTokens As String, Optional ByVal iToken As Integer = 2, Optional ByVal sDelim As String = " ") As String
    Static iCurTokenLocation As Long ' Character position of the first delimiter string
    Static nDelim As Integer            ' Length of the delimiter string
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
             sReturned = Left$(siAllTokens, iCurTokenLocation - 1)
          Else
             sReturned = sReturned & sDelim & Left$(siAllTokens, iCurTokenLocation - 1)
          End If
          siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
          iToken = iToken - 1
       Loop
    End If
End Function

