Attribute VB_Name = "modGeneral"
Option Explicit

Public Sub Main()
    frmMain.Show
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

Public Function LoadFormPosition(ByVal frmForm As Form, Optional ByVal bAutoCenter = True)
    If GetSetting(App.ProductName, frmForm.Name, "Position Saved", False) Then
       frmForm.Left = GetSetting(App.ProductName, frmForm.Name, "Form Position Left", frmForm.Left)
       frmForm.Top = GetSetting(App.ProductName, frmForm.Name, "Form Position Top", frmForm.Top)
       frmForm.Width = GetSetting(App.ProductName, frmForm.Name, "Form Position Width", frmForm.Width)
       frmForm.Height = GetSetting(App.ProductName, frmForm.Name, "Form Position Height", frmForm.Height)
    ElseIf bAutoCenter Then
       frmForm.Left = (Screen.Width - frmForm.Width) / 2
       frmForm.Top = (Screen.Height - frmForm.Height) / 2
    End If
End Function

Public Function SaveFormPosition(ByVal frmForm As Form)
    SaveSetting App.ProductName, frmForm.Name, "Position Saved", True
    SaveSetting App.ProductName, frmForm.Name, "Form Position Left", frmForm.Left
    SaveSetting App.ProductName, frmForm.Name, "Form Position Top", frmForm.Top
    SaveSetting App.ProductName, frmForm.Name, "Form Position Width", frmForm.Width
    SaveSetting App.ProductName, frmForm.Name, "Form Position Height", frmForm.Height
End Function
