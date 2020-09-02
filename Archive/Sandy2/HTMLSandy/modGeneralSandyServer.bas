Attribute VB_Name = "modGeneral"
Option Explicit

Public db As Database
Public Sandy As Object 'SliceAndDice.CSliceAndDice
Public IsSandyLoadedYet As Boolean

Public CurrentCategory As String
Public CurrentTemplate As String


Private Const msHaxValues As String = "000102030405060708090A0B0C0D0E0F101112131415161718191A1B1C1D1E1F202122232425262728292A2B2C2D2E2F303132333435363738393A3B3C3D3E3F404142434445464748494A4B4C4D4E4F505152535455565758595A5B5C5D5E5F606162636465666768696A6B6C6D6E6F707172737475767778797A7B7C7D7E7F808182838485868788898A8B8C8D8E8F909192939495969798999A9B9C9D9E9FA0A1A2A3A4A5A6A7A8A9AAABACADAEAFB0B1B2B3B4B5B6B7B8B9BABBBCBDBEBFC0C1C2C3C4C5C6C7C8C9CACBCCCDCECFD0D1D2D3D4D5D6D7D8D9DADBDCDDDEDFE0E1E2E3E4E5E6E7E8E9EAEBECEDEEEFF0F1F2F3F4F5F6F7F8F9FAFBFCFDFEFF"
Public Dummy As Byte

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
          siAllTokens = Mid(siAllTokens, iCurTokenLocation + iDelim)
       Loop
    End If
End Function

Public Function sadDecrypt(strIn As String) As String
    Dim strOut As String
    If Len(strIn) = 0 Then Exit Function
    If Left(strIn, 3) <> "EN*" Then Exit Function

    strIn = Scramble(strIn)
    Do While Len(strIn)
       strOut = strOut & Chr((255 - Val("&H" & Left(strIn, 2) & "&")) Mod 255)
       strIn = Mid(strIn, 3)
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
           strOut = strOut & Hex(255 - bytArray(CurrByte))
        Else
           strOut = strOut & "0" & Hex(255 - bytArray(CurrByte))
        End If
    Next CurrByte

    sadEncrypt = "EN* " & Scramble(strOut)
End Function

Public Function sadInvoiceEncrypt(ByVal sInvoiceNumber As String) As String
On Error Resume Next
    Dim strOut      As String
    Dim bytArray()  As Byte
    Dim CurrByte    As Long
    Dim ValueLen    As Long
    Dim OffsetLen   As Long
    Dim CharLoc     As Long
    Dim StartAt     As Long
    Dim CurrOffset  As Long
    Const Offsets   As String = "61524351678425904521800218024862068410257946231578781579516891"
    Const Values    As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"

    If Len(sInvoiceNumber) = 0 Then Exit Function

    Randomize Now + Rnd
    ValueLen = Len(Values)
    OffsetLen = Len(Offsets)
    StartAt = Int(15 * Rnd + 1)
    CurrOffset = StartAt

    sInvoiceNumber = "F" & Replace(UCase(sInvoiceNumber), "O", "0") & "S"
    bytArray = StrConv(sInvoiceNumber, vbFromUnicode)
    strOut = ""

    For CurrByte = 0 To UBound(bytArray)
        CharLoc = InStr(Values, Chr(bytArray(CurrByte)))
        If CharLoc < 1 Or CharLoc > ValueLen Then Exit Function
        CharLoc = CharLoc + Val(Mid(Offsets, CurrOffset, 1))
        If CharLoc > ValueLen Then CharLoc = CharLoc - ValueLen
        strOut = strOut & Mid(Values, CharLoc, 1)
        CurrOffset = CurrOffset + 1
        If CurrOffset > OffsetLen Then CurrOffset = 1
    Next CurrByte

    strOut = Right(Format(StartAt, "00"), 1) & strOut & Left(Format(StartAt, "00"), 1)
    sadInvoiceEncrypt = Replace(strOut, "O", ".")
End Function

Public Function sadInvoiceDecrypt(ByVal sInvoiceNumber As String) As String
On Error Resume Next
    Dim strOut      As String
    Dim bytArray()  As Byte
    Dim CurrByte    As Long
    Dim ValueLen    As Long
    Dim OffsetLen   As Long
    Dim CharLoc     As Long
    Dim StartAt     As Long
    Dim CurrOffset  As Long
    Const Offsets   As String = "61524351678425904521800218024862068410257946231578781579516891"
    Const Values    As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"

    If Len(sInvoiceNumber) = 0 Then Exit Function

    ValueLen = Len(Values)
    OffsetLen = Len(Offsets)
    sInvoiceNumber = Replace(Replace(UCase(sInvoiceNumber), "O", "0"), ".", "O")
    StartAt = Val(Right(sInvoiceNumber, 1) & Left(sInvoiceNumber, 1))
    sInvoiceNumber = Mid(sInvoiceNumber, 2, Len(sInvoiceNumber) - 2)
    CurrOffset = StartAt
    bytArray = StrConv(sInvoiceNumber, vbFromUnicode)
    strOut = ""

    For CurrByte = 0 To UBound(bytArray)
        CharLoc = InStr(Values, Chr(bytArray(CurrByte)))
        If CharLoc < 1 Or CharLoc > ValueLen Then Exit Function
        CharLoc = CharLoc - Val(Mid(Offsets, CurrOffset, 1))
        If CharLoc < 1 Then CharLoc = ValueLen - CharLoc
        strOut = strOut & Mid(Values, CharLoc, 1)
        CurrOffset = CurrOffset + 1
        If CurrOffset > OffsetLen Then CurrOffset = 1
    Next CurrByte

    If Left(strOut, 1) = "F" And Right(strOut, 1) = "S" Then
       sadInvoiceDecrypt = Mid(strOut, 2, Len(strOut) - 2)
    Else
       sadInvoiceDecrypt = ""
    End If
End Function

Public Function Scramble(ByVal strIn As String) As String
    Dim strOut As String
    Dim bytArray() As Byte
    Dim CurrByte As Long
    Dim bytStack As Byte
    Dim Shift As Integer
    Dim MaxCount As Integer

    If Left(strIn, 4) = "EN* " Then
       strIn = Mid(strIn, 5)
    End If

    bytArray = strIn
    MaxCount = UBound(bytArray)
    MaxCount = MaxCount - (MaxCount Mod 2) - 8
       For CurrByte = 0 To MaxCount Step 8 'Step 8
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
' ********************************************************************************
Public Function sReplace(ByVal sAll As String, ByVal sFind As String, ByVal sReplaceWith As String) As String
    Dim iCurFindPos As Long
    Dim iFind As Long
    Dim sOut As String

    iFind = Len(sFind)
    iCurFindPos = InStr(sAll, sFind)
    If InStr(sReplaceWith, sFind) = 0 Then
       Do While iCurFindPos > 0
          If iCurFindPos > 1 Then
             sAll = Left(sAll, iCurFindPos - 1) & sReplaceWith & Mid(sAll, iCurFindPos + iFind)
          Else
             sAll = sReplaceWith & Mid(sAll, iCurFindPos + iFind)
          End If
          iCurFindPos = InStr(sAll, sFind)
       Loop
       sReplace = sAll
    Else
       Do While iCurFindPos > 0
          If iCurFindPos > 1 Then
             sOut = sOut & Left(sAll, iCurFindPos - 1) & sReplaceWith
             sAll = Mid(sAll, iCurFindPos + iFind)
          Else
             sOut = sOut & sReplaceWith
             sAll = Mid(sAll, iCurFindPos + iFind)
          End If
          iCurFindPos = InStr(sAll, sFind)
       Loop
       
       sReplace = sOut & sAll
    End If
End Function

