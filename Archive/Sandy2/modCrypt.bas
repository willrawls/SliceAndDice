Attribute VB_Name = "modCrypt"
Option Explicit

Private Const msHaxValues As String = "000102030405060708090A0B0C0D0E0F101112131415161718191A1B1C1D1E1F202122232425262728292A2B2C2D2E2F303132333435363738393A3B3C3D3E3F404142434445464748494A4B4C4D4E4F505152535455565758595A5B5C5D5E5F606162636465666768696A6B6C6D6E6F707172737475767778797A7B7C7D7E7F808182838485868788898A8B8C8D8E8F909192939495969798999A9B9C9D9E9FA0A1A2A3A4A5A6A7A8A9AAABACADAEAFB0B1B2B3B4B5B6B7B8B9BABBBCBDBEBFC0C1C2C3C4C5C6C7C8C9CACBCCCDCECFD0D1D2D3D4D5D6D7D8D9DADBDCDDDEDFE0E1E2E3E4E5E6E7E8E9EAEBECEDEEEFF0F1F2F3F4F5F6F7F8F9FAFBFCFDFEFF"
Public Dummy As Byte
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
       strOut = strOut & Chr((255 - Val("&H" & Left$(strIn, 2) & "&")) Mod 255)
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
           strOut = strOut & Hex(255 - bytArray(CurrByte))
        Else
           strOut = strOut & "0" & Hex(255 - bytArray(CurrByte))
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
