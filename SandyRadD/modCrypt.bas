Attribute VB_Name = "modCrypt"
Option Explicit

Private Const msHaxValues As String = "000102030405060708090A0B0C0D0E0F101112131415161718191A1B1C1D1E1F202122232425262728292A2B2C2D2E2F303132333435363738393A3B3C3D3E3F404142434445464748494A4B4C4D4E4F505152535455565758595A5B5C5D5E5F606162636465666768696A6B6C6D6E6F707172737475767778797A7B7C7D7E7F808182838485868788898A8B8C8D8E8F909192939495969798999A9B9C9D9E9FA0A1A2A3A4A5A6A7A8A9AAABACADAEAFB0B1B2B3B4B5B6B7B8B9BABBBCBDBEBFC0C1C2C3C4C5C6C7C8C9CACBCCCDCECFD0D1D2D3D4D5D6D7D8D9DADBDCDDDEDFE0E1E2E3E4E5E6E7E8E9EAEBECEDEEEFF0F1F2F3F4F5F6F7F8F9FAFBFCFDFEFF"
Public Dummy As Byte

Public Function LocalGenerated(Length As Long) As String
1        Dim CurrByte As Long
2        Dim sOut As String

3        For CurrByte = 1 To Length
4            sOut = sOut & (CLng(Rnd() * 10) Mod 10)
5        Next CurrByte

6        LocalGenerated = sOut
End Function

Public Function sadDecrypt(strIn As String) As String
7        Dim strOut As String
8        If Len(strIn) = 0 Then Exit Function
9        If Left$(strIn, 3) <> "EN*" Then Exit Function

10       strIn = Scramble(strIn)
11       Do While Len(strIn)
12           strOut = strOut & Chr$((255 - Val("&H" & Left$(strIn, 2) & "&")) Mod 255)
13           strIn = Mid$(strIn, 3)
14       Loop
15       sadDecrypt = strOut
End Function

Public Function sadEncrypt(ByVal strIn As String) As String
16       Dim strOut As String
17       Dim bytArray() As Byte
18       Dim CurrByte As Long

19       bytArray = StrConv(strIn, vbFromUnicode)
20       For CurrByte = 0 To UBound(bytArray)
21           If bytArray(CurrByte) < 240 Then
22               strOut = strOut & Hex$(255 - bytArray(CurrByte))
23           Else
24               strOut = strOut & "0" & Hex$(255 - bytArray(CurrByte))
25           End If
26       Next CurrByte

27       sadEncrypt = "EN* " & Scramble(strOut)
End Function


Public Function Scramble(ByVal strIn As String) As String
28       Dim strOut As String
29       Dim bytArray() As Byte
30       Dim bytStack As Byte
31       Dim CurrByte As Long
32       Dim MaxCount As Long

33       If Left$(strIn, 4) = "EN* " Then
34           strIn = Mid$(strIn, 5)
35       End If

36       bytArray = strIn
37       MaxCount = UBound(bytArray)
38       MaxCount = MaxCount - (MaxCount Mod 2) - 8
39       For CurrByte = 0 To MaxCount Step 8
40           bytStack = bytArray(CurrByte + 0)
41           bytArray(CurrByte + 0) = bytArray(CurrByte + 6)
42           bytArray(CurrByte + 6) = bytStack
43           bytStack = bytArray(CurrByte + 2)
44           bytArray(CurrByte + 2) = bytArray(CurrByte + 4)
45           bytArray(CurrByte + 4) = bytStack
46       Next CurrByte
47       strOut = bytArray
48       Scramble = strOut
End Function
