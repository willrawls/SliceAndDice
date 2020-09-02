Attribute VB_Name = "modGeneral"
Option Explicit

Private Const msHaxValues As String = "000102030405060708090A0B0C0D0E0F101112131415161718191A1B1C1D1E1F202122232425262728292A2B2C2D2E2F303132333435363738393A3B3C3D3E3F404142434445464748494A4B4C4D4E4F505152535455565758595A5B5C5D5E5F606162636465666768696A6B6C6D6E6F707172737475767778797A7B7C7D7E7F808182838485868788898A8B8C8D8E8F909192939495969798999A9B9C9D9E9FA0A1A2A3A4A5A6A7A8A9AAABACADAEAFB0B1B2B3B4B5B6B7B8B9BABBBCBDBEBFC0C1C2C3C4C5C6C7C8C9CACBCCCDCECFD0D1D2D3D4D5D6D7D8D9DADBDCDDDEDFE0E1E2E3E4E5E6E7E8E9EAEBECEDEEEFF0F1F2F3F4F5F6F7F8F9FAFBFCFDFEFF"
Public Dummy As Byte

Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
End Type
 
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long

Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&

Const REG_NONE = 0&
Const REG_SZ = 1&
Const REG_EXPAND_SZ = 2&
Const REG_BINARY = 3&
Const REG_DWORD = 4&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ

Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte

'This constant determins wether or not to display error messages to the
'user. I have set the default value to False as an error message can and
'does become irritating after a while. Turn this value to true if you want
'to debug your programming code when reading and writing to your system
'registry, as any errors will be displayed in a message box.

Const DisplayErrorMsg = False


Public Function sadGetLicenseKey(ByVal Key As String, Optional ByVal sDefault As String) As String
1    On Error Resume Next
2        Dim sResult As String
3        sResult = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License", Key)
4        If Len(sResult) = 0 Or sResult = "Error" Then sResult = sDefault
5        If Left$(sResult, 4) = "EN* " Then
6           sResult = sadDecrypt(sResult)
7        End If
8        sadGetLicenseKey = sResult
End Function

Function SetDWORDValue(SubKey As String, Entry As String, Value As Long)

9    Call ParseKey(SubKey, MainKeyHandle)

10   If MainKeyHandle Then
11      rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
12      If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
13         rtn = RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4) 'write the value
14         If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
15            If DisplayErrorMsg = True Then 'if the user want errors displayed
16               MsgBox ErrorMsg(rtn)        'display the error
17            End If
18         End If
19         rtn = RegCloseKey(hKey) 'close the key
20      Else 'if there was an error opening the key
21         If DisplayErrorMsg = True Then 'if the user want errors displayed
22            MsgBox ErrorMsg(rtn) 'display the error
23         End If
24      End If
25   End If

End Function
Function GetDWORDValue(SubKey As String, Entry As String)

26   Call ParseKey(SubKey, MainKeyHandle)

27   If MainKeyHandle Then
28      rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
29      If rtn = ERROR_SUCCESS Then 'if the key could be opened then
30         rtn = RegQueryValueExA(hKey, Entry, 0, REG_DWORD, lBuffer, 4) 'get the value from the registry
31         If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
32            rtn = RegCloseKey(hKey)  'close the key
33            GetDWORDValue = lBuffer  'return the value
34         Else                        'otherwise, if the value couldnt be retreived
35            GetDWORDValue = "Error"  'return Error to the user
36            If DisplayErrorMsg = True Then 'if the user wants errors displayed
37               MsgBox ErrorMsg(rtn)        'tell the user what was wrong
38            End If
39         End If
40      Else 'otherwise, if the key couldnt be opened
41         GetDWORDValue = "Error"        'return Error to the user
42         If DisplayErrorMsg = True Then 'if the user wants errors displayed
43            MsgBox ErrorMsg(rtn)        'tell the user what was wrong
44         End If
45      End If
46   End If

End Function
Function GetBinaryValue(SubKey As String, Entry As String)

47   Call ParseKey(SubKey, MainKeyHandle)

48   If MainKeyHandle Then
49      rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
50      If rtn = ERROR_SUCCESS Then 'if the key could be opened
51         lBufferSize = 1
52         rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, 0, lBufferSize) 'get the value from the registry
53         sBuffer = Space(lBufferSize)
54         rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize) 'get the value from the registry
55         If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
56            rtn = RegCloseKey(hKey)  'close the key
57            GetBinaryValue = sBuffer 'return the value to the user
58         Else                        'otherwise, if the value couldnt be retreived
59            GetBinaryValue = "Error" 'return Error to the user
60            If DisplayErrorMsg = True Then 'if the user wants to errors displayed
61               MsgBox ErrorMsg(rtn)  'display the error to the user
62            End If
63         End If
64      Else 'otherwise, if the key couldnt be opened
65         GetBinaryValue = "Error" 'return Error to the user
66         If DisplayErrorMsg = True Then 'if the user wants to errors displayed
67            MsgBox ErrorMsg(rtn)  'display the error to the user
68         End If
69      End If
70   End If

End Function
Function DeleteKey(Keyname As String)

71   Call ParseKey(Keyname, MainKeyHandle)

72   If MainKeyHandle Then
73      rtn = RegOpenKeyEx(MainKeyHandle, Keyname, 0, KEY_WRITE, hKey) 'open the key
74      If rtn = ERROR_SUCCESS Then 'if the key could be opened then
75         rtn = RegDeleteKey(hKey, Keyname) 'delete the key
76         rtn = RegCloseKey(hKey)  'close the key
77      End If
78   End If

End Function

Function GetMainKeyHandle(MainKeyName As String) As Long

79   Const HKEY_CLASSES_ROOT = &H80000000
80   Const HKEY_CURRENT_USER = &H80000001
81   Const HKEY_LOCAL_MACHINE = &H80000002
82   Const HKEY_USERS = &H80000003
83   Const HKEY_PERFORMANCE_DATA = &H80000004
84   Const HKEY_CURRENT_CONFIG = &H80000005
85   Const HKEY_DYN_DATA = &H80000006
   
Select Case MainKeyName
       Case "HKEY_CLASSES_ROOT"
86               GetMainKeyHandle = HKEY_CLASSES_ROOT
87          Case "HKEY_CURRENT_USER"
88               GetMainKeyHandle = HKEY_CURRENT_USER
89          Case "HKEY_LOCAL_MACHINE"
90               GetMainKeyHandle = HKEY_LOCAL_MACHINE
91          Case "HKEY_USERS"
92               GetMainKeyHandle = HKEY_USERS
93          Case "HKEY_PERFORMANCE_DATA"
94               GetMainKeyHandle = HKEY_PERFORMANCE_DATA
95          Case "HKEY_CURRENT_CONFIG"
96               GetMainKeyHandle = HKEY_CURRENT_CONFIG
97          Case "HKEY_DYN_DATA"
98               GetMainKeyHandle = HKEY_DYN_DATA
99   End Select

End Function

Function ErrorMsg(lErrorCode As Long) As String
    
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
       Case 1009, 1015
100              ErrorMsg = "The Registry Database is corrupt!"
101         Case 2, 1010
102              ErrorMsg = "Bad Key Name"
103         Case 1011
104              ErrorMsg = "Can't Open Key"
105         Case 4, 1012
106              ErrorMsg = "Can't Read Key"
107         Case 5
108              ErrorMsg = "Access to this key is denied"
109         Case 1013
110              ErrorMsg = "Can't Write Key"
111         Case 8, 14
112              ErrorMsg = "Out of memory"
113         Case 87
114              ErrorMsg = "Invalid Parameter"
115         Case 234
116              ErrorMsg = "There is more data than the buffer has been allocated to hold."
117         Case Else
118              ErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
119  End Select

End Function



Function GetStringValue(SubKey As String, Entry As String)
120      Call ParseKey(SubKey, MainKeyHandle)
    
121      If MainKeyHandle Then
122         rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the Key
123         If rtn = ERROR_SUCCESS Then 'if the key could be opened then
124            sBuffer = Space(255)     'make a buffer
125            lBufferSize = Len(sBuffer)
126            rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize) 'get the value from the registry
127            If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
128               rtn = RegCloseKey(hKey)  'close the key
129               sBuffer = Trim(sBuffer)
130               GetStringValue = Trim(Left$(sBuffer, lBufferSize - 1))
131            Else                        'otherwise, if the value couldnt be retreived
132               GetStringValue = "Error" 'return Error to the user
133               If DisplayErrorMsg = True Then 'if the user wants errors displayed then
134                  MsgBox ErrorMsg(rtn)  'tell the user what was wrong
135               End If
136            End If
137         Else 'otherwise, if the key couldnt be opened
138            GetStringValue = "Error"       'return Error to the user
139            If DisplayErrorMsg = True Then 'if the user wants errors displayed then
140               MsgBox ErrorMsg(rtn)        'tell the user what was wrong
141            End If
142         End If
143      End If
End Function


Private Sub ParseKey(Keyname As String, Keyhandle As Long)
    
144  rtn = InStr(Keyname, "\") 'return if "\" is contained in the Keyname

145  If Left$(Keyname, 5) <> "HKEY_" Or Right$(Keyname, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
146     MsgBox "Incorrect Format:" + vbLf + vbLf + Keyname 'display error to the user
147     Exit Sub 'exit the procedure
148  ElseIf rtn = 0 Then 'if the Keyname contains no "\"
149     Keyhandle = GetMainKeyHandle(Keyname)
150     Keyname = "" 'leave Keyname blank
151  Else 'otherwise, Keyname contains "\"
152     Keyhandle = GetMainKeyHandle(Left$(Keyname, rtn - 1)) 'seperate the Keyname
153     Keyname = Right$(Keyname, Len(Keyname) - rtn)
154  End If

End Sub
Function CreateKey(SubKey As String)

155  Call ParseKey(SubKey, MainKeyHandle)

156  If MainKeyHandle Then
157     rtn = RegCreateKey(MainKeyHandle, SubKey, hKey) 'create the key
158     If rtn = ERROR_SUCCESS Then 'if the key was created then
159        rtn = RegCloseKey(hKey)  'close the key
160     End If
161  End If

End Function
Function SetStringValue(SubKey As String, Entry As String, Value As String)

162  Call ParseKey(SubKey, MainKeyHandle)

163  If MainKeyHandle Then
164     rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
165     If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
166        rtn = RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value)) 'write the value
167        If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
168           If DisplayErrorMsg = True Then 'if the user wants errors displayed
169              MsgBox ErrorMsg(rtn)        'display the error
170           End If
171        End If
172        rtn = RegCloseKey(hKey) 'close the key
173     Else 'if there was an error opening the key
174        If DisplayErrorMsg = True Then 'if the user wants errors displayed
175           MsgBox ErrorMsg(rtn)        'display the error
176        End If
177     End If
178  End If

End Function

Public Function sadDecrypt(strIn As String) As String
179      Dim strOut As String
180      If Len(strIn) = 0 Then Exit Function
181      If Left$(strIn, 3) <> "EN*" Then Exit Function

182      strIn = Scramble(strIn)
183      Do While Len(strIn)
184         strOut = strOut & Chr$((255 - Val("&H" & Left$(strIn, 2) & "&")) Mod 255)
185         strIn = Mid$(strIn, 3)
186      Loop
187      sadDecrypt = strOut
End Function

Public Function sadEncrypt(ByVal strIn As String) As String
188      Dim strOut As String
189      Dim bytArray() As Byte
190      Dim CurrByte As Long

191      bytArray = StrConv(strIn, vbFromUnicode)
192      For CurrByte = 0 To UBound(bytArray)
193          If bytArray(CurrByte) < 240 Then
194             strOut = strOut & Hex$(255 - bytArray(CurrByte))
195          Else
196             strOut = strOut & "0" & Hex$(255 - bytArray(CurrByte))
197          End If
198      Next CurrByte

199      sadEncrypt = "EN* " & Scramble(strOut)
End Function


Public Sub sadSaveLicenseKey(ByVal Key As String, ByVal Value As String)
200  On Error Resume Next
201      SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License", Key, sadEncrypt(Value)
End Sub

Public Function Scramble(ByVal strIn As String) As String
202      Dim strOut As String
203      Dim bytArray() As Byte
204      Dim CurrByte As Long
205      Dim bytStack As Byte
206      Dim Shift As Integer
207      Dim MaxCount As Integer

208      If Left$(strIn, 4) = "EN* " Then
    '   Shift = -3
209         strIn = Mid$(strIn, 5)
    'Else
    '   Shift = 3
210      End If

211      bytArray = strIn
212      MaxCount = UBound(bytArray)
213      MaxCount = MaxCount - (MaxCount Mod 2) - 8
214         For CurrByte = 0 To MaxCount Step 8 'Step 8
215             bytStack = bytArray(CurrByte + 0)
216             bytArray(CurrByte + 0) = bytArray(CurrByte + 6)
217             bytArray(CurrByte + 6) = bytStack
218             bytStack = bytArray(CurrByte + 2)
219             bytArray(CurrByte + 2) = bytArray(CurrByte + 4)
220             bytArray(CurrByte + 4) = bytStack
221         Next CurrByte
222      strOut = bytArray
223      Scramble = strOut
End Function

Public Function GetListIndex(cboToSearch As Control, ByVal sItemToFind As String) As Integer
224  On Error Resume Next
225      Static nCurItem As Integer

226      If cboToSearch Is Nothing Then Exit Function

227      If Len(sItemToFind) = 0 Or cboToSearch.ListCount = 0 Then
228         GetListIndex = -1
229         Exit Function
230      End If

231      sItemToFind = UCase$(sItemToFind)

232      For nCurItem = 0 To cboToSearch.ListCount - 1
233          If StrComp(UCase$(cboToSearch.List(nCurItem)), sItemToFind) = 0 Then
234             GetListIndex = nCurItem
235             Exit Function
236          End If
237      Next nCurItem
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
238      Static strIn As String
239      Static strOut As String
240      Static nCurrTokenStart As Long
241      Static nNextTokenStart As Long
242      Static nLenDelim As Long

  ' Handle the "simple" cases (No delimiter, or token # less than 2)
243      nLenDelim = Len(strDelim)
244      If nToken < 1 Or nLenDelim = 0 Then
     ' Nothing to extract, return nothing
245         Exit Function
246      ElseIf nToken = 1 Then
247         nCurrTokenStart = InStr(sOrigStr, strDelim)
248         If nCurrTokenStart > 0 Then
249            sExtractToken = Left$(sOrigStr, nCurrTokenStart - 1)
250            sOrigStr = Trim(Mid$(sOrigStr, nCurrTokenStart + nLenDelim))
251            Exit Function
252         Else
253            sExtractToken = sOrigStr
254            sOrigStr = ""
255            Exit Function
256         End If
257      End If

  ' Find the start of then nToken'th Token
258      strIn = sOrigStr: strOut = ""
259      nToken = nToken - 1
260      Do Until nToken = 0
261         nCurrTokenStart = InStr(strIn, strDelim)
262         If nCurrTokenStart = 0 Or Len(strIn) = 0 Then Exit Function
263         strOut = strOut & Left$(strIn, nCurrTokenStart - 1)
264         strIn = Mid$(strIn, nCurrTokenStart + nLenDelim)

     ' Check to see if this is the one the calling function is looking for
265         nToken = nToken - 1
266      Loop

  ' Now we're at the point" & gsWhere & "the token sought for resides
267      nCurrTokenStart = InStr(strIn, strDelim)
268      If nCurrTokenStart > 0 Then
269         If nCurrTokenStart > 1 Then
270            sExtractToken = Left$(strIn, nCurrTokenStart - 1)
271         Else
272            sExtractToken = ""
273         End If
     ' Rewrite the original string without the last token
274         sOrigStr = Trim(strOut & Mid$(strIn, nCurrTokenStart))
275         Exit Function
276      Else
277         sExtractToken = strIn
278         sOrigStr = Trim(strOut)
279         Exit Function
280      End If
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
281      bUserSure = (MsgBox(sPrompt, vbYesNo, "ARE YOU SURE ?") = vbYes)
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
Public Function lTokenCount(ByVal siAllTokens As String, Optional ByVal siDelim As String = " ") As Long
282      Static iCurTokenLocation As Long ' Character position of the first delimiter string
283      Static iTokensSoFar As Long      ' Used to keep track of how many tokens we've counted so far
284      Static iDelim As Long            ' Length of the delimiter string

285      iDelim = Len(siDelim)
286      If iDelim < 1 Then
     ' Empty delimiter strings means only one token equal to the string
287         lTokenCount = 1
288         Exit Function
289      ElseIf Len(siAllTokens) = 0 Then
     ' Empty input string means no tokens
290         Exit Function
291      Else
     ' Count the number of tokens
292         iTokensSoFar = 0
293         Do
294            iCurTokenLocation = InStr(siAllTokens, siDelim)
295            If iCurTokenLocation = 0 Then
296               lTokenCount = iTokensSoFar + 1 'Abs(Len(siAllTokens) > 0)
297               Exit Function
298            End If
299            iTokensSoFar = iTokensSoFar + 1
300            siAllTokens = Mid$(siAllTokens, iCurTokenLocation + iDelim)
301         Loop
302      End If
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
303      Static iCurTokenLocation As Long ' Character position of the first delimiter string
304      Static nDelim As Long            ' Length of the delimiter string
305      nDelim = Len(sDelim)

306      If iToken < 1 Or nDelim < 1 Then
     ' Negative or zeroth token or empty delimiter strings mean an empty token
307         Exit Function
308      ElseIf iToken = 1 Then
     ' Quickly extract the first token
309         iCurTokenLocation = InStr(siAllTokens, sDelim)
310         If iCurTokenLocation > 1 Then
311            sGetToken = Left$(siAllTokens, iCurTokenLocation - 1)
312         ElseIf iCurTokenLocation = 1 Then
313            sGetToken = ""
314         Else
315            sGetToken = siAllTokens
316         End If
317         Exit Function
318      Else
     ' Find the Nth token
319         Do
320            iCurTokenLocation = InStr(siAllTokens, sDelim)
321            If iCurTokenLocation = 0 Then
322               Exit Function
323            Else
324               siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
325            End If
326            iToken = iToken - 1
327         Loop Until iToken = 1

     ' Extract the Nth token (Which is the next token at this point)
328         iCurTokenLocation = InStr(siAllTokens, sDelim)
329         If iCurTokenLocation > 0 Then
330            sGetToken = Left$(siAllTokens, iCurTokenLocation - 1)
331            Exit Function
332         Else
333            sGetToken = siAllTokens
334            Exit Function
335         End If
336      End If
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
337      Static iCurTokenLocation As Long ' Character position of the first delimiter string
338      Static nDelim As Long            ' Length of the delimiter string
    
339      nDelim = Len(sDelim)
340      If iToken < 1 Or nDelim < 1 Then
     ' Negative or zeroth token or empty delimiter strings mean an empty token
341         sAfter = siAllTokens
342         Exit Function
343      ElseIf iToken = 1 Then
     ' Quickly extract the first token
344         iCurTokenLocation = InStr(siAllTokens, sDelim)
345         If iCurTokenLocation > 1 Then
346            sAfter = Mid$(siAllTokens, iCurTokenLocation + nDelim)
347            Exit Function
348         ElseIf iCurTokenLocation = 0 Then
349            sAfter = ""
350            Exit Function
351         Else
352            sAfter = Mid$(siAllTokens, nDelim + 1)
353            Exit Function
354         End If
355      Else
     ' Find the Nth token
356         Do
357            iCurTokenLocation = InStr(siAllTokens, sDelim)
358            If iCurTokenLocation = 0 Then
359               Exit Function
360            Else
361               siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
362            End If
363            iToken = iToken - 1
364         Loop Until iToken = 1

     ' Extract the Nth token (Which is the next token at this point)
365         iCurTokenLocation = InStr(siAllTokens, sDelim)
366         If iCurTokenLocation > 0 Then
367            sAfter = Mid$(siAllTokens, iCurTokenLocation + nDelim)
368            Exit Function
369         Else
370            Exit Function
371         End If
372      End If
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
373      Static iCurTokenLocation As Long ' Character position of the first delimiter string
374      Static nDelim As Long            ' Length of the delimiter string
375      Static sReturned As String

376      nDelim = Len(sDelim)
377      If iToken < 2 Or nDelim < 1 Then
     ' First, Zeroth, or Negative tokens or empty delimiter strings mean an empty string returned
378         sBefore = ""
379         Exit Function
380      ElseIf iToken = 2 Then
     ' Quickly extract the first token
381         sBefore = sGetToken(siAllTokens, 1, sDelim)
382         Exit Function
383      Else
     ' Find the Nth token
384         Do
385            iCurTokenLocation = InStr(siAllTokens, sDelim)
386            If iCurTokenLocation = 0 Or iToken = 1 Then
387               sBefore = sReturned
388               sReturned = ""
389               Exit Function
390            ElseIf Len(sReturned) = 0 Then
391               sReturned = Left$(siAllTokens, iCurTokenLocation - 1)
392            Else
393               sReturned = sReturned & sDelim & Left$(siAllTokens, iCurTokenLocation - 1)
394            End If
395            siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
396            iToken = iToken - 1
397         Loop
398      End If
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
399      Static iCurTokenLocation As Long ' Character position of the first delimiter string
400      Static nDelim As Long            ' Length of the delimiter string
401      Static sReturned As String

402      nDelim = Len(sDelim)
403      If iToken < 1 Or nDelim < 1 Then
     ' First, Zeroth, or Negative tokens or empty delimiter strings mean an empty string returned
404         sExcept = siAllTokens
405         Exit Function
406      ElseIf iToken = 1 Then
     ' Quickly Return after token 1
407         iCurTokenLocation = InStr(siAllTokens, sDelim)
408         If iCurTokenLocation = 0 Then
409            sExcept = siAllTokens
410            Exit Function
411         Else
412            sExcept = Mid$(siAllTokens, iCurTokenLocation + nDelim)
413            Exit Function
414         End If
415      Else
     ' Find the Nth token
416         Do
417            iCurTokenLocation = InStr(siAllTokens, sDelim)
418            If iToken = 1 Then
419               If iCurTokenLocation > 0 Then
420                  sExcept = sReturned & sDelim & Mid$(siAllTokens, iCurTokenLocation + nDelim)
421               Else
422                  sExcept = sReturned
423               End If
424               sReturned = ""
425               Exit Function
426            ElseIf iCurTokenLocation = 0 Then
427               sExcept = sReturned & sDelim & siAllTokens
428               sReturned = ""
429               Exit Function
430            ElseIf Len(sReturned) = 0 Then
431               sReturned = Left$(siAllTokens, iCurTokenLocation - 1)
432            Else
433               sReturned = sReturned & sDelim & Left$(siAllTokens, iCurTokenLocation - 1)
434            End If
435            siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
436            iToken = iToken - 1
437         Loop
438      End If
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
439      Dim iCurFindPos As Long
440      Dim iFind As Long
441      Dim sOut As String

442      iFind = Len(sFind)
443      iCurFindPos = InStr(sAll, sFind)
444      If InStr(sReplaceWith, sFind) = 0 Then
445         Do While iCurFindPos > 0
446            If iCurFindPos > 1 Then
447               sAll = Left$(sAll, iCurFindPos - 1) & sReplaceWith & Mid$(sAll, iCurFindPos + iFind)
448            Else
449               sAll = sReplaceWith & Mid$(sAll, iCurFindPos + iFind)
450            End If
451            iCurFindPos = InStr(sAll, sFind)
452         Loop
453         sReplace = sAll
454      Else
455         Do While iCurFindPos > 0
456            If iCurFindPos > 1 Then
457               sOut = sOut & Left$(sAll, iCurFindPos - 1) & sReplaceWith
458               sAll = Mid$(sAll, iCurFindPos + iFind)
459            Else
460               sOut = sOut & sReplaceWith
461               sAll = Mid$(sAll, iCurFindPos + iFind)
462            End If
463            iCurFindPos = InStr(sAll, sFind)
464         Loop
       
465         sReplace = sOut & sAll
466      End If
End Function

Public Function LoadFormPosition(frmToActOn As Form, Optional ByVal bAutoCenter = True, Optional ByVal bRemeberWidth As Boolean = True, Optional ByVal sSectionName As String)
467      Dim ProductName As String
468      Dim SectionName As String
    
469      If Len(ProductName) = 0 Then
470         ProductName = "Slice and Dice"
471      Else
472         ProductName = App.ProductName
473      End If
    
474      With frmToActOn
475           If Len(sSectionName) Then
476              SectionName = sSectionName
477           Else
478              SectionName = .Name
479           End If
480           If GetSetting(ProductName, SectionName, "Position Saved", False) Then
481              .Left = GetSetting(ProductName, SectionName, "Form Position Left", .Left)
482              .Top = GetSetting(ProductName, SectionName, "Form Position Top", .Top)
483              If bRemeberWidth Then .Width = GetSetting(ProductName, SectionName, "Form Position Width", .Width)
484              .Height = GetSetting(ProductName, SectionName, "Form Position Height", .Height)
485           ElseIf bAutoCenter Then
486              .Left = (Screen.Width - .Width) / 2
487              .Top = (Screen.Height - .Height) / 2
488           End If

       ' Ensure it'll fit on the screen (screen resolution change ?)
489           If .Left > Screen.Width Then .Left = 0
490           If .Top > Screen.Height Then .Top = 0
491           If .Left + .Width > Screen.Width Then .Width = Screen.Width - .Left
492           If bRemeberWidth Then
493              If .Top + .Height > Screen.Height Then .Height = Screen.Height - .Top
494           End If
495      End With
End Function

 
Public Function SaveFormPosition(frmToActOn As Form, Optional ByVal sSectionName As String)
496      Dim ProductName As String
497      Dim SectionName As String
    
498      If Len(ProductName) = 0 Then
499         ProductName = "Slice and Dice"
500      Else
501         ProductName = App.ProductName
502      End If

503      With frmToActOn
504           If Len(sSectionName) Then
505              SectionName = sSectionName
506           Else
507              SectionName = .Name
508           End If
509           SaveSetting ProductName, SectionName, "Position Saved", True
510           SaveSetting ProductName, SectionName, "Form Position Left", .Left
511           SaveSetting ProductName, SectionName, "Form Position Top", .Top
512           SaveSetting ProductName, SectionName, "Form Position Width", .Width
513           SaveSetting ProductName, SectionName, "Form Position Height", .Height
514      End With
End Function

Public Function lFindToken(ByVal sAllTokens As String, ByVal sTokenToFind As String, Optional ByVal sDelimiter As String = " ") As Long
515      Dim lTokens As Long
516      Dim l As Long

517      lTokens = lTokenCount(sAllTokens, sDelimiter)

518      For l = 1 To lTokens
519          If StrComp(UCase$(sGetToken(sAllTokens, l, sDelimiter)), UCase$(sTokenToFind)) = 0 Then
520             lFindToken = l
521             Exit Function
522          End If
523      Next l

524      lFindToken = 0
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
    Dim CheckSum    As Long
    Const Offsets   As String = "594621357894651258468953267945648551234565648965410368541687416841654106541654165416541654165719684168413841684"
    'Const Offsets   As String = "615243516784259045218002180248620684102579462315787815795168911248961534896127811596154329617581123589402160548"
    Const Values    As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
    Static LastRnd  As Long

'SOMETHING
'YUF

    If Len(sInvoiceNumber) = 0 Then Exit Function

    ValueLen = Len(Values)
    OffsetLen = Len(Offsets)

    sInvoiceNumber = "V" & UCase$(sInvoiceNumber) & "Z"
    bytArray = StrConv(sInvoiceNumber, vbFromUnicode)

    Do
        strOut = ""
        Randomize Now + Rnd
        Do
            StartAt = Int(65 * Rnd + 1)
        Loop While StartAt = LastRnd
        LastRnd = StartAt
    
        CurrOffset = StartAt
    
        strOut = vbNullString
    
        For CurrByte = 0 To UBound(bytArray)
            CheckSum = (CheckSum + bytArray(CurrByte)) Mod ValueLen
            CharLoc = InStr(Values, Chr(bytArray(CurrByte)))
            If CharLoc < 1 Or CharLoc > ValueLen Then Exit Function
            CharLoc = CharLoc + Val(Mid$(Offsets, CurrOffset, 1))
            If CharLoc > ValueLen Then CharLoc = CharLoc - ValueLen
            strOut = strOut & Mid$(Values, CharLoc, 1)
            'If ((CurrByte Mod 6) = 2) Then strOut = strOut & "-"
            CurrOffset = CurrOffset + 1
            If CurrOffset > OffsetLen Then CurrOffset = 1
        Next CurrByte
    
        If CheckSum < 1 Then CheckSum = 1
        strOut = Right$(Format(StartAt, "00"), 1) & strOut & Mid$(Values, CheckSum, 1) & Left$(Format(StartAt, "00"), 1)
        strOut = Replace$(strOut, "O", ".")
    Loop Until InStr(strOut, ".") = 0
    sadInvoiceEncrypt = strOut
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
    Dim CheckSum    As Long
    Dim CheckValue  As Long
    Const Offsets   As String = "594621357894651258468953267945648551234565648965410368541687416841654106541654165416541654165719684168413841684"
    Const Values    As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"

    If Len(sInvoiceNumber) = 0 Then Exit Function

    ValueLen = Len(Values)
    OffsetLen = Len(Offsets)
    sInvoiceNumber = Replace$(Replace$(Replace$(UCase$(sInvoiceNumber), "O", "0"), ".", "O"), "-", vbNullString)
    StartAt = Val(Right$(sInvoiceNumber, 1) & Left$(sInvoiceNumber, 1))
    sInvoiceNumber = Mid$(sInvoiceNumber, 2, Len(sInvoiceNumber) - 2)
    CheckValue = InStr(Values, Right$(sInvoiceNumber, 1))
    sInvoiceNumber = Left$(sInvoiceNumber, Len(sInvoiceNumber) - 1)
    CurrOffset = StartAt
    bytArray = StrConv(sInvoiceNumber, vbFromUnicode)
    strOut = vbNullString

    For CurrByte = 0 To UBound(bytArray)
        CharLoc = InStr(Values, Chr(bytArray(CurrByte)))
        If CharLoc < 1 Or CharLoc > ValueLen Then Exit Function
        CharLoc = CharLoc - Val(Mid$(Offsets, CurrOffset, 1))
        If CharLoc < 1 Then CharLoc = ValueLen + CharLoc
        CheckSum = (CheckSum + Asc(Mid$(Values, CharLoc, 1))) Mod ValueLen
        strOut = strOut & Mid$(Values, CharLoc, 1)
        CurrOffset = CurrOffset + 1
        If CurrOffset > OffsetLen Then CurrOffset = 1
    Next CurrByte

    If Left$(strOut, 1) = "V" And Right$(strOut, 1) = "Z" Then
       If CheckSum < 1 Then CheckSum = 1
       strOut = Mid$(strOut, 2, Len(strOut) - 2)
       If CheckSum = CheckValue Then
          sadInvoiceDecrypt = strOut
       Else
          sadInvoiceDecrypt = vbNullString
       End If
    Else
       sadInvoiceDecrypt = vbNullString
    End If
End Function



