Attribute VB_Name = "Registry"
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


Public Sub sadSaveLicenseKey(ByVal Key As String, ByVal Value As String)
1        On Error Resume Next
2        SetStringValue "HKEY_LOCAL_MACHINE" & gsBS & "SOFTWARE" & gsBS & "Zion Systems" & gsBS & "License", Key, sadEncrypt(Value)
End Sub

Public Function sadGetLicenseKey(ByVal Key As String, Optional ByVal sDefault As String) As String
3        On Error Resume Next
4        Dim sResult As String
5        sResult = GetStringValue("HKEY_LOCAL_MACHINE" & gsBS & "SOFTWARE" & gsBS & "Zion Systems" & gsBS & "License", Key)
6        If Len(sResult) = 0 Or sResult = "Error" Then sResult = sDefault
7        If Left$(sResult, 4) = "EN* " Then
8            sResult = sadDecrypt(sResult)
9        End If
10       sadGetLicenseKey = sResult
End Function


Function SetDWORDValue(SubKey As String, Entry As String, Value As Long)

11       Call ParseKey(SubKey, MainKeyHandle)

12       If MainKeyHandle Then
13           rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey)    'open the key
14           If rtn = ERROR_SUCCESS Then                   'if the key was open successfully then
15               rtn = RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4)    'write the value
16               If Not rtn = ERROR_SUCCESS Then           'if there was an error writting the value
17                   If DisplayErrorMsg = True Then        'if the user want errors displayed
18                       MsgBox ErrorMsg(rtn)              'display the error
19                   End If
20               End If
21               rtn = RegCloseKey(hKey)                   'close the key
22           Else                                          'if there was an error opening the key
23               If DisplayErrorMsg = True Then            'if the user want errors displayed
24                   MsgBox ErrorMsg(rtn)                  'display the error
25               End If
26           End If
27       End If

End Function
Function GetDWORDValue(SubKey As String, Entry As String)

28       Call ParseKey(SubKey, MainKeyHandle)

29       If MainKeyHandle Then
30           rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey)    'open the key
31           If rtn = ERROR_SUCCESS Then                   'if the key could be opened then
32               rtn = RegQueryValueExA(hKey, Entry, 0, REG_DWORD, lBuffer, 4)    'get the value from the registry
33               If rtn = ERROR_SUCCESS Then               'if the value could be retreived then
34                   rtn = RegCloseKey(hKey)               'close the key
35                   GetDWORDValue = lBuffer               'return the value
36               Else                                      'otherwise, if the value couldnt be retreived
37                   GetDWORDValue = "Error"               'return Error to the user
38                   If DisplayErrorMsg = True Then        'if the user wants errors displayed
39                       MsgBox ErrorMsg(rtn)              'tell the user what was wrong
40                   End If
41               End If
42           Else                                          'otherwise, if the key couldnt be opened
43               GetDWORDValue = "Error"                   'return Error to the user
44               If DisplayErrorMsg = True Then            'if the user wants errors displayed
45                   MsgBox ErrorMsg(rtn)                  'tell the user what was wrong
46               End If
47           End If
48       End If

End Function
Function SetBinaryValue(SubKey As String, Entry As String, Value As String)

49       Call ParseKey(SubKey, MainKeyHandle)

50       If MainKeyHandle Then
51           rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey)    'open the key
52           If rtn = ERROR_SUCCESS Then                   'if the key was open successfully then
53               lDataSize = Len(Value)
54               ReDim ByteArray(lDataSize)
55               For i = 1 To lDataSize
56                   ByteArray(i) = Asc(Mid$(Value, i, 1))
57               Next
58               rtn = RegSetValueExB(hKey, Entry, 0, REG_BINARY, ByteArray(1), lDataSize)    'write the value
59               If Not rtn = ERROR_SUCCESS Then           'if the was an error writting the value
60                   If DisplayErrorMsg = True Then        'if the user want errors displayed
61                       MsgBox ErrorMsg(rtn)              'display the error
62                   End If
63               End If
64               rtn = RegCloseKey(hKey)                   'close the key
65           Else                                          'if there was an error opening the key
66               If DisplayErrorMsg = True Then            'if the user wants errors displayed
67                   MsgBox ErrorMsg(rtn)                  'display the error
68               End If
69           End If
70       End If

End Function


Function GetBinaryValue(SubKey As String, Entry As String)

71       Call ParseKey(SubKey, MainKeyHandle)

72       If MainKeyHandle Then
73           rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey)    'open the key
74           If rtn = ERROR_SUCCESS Then                   'if the key could be opened
75               lBufferSize = 1
76               rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, 0, lBufferSize)    'get the value from the registry
77               sBuffer = Space$(lBufferSize)
78               rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize)    'get the value from the registry
79               If rtn = ERROR_SUCCESS Then               'if the value could be retreived then
80                   rtn = RegCloseKey(hKey)               'close the key
81                   GetBinaryValue = sBuffer              'return the value to the user
82               Else                                      'otherwise, if the value couldnt be retreived
83                   GetBinaryValue = "Error"              'return Error to the user
84                   If DisplayErrorMsg = True Then        'if the user wants to errors displayed
85                       MsgBox ErrorMsg(rtn)              'display the error to the user
86                   End If
87               End If
88           Else                                          'otherwise, if the key couldnt be opened
89               GetBinaryValue = "Error"                  'return Error to the user
90               If DisplayErrorMsg = True Then            'if the user wants to errors displayed
91                   MsgBox ErrorMsg(rtn)                  'display the error to the user
92               End If
93           End If
94       End If

End Function
Function DeleteKey(Keyname As String)

95       Call ParseKey(Keyname, MainKeyHandle)

96       If MainKeyHandle Then
97           rtn = RegOpenKeyEx(MainKeyHandle, Keyname, 0, KEY_WRITE, hKey)    'open the key
98           If rtn = ERROR_SUCCESS Then                   'if the key could be opened then
99               rtn = RegDeleteKey(hKey, Keyname)         'delete the key
100              rtn = RegCloseKey(hKey)                   'close the key
101          End If
102      End If

End Function

Function GetMainKeyHandle(MainKeyName As String) As Long

103      Const HKEY_CLASSES_ROOT = &H80000000
104      Const HKEY_CURRENT_USER = &H80000001
105      Const HKEY_LOCAL_MACHINE = &H80000002
106      Const HKEY_USERS = &H80000003
107      Const HKEY_PERFORMANCE_DATA = &H80000004
108      Const HKEY_CURRENT_CONFIG = &H80000005
109      Const HKEY_DYN_DATA = &H80000006

    Select Case MainKeyName
        Case "HKEY_CLASSES_ROOT"
110              GetMainKeyHandle = HKEY_CLASSES_ROOT
111          Case "HKEY_CURRENT_USER"
112              GetMainKeyHandle = HKEY_CURRENT_USER
113          Case "HKEY_LOCAL_MACHINE"
114              GetMainKeyHandle = HKEY_LOCAL_MACHINE
115          Case "HKEY_USERS"
116              GetMainKeyHandle = HKEY_USERS
117          Case "HKEY_PERFORMANCE_DATA"
118              GetMainKeyHandle = HKEY_PERFORMANCE_DATA
119          Case "HKEY_CURRENT_CONFIG"
120              GetMainKeyHandle = HKEY_CURRENT_CONFIG
121          Case "HKEY_DYN_DATA"
122              GetMainKeyHandle = HKEY_DYN_DATA
123      End Select

End Function

Function ErrorMsg(lErrorCode As Long) As String

'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

    Select Case lErrorCode
        Case 1009, 1015
124              ErrorMsg = "The Registry Database is corrupt!"
125          Case 2, 1010
126              ErrorMsg = "Bad Key Name"
127          Case 1011
128              ErrorMsg = "Can't Open Key"
129          Case 4, 1012
130              ErrorMsg = "Can't Read Key"
131          Case 5
132              ErrorMsg = "Access to this key is denied"
133          Case 1013
134              ErrorMsg = "Can't Write Key"
135          Case 8, 14
136              ErrorMsg = "Out of memory"
137          Case 87
138              ErrorMsg = "Invalid Parameter"
139          Case 234
140              ErrorMsg = "There is more data than the buffer has been allocated to hold."
141          Case Else
142              ErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
143      End Select

End Function



Function GetStringValue(SubKey As String, Entry As String)
144      Call ParseKey(SubKey, MainKeyHandle)

145      If MainKeyHandle Then
146          rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey)    'open the Key
147          If rtn = ERROR_SUCCESS Then                   'if the key could be opened then
148              sBuffer = Space$(255)                     'make a buffer
149              lBufferSize = Len(sBuffer)
150              rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize)    'get the value from the registry
151              If rtn = ERROR_SUCCESS Then               'if the value could be retreived then
152                  rtn = RegCloseKey(hKey)               'close the key
153                  sBuffer = Trim$(sBuffer)
154                  GetStringValue = Trim$(Left$(sBuffer, lBufferSize - 1))
155              Else                                      'otherwise, if the value couldnt be retreived
156                  GetStringValue = "Error"              'return Error to the user
157                  If DisplayErrorMsg = True Then        'if the user wants errors displayed then
158                      MsgBox ErrorMsg(rtn)              'tell the user what was wrong
159                  End If
160              End If
161          Else                                          'otherwise, if the key couldnt be opened
162              GetStringValue = "Error"                  'return Error to the user
163              If DisplayErrorMsg = True Then            'if the user wants errors displayed then
164                  MsgBox ErrorMsg(rtn)                  'tell the user what was wrong
165              End If
166          End If
167      End If
End Function

Private Sub ParseKey(Keyname As String, Keyhandle As Long)

168      rtn = InStr(Keyname, gsBS)                        'return if gsBS is contained in the Keyname

169      If Left$(Keyname, 5) <> "HKEY_" Or Right$(Keyname, 1) = gsBS Then    'if the is a gsBS at the end of the Keyname then
170          MsgBox "Incorrect Format:" & gs2EOL & Keyname 'display error to the user
171          Exit Sub                                      'exit the procedure
172      ElseIf rtn = 0 Then                               'if the Keyname contains no gsBS
173          Keyhandle = GetMainKeyHandle(Keyname)
174          Keyname = vbNullString                        'leave Keyname blank
175      Else                                              'otherwise, Keyname contains gsBS
176          Keyhandle = GetMainKeyHandle(Left$(Keyname, rtn - 1))    'seperate the Keyname
177          Keyname = Right$(Keyname, Len(Keyname) - rtn)
178      End If

End Sub
Function CreateKey(SubKey As String)

179      Call ParseKey(SubKey, MainKeyHandle)

180      If MainKeyHandle Then
181          rtn = RegCreateKey(MainKeyHandle, SubKey, hKey)    'create the key
182          If rtn = ERROR_SUCCESS Then                   'if the key was created then
183              rtn = RegCloseKey(hKey)                   'close the key
184          End If
185      End If

End Function
Function SetStringValue(SubKey As String, Entry As String, Value As String)
186      Dim OrigKey As String
187      OrigKey = SubKey
188      Call ParseKey(SubKey, MainKeyHandle)

189      If MainKeyHandle Then
190          rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey)    'open the key
191          If rtn = 2 Then
192              CreateKey OrigKey
193              rtn = ERROR_SUCCESS
194          End If
195          If rtn = ERROR_SUCCESS Then                   'if the key was open successfully then
196              rtn = RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value))    'write the value
197              If Not rtn = ERROR_SUCCESS Then           'if there was an error writting the value
198                  If DisplayErrorMsg = True Then        'if the user wants errors displayed
199                      MsgBox ErrorMsg(rtn)              'display the error
200                  End If
201              End If
202              rtn = RegCloseKey(hKey)                   'close the key
203          Else                                          'if there was an error opening the key
204              If DisplayErrorMsg = True Then            'if the user wants errors displayed
205                  MsgBox ErrorMsg(rtn)                  'display the error
206              End If
207          End If
208      End If
End Function

