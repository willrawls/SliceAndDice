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
On Error Resume Next
    Dim sResult As String
    sResult = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License", Key)
    If Len(sResult) = 0 Or sResult = "Error" Then sResult = sDefault
    If Left(sResult, 4) = "EN* " Then
       sResult = sadDecrypt(sResult)
    End If
    sadGetLicenseKey = sResult
End Function

Function SetDWORDValue(SubKey As String, Entry As String, Value As Long)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
         If DisplayErrorMsg = True Then 'if the user want errors displayed
            MsgBox ErrorMsg(rtn)        'display the error
         End If
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user want errors displayed
         MsgBox ErrorMsg(rtn) 'display the error
      End If
   End If
End If

End Function
Function GetDWORDValue(SubKey As String, Entry As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = RegQueryValueExA(hKey, Entry, 0, REG_DWORD, lBuffer, 4) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         GetDWORDValue = lBuffer  'return the value
      Else                        'otherwise, if the value couldnt be retreived
         GetDWORDValue = "Error"  'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants errors displayed
            MsgBox ErrorMsg(rtn)        'tell the user what was wrong
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetDWORDValue = "Error"        'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         MsgBox ErrorMsg(rtn)        'tell the user what was wrong
      End If
   End If
End If

End Function
Function GetBinaryValue(SubKey As String, Entry As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened
      lBufferSize = 1
      rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, 0, lBufferSize) 'get the value from the registry
      sBuffer = Space(lBufferSize)
      rtn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         GetBinaryValue = sBuffer 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         GetBinaryValue = "Error" 'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants to errors displayed
            MsgBox ErrorMsg(rtn)  'display the error to the user
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetBinaryValue = "Error" 'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants to errors displayed
         MsgBox ErrorMsg(rtn)  'display the error to the user
      End If
   End If
End If

End Function
Function DeleteKey(Keyname As String)

Call ParseKey(Keyname, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, Keyname, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = RegDeleteKey(hKey, Keyname) 'delete the key
      rtn = RegCloseKey(hKey)  'close the key
   End If
End If

End Function

Function GetMainKeyHandle(MainKeyName As String) As Long

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
   
Select Case MainKeyName
       Case "HKEY_CLASSES_ROOT"
            GetMainKeyHandle = HKEY_CLASSES_ROOT
       Case "HKEY_CURRENT_USER"
            GetMainKeyHandle = HKEY_CURRENT_USER
       Case "HKEY_LOCAL_MACHINE"
            GetMainKeyHandle = HKEY_LOCAL_MACHINE
       Case "HKEY_USERS"
            GetMainKeyHandle = HKEY_USERS
       Case "HKEY_PERFORMANCE_DATA"
            GetMainKeyHandle = HKEY_PERFORMANCE_DATA
       Case "HKEY_CURRENT_CONFIG"
            GetMainKeyHandle = HKEY_CURRENT_CONFIG
       Case "HKEY_DYN_DATA"
            GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function

Function ErrorMsg(lErrorCode As Long) As String
    
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
       Case 1009, 1015
            ErrorMsg = "The Registry Database is corrupt!"
       Case 2, 1010
            ErrorMsg = "Bad Key Name"
       Case 1011
            ErrorMsg = "Can't Open Key"
       Case 4, 1012
            ErrorMsg = "Can't Read Key"
       Case 5
            ErrorMsg = "Access to this key is denied"
       Case 1013
            ErrorMsg = "Can't Write Key"
       Case 8, 14
            ErrorMsg = "Out of memory"
       Case 87
            ErrorMsg = "Invalid Parameter"
       Case 234
            ErrorMsg = "There is more data than the buffer has been allocated to hold."
       Case Else
            ErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
End Select

End Function



Function GetStringValue(SubKey As String, Entry As String)
    Call ParseKey(SubKey, MainKeyHandle)
    
    If MainKeyHandle Then
       rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the Key
       If rtn = ERROR_SUCCESS Then 'if the key could be opened then
          sBuffer = Space(255)     'make a buffer
          lBufferSize = Len(sBuffer)
          rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize) 'get the value from the registry
          If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
             rtn = RegCloseKey(hKey)  'close the key
             sBuffer = Trim(sBuffer)
             GetStringValue = Trim(Left(sBuffer, lBufferSize - 1))
          Else                        'otherwise, if the value couldnt be retreived
             GetStringValue = "Error" 'return Error to the user
             If DisplayErrorMsg = True Then 'if the user wants errors displayed then
                MsgBox ErrorMsg(rtn)  'tell the user what was wrong
             End If
          End If
       Else 'otherwise, if the key couldnt be opened
          GetStringValue = "Error"       'return Error to the user
          If DisplayErrorMsg = True Then 'if the user wants errors displayed then
             MsgBox ErrorMsg(rtn)        'tell the user what was wrong
          End If
       End If
    End If
End Function


Private Sub ParseKey(Keyname As String, Keyhandle As Long)
    
rtn = InStr(Keyname, "\") 'return if "\" is contained in the Keyname

If Left(Keyname, 5) <> "HKEY_" Or Right(Keyname, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + Keyname 'display error to the user
   Exit Sub 'exit the procedure
ElseIf rtn = 0 Then 'if the Keyname contains no "\"
   Keyhandle = GetMainKeyHandle(Keyname)
   Keyname = "" 'leave Keyname blank
Else 'otherwise, Keyname contains "\"
   Keyhandle = GetMainKeyHandle(Left(Keyname, rtn - 1)) 'seperate the Keyname
   Keyname = Right(Keyname, Len(Keyname) - rtn)
End If

End Sub
Function CreateKey(SubKey As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegCreateKey(MainKeyHandle, SubKey, hKey) 'create the key
   If rtn = ERROR_SUCCESS Then 'if the key was created then
      rtn = RegCloseKey(hKey)  'close the key
   End If
End If

End Function
Function SetStringValue(SubKey As String, Entry As String, Value As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value)) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
         If DisplayErrorMsg = True Then 'if the user wants errors displayed
            MsgBox ErrorMsg(rtn)        'display the error
         End If
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         MsgBox ErrorMsg(rtn)        'display the error
      End If
   End If
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


Public Sub sadSaveLicenseKey(ByVal Key As String, ByVal Value As String)
On Error Resume Next
    SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License", Key, sadEncrypt(Value)
End Sub

Public Function Scramble(ByVal strIn As String) As String
    Dim strOut As String
    Dim bytArray() As Byte
    Dim CurrByte As Long
    Dim bytStack As Byte
    Dim Shift As Integer
    Dim MaxCount As Integer

    If Left(strIn, 4) = "EN* " Then
    '   Shift = -3
       strIn = Mid(strIn, 5)
    'Else
    '   Shift = 3
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

Public Function FileExists(sFilename As String) As Boolean
On Error Resume Next
    Err.Clear
       FileExists = Len(Dir(sFilename)) > 0
    Err.Clear
End Function

Public Function GetListIndex(cboToSearch As Control, ByVal sItemToFind As String) As Integer
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

Public Function LogError(ByVal sModuleName As String, sProcName As String, lError As Long, sErrorMsg As String) As Boolean
    Dim fh As Long
    Dim sMessage As String

    fh = FreeFile
    Open "sadRegister.LOG" For Append As #fh
         sMessage = "***** Error " & Format(lError, "00000") & " at: " & Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM")
         sMessage = sMessage & Chr(13) & "  *** Module:         " & sModuleName
         sMessage = sMessage & Chr(13) & "  *** Procedure:      " & sProcName
         sMessage = sMessage & Chr(13) & "  *** Description:    " & sErrorMsg
         Print #fh, sMessage
         MsgBox sMessage
         Print #fh, "  *** Program continued by user after error."
    Close #fh

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
          sExtractToken = Left(sOrigStr, nCurrTokenStart - 1)
          sOrigStr = Trim(Mid(sOrigStr, nCurrTokenStart + nLenDelim))
          Exit Function
       Else
          sExtractToken = sOrigStr
          sOrigStr = ""
          Exit Function
       End If
    End If

  ' Find the start of then nToken'th Token
    strIn = sOrigStr: strOut = ""
    nToken = nToken - 1
    Do Until nToken = 0
       nCurrTokenStart = InStr(strIn, strDelim)
       If nCurrTokenStart = 0 Or Len(strIn) = 0 Then Exit Function
       strOut = strOut & Left(strIn, nCurrTokenStart - 1)
       strIn = Mid(strIn, nCurrTokenStart + nLenDelim)

     ' Check to see if this is the one the calling function is looking for
       nToken = nToken - 1
    Loop

  ' Now we're at the point" & gsWhere & "the token sought for resides
    nCurrTokenStart = InStr(strIn, strDelim)
    If nCurrTokenStart > 0 Then
       If nCurrTokenStart > 1 Then
          sExtractToken = Left(strIn, nCurrTokenStart - 1)
       Else
          sExtractToken = ""
       End If
     ' Rewrite the original string without the last token
       sOrigStr = Trim(strOut & Mid(strIn, nCurrTokenStart))
       Exit Function
    Else
       sExtractToken = strIn
       sOrigStr = Trim(strOut)
       Exit Function
    End If
End Function

