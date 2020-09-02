VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Easy access to Windows Registry"
Option Explicit

'*******************************************************************************
'
'   Name:           Setting.cls
'
'   Description:    Registry Setting class
'
'   Author:         Kenneth LeFebvre
'
'   ------------------------------------------------------------------------
'
'   PUBLIC PROPERTIES:
'       KeyPath                 RW: String, Full path to any key in registry
'                                   (supercedes Application, Section, and Key)
'       Application             RW: String, Application name (defaults to App.ProductName)
'       Section                 RW: String, Section name (defaults to "Recent")
'       Key                     RW: String, Key name (no default)
'       Default                 RW: String, Default value
'       DataType                RW: Integer, Type of value to read/write
'       Value                   RW: String, Current value in registry
'       AutoUpdate              RW: Boolean, Automatically rewrite defaults to registry?
'                                   (defaults to True)
'       Count                   RO: Integer, Number of values in section
'
'   PUBLIC METHODS:
'       GetAll                  Variant, Returns an array of all values in section
'
'*******************************************************************************
' Copyright (c) 1996, Kenneth LeFebvre
'*******************************************************************************

Private msKeyPath               As String
Private msApplication           As String
Private msSection               As String
Private msKey                   As String
Private msDefault               As String
Private msValue                 As String
Private mbAutoUpdate            As Boolean
Private miType                  As Integer

Private mlRootKey               As Long
Private msSubKey                As String
Private msValueName             As String

Private mhKey                As Long         'Handle to open key
Private msClass              As String       'Class name
Private mlClassLen           As Long         'Length of class name
Private mlNoSubKeys          As Long         'Number of subkeys
Private mlSubKeyMaxLen       As Long         'Length of longest subkey name
Private mlClassMaxLen        As Long         'Length of longest class name
Private mlNoValues           As Long         'Number of values
Private mlValueNameMaxLen    As Long         'Length of longest value name
Private mlValueMaxLen        As Long         'Length of longest value
Private mlSecurityLen        As Long         'Length of security descriptor
Private mftLastWrite         As FILETIME     'Last modification time


Public Sub LoadPosition(frmToPosition As Object, sSection As String)
    Dim frm As VB.Form
    Set frm = frmToPosition
    With frm
         If Application = vbNullString Then Application = App.ProductName
         Section = sSection
         Key = "Top":    .Top = IIf(Value = vbNullString, .Top, Value)
         Key = "Left":   .Left = IIf(Value = vbNullString, .Left, Value)
         Key = "Width":  .Width = IIf(Value = vbNullString, .Width, Value)
         Key = "Height": .Height = IIf(Value = vbNullString, .Height, Value)
    End With
    Set frm = Nothing
End Sub


Public Sub SavePosition(frmToPosition As Object, sSection As String)
    Dim frm As VB.Form
    Set frm = frmToPosition
    With frm
         If Application = vbNullString Then Application = App.ProductName
         Section = sSection
         Key = "Top":    Value = .Top
         Key = "Left":   Value = .Left
         Key = "Width":  Value = .Width
         Key = "Height": Value = .Height
    End With
    Set frm = Nothing
End Sub

Public Property Get Root() As String
    Select Case mlRootKey
           Case HKEY_CLASSES_ROOT:      Root = "HKEY_CLASSES_ROOT"
           Case HKEY_CURRENT_USER:      Root = "HKEY_CURRENT_USER"
           Case HKEY_LOCAL_MACHINE:     Root = "HKEY_LOCAL_MACHINE"
           Case HKEY_PERFORMANCE_DATA:  Root = "HKEY_PERFORMANCE_DATA"
           Case HKEY_USERS:             Root = "HKEY_USERS"
      '    Case HKEY_CURRENT_CONFIG:    Root = HKEY_CURRENT_CONFIG
      '    Case HKEY_DYN_DATA:          Root = HKEY_DYN_DATA
    End Select
End Property

Public Property Let Root(ByVal NewRoot As String)
    Dim iSlash As Integer

    iSlash = InStr(NewRoot, "\")

    Select Case UCase$(Left$(msKeyPath, iSlash - 1))
           Case "HKEY_CLASSES_ROOT":      mlRootKey = HKEY_CLASSES_ROOT
           Case "HKEY_CURRENT_USER":      mlRootKey = HKEY_CURRENT_USER
           Case "HKEY_LOCAL_MACHINE":     mlRootKey = HKEY_LOCAL_MACHINE
           Case "HKEY_PERFORMANCE_DATA":  mlRootKey = HKEY_PERFORMANCE_DATA
           Case "HKEY_USERS":             mlRootKey = HKEY_USERS
      '    Case "HKEY_CURRENT_CONFIG":    mlRootKey = HKEY_CURRENT_CONFIG
      '    Case "HKEY_DYN_DATA":          mlRootKey = HKEY_DYN_DATA
           Case Else:                     mlRootKey = HKEY_CLASSES_ROOT
    End Select
    
End Property

'*******************************************************************************
Private Sub Class_Initialize()

    Application = App.ProductName
    Section = "Recent"
    DataType = REG_SZ
    AutoUpdate = True

End Sub


'*******************************************************************************
'       KeyPath                 RW: String, Full path to any key in registry
'                                   (supercedes Application, Section, and Key)
Public Property Get KeyPath() As String

    KeyPath = msKeyPath

End Property
Public Property Let KeyPath(Value As String)
Dim iSlash      As Integer

    If Right$(Value, 1) = "\" Then
        msKeyPath = Value & "(default)"
    Else
        msKeyPath = Value
    End If

    iSlash = InStr(msKeyPath, "\")

    Select Case UCase$(Left$(msKeyPath, iSlash - 1))
           Case "HKEY_CLASSES_ROOT":           mlRootKey = HKEY_CLASSES_ROOT
           Case "HKEY_CURRENT_USER":           mlRootKey = HKEY_CURRENT_USER
           Case "HKEY_LOCAL_MACHINE":          mlRootKey = HKEY_LOCAL_MACHINE
           Case "HKEY_PERFORMANCE_DATA":       mlRootKey = HKEY_PERFORMANCE_DATA
           Case "HKEY_USERS":                  mlRootKey = HKEY_USERS
      '    Case "HKEY_CURRENT_CONFIG":         mlRootKey = HKEY_CURRENT_CONFIG
      '    Case "HKEY_DYN_DATA":               mlRootKey = HKEY_DYN_DATA
    End Select

    msSubKey = Mid$(msKeyPath, iSlash + 1)

    iSlash = Len(msSubKey)
    Do
        iSlash = iSlash - 1
    Loop Until Mid$(msSubKey, iSlash, 1) = "\"

    msValueName = Mid$(msSubKey, iSlash + 1)
    If msValueName = "(default)" Then msValueName = vbNullString
    msSubKey = Left$(msSubKey, iSlash - 1)

End Property


'*******************************************************************************
'       Application             RW: Application name (defaults to App.ProductName)
Public Property Get Application() As String

    Application = msApplication

End Property
Public Property Let Application(Value As String)

    msApplication = Value

End Property


'*******************************************************************************
'       Section                 RW: Section name (defaults to "Recent")
Public Property Get Section() As String

    Section = msSection

End Property
Public Property Let Section(Value As String)

    msSection = Value

End Property


'*******************************************************************************
'       Key                     RW: Key name (no default)
Public Property Get Key() As String

    Key = msKey

End Property
Public Property Let Key(Value As String)

    msKey = Value

End Property


'*******************************************************************************
'       Default                 RW: Default value
Public Property Get Default() As String

    Default = msDefault

End Property
Public Property Let Default(Value As String)

    msDefault = Value

End Property


'*******************************************************************************
'       Value                   RW: Current value in registry
Public Property Get Value() As Variant
Dim vValue              As Variant
Dim lRetVal             As Long                         ' Return Value
Dim i As Long                                           ' Loop Counter
Dim hKey As Long                                        ' Handle To An Open Registry Key
Dim hDepth As Long                                      '
Dim lKeyValType As Long                                 ' Data Type Of A Registry Key
Dim sTempValue As String                                ' Tempory Storage For A Registry Key Value
Dim lKeyValSize As Long                                 ' Size Of Registry Key Variable

    If msKeyPath = vbNullString Then
        vValue = GetSetting(msApplication, msSection, msKey, msDefault)
    Else
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        i = 0
RETRY:
        lRetVal = RegOpenKeyEx(mlRootKey, msSubKey, 0, KEY_ALL_ACCESS, hKey)   ' Open Registry Key

        If (lRetVal <> ERROR_SUCCESS) Then
            i = i + 1
            If i < 5 Then
                GoTo RETRY
            Else
                GoTo GetKeyError
            End If
        End If

        sTempValue = String$(1024, 0)                           ' Allocate Variable Space
        lKeyValSize = 1024                                      ' Mark Variable Size

        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        lRetVal = RegQueryValueEx(hKey, msValueName, 0, _
                             lKeyValType, sTempValue, lKeyValSize)    ' Get/Create Key Value

        If (lRetVal <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
        
        If (Asc(Mid$(sTempValue, lKeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
            sTempValue = Left$(sTempValue, lKeyValSize - 1)               ' Null Found, Extract From String
        Else                                                    ' WinNT Does NOT Null Terminate String...
            sTempValue = Left$(sTempValue, lKeyValSize)                   ' Null Not Found, Extract String Only
        End If
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        'REG_NONE = 0                       ' No value type
        'REG_SZ = 1                         ' Unicode nul terminated string
        'REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
        'REG_BINARY = 3                     ' Free form binary
        'REG_DWORD = 4                      ' 32-bit number
        'REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
        'REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
        'REG_LINK = 6                       ' Symbolic Link (unicode)
        'REG_MULTI_SZ = 7                   ' Multiple Unicode strings
        'REG_RESOURCE_LIST = 8              ' Resource list in the resource map
        'REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
        'REG_RESOURCE_REQUIREMENTS_LIST = 10

        If DataType = REG_BOOLEAN And lKeyValType = REG_SZ Then
            lKeyValType = REG_BOOLEAN
        End If

        Select Case lKeyValType                                  ' Search Data Types...
        Case REG_SZ:                                             ' String Registry Key Data Type
            vValue = sTempValue                                    ' Copy String Value
        Case REG_DWORD:                                          ' Double Word Registry Key Data Type
            For i = Len(sTempValue) To 1 Step -1                    ' Convert Each Bit
                vValue = vValue + Hex(Asc(Mid$(sTempValue, i, 1))) ' Build Value Char. By Char.
            Next
            vValue = CLng(vValue)
        Case REG_BOOLEAN:
            On Error Resume Next
            vValue = CBool(sTempValue)
            If Err.Number <> 0 Then
                If Err.Number = 13 Then
                    If sTempValue = vbNullString Then
                        vValue = False
                    Else
                        vValue = True
                    End If
                Else
                    Err.Raise Err.Number
                End If
            End If
            On Error GoTo 0
        End Select

        DataType = lKeyValType
        
        If Trim$(vValue) = vbNullString Then
            vValue = msDefault
        End If

        Value = vValue                                     ' Return Success
        lRetVal = RegCloseKey(hKey)                                  ' Close Registry Key
        GoTo EXITSUB
        
    '------------------------------------------------------------
GetKeyError:        ' Cleanup After An Error Has Occured...
    '------------------------------------------------------------
        If lRetVal = ERROR_FILE_NOT_FOUND Then
            If AutoUpdate Then
                Value = msDefault
                i = 0
                GoTo RETRY
            End If
        End If
        Value = msDefault                                    ' Return Failure
        lRetVal = RegCloseKey(hKey)                                  ' Close Registry Key
    '------------------------------------------------------------
    End If

EXITSUB:
    If AutoUpdate Then
        Value = vValue
    End If

    Value = vValue


End Property
Public Property Let Value(Text As Variant)
Dim rc As Long                                      ' Return Code
Dim hKey As Long                                    ' Handle To A Registry Key
Dim hDepth As Long                                  '
Dim lpAttr As SECURITY_ATTRIBUTES                   ' Registry Security Type
Dim sText       As String
Dim iDataType       As Integer

    sText = Text

    If msKeyPath = vbNullString Then
        SaveSetting appname:=msApplication, _
                    Section:=msSection, _
                    Key:=msKey, _
                    Setting:=sText
    Else
        lpAttr.nLength = 50                                 ' Set Security Attributes To Defaults...
        lpAttr.lpSecurityDescriptor = 0                     ' ...
        lpAttr.bInheritHandle = True                        ' ...

        If DataType = REG_BOOLEAN Then
            iDataType = REG_SZ
        Else
            iDataType = DataType
        End If

        '------------------------------------------------------------
        '- Create/Open Registry Key...
        '------------------------------------------------------------
        rc = RegCreateKeyEx(mlRootKey, msSubKey, _
                            0, iDataType, _
                            REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                            hKey, hDepth)                   ' Create/Open //KeyRoot//KeyName
        
        If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Errors...
        
        '------------------------------------------------------------
        '- Create/Modify Key Value...
        '------------------------------------------------------------
        If (sText = vbNullString) Then sText = " "       ' A Space Is Needed For RegSetValueEx() To Work...
        
        rc = RegSetValueEx(hKey, msValueName, _
                           0, iDataType, _
                           sText, Len(sText))   ' Create/Modify Key Value
    
        If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Error
        '------------------------------------------------------------
        '- Close Registry Key...
        '------------------------------------------------------------
        rc = RegCloseKey(hKey)                              ' Close Key
        
        GoTo EXITSUB                                       ' Exit
    '------------------------------------------------------------
CreateKeyError:
    '------------------------------------------------------------
        rc = RegCloseKey(hKey)                              ' Attempt To Close Key
    '------------------------------------------------------------
    End If
EXITSUB:

End Property


'*******************************************************************************
'       AutoUpdate              RW: Boolean, Automatically rewrite defaults to registry?
Public Property Get AutoUpdate() As Boolean

    AutoUpdate = mbAutoUpdate

End Property
Public Property Let AutoUpdate(Value As Boolean)

    mbAutoUpdate = Value

End Property


'*******************************************************************************
'       Count                   RO: Integer, Number of values in section
Public Property Get Count() As Integer
Dim lRetVal             As Long
Dim hKey                As Long

    lRetVal = RegOpenKeyEx(mlRootKey, msSubKey, 0, KEY_ALL_ACCESS, hKey)   ' Open Registry Key
    If lRetVal <> ERROR_SUCCESS Then GoTo ERROR   ' Handle Errors...

    If mhKey = hKey Then
        Count = mlNoValues
    Else
        msClass = String(1024, " ")
        mlClassLen = Len(msClass)
        lRetVal = RegQueryInfoKey(hKey, msClass, mlClassLen, 0, _
                     mlNoSubKeys, mlSubKeyMaxLen, mlClassMaxLen, _
                     mlNoValues, mlValueNameMaxLen, mlValueMaxLen, _
                     mlSecurityLen, mftLastWrite)
        If lRetVal <> ERROR_SUCCESS Then GoTo ERROR
        Count = mlNoValues
    End If

    GoTo EXITSUB

ERROR:
    Count = 0

EXITSUB:
    Exit Property

End Property


'*******************************************************************************
'       GetAll                  Variant, Returns an array of all values in section
Public Function GetAll() As Variant
Dim lRetVal             As Long
Dim hKey                As Long
Dim iCtr                As Integer
Dim sValueName          As String       'Name of value
Dim lValueNameLen       As Long         'Length of value name
Dim lType               As Long         'Type code of value data
Dim bData               As Byte         'Pointer to Data buffer
Dim sData               As String       'Data buffer
Dim lDataLen            As Long         'Length of data buffer
Dim asGetAll()          As String       'Working output buffer
Dim sTempPath           As String
Dim sTempName           As String

    lRetVal = RegOpenKeyEx(mlRootKey, msSubKey, 0, KEY_ALL_ACCESS, hKey)   ' Open Registry Key
    If lRetVal <> ERROR_SUCCESS Then GoTo ERROR   ' Handle Errors...

    If mhKey <> hKey Then
        msClass = String(1024, " ")
        mlClassLen = Len(msClass)
        lRetVal = RegQueryInfoKey(hKey, msClass, mlClassLen, 0, _
                     mlNoSubKeys, mlSubKeyMaxLen, mlClassMaxLen, _
                     mlNoValues, mlValueNameMaxLen, mlValueMaxLen, _
                     mlSecurityLen, mftLastWrite)
        If lRetVal <> ERROR_SUCCESS Then GoTo ERROR
    End If

    If mlNoValues = 0 Then
        GoTo ERROR
    End If

    sTempPath = String(1024, " ")
    lDataLen = Len(sTempPath)
    lRetVal = GetTempPath(lDataLen, sTempPath)
    If InStr(sTempPath, vbNullChar) > 0 Then
        sTempPath = Left$(sTempPath, InStr(sTempPath, vbNullChar) - 1)
    End If

    sTempName = String(1024, " ")
    lDataLen = Len(sTempName)
    lRetVal = GetTempFileName(sTempPath, "ats", 0, sTempName)
    If InStr(sTempName, vbNullChar) > 0 Then
        sTempName = Left$(sTempName, InStr(sTempName, vbNullChar) - 1)
    End If

    lRetVal = RegSaveKey(hKey, sTempName, 0)

    ReDim asGetAll(1 To mlNoValues, 0 To 1) As String
    
    For iCtr = 0 To mlNoValues - 1
        sValueName = String(1024, " ")
        lValueNameLen = Len(sValueName)
        sData = String(1024, " ")
        lDataLen = Len(sData)
        lRetVal = RegEnumValue(hKey, iCtr, sValueName, lValueNameLen, 0, lType, bData, lDataLen)
        asGetAll(iCtr + 1, 0) = Trim$(sValueName)
        asGetAll(iCtr + 1, 1) = Trim$(sData)
    Next iCtr

    GetAll = asGetAll()

    GoTo EXITSUB

ERROR:
    GetAll = Empty

EXITSUB:
    Exit Function

End Function

'*******************************************************************************
'       DataType                RW: Integer, Type of value to read/write
Public Property Get DataType() As Integer

    DataType = miType

End Property
Public Property Let DataType(Value As Integer)

    miType = Value

End Property
