VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSqlSp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'COPYRIGHT  : Copyright (c) 2000 Chung Lau.
'Disclaimer : This SOFTWARE is provided "as is" without representations or
'           : warranties of any kind, whether expressed or implied.
'           : The USER must assume the entire risk of using the SOFTWARE,
'           : and Chung Lau shall have no liability to the USER or any other
'           : third-party for any damages whatsoever, including, but not limited to,
'           : any economic or data loss, even if such loss was foreseeable by Chung Lau.
'           :
'PURPOSE    : Contains shared database access routines
'EFFECT     :
'LOGIC      :
'TOOLS      :
'           :
'---------- DEVELOPMENT ENVIRONMENT -------------------------------------------
'OS         : Windows 2000
'HARDWARE   :
'COMPILER   : MS Visual Basic 6.0 (Enterprise Edition) 32bit SP3
'           :
'TOOLS      :
'           :
'---------- PRODUCTION ENVIRONMENT --------------------------------------------
'OS         :
'HARDWARE   :
'RUN FILES  : List of files needed to run this program
'           :
'---------- VERSION LOG -------------------------------------------------------
' CL - Chung Lau, clau@globehead.com
'
' Ver     Date   Init  Description
'-----  -------- ----  --------------------------------------------------------
'01.00  07/31/00  CL   Started
'01.00  08/30/00  CL   Added user specified parameter values for spSelectList function
'
'******************************************************************************
Option Explicit
Option Base 0
Option Compare Text

Private Const adolMAXROWS As Long = 10000

'Connect string
Private m_sConnString         As String
Private m_sProvider                 As String
Private m_sServerName               As String
Private m_sDriver                   As String
Private m_sDatabase                 As String
Private m_sDataSourceName           As String
Private m_lCursorLocation           As Long


Private m_sUID                      As String
Private m_sPassword                 As String

Private adMainConn                  As New adodb.Connection

'Connection flag.
Private m_bConnected                As Boolean

Private Type udtDataType
    lAdoDataTypeLen As Long
    sAdoDataTypeName As String
    sSqlDataTypeName As String
    sVbDataTypeName As String
    sVbNamingNotation As String
    vntDefaultValue As Variant
End Type
Private audtDataType(255) As udtDataType
Private asParamDirectionType(255) As String
Private vntParams() As Variant

'-<Private Class Methods>----------------------------------------------------------------------------------------------------------------------
Private Sub BuildDataTypeTable(ByVal pvbyAdoDataTypeNum As Byte, ByVal pvsAdoDataTypeName As String, ByVal pvlAdoDataTypeLen As Long, ByVal pvsSqlDataTypeName As String, ByVal pvsVBDataTypeName As String, ByVal pvsVBNamingNotation As String, ByVal pvvntDefaultValue As Variant)
    With audtDataType(pvbyAdoDataTypeNum)
        .lAdoDataTypeLen = pvlAdoDataTypeLen
        .sAdoDataTypeName = pvsAdoDataTypeName
        .sSqlDataTypeName = pvsSqlDataTypeName
        .sVbDataTypeName = pvsVBDataTypeName
        .sVbNamingNotation = pvsVBNamingNotation
        .vntDefaultValue = pvvntDefaultValue
    End With
End Sub
Private Sub Class_Initialize()

    '--- Build the ADO Data Type Table ---
    '==>  adArray - Or'd together with another type to indicate that the data is a safe-array of that type (DBTYPE_ARRAY).
    '==>  adBigInt - An 8-byte signed integer (DBTYPE_I8).
    Call BuildDataTypeTable(adBigInt, "adBigInt", 8, "Not Supported", "Not Supported", "", 0)
    '==>  adBinary - A binary value (DBTYPE_BYTES).
    Call BuildDataTypeTable(adBinary, "adBinary", 0, "Binary/Timestamp", "Variant", "vnt", Null)
    '==>  adBoolean - A Boolean value (DBTYPE_BOOL).
    Call BuildDataTypeTable(adBoolean, "adBoolean", 1, "bit", "Boolean", "b", False)
    '==>  adByRef - Or'd together with another type to indicate that the data is a pointer to data of the other type (DBTYPE_BYREF).
    '==>  adBSTR - A null-terminated character string (Unicode) (DBTYPE_BSTR).
    Call BuildDataTypeTable(adBSTR, "adBSTR", 0, "Not Supported", "Not Supported", "", Null)
    '==>  adChar - A String value (DBTYPE_STR).
    Call BuildDataTypeTable(adChar, "adChar", 0, "Char/EmpID", "String", "s", "")
    '==>  adCurrency - A currency value (DBTYPE_CY). Currency is a fixed-point number with 4 digits to the right of the decimal point. It is stored in an 8-byte signed integer scaled by 10,000.
    Call BuildDataTypeTable(adCurrency, "adCurrency", 8, "Money/SmallMoney", "Currency", "c", 0@)
    '==>  adDate - A Date value (DBTYPE_DATE). A date is stored as a Double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day.
    Call BuildDataTypeTable(adDate, "adDate", 8, "Not Supported", "Date", "dt", #1/1/1999#)
    '==>  adDBDate - A date value (yyyymmdd) (DBTYPE_DBDATE).
    Call BuildDataTypeTable(adDBDate, "adDBDate", 8, "Not Supported", "Date", "dt", #1/1/1999#)
    '==>  adDBTime - A time value (hhmmss) (DBTYPE_DBTIME).
    Call BuildDataTypeTable(adDBTime, "adDBTime", 8, "Not Supported", "Not Supported", "", Null)
    '==>  adDBTimeStamp - A date-time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP).
    Call BuildDataTypeTable(adDBTimeStamp, "adDBTimeStamp", 8, "Datetime/SmallDatetime", "Date", "dtm", #1/1/1999#)
    '==>  adDecimal - An exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL).
    Call BuildDataTypeTable(adDecimal, "adDecimal", 16, "Decimal", "String", "s", "")
    '==>  adDouble - A double-precision floating point value (DBTYPE_R8).
    Call BuildDataTypeTable(adDouble, "adDouble", 8, "Float", "Double", "d", 0#)
    '==>  adEmpty - No value was specified (DBTYPE_EMPTY).
    Call BuildDataTypeTable(adEmpty, "adEmpty", 0, "Not Supported", "Not Supported", "", Null)
    '==>  adError - A 32-bit error code (DBTYPE_ERROR).
    Call BuildDataTypeTable(adError, "adError", 0, "Not Supported", "Not Supported", "", Null)
    '==>  adGUID - A globally unique identifier (GUID) (DBTYPE_GUID).
    Call BuildDataTypeTable(adGUID, "adGUID", 0, "UniqueIdentifier", "Variant", "", Null)
    '==>  adIDispatch - A pointer to an IDispatch interface on an OLE object (DBTYPE_IDISPATCH).
    Call BuildDataTypeTable(adIDispatch, "adIDispatch", 0, "Not Supported", "Not Supported", "", Null)
    '==>  adInteger - A 4-byte signed integer (DBTYPE_I4).
    Call BuildDataTypeTable(adInteger, "adInteger", 4, "int/identity", "Long", "l", 0&)
    '==>  adIUnknown - A pointer to an IUnknown interface on an OLE object (DBTYPE_IUNKNOWN).
    Call BuildDataTypeTable(adIUnknown, "adIUnknown", 0, "Not Supported", "Not Supported", "", Null)
    '==>  adLongVarBinary - A long binary value (Parameter object only).
    Call BuildDataTypeTable(adLongVarBinary, "adLongVarBinary", 0, "Image", "Variant", "vnt", Null)
    '==>  adLongVarChar - A long String value (Parameter object only).
    Call BuildDataTypeTable(adLongVarChar, "adLongVarChar", 0, "Text", "String", "s", "")
    '==>  adLongVarWChar - A long null-terminated string value (Parameter object only).
    Call BuildDataTypeTable(adLongVarWChar, "adLongVarWChar", 0, "nText", "String", "s", "")
    '==>  adNumeric - An exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
    Call BuildDataTypeTable(adNumeric, "adNumeric", 16, "Numeric", "String", "s", "")
        '==>  adSingle - A single-precision floating point value (DBTYPE_R4).
    Call BuildDataTypeTable(adSingle, "adSingle", 4, "Real", "Single", "f", 0!)
    '==>  adSmallInt - A 2-byte signed integer (DBTYPE_I2).
    Call BuildDataTypeTable(adSmallInt, "adSmallInt", 2, "SmallInt", "Integer", "n", 0)
    '==>  adTinyInt - A 1-byte signed integer (DBTYPE_I1).
    Call BuildDataTypeTable(adTinyInt, "adTinyInt", 1, "Not Supported", "Not Supported", "", Null)
    '==>  adUnsignedBigInt - An 8-byte unsigned integer (DBTYPE_UI8).
    Call BuildDataTypeTable(adUnsignedBigInt, "adUnsignedBigInt", 8, "Not Supported", "Not Supported", "", Null)
    '==>  adUnsignedInt - A 4-byte unsigned integer (DBTYPE_UI4).
    Call BuildDataTypeTable(adUnsignedInt, "adUnsignedInt", 4, "Not Supported", "Not Supported", "", Null)
    '==>  adUnsignedSmallInt - A 2-byte unsigned integer (DBTYPE_UI2).
    Call BuildDataTypeTable(adUnsignedSmallInt, "adUnsignedSmallInt", 2, "Not Supported", "Not Supported", "", Null)
    '==>  adUnsignedTinyInt - A 1-byte unsigned integer (DBTYPE_UI1).
    Call BuildDataTypeTable(adUnsignedTinyInt, "adUnsignedTinyInt", 1, "Not Supported", "Not Supported", "", Null)
    '==>  adUserDefined - A user-defined variable (DBTYPE_UDT).
    Call BuildDataTypeTable(adUserDefined, "adUserDefined", 0, "Not Supported", "Not Supported", "", Null)
    '==>  adVarBinary - A binary value (Parameter object only).
    Call BuildDataTypeTable(adVarBinary, "adVarBinary", 0, "VarBinary", "Variant", "vnt", Null)
    '==>  adVarChar - A String value (Parameter object only).
    Call BuildDataTypeTable(adVarChar, "adVarChar", 0, "VarChar/TID/SysName", "String", "s", "")
    '==>  adVariant - An Automation Variant (DBTYPE_VARIANT).
    Call BuildDataTypeTable(adVariant, "adVariant", 0, "Not Supported", "Not Supported", "", Null)
    '==>  adVector - Or'd together with another type to indicate that the data is a DBVECTOR structure, as defined by OLE DB, that contains a count of elements and a pointer to data of the other type (DBTYPE_VECTOR).
    '==>  adVarWChar - A null-terminated Unicode character string (Parameter object only).
    Call BuildDataTypeTable(adVarWChar, "adVarWChar", 0, "nVarChar/SysName", "String", "s", "")
   '==>  adWChar - A null-terminated Unicode character string (DBTYPE_WSTR).
    Call BuildDataTypeTable(adWChar, "adWChar", 0, "nChar", "String", "s", "")

    '--- Initialize ADO Parameter Direction names ---
    asParamDirectionType(adParamUnknown) = "adParamUnknown"
    asParamDirectionType(adParamReturnValue) = "adParamReturnValue"
    asParamDirectionType(adParamOutput) = "adParamOutput"
    asParamDirectionType(adParamInputOutput) = "adParamInputOutput"
    asParamDirectionType(adParamInput) = "adParamInput"

End Sub

Private Sub Class_Terminate()
   If m_bConnected Then Disconnect

End Sub

'-<Public Methods/Functions/Subs>------------------------------------------------------------------------------------------------------------
Public Property Get ConnString() As String
    ConnString = m_sConnString
End Property
Public Property Let ConnString(sConnectString As String)
    Dim DataString As String
    ReDim substr(0) As String
    Dim nSubstrCnt As Integer, i As Integer
    
    m_sConnString = sConnectString
    If InStr(sConnectString, ";") Then
        DataString = m_sConnString
        nSubstrCnt = ParseString(substr(), DataString, ";")
        
        For i = 1 To nSubstrCnt
            'Debug.Print substr(i)
            Select Case UCase(Left(substr(i), InStr(substr(i), "=") - 1))
                Case "DRIVER"
                    Driver = Mid(substr(i), InStr(substr(i), "=") + 1)
                Case "DATABASE"
                    Database = Mid(substr(i), InStr(substr(i), "=") + 1)
                Case "SERVER"
                    ServerName = Mid(substr(i), InStr(substr(i), "=") + 1)
                Case "UID"
                    UID = Mid(substr(i), InStr(substr(i), "=") + 1)
                Case "PWD"
                    Password = Mid(substr(i), InStr(substr(i), "=") + 1)
            End Select
        Next
    End If
End Property

Public Sub BeginTransaction()
    adMainConn.BeginTrans
End Sub
Public Sub CommitTransaction()
    adMainConn.CommitTrans
End Sub
Public Sub RollbackTransaction()
    adMainConn.RollbackTrans
End Sub

Public Property Get UID() As String
    UID = m_sUID
End Property
Public Property Let UID(ByVal UID As String)
    m_sUID = UID
End Property

Public Property Get Password() As String
    Password = m_sPassword
End Property
Public Property Let Password(ByVal Password As String)
    m_sPassword = Password
End Property

Public Property Get ServerName() As String
    ServerName = m_sServerName
End Property
Public Property Let ServerName(ByVal ServerName As String)
    m_sServerName = ServerName
End Property

Public Property Get Driver() As String
    Driver = m_sDriver
End Property
Public Property Let Driver(ByVal Driver As String)
    m_sDriver = Driver
End Property

Public Property Get Database() As String
    Database = m_sDatabase
End Property
Public Property Let Database(ByVal Database As String)
    m_sDatabase = Database
End Property

Public Property Get dataSourceName() As String
    dataSourceName = m_sDataSourceName
End Property
Public Property Let dataSourceName(ByVal dataSourceName As String)
    m_sDataSourceName = dataSourceName
End Property

Public Property Get Provider() As String
    Provider = m_sProvider
End Property
Public Property Let Provider(ByVal sProvider As String)
    m_sProvider = sProvider
End Property

Public Property Let CursorLocation(Location As CursorLocationEnum)
    m_lCursorLocation = Location
End Property
Public Property Get CursorLocation() As CursorLocationEnum
    CursorLocation = m_lCursorLocation
End Property

Public Property Get Connected() As Boolean
    Connected = m_bConnected
End Property

'cl 2/9/99 Below property added to allow outside access of the Ado Connection
Public Property Get Connection() As adodb.Connection
    Set Connection = adMainConn
End Property

Public Function Connect() As Boolean
    Dim strConnect As String
    Dim bRet As Boolean

    On Error GoTo ConnectError


    'Build the DSN string from the parameters
    'm_sConnString = "DRIVER=" & Driver & ";SERVER=" & ServerName & ";DATABASE=" & Database & ";UID=" & UID & ";PWD=" & Password
    'ConcatDSN(sDSN, sUserID, sPassword, sDatabase)
    
    With adMainConn
        If Len(Trim(Provider)) > 0 Then
            .Provider = m_sProvider
        End If
        
        If Len(Trim(m_sConnString)) > 0 Then
            .ConnectionString = m_sConnString
        End If
        '"DRIVER={SQL Server};SERVER=CWCDEVSERV;DATABASE=BILLING;UID=sa;PWD=;"
'        .CursorLocation = m_lCursorLocation
        .Open
    End With

    bRet = True

    If bRet Then
        ' Get First connection.
        'Set DBCFramework.Connection = adMainConn
        'cl 2/9/99 Above line commented to make connection independent
    End If

    m_bConnected = bRet

ConnectExit:
    Connect = m_bConnected
    Exit Function

ConnectError:
    bRet = False
    On Error GoTo 0
    m_bConnected = False
    Resume ConnectExit

End Function

Public Sub Disconnect()
   On Error Resume Next
   
    adMainConn.Close
    
   'Let the contract object handle disconnecting rdo.
    'DBCFramework.Disconnect
    'cl 2/9/99 Above line commented to make connection independent
        
    m_bConnected = False
   
   'Dereference object refs.
    'Set DBCFramework = Nothing
    'cl 2/9/99 Above line commented to make connection independent

    Set adMainConn = Nothing
End Sub

'-<Private Methods/Functions>-----------------------------------------------------------------------------------------

Private Function ParseString(SubStrs() As String, ByVal SrcStr As String, ByVal Delimiter As String) As Integer
    ReDim SubStrs(0) As String
    Dim CurPos As Long
    Dim NextPos As Long
    Dim DelLen As Integer
    Dim nCount As Integer
    Dim TStr As String
      
    ' Add delimiters to start and end of string to make loop simpler:
    'SrcStr = Delimiter & SrcStr & Delimiter
    SrcStr = SrcStr & Delimiter
    
    ' Calculate the delimiter length only once:
    DelLen = Len(Delimiter)
    ' Initialize the count and position:
    nCount = 0
    CurPos = 1
    NextPos = InStr(CurPos + DelLen, SrcStr, Delimiter)
    ' Loop searching for delimiters:
    Do Until NextPos = 0
        ' Extract a sub-string:
        If CurPos = 1 Then
            TStr = Mid$(SrcStr, CurPos, NextPos - CurPos)
        Else
            TStr = Mid$(SrcStr, CurPos + DelLen, NextPos - CurPos - DelLen)
        End If
        ' Increment the sub string counter:
        nCount = nCount + 1
        ' Add room for the new sub-string in the array:
        ReDim Preserve SubStrs(nCount) As String
        ' Put the sub-string in the array:
        SubStrs(nCount) = TStr
        ' Position to the last found delimiter:
        CurPos = NextPos
        ' Find the next delimiter:
        NextPos = InStr(CurPos + DelLen, SrcStr, Delimiter)
    Loop
    ' Return the number of sub-strings found:
    ParseString = nCount
   
End Function

Public Function grstSpSelect(ByVal pvsSpName As String, Optional ByVal pavntParam As Variant) As adodb.Recordset
'------------------------------------------------------------------------------
'PURPOSE: Executes MS Sql Server Stored Procedure with parameters
'INPUT  :
'           pvsSpName - Name of the SP Name in SQL Server
'           pavntParam - Optionally a list of parameters
'ASSUME : adMainConn is already connected to a database
'RETURN : Result record set if success, Nothing if failed
'EFFECT : Error must be handle by the caller
'LOGIC  : Using the Parameters.Refresh method get the parameter list from
'           SQL SP so that we don't have to explicitly define each parameter
'           in code.
'------------------------------------------------------------------------------
Dim adCmd As New adodb.Command
Dim i As Integer
Dim nParam As Integer

    With adCmd
        Set .ActiveConnection = adMainConn
        .CommandText = pvsSpName
        .CommandType = adCmdStoredProc
    
        If Not IsMissing(pavntParam) Then
            .Parameters.Refresh
            For i = LBound(pavntParam) To UBound(pavntParam)
                If i > 0 Then
                    nParam = nParam + 1
                    .Parameters(nParam) = pavntParam(i)
                    .CreateParameter
                End If
            Next i
        End If
        
        Set grstSpSelect = .Execute
    End With
    Set adCmd = Nothing
End Function

Public Function gprmSpParam(ByVal pvsSpName As String) As adodb.Parameters
'------------------------------------------------------------------------------
'PURPOSE: Executes MS Sql Server Stored Procedure with parameters
'INPUT  :
'           pvsSpName - Name of the SP Name in SQL Server

'ASSUME : adMainConn is already connected to a database
'RETURN : Result parameter string if success, empty string if failed
'EFFECT : Error must be handle by the caller
'LOGIC  : Using the Parameters.Refresh method get the parameter list from
'           SQL SP so that we don't have to explicitly define each parameter
'           in code.
'------------------------------------------------------------------------------
On Error Resume Next
Dim adCmd As New adodb.Command

    With adCmd
        Set .ActiveConnection = adMainConn
        .CommandText = pvsSpName
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        Set gprmSpParam = .Parameters
    End With
    Set adCmd = Nothing
End Function

Public Function spParamList(ByVal pvsSpName As String, ByVal pvsConStr As String) As String
'------------------------------------------------------------------------------
'PURPOSE: Gets parameter information from a Stored Procedure
'INPUT  :
'           pvsSpName - Name of the SP Name in SQL Server
'           pvsConStr - Connection string, such as DSN=dbTest
'
'ASSUME : adMainConn is already connected to a database
'RETURN : Result parameter string if success, empty string if failed
'EFFECT : Error must be handle by the caller
'LOGIC  :
'
'------------------------------------------------------------------------------
Dim prms As adodb.Parameters
Dim prm As adodb.Parameter
Dim bIncludeParam As Boolean

    On Error GoTo spParamList_Error
    m_sConnString = pvsConStr
    If Connect() Then
        Set prms = gprmSpParam(pvsSpName)
        If Err.Number <> 0 Then
            MsgBox Err.Description
           Exit Function
        End If
        For Each prm In prms
            With prm
                If bIncludeParam Then
                    spParamList = spParamList & .Name & " " & audtDataType(.Type).sAdoDataTypeName & " " & audtDataType(.Type).sVbDataTypeName & " " & CStr(.Size) & " " & asParamDirectionType(.Direction) & ";"
                End If
            End With
            bIncludeParam = True
        Next
        Set prms = Nothing
        Set prm = Nothing
        Call Disconnect
    End If
   'spParamList = spParamList & "; "

spParamList_Exit:
    Exit Function

spParamList_Error:
    MsgBox Err.Description
    Resume spParamList_Exit
End Function

Public Function spSelectList(ByVal pvsSpName As String, ByVal pvsConStr As String, ByVal pvsParamValue As String) As String
'------------------------------------------------------------------------------
'PURPOSE: Gets parameter and recordset information from a Stored Procedure
'INPUT  :
'           pvsSpName - Name of the SP Name in SQL Server
'           pvsParamValue - comma delimited listing of parameter values
'           pvsConStr - Connection string, such as DSN=dbTest
'
'ASSUME : adMainConn is already connected to a database
'RETURN : Result parameter string and fields if success, empty string if failed
'EFFECT : Error must be handle by the caller
'LOGIC  :
'
'------------------------------------------------------------------------------
Dim nParam As Integer
Dim prms As adodb.Parameters
Dim prm As adodb.Parameter
Dim rst As adodb.Recordset
Dim fld As adodb.Field
Dim bIncludeParam As Boolean
Dim sParamList As String
    
    On Error GoTo spSelectList_Error

    ConnString = pvsConStr
    If Connect() Then
        Set prms = gprmSpParam(pvsSpName)
        If Err.Number <> 0 Then
            MsgBox Err.Description
           Exit Function
        End If
        For Each prm In prms
            ReDim Preserve vntParams(nParam)
            With prm
                If bIncludeParam Then
                    sParamList = sParamList & .Name & " " & audtDataType(.Type).sAdoDataTypeName & " " & audtDataType(.Type).sVbDataTypeName & " " & CStr(.Size) & " " & asParamDirectionType(.Direction) & " "
                    If pvsParamValue = "" Or pvsParamValue = ";" Then
                        vntParams(nParam) = audtDataType(.Type).vntDefaultValue
                    Else
                        vntParams(nParam) = vntGetToken(pvsParamValue, nParam, .Type)
                    End If
                End If
            End With
            nParam = nParam + 1
            bIncludeParam = True
        Next
        
        If nParam = 1 Then
            Set rst = grstSpSelect(pvsSpName)
        Else
            Set rst = grstSpSelect(pvsSpName, vntParams)
        End If
        
        With rst
            For Each fld In .Fields
                With fld
                    spSelectList = spSelectList & .Name & " " & audtDataType(.Type).sVbDataTypeName & " " & CStr(.DefinedSize) & " "
                End With
            Next
        End With
        Set prms = Nothing
        Set prm = Nothing
        Set rst = Nothing
        Set fld = Nothing
        Call Disconnect
    End If
    spSelectList = sParamList & "; " & spSelectList
    
spSelectList_Exit:
    Exit Function

spSelectList_Error:
    MsgBox Err.Description
    spSelectList = "; "
    Resume spSelectList_Exit
End Function

Private Function vntGetToken(ByVal pvsParamValue As String, pvnParam As Integer, ByVal pvlType As Byte) As Variant
Dim sToken As String

    vntGetToken = Null
    sToken = sGetToken(pvsParamValue, pvnParam, ",")
    Select Case pvlType
        '==>  adBigInt - An 8-byte signed integer (DBTYPE_I8).
        Case adBigInt ' "adBigInt", 8, "Not Supported", "Not Supported", "", Null)
        '==>  adBinary - A binary value (DBTYPE_BYTES).
        Case adBinary ' "adBinary", 0, "Binary/Timestamp", "Variant", "vnt", Null)
        '==>  adBoolean - A Boolean value (DBTYPE_BOOL).
        Case adBoolean ' "adBoolean", 1, "bit", "Boolean", "b", False)
            If sToken = "0" Then
                vntGetToken = 0
            Else
                vntGetToken = 1
            End If
            
        '==>  adByRef - Or'd together with another type to indicate that the data is a pointer to data of the other type (DBTYPE_BYREF).
        '==>  adBSTR - A null-terminated character string (Unicode) (DBTYPE_BSTR).
        Case adBSTR ' "adBSTR", 0, "Not Supported", "Not Supported", "", Null)
        '==>  adChar - A String value (DBTYPE_STR).
        Case adChar ' "adChar", 0, "Char/EmpID", "String", "s", Null)
        '==>  adCurrency - A currency value (DBTYPE_CY). Currency is a fixed-point number with 4 digits to the right of the decimal point. It is stored in an 8-byte signed integer scaled by 10,000.
        Case adCurrency ' "adCurrency", 8, "Money/SmallMoney", "Currency", "c", 0@)
            vntGetToken = CCur(sToken)
        
        '==>  adDate - A Date value (DBTYPE_DATE). A date is stored as a Double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day.
        Case adDate ' "adDate", 8, "Not Supported", "Date", "dt", #1/1/1999#)
            vntGetToken = CDate(sToken)
            
        '==>  adDBDate - A date value (yyyymmdd) (DBTYPE_DBDATE).
        Case adDBDate ' "adDBDate", 8, "Not Supported", "Date", "dt", #1/1/1999#)
        '==>  adDBTime - A time value (hhmmss) (DBTYPE_DBTIME).
        Case adDBTime ' "adDBTime", 8, "Not Supported", "Not Supported", "", Null)
        '==>  adDBTimeStamp - A date-time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP).
        Case adDBTimeStamp ' "adDBTimeStamp", 8, "Datetime/SmallDatetime", "Date", "dtm", #1/1/1999#)
            vntGetToken = CCur(sToken)

        '==>  adDecimal - An exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL).
        Case adDecimal ' "adDecimal", 16, "Decimal", "String", "s", Null)
        '==>  adDouble - A double-precision floating point value (DBTYPE_R8).
        Case adDouble ' "adDouble", 8, "Float", "Double", "d", 0#)
            vntGetToken = CDbl(sToken)
        
        '==>  adEmpty - No value was specified (DBTYPE_EMPTY).
        Case adEmpty ' "adEmpty", 0, "Not Supported", "Not Supported", "", Null)
        '==>  adError - A 32-bit error code (DBTYPE_ERROR).
        Case adError ' "adError", 0, "Not Supported", "Not Supported", "", Null)
        '==>  adGUID - A globally unique identifier (GUID) (DBTYPE_GUID).
        Case adGUID ' "adGUID", 0, "UniqueIdentifier", "Variant", "", Null)
        '==>  adIDispatch - A pointer to an IDispatch interface on an OLE object (DBTYPE_IDISPATCH).
        Case adIDispatch ' "adIDispatch", 0, "Not Supported", "Not Supported", "", Null)
        '==>  adInteger - A 4-byte signed integer (DBTYPE_I4).
        Case adInteger ' "adInteger", 4, "int/identity", "Long", "l", 0&)
            vntGetToken = CLng(sToken)
        
        '==>  adIUnknown - A pointer to an IUnknown interface on an OLE object (DBTYPE_IUNKNOWN).
        Case adIUnknown ' "adIUnknown", 0, "Not Supported", "Not Supported", "", Null)
        '==>  adLongVarBinary - A long binary value (Parameter object only).
        Case adLongVarBinary ' "adLongVarBinary", 0, "Image", "Variant", "vnt", Null)
        '==>  adLongVarChar - A long String value (Parameter object only).
        Case adLongVarChar ' "adLongVarChar", 0, "Text", "String", "s", Null)
            vntGetToken = sToken
        
        '==>  adLongVarWChar - A long null-terminated string value (Parameter object only).
        Case adLongVarWChar ' "adLongVarWChar", 0, "nText", "String", "s", Null)
            vntGetToken = sToken
        
        '==>  adNumeric - An exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
        Case adNumeric ' "adNumeric", 16, "Numeric", "String", "s", Null)
            vntGetToken = sToken
        
        '==>  adSingle - A single-precision floating point value (DBTYPE_R4).
        Case adSingle ' "adSingle", 4, "Real", "Single", "f", 0!)
            vntGetToken = CSng(sToken)
        
        '==>  adSmallInt - A 2-byte signed integer (DBTYPE_I2).
        Case adSmallInt ' "adSmallInt", 2, "SmallInt", "Integer", "n", 0)
            vntGetToken = CInt(sToken)
        
        '==>  adTinyInt - A 1-byte signed integer (DBTYPE_I1).
        Case adTinyInt ' "adTinyInt", 1, "Not Supported", "Not Supported", "", Null)
        '==>  adUnsignedBigInt - An 8-byte unsigned integer (DBTYPE_UI8).
        Case adUnsignedBigInt ' "adUnsignedBigInt", 8, "Not Supported", "Not Supported", "", Null)
        '==>  adUnsignedInt - A 4-byte unsigned integer (DBTYPE_UI4).
        Case adUnsignedInt ' "adUnsignedInt", 4, "Not Supported", "Not Supported", "", Null)
        '==>  adUnsignedSmallInt - A 2-byte unsigned integer (DBTYPE_UI2).
        Case adUnsignedSmallInt ' "adUnsignedSmallInt", 2, "Not Supported", "Not Supported", "", Null)
        '==>  adUnsignedTinyInt - A 1-byte unsigned integer (DBTYPE_UI1).
        Case adUnsignedTinyInt ' "adUnsignedTinyInt", 1, "Not Supported", "Not Supported", "", Null)
        '==>  adUserDefined - A user-defined variable (DBTYPE_UDT).
        Case adUserDefined ' "adUserDefined", 0, "Not Supported", "Not Supported", "", Null)
        '==>  adVarBinary - A binary value (Parameter object only).
        Case adVarBinary ' "adVarBinary", 0, "VarBinary", "Variant", "vnt", Null)
        '==>  adVarChar - A String value (Parameter object only).
        Case adVarChar ' "adVarChar", 0, "VarChar/TID/SysName", "String", "s", "")
            vntGetToken = sToken
        
        '==>  adVariant - An Automation Variant (DBTYPE_VARIANT).
        Case adVariant ' "adVariant", 0, "Not Supported", "Not Supported", "", Null)
        '==>  adVector - Or'd together with another type to indicate that the data is a DBVECTOR structure, as defined by OLE DB, that contains a count of elements and a pointer to data of the other type (DBTYPE_VECTOR).
        '==>  adVarWChar - A null-terminated Unicode character string (Parameter object only).
        Case adVarWChar ' "adVarWChar", 0, "nVarChar/SysName", "String", "s", Null)
            vntGetToken = sToken
            
        '==>  adWChar - A null-terminated Unicode character string (DBTYPE_WSTR).
        Case adWChar ' "adWChar", 0, "nChar", "String", "s", Null)
            vntGetToken = sToken
        
    End Select
End Function

' ================================================================================
' Synopsis          Returns the Nth Token from siAllTokens delimited by sDelim
'
' Parameters
'
'   siAllTokens                 (I) Required. The string containing all the tokens
'   iToken                      (I) Optional. The index of the token to return
'                                   DEFAULT = 1
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = gsS
' Description
'  For the following:
'    siAllTokens         iToken   sDelim  Returns       Notes
'   "William M Rawls"    1       gsS     "William"      First word
'   "William M Rawls"    2       gsS     "M"            Second word
'   "William M Rawls"    3       gsS     "Rawls"        Third word
'   "William M Rawls"    4       gsS     vbNullString             No forth word
'   "William M Rawls"    0       gsS     vbNullString             Zeroth token is always empty
'   "William M Rawls"   -1       gsS     vbNullString             Negative tokesn always empty
'   "William M Rawls"    1       vbNullString      vbNullString             No delimiter ? Token empty
' ================================================================================
Private Function sGetToken(ByVal siAllTokens As String, Optional ByVal iToken As Long = 1, Optional ByVal sDelim As String = ",") As String
244      Static iCurTokenLocation As Long                  ' Character position of the first delimiter string
245      Static nDelim As Long                             ' Length of the delimiter string
246      nDelim = Len(sDelim)

247      If iToken < 1 Or nDelim < 1 Then
        ' Negative or zeroth token or empty delimiter strings mean an empty token
248          Exit Function
249      ElseIf iToken = 1 Then
        ' Quickly extract the first token
250          iCurTokenLocation = InStr(siAllTokens, sDelim)
251          If iCurTokenLocation > 1 Then
252              sGetToken = Left$(siAllTokens, iCurTokenLocation - 1)
253          ElseIf iCurTokenLocation = 1 Then
254              sGetToken = vbNullString
255          Else
256              sGetToken = siAllTokens
257          End If
258          Exit Function
259      Else
        ' Find the Nth token
260          Do
261              iCurTokenLocation = InStr(siAllTokens, sDelim)
262              If iCurTokenLocation = 0 Then
263                  Exit Function
264              Else
265                  siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
266              End If
267              iToken = iToken - 1
268          Loop Until iToken = 1

        ' Extract the Nth token (Which is the next token at this point)
269          iCurTokenLocation = InStr(siAllTokens, sDelim)
270          If iCurTokenLocation > 0 Then
271              sGetToken = Left$(siAllTokens, iCurTokenLocation - 1)
272              Exit Function
273          Else
274              sGetToken = siAllTokens
275              Exit Function
276          End If
277      End If
End Function

Public Function ListStoredProcedures(ByVal DSN As String) As String
On Error Resume Next
    Dim catProbe        As ADOX.Catalog
    Dim stpProbe        As ADOX.Procedures
    Dim CurrStoredProc  As ADOX.Procedure
    Dim strListOut      As String

    Set catProbe = New ADOX.Catalog
            catProbe.ActiveConnection = DSN
            If Err.Number = 0 Then
                With catProbe
                     Set stpProbe = .Procedures
                     For Each CurrStoredProc In stpProbe
                         With CurrStoredProc
                              strListOut = strListOut & ";" & sGetToken(.Name, 1, ";")
                         End With
                     Next CurrStoredProc
                End With
            End If
    Set catProbe = Nothing

    If Len(strListOut) Then
       ListStoredProcedures = Mid$(strListOut, 2)
    End If
End Function


