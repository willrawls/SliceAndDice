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
    Call BuildDataTypeTable(adBigInt, "adBigInt", 8, "Not Supported", "Not Supported", "", Null)
    '==>  adBinary - A binary value (DBTYPE_BYTES).
    Call BuildDataTypeTable(adBinary, "adBinary", 0, "Binary/Timestamp", "Variant", "vnt", Null)
    '==>  adBoolean - A Boolean value (DBTYPE_BOOL).
    Call BuildDataTypeTable(adBoolean, "adBoolean", 1, "bit", "Boolean", "b", False)
    '==>  adByRef - Or'd together with another type to indicate that the data is a pointer to data of the other type (DBTYPE_BYREF).
    '==>  adBSTR - A null-terminated character string (Unicode) (DBTYPE_BSTR).
    Call BuildDataTypeTable(adBSTR, "adBSTR", 0, "Not Supported", "Not Supported", "", Null)
    '==>  adChar - A String value (DBTYPE_STR).
    Call BuildDataTypeTable(adChar, "adChar", 0, "Char/EmpID", "String", "s", Null)
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
    Call BuildDataTypeTable(adDecimal, "adDecimal", 16, "Decimal", "String", "s", Null)
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
    Call BuildDataTypeTable(adLongVarChar, "adLongVarChar", 0, "Text", "String", "s", Null)
    '==>  adLongVarWChar - A long null-terminated string value (Parameter object only).
    Call BuildDataTypeTable(adLongVarWChar, "adLongVarWChar", 0, "nText", "String", "s", Null)
    '==>  adNumeric - An exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
    Call BuildDataTypeTable(adNumeric, "adNumeric", 16, "Numeric", "String", "s", Null)
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
    Call BuildDataTypeTable(adVarWChar, "adVarWChar", 0, "nVarChar/SysName", "String", "s", Null)
   '==>  adWChar - A null-terminated Unicode character string (DBTYPE_WSTR).
    Call BuildDataTypeTable(adWChar, "adWChar", 0, "nChar", "String", "s", Null)

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

Public Function grstSpSelect(ByVal pvsSpName As String, ByVal pavntParam As Variant) As adodb.Recordset
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
            For i = LBound(pavntParam) To UBound(pavntParam) - 1
                nParam = nParam + 1
                .Parameters(nParam) = pavntParam(i)
                .CreateParameter
            Next i
        End If
        
        Set grstSpSelect = .Execute
    End With
    Set adCmd = Nothing
End Function

Public Function gParmSpParam(ByVal pvsSpName As String) As adodb.Parameters
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
        Set gParmSpParam = .Parameters
    End With
    Set adCmd = Nothing
End Function

Public Function spParamList(ByVal pvsConStr As String, ByVal pvsSpName As String) As String
Dim parms As adodb.Parameters
Dim parm As adodb.Parameter
Dim bIncludeParam As Boolean

    m_sConnString = pvsConStr
    If Connect() Then
        Set parms = gParmSpParam(pvsSpName)
        If Err.Number <> 0 Then
           Exit Function
        End If
        For Each parm In parms
            With parm
                If bIncludeParam Then
                    spParamList = spParamList & .Name & " " & audtDataType(.Type).sAdoDataTypeName & " " & audtDataType(.Type).sVbDataTypeName & " " & CStr(.Size) & " " & asParamDirectionType(.Direction) & " "
                End If
            End With
            bIncludeParam = True
        Next
        Set parms = Nothing
        Set parm = Nothing
        Call Disconnect
    End If
    spParamList = spParamList & " ;"
End Function

Public Function spSelectList(ByVal pvsConStr As String, ByVal pvsSpName As String) As String
Dim nParam As Integer
Dim parms As adodb.Parameters
Dim parm As adodb.Parameter
Dim rst As adodb.Recordset
Dim fld As adodb.Field
Dim bIncludeParam As Boolean
Dim sParamList As String

    ConnString = pvsConStr
    If Connect() Then
        Set parms = gParmSpParam(pvsSpName)
        If Err.Number <> 0 Then
           Exit Function
        End If
        For Each parm In parms
            ReDim Preserve vntParams(nParam)
            With parm
                If bIncludeParam Then
                    sParamList = sParamList & .Name & " " & audtDataType(.Type).sAdoDataTypeName & " " & audtDataType(.Type).sVbDataTypeName & " " & CStr(.Size) & " " & asParamDirectionType(.Direction) & " "
                End If
                vntParams(nParam) = audtDataType(.Type).vntDefaultValue
            End With
            nParam = nParam + 1
            bIncludeParam = True
        Next
        
        Set rst = grstSpSelect(pvsSpName, vntParams)
        With rst
            For Each fld In .Fields
                With fld
                    spSelectList = spSelectList & .Name & " " & audtDataType(.Type).sVbDataTypeName & " " & CStr(.DefinedSize) & " "
                End With
            Next
        End With
        Set parms = Nothing
        Set parm = Nothing
        Set rst = Nothing
        Set fld = Nothing
        Call Disconnect
    End If
    spSelectList = sParamList & "; " & spSelectList & " ;"
End Function
