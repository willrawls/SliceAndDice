VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAutoWebBrowsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection
Public Parent As frmMain
'Public KEY_FIELD_NAME As String

'Implements IObjectSafety
'
'Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByVal riid As Long, pdwSupportedOptions As Long, pdwEnabledOptions As Long)
'    modIObjectSafety.GetInterfaceSafetyOptions riid, pdwSupportedOptions, pdwEnabledOptions
'End Sub
'
'Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByVal riid As Long, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
'    modIObjectSafety.SetInterfaceSafetyOptions riid, dwOptionsSetMask, dwEnabledOptions
'End Sub

Public Function Add(ByVal sKey As String) As frmBrowser
    Dim objNewMember As frmBrowser

    Set objNewMember = New frmBrowser
        Set objNewMember.Parent = Me
        If Len(sKey) Then
           mCol.Add objNewMember, UCase(sKey)
        Else
           mCol.Add objNewMember
        End If
        objNewMember.Index = mCol.Count
        Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As frmBrowser
Attribute Item.VB_UserMemId = 0
On Error GoTo EH_Item
    Dim NewItem As frmBrowser

    Set Item = Nothing
    If VarType(vntIndexKey) = vbString Then
       If Len(vntIndexKey) = 0 Then Exit Property
       Set Item = mCol(UCase(vntIndexKey))
    Else
       Set Item = mCol(vntIndexKey)
    End If
    
EH_Item_Continue:
    Set NewItem = Nothing
    Exit Property

EH_Item:
'On Error Resume Next
'    Set NewItem = New frmBrowser
'    Set NewItem.Parent = Me
'    If NewItem.ReadRecord(CStr(vntIndexKey)) Then
'       mCol.Add NewItem, UCase(CStr(vntIndexKey))
'       Set Item = NewItem
'    Else
'       Set NewItem.Parent = Nothing
'    End If
    Resume EH_Item_Continue
End Property

Public Property Get Count() As Long
On Error Resume Next
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
On Error Resume Next
    If VarType(vntIndexKey) = vbString Then
       If Not mCol(UCase(vntIndexKey)) Is Nothing Then
          Set mCol(UCase(vntIndexKey)).Parent = Nothing
          mCol.Remove UCase(vntIndexKey)
       End If
    Else
       If Not mCol(vntIndexKey) Is Nothing Then
          Set mCol(vntIndexKey).Parent = Nothing
          mCol.Remove vntIndexKey
       End If
    End If
End Sub

Public Sub Clear(Optional ByVal bRecreateCollection As Boolean = True)
On Error Resume Next
    Dim CurrfrmBrowser As frmBrowser
    For Each CurrfrmBrowser In mCol
        Set CurrfrmBrowser.Parent = Nothing
        Unload CurrfrmBrowser
    Next CurrfrmBrowser
    Set mCol = Nothing
    If bRecreateCollection Then Set mCol = New Collection
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
On Error Resume Next
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
On Error Resume Next
   'KEY_FIELD_NAME = ???
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Clear False
    Set Parent = Nothing
End Sub

Public Function ReadRecords() As Boolean
'On Error Resume Next
'    Dim sSQL As String
'    Dim rst As RDO.rdoResultset
'    Dim CurrfrmBrowser As frmBrowser
'    Dim bOkaySoFar As Boolean
'
'    Clear
'
'    sSQL = "select * from frmBrowser;" & Chr(13) & Chr(10)
'    Set rst = Parent.Oracle.OpenResultset(sSQL, rdOpenForwardOnly, rdConcurReadOnly)
'    If RecordsetHasRecords(rst) Then
'       bOkaySoFar = True
'       With rst
'            Do Until .EOF Or (Not bOkaySoFar)
'               Add("" & .rdoColumns(KEY_FIELD_NAME)).LoadProperties rst
'               .MoveNext
'            Loop
'       End With
'       Set CurrfrmBrowser = Nothing
'       ReadRecords = bOkaySoFar
'    Else
'       ReadRecords = True
'    End If
'    rst.Close
End Function
