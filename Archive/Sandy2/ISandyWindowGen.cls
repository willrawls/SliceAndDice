VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISandyWindowGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Parent As ISandyWindowMain

'Public Sub GenerateChildren(ByRef asaPass As CAssocArray, sDataLibraryType As String, nodChild As Node)
Public Sub GenerateChildren(ByRef asaPass As CAssocArray, sDataLibraryType As String, nodChild As Object)
End Sub

'Public Sub GenerateClass(ByRef asaPass As CAssocArray, sDataLibraryType As String, tvwTables As TreeView, lvwFields As ListView)
Public Sub GenerateClass(ByRef asaPass As CAssocArray, sDataLibraryType As String, tvwTables As Object, lvwFields As Object)
End Sub

Public Sub TriggerClassGeneration()
End Sub

Public Sub AddTable(sTableName As String, sParentTable As String)
End Sub

Public Property Get Canceled() As Boolean
End Property

Public Property Get ConnectString() As String
End Property

Public Property Get DBName() As String
End Property

Public Property Get DBPathAndFilename() As String
End Property

Public Property Get ODBCDatabaseName() As String
End Property

Public Sub UpdateFavorites()
End Sub

Public Property Get GenerateBranch() As Boolean
End Property

Public Property Get GenerateDatabase() As Boolean
End Property

Public Sub RefreshCategories()
End Sub

Public Sub PopulateTree()
End Sub

Public Function sFieldType(iFieldType As Long) As String
End Function

Public Sub OpenFile()
End Sub

Public Sub NodeClick(ByRef Node As MSComctlLib.Node)
End Sub

Public Sub Show(Optional ByVal ModalSetting As Integer, Optional ParentWindow As Object)
End Sub

Public Sub Hide()
End Sub

Public Sub ZOrder()
End Sub

Public Sub SetColors(ByVal ForeColor As Long, ByVal BackColor As Long)
End Sub
