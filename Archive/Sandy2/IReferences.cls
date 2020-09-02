VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Link As Object

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
End Property

Public Property Get Item(vIndex As Variant) As IReference
Attribute Item.VB_UserMemId = 0
End Property

Public Function AddFromFile(ByVal sFilename As String) As IReference
End Function

Public Function AddFromGuid(ByVal sGUID As String, ByVal Major As Long, ByVal Minor As Long) As IReference
End Function

