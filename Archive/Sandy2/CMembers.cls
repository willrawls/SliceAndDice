VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyMembers As VBIDE.Members

Implements SandySupport.IMembers

Private Property Get IMembers_Count() As Long
    IMembers_Count = MyMembers.Count
End Property

Private Property Get IMembers_Item(vIndex As Variant) As SandySupport.IMember
On Error Resume Next
    Dim X As IMember
    Set X = New CMember
        Set X.Link = MyMembers.Item(vIndex)
        Set IMembers_Item = X
    Set X = Nothing
End Property


Private Property Set IMembers_Link(ByVal RHS As Object)
    Set MyMembers = RHS
End Property

Private Property Get IMembers_Link() As Object
    Set IMembers_Link = MyMembers
End Property


Private Property Get IMembers_NewEnum() As stdole.IUnknown
    Set IMembers_NewEnum = MyMembers.[_NewEnum]
End Property


