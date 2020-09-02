VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyReferences As VBIDE.References

Implements SandySupport.IReferences

Private Sub Class_Terminate()
    Set MyReferences = Nothing
End Sub


Private Function IReferences_AddFromFile(ByVal sFilename As String) As SandySupport.IReference
    Dim X As IReference

    Set X = New CReference
        Set X.Link = MyReferences.AddFromFile(sFilename)
        Set IReferences_AddFromFile = X
    Set X = Nothing
End Function

Private Function IReferences_AddFromGuid(ByVal sGUID As String, ByVal Major As Long, ByVal Minor As Long) As SandySupport.IReference
    Dim X As IReference

    Set X = New CReference
        Set X.Link = MyReferences.AddFromGuid(sGUID, Major, Minor)
        Set IReferences_AddFromGuid = X
    Set X = Nothing
End Function


Private Property Get IReferences_Item(vIndex As Variant) As SandySupport.IReference
    Dim X As IReference

    Set X = New CReference
        Set X.Link = MyReferences(vIndex)
        Set IReferences_Item = X
    Set X = Nothing
End Property

Private Property Set IReferences_Link(ByVal RHS As Object)
    Set MyReferences = RHS
End Property

Private Property Get IReferences_Link() As Object
    Set IReferences_Link = MyReferences
End Property


Private Property Get IReferences_NewEnum() As stdole.IUnknown
    Dim X As IReference

    Set X = New CReference
        Set X.Link = MyReferences.[_NewEnum]
        Set IReferences_NewEnum = X
    Set X = Nothing
End Property


