VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyReference As VBIDE.Reference

Implements SandySupport.IReference

Private Sub Class_Terminate()
    Set MyReference = Nothing
End Sub


Private Property Get IReference_FullPath() As String
    IReference_FullPath = MyReference.FullPath
End Property


Private Property Set IReference_Link(ByVal RHS As Object)
    Set MyReference = RHS
End Property

Private Property Get IReference_Link() As Object
    Set IReference_Link = MyReference
End Property


