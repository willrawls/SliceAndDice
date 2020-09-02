VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyDesigner As Object

Implements SandySupport.IDesigner

Private Sub Class_Terminate()
    Set MyDesigner = Nothing
End Sub


Private Property Get IDesigner_Controls() As SandySupport.IControls
    Dim X As IControls

    Set X = New CControls
        Set X.Link = MyDesigner.VBControls
        Set IDesigner_Controls = X
    Set X = Nothing
End Property

Private Property Set IDesigner_Link(ByVal RHS As Object)
    Set MyDesigner = RHS
End Property

Private Property Get IDesigner_Link() As Object
    Set IDesigner_Link = MyDesigner
End Property


