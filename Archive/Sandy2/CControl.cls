VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyControl As VBIDE.VBControl

Implements SandySupport.IControl

Private Sub Class_Terminate()
    Set MyControl = Nothing
End Sub


Private Property Get IControl_ClassName() As String
    IControl_ClassName = MyControl.ClassName
End Property

Private Property Get IControl_ContainedControls() As SandySupport.IControls
    Dim X As IControls

    Set X = New CControls
        Set X.Link = MyControl.ContainedVBControls
        Set IControl_ContainedControls = X
    Set X = Nothing
End Property


Private Property Set IControl_Link(ByVal RHS As Object)
    Set MyControl = RHS
End Property

Private Property Get IControl_Link() As Object
    Set IControl_Link = MyControl
End Property


Private Property Get IControl_Properties(vIndex As Variant) As Variant
    IControl_Properties = MyControl.Properties(vIndex)
End Property


