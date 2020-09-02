VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyForm As VBForm

Implements SandySupport.IForm

Private Sub Class_Terminate()
    Set MyForm = Nothing
End Sub


Private Property Get IForm_Controls() As SandySupport.IControls
    Dim X As IControls

    Set X = New CControls
        Set X.Link = MyForm.VBControls
        Set IForm_Controls = X
    Set X = Nothing
End Property

Private Property Set IForm_Link(ByVal RHS As Object)
    Set MyForm = RHS
End Property

Private Property Get IForm_Link() As Object
    Set IForm_Link = MyForm
End Property


