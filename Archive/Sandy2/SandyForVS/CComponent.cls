VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyComponent As VBIDE.VBComponent

Implements SandySupport.IComponent

Private Sub Class_Terminate()
    Set MyComponent = Nothing
End Sub


Private Property Get IComponent_CodeModule() As SandySupport.ICodeModule
    Dim X As ICodeModule
    Set X = New CCodeModule
        Set X.Link = MyComponent.CodeModule
        Set IComponent_CodeModule = X
    Set X = Nothing
End Property

Private Property Get IComponent_Designer() As SandySupport.IDesigner
    Set IComponent_Designer = MyComponent.Designer
End Property


Private Property Set IComponent_Link(ByVal RHS As Object)
    Set MyComponent = RHS
End Property

Private Property Get IComponent_Link() As Object
    Set IComponent_Link = MyComponent
End Property


Private Property Get IComponent_Properties(sPropertyName As String) As Variant
    IComponent_Properties = MyComponent.Properties(sPropertyName)
End Property


