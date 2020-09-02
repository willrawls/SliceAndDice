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

Private Sub IComponent_Activate()
    MyComponent.Activate
End Sub

Private Property Get IComponent_CodeModule() As SandySupport.ICodeModule
    Dim X As ICodeModule
    Set X = New CCodeModule
        Set X.Link = MyComponent.CodeModule
        Set IComponent_CodeModule = X
    Set X = Nothing
End Property

Private Property Get IComponent_ComponentType() As Long
    IComponent_ComponentType = MyComponent.Type
End Property

Private Property Get IComponent_ComponentTypeDescription() As String
On Error Resume Next
    Select Case MyComponent.Type
           Case vbext_ct_StdModule:          IComponent_ComponentTypeDescription = "MODULE"
           Case vbext_ct_ClassModule:        IComponent_ComponentTypeDescription = "CLASS"
           Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm
                                             IComponent_ComponentTypeDescription = "FORM"
           Case vbext_ct_ActiveXDesigner:    IComponent_ComponentTypeDescription = "DESIGNER"
           Case vbext_ct_DocObject, vbext_ct_RelatedDocument, 100
                                             IComponent_ComponentTypeDescription = "DOCUMENT"
           Case vbext_ct_PropPage:           IComponent_ComponentTypeDescription = "PROPERTYPAGE"
           Case vbext_ct_ResFile:            IComponent_ComponentTypeDescription = "RESOURCEFILE"
           Case vbext_ct_UserControl:        IComponent_ComponentTypeDescription = "CONTROL"
    End Select
End Property

Private Property Get IComponent_Designer() As SandySupport.IDesigner
On Error Resume Next
    Dim X As IDesigner
    
    Set X = New CDesigner
        Set X.Link = MyComponent.Designer
        Set IComponent_Designer = X
    Set X = Nothing
End Property


Private Property Set IComponent_Link(ByVal RHS As Object)
    Set MyComponent = RHS
End Property

Private Property Get IComponent_Link() As Object
    Set IComponent_Link = MyComponent
End Property


Private Property Get IComponent_Name() As String
    IComponent_Name = MyComponent.Name
End Property

Private Property Let IComponent_Name(RHS As String)
    MyComponent.Name = RHS
End Property

Private Property Get IComponent_Properties(sPropertyName As String) As Variant
    IComponent_Properties = MyComponent.Properties(sPropertyName)
End Property


