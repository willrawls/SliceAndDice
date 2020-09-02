VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComponents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyComponents As VBIDE.VBComponents

Implements SandySupport.IComponents


Private Sub Class_Terminate()
    Set MyComponents = Nothing
End Sub

Private Function IComponents_Add(ByVal lComponentType As Long) As SandySupport.IComponent
    Dim X As IComponent

    Set X = New CComponent
        Set X.Link = MyComponents.Add(lComponentType)
        Set IComponents_Add = X
    Set X = Nothing
End Function

Private Property Get IComponents_Count() As Long
    IComponents_Count = MyComponents.Count
End Property

Private Property Get IComponents_Item(vIndex As Variant) As SandySupport.IComponent
    Dim X As IComponent

    Set X = New CComponent
        Set X.Link = MyComponents.Item(vIndex)
        Set IComponents_Item = X
    Set X = Nothing
End Property

Private Property Set IComponents_Link(ByVal RHS As Object)
    Set MyComponents = RHS
End Property

Private Property Get IComponents_Link() As Object
    Set IComponents_Link = MyComponents
End Property


Private Property Get IComponents_NewEnum() As stdole.IUnknown
    'Set IComponents_NewEnum = MyComponents.[_NewEnum]
End Property


