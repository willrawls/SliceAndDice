VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Link As Object

Public Sub Activate()
End Sub

Public Property Get CodeModule() As SandySupport.ICodeModule
End Property

Public Property Get Designer() As SandySupport.IDesigner
End Property


Public Property Let Name(RHS As String)
End Property

Public Property Get Name() As String
End Property

Public Property Get Properties(sPropertyName As String) As Variant
End Property

Public Property Get ComponentType() As Long
End Property

Public Property Get ComponentTypeDescription() As String
End Property


