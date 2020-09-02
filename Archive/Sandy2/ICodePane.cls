VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICodePane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Parent As Object
Public Link   As Object

Public Property Get Window() As Object
End Property

Public Property Get CodeModule() As SandySupport.ICodeModule
End Property

Public Sub GetSelection(ByRef lLine As Long, ByRef lFirstCol As Long, ByRef lLastLine As Long, ByRef lLastCol As Long)
End Sub

Public Sub Show()
End Sub
