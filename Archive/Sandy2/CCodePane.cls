VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCodePane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyPane As VBIDE.CodePane
Private Parent As Object

Implements SandySupport.ICodePane

Private Sub Class_Terminate()
    Set Parent = Nothing
    Set MyPane = Nothing
End Sub


Private Property Get ICodePane_CodeModule() As SandySupport.ICodeModule
    Dim x As ICodeModule
    Set x = New CCodeModule
    Set x.Link = MyPane.CodeModule
    Set ICodePane_CodeModule = x
    Set x = Nothing
End Property

Private Sub ICodePane_GetSelection(lLine As Long, lFirstCol As Long, lLastLine As Long, lLastCol As Long)
    MyPane.GetSelection lLine, lFirstCol, lLastLine, lLastCol
End Sub


Private Property Set ICodePane_Link(ByVal RHS As Object)
    Set MyPane = RHS
End Property

Private Property Get ICodePane_Link() As Object
    Set ICodePane_Link = MyPane
End Property


Private Property Set ICodePane_Parent(ByVal RHS As Object)
    Set Parent = RHS
End Property

Private Property Get ICodePane_Parent() As Object
    Set ICodePane_Parent = Parent
End Property


Private Sub ICodePane_Show()
    MyPane.Show
End Sub

Private Property Get ICodePane_Window() As Object
    Set ICodePane_Window = MyPane.Window
End Property


