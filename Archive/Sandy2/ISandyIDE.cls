VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISandyIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public ActiveProject        As SandySupport.IProject
Public ActiveWindow         As Object

Public SelectedComponent    As SandySupport.IComponent
Public Components           As Object

Public Projects             As SandySupport.IProjects
Public Property Get ActiveCodeModule() As SandySupport.ICodeModule
End Property


Public Property Get ActiveCodePane() As SandySupport.ICodePane
End Property


Public Function AddComponent(ComponentType As Long) As IComponent
End Function

