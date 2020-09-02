VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyProjects As VBIDE.VBProjects

Implements IProjects

Private Sub Class_Terminate()
    Set MyProjects = Nothing
End Sub


Private Function IProjects_Add(ByVal lProjectType As Long, Optional ByVal Exclusive As Boolean = False) As SandySupport.IProject
    Dim X As IProject

    Set X = New CProject
        Set X.Link = MyProjects.Add(lProjectType, Exclusive)
        Set IProjects_Add = X
    Set X = Nothing
End Function

Private Property Get IProjects_Count() As Long
    IProjects_Count = MyProjects.Count
End Property

Private Property Get IProjects_Item(vIndex As Variant) As SandySupport.IProject
    Dim X As IProject

    Set X = New CProject
        Set X.Link = MyProjects.Item(vIndex)
        Set IProjects_Item = X
    Set X = Nothing
End Property

Private Property Set IProjects_Link(ByVal RHS As Object)
    Set MyProjects = RHS
End Property

Private Property Get IProjects_Link() As Object
    Set IProjects_Link = MyProjects
End Property

