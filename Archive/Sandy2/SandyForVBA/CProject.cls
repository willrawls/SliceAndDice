VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyProject As VBIDE.VBProject

Implements SandySupport.IProject

Private Sub Class_Terminate()
    Set MyProject = Nothing
End Sub


Private Sub IProject_AddToolboxProgID(ByVal sProgID As String, Optional ByVal sFilename As String)
On Error Resume Next
    'If IsMissing(sFilename) Then
       MyProject.References.AddFromFile sProgID
    'Else
    '   MyProject.AddToolboxProgID sProgID, sFilename
    'End If
End Sub

Private Property Set IProject_Link(ByVal RHS As Object)
    Set MyProject = RHS
End Property

Private Property Get IProject_Link() As Object
    Set IProject_Link = MyProject
End Property


Private Property Let IProject_Name(ByVal RHS As String)
    MyProject.Name = RHS
End Property

Private Property Get IProject_Name() As String
    IProject_Name = MyProject.Name
End Property


Private Property Get IProject_References() As SandySupport.IReferences
On Error Resume Next
    Dim X As IReferences

    Set X = New CReferences
        Set X.Link = MyProject.References
        Set IProject_References = X
    Set X = Nothing
End Property


