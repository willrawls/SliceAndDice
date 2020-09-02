VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Link As Object

Public Property Get Item(vIndex As Variant) As IProject
Attribute Item.VB_UserMemId = 0
End Property

Public Property Get Count() As Long
End Property

Public Function Add(ByVal lProjectType As Long, Optional ByVal Exclusive As Boolean = False) As IProject
End Function

