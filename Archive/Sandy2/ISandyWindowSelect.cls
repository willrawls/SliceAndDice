VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISandyWindowSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Choices As New CAssocArray
Public Key As String

Public Property Get Choice() As String
End Property

Public Sub Initialize(sChoices As String, Optional ByVal sDelimiter As String = ";", Optional ByVal sDefault As String)
End Sub

Public Sub Show(Optional ByVal ModalSetting As Integer, Optional ParentWindow As Object)
End Sub

Public Sub Hide()
End Sub

Public Sub ZOrder()
End Sub


