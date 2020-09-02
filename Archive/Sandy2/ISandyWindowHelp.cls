VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISandyWindowHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SadCommandSet    As CSadCommands
Public CurrCommand      As CSadCommand
Public vCurrCommandKey  As Variant
Public BackColor        As OLE_COLOR
Public ForeColor        As OLE_COLOR

Public Property Let CurrCommandKey(ByVal vKey As Variant)
End Property

Public Sub Populate()
End Sub

Public Sub FileExit()
End Sub

Public Sub Show(Optional ByVal ModalSetting As Integer, Optional ParentWindow As Object)
End Sub

Public Sub Hide()
End Sub

Public Sub ZOrder()
End Sub

