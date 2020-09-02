VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDELink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy 2 - Slice and Dice for VB6"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private SandyIDE As CSandyIDE

Implements IDTExtensibility

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
    ' Comment to prevent procedure from being deleted on compilation.
End Sub

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
On Error Resume Next
    Set SandyIDE = CreateObject("SandyIDEforVB6.CSandyIDE")
    If SandyIDE Is Nothing Then
       MsgBox "Failed to create a 'SandyIDEforVB6.CSandyIDE' object. Can't start Slice and Dice."
       Exit Sub
    End If

    SandyIDE.OnConnection VBInst
End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)
On Error Resume Next
    If Not SandyIDE Is Nothing Then
       SandyIDE.OnDisconnection
       Set SandyIDE = Nothing
    End If
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
  ' Comment to prevent procedure from being deleted on compilation.
End Sub
