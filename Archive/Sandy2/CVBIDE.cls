VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVBIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private vbInst As VBIDE.VBE

Implements SliceAndDice.ISandyIDE

Private Sub Class_Initialize()

    ' LogEvent "CVBIDE: Initialize"
End Sub


Private Sub Class_Terminate()

    ' LogEvent "CVBIDE: Terminate"
End Sub


Private Property Get ISandyIDE_ActiveCodePane() As VBIDE.CodePane
On Error Resume Next
    With vbInst
         If .ActiveCodePane Is Nothing Then
            If .SelectedVBComponent Is Nothing Then
               If .ActiveVBProject Is Nothing Then
                  MsgBox "No projects."
               ElseIf .ActiveVBProject.VBComponents.Count > 0 Then
                  .ActiveVBProject.VBComponents(1).CodeModule.CodePane.Show
               Else
                  .ActiveVBProject.VBComponents.Add(vbext_ct_StdModule).CodeModule.CodePane.Show
               End If
            ElseIf Not .SelectedVBComponent.CodeModule.CodePane Is Nothing Then
               .SelectedVBComponent.CodeModule.CodePane.Show
            Else
               MsgBox "Can't do an insertion since no code pane is active.", vbInformation
               Exit Property
            End If
         End If
         ISandyIDE_ActiveCodePane = .ActiveCodePane
    End With
End Property


