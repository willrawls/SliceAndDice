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

Private Property Get ISandyIDE_ActiveCodePane() As VBIDE.CodePane
1        On Error Resume Next
2        With vbInst
3            If .ActiveCodePane Is Nothing Then
4                If .SelectedVBComponent Is Nothing Then
5                    If .ActiveVBProject Is Nothing Then
6                        MsgBox "No projects."
7                    ElseIf .ActiveVBProject.VBComponents.Count > 0 Then
8                        .ActiveVBProject.VBComponents(1).CodeModule.CodePane.Show
9                    Else
10                       .ActiveVBProject.VBComponents.Add(vbext_ct_StdModule).CodeModule.CodePane.Show
11                   End If
12               ElseIf Not .SelectedVBComponent.CodeModule.CodePane Is Nothing Then
13                   .SelectedVBComponent.CodeModule.CodePane.Show
14               Else
15                   MsgBox "Can't do an insertion since no code pane is active.", vbInformation
16                   Exit Property
17               End If
18           End If
19           ISandyIDE_ActiveCodePane = .ActiveCodePane
20       End With
End Property


