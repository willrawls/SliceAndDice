VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISandyEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements VBIDE.VBE

Private moVBE As VBIDE.VBE
Private moRHS As VBIDE.CodePane

Public Property Let sadEnvironment(NewEnvironment As Object)
    Set moVBE = NewEnvironment
End Property

Public Property Get sadEnvironment() As Object
    Set sadEnvironment = moVBE
End Property


Private Property Set VBE_ActiveCodePane(ByVal RHS As VBIDE.CodePane)
    Set moRHS = RHS
End Property

Private Property Get VBE_ActiveCodePane() As VBIDE.CodePane
    Set VBE_ActiveCodePane = moVBE.ActiveCodePane
End Property

Private Property Set VBE_ActiveVBProject(ByVal RHS As VBIDE.VBProject)
'
End Property

Private Property Get VBE_ActiveVBProject() As VBIDE.VBProject
'
End Property

Private Property Get VBE_ActiveWindow() As VBIDE.Window
'
End Property


Private Property Get VBE_Addins() As VBIDE.Addins
'
End Property

Private Property Get VBE_CodePanes() As VBIDE.CodePanes
'
End Property


Private Property Get VBE_CommandBars() As Office.CommandBars
'
End Property


Private Property Let VBE_DisplayModel(ByVal RHS As VBIDE.vbext_VBADisplayModel)
'
End Property

Private Property Get VBE_DisplayModel() As VBIDE.vbext_VBADisplayModel
'
End Property


Private Property Get VBE_Events() As VBIDE.Events
'
End Property

Private Property Get VBE_FullName() As String
'
End Property


Private Property Let VBE_LastUsedPath(ByVal RHS As String)
'
End Property

Private Property Get VBE_LastUsedPath() As String
'
End Property


Private Property Get VBE_MainWindow() As VBIDE.Window
'
End Property

Private Property Get VBE_Name() As String

End Property

Private Sub VBE_Quit()
'
End Sub


Private Property Let VBE_ReadOnlyMode(ByVal RHS As Integer)
'
End Property

Private Property Get VBE_ReadOnlyMode() As Integer
'
End Property


Private Property Get VBE_SelectedVBComponent() As VBIDE.VBComponent
'
End Property


Private Property Get VBE_TemplatePath() As String
'
End Property


Private Property Get VBE_VBProjects() As VBIDE.VBProjects
'
End Property


Private Property Get VBE_Version() As String
'
End Property


Private Property Get VBE_Windows() As VBIDE.Windows
'
End Property


