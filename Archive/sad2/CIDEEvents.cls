VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIDEEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents FileControlEventsHandler As FileControlEvents
Attribute FileControlEventsHandler.VB_VarHelpID = -1
Private WithEvents ReferencesEventsHandler As ReferencesEvents
Attribute ReferencesEventsHandler.VB_VarHelpID = -1
Private WithEvents SelectedVBControlsEventsHandler As SelectedVBControlsEvents
Attribute SelectedVBControlsEventsHandler.VB_VarHelpID = -1
Private WithEvents VBComponentsEventsHandler As VBComponentsEvents
Attribute VBComponentsEventsHandler.VB_VarHelpID = -1
Private WithEvents VBControlsEventsHanlder As VBControlsEvents
Attribute VBControlsEventsHanlder.VB_VarHelpID = -1
Private WithEvents VBProjectsEventsHandler As VBProjectsEvents
Attribute VBProjectsEventsHandler.VB_VarHelpID = -1

Private m_oUI As frmMain
Public Sub Init(ByRef IDE As VBIDE.VBE, UI As frmMain)
    Set m_oUI = UI
    Set FileControlEventsHandler = IDE.Events.FileControlEvents(Nothing)
    Set ReferencesEventsHandler = IDE.Events.ReferencesEvents(Nothing)
    Set SelectedVBControlsEventsHandler = IDE.Events.SelectedVBControlsEvents(Nothing, Nothing)
    Set VBComponentsEventsHandler = IDE.Events.VBComponentsEvents(Nothing)
    Set VBControlsEventsHanlder = IDE.Events.VBControlsEvents(Nothing, Nothing)
    Set VBProjectsEventsHandler = IDE.Events.VBProjectsEvents
End Sub


Private Sub Class_Terminate()
    Set m_oUI = Nothing

    Set FileControlEventsHandler = Nothing
    Set ReferencesEventsHandler = Nothing
    Set SelectedVBControlsEventsHandler = Nothing
    Set VBComponentsEventsHandler = Nothing
    Set VBControlsEventsHanlder = Nothing
    Set VBProjectsEventsHandler = Nothing
End Sub


Private Sub FileControlEventsHandler_AfterAddFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
    m_oUI.HandleIDEEvents "File - After Add", VBProject
End Sub

Private Sub FileControlEventsHandler_AfterChangeFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal NewName As String, ByVal OldName As String)
    m_oUI.HandleIDEEvents "File - After Name Change", VBProject
End Sub

Private Sub FileControlEventsHandler_AfterCloseFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal WasDirty As Boolean)
    m_oUI.HandleIDEEvents "File - After Close", VBProject
End Sub


Private Sub FileControlEventsHandler_AfterRemoveFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
    m_oUI.HandleIDEEvents "File - After Remove", VBProject
End Sub


Private Sub FileControlEventsHandler_AfterWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal Result As Integer)
    m_oUI.HandleIDEEvents "File - After Write", VBProject
End Sub


Private Sub FileControlEventsHandler_BeforeLoadFile(ByVal VBProject As VBIDE.VBProject, FileNames() As String)
    m_oUI.HandleIDEEvents "File - Before Load", VBProject
End Sub


Private Sub ReferencesEventsHandler_ItemAdded(ByVal Reference As VBIDE.Reference)
    m_oUI.HandleIDEEvents "References - Item Added"
End Sub

Private Sub ReferencesEventsHandler_ItemRemoved(ByVal Reference As VBIDE.Reference)
    m_oUI.HandleIDEEvents "References - Item Removed"
End Sub


Private Sub VBComponentsEventsHandler_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
    m_oUI.HandleIDEEvents "Components - Item Added", , VBComponent
End Sub


Private Sub VBComponentsEventsHandler_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
    m_oUI.HandleIDEEvents "Components - Item Removed", , VBComponent
End Sub


Private Sub VBComponentsEventsHandler_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
    m_oUI.HandleIDEEvents "Components - Item Renamed", , VBComponent
End Sub


Private Sub VBComponentsEventsHandler_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    m_oUI.HandleIDEEvents "Components - Item Selected", , VBComponent
End Sub


Private Sub VBProjectsEventsHandler_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    m_oUI.HandleIDEEvents "Project - Item Activated", VBProject
End Sub

Private Sub VBProjectsEventsHandler_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    m_oUI.HandleIDEEvents "Project - Item Added", VBProject
End Sub


Private Sub VBProjectsEventsHandler_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    m_oUI.HandleIDEEvents "Project - Item Removed", VBProject
End Sub


Private Sub VBProjectsEventsHandler_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
    m_oUI.HandleIDEEvents "Project - Item Renamed", VBProject
End Sub



