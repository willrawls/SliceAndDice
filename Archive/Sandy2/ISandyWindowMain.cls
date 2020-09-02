VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISandyWindowMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public CurrentCodeArea              As Integer
Public Parent                       As ISandyWizard
Public SliceAndDice                 As CSliceAndDice
Public CurrentTemplate              As CTemplate
Public InternalCurrentTemplate      As CTemplate
Public Complete                     As CSadCommands
Public SadCommandSetCount           As Long
Public OkayToDoAction               As Boolean
Public FavoriteCalledFromIDE        As Boolean
Public OkayToUnload                 As Boolean
Public CurrentEventResponseCategory As String
'Public DBClassGen                   As ISandyWindowGen
Public Visible                      As Boolean
Public mHotKeyOpenWindow            As cRegHotKey
Public hWnd                         As Long

Public Function AddSadCommandSet(ByRef oCommands As ISadAddin) As Boolean
End Function

Public Property Get CurrentTemplateNameAndCategory() As String
End Property

Public Sub DeleteTemplate(Optional ByVal bAutoDelete As Boolean = False)
End Sub

Public Property Get ExitAfterInsert() As Boolean
End Property

Public Sub FillAddInScreen()
End Sub

Public Sub FormUnload()
End Sub

Public Sub GetCategoryAndName(ByVal sCategoryAndName As String, ByRef sCategory As String, ByRef sShortName As String)
End Sub

Public Sub HideAllWindows(Optional ByVal bUnloadAsWell = False)
End Sub

Public Function InitializeAddinDLLs(ByVal sAddinList As String) As Boolean
End Function

Public Sub NewTemplate(Optional ByVal bAutoCreate As Boolean = False, Optional ByVal sTitle As String, Optional ByVal sDefaultShortName As String, Optional ByVal bJumpToAfterCreate As Boolean = True)
End Sub

Public Sub QueueAction(ByVal sAction As String, Optional ByVal sParam As String, Optional ByVal Interval As Integer = 150)
End Sub

Public Property Let QueuedInsertions(New_QueuedInsertions As String)
End Property

Public Function RefreshDatabaseConnection() As Boolean
End Function

Public Sub DoInsertion(asaV As CAssocArray, sTemplateToInsert As String, Optional ByVal bSkipDeclarations As Boolean = False)
End Sub

Public Function FillTemplateWithUserInput(ByRef asaX As CAssocArray, ByVal sToParse As String, ByRef sCodeToInsert As String, ByVal sMsgBoxTitle As String) As Boolean
End Function

Public Function InternalInsertTemplate(II As CInsertionInfo) As Boolean
End Function

Public Sub GetProcAtLine(ByRef lCurrentLine As Long, ByRef sProcName As String, ByRef lProcType As Long)
End Sub

Public Function FindLastProcLine(sProcName As String, lProcType As Long) As Long
End Function

Public Function JumpTo(ByVal sTemplateName As String, Optional ByVal bRecordInHistory As Boolean = True, Optional ByVal bSyncCategoryList As Boolean = False) As Boolean
End Function

Public Sub RefillList()
End Sub

Private Sub SetColors(ByVal BackColor As String, ByVal ForeColor As String)
End Sub

Public Function SetInternalCurrentTemplate(ByVal sTemplateName As String) As Boolean
End Function

Public Function sGetCurrentLineAtCharacter(ByVal sTextToSearch As String, ByVal lCharToStart As Long) As String
End Function

Public Sub ShowExternalsMenu()
End Sub

Public Sub ShowFavMenu()
End Sub

Public Function ShutdownDLLs() As Boolean
End Function

Public Property Get TemplateDatabaseName() As String
End Property

Public Sub SaveTemplate()
End Sub

Public Sub UpdateFavorites()
End Sub

Public Sub UpdateHotKeys()
End Sub

Public Function sChooseDatabase(Optional ByVal sPath As String, Optional ByVal sFilename As String) As String
End Function

Public Function sChooseFile(Optional ByVal sPath As String, Optional ByVal sFilename As String, Optional ByVal sFilter As String = vbNullString) As String
End Function

Public Function sChooseColor(Optional ByVal sInitialColor As String) As String
End Function

Public Sub FindInCurrent(Optional ByVal bRepeatLastSearch As Boolean = False, Optional ByVal bReplace As Boolean = False, Optional ByVal bAuto As Boolean = False)
End Sub

Public Function sPropertyType(sFieldType As String) As String
End Function

Public Function GetCurrentTextSelection() As String
End Function

Public Sub DeleteCurrentTextSelection()
End Sub

Public Function DetermineLastLineInSelection() As Long
End Function

Public Function DetermineFirstLineInSelection() As Long
End Function

Public Function DetermineFirstColumnInSelection() As Long
End Function

Public Function DetermineLastColumnInSelection() As Long
End Function

Public Sub Show(Optional ByVal ModalSetting As Integer, Optional ParentWindow As Object)
End Sub

Public Sub Hide()
End Sub

Public Sub ZOrder()
End Sub

Public Sub SetFocus()
End Sub

