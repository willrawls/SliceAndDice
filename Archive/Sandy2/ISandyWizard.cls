VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISandyWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public InsertionCancelled As Boolean

Public Function JumpTo(ByVal sTemplateName As String, Optional ByVal bRecordInHistory As Boolean = True, Optional ByVal bSyncCategoryList As Boolean = False) As Boolean
End Function

Public Sub NewTemplate(Optional ByVal bAutoCreate As Boolean = False, Optional ByVal sTitle As String, Optional ByVal sDefaultShortName As String, Optional ByVal bJumpToAfterCreate As Boolean = True)
End Sub

Public Function sChooseColor(Optional ByVal sInitialColor As String) As String
End Function

Public Function sChooseFile(Optional ByVal sPath As String, Optional ByVal sFilename As String, Optional ByVal sFilter As String) As String
End Function

Public Sub ShowSplashScreen()
End Sub

Public Property Get TemplateDatabaseName() As String
End Property

Public Function GetATemplate(sCategory As String, sTemplate As String) As CTemplate
End Function

Public Function sGetGUID(ByVal sProgID As String) As String
End Function

Public Function sChoose(sChoices As String, Optional ByVal sDelimiter As String = ";", Optional ByVal sDefault As String)
End Function

Public Property Get TemplateDatabasePath() As String
End Property

Public Property Get Version() As String
End Property

Public Sub DeleteCurrentTextSelection()
End Sub

Public Function DetermineFirstLineInSelection() As Long
End Function

Public Function DetermineLastLineInSelection() As Long
End Function

Public Sub DoInsertion(asaV As CAssocArray, sTemplateToInsert As String, Optional ByVal bSkipDeclarations As Boolean = False)
End Sub

Public Function FillTemplateWithUserInput(ByRef asaX As CAssocArray, ByVal sToParse As String, ByRef sCodeToInsert As String, ByVal sMsgBoxTitle As String) As Boolean
End Function

Public Function FindLastProcLine(sProcName As String, lProcType As Long) As Long
End Function

Public Sub GetCategoryAndName(ByVal sCategoryAndName As String, ByRef sCategory As String, ByRef sShortName As String)
End Sub

Public Function GetCurrentTextSelection() As String
End Function

Public Sub GetProcAtLine(ByVal lCurrentLine As Long, ByRef sProcName As String, ByRef lProcType As Long)
End Sub

Public Function InsertTemplate(ByVal sToParseIn As String, Optional ByVal lLineIn As Long = 1, Optional ByRef SoftVars As CAssocArray, Optional ByVal sTemplateName As String = "Slice and Dice", Optional ByVal sExternalFilename As String = vbNullString) As Boolean
End Function

Public Property Get InternalCurrentTemplate() As CTemplate
End Property

Public Property Get CurrentTemplate() As CTemplate
End Property

Public Function sChooseDatabase(Optional ByVal sPath As String, Optional ByVal sFilename As String) As String
End Function

Public Property Get SliceAndDice() As CSliceAndDice
End Property

Public Function sPropertyType(sFieldType As String) As String
End Function

Public Property Get UI() As ISandyWindowMain
End Property

Public Property Get SandyIDE() As ISandyIDE
End Property

Public Property Get CurrentTemplateNameAndCategory() As String
End Property

Public Function OnConnection(xIDE As SandySupport.ISandyIDE) As Boolean
End Function

Public Sub OnDisconnection()
End Sub

Public Function sTableToPropertyName(ByVal sTableName As String) As String
End Function

