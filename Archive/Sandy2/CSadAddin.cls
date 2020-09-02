VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISadAddin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy - Addin Abstract Class"
Option Explicit

Public Property Get CommandSet() As Object
End Property

Public Property Get Externals() As CAssocArray
End Property

Public Function ExecuteSoftCommand(II As SandySupport.CInsertionInfo) As Boolean
End Function

Public Function Shutdown() As Boolean
End Function

Public Function Startup(ByRef oSliceAndDiceWizard As Object, Application As Object) As Boolean
End Function

Public Function ExecuteSoftCommandInline(ByRef SoftVars As SandySupport.CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, ByRef sResults As String) As Boolean
End Function

Public Function ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
End Function

