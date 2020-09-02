VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsertionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy - InsertionInfo"
Option Explicit

Public TemplateName             As String

Public OriginalCodeToInsert     As String

Public ExternalFilename         As String
Public TextToSendToFile         As String

Public SoftCommandName          As String
Public LinesLeftToProcess       As String
Public PointOfInsertion         As Long
Public CurrentLineToProcess     As String
Public Result                   As String
Public Expression               As String

Public sParam                   As String
Public AllParameters            As String

Public ParamLineOffset          As Long

Public SoftVars                 As SandySupport.CAssocArray
Public CurrControl              As SandySupport.IControl

Private Sub Class_Initialize()

    ' LogEvent "CInsertionInfo: Initialize"
End Sub


Private Sub Class_Terminate()
    Set CurrControl = Nothing
    Set SoftVars = Nothing
    ' LogEvent "CInsertionInfo: Terminate"
End Sub

