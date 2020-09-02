VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSandyWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function CreateForm(sFormType As String) As Object
    Select Case UCase$(sFormType)
           Case "HELP", "COMMAND", "REFERENCE", "COMMANDREFERENCE", "COMMAND REFERENCE"
                Set CreateForm = New frmCommandHelp
           Case "FIND", "FINDREPLACE", "REPLACE"
                Set CreateForm = New frmFindReplace
           Case "LIST", "SELECT", "LISTSELECT"
                Set CreateForm = New frmListSelect
           Case "MAIN", "SANDY", "SLICEANDDICE", "SLICE AND DICE"
                Set CreateForm = New frmMain
    End Select
End Function


