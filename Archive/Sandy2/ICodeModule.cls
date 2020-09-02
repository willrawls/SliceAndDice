VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICodeModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Parent As Object
Public Link   As Object

Public Function Members() As SandySupport.IMembers
End Function

Public Property Get CodePane() As ICodePane
End Property

Public Function ProcBodyLine(ByVal sProcName As String, ByVal lProcType As Long) As Long
End Function

Public Function Lines(ByVal lStartLine As Long, ByVal lLineCount As Long) As String
End Function

Public Function CountOfDeclarationLines() As Long
End Function

Public Function CountOfLines() As Long
End Function

Public Function ProcOfLine(ByVal lLineNumberOfProc As Long, ByRef lProcType As Long) As String
End Function

Public Function ProcStartLine(ByVal sProcName As String, ByVal lProcType As Long) As Long
End Function

Public Function CreateEventProc(ByVal Eventname As String, ByVal ObjectName As String) As Long
End Function

Public Function ProcCountLines(ByVal sProcName As String, ByVal lProcType As Long) As Long
End Function

Public Sub DeleteLines(ByVal lStartLine As Long, Optional ByVal lLineCount As Long = 1)
End Sub

Public Sub InsertLines(ByVal lStartLine As Long, ByVal LinesToInsert As String)
End Sub

Public Function Find(ByVal TextToFind As String, ByVal lStartLine As Long, ByVal lStartColFound As Long, ByVal lEndLine As Long, lEndColFound As Long, Optional ByVal WholeWord As Boolean, Optional ByVal MatchCase As Boolean, Optional ByVal PatternSearch As Boolean) As Long
End Function

Public Sub AddFromFile(sFilename As String)
End Sub
