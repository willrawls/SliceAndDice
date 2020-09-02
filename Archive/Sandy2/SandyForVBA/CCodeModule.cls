VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCodeModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Parent   As Object
Private MyModule As VBIDE.CodeModule
Attribute MyModule.VB_VarHelpID = -1

Implements SandySupport.ICodeModule

Private Sub Class_Terminate()
    Set Parent = Nothing
    Set MyModule = Nothing
End Sub


Private Sub ICodeModule_AddFromFile(sFilename As String)
    MyModule.AddFromFile sFilename
End Sub

Private Property Get ICodeModule_CodePane() As SandySupport.ICodePane
    Dim X As ICodePane

    Set X = New CCodePane
        Set X.Link = MyModule.CodePane
        Set ICodeModule_CodePane = X
    Set X = Nothing
End Property


Private Function ICodeModule_CountOfDeclarationLines() As Long
    ICodeModule_CountOfDeclarationLines = MyModule.CountOfDeclarationLines
End Function

Private Function ICodeModule_CountOfLines() As Long
    ICodeModule_CountOfLines = MyModule.CountOfLines
End Function


Private Function ICodeModule_CreateEventProc(ByVal Eventname As String, ByVal ObjectName As String) As Long
    If ObjectName = "Form" Then ObjectName = "UserForm"
    If Eventname = "Load" Then Eventname = "Initialize"
    If Eventname = "Unload" Then Eventname = "Terminate"
    ICodeModule_CreateEventProc = MyModule.CreateEventProc(Eventname, ObjectName)
End Function


Private Sub ICodeModule_DeleteLines(ByVal lStartLine As Long, Optional ByVal lLineCount As Long = 1)
    MyModule.DeleteLines lStartLine, lLineCount
End Sub


Private Function ICodeModule_Find(ByVal TextToFind As String, ByVal lStartLine As Long, ByVal lStartColFound As Long, ByVal lEndLine As Long, lEndColFound As Long, Optional ByVal WholeWord As Boolean, Optional ByVal MatchCase As Boolean, Optional ByVal PatternSearch As Boolean) As Long
    ICodeModule_Find = MyModule.Find(TextToFind, lStartLine, lStartColFound, lEndLine, lEndColFound, WholeWord, MatchCase, PatternSearch)
End Function

Private Sub ICodeModule_InsertLines(ByVal lStartLine As Long, ByVal LinesToInsert As String)
    MyModule.InsertLines lStartLine, LinesToInsert
End Sub


Private Function ICodeModule_Lines(ByVal lStartLine As Long, ByVal lLineCount As Long) As String
On Error Resume Next
    ICodeModule_Lines = MyModule.Lines(lStartLine, lLineCount)
End Function


Private Property Set ICodeModule_Link(ByVal RHS As Object)
    Set MyModule = RHS
End Property

Private Property Get ICodeModule_Link() As Object
    Set ICodeModule_Link = MyModule
End Property

Private Function ICodeModule_Members(ByVal sMemberName As String) As Object
   'Set ICodeModule_Members = MyModule.Members(sMemberName)
End Function

Private Property Set ICodeModule_Parent(ByVal RHS As Object)
    Set Parent = RHS
End Property

Private Property Get ICodeModule_Parent() As Object
    Set ICodeModule_Parent = MyModule.Parent
End Property


Private Function ICodeModule_ProcBodyLine(ByVal sProcName As String, ByVal lProcType As Long) As Long
On Error Resume Next
    If sProcName = "Form_Load" Then sProcName = "UserForm_Initialize"
    If sProcName = "Form_Unload" Then sProcName = "UserForm_Terminate"
    
    If lProcType = 0 Then
       lProcType = 0
    ElseIf lProcType = 1 Then
       lProcType = 1
    ElseIf lProcType = 2 Then
       lProcType = 3
    ElseIf lProcType = 3 Then
       lProcType = 2
    End If
    
    ICodeModule_ProcBodyLine = MyModule.ProcBodyLine(sProcName, lProcType)
End Function


Private Function ICodeModule_ProcCountLines(ByVal sProcName As String, ByVal lProcType As Long) As Long
    If sProcName = "Form_Load" Then sProcName = "UserForm_Initialize"
    If sProcName = "Form_Unload" Then sProcName = "UserForm_Terminate"
    ICodeModule_ProcCountLines = MyModule.ProcCountLines(sProcName, lProcType)
End Function


Private Function ICodeModule_ProcOfLine(ByVal lLineNumberOfProc As Long, ByVal lProcType As Long) As String
    ICodeModule_ProcOfLine = MyModule.ProcOfLine(lLineNumberOfProc, lProcType)
End Function


Private Function ICodeModule_ProcStartLine(ByVal sProcName As String, ByVal lProcType As Long) As Long
On Error Resume Next
    If sProcName = "Form_Load" Then sProcName = "UserForm_Initialize"
    If sProcName = "Form_Unload" Then sProcName = "UserForm_Terminate"
    ICodeModule_ProcStartLine = MyModule.ProcStartLine(sProcName, lProcType)
End Function


