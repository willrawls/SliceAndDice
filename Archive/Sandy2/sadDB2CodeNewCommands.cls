VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NewCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements SandySupport.ISadAddin

Public Parent As SandySupport.ISandyWizard
Public MySadCommands As SandySupport.CSadCommands
Public Externals As SandySupport.CAssocArray

Private Sub Class_Initialize()
On Error Resume Next
    Set MySadCommands = New SandySupport.CSadCommands
    Set Externals = New SandySupport.CAssocArray
    With MySadCommands
         .Clear
         .ParameterDelimiter = ","
         .ParameterTypeDelimiter = ":"
         .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision
    End With

    With Externals
         .Clear
         .Item("&Database to Code Generator (Primary)") = "Db2Code - Primary"
    End With
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Set Externals = Nothing
    Set MySadCommands = Nothing
    Set frmDBClassGen.Parent = Nothing
    Unload frmDBClassGen
End Sub

Private Property Get ISadAddin_CommandSet() As Object
On Error Resume Next
    Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
    Dim X As Form
    Dim sMessage As String

On Error Resume Next
    If frmDBClassGen.Parent Is Nothing Then
       Set frmDBClassGen.Parent = Me
    End If
    
    Select Case UCase$(sValue)
           Case "DB2CODE - PRIMARY"
                Set frmDBClassGen.Parent = Parent.UI

                frmDBClassGen.RefreshCategories

                If frmDBClassGen.Visible Then
                   frmDBClassGen.Hide
                Else
                   Set X = Nothing
                   Set X = Parent.UI
                       If X Is Nothing Then
                          frmDBClassGen.Show
                       Else
                          frmDBClassGen.Show 0, X
                       End If
                       If Not frmDBClassGen.Visible Then
                          frmDBClassGen.Show
                       End If
                   Set X = Nothing
                End If

           Case "HIDE ALL WINDOWS", "HIDEALLWINDOWS"
               'Set frmDBClassGen.Parent = Nothing
                If frmDBClassGen.Visible Then
                   frmDBClassGen.Hide
                End If

           Case "UNLOAD"
                Set frmDBClassGen.Parent = Nothing
                Unload frmDBClassGen
    End Select
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
    '
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
    '
End Function

Private Property Get ISadAddin_Externals() As SandySupport.CAssocArray
On Error Resume Next
    Set ISadAddin_Externals = Externals
End Property

Private Function ISadAddin_Shutdown() As Boolean
On Error Resume Next
    MySadCommands.Clear
    Set MySadCommands.Parent = Nothing
    Set MySadCommands = Nothing
    Set Parent = Nothing

    ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oSliceAndDiceWizard As Object, Application As Object) As Boolean
On Error Resume Next
    Err.Clear

    Set Parent = oSliceAndDiceWizard
    Set MySadCommands.Parent = oSliceAndDiceWizard

    ISadAddin_Startup = (Err.Number = 0)
End Function

