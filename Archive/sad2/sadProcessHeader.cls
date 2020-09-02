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

Implements SliceAndDice.ISadAddin

Public Parent As SliceAndDice.Wizard
Public MySadCommands As SliceAndDice.CSadCommands
Public Externals As SliceAndDice.CAssocArray
Private Sub Class_Initialize()
1    On Error Resume Next
2        Set MySadCommands = New SliceAndDice.CSadCommands
3        With MySadCommands
4             .Clear
5             .ParameterDelimiter = ","
6             .ParameterTypeDelimiter = ":"
7             .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision & " Beta"
8        End With
    
9        Set Externals = New SliceAndDice.CAssocArray
10       With Externals
11            .Clear
12            .Item("&Process C++ Header") = "PROCESS CPP HEADER"
13       End With
End Sub

Private Sub Class_Terminate()
14       Set MySadCommands = Nothing
15       Set Externals = Nothing
End Sub

Private Property Get ISadAddin_CommandSet() As SliceAndDice.CSadCommands
16   On Error Resume Next
17       Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
19 On Error Resume Next
18   Dim x As Form
     Select Case UCase$(sValue)
            Case "PROCESS CPP HEADER"
20                   If frmOptions.Parent Is Nothing Then
21                      Set frmOptions.Parent = Me
22                   End If
                
23                   If frmOptions.Visible Then
24                      frmOptions.Hide
25                   Else
26                      Set x = Parent.SandyWindow
27                      If x Is Nothing Then
28                         frmOptions.Show
29                      Else
30                         frmOptions.Show , x
31                      End If
32                      Set x = Nothing
33                   End If

34              Case "HIDE ALL WINDOWS"
35                   If frmOptions.Visible Then
36                      frmOptions.Hide
37                   End If

38              Case "UNLOAD"
39                   frmOptions.Parent = Nothing
40                   Unload frmOptions
41       End Select
End Function


Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
'
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
'
End Function

Private Property Get ISadAddin_Externals() As SliceAndDice.CAssocArray
100      Set ISadAddin_Externals = Externals
End Property

Private Function ISadAddin_Shutdown() As Boolean
101 On Error Resume Next
102      MySadCommands.Clear
103      Set MySadCommands.Parent = Nothing
104      Set MySadCommands = Nothing
    
105      Externals.Clear
106      Set Externals = Nothing

107      frmOptions.Parent = Nothing
108      Unload frmOptions

109      Set Parent = Nothing

110      ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard, vbInst As VBIDE.VBE) As Boolean
111  On Error Resume Next
112      Err.Clear

113      Set Parent = oParent
114      Set MySadCommands.Parent = oParent

117      ISadAddin_Startup = (Err.Number = 0)
118      Err.Clear
End Function
