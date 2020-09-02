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
18       Dim x As Form
19   On Error Resume Next
    Select Case UCase(sValue)
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
42   On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommand
43       Dim bEOLAtEndOfLine As Boolean

44       Dim lWrapLength As Long
45       Dim lThisWrap As Long
46       Dim lLineOffset As Long

47       Dim sWordWrapped As String
48       Dim sToWrap As String
49       Dim Token1 As String
50       Dim Token2 As String
51       Dim SoftVar1 As String
52       Dim SoftVar2 As String
53       Dim sOperator As String

54       With II
55            If Not MySadCommands(.SoftCommandName & "*C") Is Nothing Then
56               If MySadCommands(.SoftCommandName & "*C").IsInline Then Exit Function

            Select Case UCase(.SoftCommandName)
                   Case "TESTING"
57                            MsgBox "Soft command 'Testing' executed."
58                            ISadAddin_ExecuteSoftCommand = True

59                      Case "ANOTHERONE"
60                            MsgBox "Soft command 'AnotherOne' executed."
61                            ISadAddin_ExecuteSoftCommand = True

                  'Case "X"
                  'Case "Y"
                  'Case "Z"
62               End Select
63            End If
64       End With

65 EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:
66       Exit Function

67 EH_SADAddin_ISadAddin_ExecuteSoftCommand:
68       Parent.LogError "sadProcessHeader", "ISadAddin_ExecuteSoftCommand", Err.Number, Err.Description
69       Resume EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue

70       Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
71   On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline
72       Dim Template As CTemplate

73       Dim Area As Long
74       Dim CurrSet As Long
75       Dim bInlineCommandExecuted As Boolean
76       Dim lParamCount As Long

77       Dim sDefault As String
78       Dim sT As String
79       Dim sVar1 As String
80       Dim sVar2 As String
81       Dim sVar3 As String

82       If Not MySadCommands(sInlineSoftCommandName & "*I") Is Nothing Then
83          If Not MySadCommands(sInlineSoftCommandName & "*I").IsInline Then Exit Function
       Select Case sInlineSoftCommandName
              Case "INLINE1"
84                      sResults = InputBox("What do you want in here ?")
85                      ISadAddin_ExecuteSoftCommandInline = True

             'Case "INLINEX"
86                      sResults = ""
87                      ISadAddin_ExecuteSoftCommandInline = True

             'Case "INLINEY"
88                      sResults = ""
89                      ISadAddin_ExecuteSoftCommandInline = True

             'Case "INLINEZ"
90                      sResults = ""
91                      ISadAddin_ExecuteSoftCommandInline = True

92           End Select
93       End If

94 EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue:
95       Exit Function

96 EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline:
97       Parent.LogError "sadProcessHeader", "ISadAddin_ExecuteSoftCommandInline", Err.Number, Err.Description
98       Resume EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue

99       Resume
End Function

Private Property Get ISadAddin_Externals() As SliceAndDice.CAssocArray
100      Set ISadAddin_Externals = Externals
End Property

Private Function ISadAddin_Shutdown() As Boolean
101  On Error Resume Next
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

'    If UCase(oParent.Version) <> UCase(MySadCommands.Attributes("Version")) Then
'       If MsgBox("Version mismatch:" & Chr(13) & Chr(9) & "sadProcessHeader.NewCommands = " & MySadCommands.Attributes("Version") & Chr(13) & Chr(9) & "Slice and Dice = " & oParent.Version & Chr(13) & Chr(13) & "Continue loading DLL ?", vbYesNo, "*** WARNING - DLL Version mismatch **") = vbNo Then
'          Exit Function
'       End If
'    End If
'
113      Set Parent = oParent
114      Set MySadCommands.Parent = oParent

115      Err.Clear
116          MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath & "sadProcessHeader.txt")
117          ISadAddin_Startup = (Err.Number = 0)
118      Err.Clear
End Function
