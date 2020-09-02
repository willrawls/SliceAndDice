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

Public Parent           As SliceAndDice.Wizard
Public MySadCommands    As SliceAndDice.CSadCommands
Public Externals        As SliceAndDice.CAssocArray

Implements SliceAndDice.ISadAddin


Private Sub Class_Initialize()
1    On Error Resume Next
2        Set MySadCommands = New SliceAndDice.CSadCommands
3        Set Externals = New SliceAndDice.CAssocArray
4        With MySadCommands
5             .Clear
6             .ParameterDelimiter = ","
7             .ParameterTypeDelimiter = ":"
8             .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision
9             With .Add("GetURL", False)
10                 .Aliases = "FileURL, GetWebPage, GetPage, FetchURL"
11                 .Comments = "Timeout defaults to 120 seconds"
12                 .OneLineDescription = "Fetches the contents of the specified URL to the indicated SoftVarOut"
13                 .Syntax = "SoftVarOut:SoftVar = URL:String"
14            End With
15            With .Add("GetFileURL", False)
16                 .Aliases = "URL, GetFileFromPage, GetRemoteFile, FetchFile"
17                 .Comments = "Timeout defaults to 120 seconds"
18                 .OneLineDescription = "Fetches a binary file from the specified URL to the indicated filename"
19                 .Syntax = "FilenameOut:String = URL:String"
20            End With
21       End With

22       With Externals
23            .Clear
        '.Item("&Register Slice and Dice") = "Register"
24            .Item("&Check for version updates") = "Check for updated files"
25            .Item("&Submit current Template to CTL") = "Submit Template"
26       End With
End Sub

Private Sub Class_Terminate()
27   On Error Resume Next
28       Set Externals = Nothing
29       Set MySadCommands = Nothing
30       Set frmRegister.Parent = Nothing
31       Unload frmRegister
End Sub

Private Property Get ISadAddin_CommandSet() As SliceAndDice.CSadCommands
32   On Error Resume Next
33       Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
34   On Error Resume Next
35       Dim X           As Form
36       Dim sMessage    As String

37       If frmRegister.Parent Is Nothing Then
38          Set frmRegister.Parent = Me
39       End If

    Select Case UCase$(sValue)
'           Case "REGISTER"
'                If MySadCommands.Attributes("Registered") = "True" Then
'                   MsgBox "You have already registered Slice and Dice successfully." & vbCr & vbTab & "Invoice # " & sadGetLicenseKey("Invoice Number", ""), vbInformation
'                Else
'                   If frmRegister.Visible Then
'                      frmRegister.Hide
'                   Else
'                      Set X = Nothing
'                      Set X = Parent.SandyWindow
'                          If X Is Nothing Then
'                             frmRegister.Show
'                          Else
'                             frmRegister.Show 0, X
'                          End If
'                          If Not frmRegister.Visible Then
'                             frmRegister.Show
'                          End If
'                      Set X = Nothing
'                   End If
'                End If
'
           Case "UPDATE FILES"
40                    If frmRegister.GetCentralUpdateInfo Then
41                       MsgBox "Success !" & vbCr & vbCr & "Please note, no product update has been done." & vbNewLine & "This feature is still beta." & vbCr & vbTab & "Please check http://www.sliceanddice.com for an update."
42                       sMessage = sadGetLicenseKey("Message")
43                       If Len(sMessage) Then
44                          sMessage = sReplace(sReplace(sMessage, "$%$EOL$%$", vbCr), "$%$TAB$%$", vbTab)
45                          MsgBox "A message was received from the Central Server as follows:" & vbCr & vbCr & sMessage
46                       End If
47                    End If

48              Case "SUBMIT TEMPLATE"
49                    frmRegister.SubmitTemplate

50              Case "HIDE ALL WINDOWS", "HIDEALLWINDOWS"
51                   If frmRegister.Visible Then frmRegister.Hide

52              Case "UNLOAD"
53                   Unload frmRegister
54       End Select
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
55   On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommand

56       With II
57            If Not MySadCommands(.SoftCommandName & "*C") Is Nothing Then
58               If MySadCommands(.SoftCommandName & "*C").IsInline Then Exit Function

59   On Error Resume Next
            Select Case UCase$(.SoftCommandName)
                   Case "URL", "GETWEBPAGE", "GETPAGE", "FETCHURL", "GETURL"
60                            .SoftVars(.Result) = frmRegister.GetURL(.Expression)
61                            ISadAddin_ExecuteSoftCommand = True
62                      Case "GETFILEURL", "FILEURL", "GETFILEFROMPAGE", "GETREMOTEFILE", "FETCHFILE"
63                            Call frmRegister.GetFile(.Expression, .Result)
64                            ISadAddin_ExecuteSoftCommand = True
65               End Select
66            End If
67       End With

68 EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:
69       Exit Function

70 EH_SADAddin_ISadAddin_ExecuteSoftCommand:
71       ErrorsOcurred II
72       Resume EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue

73       Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
    '
End Function

Private Property Get ISadAddin_Externals() As SliceAndDice.CAssocArray
74   On Error Resume Next
75       Set ISadAddin_Externals = Externals
End Property

Private Function ISadAddin_Shutdown() As Boolean
76   On Error Resume Next
77       MySadCommands.Clear
78       Set MySadCommands.Parent = Nothing
79       Set MySadCommands = Nothing
80       Set Parent = Nothing

81       ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard, vbInst As VBIDE.VBE) As Boolean
82   On Error Resume Next
83       Err.Clear

84       Set Parent = oParent
85       Set MySadCommands.Parent = oParent

86       ISadAddin_Startup = (Err.Number = 0)
End Function

Private Function ErrorsOcurred(ByRef II As CInsertionInfo) As Boolean
87   On Error Resume Next
88       Dim fh          As Long
89       Dim sMessage    As String

90       If Err.Number <> 0 Then
91          sMessage = "Error executing SoftCode:" & vbNewLine
92          sMessage = sMessage & vbTab & "Occured:      " & Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM")
93          sMessage = sMessage & vbTab & "Sandal:       sadRegister.NewCommands" & vbNewLine
94          If Erl <> 0 Then sMessage = sMessage & vbTab & "Sandal Line:  " & Erl & vbNewLine
95          sMessage = sMessage & vbTab & "Error Number: " & Err.Number & vbNewLine
96          sMessage = sMessage & vbTab & "Description:  " & Err.Description & vbNewLine

97          sMessage = sMessage & vbNewLine & vbTab & "SoftCode Parameters (Resolved):" & vbNewLine
98          sMessage = sMessage & vbTab & vbTab & "-No variables-" & vbNewLine
 
99          If Not II Is Nothing Then
100            sMessage = sMessage & vbNewLine & vbTab & "(Unresolved) CInsertionInfo contents:" & vbNewLine
101            sMessage = sMessage & vbTab & vbTab & "TemplateName = " & II.TemplateName & vbNewLine
102            sMessage = sMessage & vbTab & vbTab & "CurrentLineToProcess = " & II.CurrentLineToProcess & vbNewLine
103            sMessage = sMessage & vbTab & vbTab & "SoftCommandName = " & II.SoftCommandName & vbNewLine
104            sMessage = sMessage & vbTab & vbTab & "AllParameters = " & II.AllParameters & vbNewLine
105            sMessage = sMessage & vbTab & vbTab & "Result     (LHS) = " & II.Result & vbNewLine
106            sMessage = sMessage & vbTab & vbTab & "Expression (RHS) = " & II.Expression & vbNewLine
107            sMessage = sMessage & vbTab & vbTab & "ExternalFilename = " & II.ExternalFilename & vbNewLine
108         End If

109         sMessage = sMessage & vbNewLine & vbNewLine & "Would you like to cancel processing ?" & vbNewLine
110         sMessage = sMessage & vbNewLine & vbTab & "IF YOU SELECT:" & vbNewLine
111         sMessage = sMessage & vbTab & vbTab & "YES: This template should be cancelled."
112         sMessage = sMessage & vbTab & vbTab & vbTab & "NOTE: Due to the nature of this addin, processing is not guarenteed to stop."
113         sMessage = sMessage & vbTab & vbTab & "NO : Processing will continue with the next line of SoftCode."
114         sMessage = sMessage & vbTab & "NOTE: This information will be stored in: """ & App.Path & "\sadRegisterErrors.Log"""

115         If MsgBox(sMessage, vbYesNo, "CANCEL PROCESSING ?") = vbYes Then
116            If Not II Is Nothing Then
117               II.LinesLeftToProcess = vbNewLine & vbNewLine & "~~' Cancel inserted at user's request" & vbNewLine & "~~Cancel 1" & vbNewLine
118            End If
119            sMessage = sMessage & vbNewLine & "  *** User decided to CANCEL processing."
120            ErrorsOcurred = True ' Tell parent function processing has been cancelled.
121         Else
122            sMessage = sMessage & vbNewLine & "  *** User choose to CONTINUE after error."
123            Err.Clear
124         End If

125         fh = FreeFile
126         Open App.Path & "\sadRegisterError.Log" For Append As #fh
127              Print #fh, sMessage
128         Close #fh
129      End If
End Function


