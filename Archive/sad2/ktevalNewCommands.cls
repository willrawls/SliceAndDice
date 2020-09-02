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

Private Parent          As SliceAndDice.Wizard
Private MySadCommands   As SliceAndDice.CSadCommands

Private Sub Class_Initialize()
1        On Error Resume Next
2        Set MySadCommands = New SliceAndDice.CSadCommands
3        With MySadCommands
4            .ParameterDelimiter = ","
5            .ParameterTypeDelimiter = ":"
6            .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision
7            With .Add("KTEval", False)
8                .Syntax = "[Result As SoftVar]=[Expression As SoftVarOrString]"
9                .OneLineDescription = "Significantly more advanced Eval command"
10               .LongDescription = "'Copyright© 1999, Tretyakov Konstantin" & vbNewLine & _
                    "'_____________________________________________________" & vbNewLine & _
                    "'This is the 'Evaluator' class: it inputs a string" & vbNewLine & _
                    "'like ""2+2"" or ""2+4*sin(3.4)^2-8*arccos(0.55)"", etc" & vbNewLine & _
                    "'_____________________________________________________" & vbNewLine & _
                    "'You may use the code for free, if you give me credit." & vbNewLine & _
                    "'If you modify it or make your own program with it," & vbNewLine & _
                    "'I would VERY APPRECIATE, if you mail me it (or better-" & vbNewLine & _
                    "'a link to it)" & vbNewLine & _
                    "'On the whole - just do not stamp your name on what you haven't" & vbNewLine & _
                    "'done quite alone." & vbNewLine & _
                    "'This code was written totally by me, and 'it took me about" & vbNewLine & _
                    "'2 days to code it (and about a year" & vbNewLine & _
                    "'-that is,from the very moment I got interested in programming-" & vbNewLine & _
                    "'I spent dreaming of having such a thing)" & vbNewLine & _
                    "" & vbNewLine & _
                    "'(BTW this code seems to be quite unique-" & vbNewLine & _
                    "'I searched all over the Internet for such, but NOONE" & vbNewLine & _
                    "'is giving the source for such things)" & vbNewLine & _
                    "'______________________________________________________" & vbNewLine & _
                    "'Yours Sincerely, Konstantin Tretyakov (kt_ee@yahoo.com)" & vbNewLine
11           End With
12       End With
End Sub

Private Sub Class_Terminate()
1        Set MySadCommands = Nothing
End Sub

Private Property Get ISadAddin_CommandSet() As SliceAndDice.CSadCommands
1        On Error Resume Next
2        Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
'    Select Case UCase$(sKey)
'           Case "TESTING EXTERNALS"
'                MsgBox sValue
'    End Select
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
1        On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommand
2        Dim sVar1 As String
3        Dim sVar2 As String
4        Dim sVar3 As String
5        Dim sVar4 As String
6        Dim ktEval As Evaluator

7        With II
8            If Not MySadCommands(.SoftCommandName & "*C") Is Nothing Then

9                If MySadCommands(.SoftCommandName & "*C").IsInline Then Exit Function

10               sVar1 = Parent.sGetToken(II.AllParameters, 1, ","): If Len(.SoftVars(sVar1)) Then sVar1 = .SoftVars(sVar1)
11               sVar2 = Parent.sGetToken(II.AllParameters, 2, ","): If Len(.SoftVars(sVar2)) Then sVar2 = .SoftVars(sVar2)
            'sVar3 = Parent.sGetToken(II.AllParameters, 3, ","): If Len(.SoftVars(sVar3)) Then sVar3 = .SoftVars(sVar3)
            'sVar4 = Parent.sGetToken(II.AllParameters, 4, ","): If Len(.SoftVars(sVar4)) Then sVar4 = .SoftVars(sVar4)

            'On Error Resume Next
12               Err.Clear
            Select Case UCase$(.SoftCommandName)
                Case "KTEVAL"
                    ' Do something here
13                       Set ktEval = New Evaluator
14                       If Not ktEval Is Nothing Then
15                           .SoftVars(II.Result) = ktEval.Evaluate(II.Expression, True) & vbNullString
16                       End If
17                       Set ktEval = Nothing

18                       If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
                        ' Clean up for immediate exit here
19                           ISadAddin_ExecuteSoftCommand = True
20                           Exit Function
21                       End If
22                       ISadAddin_ExecuteSoftCommand = True

                    ' More Commands by Case here
23               End Select
24           End If
25       End With

26 EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:
27       Exit Function

28 EH_SADAddin_ISadAddin_ExecuteSoftCommand:

29       ErrorsOcurred II, sVar1, sVar2, sVar3, sVar4, vbNullString
30       Resume EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue

31       Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
'On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommandInline
'    Dim Template                As CTemplate
'    Dim Area                    As Long
'    Dim CurrSet                 As Long
'    Dim bInlineCommandExecuted  As Boolean
'    Dim lParameterCount         As Long
'    Dim sDefault                As String
'    Dim sT                      As String
'    Dim sVar()                  As String
'
'    Dim CurrParam               As Long
'    Dim ParameterCount          As Long
'
'    ReDim sVar(1 To 5) As String
'
'    If Not MySadCommands(sInlineSoftCommandName & "*I") Is Nothing Then
'       If Not MySadCommands(sInlineSoftCommandName & "*I").IsInline Then Exit Function
'
'       ParameterCount = Parent.lTokenCount(sParameters, ",")
'
'       If ParameterCount > 0 Then
'          If ParameterCount < 5 Then
'             ReDim sVar(1 To 5) As String
'          Else
'             ReDim sVar(1 To ParameterCount) As String
'          End If
'          For CurrParam = 1 To ParameterCount
'              sVar(CurrParam) = Parent.sGetToken(sParameters, CurrParam, ",")
'              If Len(SoftVars(sVar(CurrParam))) Then sVar(CurrParam) = SoftVars(sVar(CurrParam))
'          Next CurrParam
'       End If
'
'       Select Case sInlineSoftCommandName
'              Case "DOSOMETHING"
'                 ' Do something and return results through the string 'sResults'
'                   sResults = vbNullString
'
'                   ISadAddin_ExecuteSoftCommandInline = True
'
'            ' More inline commands go here
'        End Select
'    End If
'
'EH_SADAddin_ISadAddin_ExecuteSoftCommandInline_Continue:
'    Exit Function
'
'EH_SADAddin_ISadAddin_ExecuteSoftCommandInline:
'    Dim X As CInsertionInfo
'
'    Set X = New CInsertionInfo
'    With X
'         .AllParameters = sParameters
'         .Result = sResults
'         Set .SoftVars = SoftVars
'         .SoftCommandName = sInlineSoftCommandName
'         .CurrentLineToProcess = "-Inline Substition-"
'    End With
'
'    ErrorsOcurred X, sVar(1), sVar(2), sVar(3), sVar(4), sVar(5)
'
'    Set X.SoftVars = Nothing
'    Set X = Nothing
'
'    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommandInline_Continue
'    Resume
End Function

Private Property Get ISadAddin_Externals() As SliceAndDice.CAssocArray
'    Dim Externals As CAssocArray
'    Set Externals = New CAssocArray
'    Externals.All = "Testing Externals=Testing 123"
'    Set ISadAddin_Externals = Externals
'    Set Externals = Nothing
End Property

Private Function ISadAddin_Shutdown() As Boolean
1        On Error Resume Next
2        MySadCommands.Clear
3        Set MySadCommands.Parent = Nothing
4        Set MySadCommands = Nothing
5        Set Parent = Nothing

6        ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard, vbInst As VBIDE.VBE) As Boolean
1        On Error Resume Next
2        Err.Clear

3        Set Parent = oParent
4        Set MySadCommands.Parent = oParent

5        Err.Clear
    'MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath & "sadKTEval.txt")
6        ISadAddin_Startup = (Err.Number = 0)
7        Err.Clear
End Function

Private Function ErrorsOcurred(ByRef II As CInsertionInfo, ByVal sVar1 As String, ByVal sVar2 As String, ByVal sVar3 As String, ByVal sVar4 As String, ByVal sVar5 As String) As Boolean
1        On Error Resume Next
2        Dim fh          As Long
3        Dim sMessage    As String

4        If Err.Number <> 0 Then
5            sMessage = "Error executing SoftCode:" & vbNewLine
6            sMessage = sMessage & vbTab & "Occured:      " & Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM")
7            sMessage = sMessage & vbTab & "Sandal:       sadKTEval.NewCommands" & vbNewLine
8            If Erl <> 0 Then sMessage = sMessage & vbTab & "Sandal Line:  " & Erl & vbNewLine
9            sMessage = sMessage & vbTab & "Error Number: " & Err.Number & vbNewLine
10           sMessage = sMessage & vbTab & "Description:  " & Err.Description & vbNewLine

11           sMessage = sMessage & vbNewLine & vbTab & "SoftCode Parameters (Resolved):" & vbNewLine
12           sMessage = sMessage & vbTab & vbTab & "sVar(1) = " & sVar1 & vbNewLine
13           sMessage = sMessage & vbTab & vbTab & "sVar(2) = " & sVar2 & vbNewLine
14           sMessage = sMessage & vbTab & vbTab & "sVar(3) = " & sVar3 & vbNewLine
15           sMessage = sMessage & vbTab & vbTab & "sVar(4) = " & sVar4 & vbNewLine
16           sMessage = sMessage & vbTab & vbTab & "sVar(5) = " & sVar5 & vbNewLine

17           If Not II Is Nothing Then
18               sMessage = sMessage & vbNewLine & vbTab & "(Unresolved) CInsertionInfo contents:" & vbNewLine
19               sMessage = sMessage & vbTab & vbTab & "TemplateName = " & II.TemplateName & vbNewLine
20               sMessage = sMessage & vbTab & vbTab & "CurrentLineToProcess = " & II.CurrentLineToProcess & vbNewLine
21               sMessage = sMessage & vbTab & vbTab & "SoftCommandName = " & II.SoftCommandName & vbNewLine
22               sMessage = sMessage & vbTab & vbTab & "AllParameters = " & II.AllParameters & vbNewLine
23               sMessage = sMessage & vbTab & vbTab & "Result     (LHS) = " & II.Result & vbNewLine
24               sMessage = sMessage & vbTab & vbTab & "Expression (RHS) = " & II.Expression & vbNewLine
25               sMessage = sMessage & vbTab & vbTab & "ExternalFilename = " & II.ExternalFilename & vbNewLine
26           End If

27           sMessage = sMessage & vbNewLine & vbNewLine & "Would you like to cancel processing ?" & vbNewLine
28           sMessage = sMessage & vbNewLine & vbTab & "IF YOU SELECT:" & vbNewLine
29           sMessage = sMessage & vbTab & vbTab & "YES: This template should be cancelled."
30           sMessage = sMessage & vbTab & vbTab & vbTab & "NOTE: Due to the nature of this addin, processing is not guarenteed to stop."
31           sMessage = sMessage & vbTab & vbTab & "NO : Processing will continue with the next line of SoftCode."
32           sMessage = sMessage & vbTab & "NOTE: This information will be stored in: """ & App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "\sadKTEvalError.Log"""

33           If MsgBox(sMessage, vbYesNo, "CANCEL PROCESSING ?") = vbYes Then
34               If Not II Is Nothing Then
35                   II.LinesLeftToProcess = vbNullString
36               End If
37               sMessage = sMessage & vbNewLine & "  *** User decided to CANCEL processing."
38               ErrorsOcurred = True                      ' Tell parent function processing has been cancelled.
39           Else
40               sMessage = sMessage & vbNewLine & "  *** User choose to CONTINUE after error."
41               Err.Clear
42           End If

43           fh = FreeFile
44           Open App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "\sadKTEvalError.Log" For Append As #fh
45           Print #fh, sMessage
46           Close #fh
47       End If
End Function
