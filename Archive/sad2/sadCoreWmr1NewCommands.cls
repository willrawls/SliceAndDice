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
Private CommandsSupported   As SliceAndDice.CSadCommands
Private Sub Class_Initialize()
On Error Resume Next
    Set CommandsSupported = New SliceAndDice.CSadCommands
    With CommandsSupported
         .ParameterDelimiter = ","
         .ParameterTypeDelimiter = ":"
         .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision
    End With
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Set CommandsSupported = Nothing
End Sub

Private Property Get ISadAddin_CommandSet() As SliceAndDice.CSadCommands
On Error Resume Next
    Set ISadAddin_CommandSet = CommandsSupported
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
'    Select Case UCase$(sKey)
'           Case "TESTING EXTERNALS"
'                MsgBox sValue
'    End Select
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
On Error GoTo SandalError_ExecuteSoftCommand
    Dim sVar1 As String
    Dim sVar2 As String
    Dim sVar3 As String
    Dim sVar4 As String

    With II
         If Not CommandsSupported(.SoftCommandName & "*C") Is Nothing Then

            If CommandsSupported(.SoftCommandName & "*C").IsInline Then Exit Function

            sVar1 = Parent.sGetToken(II.AllParameters, 1, ","): If Len(.SoftVars(sVar1)) Then sVar1 = .SoftVars(sVar1)
            sVar2 = Parent.sGetToken(II.AllParameters, 2, ","): If Len(.SoftVars(sVar2)) Then sVar2 = .SoftVars(sVar2)
            sVar3 = Parent.sGetToken(II.AllParameters, 3, ","): If Len(.SoftVars(sVar3)) Then sVar3 = .SoftVars(sVar3)
            sVar4 = Parent.sGetToken(II.AllParameters, 4, ","): If Len(.SoftVars(sVar4)) Then sVar4 = .SoftVars(sVar4)

On Error Resume Next
            Err.Clear
            Select Case UCase$(.SoftCommandName)
                   'Case "SOMETHING"
                   '      ISadAddin_ExecuteSoftCommand = True
            End Select
         End If
    End With

SandalError_ExecuteSoftCommand_Continue:
    Exit Function

SandalError_ExecuteSoftCommand:

    ErrorsOcurred II, sVar1, sVar2, sVar3, sVar4, vbNullString
    Resume SandalError_ExecuteSoftCommand_Continue

    Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
On Error GoTo SandalError_ExecuteSoftCommandInline
    Dim Template                As CTemplate
    Dim Area                    As Long
    Dim CurrSet                 As Long
    Dim bInlineCommandExecuted  As Boolean
    Dim lParameterCount         As Long
    Dim sDefault                As String
    Dim sT                      As String
    Dim sVar()                  As String
    
    Dim CurrParam               As Long
    Dim ParameterCount          As Long

    ReDim sVar(1 To 5) As String

    If Not CommandsSupported(sInlineSoftCommandName & "*I") Is Nothing Then
       If Not CommandsSupported(sInlineSoftCommandName & "*I").IsInline Then Exit Function

       ParameterCount = Parent.lTokenCount(sParameters, ",")

       If ParameterCount > 0 Then
          If ParameterCount < 5 Then
             ReDim sVar(1 To 5) As String
          Else
             ReDim sVar(1 To ParameterCount) As String
          End If
          For CurrParam = 1 To ParameterCount
              sVar(CurrParam) = Parent.sGetToken(sParameters, CurrParam, ",")
              If Len(SoftVars(sVar(CurrParam))) Then sVar(CurrParam) = SoftVars(sVar(CurrParam))
          Next CurrParam
       End If

       Select Case sInlineSoftCommandName
              Case "FINDPATTERNINFILE", "FINDINFILE", "FINDREGEXPINFILE"
                   sResults = FindPatternInFile(Parent.sGetToken(sParameters, 1, "="), Parent.sAfter(sParameters, 1, "="))
                   ISadAddin_ExecuteSoftCommandInline = True
        End Select
    End If

SandalError_ExecuteSoftCommandInline_Continue:
    Exit Function

SandalError_ExecuteSoftCommandInline:
    Dim X As CInsertionInfo

    Set X = New CInsertionInfo
    With X
         .AllParameters = sParameters
         .Result = sResults
         Set .SoftVars = SoftVars
         .SoftCommandName = sInlineSoftCommandName
         .CurrentLineToProcess = "-Inline Substition-"
    End With

    ErrorsOcurred X, sVar(1), sVar(2), sVar(3), sVar(4), sVar(5)

    Set X.SoftVars = Nothing
    Set X = Nothing

    Resume SandalError_ExecuteSoftCommandInline_Continue
    Resume
End Function

Private Property Get ISadAddin_Externals() As SliceAndDice.CAssocArray
'    Dim Externals As CAssocArray
'    Set Externals = New CAssocArray
'    Externals.All = "Testing Externals=Testing 123"
'    Set ISadAddin_Externals = Externals
'    Set Externals = Nothing
End Property

Private Function ISadAddin_Shutdown() As Boolean
On Error Resume Next
    CommandsSupported.Clear
    Set CommandsSupported.Parent = Nothing
    Set CommandsSupported = Nothing
    Set Parent = Nothing

    ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard, vbInst As VBIDE.VBE) As Boolean
On Error Resume Next
    Err.Clear

    Set Parent = oParent
    Set CommandsSupported.Parent = oParent

    Err.Clear
        With CommandsSupported.Add("FindInFile", True)
             .Aliases = "FindPatternInFile"
             .Examples = "~~FindInFile c:\sadResults.txt=occ*"
             .OneLineDescription = "[FN As PathAndFile] = [Expression As WSHRegExpString]"
             .Comments = "Expression is as the WSH "
        End With
        
        CommandsSupported.All = Parent.sFileContents(Parent.TemplateDatabasePath & "sadCoreWmr.txt")
        ISadAddin_Startup = (Err.Number = 0)
    Err.Clear
End Function

Private Function ErrorsOcurred(ByRef II As CInsertionInfo, ByVal sVar1 As String, ByVal sVar2 As String, ByVal sVar3 As String, ByVal sVar4 As String, ByVal sVar5 As String) As Boolean
On Error Resume Next
    Dim fh          As Long
    Dim sMessage    As String

    If Err.Number <> 0 Then
       sMessage = "Error executing SoftCode:" & vbNewLine
       sMessage = sMessage & vbTab & "Occured:      " & Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM")
       sMessage = sMessage & vbTab & "Sandal:       sadCoreWmr.NewCommands" & vbNewLine
       If Erl <> 0 Then sMessage = sMessage & vbTab & "Sandal Line:  " & Erl & vbNewLine
       sMessage = sMessage & vbTab & "Error Number: " & Err.Number & vbNewLine
       sMessage = sMessage & vbTab & "Description:  " & Err.Description & vbNewLine

       sMessage = sMessage & vbNewLine & vbTab & "SoftCode Parameters (Resolved):" & vbNewLine
       sMessage = sMessage & vbTab & vbTab & "sVar(1) = " & sVar1 & vbNewLine
       sMessage = sMessage & vbTab & vbTab & "sVar(2) = " & sVar2 & vbNewLine
       sMessage = sMessage & vbTab & vbTab & "sVar(3) = " & sVar3 & vbNewLine
       sMessage = sMessage & vbTab & vbTab & "sVar(4) = " & sVar4 & vbNewLine
       sMessage = sMessage & vbTab & vbTab & "sVar(5) = " & sVar5 & vbNewLine

       If Not II Is Nothing Then
          sMessage = sMessage & vbNewLine & vbTab & "(Unresolved) CInsertionInfo contents:" & vbNewLine
          sMessage = sMessage & vbTab & vbTab & "TemplateName = " & II.TemplateName & vbNewLine
          sMessage = sMessage & vbTab & vbTab & "CurrentLineToProcess = " & II.CurrentLineToProcess & vbNewLine
          sMessage = sMessage & vbTab & vbTab & "SoftCommandName = " & II.SoftCommandName & vbNewLine
          sMessage = sMessage & vbTab & vbTab & "AllParameters = " & II.AllParameters & vbNewLine
          sMessage = sMessage & vbTab & vbTab & "Result     (LHS) = " & II.Result & vbNewLine
          sMessage = sMessage & vbTab & vbTab & "Expression (RHS) = " & II.Expression & vbNewLine
          sMessage = sMessage & vbTab & vbTab & "ExternalFilename = " & II.ExternalFilename & vbNewLine
       End If

       sMessage = sMessage & vbNewLine & vbNewLine & "Would you like to cancel processing ?" & vbNewLine
       sMessage = sMessage & vbNewLine & vbTab & "IF YOU SELECT:" & vbNewLine
       sMessage = sMessage & vbTab & vbTab & "YES: This template should be cancelled."
       sMessage = sMessage & vbTab & vbTab & vbTab & "NOTE: Due to the nature of this addin, processing is not guarenteed to stop."
       sMessage = sMessage & vbTab & vbTab & "NO : Processing will continue with the next line of SoftCode."
       sMessage = sMessage & vbTab & "NOTE: This information will be stored in: """ & App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "\sadCoreWmrError.Log"""

       If MsgBox(sMessage, vbYesNo, "CANCEL PROCESSING ?") = vbYes Then
          If Not II Is Nothing Then
             II.LinesLeftToProcess = vbNullString
          End If
          sMessage = sMessage & vbNewLine & "  *** User decided to CANCEL processing."
          ErrorsOcurred = True ' Tell parent function processing has been cancelled.
       Else
          sMessage = sMessage & vbNewLine & "  *** User choose to CONTINUE after error."
          Err.Clear
       End If

       fh = FreeFile
       Open App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "\sadCoreWmrError.Log" For Append As #fh
            Print #fh, sMessage
       Close #fh
    End If
End Function

Private Function FindPatternInFile(sFilename As String, sRegularExpression As String) As Boolean
    Dim fs As FileSystemObject
    Dim ts As TextStream
    Dim re As RegExp
    Dim text As String

    Set fs = New FileSystemObject
    Set ts = fs.OpenTextFile(sFilename, ForReading)
    text = ts.ReadAll
    Set re = New RegExp
    re.Pattern = sRegularExpression
    FindPatternInFile = re.Test(text)
End Function


