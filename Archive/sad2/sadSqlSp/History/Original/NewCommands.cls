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
On Error Resume Next
    Set MySadCommands = New SliceAndDice.CSadCommands
    With MySadCommands
         .ParameterDelimiter = ","
         .ParameterTypeDelimiter = ":"
         .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision
         With .Add("spParamList", True)
              .Syntax = "[StoredProcedureName As StringOrSoftVar] , [DsnInformation As StringOrSoftVar]"
              .OneLineDescription = "Returns a SQL Server stored procedure's parameters"
              .LongDescription = "Parameters are space delimited"
              .Examples = "~~Set sParams=%%spParamList::_Test,dbTest%%"
              .Comments = "Use at your own risk. Contact lau800@hotmail.com for help."
              .SeeAlso = "spSelectList"
              .Aliases = ", spParmList, spParameterList, spParameters, "
         End With
         
         With .Add("spSelectList", True)
              .OneLineDescription = "Returns a SQL Server stored procedure's parameters"
              .LongDescription = "Parameters are space delimited"
              .Examples = "~~Set sParams=%%spSelectList::_Test,dbTest%%"
              .Comments = "Use at your own risk. Contact lau800@hotmail.com for help."
              .Syntax = "[StoredProcedureName As StringOrSoftVar] , [DsnInformation As StringOrSoftVar]"
              .SeeAlso = "spParamList"
              .Aliases = ", spFieldList, spFields, "
         End With
         
         
'         With .Add("SqlSp", False)
'              .Syntax = "[Result As SoftVar]=[Expression As SoftVarOrString]"
'              .OneLineDescription = "Gets MS SqlServer (v6.5 or 7.0) Stored Procedures Info"
'              .LongDescription = "'Copyright© 2000, Chung Lau." & vbNewLine & _
'                    "'_____________________________________________________" & vbNewLine & _
'                    "'WARNING: This module contains codes that access to your SQL" & vbNewLine & _
'                    "'Server Stored Procedures. Use at your own risk." & vbNewLine & _
'                    "'Chung Lau DISCLAIMS ALL OTHER WARRANTIES, EITHER EXPRESS OR IMPLIED, " & vbNewLine & _
'                    "'INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF MERCHANTABILITY AND " & vbNewLine & _
'                    "'FITNESS FOR A PARTICULAR PURPOSE, WITH RESPECT TO THE SOFTWARE AND " & vbNewLine & _
'                    "'THE ACCOMPANYING WRITTEN MATERIALS." & vbNewLine & _
'                    "'_____________________________________________________" & vbNewLine & _
'                    "'______________________________________________________" & vbNewLine & _
'                    "'Please contact lau800@hotmail.com for help." & vbNewLine
'         End With
    End With
End Sub

Private Sub Class_Terminate()
    Set MySadCommands = Nothing
End Sub

Private Property Get ISadAddin_CommandSet() As SliceAndDice.CSadCommands
On Error Resume Next
    Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
'    Select Case UCase$(sKey)
'           Case "TESTING EXTERNALS"
'                MsgBox sValue
'    End Select
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
'On Error GoTo EH_sadSqlSp_ISadAddin_ExecuteSoftCommand
'    Dim sVar1 As String
'    Dim sVar2 As String
'    Dim sVar3 As String
'    Dim sVar4 As String
'    Dim oSqlSp As clsSqlSp
'
'    With II
'         If Not MySadCommands(.SoftCommandName & "*C") Is Nothing Then
'
'            If MySadCommands(.SoftCommandName & "*C").IsInline Then Exit Function
'
'            sVar1 = Parent.sGetToken(II.AllParameters, 1, ","): If Len(.SoftVars(sVar1)) Then sVar1 = .SoftVars(sVar1)
'            sVar2 = Parent.sGetToken(II.AllParameters, 2, ","): If Len(.SoftVars(sVar2)) Then sVar2 = .SoftVars(sVar2)
'           'sVar3 = Parent.sGetToken(II.AllParameters, 3, ","): If Len(.SoftVars(sVar3)) Then sVar3 = .SoftVars(sVar3)
'           'sVar4 = Parent.sGetToken(II.AllParameters, 4, ","): If Len(.SoftVars(sVar4)) Then sVar4 = .SoftVars(sVar4)
'
''On Error Resume Next
'            Err.Clear
'            Select Case UCase$(.SoftCommandName)
''                   Case "SPPARMLIST"
''                         Set oSqlSp = New clsSqlSp
''                         If Not oSqlSp Is Nothing Then
''                            .SoftVars(II.Result) = oSqlSp.spParamList(sVar1, sVar2) & vbNullString
''                         End If
''                         Set oSqlSp = Nothing
''
''                         If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
''                          ' Clean up for immediate exit here
''                            ISadAddin_ExecuteSoftCommand = True
''                            Exit Function
''                         End If
''                         ISadAddin_ExecuteSoftCommand = True
''
''                   Case "SPSELECTLIST"
''                         Set oSqlSp = New clsSqlSp
''                         If Not oSqlSp Is Nothing Then
''                            .SoftVars(II.Result) = oSqlSp.spSelectList(sVar1, sVar2) & vbNullString
''                         End If
''                         Set oSqlSp = Nothing
''
''                         If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
''                          ' Clean up for immediate exit here
''                            ISadAddin_ExecuteSoftCommand = True
''                            Exit Function
''                         End If
''                         ISadAddin_ExecuteSoftCommand = True
'
'                 ' More Commands by Case here
'            End Select
'         End If
'    End With
'
'EH_sadSqlSp_ISadAddin_ExecuteSoftCommand_Continue:
'    Exit Function
'
'EH_sadSqlSp_ISadAddin_ExecuteSoftCommand:
'
'    ErrorsOcurred II, sVar1, sVar2, sVar3, sVar4, vbNullString
'    Resume EH_sadSqlSp_ISadAddin_ExecuteSoftCommand_Continue
'
'    Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
On Error GoTo EH_sadSqlSp_ISadAddin_ExecuteSoftCommandInline

    Dim ParameterCount As Long

    Dim sVar()         As String
    ReDim sVar(1 To 2) As String

    Dim oSqlSp         As clsSqlSp

    If Not MySadCommands(sInlineSoftCommandName & "*I") Is Nothing Then
       If Not MySadCommands(sInlineSoftCommandName & "*I").IsInline Then Exit Function

       ParameterCount = Parent.lTokenCount(sParameters, ",")

       If ParameterCount > 0 Then
'          If ParameterCount < 5 Then
'             ReDim sVar(1 To 5) As String
'          Else
'             ReDim sVar(1 To ParameterCount) As String
'          End If
'          For CurrParam = 1 To ParameterCount
'              sVar(CurrParam) = Parent.sGetToken(sParameters, CurrParam, ",")
'              If Len(SoftVars(sVar(CurrParam))) Then sVar(CurrParam) = SoftVars(sVar(CurrParam))
'          Next CurrParam

          'ReDim sVar(1 To 2) As String
           sVar(1) = Parent.sGetToken(sParameters, 1, ",")
               If Len(SoftVars(sVar(1))) Then sVar(1) = SoftVars(sVar(1))
           sVar(2) = Parent.sAfter(sParameters, 1, ",")
               If Len(SoftVars(sVar(2))) Then sVar(1) = SoftVars(sVar(2))
       End If

       Select Case sInlineSoftCommandName
              Case "SPPARMLIST", "SPPARAMLIST", "SPPARAMETERLIST", "SPPARAMETERS"
                    Set oSqlSp = New clsSqlSp
                        If Not oSqlSp Is Nothing Then
                           sResults = oSqlSp.spParamList(sVar(2), sVar(1)) & vbNullString
                        End If
                    Set oSqlSp = Nothing
                    ISadAddin_ExecuteSoftCommandInline = True   ' Always return true so Slice and Dice knows the SoftCommand was handled, even if the command failed
            
              Case "SPSELECTLIST", "SPFIELDLIST", "SPFIELDS"
                    Set oSqlSp = New clsSqlSp
                        If Not oSqlSp Is Nothing Then
                           sResults = oSqlSp.spSelectList(sVar(2), sVar(1)) & vbNullString
                        End If
                    Set oSqlSp = Nothing
                    ISadAddin_ExecuteSoftCommandInline = True   ' Always return true so Slice and Dice knows the SoftCommand was handled, even if the command failed
        End Select
    End If

EH_sadSqlSp_ISadAddin_ExecuteSoftCommandInline_Continue:
    Exit Function

EH_sadSqlSp_ISadAddin_ExecuteSoftCommandInline:
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

    Resume EH_sadSqlSp_ISadAddin_ExecuteSoftCommandInline_Continue
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
    MySadCommands.Clear
    Set MySadCommands.Parent = Nothing
    Set MySadCommands = Nothing
    Set Parent = Nothing

    ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard, vbInst As VBIDE.VBE) As Boolean
On Error Resume Next
    Err.Clear

    Set Parent = oParent
    Set MySadCommands.Parent = oParent

    Err.Clear
       'MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath & "sadosqlsp.txt")
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
       sMessage = sMessage & vbTab & "Sandal:       sadSqlSp.NewCommands" & vbNewLine
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
       sMessage = sMessage & vbTab & "NOTE: This information will be stored in: """ & App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "\sadSqlSpError.Log"""

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
       Open App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "\sadSqlSpError.Log" For Append As #fh
            Print #fh, sMessage
       Close #fh
    End If
End Function


