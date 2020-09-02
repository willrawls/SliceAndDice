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
Attribute VB_Description = "Sandle - Main Commands"
Option Explicit

Implements SandySupport.ISadAddin

Private Parent          As SandySupport.ISandyWizard
Private MySadCommands   As SandySupport.CSadCommands
Private asaSaved        As SandySupport.CAssocArray

Private Const vbext_mt_Method = 1
Private Const vbext_mt_Property = 2
Private Const vbext_mt_Variable = 3
Private Const vbext_mt_Event = 4
Private Const vbext_mt_Const = 5

Private Sub Class_Initialize()
On Error Resume Next
    Set MySadCommands = New SandySupport.CSadCommands
    With MySadCommands
         .ParameterDelimiter = ","
         .ParameterTypeDelimiter = ":"
         .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision
    End With
End Sub

Private Sub Class_Terminate()
    Set MySadCommands = Nothing
End Sub

Private Property Get ISadAddin_CommandSet() As Object
On Error Resume Next
    Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
    '
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As SandySupport.CInsertionInfo) As Boolean
On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommand
    Dim bEOLAtEndOfLine As Boolean

    Dim lWrapLength As Long
    Dim lThisWrap As Long
    Dim lLineOffset As Long

    Dim Decision As VbMsgBoxResult
    Dim sWordWrapped As String
    Dim sToWrap As String
    Dim Token1 As String
    Dim Token2 As String
    Dim SoftVar1 As String
    Dim SoftVar2 As String
    Dim sOperator As String

    With II
         If Not MySadCommands(.SoftCommandName & "*C") Is Nothing Then
            If MySadCommands(.SoftCommandName & "*C").IsInline Then Exit Function

            Select Case UCase(.SoftCommandName)
                   Case "SAVEVARS"
                        Set asaSaved = New SandySupport.CAssocArray
                        asaSaved.All = .SoftVars.All

                   Case "RESTOREVARS"
                        If Not asaSaved Is Nothing Then
                        '   MsgBox "You must execute the ~~SaveVars command first (See F1 help) before using the ~~RestoreVars command."
                        'Else
                           asaSaved.KeyValueDelimiter = " saved="
                           .SoftVars.All = Replace(asaSaved.All & vbNewLine & .SoftVars.All, "saved saved", "saved")
                           Do While InStr(.SoftVars.All, "saved saved"): .SoftVars.All = Replace(.SoftVars.All, "saved saved", "saved"): Loop
                           asaSaved.KeyValueDelimiter = " ="
                        End If
                   
                   Case "APPEND"
                        .SoftVars(.Result) = .SoftVars(.Result) & Replace(Replace(Replace(.Expression, "$EOL$", vbNewLine), "$13$", vbCr), "$9$", vbTab)
                         ISadAddin_ExecuteSoftCommand = True
                   
                   Case "CHOOSE"
                         If InStr(.Expression, IIf(Len(.SoftVars("Choose Delimiter")) = 0, ";", .SoftVars("Choose Delimiter"))) Then
                            .SoftVars(.Result) = Parent.sChoose(.Expression, IIf(Len(.SoftVars("Choose Delimiter")) = 0, ";", .SoftVars("Choose Delimiter")), .SoftVars("Choose Default"))
                         ElseIf Len(.SoftVars(.Expression)) Then
                            .SoftVars(.Result) = Parent.sChoose(.SoftVars(.Expression))
                         End If
                         ISadAddin_ExecuteSoftCommand = True

                   Case "MANGLE"
                         .SoftVars(.Result) = Mangle(.Expression)
                         ISadAddin_ExecuteSoftCommand = True

                   Case "LTRIM", "LEFTTRIM"
                         .SoftVars(.AllParameters).Value = LTrim$(.SoftVars(.AllParameters))
                         ISadAddin_ExecuteSoftCommand = True

                   Case "RTRIM", "RIGHTTRIM"
                         .SoftVars(.AllParameters).Value = RTrim$(.SoftVars(.AllParameters))
                         ISadAddin_ExecuteSoftCommand = True

                   Case "TRIM"
                         .SoftVars(.AllParameters).Value = Trim$(.SoftVars(.AllParameters))
                         ISadAddin_ExecuteSoftCommand = True

                   Case "SET", "SETVALUE"
                         .SoftVars(.Result) = .Expression
                         ISadAddin_ExecuteSoftCommand = True

                   Case "STRCOPY"
                         If Len(.SoftVars(.Expression)) > 0 Then
                            .SoftVars(.Result).Value = .SoftVars(.Expression)
                         Else
                            .SoftVars(.Result).Value = .Expression
                         End If
                         ISadAddin_ExecuteSoftCommand = True

                   Case "LEFT"
                         If Val(sGetToken(.Result)) > 0 Then
                            .SoftVars(sAfter(.Result)) = Left(.Expression, Val(sGetToken(.Result)))
                         Else
                            .SoftVars(.Result) = .Expression
                         End If
                         ISadAddin_ExecuteSoftCommand = True

                   Case "MID"
                         If Val(sGetToken(.Result)) > 0 Then
                            If Val(sGetToken(.Result, 2)) > 0 Then
                               .SoftVars(sAfter(.Result)) = Mid$(.Expression, Val(sGetToken(.Result)), Val(sGetToken(.Result, 2)))
                            Else
                               .SoftVars(.Result) = Mid$(.Expression, Val(sGetToken(.Result)))
                            End If
                         Else
                            .SoftVars(.Result) = .Expression
                         End If
                         ISadAddin_ExecuteSoftCommand = True

                   Case "RIGHT"
                         If Val(sGetToken(.Result)) > 0 Then
                            .SoftVars(sAfter(.Result)) = Right(.Expression, Val(sGetToken(.Result)))
                         Else
                            .SoftVars(.Result) = .Expression
                         End If
                         ISadAddin_ExecuteSoftCommand = True

                   Case "WORDWRAP"                    ' Wrap the indicated text at the indicated point, preceding each wrap with the indicated prefix
                         lWrapLength = Val(sGetToken(II.AllParameters))
                         sToWrap = Trim$(sAfter(.sParam, 2, """"))
                         .sParam = sGetToken(.sParam, 2, """")
                         .LinesLeftToProcess = sAfter(.LinesLeftToProcess, 1, .SoftVars("EOL"))
                         Do Until Len(sToWrap) = 0
                            If Len(sToWrap) > lWrapLength Then
                               lThisWrap = lWrapLength
                               If lThisWrap > 0 Then
                                  Do While StrComp(Mid$(sToWrap, lThisWrap, 1), " ") <> 0
                                     lThisWrap = lThisWrap - 1
                                  Loop
                               End If
                               sWordWrapped = sWordWrapped & .sParam & Left(sToWrap, lThisWrap - 1) & .SoftVars("EOL")
                               sToWrap = Mid$(sToWrap, lThisWrap + 1)
                            Else
                               sWordWrapped = sWordWrapped & .sParam & sToWrap & .SoftVars("EOL")
                               sToWrap = vbNullString
                            End If
                         Loop
                         .LinesLeftToProcess = .SoftVars("EOL") & sWordWrapped & .LinesLeftToProcess
                         ISadAddin_ExecuteSoftCommand = True

                   Case "INCREMENT", "INC"
                         With .SoftVars(sGetToken(.AllParameters))
                              If Val(sAfter(II.AllParameters)) = 0 Then
                                 .Value = Val(.Value) + 1
                              Else
                                 .Value = Val(.Value) + Val(sAfter(II.AllParameters))
                              End If
                         End With
                         ISadAddin_ExecuteSoftCommand = True

                    Case "DECREMENT", "DEC"
                         With .SoftVars(sGetToken(.AllParameters))
                              If Val(sAfter(II.AllParameters)) = 0 Then
                                 .Value = Val(.Value) - 1
                              Else
                                 .Value = Val(.Value) - Val(sAfter(II.AllParameters))
                              End If
                         End With
                         ISadAddin_ExecuteSoftCommand = True
                         
                    Case "EVALUATE", "EVAL"
                         .SoftVars(.Result) = Evaluate(.Expression, .SoftVars)
                         ISadAddin_ExecuteSoftCommand = True

                    Case "STRING", "STR", "REPSTR", "REPLICATESTRING"
                         SoftVar1 = sGetToken(.AllParameters)
                         lLineOffset = sGetToken(.AllParameters, 2)
                         If lLineOffset < 1 Then lLineOffset = 1
                         If InStr(UCase(.AllParameters), "CHR(") > 0 Then
                            SoftVar2 = Chr(Val(sGetToken(sGetToken(UCase(.AllParameters), 2, "CHR("), 1, ")")))
                            .SoftVars(SoftVar1) = String(lLineOffset, SoftVar2)
                         Else
                            SoftVar2 = Left(sGetToken(.AllParameters, 2, """") & " ", 1)
                            .SoftVars(SoftVar1) = String(lLineOffset, SoftVar2)
                         End If
                         ISadAddin_ExecuteSoftCommand = True

                    Case "AUTOSPACE"
                         .SoftVars(.Result) = sInsertSpaces(.Expression)
                         ISadAddin_ExecuteSoftCommand = True

                    Case "REVERSETOKENS", "FLIPFLOP"
                         SoftVar1 = sGetToken(.AllParameters)
                         SoftVar2 = sGetToken(.AllParameters, 2)
                         sOperator = sGetToken(.AllParameters, 2, """")
                         .Expression = .SoftVars(SoftVar1)
                         With II.SoftVars(SoftVar2)
                              .Value = vbNullString
                              Do While Len(II.Expression)
                                 II.Result = sGetToken(II.Expression, 1, II.SoftVars("EOL"))
                                 bEOLAtEndOfLine = InStr(II.Expression, II.SoftVars("EOL"))
                                 II.Expression = sAfter(II.Expression, 1, II.SoftVars("EOL"))
                                 If InStr(II.Result, sOperator) Then
                                    Token1 = sGetToken(II.Result, 1, sOperator)
                                    Token2 = sAfter(II.Result, 1, sOperator)
                                    If Len(Token1) - Len(LTrim$(Token1)) > 0 Then
                                       .Value = .Value & String(Len(Token1) - Len(LTrim$(Token1)), " ") & Token2 & sOperator & LTrim$(Token1) & IIf(bEOLAtEndOfLine, II.SoftVars("EOL"), vbNullString)
                                    Else
                                       .Value = .Value & Token2 & sOperator & LTrim$(Token1) & IIf(bEOLAtEndOfLine, II.SoftVars("EOL"), vbNullString)
                                    End If
                                 Else
                                    .Value = II.Result & IIf(bEOLAtEndOfLine, II.SoftVars("EOL"), vbNullString)
                                 End If
                              Loop
                         End With
                         ISadAddin_ExecuteSoftCommand = True

                    Case "STRINGLENGTH", "LEN", "LENGTH", "STRLEN"
                         .SoftVars(.Result) = Len(.Expression)
                         ISadAddin_ExecuteSoftCommand = True

                    Case "MSGBOX"
                         Parent.UI.ZOrder
                         MsgBox Replace(Replace(.AllParameters, "$EOL$", vbNewLine), "$TAB$", vbTab)
                         ISadAddin_ExecuteSoftCommand = True

                    Case "DECIDE", "YESNO", "YESNOCANCEL", "ASKUSER"
                          Decision = MsgBox(II.Expression, vbYesNoCancel)
                          If Decision = vbCancel Then
                             II.LinesLeftToProcess = ""
                          Else
                             II.SoftVars(II.Result) = IIf(Decision = vbYes, "1", "0")
                          End If
            End Select
         End If
    End With

EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:
    Exit Function

EH_SADAddin_ISadAddin_ExecuteSoftCommand:
    LogError "SADAddin", "ISadAddin_ExecuteSoftCommand", Err.Number, Err.Description
    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue

    Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline
    Dim Template As CTemplate

    Dim Area                As Long
    Dim CurrSet             As Long
    Dim bCommandExecuted    As Boolean
    Dim lParamCount         As Long

    Dim Decision            As VbMsgBoxResult
    Dim CurrProject         As IProject
    Dim CurrComponent       As IComponent
    Dim CurrMember          As IMember

    Dim sDefault            As String
    Dim sT                  As String
    Dim sVar1               As String
    Dim sVar2               As String
    Dim sVar3               As String

    If Not MySadCommands(sInlineSoftCommandName & "*I") Is Nothing Then
       If Not MySadCommands(sInlineSoftCommandName & "*I").IsInline Then Exit Function
       Select Case sInlineSoftCommandName
              Case "MODULECONTENTS", "CONTENTS", "MEMBERS"
                    sResults = vbNullString
                    For Each CurrMember In Parent.SandyIDE.ActiveCodePane.CodeModule.Members
                        Select Case CurrMember.MemberType
                              'Case vbext_mt_Const:      sResults = sResults & CurrMember.Name & "=CONSTANT$$$$"
                              'Case vbext_mt_Variable:   sResults = sResults & CurrMember.Name & "=VARIABLE$$$$"
                               Case vbext_mt_Event:      sResults = sResults & CurrMember.Name & "=EVENT$$$$"
                               Case vbext_mt_Method:     sResults = sResults & CurrMember.Name & "=METHOD$$$$"
                               Case vbext_mt_Property:   sResults = sResults & CurrMember.Name & "=PROPERTY$$$$"
                        End Select
                    Next CurrMember
                    ISadAddin_ExecuteSoftCommandInline = True
                    
              Case "MODULELIST"
                    sResults = vbNullString
                    For Each CurrComponent In Parent.SandyIDE.ActiveProject.Components
                        Select Case CurrComponent
                               Case 1:    sResults = sResults & CurrComponent.Name & "=MODULE$$$$"
                               Case 2:    sResults = sResults & CurrComponent.Name & "=CLASS$$$$"
                               Case 3, 5: sResults = sResults & CurrComponent.Name & "=FORM$$$$"
                               Case 11:   sResults = sResults & CurrComponent.Name & "=DESIGNER$$$$"
                               Case 100:  sResults = sResults & CurrComponent.Name & "=DOCUMENT$$$$"
'                               Case vbext_ct_StdModule:         sResults = sResults & CurrComponent.Name & "=MODULE$$$$"
'                               Case vbext_ct_ClassModule:       sResults = sResults & CurrComponent.Name & "=CLASS$$$$"
'                               Case vbext_ct_MSForm, 5:         sResults = sResults & CurrComponent.Name & "=FORM$$$$"
'                               Case vbext_ct_ActiveXDesigner:   sResults = sResults & CurrComponent.Name & "=DESIGNER$$$$"
'                               Case 100:                        sResults = sResults & CurrComponent.Name & "=DOCUMENT$$$$"
                        End Select
                    Next CurrComponent
                    ISadAddin_ExecuteSoftCommandInline = True
                    
              Case "PROJECTLIST"
                    sResults = vbNullString
                    For Each CurrProject In Parent.SandyIDE.Projects
                        sResults = sResults & CurrProject.Name & "=" & CurrProject.FileName & "$$$$"
                    Next CurrProject
                    ISadAddin_ExecuteSoftCommandInline = True

              Case "DECIDE", "YESNO", "YESNOCANCEL", "ASKUSER"
                    sResults = IIf(MsgBox(sParameters, vbYesNo) = vbYes, "1", "0")
                    ISadAddin_ExecuteSoftCommandInline = True

              Case "LCASE", "LOWERCASE", "LWCASE"
                    If Len(SoftVars(sParameters)) > 0 Then
                       sResults = LCase(SoftVars(sParameters))
                    Else
                       sResults = LCase(sParameters)
                    End If
                    ISadAddin_ExecuteSoftCommandInline = True

              Case "UCASE", "UPPERCASE", "UPCASE"
                    If Len(SoftVars(sParameters)) > 0 Then
                       sResults = UCase(SoftVars(sParameters))
                    Else
                       sResults = UCase(sParameters)
                    End If
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "CHR"
                   If Len(sParameters) = 0 Then sParameters = "48"
                   If Val(sParameters) < 0 Then sParameters = "48"
                   If Val(sParameters) > 255 Then sParameters = "48"
                   sResults = Chr(Val(sParameters))
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "ASC"
                   If Len(sParameters) = 0 Then sParameters = "0"
                   sResults = Asc(sParameters)
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "STRINGLENGTH", "LEN", "LENGTH", "STRLEN", "TRIM", "LTRIM", "LEFTTRIM", "RTRIM", "RIGHTTRIM"               ' %%Function::SoftVarName%%
                   Select Case sInlineSoftCommandName
                          Case "STRINGLENGTH", "LEN", "LENGTH", "STRLEN"
                                                        sResults = Len(SoftVars(sParameters))
                          Case "TRIM":                  sResults = Trim$(SoftVars(sParameters))
                          Case "LTRIM", "LEFTTRIM":     sResults = LTrim$(SoftVars(sParameters))
                          Case "RTRIM", "RIGHTTRIM":    sResults = RTrim$(SoftVars(sParameters))
                   End Select
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "PAD"
                    lParamCount = lTokenCount(sParameters, ",")
                    sVar1 = sGetToken(sParameters, 1, ",")
                    If Len(SoftVars(sVar1)) > 0 Then
                       sVar1 = SoftVars(sVar1)
                    End If
                    If lParamCount > 1 Then
                       lParamCount = Val(sAfter(sParameters, 1, ","))
                    Else
                       lParamCount = 50
                    End If
                    If lParamCount < 1 Then lParamCount = 1
                    lParamCount = lParamCount - Len(sVar1)
                    If lParamCount > 0 Then
                       sResults = sVar1 & String$(lParamCount, " ")
                    Else
                       sResults = sVar1
                    End If
                    ISadAddin_ExecuteSoftCommandInline = True

              Case "TOKENCOUNT", "TCOUNT"                                                          ' %%Function::String,[Delim]%%
                   lParamCount = lTokenCount(sParameters, ",")
                   sVar1 = sGetToken(sParameters, 1, ",")
                   If lParamCount > 1 Then
                      sVar2 = sAfter(sParameters, 1, ",")
                   Else
                      sVar2 = " "
                   End If
                   sResults = lTokenCount(SoftVars(sVar1), sVar2)
                   ISadAddin_ExecuteSoftCommandInline = True

'              Case "TOKENCOUNT", "TCOUNT"                                                          ' %%Function::String,[Delim]%%
'                   lParamCount = lTokenCount(sParameters, ",")
'                   sVar1 = sGetToken(sParameters, 1, ",")
'                   If lParamCount > 1 Then
'                      sVar2 = sGetToken(sParameters, 2, ",")
'                      If Val(sVar2) = 0 Then
'                         sVar2 = " "
'                      End If
'                   Else
'                      sVar2 = " "
'                   End If
'                   sResults = lTokenCount(SoftVars(sVar1), sVar2)
'                   ISadAddin_ExecuteSoftCommandInline = True
'
              Case "REPLACE"
                   sVar1 = sGetToken(sParameters, 1, ",")
                   If Len(SoftVars(sVar1)) Then sVar1 = SoftVars(sVar1)
                   sVar2 = sGetToken(sParameters, 2, ",")
                   If Len(SoftVars(sVar2)) Then sVar2 = SoftVars(sVar2)
                   sVar3 = sGetToken(sParameters, 3, ",")
                   If Len(SoftVars(sVar3)) Then sVar3 = SoftVars(sVar3)
                   
                   sResults = Replace(sVar1, sVar2, sVar3)
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "TOKEN", "BEFORE", "AFTER", "EXCEPT", "EXTRACT", "LEFT", "RIGHT", "MID", "FIND", "FINDTOKEN"
                   lParamCount = lTokenCount(sParameters, ",")
                   sVar1 = sGetToken(sParameters, 1, ",")
                   If lParamCount > 1 Then
                      sVar2 = sGetToken(sParameters, 2, ",")
                      If Val(sVar2) = 0 And (sInlineSoftCommandName <> "FIND" And sInlineSoftCommandName <> "FINDTOKEN") Then
                         If Val(SoftVars(sVar2)) > 0 Then
                            sVar2 = SoftVars(sVar2)
                         Else
                            sVar2 = IIf(sInlineSoftCommandName = "BEFORE", 2, 1)
                         End If
                      End If
                      If lParamCount > 2 Then
                         sVar3 = sGetToken(sParameters, 3, ",")
                         If InStr(sVar3, """") Then
                            sVar3 = sGetToken(sVar3, 2, """")
                         End If
                      Else
                         sVar3 = IIf(sInlineSoftCommandName = "MID", "0", " ")
                      End If
                   Else
                      sVar2 = 1
                      sVar3 = IIf(sInlineSoftCommandName = "MID", "0", " ")
                   End If
                   
                   sVar1 = IIf(Len(SoftVars(sVar1)), SoftVars(sVar1), sVar1)
                   sVar2 = IIf(Len(SoftVars(sVar2)), SoftVars(sVar2), sVar2)
                   sVar3 = IIf(Len(SoftVars(sVar3)), SoftVars(sVar3), sVar3)

                   Select Case sInlineSoftCommandName
                          Case "TOKEN":    sResults = sGetToken(sVar1, Val(sVar2), sVar3)
                          Case "BEFORE":   sResults = sBefore(sVar1, Val(sVar2), sVar3)
                          Case "AFTER":
                               sResults = sAfter(sVar1, Val(sVar2), sVar3)
                          Case "EXCEPT":   sResults = sExcept(sVar1, Val(sVar2), sVar3)
                          Case "EXTRACT":  sResults = sGetToken(sVar1, Val(sVar2), sVar3)
                                           sVar1 = sExcept(sVar1, Val(sVar2), sVar3)
                          Case "LEFT":     sResults = Left(sVar1, Val(sVar2))
                          Case "RIGHT":    sResults = Right(sVar1, Val(sVar2))
                          Case "MID"
                               If Val(sVar3) = 0 Then
                                  sResults = Mid$(sVar1, Val(sVar2))
                               Else
                                  sResults = Mid$(sVar1, Val(sVar2), Val(sVar3))
                               End If
                          Case "FIND", "FINDTOKEN"
                               sResults = vbNullString & lFindToken(sVar1, sVar2, sVar3)
                   End Select
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "INCLUDE"
                   If Val(sGetToken(sParameters, 1, ",")) > 0 Then
                      Area = Val(sGetToken(sParameters, 1, ","))
                      sParameters = sAfter(sParameters, 1, ",")
                   Else
                      Area = 0
                   End If
                   Set Template = Parent.GetATemplate(sGetToken(sParameters, 1, " - "), sAfter(sParameters, 1, " - "))
                   If Not Template Is Nothing Then
                      With Template
                           Select Case Area
                                  Case 1:     sResults = .memoCodeAtTop
                                  Case 2:     sResults = .memoCodeAtCursor
                                  Case 3:     sResults = .memoCodeAtBottom
                                  Case 0
                                       sResults = vbNullString
                                       If Len(.memoCodeAtCursor) Then sResults = sResults & .memoCodeAtCursor
                                       If Len(.memoCodeAtTop) Then sResults = sResults & IIf(Len(sResults) And Right(sResults, 2) <> vbNewLine, vbNewLine, vbNullString) & "~~GotoDec End" & vbNewLine & .memoCodeAtTop
                                       If Len(.memoCodeAtBottom) Then sResults = sResults & IIf(Len(sResults) And Right(sResults, 2) <> vbNewLine, vbNewLine, vbNullString) & "~~GotoEnd" & vbNewLine & .memoCodeAtTop
                                  Case Else:  sResults = .memoCodeToFile
                           End Select
                      End With
                      If Len(sResults) Then
                         sResults = "~~'" & vbNewLine & sResults
                      End If
                      Set Template = Nothing
                      ISadAddin_ExecuteSoftCommandInline = True
                   End If

              Case "NOW"
                    If Len(sParameters) Then
                       On Error Resume Next
                          sResults = Format(Now(), sParameters)
                       On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline
                    Else
                       sResults = Format(Now(), "M/D/YYYY H:NN AM/PM")
                    End If
                    ISadAddin_ExecuteSoftCommandInline = True

              Case "CHOOSE"
                    If InStr(sParameters, IIf(Len(SoftVars("Choose Delimiter")) = 0, ";", SoftVars("Choose Delimiter"))) Then
                       sResults = Parent.sChoose(sParameters, IIf(Len(SoftVars("Choose Delimiter")) = 0, ";", SoftVars("Choose Delimiter")), SoftVars("Choose Default"))
                    ElseIf Len(SoftVars(sParameters)) Then
                       sResults = Parent.sChoose(SoftVars(sParameters))
                    Else
                       sResults = sParameters
                    End If
                    ISadAddin_ExecuteSoftCommandInline = True

               Case "MANGLE"
                     If Len(SoftVars(sParameters)) Then
                        sResults = Mangle(SoftVars(sParameters))
                     Else
                        sResults = Mangle(sParameters)
                     End If
                     ISadAddin_ExecuteSoftCommandInline = True
        End Select
    End If

EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue:
    Exit Function

EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline:
    LogError "SADAddin", "ISadAddin_ExecuteSoftCommandInline", Err.Number, Err.Description
    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue

    Resume
End Function

Private Property Get ISadAddin_Externals() As SandySupport.CAssocArray
    '
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

    Err.Clear
        MySadCommands.All = sFileContents(Parent.TemplateDatabasePath & "s2kAddin.txt")
        ISadAddin_Startup = (Err.Number = 0)
    Err.Clear
End Function
