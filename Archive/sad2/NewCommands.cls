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

Private mvbInst         As VBIDE.VBE
Private Parent          As SliceAndDice.Wizard
Private asaSaved        As SliceAndDice.CAssocArray
Private MySadCommands   As SliceAndDice.CSadCommands

Private Const gsC As String = ","

Implements SliceAndDice.ISadAddin


Private Sub Class_Initialize()
1    On Error Resume Next
2        Set MySadCommands = New SliceAndDice.CSadCommands
3        With MySadCommands
4             .ParameterDelimiter = gsC
5             .ParameterTypeDelimiter = ":"
6             .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision
7        End With
End Sub

Private Sub Class_Terminate()
8        Set MySadCommands = Nothing
End Sub

Private Property Get ISadAddin_CommandSet() As SliceAndDice.CSadCommands
9    On Error Resume Next
10       Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
    ' %%xmlList::
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
11   On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommand
12       Dim bEOLAtEndOfLine     As Boolean

13       Dim lWrapLength         As Long
14       Dim lThisWrap           As Long
15       Dim lLineOffset         As Long
16       Dim Decision            As VbMsgBoxResult

17       Dim sWordWrapped        As String
18       Dim sToWrap             As String
19       Dim Token1              As String
20       Dim Token2              As String
21       Dim SoftVar1            As String
22       Dim SoftVar2            As String
23       Dim sOperator           As String

24       With II
25            If Not MySadCommands(.SoftCommandName & "*C") Is Nothing Then
26               If MySadCommands(.SoftCommandName & "*C").IsInline Then Exit Function

            Select Case UCase(.SoftCommandName)
                        Case "SAVEVARS"
27                           Set asaSaved = New SliceAndDice.CAssocArray
28                           asaSaved.All = .SoftVars.All

29                      Case "RESTOREVARS"
30                           If Not asaSaved Is Nothing Then
                        '   MsgBox "You must execute the ~~SaveVars command first (See F1 help) before using the ~~RestoreVars command."
                        'Else
31                              asaSaved.KeyValueDelimiter = " saved="
32                              .SoftVars.All = Replace(asaSaved.All & vbNewLine & .SoftVars.All, "saved saved", "saved")
33                              Do While InStr(.SoftVars.All, "saved saved"): .SoftVars.All = Replace(.SoftVars.All, "saved saved", "saved"): Loop
34                              asaSaved.KeyValueDelimiter = " ="
35                           End If

36                      Case "APPEND"
37                           .SoftVars(.Result) = .SoftVars(.Result) & Replace(Replace(Replace(Replace(IIf(Len(.SoftVars(.Expression)) = 0, .Expression, .SoftVars(.Expression)), "$EOL$", vbNewLine), "$13$", Chr(13)), "$9$", Chr(9)), "$SP$", " ")
38                            ISadAddin_ExecuteSoftCommand = True
                   
39                      Case "CHOOSE"
40                            If InStr(.Expression, IIf(Len(.SoftVars("Choose Delimiter")) = 0, ";", .SoftVars("Choose Delimiter"))) Then
41                               .SoftVars(.Result) = Parent.sChoose(.Expression, IIf(Len(.SoftVars("Choose Delimiter")) = 0, ";", .SoftVars("Choose Delimiter")), .SoftVars("Choose Default"))
42                            ElseIf Len(.SoftVars(.Expression)) Then
43                               .SoftVars(.Result) = Parent.sChoose(.SoftVars(.Expression))
44                            End If
45                            ISadAddin_ExecuteSoftCommand = True

46                      Case "MANGLE"
47                            .SoftVars(.Result) = Mangle(.Expression)
48                            ISadAddin_ExecuteSoftCommand = True

49                      Case "LTRIM", "LEFTTRIM"
50                            .SoftVars(.AllParameters).Value = LTrim$(.SoftVars(.AllParameters))
51                            ISadAddin_ExecuteSoftCommand = True

52                      Case "RTRIM", "RIGHTTRIM"
53                            .SoftVars(.AllParameters).Value = RTrim$(.SoftVars(.AllParameters))
54                            ISadAddin_ExecuteSoftCommand = True

55                      Case "TRIM"
56                            .SoftVars(.AllParameters).Value = Trim$(.SoftVars(.AllParameters))
57                            ISadAddin_ExecuteSoftCommand = True

58                      Case "SET", "SETVALUE", "LET", "!"
59                            .SoftVars(.Result) = .Expression
                              ISadAddin_ExecuteSoftCommand = True

                        Case "LSET"
                              Token1 = .SoftVars(.Result)
                              LSet Token1 = .Expression
                              .SoftVars(.Result) = Token1
                              ISadAddin_ExecuteSoftCommand = True

                        Case "RSET"
                              Token1 = .SoftVars(.Result)
                              RSet Token1 = .Expression
                              .SoftVars(.Result) = Token1
                              ISadAddin_ExecuteSoftCommand = True

61                      Case "STRCOPY"
62                            If Len(.SoftVars(.Expression)) > 0 Then
63                               .SoftVars(.Result).Value = .SoftVars(.Expression)
64                            Else
65                               .SoftVars(.Result).Value = .Expression
66                            End If
67                            ISadAddin_ExecuteSoftCommand = True

68                      Case "LEFT"
69                            If Val(Parent.sGetToken(.Result)) > 0 Then
70                               .SoftVars(Parent.sAfter(.Result)) = Left$(.Expression, Val(Parent.sGetToken(.Result)))
71                            Else
72                               .SoftVars(.Result) = .Expression
73                            End If
74                            ISadAddin_ExecuteSoftCommand = True

75                      Case "MID"
76                            If Val(Parent.sGetToken(.Result)) > 0 Then
77                               If Val(Parent.sGetToken(.Result, 2)) > 0 Then
78                                  .SoftVars(Parent.sAfter(.Result)) = Mid$(.Expression, Val(Parent.sGetToken(.Result)), Val(Parent.sGetToken(.Result, 2)))
79                               Else
80                                  .SoftVars(.Result) = Mid$(.Expression, Val(Parent.sGetToken(.Result)))
81                               End If
82                            Else
83                               .SoftVars(.Result) = .Expression
84                            End If
85                            ISadAddin_ExecuteSoftCommand = True

86                      Case "RIGHT"
87                            If Val(Parent.sGetToken(.Result)) > 0 Then
88                               .SoftVars(Parent.sAfter(.Result)) = Right$(.Expression, Val(Parent.sGetToken(.Result)))
89                            Else
90                               .SoftVars(.Result) = .Expression
91                            End If
92                            ISadAddin_ExecuteSoftCommand = True

93                      Case "WORDWRAP"                    ' Wrap the indicated text at the indicated point, preceding each wrap with the indicated prefix
94                            lWrapLength = Val(Parent.sGetToken(II.AllParameters))
95                            sToWrap = Trim(Parent.sAfter(.sParam, 2, """"))
96                            .sParam = Parent.sGetToken(.sParam, 2, """")
97                            .LinesLeftToProcess = Parent.sAfter(.LinesLeftToProcess, 1, .SoftVars("EOL"))
98                            Do Until Len(sToWrap) = 0
99                               If Len(sToWrap) > lWrapLength Then
100                                 lThisWrap = lWrapLength
101                                 If lThisWrap > 0 Then
102                                    Do While Mid$(sToWrap, lThisWrap, 1) <> " "
103                                       lThisWrap = lThisWrap - 1
104                                    Loop
105                                 End If
106                                 sWordWrapped = sWordWrapped & .sParam & Left$(sToWrap, lThisWrap - 1) & .SoftVars("EOL")
107                                 sToWrap = Mid$(sToWrap, lThisWrap + 1)
108                              Else
109                                 sWordWrapped = sWordWrapped & .sParam & sToWrap & .SoftVars("EOL")
110                                 sToWrap = vbNullString
111                              End If
112                           Loop
113                           .LinesLeftToProcess = .SoftVars("EOL") & sWordWrapped & .LinesLeftToProcess
114                           ISadAddin_ExecuteSoftCommand = True

115                     Case "INCREMENT", "INC"
116                           With .SoftVars(Parent.sGetToken(.AllParameters))
117                                If Val(Parent.sAfter(II.AllParameters)) = 0 Then
118                                   .Value = Val(.Value) + 1
119                                Else
120                                   .Value = Val(.Value) + Val(Parent.sAfter(II.AllParameters))
121                                End If
122                           End With
123                           ISadAddin_ExecuteSoftCommand = True

124                      Case "DECREMENT", "DEC"
125                           With .SoftVars(Parent.sGetToken(.AllParameters))
126                                If Val(Parent.sAfter(II.AllParameters)) = 0 Then
127                                   .Value = Val(.Value) - 1
128                                Else
129                                   .Value = Val(.Value) - Val(Parent.sAfter(II.AllParameters))
130                                End If
131                           End With
132                           ISadAddin_ExecuteSoftCommand = True
                         
133                      Case "EVALUATE", "EVAL"
134                           .SoftVars(.Result) = Parent.Evaluate(.Expression, .SoftVars)
135                           ISadAddin_ExecuteSoftCommand = True

136                      Case "STRING", "STR", "REPSTR", "REPLICATESTRING"
137                           SoftVar1 = Parent.sGetToken(.AllParameters)
138                           lLineOffset = Parent.sGetToken(.AllParameters, 2)
139                           If lLineOffset < 1 Then lLineOffset = 1
140                           If InStr(UCase(.AllParameters), "CHR(") > 0 Then
141                              SoftVar2 = Chr(Val(Parent.sGetToken(Parent.sGetToken(UCase(.AllParameters), 2, "CHR("), 1, ")")))
142                              .SoftVars(SoftVar1) = String(lLineOffset, SoftVar2)
143                           Else
144                              SoftVar2 = Left$(Parent.sGetToken(.AllParameters, 2, """") & " ", 1)
145                              .SoftVars(SoftVar1) = String(lLineOffset, SoftVar2)
146                           End If
147                           ISadAddin_ExecuteSoftCommand = True

148                      Case "AUTOSPACE"
149                           .SoftVars(.Result) = Parent.sInsertSpaces(.Expression)
150                           ISadAddin_ExecuteSoftCommand = True

151                      Case "REVERSETOKENS", "FLIPFLOP"
152                           SoftVar1 = Parent.sGetToken(.AllParameters)
153                           SoftVar2 = Parent.sGetToken(.AllParameters, 2)
154                           sOperator = Parent.sGetToken(.AllParameters, 2, """")
155                           .Expression = .SoftVars(SoftVar1)
156                           With II.SoftVars(SoftVar2)
157                                .Value = vbNullString
158                                Do While Len(II.Expression)
159                                   II.Result = Parent.sGetToken(II.Expression, 1, II.SoftVars("EOL"))
160                                   bEOLAtEndOfLine = InStr(II.Expression, II.SoftVars("EOL"))
161                                   II.Expression = Parent.sAfter(II.Expression, 1, II.SoftVars("EOL"))
162                                   If InStr(II.Result, sOperator) Then
163                                      Token1 = Parent.sGetToken(II.Result, 1, sOperator)
164                                      Token2 = Parent.sAfter(II.Result, 1, sOperator)
165                                      If Len(Token1) - Len(LTrim(Token1)) > 0 Then
166                                         .Value = .Value & String(Len(Token1) - Len(LTrim(Token1)), " ") & Token2 & sOperator & LTrim(Token1) & IIf(bEOLAtEndOfLine, II.SoftVars("EOL"), vbNullString)
167                                      Else
168                                         .Value = .Value & Token2 & sOperator & LTrim(Token1) & IIf(bEOLAtEndOfLine, II.SoftVars("EOL"), vbNullString)
169                                      End If
170                                   Else
171                                      .Value = II.Result & IIf(bEOLAtEndOfLine, II.SoftVars("EOL"), vbNullString)
172                                   End If
173                                Loop
174                           End With
175                           ISadAddin_ExecuteSoftCommand = True

176                      Case "STRINGLENGTH", "LEN", "LENGTH", "STRLEN"
177                           .SoftVars(.Result) = Len(.Expression)
178                           ISadAddin_ExecuteSoftCommand = True

179                      Case "MSGBOX"
180                           Parent.SandyWindow.ZOrder
181                           MsgBox Replace(Replace(.AllParameters, "$EOL$", vbNewLine), "$TAB$", vbTab)
182                           ISadAddin_ExecuteSoftCommand = True
            
183                      Case "DECIDE", "YESNO", "YESNOCANCEL", "ASKUSER"
184                           Decision = MsgBox(IIf(Len(II.SoftVars(II.Expression)) > 0, II.SoftVars(II.Expression), II.Expression), vbYesNoCancel)
185                           If Decision = vbCancel Then
186                              II.LinesLeftToProcess = vbNullString
187                           Else
188                              II.SoftVars(II.Result) = IIf(Decision = vbYes, "1", "0")
189                           End If
190                           ISadAddin_ExecuteSoftCommand = True

191                      Case "REPLACE"
192                           If InStr(II.AllParameters, "=") = 0 Then
193                              SoftVar1 = II.AllParameters
194                           Else
195                              SoftVar1 = II.Expression
196                           End If
197                           If Len(II.SoftVars(SoftVar1)) Then SoftVar1 = II.SoftVars(SoftVar1)

198                           If Len(SoftVar1) > 0 And Len(II.SoftVars("ToFind")) > 0 Then
199                              SoftVar2 = II.SoftVars("ToFind")
                               'If Len(II.SoftVars(SoftVar2)) Then SoftVar2 = II.SoftVars(SoftVar2)
200                              Token1 = II.SoftVars("ReplaceWith")
                               'If Len(II.SoftVars(Token1)) Then Token1 = II.SoftVars(Token1)
201                              If InStr(II.AllParameters, "=") = 0 Then
202                                 II.SoftVars(II.AllParameters) = Replace(SoftVar1, SoftVar2, Token1)
203                              Else
204                                 II.SoftVars(II.Result) = Replace(SoftVar1, SoftVar2, Token1)
205                              End If
206                           End If
207                           ISadAddin_ExecuteSoftCommand = True
208              End Select
209           End If
210      End With

211 EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:
212      Exit Function

213 EH_SADAddin_ISadAddin_ExecuteSoftCommand:
214      ErrorsOcurred II, SoftVar1, SoftVar2, Token1, Token2

215      Resume EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue
216      Resume
End Function

Private Function Mangle(strIn As String) As String
217  On Error Resume Next
218      Dim bPreviousWasAnUnderscore    As Boolean
219      Dim sOrig                       As String
220      Dim CurrChar                    As Long
221      Dim LenOrig                     As Long
222      Dim sChar                       As String
223      Dim sOut                        As String

224      sOrig = strIn
225      LenOrig = Len(sOrig)
226      bPreviousWasAnUnderscore = True
227      sOut = vbNullString

228      If InStr(strIn, "_") > 0 Then
229         For CurrChar = 1 To LenOrig
230             sChar = Mid$(strIn, CurrChar, 1)
           Select Case sChar
                  Case "0" To "9"
231                         sOut = sOut & sChar
232                         bPreviousWasAnUnderscore = False
   
233                    Case "A" To "Z", "a" To "z"
234                         If bPreviousWasAnUnderscore Then
235                            sOut = sOut & UCase(sChar)
236                         Else
237                            sOut = sOut & LCase(sChar)
238                         End If
239                         bPreviousWasAnUnderscore = False
   
240                    Case "_"
241                         bPreviousWasAnUnderscore = True
242             End Select
243         Next CurrChar
244      Else
245         sOut = strIn
246      End If
247      Mangle = sOut
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
248  On Error GoTo ErrorHandler
249      Dim Template                As CTemplate

250      Dim Area                    As Long
253      Dim lParamCount             As Long
         Dim lProcType       As Long

'254      Dim Decision                As VbMsgBoxResult
255      Dim CurrProject             As VBProject
256      Dim CurrModule              As VBComponent
257      Dim CurrMember              As Member

   'Dim sDefault                As String
258      Dim sT                      As String
259      Dim sVar1                   As String
260      Dim sVar2                   As String
261      Dim sVar3                   As String
262      Dim ProcType                As vbext_ProcKind

263      If Not MySadCommands(sInlineSoftCommandName & "*I") Is Nothing Then
264         If Not MySadCommands(sInlineSoftCommandName & "*I").IsInline Then Exit Function
       Select Case sInlineSoftCommandName
              Case "MODULECONTENTS", "CONTENTS", "MEMBERS"
265                 sResults = vbNullString
266                 For Each CurrMember In mvbInst.ActiveCodePane.CodeModule.Members
                        Select Case CurrMember.Type
                               Case vbext_mt_Const     ':   sResults = sResults & CurrMember.Name & "=CONSTANT=" & X & "$$$$"
267                            Case vbext_mt_Variable  ':   sResults = sResults & CurrMember.Name & "=VARIABLE=" & X & "=" & X & "=" & X & "=$$$$"
268                            Case Else
                                    Select Case CurrMember.Type
                                           Case vbext_mt_Event:    sResults = sResults & CurrMember.Name & "=EVENT=="
270                                        Case vbext_mt_Method:   sResults = sResults & CurrMember.Name & "=METHOD=="
272                                        Case vbext_mt_Property: sT = mvbInst.ActiveCodePane.CodeModule.ProcOfLine(CurrMember.CodeLocation, lProcType)
274                                                                sResults = sResults & CurrMember.Name & "=PROPERTY=" & Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString) & "="
275                                 End Select
276                                 sResults = sResults & CurrMember.CodeLocation & "=" & CurrMember.UIDefault & "=" & CurrMember.Hidden & "=" & CurrMember.Scope & "=$$$$"
277                     End Select
278                 Next CurrMember
279                 ISadAddin_ExecuteSoftCommandInline = True

280                Case "MODULELIST"
281                      sResults = vbNullString
282                      For Each CurrModule In mvbInst.ActiveVBProject.VBComponents
                        Select Case CurrModule.Type
                               Case vbext_ct_StdModule:         sResults = sResults & CurrModule.Name & "=MODULE$$$$"
283                                 Case vbext_ct_ClassModule:       sResults = sResults & CurrModule.Name & "=CLASS$$$$"
284                                 Case vbext_ct_MSForm, 5:         sResults = sResults & CurrModule.Name & "=FORM$$$$"
285                                 Case vbext_ct_ActiveXDesigner:   sResults = sResults & CurrModule.Name & "=DESIGNER$$$$"
286                                 Case 100:                        sResults = sResults & CurrModule.Name & "=DOCUMENT$$$$"
287                          End Select
288                      Next CurrModule
289                      ISadAddin_ExecuteSoftCommandInline = True
                    
290                Case "PROJECTLIST"
291                      sResults = vbNullString
292                      For Each CurrProject In mvbInst.VBProjects
293                          sResults = sResults & CurrProject.Name & "=" & CurrProject.FileName & "$$$$"
294                      Next CurrProject
295                      ISadAddin_ExecuteSoftCommandInline = True

296                Case "DECIDE", "YESNO", "YESNOCANCEL", "ASKUSER"
297                      sResults = IIf(MsgBox(IIf(Len(SoftVars(sParameters)) > 0, SoftVars(sParameters), sParameters), vbYesNo) = vbYes, "1", "0")
298                      ISadAddin_ExecuteSoftCommandInline = True

299                Case "LCASE", "LOWERCASE", "LWCASE"
300                      If Len(SoftVars(sParameters)) > 0 Then
301                         sResults = LCase(SoftVars(sParameters))
302                      Else
303                         sResults = LCase(sParameters)
304                      End If
305                      ISadAddin_ExecuteSoftCommandInline = True

306                Case "UCASE", "UPPERCASE", "UPCASE"
307                      If Len(SoftVars(sParameters)) > 0 Then
308                         sResults = UCase(SoftVars(sParameters))
309                      Else
310                         sResults = UCase(sParameters)
311                      End If
312                     ISadAddin_ExecuteSoftCommandInline = True

313                Case "PAD"
314                      lParamCount = Parent.lTokenCount(sParameters, gsC)
315                      sVar1 = Parent.sGetToken(sParameters, 1, gsC)
316                      If Len(SoftVars(sVar1)) > 0 Then
317                         sVar1 = SoftVars(sVar1)
318                      End If
319                      If lParamCount > 1 Then
320                         lParamCount = Val(Parent.sAfter(sParameters, 1, gsC))
321                      Else
322                         lParamCount = 50
323                      End If
324                      If lParamCount < 1 Then lParamCount = 1
325                      lParamCount = lParamCount - Len(sVar1)
326                      If lParamCount > 0 Then
327                         sResults = sVar1 & String$(lParamCount, " ")
328                      Else
329                         sResults = sVar1
330                      End If
331                      ISadAddin_ExecuteSoftCommandInline = True

332                Case "CHR"
333                     If Len(sParameters) = 0 Then sParameters = "48"
334                     If Val(sParameters) < 0 Then sParameters = "48"
335                     If Val(sParameters) > 255 Then sParameters = "48"
336                     sResults = Chr(Val(sParameters))
337                     ISadAddin_ExecuteSoftCommandInline = True

338                Case "ASC"
339                     If Len(sParameters) = 0 Then sParameters = "0"
340                     sResults = Asc(sParameters)
341                     ISadAddin_ExecuteSoftCommandInline = True

342                Case "STRINGLENGTH", "LEN", "LENGTH", "STRLEN", "TRIM", "LTRIM", "LEFTTRIM", "RTRIM", "RIGHTTRIM"               ' %%Function::SoftVarName%%
                   Select Case sInlineSoftCommandName
                          Case "STRINGLENGTH", "LEN", "LENGTH", "STRLEN"
343                                                          sResults = Len(SoftVars(sParameters))
344                            Case "TRIM":                  sResults = Trim(SoftVars(sParameters))
345                            Case "LTRIM", "LEFTTRIM":     sResults = LTrim(SoftVars(sParameters))
346                            Case "RTRIM", "RIGHTTRIM":    sResults = RTrim(SoftVars(sParameters))
347                     End Select
348                     ISadAddin_ExecuteSoftCommandInline = True

                   Case "INSTR"                                                          ' %%Function::String,[Delim]%%
                        lParamCount = Parent.lTokenCount(sParameters, gsC)
                        sVar1 = Parent.sGetToken(sParameters, 1, gsC)
                        If lParamCount > 1 Then
                           sVar2 = Parent.sAfter(sParameters, 1, gsC)
                        Else
                           sVar2 = " "
                        End If
                        sVar1 = IIf(Len(SoftVars(sVar1)), SoftVars(sVar1), sVar1)
                        sVar2 = IIf(Len(SoftVars(sVar2)), SoftVars(sVar2), sVar2)
                        sResults = InStr(sVar1, sVar2)
                        ISadAddin_ExecuteSoftCommandInline = True

349                Case "TOKENCOUNT", "TCOUNT"                                                          ' %%Function::String,[Delim]%%
350                     lParamCount = Parent.lTokenCount(sParameters, gsC)
351                     sVar1 = Parent.sGetToken(sParameters, 1, gsC)
352                     If lParamCount > 1 Then
353                        sVar2 = Parent.sAfter(sParameters, 1, gsC)
354                     Else
355                        sVar2 = " "
356                     End If
                        sVar1 = IIf(Len(SoftVars(sVar1)), SoftVars(sVar1), sVar1)
                        sVar2 = IIf(Len(SoftVars(sVar2)), SoftVars(sVar2), sVar2)
357                     sResults = Parent.lTokenCount(sVar1, sVar2)
358                     ISadAddin_ExecuteSoftCommandInline = True

359                Case "REPLACE"
360                     sVar1 = Parent.sGetToken(sParameters, 1, gsC)
361                     If Len(SoftVars(sVar1)) Then sVar1 = SoftVars(sVar1)
362                     sVar2 = Parent.sGetToken(sParameters, 2, gsC)
363                     If Len(SoftVars(sVar2)) Then sVar2 = SoftVars(sVar2)
364                     sVar3 = Parent.sGetToken(sParameters, 3, gsC)
365                     If Len(SoftVars(sVar3)) Then sVar3 = SoftVars(sVar3)
                   
366                     sResults = Replace(sVar1, sVar2, sVar3)
367                     ISadAddin_ExecuteSoftCommandInline = True

368                Case "TOKEN", "BEFORE", "AFTER", "EXCEPT", "EXTRACT", "LEFT", "RIGHT", "MID", "FIND", "FINDTOKEN"
369                     lParamCount = Parent.lTokenCount(sParameters, gsC)
370                     sVar1 = Parent.sGetToken(sParameters, 1, gsC)
371                     If lParamCount > 1 Then
372                        sVar2 = Parent.sGetToken(sParameters, 2, gsC)
373                        If Val(sVar2) = 0 And (sInlineSoftCommandName <> "FIND" And sInlineSoftCommandName <> "FINDTOKEN") Then
374                           If Val(SoftVars(sVar2)) > 0 Then
375                              sVar2 = SoftVars(sVar2)
376                           Else
377                              sVar2 = IIf(sInlineSoftCommandName = "BEFORE", 2, 1)
378                           End If
379                        End If
380                        If lParamCount > 2 Then
381                           sVar3 = Parent.sGetToken(sParameters, 3, gsC)
382                           If InStr(sVar3, """") Then
383                              sVar3 = Parent.sGetToken(sVar3, 2, """")
384                           End If
385                        Else
386                           sVar3 = IIf(sInlineSoftCommandName = "MID", "0", " ")
387                        End If
388                     Else
389                        sVar2 = 1
390                        sVar3 = IIf(sInlineSoftCommandName = "MID", "0", " ")
391                     End If
                   
392                     sVar1 = IIf(Len(SoftVars(sVar1)), SoftVars(sVar1), sVar1)
393                     sVar2 = IIf(Len(SoftVars(sVar2)), SoftVars(sVar2), sVar2)
394                     sVar3 = IIf(Len(SoftVars(sVar3)), SoftVars(sVar3), sVar3)

                   Select Case sInlineSoftCommandName
                          Case "TOKEN":    sResults = Parent.sGetToken(sVar1, Val(sVar2), sVar3)
395                            Case "BEFORE":   sResults = Parent.sBefore(sVar1, Val(sVar2), sVar3)
396                            Case "AFTER":
397                                 sResults = Parent.sAfter(sVar1, Val(sVar2), sVar3)
398                            Case "EXCEPT":   sResults = Parent.sExcept(sVar1, Val(sVar2), sVar3)
399                            Case "EXTRACT":  sResults = Parent.sGetToken(sVar1, Val(sVar2), sVar3)
400                                             sVar1 = Parent.sExcept(sVar1, Val(sVar2), sVar3)
401                            Case "LEFT":     sResults = Left$(sVar1, Val(sVar2))
402                            Case "RIGHT":    sResults = Right$(sVar1, Val(sVar2))
403                            Case "MID"
404                                 If Val(sVar3) = 0 Then
405                                    sResults = Mid$(sVar1, Val(sVar2))
406                                 Else
407                                    sResults = Mid$(sVar1, Val(sVar2), Val(sVar3))
408                                 End If
409                            Case "FIND", "FINDTOKEN"
410                                 sResults = vbNullString & Parent.lFindToken(sVar1, sVar2, sVar3)
411                     End Select
412                     ISadAddin_ExecuteSoftCommandInline = True

413                Case "INCLUDE"
414                     If Val(Parent.sGetToken(sParameters, 1, gsC)) > 0 Then
415                        Area = Val(Parent.sGetToken(sParameters, 1, gsC))
416                        sParameters = Parent.sAfter(sParameters, 1, gsC)
417                     Else
418                        Area = 0
419                     End If

420                     If InStr(sParameters, " - ") = 0 Then
421                        sParameters = "**CURRENT CATEGORY** - " & sParameters
422                     End If

423                     Set Template = Parent.GetATemplate(Parent.sGetToken(sParameters, 1, " - "), Parent.sAfter(sParameters, 1, " - "))
424                     If Not Template Is Nothing Then
425                        With Template
                           Select Case Area
                                  Case 1:     sResults = .memoCodeAtTop
426                                    Case 2:     sResults = .memoCodeAtCursor
427                                    Case 3:     sResults = .memoCodeAtBottom
428                                    Case 0
429                                         sResults = vbNullString
430                                         If Len(.memoCodeAtCursor) Then sResults = sResults & .memoCodeAtCursor
431                                         If Len(.memoCodeAtTop) Then sResults = sResults & IIf(Len(sResults) And Right$(sResults, 2) <> vbNewLine, vbNewLine, vbNullString) & "~~GotoDec End" & vbNewLine & .memoCodeAtTop
432                                         If Len(.memoCodeAtBottom) Then sResults = sResults & IIf(Len(sResults) And Right$(sResults, 2) <> vbNewLine, vbNewLine, vbNullString) & "~~GotoEnd" & vbNewLine & .memoCodeAtTop
433                                    Case Else:  sResults = .memoCodeToFile
434                             End Select
435                        End With
436                        If Len(sResults) Then
437                           sResults = "~~'" & vbNewLine & sResults
438                        End If
439                        Set Template = Nothing
440                        ISadAddin_ExecuteSoftCommandInline = True
441                     End If

442                Case "NOW"
443                      If Len(sParameters) Then
444                         On Error Resume Next
445                            sResults = Format(Now(), sParameters)
446                         On Error GoTo ErrorHandler
447                      Else
448                         sResults = Format(Now(), "M/D/YYYY H:NN AM/PM")
449                      End If
450                      ISadAddin_ExecuteSoftCommandInline = True

451                Case "CHOOSE"
452                      If InStr(sParameters, IIf(Len(SoftVars("Choose Delimiter")) = 0, ";", SoftVars("Choose Delimiter"))) Then
453                         sResults = Parent.sChoose(sParameters, IIf(Len(SoftVars("Choose Delimiter")) = 0, ";", SoftVars("Choose Delimiter")), SoftVars("Choose Default"))
454                      ElseIf Len(SoftVars(sParameters)) Then
455                         sResults = Parent.sChoose(SoftVars(sParameters))
456                      Else
457                         sResults = sParameters
458                      End If
459                      ISadAddin_ExecuteSoftCommandInline = True

460                 Case "MANGLE"
461                       If Len(SoftVars(sParameters)) Then
462                          sResults = Mangle(SoftVars(sParameters))
463                       Else
464                          sResults = Mangle(sParameters)
465                       End If
466                       ISadAddin_ExecuteSoftCommandInline = True
467          End Select
468      End If

469 ErrorHandler_Continue:
470      Exit Function

471 ErrorHandler:
472      Dim X As CInsertionInfo

473      Set X = New CInsertionInfo
474      With X
475           .AllParameters = sParameters
476           .Result = sResults
477           Set .SoftVars = SoftVars
478           .SoftCommandName = sInlineSoftCommandName
479           .CurrentLineToProcess = "-Inline Substition-"
480      End With

481      ErrorsOcurred X, sVar1, sVar2, sVar3, "-Inline Soft Command-"

482      Set X.SoftVars = Nothing
483      Set X = Nothing

484      Resume ErrorHandler_Continue
485      Resume
End Function

Private Property Get ISadAddin_Externals() As SliceAndDice.CAssocArray
    '
End Property

Private Function ISadAddin_Shutdown() As Boolean
486  On Error Resume Next
487      MySadCommands.Clear
488      Set MySadCommands.Parent = Nothing
489      Set MySadCommands = Nothing
490      Set Parent = Nothing

491      ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard, vbInst As VBIDE.VBE) As Boolean
492  On Error Resume Next
493      Err.Clear

    'If Now() > 36300 Then
    '   MsgBox "The Slice and Dice evaluation has expired. Thank you for participating." & Chr(13) & Chr(13) & "Please download the latest evaluation from: http://www.sliceanddice.com/VB5CodeWalker"
    '   Exit Function
    'End If

    'If UCase(oParent.Version) <> UCase(MySadCommands.Attributes("Version")) Then
    '   If MsgBox("Version mismatch:" & Chr(13) & Chr(9) & "SADAddin.NewCommands = " & MySadCommands.Attributes("Version") & Chr(13) & Chr(9) & "Slice and Dice = " & oParent.Version & Chr(13) & Chr(13) & "Continue loading DLL ?", vbYesNo, "*** WARNING - DLL Version mismatch **") = vbNo Then
    '      Exit Function
    '   End If
    'End If

494      Set Parent = oParent
495      Set MySadCommands.Parent = oParent
    
496      Set mvbInst = vbInst

497      Err.Clear
498          MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath & "SADAddin.txt")
499          ISadAddin_Startup = (Err.Number = 0)
500      Err.Clear
End Function

Private Function ErrorsOcurred(ByRef II As CInsertionInfo, ByVal sVar1 As String, ByVal sVar2 As String, ByVal sVar3 As String, ByVal sVar4 As String) As Boolean
501  On Error Resume Next
502      Dim fh          As Long
503      Dim sMessage    As String

504      If Err.Number <> 0 Then
505         sMessage = "Error executing SoftCode:" & vbNewLine
506         sMessage = sMessage & vbTab & "Occured:      " & Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM")
507         sMessage = sMessage & vbTab & "Sandal:       sadAddin.NewCommands" & vbNewLine
508         If Erl <> 0 Then sMessage = sMessage & vbTab & "Sandal Line:  " & Erl & vbNewLine
509         sMessage = sMessage & vbTab & "Error Number: " & Err.Number & vbNewLine
510         sMessage = sMessage & vbTab & "Description:  " & Err.Description & vbNewLine

511         sMessage = sMessage & vbNewLine & vbTab & "SoftCode Parameters (Resolved):" & vbNewLine
512         sMessage = sMessage & vbTab & vbTab & "sVar1 = " & sVar1 & vbNewLine
513         sMessage = sMessage & vbTab & vbTab & "sVar2 = " & sVar2 & vbNewLine
514         sMessage = sMessage & vbTab & vbTab & "sVar3 = " & sVar3 & vbNewLine
515         sMessage = sMessage & vbTab & vbTab & "sVar4 = " & sVar4 & vbNewLine
 
516         If Not II Is Nothing Then
517            sMessage = sMessage & vbNewLine & vbTab & "(Unresolved) CInsertionInfo contents:" & vbNewLine
518            sMessage = sMessage & vbTab & vbTab & "TemplateName = " & II.TemplateName & vbNewLine
519            sMessage = sMessage & vbTab & vbTab & "CurrentLineToProcess = " & II.CurrentLineToProcess & vbNewLine
520            sMessage = sMessage & vbTab & vbTab & "SoftCommandName = " & II.SoftCommandName & vbNewLine
521            sMessage = sMessage & vbTab & vbTab & "AllParameters = " & II.AllParameters & vbNewLine
522            sMessage = sMessage & vbTab & vbTab & "Result     (LHS) = " & II.Result & vbNewLine
523            sMessage = sMessage & vbTab & vbTab & "Expression (RHS) = " & II.Expression & vbNewLine
524            sMessage = sMessage & vbTab & vbTab & "ExternalFilename = " & II.ExternalFilename & vbNewLine
525         End If

526         sMessage = sMessage & vbNewLine & vbNewLine & "Would you like to cancel processing ?" & vbNewLine
527         sMessage = sMessage & vbNewLine & vbTab & "IF YOU SELECT:" & vbNewLine
528         sMessage = sMessage & vbTab & vbTab & "YES: This template should be cancelled."
529         sMessage = sMessage & vbTab & vbTab & vbTab & "NOTE: Due to the nature of this addin, processing is not guarenteed to stop."
530         sMessage = sMessage & vbTab & vbTab & "NO : Processing will continue with the next line of SoftCode."
531         sMessage = sMessage & vbTab & "NOTE: This information will be stored in: """ & App.Path & "\sadAddinError.Log"""

532         If MsgBox(sMessage, vbYesNo, "CANCEL PROCESSING ?") = vbYes Then
533            If Not II Is Nothing Then
534               II.LinesLeftToProcess = vbNewLine & vbNewLine & "~~' Cancel inserted at user's request" & vbNewLine & "~~Cancel 1" & vbNewLine
535            End If
536            sMessage = sMessage & vbNewLine & "  *** User decided to CANCEL processing."
537            ErrorsOcurred = True ' Tell parent function processing has been cancelled.
538         Else
539            sMessage = sMessage & vbNewLine & "  *** User choose to CONTINUE after error."
540            Err.Clear
541         End If

542         fh = FreeFile
543         Open App.Path & "\sadAddinError.Log" For Append As #fh
544              Print #fh, sMessage
545         Close #fh
546      End If
End Function

