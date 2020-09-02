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
Attribute VB_Description = "Sandle - File System Commands"
Option Explicit

Implements SliceAndDice.ISadAddin

Private Parent          As SliceAndDice.Wizard
Private MySadCommands   As SliceAndDice.CSadCommands

Private fh As Long      ' File handle that can hang around between calls
Private Sub Class_Initialize()
1    On Error Resume Next
2        Set MySadCommands = New SliceAndDice.CSadCommands
3        With MySadCommands
4             .ParameterDelimiter = ","
5             .ParameterTypeDelimiter = ":"
6             .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision
7        End With

8        fh = 0
End Sub

Private Sub Class_Terminate()
9    On Error Resume Next
10       If fh <> 0 Then
11          Close #fh
12          fh = 0
13       End If

14       Set MySadCommands = Nothing
End Sub

Private Property Get ISadAddin_CommandSet() As SliceAndDice.CSadCommands
15   On Error Resume Next
16       Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
'    Select Case UCase$(sKey)
'           Case "TESTING EXTERNALS"
'                MsgBox sValue
'    End Select
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
17   On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommand
18       Dim sVar1 As String
19       Dim sVar2 As String
20       Dim sVar3 As String
21       Dim sVar4 As String

22       With II
23            If Not MySadCommands(.SoftCommandName & "*C") Is Nothing Then

24               If MySadCommands(.SoftCommandName & "*C").IsInline Then Exit Function

25               sVar1 = Parent.sGetToken(II.AllParameters, 1, ","): If Len(.SoftVars(sVar1)) Then sVar1 = .SoftVars(sVar1)
26               sVar2 = Parent.sGetToken(II.AllParameters, 2, ","): If Len(.SoftVars(sVar2)) Then sVar2 = .SoftVars(sVar2)
27               sVar3 = Parent.sGetToken(II.AllParameters, 3, ","): If Len(.SoftVars(sVar3)) Then sVar3 = .SoftVars(sVar3)
28               sVar4 = Parent.sGetToken(II.AllParameters, 4, ","): If Len(.SoftVars(sVar4)) Then sVar4 = .SoftVars(sVar4)

29   On Error Resume Next
30               Err.Clear
            Select Case UCase$(.SoftCommandName)
                   Case "CHANGEDRIVE", "DRIVE", "CHDRIVE"
31                            ChDrive sVar1
32                            If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
33                               If fh <> 0 Then Close #fh: fh = 0
34                               ISadAddin_ExecuteSoftCommand = True
35                               Exit Function
36                            End If

37                            ISadAddin_ExecuteSoftCommand = True

38                      Case "CHANGEDIRECTORY", "CHANGEDIR", "CHDIR", "CD"
39                            ChDir sVar1
40                            If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
41                               If fh <> 0 Then Close #fh: fh = 0
42                               ISadAddin_ExecuteSoftCommand = True
43                               Exit Function
44                            End If

45                            ISadAddin_ExecuteSoftCommand = True
                        
46                      Case "CREATEFILE", "MAKEFILE", "MKFILE"
47                            fh = FreeFile
48                            Open sVar1 For Output Access Write As #fh
49                            If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
50                               If fh <> 0 Then Close #fh: fh = 0
51                               ISadAddin_ExecuteSoftCommand = True
52                               Exit Function
53                            End If

54                                 If Len(sVar2) > 0 Then
55                                    Print #fh, sVar2
56                                 End If

57                            If fh <> 0 Then Close #fh: fh = 0

58                            ISadAddin_ExecuteSoftCommand = True

59                      Case "APPENDTOFILE", "APPENDFILE"
60                            fh = FreeFile
61                            Open sVar1 For Append Access Write As #fh
62                            If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
63                               If fh <> 0 Then Close #fh: fh = 0
64                               ISadAddin_ExecuteSoftCommand = True
65                               Exit Function
66                            End If

67                                 If Len(sVar2) > 0 Then
68                                    Print #fh, sVar2
69                                 End If
70                            If fh <> 0 Then Close #fh: fh = 0
71                            ISadAddin_ExecuteSoftCommand = True

72                      Case "DELETEFILE", "DELFILE", "DEL"
73                            If Len(Dir(sVar1)) > 0 Then
74                               Err.Clear
75                               Kill sVar1
76                            End If
77                            If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
78                               If fh <> 0 Then Close #fh: fh = 0
79                               ISadAddin_ExecuteSoftCommand = True
80                               Exit Function
81                            End If

82                            ISadAddin_ExecuteSoftCommand = True

83                      Case "CREATEFOLDER", "CREATEDIRECTORY", "MAKEDIR", "MKDIR", "MD"
84                            MkDir sVar1

85                            If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
86                               If fh <> 0 Then Close #fh: fh = 0
87                               ISadAddin_ExecuteSoftCommand = True
88                               Exit Function
89                            End If

90                            ISadAddin_ExecuteSoftCommand = True

91                      Case "DELETEFOLDER", "DELETEDIRECTORY", "DELDIR", "DELETEDIR"
92                            RmDir sVar1

93                            If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
94                               If fh <> 0 Then Close #fh: fh = 0
95                               ISadAddin_ExecuteSoftCommand = True
96                               Exit Function
97                            End If

98                            ISadAddin_ExecuteSoftCommand = True

99                      Case "RENAME", "REN", "RENAMEFILE"
100                           Name sVar1 As sVar2

101                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
102                              If fh <> 0 Then Close #fh: fh = 0
103                              ISadAddin_ExecuteSoftCommand = True
104                              Exit Function
105                           End If

106                           ISadAddin_ExecuteSoftCommand = True

107                     Case "FILECONTENTS", "GETFILE", "FILETOSOFTVAR", "FILETOVAR"
108                           II.SoftVars(II.Result) = Parent.sFileContents(II.Expression)

109                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
110                              If fh <> 0 Then Close #fh: fh = 0
111                              ISadAddin_ExecuteSoftCommand = True
112                              Exit Function
113                           End If

114                           ISadAddin_ExecuteSoftCommand = True

115                     Case "VARTOFILE", "PUTFILE", "SAVETOFILE"
116                           fh = FreeFile
117                           Open II.Result For Output Access Write As #fh

118                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
119                              If fh <> 0 Then Close #fh: fh = 0
120                              ISadAddin_ExecuteSoftCommand = True
121                              Exit Function
122                           End If

123                                Print #fh, II.SoftVars(II.Expression)

124                           Close #fh
125                           fh = 0

126                           ISadAddin_ExecuteSoftCommand = True

127                     Case "APPENDVARTOFILE", "APPENDVAR"
128                           fh = FreeFile
129                           Open II.Result For Output Access Write As #fh

130                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
131                              If fh <> 0 Then Close #fh: fh = 0
132                              ISadAddin_ExecuteSoftCommand = True
133                              Exit Function
134                           End If

135                                Print #fh, II.SoftVars(II.Expression)
136                           If fh <> 0 Then Close #fh: fh = 0
137                           ISadAddin_ExecuteSoftCommand = True

138                     Case "RUNPROGRAM", "SHELL", "EXECUTE"
139                           II.SoftVars("SHELLID") = Shell(sVar1, vbNormalFocus)

140                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
141                              If fh <> 0 Then Close #fh: fh = 0
142                              ISadAddin_ExecuteSoftCommand = True
143                              Exit Function
144                           End If

145                           ISadAddin_ExecuteSoftCommand = True

146                     Case "ACTIVATE", "APPACTIVATE", "ACTIVATEAPPLICATION", "SWITCHTOAPP"
147                           AppActivate sVar1

148                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
149                              If fh <> 0 Then Close #fh: fh = 0
150                              ISadAddin_ExecuteSoftCommand = True
151                              Exit Function
152                           End If

153                           ISadAddin_ExecuteSoftCommand = True

154                     Case "SENDKEYS"
155                           SendKeys sVar1

156                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
157                              If fh <> 0 Then Close #fh: fh = 0
158                              ISadAddin_ExecuteSoftCommand = True
159                              Exit Function
160                           End If

161                           ISadAddin_ExecuteSoftCommand = True

162                     Case "SENDKEYSWITHPAUSE", "SENDKEYSANDWAIT", "SENDKEYSANDPAUSE", "SENDKEYSWITHWAIT"
163                           SendKeys sVar1, True

164                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
165                              If fh <> 0 Then Close #fh: fh = 0
166                              ISadAddin_ExecuteSoftCommand = True
167                              Exit Function
168                           End If

169                           ISadAddin_ExecuteSoftCommand = True

170                     Case "CHECKFORFILE", "PAUSEFORFILE", "WAITFORFILE", "WAITFOR"
171                           MsgBox "CHECKFORFILE not currently implemented (sorry, please request completion)."

172                           ISadAddin_ExecuteSoftCommand = True

173                     Case "PAUSE", "WAIT"
                         'If Val(II.AllParameters) > 0 Then
                         '   Pause
                         'End If
174                           ISadAddin_ExecuteSoftCommand = True

175                     Case "OPENFORINPUT", "OPENFILE"
176                           fh = FreeFile
177                           Open sVar1 For Input Access Read As #fh

178                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
179                              If fh <> 0 Then Close #fh: fh = 0
180                              ISadAddin_ExecuteSoftCommand = True
181                              Exit Function
182                           End If

183                           ISadAddin_ExecuteSoftCommand = True

184                     Case "OPENFOROUTPUT"
185                           Open sVar1 For Output Access Write As #fh

186                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
187                              If fh <> 0 Then Close #fh: fh = 0
188                              ISadAddin_ExecuteSoftCommand = True
189                              Exit Function
190                           End If

191                           ISadAddin_ExecuteSoftCommand = True

192                     Case "NEXTLINE", "LINEINPUTFROMFILE", "READLINEFROMFILE", "READFILELINE", "FILEREADLINE"
193                           Line Input #fh, sVar3

194                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
195                              If fh <> 0 Then Close #fh: fh = 0
196                              ISadAddin_ExecuteSoftCommand = True
197                              Exit Function
198                           End If

199                           II.SoftVars(sVar1) = sVar3
200                           ISadAddin_ExecuteSoftCommand = True

201                     Case "CLOSEFILE"
202                           If fh <> 0 Then
203                              Close #fh
204                              fh = 0
205                           End If
206                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
207                              ISadAddin_ExecuteSoftCommand = True
208                              Exit Function
209                           End If
210                           ISadAddin_ExecuteSoftCommand = True

211                     Case "COPYFILE", "FILECOPY"
212                           FileCopy II.Expression, II.Result

213                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
214                              If fh <> 0 Then Close #fh: fh = 0
215                              ISadAddin_ExecuteSoftCommand = True
216                              Exit Function
217                           End If

218                           ISadAddin_ExecuteSoftCommand = True

219                     Case "SAVESETTING", "SAVEREGISTRYSETTING", "SETREGKEY"
220                           SaveSetting "Slice and Dice", sVar1, sVar2, sVar3

221                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
222                              If fh <> 0 Then Close #fh: fh = 0
223                              ISadAddin_ExecuteSoftCommand = True
224                              Exit Function
225                           End If

226                           ISadAddin_ExecuteSoftCommand = True

227                     Case "GETSETTING", "GETREGISTRYSETTING", "GETREGKEY"
228                           .SoftVars(sVar1) = GetSetting("Slice and Dice", sVar2, sVar3, sVar4)

229                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
230                              If fh <> 0 Then Close #fh: fh = 0
231                              ISadAddin_ExecuteSoftCommand = True
232                              Exit Function
233                           End If

234                           ISadAddin_ExecuteSoftCommand = True

235                     Case "DELETESETTING", "DELETEREGISTRYSETTING", "DELETEREGKEY", "DELREGKEY"
236                           DeleteSetting "Slice and Dice", sVar1, sVar2

237                           If ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, vbNullString) Then
238                              If fh <> 0 Then Close #fh: fh = 0
239                              ISadAddin_ExecuteSoftCommand = True
240                              Exit Function
241                           End If

242                           ISadAddin_ExecuteSoftCommand = True

'                   Case "GOTO", "GOSUB", "RETURNFROMGOSUB", "RETURN"
'                         ISadAddin_ExecuteSoftCommand = True
'
'                   Case "SELECTCASE", "SWITCHCASE", "SWITCH"
'                         ISadAddin_ExecuteSoftCommand = True
'
'                   Case "CREATETREE", "CREATEDIRECTORYTREE"
'                         ISadAddin_ExecuteSoftCommand = True
'
'                   Case "DELETETREE", "DELETEDIRECTORYTREE"
'                         ISadAddin_ExecuteSoftCommand = True
'
'                   Case "MODIFYTREE", "MODIFYDIRECTORYTREE"
'                         ISadAddin_ExecuteSoftCommand = True
'
'                   Case "FOREACHFILE", "STARTEACHFILE", "ENDEACHFILE"
'                        sResults = vbNullString
'                        ISadAddin_ExecuteSoftCommandInline = True
'
'                   Case "FOREACHDIRECTORY", "STARTEACHDIRECTORY", "ENDEACHDIRECTORY"
'                        sResults = vbNullString
'                        ISadAddin_ExecuteSoftCommandInline = True
'
243              End Select
244           End If
245      End With

246 EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:
247      Exit Function

248 EH_SADAddin_ISadAddin_ExecuteSoftCommand:

249      ErrorsOcurred II, sVar1, sVar2, sVar3, sVar4, vbNullString
250      Resume EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue

251      Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
252  On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommandInline
253      Dim Template                As CTemplate
254      Dim Area                    As Long
255      Dim CurrSet                 As Long
256      Dim bInlineCommandExecuted  As Boolean
257      Dim lParameterCount         As Long
258      Dim sDefault                As String
259      Dim sT                      As String
260      Dim sVar()                  As String
    
261      Dim CurrParam               As Long
262      Dim ParameterCount          As Long

263      ReDim sVar(1 To 5) As String

264      If Not MySadCommands(sInlineSoftCommandName & "*I") Is Nothing Then
265         If Not MySadCommands(sInlineSoftCommandName & "*I").IsInline Then Exit Function

266         ParameterCount = Parent.lTokenCount(sParameters, ",")

267         If ParameterCount > 0 Then
268            If ParameterCount < 5 Then
269               ReDim sVar(1 To 5) As String
270            Else
271               ReDim sVar(1 To ParameterCount) As String
272            End If
273            For CurrParam = 1 To ParameterCount
274                sVar(CurrParam) = Parent.sGetToken(sParameters, CurrParam, ",")
275                If Len(SoftVars(sVar(CurrParam))) Then sVar(CurrParam) = SoftVars(sVar(CurrParam))
276            Next CurrParam
277         End If

       Select Case sInlineSoftCommandName
              Case "CHOOSEDIR", "CHOOSEDIRECTORY", "PICKDIR", "PICKDIRECTORY", "GETPATH", "BROWSEFORFOLDER", "CHOOSEFOLDER", "BROWSEFORDIR", "BROWSEFORDIRECTORY"
                   If Len(sVar(1)) = 0 Then
                      sResults = Parent.BrowseForFolder("Choose a Folder")
                   Else
                      sResults = Parent.BrowseForFolder(sVar(1))
                   End If
                   If Len(sResults) = 0 Then
                      If Parent.bUserSure("No Folder selected." & vbNewLine & vbTab & "Cancel processing ?") Then
                         sResults = vbNewLine & vbNewLine & "~~' Cancel inserted at user's request" & vbNewLine & "~~Cancel 1" & vbNewLine
                      End If
                   End If
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "CHOOSEFILE", "PICKFILE", "GETPATHANDFILE"
                   Select Case ParameterCount
                          Case Is < 2: sResults = Parent.sChooseFile(sVar(1))
278                            Case 2:    sResults = Parent.sChooseFile(sVar(1), sVar(2))
279                            Case Else: sResults = Parent.sChooseFile(sVar(1), sVar(2), sVar(3))
280                     End Select

281                     If Len(sResults) = 0 Then
282                        If Parent.bUserSure("No file selected." & vbNewLine & vbTab & "Cancel processing ?") Then
283                           sResults = vbNewLine & vbNewLine & "~~' Cancel inserted at user's request" & vbNewLine & "~~Cancel 1" & vbNewLine
284                        End If
285                     End If

286                     ISadAddin_ExecuteSoftCommandInline = True

287                Case "CHOOSEDATABASE", "CHOOSEDB", "PICKDATABASE", "PICKDB", "GETDBPATHANDFILE"
                   Select Case ParameterCount
                          Case Is < 2: sResults = Parent.sChooseDatabase(sVar(1))
288                            Case Else: sResults = Parent.sChooseDatabase(sVar(1), sVar(2))
289                     End Select

290                     If Len(sResults) = 0 Then
291                        If Parent.bUserSure("No database selected." & vbNewLine & vbTab & "Cancel processing ?") Then
292                           sResults = vbNewLine & vbNewLine & "~~' Cancel inserted at user's request" & vbNewLine & "~~Cancel 1" & vbNewLine
293                        End If
294                     End If

295                     ISadAddin_ExecuteSoftCommandInline = True

296                Case "CHOOSECOLOR", "PICKCOLOR"
                   Select Case ParameterCount
                          Case Is > 0: sResults = Parent.sChooseColor(sVar(1))
297                            Case Else: sResults = Parent.sChooseColor
298                     End Select

299                     If Len(sResults) = 0 Then
300                        If Parent.bUserSure("No color selected." & vbNewLine & vbTab & "Cancel processing ?") Then
301                           sResults = vbNewLine & vbNewLine & "~~' Cancel inserted at user's request" & vbNewLine & "~~Cancel 1" & vbNewLine
302                        End If
303                     End If

304                     ISadAddin_ExecuteSoftCommandInline = True

305                Case "READTREE", "READDIRECTORYTREE", "TREETOSOFTVAR"
306                     If Parent.bUserSure("The SoftCommand 'READDIRECTORYTREE' has not been implemented yet (sorry)." & vbNewLine & vbNewLine & vbTab & "Cancel processing ?") Then
307                        sResults = vbNewLine & vbNewLine & "~~' Cancel inserted at user's request" & vbNewLine & "~~Cancel 1" & vbNewLine
308                     End If
309                     ISadAddin_ExecuteSoftCommandInline = True

310                Case "FILEEXISTS", "EXISTENCE", "DOESFILEEXIST"  ' inline
311                     If Len(SoftVars(sParameters)) > 0 Then
312                        sResults = Abs(Len(Dir(SoftVars(sParameters))) > 0)
313                     Else
314                        sResults = Abs(Len(Dir(sParameters)) > 0) ' "1" if the file exists, "0" if it doesn't
315                     End If
316                     ISadAddin_ExecuteSoftCommandInline = True

317                Case "MATCHFILE", "DIR", "FINDFIRST", "FIRSTFILE"         'inline
318                     sResults = Dir(sVar(1), Val(sVar(2)))
319                     ISadAddin_ExecuteSoftCommandInline = True

320                Case "MATCHNEXTFILE", "FINDNEXT", "NEXTFILE"             'inline
321                     sResults = Dir
322                     ISadAddin_ExecuteSoftCommandInline = True

323                Case "GETSETTING", "GETREGISTRYSETTING", "GETREGKEY"         'inline
324                     sResults = GetSetting("Slice and Dice", sVar(1), sVar(2), sVar(3))
325                     ISadAddin_ExecuteSoftCommandInline = True

326                Case "FILECONTENTS", "CONTENTS", "INCLUDEFILE"
327                     sResults = Parent.sFileContents(sParameters)
328                     ISadAddin_ExecuteSoftCommandInline = True

329                Case "LISTDIRS", "LISTDIRECTORIES", "DIRECTORIES", "DIRS", "ENUMDIRS"
330                     sResults = Parent.EnumFiles(sVar(1), sVar(2), "DIRECTORY")
331                     ISadAddin_ExecuteSoftCommandInline = True

332                Case "LISTFILES", "FILES", "ENUMFILES"
333                     sResults = Parent.EnumFiles(sVar(1), sVar(2), sVar(3))
334                     ISadAddin_ExecuteSoftCommandInline = True

335                Case "TESTEOF", "EOF", "ENDOFFILE", "ISENDOFFILE", "TESTENDOFFILE"
336                     If fh = 0 Then
337                        sResults = "0"
338                     Else
339                        sResults = EOF(fh) & vbNullString
340                     End If
341                     ISadAddin_ExecuteSoftCommandInline = True
342          End Select
343      End If

344 EH_SADAddin_ISadAddin_ExecuteSoftCommandInline_Continue:
345      Exit Function

346 EH_SADAddin_ISadAddin_ExecuteSoftCommandInline:
347      Dim X As CInsertionInfo

348      Set X = New CInsertionInfo
349      With X
350           .AllParameters = sParameters
351           .Result = sResults
352           Set .SoftVars = SoftVars
353           .SoftCommandName = sInlineSoftCommandName
354           .CurrentLineToProcess = "-Inline Substition-"
355      End With

356      ErrorsOcurred X, sVar(1), sVar(2), sVar(3), sVar(4), sVar(5)

357      Set X.SoftVars = Nothing
358      Set X = Nothing

359      Resume EH_SADAddin_ISadAddin_ExecuteSoftCommandInline_Continue
360      Resume
End Function

Private Property Get ISadAddin_Externals() As SliceAndDice.CAssocArray
'    Dim Externals As CAssocArray
'    Set Externals = New CAssocArray
'    Externals.All = "Testing Externals=Testing 123"
'    Set ISadAddin_Externals = Externals
'    Set Externals = Nothing
End Property

Private Function ISadAddin_Shutdown() As Boolean
361  On Error Resume Next
362      MySadCommands.Clear
363      Set MySadCommands.Parent = Nothing
364      Set MySadCommands = Nothing
365      Set Parent = Nothing

366      ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard, vbInst As VBIDE.VBE) As Boolean
367  On Error Resume Next
368      Err.Clear

'    If UCase$(oParent.Version) <> UCase$(MySadCommands.Attributes("Version")) Then
'       If MsgBox("Version mismatch:" & Chr(13) & Chr(9) & "SADAddin.NewCommands = " & MySadCommands.Attributes("Version") & Chr(13) & Chr(9) & "Slice and Dice = " & oParent.Version & Chr(13) & Chr(13) & "Continue loading DLL ?", vbYesNo, "*** WARNING - DLL Version mismatch **") = vbNo Then
'          Exit Function
'       End If
'    End If

369      Set Parent = oParent
370      Set MySadCommands.Parent = oParent

371      Err.Clear
372          MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath & "sadFile.txt")
373          ISadAddin_Startup = (Err.Number = 0)
374      Err.Clear
End Function

Private Function ErrorsOcurred(ByRef II As CInsertionInfo, ByVal sVar1 As String, ByVal sVar2 As String, ByVal sVar3 As String, ByVal sVar4 As String, ByVal sVar5 As String) As Boolean
375  On Error Resume Next
376      Dim fh          As Long
377      Dim sMessage    As String

378      If Err.Number <> 0 Then
379         sMessage = "Error executing SoftCode:" & vbNewLine
380         sMessage = sMessage & vbTab & "Occured:      " & Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM")
381         sMessage = sMessage & vbTab & "Sandal:       sadFile.NewCommands" & vbNewLine
382         If Erl <> 0 Then sMessage = sMessage & vbTab & "Sandal Line:  " & Erl & vbNewLine
383         sMessage = sMessage & vbTab & "Error Number: " & Err.Number & vbNewLine
384         sMessage = sMessage & vbTab & "Description:  " & Err.Description & vbNewLine

385         sMessage = sMessage & vbNewLine & vbTab & "SoftCode Parameters (Resolved):" & vbNewLine
386         sMessage = sMessage & vbTab & vbTab & "sVar(1) = " & sVar1 & vbNewLine
387         sMessage = sMessage & vbTab & vbTab & "sVar(2) = " & sVar2 & vbNewLine
388         sMessage = sMessage & vbTab & vbTab & "sVar(3) = " & sVar3 & vbNewLine
389         sMessage = sMessage & vbTab & vbTab & "sVar(4) = " & sVar4 & vbNewLine
390         sMessage = sMessage & vbTab & vbTab & "sVar(5) = " & sVar5 & vbNewLine

391         If Not II Is Nothing Then
392            sMessage = sMessage & vbNewLine & vbTab & "(Unresolved) CInsertionInfo contents:" & vbNewLine
393            sMessage = sMessage & vbTab & vbTab & "TemplateName = " & II.TemplateName & vbNewLine
394            sMessage = sMessage & vbTab & vbTab & "CurrentLineToProcess = " & II.CurrentLineToProcess & vbNewLine
395            sMessage = sMessage & vbTab & vbTab & "SoftCommandName = " & II.SoftCommandName & vbNewLine
396            sMessage = sMessage & vbTab & vbTab & "AllParameters = " & II.AllParameters & vbNewLine
397            sMessage = sMessage & vbTab & vbTab & "Result     (LHS) = " & II.Result & vbNewLine
398            sMessage = sMessage & vbTab & vbTab & "Expression (RHS) = " & II.Expression & vbNewLine
399            sMessage = sMessage & vbTab & vbTab & "ExternalFilename = " & II.ExternalFilename & vbNewLine
400         End If

401         sMessage = sMessage & vbNewLine & vbNewLine & "Would you like to cancel processing ?" & vbNewLine
402         sMessage = sMessage & vbNewLine & vbTab & "IF YOU SELECT:" & vbNewLine
403         sMessage = sMessage & vbTab & vbTab & "YES: This template should be cancelled."
404         sMessage = sMessage & vbTab & vbTab & vbTab & "NOTE: Due to the nature of this addin, processing is not guarenteed to stop."
405         sMessage = sMessage & vbTab & vbTab & "NO : Processing will continue with the next line of SoftCode."
406         sMessage = sMessage & vbTab & "NOTE: This information will be stored in: """ & App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "\sadFileError.Log"""

407         If MsgBox(sMessage, vbYesNo, "CANCEL PROCESSING ?") = vbYes Then
408            If Not II Is Nothing Then
409               II.LinesLeftToProcess = vbNullString
410            End If
411            sMessage = sMessage & vbNewLine & "  *** User decided to CANCEL processing."
412            ErrorsOcurred = True ' Tell parent function processing has been cancelled.
413         Else
414            sMessage = sMessage & vbNewLine & "  *** User choose to CONTINUE after error."
415            Err.Clear
416         End If

417         fh = FreeFile
418         Open App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "\sadFileError.Log" For Append As #fh
419              Print #fh, sMessage
420         Close #fh
421      End If
End Function

