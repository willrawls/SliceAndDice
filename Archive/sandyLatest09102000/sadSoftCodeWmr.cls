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

Private Parent              As SliceAndDice.Wizard
Private CommandsSupported   As SliceAndDice.CSadCommands

Private ItemLists           As Collection
Private ItemsToProcess      As Variant

Public Function sGinsu(ByVal xmlFields As String, ByVal xmlValues As String, ByVal xmlWrapper As String, Optional ByVal FieldDelimiter As String = ";", Optional ByVal ValueDelimiter As String = vbNewLine) As String
On Error Resume Next
    Dim sResult       As String
    Dim FieldList     As Variant
    Dim ValueList     As Variant
    Dim CurrField     As Variant
    Dim CurrValue     As Long
    Dim MaxValue      As Long

    If Len(xmlWrapper) Then sResult = "<" & xmlWrapper & ">" & vbNewLine

    If Len(FieldDelimiter) = 0 Then FieldDelimiter = ";"
    If Len(ValueDelimiter) = 0 Then ValueDelimiter = vbNewLine
    
    FieldList = Split(xmlFields, FieldDelimiter)
    ValueList = Split(xmlValues, ValueDelimiter)

    CurrValue = -1
    MaxValue = UBound(ValueList)

    Do While CurrValue < MaxValue
       For Each CurrField In FieldList
           CurrValue = CurrValue + 1
           If CurrValue <= MaxValue Then
              sResult = sResult & vbTab & "<" & CurrField & ">" & ValueList(CurrValue) & "</" & CurrField & ">" & vbNewLine
           Else
              sResult = sResult & vbTab & "<" & CurrField & "></" & CurrField & ">" & vbNewLine
           End If
       Next CurrField
       sResult = sResult & "</" & xmlWrapper & ">" & vbNewLine
       If CurrValue < MaxValue Then sResult = sResult & "<" & xmlWrapper & ">" & vbNewLine
    Loop

    sGinsu = sResult
End Function


Private Sub Class_Terminate()
On Error Resume Next
    Set CommandsSupported = Nothing
    Set ItemLists = Nothing
End Sub

Private Property Get ISadAddin_CommandSet() As SliceAndDice.CSadCommands
On Error Resume Next
    Set ISadAddin_CommandSet = CommandsSupported
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
On Error Resume Next
    Dim Commands As CSadCommands
    Dim CurrCommand As CSadCommand
    Dim sHTML As String

    Select Case UCase$(sValue)
           Case "REVISIONHISTORY"
                If Len(Dir$(App.Path & "\SandyRevisions.txt")) = 0 Then
                   MsgBox "Revision history file ('SandyRevisions.txt') not in the application path." & vbNewLine & vbTab & "Unable to view at this time."
                Else
                   Parent.ShowMessage Parent.sFileContents(App.Path & "\SandyRevisions.txt"), "Slice and Dice Revision History", "Abridged history of changes made between revisions of Slice and Dice."
                End If

           Case "HTMLREFERENCE"
                Screen.MousePointer = vbHourglass
                Set Commands = Parent.SoftCommands
                sHTML = "<HTML><BODY><TABLE>"
                For Each CurrCommand In Commands
                    With CurrCommand
                         sHTML = sHTML & vbTab & "<TR>" & vbNewLine
                         sHTML = sHTML & vbTab & vbTab & "<TD><B><H2>" & .SoftCommandName & "</h2></B><UL><BR>" & vbNewLine
                         If Len(.Aliases) > 0 Then sHTML = sHTML & vbTab & vbTab & "<LI><B><H3>Aliases</H3></B> " & "<BLOCKQUOTE><H3>" & Left$(Mid$(.Aliases, 3), Len(.Aliases) - 4) & "</H3></BLOCKQUOTE><BR>" & vbNewLine
                         If Len(.OneLineDescription) > 0 Then sHTML = sHTML & vbTab & vbTab & "<LI><B><H3>Summary</H3></B> " & "<BLOCKQUOTE>" & .OneLineDescription & "</BLOCKQUOTE><BR>" & vbNewLine
                         If Len(.SeeAlso) > 0 Then sHTML = sHTML & vbTab & vbTab & "<LI><B><H3>See Also</H3></B> " & "<BLOCKQUOTE>" & Replace(.SeeAlso, vbNewLine, "<BR>" & vbNewLine) & "</BLOCKQUOTE><BR>" & vbNewLine
                         If Len(.LongDescription) > 0 Then sHTML = sHTML & vbTab & vbTab & "<LI><B><H3>Long Description</H3></B> " & "<BLOCKQUOTE>" & Replace(.LongDescription, vbNewLine, "<BR>" & vbNewLine) & "</BLOCKQUOTE><BR>" & vbNewLine
                         If Len(.Comments) > 0 Then sHTML = sHTML & vbTab & vbTab & "<LI><B><H3>Comments</H3></B> " & "<BLOCKQUOTE>" & Replace(.Comments, vbNewLine, "<BR>" & vbNewLine) & "</BLOCKQUOTE><BR>" & vbNewLine
                         If Len(.Examples) > 0 Then sHTML = sHTML & vbTab & vbTab & "<LI><B><H3>Examples</H3></B> " & "<BLOCKQUOTE>" & Replace(.Examples, vbNewLine, "<BR>" & vbNewLine) & "</BLOCKQUOTE><BR>" & vbNewLine
                         sHTML = sHTML & vbTab & "</UL></TD></TR>" & vbNewLine
                    End With
                Next CurrCommand
                sHTML = sHTML & vbNewLine & "</TABLE></BODY></HTML>" & vbNewLine
                Parent.SaveToFile Parent.TemplateDatabasePath & "sadReference.html", sHTML
                sHTML = vbNullString
                Parent.BrowseTo Parent.TemplateDatabasePath & "sadReference.html"
                ISadAddin_ExecuteExternal = True
                Screen.MousePointer = vbDefault
         End Select
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
On Error GoTo SandalError_ExecuteSoftCommand
         Dim vParams(0 To 3) As Variant
         Dim CurrItem        As Long
         Dim CurrSection     As Long
         Dim EditedMessage   As String
         Dim ktEval          As Evaluator

11       With II
12            If Not CommandsSupported(.SoftCommandName & "*C") Is Nothing Then

13               If CommandsSupported(.SoftCommandName & "*C").IsInline Then Exit Function

14               vParams(0) = Split(II.AllParameters, ",")
15               vParams(1) = Array(Parent.sGetToken(II.AllParameters, 1, " - "), Parent.sAfter(II.AllParameters, 1, " - "))
16               vParams(2) = vParams(0)
17               vParams(3) = vParams(1)

18               For CurrSection = 0 To 1
19                   For CurrItem = 0 To UBound(vParams(CurrSection))
20                       If Len(.SoftVars(vParams(CurrSection)(CurrItem) & vbNullString)) Then
21                          vParams(2)(CurrItem) = vParams(CurrSection)(CurrItem)
22                       End If
23                   Next CurrItem
24               Next CurrSection

25   On Error Resume Next
26               Err.Clear
            Select Case UCase$(.SoftCommandName)
                   Case "GINSU", "DOUBLEXML"
                        .SoftVars(II.AllParameters) = sGinsu(II.SoftVars("xmlFields"), II.SoftVars("xmlValues"), II.SoftVars("xmlWrapper"), II.SoftVars("FieldDelimiter"), II.SoftVars("ValueDelimiter"))
                        ISadAddin_ExecuteSoftCommand = True

                   Case "KTEVAL"
                         Set ktEval = New Evaluator
                         If Not ktEval Is Nothing Then
                            .SoftVars(II.Result) = ktEval.Evaluate(II.Expression, True) & vbNullString
                         End If
                         Set ktEval = Nothing
                         ISadAddin_ExecuteSoftCommand = True

                   Case "PARSE", "SPLIT", "TOKENIZE"
     On Error Resume Next
                           If Len(II.Expression) > 0 Then
                              If ItemLists Is Nothing Then Set ItemLists = New Collection
                              If Not ItemLists(II.Result) Is Nothing Then ItemLists.Remove II.Result

                              vParams(2)(0) = II.Expression
                              If Len(.SoftVars(vParams(2)(0) & vbNullString)) > 0 Then
                                 vParams(2)(0) = .SoftVars(vParams(2)(0) & vbNullString)
                              End If
                              ItemLists.Add Split(vParams(2)(0), II.SoftVars("ItemDelimiter")), II.Result

                              II.SoftVars("Next Item in " & II.Result) = "1"
                              II.SoftVars("Current Item in " & II.Result) = "0"
                              II.SoftVars("Next Item") = "1"
                              II.SoftVars("Current Item") = "0"
                              II.SoftVars("ItemCount") = UBound(ItemLists(II.Result)) + 1
                              II.SoftVars("Current ItemList") = II.Result
                           End If
                           ISadAddin_ExecuteSoftCommand = True

                   Case "EXECUTESOFTCODE", "RUNSOFTCODE"
27   On Error Resume Next
28                            Parent.InsertTemplate II.AllParameters
29                            ISadAddin_ExecuteSoftCommand = True

30                      Case "GOTOSANDYTEMPLATE"
31   On Error Resume Next
32                            If Parent.SliceAndDice.Categorys(vParams(2)(0)) Is Nothing Then
33                               Parent.NewTemplate True, II.AllParameters, vParams(2)(1), True
34                            ElseIf Parent.SliceAndDice.Categorys(vParams(2)(0)).Templates(vParams(2)(1)) Is Nothing Then
35                               Parent.NewTemplate True, II.AllParameters, vParams(2)(1), True
36                            Else
37                               Parent.JumpTo II.AllParameters, False, False
38                            End If
39                            ISadAddin_ExecuteSoftCommand = True

40                      Case "SETCODEAREATEXT", "SETCODEAREA"
                         Select Case UCase$(vParams(0)(0))
                                Case "ATTOP", "AT TOP", "TOP", "(DECLARATIONS)", "DECLARATIONS", "DECL"
41                                        vParams(0)(0) = "0"
42                                        vParams(2)(0) = "0"
43                                   Case "ATCURSOR", "AT CURSOR", "CURSOR"
44                                        vParams(0)(0) = "1"
45                                        vParams(2)(0) = "1"
46                                   Case "ATBOTTOM", "AT BOTTOM", "BOTTOM"
47                                        vParams(0)(0) = "2"
48                                        vParams(2)(0) = "2"
49                                   Case "INFILE", "IN FILE", "FILE"
50                                        vParams(0)(0) = "3"
51                                        vParams(2)(0) = "3"
52                            End Select
                         
53                            If Parent.InternalCurrentTemplate Is Nothing Then
                            Select Case UCase$(vParams(0)(0))
                                   Case "0": Parent.InternalCurrentTemplate.memoCodeAtTop = vParams(2)(1)
54                                      Case "1": Parent.InternalCurrentTemplate.memoCodeAtCursor = vParams(2)(1)
55                                      Case "2": Parent.InternalCurrentTemplate.memoCodeAtBottom = vParams(2)(1)
56                                      Case "3": Parent.InternalCurrentTemplate.memoCodeToFile = vParams(2)(1)
57                               End Select
58                               Parent.SliceAndDice.Save Parent.TemplateDatabasePath & Parent.TemplateDatabaseName
59                            End If

60                            ISadAddin_ExecuteSoftCommand = True
61               End Select
62            End If
63       End With

64 SandalError_ExecuteSoftCommand_Continue:
65       Exit Function

66 SandalError_ExecuteSoftCommand:

67       ErrorsOcurred II, vParams(2)(0), vParams(2)(1), vbNullString, vbNullString, vbNullString
68       Resume SandalError_ExecuteSoftCommand_Continue

69       Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
On Error GoTo SandalError_ExecuteSoftCommandInline
    Dim vParams As Variant

    If Not CommandsSupported(sInlineSoftCommandName & "*I") Is Nothing Then
       If Not CommandsSupported(sInlineSoftCommandName & "*I").IsInline Then Exit Function

       Select Case sInlineSoftCommandName
              Case "LASTTOKEN", "TLAST"
                   If Len(sParameters) Then
                      vParams = Split(sParameters, ",")
                      If Len(SoftVars.Item(vParams(0))) Then vParams(0) = SoftVars(vParams(0))
                      If UBound(vParams) = 0 Then
                       ' Only a string or soft variable was passed. Use the default delimiter ("\")
                         sResults = Parent.sGetToken(vParams(0), Parent.lTokenCount(vParams(0), "\"), "\")
                      Else
                         If Len(SoftVars(vParams(1))) Then vParams(1) = SoftVars(vParams(1))
                         sResults = Parent.sGetToken(vParams(0), Parent.lTokenCount(vParams(0), vParams(1)), vParams(1))
                      End If
                   Else
                      sResults = vbNullString
                   End If
                   ISadAddin_ExecuteSoftCommandInline = True
                
              Case "BEFORELASTTOKEN", "BEFORELAST", "EXCEPTLASTTOKEN", "EXCEPTLAST"
                   If Len(sParameters) Then
                      vParams = Split(sParameters, ",")
                      If Len(SoftVars(vParams(0))) Then vParams(0) = SoftVars(vParams(0))
                      If UBound(vParams) = 0 Then
                       ' Only a string or soft variable was passed. Use the default delimiter ("\")
                         sResults = Parent.sBefore(vParams(0), Parent.lTokenCount(vParams(0), "\"), "\")
                      Else
                         If Len(SoftVars(vParams(1))) Then vParams(1) = SoftVars(vParams(1))
                         sResults = Parent.sBefore(vParams(0), Parent.lTokenCount(vParams(0), vParams(1)), vParams(1))
                      End If
                   Else
                      sResults = vbNullString
                   End If
                   ISadAddin_ExecuteSoftCommandInline = True
              
              Case "NEXTITEM", "ITEMTOPROCESS", "NEXTITEMTOPROCESS", "NEXTTOKEN"
On Error Resume Next
                   If Len(sParameters) = 0 Then sParameters = SoftVars("Current ItemList")
                   If Len(sParameters) > 0 Then
                      If Not ItemLists(sParameters) Is Nothing Then
                         With SoftVars("Next Item in " & sParameters)
                         
                              If Val(.Value) > 1 Then
                                 SoftVars("Previous Value") = ItemLists(sParameters)(Val(.Value) - 2)
                              Else
                                 SoftVars("Previous Value") = vbNullString
                              End If
                              If Val(.Value) > 0 Then
                                 SoftVars("Current Value") = ItemLists(sParameters)(Val(.Value) - 1)
                              Else
                                 SoftVars("Current Value") = vbNullString
                              End If
                         
                              sResults = ItemLists(sParameters)(Val(.Value) - 1)
                              SoftVars("Current Item in " & sParameters) = .Value
                              SoftVars("Previous Item") = SoftVars("Current Item")
                              SoftVars("Current Item") = .Value
                                .Value = .Value + 1
                                If Val(.Value) > UBound(ItemLists(sParameters)) + 1 Then
                                   .Value = "0"
                                   SoftVars("Next Value") = vbNullString
                                Else
                                   SoftVars("Next Value") = ItemLists(sParameters)(Val(.Value) - 1)
                                End If
                              SoftVars("Next Item") = .Value
                         End With
                      End If
                   End If
                   ISadAddin_ExecuteSoftCommandInline = True
        
              Case "GETFILELIST", "LISTFILES"
                   Select Case Parent.lTokenCount(sParameters, ",")
                          Case 3:   vParams = Split(sParameters, ",")
                          Case 2:   vParams = Split(sParameters & ", ", ",")
                          Case 1:   vParams = Split(sParameters & ",*.*, ", ",")
                   End Select

                   If Len(SoftVars(vParams(0) & vbNullString)) Then vParams(0) = SoftVars(vParams(0) & vbNullString)
                   If Len(SoftVars(vParams(1) & vbNullString)) Then vParams(1) = SoftVars(vParams(1) & vbNullString)
                   If Len(SoftVars(vParams(2) & vbNullString)) Then vParams(2) = SoftVars(vParams(2) & vbNullString)

                   sResults = GetFileList(vParams(0), vParams(1), vParams(2))

                   ISadAddin_ExecuteSoftCommandInline = True
        End Select
99       End If

100 SandalError_ExecuteSoftCommandInline_Continue:
101      Exit Function

102 SandalError_ExecuteSoftCommandInline:
103      Dim X As CInsertionInfo

104      Set X = New CInsertionInfo
105      With X
106           .AllParameters = sParameters
107           .Result = sResults
108           Set SoftVars = SoftVars
109           .SoftCommandName = sInlineSoftCommandName
110           .CurrentLineToProcess = "-Inline Substition-"
111      End With

    'ErrorsOcurred X, sVar(1), sVar(2), sVar(3), sVar(4), sVar(5)

    'Set XSoftVars = Nothing
112      Set X = Nothing

113      Resume SandalError_ExecuteSoftCommandInline_Continue
114      Resume
End Function

Private Property Get ISadAddin_Externals() As SliceAndDice.CAssocArray
On Error Resume Next
    Dim Externals As CAssocArray
    Set Externals = New CAssocArray
        With Externals
             .Clear True
             .Item("SoftCommand HTML Reference") = "HTMLREFERENCE"
             .Item("Revision History") = "REVISIONHISTORY"
        End With
        Set ISadAddin_Externals = Externals
    Set Externals = Nothing
End Property

Private Function ISadAddin_Shutdown() As Boolean
On Error Resume Next
116      CommandsSupported.Clear
117      Set CommandsSupported.Parent = Nothing
118      Set CommandsSupported = Nothing
119      Set Parent = Nothing

120      ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard, vbInst As VBIDE.VBE) As Boolean
On Error Resume Next
122      Err.Clear

123      Set Parent = oParent
    
124  On Error Resume Next
        'Set asaList = New CAssocArray

125      Set CommandsSupported = New SliceAndDice.CSadCommands
126      With CommandsSupported
127           Set .Parent = oParent
128           .ParameterDelimiter = ","
129           .ParameterTypeDelimiter = ":"
130           .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision

              With .Add("LastToken", True)
                   .Aliases = ", tLast, "
                   .Examples = "~~Set X=1;2;3" & vbNewLine & "~~Set Y=%%LastToken::X,;%%" & vbNewLine & "~~' Y is now equal to '3'"
                   .Syntax = "[TheTokens As SoftVarOrString],[Opt TokenDelimiter As SoftVarOrString = ""\""]"
                   .OneLineDescription = "Returns the last token in a string or SoftVariable"
                   .SeeAlso = "BeforeLastToken, FirstToken, Token, TokenCount, Before, After"
              End With
              
              With .Add("BeforeLastToken", True)
                   .Aliases = ", BeforeLast, ExceptLastToken, ExceptLast, "
                   .Examples = "~~Set X=1;2;3" & vbNewLine & "~~Set Y=%%BeforeLastToken::X,;%%" & vbNewLine & "~~' Y is now equal to '3'"
                   .Syntax = "[TheTokens As SoftVarOrString],[Opt TokenDelimiter As SoftVarOrString = "" ""]"
                   .OneLineDescription = "Returns the last token in a string or SoftVariable"
                   .SeeAlso = "BeforeLastToken, Token, TokenCount, Before, After"
              End With
              
              With .Add("Ginsu")
                   .Syntax = "Special"
                   .OneLineDescription = "Wraps Values with 2 layers of XML tags"
                   .Examples = "~~GotoMessage" & vbNewLine & _
                               "~~Set xmlFields=Field1;Field2;Field3" & vbNewLine & _
                               "~~Set xmlValues=1;2;3;4;5;6" & vbNewLine & _
                               "~~Set xmlWrapper=OuterLayer" & vbNewLine & _
                               "%%Ginsu%%"
              End With
              
              With .Add("GetFileList", True)
                   .Syntax = "[StartingPath As StringOrSoftVar],[Opt FileMask = ""*.*""],[Opt FileDelimiter = "" ""]"
                   .OneLineDescription = "Returns a delimited list of files fitting the mask."
              End With

              With .Add("KTEval", False)
                 .Syntax = "[Result As SoftVar]=[Expression As SoftVarOrString]"
                 .OneLineDescription = "Significantly more advanced Eval command"
                 .LongDescription = "'Copyright© 1999, Tretyakov Konstantin" & vbNewLine & _
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
              End With

131           With .Add("ExecuteSoftCode", True)
132               .Aliases = ", RunSoftCode, "
133               .Examples = "~~GetClipboardText cbText" & vbNewLine & "~~ExecuteSoftCode cbText"
134               .Syntax = "[CodeToExecute As SoftCode]"
135           End With

136           With .Add("GotoSandyTemplate")
137               .Examples = "~~GotoSandyTemplate Move, modify, and store - Think SoftCode 1"
138               .Syntax = "[CategoryName & "" - "" & TemplateName]"
139               .Comments = "Same as clicking on a template, does NOT affect current insertion."
140           End With

141           With .Add("SetCodeArea")
142               .Aliases = ", SetCodeAreaText, "
143               .Examples = "~~GotoSandyTemplate Move, modify, and store - Think SoftCode 1" & vbNewLine & "~~Set example=This will get inserted into the template" & vbNewLine & "~~SetCodeArea 0 example"
144               .Syntax = "[CodeToExecute As SoftCode]"
145           End With

              With .Add("Parse")
                   .Aliases = ", Split, Tokenize, "
                   .Examples = "~~Set cbText=A token;Another token;Now a third one" & vbNewLine & _
                               "~~Set ItemDelimiter=;" & vbNewLine & _
                               "~~Parse Test List=cbText" & vbNewLine & _
                               "~~_Outer_StartCoda" & vbNewLine & _
                               "  ~~Set cItem=%%NextItem::Test List%%" & vbNewLine & _
                               "' %%Current Item in Test List%% : ""%%cItem%%""" & vbNewLine & _
                               "  ~~String ItemDelimiter 1 "" """ & vbNewLine & _
                               "  ~~Parse Inner List=cItem" & vbNewLine & _
                               "  ~~_Inner_StartCoda" & vbNewLine & _
                               "    ~~Set cWord=%%NextItem::Inner List%%" & vbNewLine & _
                               "'   %%Current Item in Inner List%% : ""%%cWord%%""" & vbNewLine & _
                               "  ~~_Inner_Coda %%Next Item in Inner List%%" & vbNewLine & _
                               "~~_Outer_Coda %%Next Item in Test List%%" & vbNewLine
                        
                   .Syntax = "[ListName As String] = [SoftVarToParse As SoftVar]"
              End With
              With .Add("NextToken")
                   .IsInline = True
                   .Aliases = ", NextItem, ItemToProcess, NextItemToProcess, "
                   .Examples = "~~Set cbText=A token;Another token;Now a third one" & vbNewLine & _
                               "~~Set ItemDelimiter=;" & vbNewLine & _
                               "~~Parse Test List=cbText" & vbNewLine & _
                               "~~_Outer_StartCoda" & vbNewLine & _
                               "  ~~Set cItem=%%NextItem::Test List%%" & vbNewLine & _
                               "' %%Current Item in Test List%% : ""%%cItem%%""" & vbNewLine & _
                               "  ~~String ItemDelimiter 1 "" """ & vbNewLine & _
                               "  ~~Parse Inner List=cItem" & vbNewLine & _
                               "  ~~_Inner_StartCoda" & vbNewLine & _
                               "    ~~Set cWord=%%NextItem::Inner List%%" & vbNewLine & _
                               "'   %%Current Item in Inner List%% : ""%%cWord%%""" & vbNewLine & _
                               "  ~~_Inner_Coda %%Next Item in Inner List%%" & vbNewLine & _
                               "~~_Outer_Coda %%Next Item in Test List%%" & vbNewLine
                   .Syntax = "[Opt ListName As String = LastListAccessed]"
                   .Comments = "[ListName] will default to the last ListName parsed or from whom 'NextToken' was last called."
              End With
146       End With
    

147      Err.Clear

148          ISadAddin_Startup = (Err.Number = 0)
149      Err.Clear
End Function

Private Function ErrorsOcurred(ByRef II As CInsertionInfo, ByVal sVar1 As String, ByVal sVar2 As String, ByVal sVar3 As String, ByVal sVar4 As String, ByVal sVar5 As String) As Boolean
On Error Resume Next
151      Dim fh          As Long
152      Dim sMessage    As String

153      If Err.Number <> 0 Then
154         sMessage = "Error executing SoftCode:" & vbNewLine
155         sMessage = sMessage & vbTab & "Occured:      " & Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM")
156         sMessage = sMessage & vbTab & "Sandal:       sadSoftCodeWmr.NewCommands" & vbNewLine
157         If Erl <> 0 Then sMessage = sMessage & vbTab & "Sandal Line:  " & Erl & vbNewLine
158         sMessage = sMessage & vbTab & "Error Number: " & Err.Number & vbNewLine
159         sMessage = sMessage & vbTab & "Description:  " & Err.Description & vbNewLine

160         sMessage = sMessage & vbNewLine & vbTab & "SoftCode Parameters (Resolved):" & vbNewLine
161         sMessage = sMessage & vbTab & vbTab & "sVar(1) = " & sVar1 & vbNewLine
162         sMessage = sMessage & vbTab & vbTab & "sVar(2) = " & sVar2 & vbNewLine
163         sMessage = sMessage & vbTab & vbTab & "sVar(3) = " & sVar3 & vbNewLine
164         sMessage = sMessage & vbTab & vbTab & "sVar(4) = " & sVar4 & vbNewLine
165         sMessage = sMessage & vbTab & vbTab & "sVar(5) = " & sVar5 & vbNewLine

166         If Not II Is Nothing Then
167            sMessage = sMessage & vbNewLine & vbTab & "(Unresolved) CInsertionInfo contents:" & vbNewLine
168            sMessage = sMessage & vbTab & vbTab & "TemplateName = " & II.TemplateName & vbNewLine
169            sMessage = sMessage & vbTab & vbTab & "CurrentLineToProcess = " & II.CurrentLineToProcess & vbNewLine
170            sMessage = sMessage & vbTab & vbTab & "SoftCommandName = " & II.SoftCommandName & vbNewLine
171            sMessage = sMessage & vbTab & vbTab & "AllParameters = " & II.AllParameters & vbNewLine
172            sMessage = sMessage & vbTab & vbTab & "Result     (LHS) = " & II.Result & vbNewLine
173            sMessage = sMessage & vbTab & vbTab & "Expression (RHS) = " & II.Expression & vbNewLine
174            sMessage = sMessage & vbTab & vbTab & "ExternalFilename = " & II.ExternalFilename & vbNewLine
175         End If

176         sMessage = sMessage & vbNewLine & vbNewLine & "Would you like to cancel processing ?" & vbNewLine
177         sMessage = sMessage & vbNewLine & vbTab & "IF YOU SELECT:" & vbNewLine
178         sMessage = sMessage & vbTab & vbTab & "YES: This template should be cancelled."
179         sMessage = sMessage & vbTab & vbTab & vbTab & "NOTE: Due to the nature of this addin, processing is not guarenteed to stop."
180         sMessage = sMessage & vbTab & vbTab & "NO : Processing will continue with the next line of SoftCode."
181         sMessage = sMessage & vbTab & "NOTE: This information will be stored in: """ & App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "\sadSoftCodeWmrError.Log"""

182         If MsgBox(sMessage, vbYesNo, "CANCEL PROCESSING ?") = vbYes Then
183            If Not II Is Nothing Then
184               II.LinesLeftToProcess = vbNullString
185            End If
186            sMessage = sMessage & vbNewLine & "  *** User decided to CANCEL processing."
187            ErrorsOcurred = True ' Tell parent function processing has been cancelled.
188         Else
189            sMessage = sMessage & vbNewLine & "  *** User choose to CONTINUE after error."
190            Err.Clear
191         End If

192         fh = FreeFile
193         Open App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "\sadSoftCodeWmrError.Log" For Append As #fh
194              Print #fh, sMessage
195         Close #fh
196      End If
End Function

