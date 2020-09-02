VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSadCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy - Soft Command Collection"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CSadCommand"
Attribute VB_Ext_KEY = "Member0" ,"CSadCommand"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public ParameterDelimiter As String
Public ParameterTypeDelimiter As String


Public Parent As SliceAndDice.Wizard
Private MyHelpScreen As frmCommandHelp

Private masaParams As CAssocArray
Private masaSyntax As CAssocArray
Private masaAttributes As CAssocArray

Public Enum SadCommandSectionType
    InNone = 0
    InSyntax
    InLong
    InComments
    InSeeAlso
    InExample
End Enum

Public Sub Append(NewItem As Object)
1        On Error Resume Next
2        If TypeOf NewItem Is CSadCommand Then
             NewItem.Index = mCol.Count + 1
3            mCol.Add NewItem, UCase$(NewItem.SoftCommandName) & IIf(NewItem.IsInline, "*I", "*C")
4        End If
End Sub


Public Property Get Attributes() As CAssocArray
5        Set Attributes = masaAttributes
End Property

Public Function CheckSyntax(SoftCommandName As String) As Boolean
6        On Error GoTo EH_CSadCommands_DetermineIfSyntaxIsValid
7        Static CurrSyntaxItem As CAssocItem
8        Static RequiredParams As Long
    'Dim CurrParamItem As CAssocItem
    'Dim ParamType As String

9        If Not Item(SoftCommandName) Is Nothing Then
10           With Item(SoftCommandName)
11               masaSyntax.OrderedMultiAdd .Syntax
12               RequiredParams = 0
13               For Each CurrSyntaxItem In masaSyntax.mCol
14                   If InStr(UCase$(sGetToken(CurrSyntaxItem.Value, 2, ParameterTypeDelimiter)), "OPT") = 0 Then
15                       RequiredParams = RequiredParams + 1
16                   End If
17               Next CurrSyntaxItem
18               masaParams.OrderedMultiAdd .Parameters
19               CheckSyntax = (masaParams.Count >= RequiredParams)

            'mbSyntaxIsValid = True
            'For Each CurrSyntaxItem In masaSyntax.mCol
            'Set CurrParamItem = masaParams(CurrSyntaxItem.Key)
            'ParamType = sAfter(CurrSyntaxItem.Value, 1, Parent.ParameterTypeDelimiter)
            'Select Case UCase$(sGetToken(ParamType))
            '       Case "SOFTVAR"
            '       Case "STRINGORSOFTVAR"
            '       Case "STRING"
            '       Case "QUOTEDSTRING"
            '       Case "LONG"
            '       Case vbNullString
            'End Select
            'Next CurrSyntaxItem
20           End With
21       End If

22 EH_CSadCommands_DetermineIfSyntaxIsValid_Continue:
23       Set CurrSyntaxItem = Nothing
24       Exit Function

25 EH_CSadCommands_DetermineIfSyntaxIsValid:
26       LogError "CSadCommands", "DetermineIfSyntaxIsValid", Err.Number, Err.Description, Erl
27       Resume EH_CSadCommands_DetermineIfSyntaxIsValid_Continue

28       Resume
End Function


Public Function Add(SoftCommandName As String, Optional bInLine As Boolean = False) As CSadCommand
29       On Error GoTo EH_CSadCommands_Add
30       Dim sKey           As String
31       Dim CurrMember     As CSadCommand
32       Dim objNewMember   As CSadCommand

33       Set objNewMember = New CSadCommand

34       With objNewMember
35           .SoftCommandName = SoftCommandName
36           .IsInline = bInLine
37       End With

    ' Fit key into current list in alphabetic order
38       sKey = UCase$(SoftCommandName) & IIf(bInLine, "*I", "*C")
39       If mCol.Count = 0 Then
40           mCol.Add objNewMember, sKey
41       Else
42           For Each CurrMember In mCol
43               If StrComp(sKey, UCase$(CurrMember.SoftCommandName) & IIf(CurrMember.IsInline, "*I", "*C")) < 0 Then
44                   mCol.Add objNewMember, sKey, UCase$(CurrMember.SoftCommandName) & IIf(CurrMember.IsInline, "*I", "*C")
45                   Set Add = objNewMember
46                   GoTo EH_CSadCommands_Add_Continue
47               End If
48           Next CurrMember

        ' Not found, Insert at the end of the collection
49           mCol.Add objNewMember, sKey
50       End If
51       Set Add = objNewMember

52 EH_CSadCommands_Add_Continue:
53       Set objNewMember = Nothing
54       Exit Function

55 EH_CSadCommands_Add:
56       LogError "CSadCommands", "Add", Err.Number, "SoftCommandName=" & SoftCommandName & gsEolTab & Err.Description, Erl
57       Resume EH_CSadCommands_Add_Continue

58       Resume
End Function

Public Property Let All(ByVal NewData As String)
59       On Error Resume Next
60       Dim CurrLine As String
61       Dim CurrSection As SadCommandSectionType
62       Dim CurrCommand As CSadCommand
63       Dim CurrAlias As CSadCommand
64       Dim CurrIndex As Long

65       Clear
66       Do While Len(NewData)
67           CurrLine = sExtractToken(NewData, 1, vbNewLine)
68           If Left$(CurrLine, 3) = gsSpecialLineItemDelimiter & gsS Then
69               CurrLine = Mid$(CurrLine, 4)
            Select Case UCase$(Trim$(sGetToken(CurrLine)))
                Case gsA, "REM"                       ' Comment
70                       CurrSection = InNone

71                   Case "COMMAND", "SOFTCOMMAND"
72                       Set CurrCommand = Add(Trim$(sAfter(CurrLine)), False)
73                       CurrSection = InNone

74                   Case "INLINE", "INLINECOMMAND", "INLINESOFTCOMMAND"
75                       Set CurrCommand = Add(Trim$(sAfter(CurrLine)), True)
76                       CurrSection = InNone

77                   Case "ALIAS", "ALIASES"
78                       If Not CurrCommand Is Nothing Then CurrCommand.Aliases = sAfter(CurrLine)
79                       If Len(CurrCommand.Aliases) Then
80                           CurrCommand.Aliases = Replace(Replace(CurrCommand.Aliases, gsC, ", "), ",  ", ", ")
81                           If Right$(CurrCommand.Aliases, 2) <> ", " Then
82                               CurrCommand.Aliases = CurrCommand.Aliases & ", "
83                           End If
84                           If Left$(CurrCommand.Aliases, 2) <> ", " Then
85                               CurrCommand.Aliases = ", " & CurrCommand.Aliases
86                           End If
87                       End If
88                       CurrSection = InNone

89                   Case "SHORT", "ONELINE", "SHORTDESCRIPTION", "SHORTDESC", "ONE"
90                       If Not CurrCommand Is Nothing Then CurrCommand.OneLineDescription = sAfter(CurrLine)
91                       CurrSection = InNone

92                   Case "SYNTAX"
93                       CurrSection = InSyntax
94                       If Not CurrCommand Is Nothing Then CurrCommand.Syntax = sAfter(CurrLine)

95                   Case "LONG", "LONGDESCRIPTION", "LONGDESC"
96                       CurrSection = InLong
97                       If Not CurrCommand Is Nothing Then CurrCommand.LongDescription = sAfter(CurrLine)

98                   Case "COMMENTS", "REMARKS", "NOTES"
99                       CurrSection = InComments
100                      If Not CurrCommand Is Nothing Then CurrCommand.Comments = sAfter(CurrLine)

101                  Case "SEEALSO", "SEE"
102                      CurrSection = InSeeAlso
103                      If Not CurrCommand Is Nothing Then CurrCommand.SeeAlso = sAfter(CurrLine)

104                  Case "EXAMPLE", "EXAMPLES"
105                      CurrSection = InExample
106                      If Not CurrCommand Is Nothing Then CurrCommand.Examples = sAfter(CurrLine)

107                  Case "HELPFILE", "HELP"
108                      If Not CurrCommand Is Nothing Then CurrCommand.HelpFile = sAfter(CurrLine)
109                      CurrSection = InNone

110                  Case "HELP TOPIC", "TOPIC"
111                      If Not CurrCommand Is Nothing Then CurrCommand.HelpTopic = sAfter(CurrLine)
112                      CurrSection = InNone

113                  Case "END"
114                      NewData = vbNullString
115                      CurrSection = InNone

116                  Case Else
117                      CurrSection = InNone
118              End Select
119          Else
            Select Case CurrSection
                Case InNone
120                  Case InSyntax: If Not CurrCommand Is Nothing Then CurrCommand.Syntax = CurrCommand.Syntax & CurrLine & vbNewLine
121                  Case InLong: If Not CurrCommand Is Nothing Then CurrCommand.LongDescription = CurrCommand.LongDescription & CurrLine & vbNewLine
122                  Case InComments: If Not CurrCommand Is Nothing Then CurrCommand.Comments = CurrCommand.Comments & CurrLine & vbNewLine
123                  Case InSeeAlso: If Not CurrCommand Is Nothing Then CurrCommand.SeeAlso = CurrCommand.SeeAlso & CurrLine & vbNewLine
124                  Case InExample: If Not CurrCommand Is Nothing Then CurrCommand.Examples = CurrCommand.Examples & CurrLine & vbNewLine
125              End Select
126          End If
127      Loop

128      CurrIndex = 1
129      For Each CurrCommand In mCol
130          CurrCommand.Index = CurrIndex
131          CurrIndex = CurrIndex + 1
132      Next CurrCommand
End Property

Public Sub Clear(Optional ByVal bRecreateCollection As Boolean = True)
133      On Error Resume Next
134      Set mCol = Nothing
135      If bRecreateCollection Then Set mCol = New Collection
End Sub

Public Sub HideWindow(Optional ByVal bUnloadAsWell As Boolean = False)
On Error Resume Next
136      If Not MyHelpScreen Is Nothing Then
137          MyHelpScreen.Hide
138          If bUnloadAsWell Then
139              Unload MyHelpScreen
140              Set MyHelpScreen = Nothing
141          End If
142      End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As CSadCommand
Attribute Item.VB_UserMemId = 0
143      On Error Resume Next
144      Dim TempItem As CSadCommand
145      Dim TempKey As String

146      If VarType(vntIndexKey) = vbString Then
147          Err.Clear
148          Set TempItem = Nothing
149          Set TempItem = mCol(UCase$(vntIndexKey))
150          If TempItem Is Nothing Then
151              Set TempItem = mCol(UCase$(vntIndexKey) & "*C")
152              If TempItem Is Nothing Then
153                  Set TempItem = mCol(UCase$(vntIndexKey) & "*I")
154                  If Not TempItem Is Nothing Then
155                      Set Item = TempItem
156                  End If
157              Else
158                  Set Item = TempItem
159              End If
160          Else
161              Set Item = TempItem
162          End If
163          If TempItem Is Nothing Then
            ' Search Aliases
164              If InStr(vntIndexKey, "*I") Then
165                  TempKey = sGetToken(vntIndexKey, 1, "*I")
166              ElseIf InStr(vntIndexKey, "*C") Then
167                  TempKey = sGetToken(vntIndexKey, 1, "*C")
168              Else
169                  TempKey = vntIndexKey
170              End If
171              TempKey = ", " & TempKey & ", "
172              For Each TempItem In mCol
173                  If InStr(UCase$(TempItem.Aliases), UCase$(TempKey)) Then
174                      Set Item = TempItem
175                      Exit For
176                  End If
177              Next TempItem
178          End If
179      Else
180          Set Item = mCol(vntIndexKey)
181      End If

182      Set TempItem = Nothing

End Property

Public Property Get Count() As Long
183      On Error Resume Next
184      Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As String)
185      On Error Resume Next
186      mCol.Remove UCase$(vntIndexKey)
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
187      On Error Resume Next
188      Set NewEnum = mCol.[_NewEnum]
End Property

Public Sub ShowHelpScreen(Optional ByVal sSoftCommand As String)
189      On Error Resume Next
190      With MyHelpScreen
191          Set .SadCommandSet = Me
192          If Len(sSoftCommand) Then
193              If Not Item(sSoftCommand) Is Nothing Then
194                  .CurrCommandKey = Item(sSoftCommand).Index
195              End If
196          Else
197              .CurrCommandKey = 1
198          End If
199          .txtAliases.BackColor = "&H" & GetSetting$(App.ProductName, "Last", "Background Color", "&H80000018&")
200          .txtSyntax.BackColor = .txtAliases.BackColor
201          .txtOneLineDescription.BackColor = .txtAliases.BackColor
202          .txtHelpFile.BackColor = .txtAliases.BackColor
203          .txtHelpTopic.BackColor = .txtAliases.BackColor
204          .txtLongDescription.BackColor = .txtAliases.BackColor
205          .txtComments.BackColor = .txtAliases.BackColor
206          .txtSeeAlso.BackColor = .txtAliases.BackColor
207          .txtExamples.BackColor = .txtAliases.BackColor

208          .txtAliases.ForeColor = "&H" & GetSetting$(App.ProductName, "Last", "Foreground Color", "&H80000008&")
209          .txtSyntax.ForeColor = .txtAliases.ForeColor
210          .txtOneLineDescription.ForeColor = .txtAliases.ForeColor
211          .txtHelpFile.ForeColor = .txtAliases.ForeColor
212          .txtHelpTopic.ForeColor = .txtAliases.ForeColor
213          .txtLongDescription.ForeColor = .txtAliases.ForeColor
214          .txtComments.ForeColor = .txtAliases.ForeColor
215          .txtSeeAlso.ForeColor = .txtAliases.ForeColor
216          .txtExamples.ForeColor = .txtAliases.ForeColor
217          .Show 0, Parent.UI
218      End With
End Sub

Private Sub Class_Initialize()
219      On Error Resume Next
220      Set mCol = New Collection
221      ParameterDelimiter = gsC
222      ParameterTypeDelimiter = ":"

223      Set masaParams = New CAssocArray
224      masaParams.ItemDelimiter = ParameterDelimiter
225      masaParams.FieldDelimiter = ParameterTypeDelimiter

226      Set masaSyntax = New CAssocArray
227      masaSyntax.ItemDelimiter = ParameterDelimiter
228      masaSyntax.FieldDelimiter = ParameterTypeDelimiter

229      Set masaAttributes = New CAssocArray

230      Set MyHelpScreen = New frmCommandHelp
231      With MyHelpScreen
232          Set .SadCommandSet = Me
233          .CurrCommandKey = 1
234      End With
End Sub

Private Sub Class_Terminate()
235      On Error Resume Next
236      With MyHelpScreen
237          .mnuFileExit_Click
238          Set .SadCommandSet = Nothing
239          Set .CurrCommand = Nothing
240      End With
241      Unload MyHelpScreen
242      Set MyHelpScreen = Nothing

243      Set masaParams = Nothing
244      Set masaSyntax = Nothing
245      Set masaAttributes = Nothing
246      Set mCol = Nothing
End Sub

