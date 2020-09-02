VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy - Slice and Dice for Visual Basic"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

' ================================================================================
' Class Module      Wizard
'
' Filename          Wizard.cls
'
' Author            William M. Rawls
'
' Created On        9/3/1997 8:00 pm
'
' Description
'
' The actual Add-in object attached to the VB environment
'
' Revisions
'
' <Date>, William M. Rawls
' <Description of Revision>
'
' ================================================================================
Option Explicit

Implements IDTExtensibility

Private mcbAddinButton As Office.CommandBarButton
Private mcbEditButton As Office.CommandBarButton
Private mcbShortcutButton As Office.CommandBarButton
Private mcbChangeToButton As Office.CommandBarButton
Private mcbAltChangeToButton As Office.CommandBarButton
Private mcbFavoritesButton As Office.CommandBarButton
Private mcbExternalsButton As Office.CommandBarButton

Private WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
Private WithEvents BarHandler As CommandBarEvents
Attribute BarHandler.VB_VarHelpID = -1
Private WithEvents ShortcutHandler As CommandBarEvents
Attribute ShortcutHandler.VB_VarHelpID = -1
Private WithEvents ChangeToHandler As CommandBarEvents
Attribute ChangeToHandler.VB_VarHelpID = -1
Private WithEvents AltChangeToHandler As CommandBarEvents
Attribute AltChangeToHandler.VB_VarHelpID = -1
Private WithEvents FavoritesHandler As CommandBarEvents
Attribute FavoritesHandler.VB_VarHelpID = -1
Private WithEvents ExternalsHandler As CommandBarEvents
Attribute ExternalsHandler.VB_VarHelpID = -1

Private m_oUI     As frmMain
Private m_oWindow As frmMain
Private m_oVBInst As VBIDE.VBE

Private PropertyStack As CAssocArray
Private asaCustom     As CAssocArray

Public HostedByVB As Boolean

Public Function EnumFiles(sPath As String, sMask As String, sAttr As String) As String
1        Dim Attr As VbFileAttribute

    Select Case Attr
        Case "ALIAS": Attr = vbAlias
2            Case "ARCHIVE": Attr = vbArchive
3            Case "DIRECTORY": Attr = vbDirectory
4            Case "HIDDEN": Attr = vbHidden
5            Case "READONLY": Attr = vbReadOnly
6            Case "SYSTEM": Attr = vbSystem
7            Case "VOLUME": Attr = vbVolume
8            Case Else: Attr = vbNormal
9        End Select
10       EnumFiles = modGeneral.EnumFiles(sPath, sMask, Attr)
End Function

Public Function Evaluate(ByVal sExpression As String, ByRef asaVar As CAssocArray)
11       Evaluate = m_oUI.Evaluate(sExpression, asaVar)
End Function


Public Function FileExists(sFilename As String) As Boolean
12       FileExists = modGeneral.FileExists(sFilename)
End Function

Public Sub HandleKeyPress(ByRef KeyCode As Integer, ByRef Shift As Integer)
13       If Shift = 3 Then
        Select Case KeyCode
            Case 69                                   ' "E"xternals window
14                   ExternalsHandler_Click Nothing, False, False
15                   KeyCode = 0
16                   Shift = 0
17               Case 70                                   ' "F"avorites
18                   FavoritesHandler_Click Nothing, False, False
19                   KeyCode = 0
20                   Shift = 0
21               Case 83                                   ' "S"lice and Dice window
22                   MenuHandler_Click Nothing, False, False
23                   KeyCode = 0
24                   Shift = 0
25               Case Else
26                   MsgBox "Combination key Shift-Ctrl-" & KeyCode & " pressed"
27                   KeyCode = 0
28                   Shift = 0
29           End Select
30       End If
End Sub

Public Function JumpTo(ByVal sTemplateName As String, Optional ByVal bRecordInHistory As Boolean = True, Optional ByVal bSyncCategoryList As Boolean = False) As Boolean
31       m_oUI.JumpTo sTemplateName, bRecordInHistory, bSyncCategoryList
End Function

Public Sub NewTemplate(Optional ByVal bAutoCreate As Boolean = False, Optional ByVal sTitle As String, Optional ByVal sDefaultShortName As String, Optional ByVal bJumpToAfterCreate As Boolean = True)
32       m_oUI.NewTemplate bAutoCreate, sTitle, sDefaultShortName, bJumpToAfterCreate
End Sub

Public Function sChooseColor(Optional ByVal sInitialColor As String) As String
33       sChooseColor = m_oUI.sChooseColor(sInitialColor)
End Function

Public Function sChooseFile(Optional ByVal sPath As String, Optional ByVal sFilename As String, Optional ByVal sFilter As String) As String
34       sChooseFile = m_oUI.sChooseFile(sPath, sFilename, sFilter)
End Function

Public Sub ShowSplashScreen()
35       On Error Resume Next
36       frmSplash.DetermineRegistration
37       frmSplash.Show
End Sub

Public Property Get TemplateDatabaseName() As String
38       On Error Resume Next
39       TemplateDatabaseName = m_oUI.TemplateDatabaseName
End Property

Public Function GetATemplate(sCategory As String, sTemplate As String) As CTemplate
40       On Error Resume Next
41       Set GetATemplate = Nothing

42       If StrComp(sCategory, gsSpecialLineItemDelimiter & "CURRENT " & UCase$(gsCategory) & gsSpecialLineItemDelimiter, vbTextCompare) = 0 Then
43           Set GetATemplate = m_oUI.SliceAndDice.Categorys(sGetToken(m_oUI.InternalCurrentTemplate, 1, gsCategoryTemplateDelimiter)).Templates(sTemplate)
44       Else
45           Set GetATemplate = m_oUI.SliceAndDice.Categorys(sCategory).Templates(sTemplate)
46       End If
End Function

Public Function LogError(ByVal sModuleName As String, sProcName As String, lError As Long, sErrorMsg As String) As Boolean
47       LogError = modGeneral.LogError(sModuleName, sProcName, lError, sErrorMsg, Erl)
End Function

Public Function sFileContents(sPathAndFilename As String) As String
48       sFileContents = modGeneral.sFileContents(sPathAndFilename)
End Function

Public Function sGetGUID(ByVal sProgID As String) As String
49       sGetGUID = modGeneral.sGetGUID(sProgID)
End Function

Public Function sChoose(sChoices As String, Optional ByVal sDelimiter As String = gsSC, Optional ByVal sDefault As String)
50       sChoose = modGeneral.sChoose(sChoices, sDelimiter, sDefault)
End Function

Public Function sExtractToken(ByRef sOrigStr As String, Optional ByVal nToken As Integer = 1, Optional ByVal strDelim As String = gsS)
51       sExtractToken = modGeneral.sExtractToken(sOrigStr, nToken, strDelim)
End Function

Public Function BrowseForFolder(ByVal sPrompt As String) As String
52       BrowseForFolder = modGeneral.BrowseForFolder(m_oUI.hwnd, sPrompt)
End Function

Public Function bUserSure(Optional ByVal sPrompt As String = "Are you sure this is what you want to do ?") As Boolean
53       bUserSure = modGeneral.bUserSure(sPrompt)
End Function

Public Function NextNegativeUnique() As Long
54       NextNegativeUnique = modGeneral.NextNegativeUnique
End Function

Public Function lTokenCount(ByVal sAllTokens As String, Optional ByVal sDelim As String = gsS) As Long
55       lTokenCount = modGeneral.lTokenCount(sAllTokens, sDelim)
End Function

Public Function nZ(ByRef vData As Variant, Optional sDefault As String = vbNullString) As String
56       nZ = modGeneral.nZ(vData, sDefault)
End Function

Public Function sDenormalize(sLine As String) As String
57       sDenormalize = modGeneral.sDenormalize(sLine)
End Function

Public Function sGetToken(ByVal sAllTokens As String, Optional ByVal lToken As Long = 1, Optional ByVal sDelim As String = gsS) As String
58       sGetToken = modGeneral.sGetToken(sAllTokens, lToken, sDelim)
End Function

Public Function sAfter(ByVal sAllTokens As String, Optional ByVal lToken As Long = 1, Optional ByVal sDelim As String = gsS) As String
59       sAfter = modGeneral.sAfter(sAllTokens, lToken, sDelim)
End Function

Public Function sBefore(ByVal sAllTokens As String, Optional ByVal lToken As Long = 2, Optional ByVal sDelim As String = gsS) As String
60       sBefore = modGeneral.sBefore(sAllTokens, lToken, sDelim)
End Function

Public Function sExcept(ByVal sAllTokens As String, Optional ByVal lToken As Long = 1, Optional ByVal sDelim As String = gsS) As String
61       sExcept = modGeneral.sExcept(sAllTokens, lToken, sDelim)
End Function

Public Function sInsertSpaces(ByVal sToInsertInto As String) As String
62       sInsertSpaces = modGeneral.sInsertSpaces(sToInsertInto)
End Function

Public Function sNormalize(sLine As String) As String
63       sNormalize = modGeneral.sNormalize(sLine)
End Function

Public Function sReplace(ByVal sAll As String, ByVal sFind As String, ByVal sReplaceWith As String) As String
64       sReplace = Replace(sAll, sFind, sReplaceWith)
End Function

Public Property Get TemplateDatabasePath() As String
65       On Error Resume Next
66       TemplateDatabasePath = sBefore(m_oUI.TemplateDatabaseName, lTokenCount(m_oUI.TemplateDatabaseName, gsBS), gsBS) & gsBS
End Property

Public Property Get Version() As String
67       Version = App.Major & gsP & App.Minor & gsP & App.Revision
End Property

Public Function zn(sData As String) As Variant
68       zn = modGeneral.zn(sData)
End Function

Public Function lFindToken(ByVal sAllTokens As String, ByVal sTokenToFind As String, Optional ByVal sDelimiter As String = gsS) As Long
69       lFindToken = modGeneral.lFindToken(sAllTokens, sTokenToFind, sDelimiter)
End Function

Public Function AddPopup(mcp As CommandBarPopup, sMenu As String, nBefore As Long, sCaption As String, bBeginGroup As Boolean, bNextBeginGroup As Boolean) As String
70       On Error GoTo EH_Wizard_AddPopup
71       Static bInHereAlready As Boolean
72       If bInHereAlready Then Exit Function
73       bInHereAlready = True

         If Not HostedByVB Then Exit Function     ' Shell App Override

74       If m_oVBInst.CommandBars(sMenu) Is Nothing Then
75           MsgBox "Hmm... There doesn't seem to be anywhere I can place the following " & gsSliceAndDice & " button on the (nonexistant) '" & sMenu & "' menu:" & gs2EOLTab & "With Caption: " & sCaption & gs2EOLTab & "At position: " & nBefore
76           Exit Function
77       End If

78       If m_oVBInst.CommandBars(sMenu).Controls.Count < nBefore Then
79           If m_oVBInst.CommandBars(sMenu).Controls.Count > 0 Then
80               nBefore = m_oVBInst.CommandBars(sMenu).Controls.Count
81           Else
            ' No where to put that button
82               MsgBox "Hmm... There doesn't seem to be anywhere I can place the following " & gsSliceAndDice & " button on the '" & sMenu & "' menu:" & gs2EOLTab & "With Caption: " & sCaption & gs2EOLTab & "At position: " & nBefore
83               Exit Function
84           End If
85       End If

86       Set mcp = m_oVBInst.CommandBars(sMenu).Controls.Add(before:=nBefore, Type:=msoControlPopup)
87       If mcp Is Nothing Then
88           mcp.BeginGroup = bBeginGroup
89           mcp.Caption = sCaption
90           If bNextBeginGroup = True Then m_oVBInst.CommandBars(sMenu).Controls(nBefore + 1).BeginGroup = True
91           AddPopup = mcp.CommandBar.Name
92       End If

93 EH_Wizard_AddPopup_Continue:
94       bInHereAlready = False
95       Exit Function

96 EH_Wizard_AddPopup:
97       modGeneral.LogError "Wizard", "AddPopup", Err.Number, Err.Description, Erl

98       Resume EH_Wizard_AddPopup_Continue

99       Resume
End Function

Public Function AddButton(mcb As CommandBarButton, sMenu As String, nBefore As Long, sCaption As String, nButtonFace As Long, bBeginGroup As Boolean, bNextBeginGroup As Boolean) As CommandBarEvents
100      On Error GoTo EH_Wizard_AddButton
101      Static bInHereAlready As Boolean
102      If bInHereAlready Then Exit Function
103      bInHereAlready = True

         If Not HostedByVB Then Exit Function     ' Shell App Override

104      If m_oVBInst.CommandBars(sMenu) Is Nothing Then
105          MsgBox "Hmm... There doesn't seem to be anywhere I can place the following " & gsSliceAndDice & " button on the (nonexistant) '" & sMenu & "' menu:" & gs2EOLTab & "With Caption: " & sCaption & gs2EOLTab & "At position: " & nBefore
106          Exit Function
107      End If

108      If m_oVBInst.CommandBars(sMenu).Controls.Count < nBefore Then
109          If m_oVBInst.CommandBars(sMenu).Controls.Count > 0 Then
110              nBefore = m_oVBInst.CommandBars(sMenu).Controls.Count
111          Else
            ' No where to put that button
112              MsgBox "Hmm... There doesn't seem to be anywhere I can place the following " & gsSliceAndDice & " button on the '" & sMenu & "' menu:" & gs2EOLTab & "With Caption: " & sCaption & gs2EOLTab & "At position: " & nBefore
113              Exit Function
114          End If
115      End If

116      Set mcb = m_oVBInst.CommandBars(sMenu).Controls.Add(before:=nBefore)
117      If Not mcb Is Nothing Then
118          mcb.BeginGroup = bBeginGroup
119          If Len(sCaption) > 0 Then mcb.Caption = sCaption
120          If nButtonFace <> 0 Then mcb.FaceId = nButtonFace
121          On Error Resume Next
122          If m_oVBInst.CommandBars(sMenu).Controls.Count >= nBefore + 1 Then
123              m_oVBInst.CommandBars(sMenu).Controls(nBefore + 1).BeginGroup = bNextBeginGroup
124              On Error GoTo EH_Wizard_AddButton
125          End If
126          Set AddButton = m_oVBInst.Events.CommandBarEvents(mcb)
127      End If

128 EH_Wizard_AddButton_Continue:
129      bInHereAlready = False
130      Exit Function

131 EH_Wizard_AddButton:
132      modGeneral.LogError "Wizard", "AddButton", Err.Number, Err.Description & gs2EOLTab & "Menu: " & sMenu & gs2EOLTab & "Caption: " & sCaption & gs2EOLTab & "At position: " & nBefore, Erl

133      Resume EH_Wizard_AddButton_Continue

134      Resume
End Function

Public Sub DeleteCurrentTextSelection()
135      m_oUI.DeleteCurrentTextSelection
End Sub

Public Function DetermineFirstLineInSelection() As Long
136      DetermineFirstLineInSelection = m_oUI.DetermineFirstLineInSelection
End Function

Public Function DetermineLastLineInSelection() As Long
137      DetermineLastLineInSelection = m_oUI.DetermineLastLineInSelection
End Function


Public Sub DoInsertion(asaV As CAssocArray, sTemplateToInsert As String, Optional ByVal bSkipDeclarations As Boolean = False)
138      m_oUI.DoInsertion asaV, sTemplateToInsert, bSkipDeclarations
End Sub


Public Function FillTemplateWithUserInput(ByRef asaX As CAssocArray, ByVal sToParse As String, ByRef sCodeToInsert As String, ByVal sMsgBoxTitle As String) As Boolean
139      FillTemplateWithUserInput = m_oUI.FillTemplateWithUserInput(asaX, sToParse, sCodeToInsert, sMsgBoxTitle)
End Function

Public Function FindLastProcLine(sProcName As String, lProcType As Long) As Long
140      FindLastProcLine = m_oUI.FindLastProcLine(sProcName, lProcType)
End Function

Public Sub GetCategoryAndName(ByVal sCategoryAndName As String, ByRef sCategory As String, ByRef sShortName As String)
141      m_oUI.GetCategoryAndName sCategoryAndName, sCategory, sShortName
End Sub

Public Function GetCurrentTextSelection() As String
142      GetCurrentTextSelection = m_oUI.GetCurrentTextSelection
End Function

Public Sub GetProcAtLine(ByVal lCurrentLine As Long, ByRef sProcName As String, ByRef lProcType As Long)
143      m_oUI.GetProcAtLine lCurrentLine, sProcName, lProcType
End Sub

Public Function InsertTemplate(ByVal sToParseIn As String, Optional ByVal lLineIn As Long = 1, Optional ByRef SoftVars As CAssocArray, Optional ByVal sTemplateName As String = "Slice and Dice", Optional ByVal sExternalFilename As String = vbNullString) As Boolean
144      Dim InsertionInfo As CInsertionInfo

145      Set InsertionInfo = New CInsertionInfo
146      With InsertionInfo
147          If SoftVars Is Nothing Then
148              Set .SoftVars = New CAssocArray
149          Else
150              Set .SoftVars = SoftVars
151          End If

             With .SoftVars
                ' Set up commonly used variables
                  .Item("Windows Path") = WindowsDirectory


                  .Item("System Path") = WindowsSystemDirectory
             End With

152          .OriginalCodeToInsert = sToParseIn
153          .PointOfInsertion = lLineIn
154          .ExternalFilename = sExternalFilename
155          .TemplateName = sTemplateName
156      End With

157      InsertTemplate = m_oUI.InternalInsertTemplate(InsertionInfo)
158      Set InsertionInfo.SoftVars = Nothing
159      Set InsertionInfo = Nothing
End Function

Public Property Get InternalCurrentTemplate() As CTemplate
160      Set InternalCurrentTemplate = m_oUI.InternalCurrentTemplate
End Property

Public Property Get CurrentTemplate() As CTemplate
161      Set CurrentTemplate = m_oUI.CurrentTemplate
End Property


Public Function sChooseDatabase(Optional ByVal sPath As String, Optional ByVal sFilename As String) As String
162      sChooseDatabase = m_oUI.sChooseDatabase(sPath, sFilename)
End Function

Public Property Get SliceAndDice() As CSliceAndDice
163      Set SliceAndDice = m_oUI.SliceAndDice
End Property

Public Function sPropertyType(sFieldType As String) As String
164      sPropertyType = m_oUI.sPropertyType(sFieldType)
End Function

Friend Property Get UI() As frmMain
165      Set UI = m_oUI
End Property

Public Property Get SandyWindow() As Object
166      Set SandyWindow = m_oUI
End Property

Friend Property Get Window() As Window
167      Set Window = m_oWindow
End Property

Friend Property Get vbInst() As VBE
168      Set vbInst = m_oVBInst
End Property

Public Property Get CurrentTemplateNameAndCategory() As String
169      CurrentTemplateNameAndCategory = m_oUI.txtName.Text
End Property

Private Sub AltChangeToHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
170      ChangeToHandler_Click CommandBarControl, handled, CancelDefault
End Sub

Private Sub BarHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
171      MenuHandler_Click CommandBarControl, handled, CancelDefault
End Sub

Private Sub ChangeToHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
         If Not HostedByVB Then Exit Sub

172      Dim CurControl As VBControl
173      Dim NewControl As VBControl
174      Dim asaVar As CAssocArray

175      Dim sChoices As String
176      Dim sTemplate As String
177      Dim sProgID As String
178      Dim sLine As String
179      Dim sLastClassName As String
180      Dim sLastChoice As String
181      Dim sCodeToInsert As String
182      Dim sCodeToInsert2 As String
183      Dim sPropertyName As String

184      Dim bInTemplate As Boolean

185      Set asaVar = New CAssocArray
186      On Error Resume Next

187 ChangeToHandler_Click_Start_Over:
188      With m_oVBInst.SelectedVBComponent.Designer
189          gbCancelInsertion = False
190          For Each CurControl In .SelectedVBControls
191              If gbCancelInsertion Then Exit For
192              With CurControl
193                  If Not m_oUI.SetInternalCurrentTemplate("Change from" & gsCategoryTemplateDelimiter & .ClassName) Then
194                      If Not bUserSure("Please set up a 'Change to' " & gsTemplate & " named" & gsEolTab & "'Change from" & gsCategoryTemplateDelimiter & .ClassName & gsA & vbNewLine & "before using this function on the '" & .ClassName & "' type control." & gs2EOL & "Select Yes to create the new " & gsTemplate & gsP & vbNewLine & "Select No to abort.") Then
195                          Exit Sub
196                      End If

197                      m_oWindow.Visible = True
198                      m_oUI.NewTemplate True, "Change from" & gsCategoryTemplateDelimiter & .ClassName
199                      Exit Sub
200                  End If

201                  If .ClassName = sLastClassName Then
202                      sProgID = sLastChoice
203                  Else
204                      sChoices = vbNullString
205                      If m_oUI.SliceAndDice.Categorys("Change from").Templates("All Types") Is Nothing Or InStr(1, m_oUI.InternalCurrentTemplate.memoCodeAtTop, "NoAllTypes", vbTextCompare) Then
206                          sTemplate = m_oUI.InternalCurrentTemplate.memoCodeAtTop
207                      Else
208                          With m_oUI.SliceAndDice.Categorys("Change from").Templates("All Types")
209                              If Right$(.memoCodeAtTop, 2) = vbNewLine Then
210                                  sTemplate = .memoCodeAtTop & m_oUI.InternalCurrentTemplate.memoCodeAtTop
211                              Else
212                                  sTemplate = .memoCodeAtTop & vbNewLine & m_oUI.InternalCurrentTemplate.memoCodeAtTop
213                              End If
214                          End With
215                      End If

216                      Do Until Len(sTemplate) = 0
217                          sLine = sGetToken(sTemplate, 1, vbNewLine)
218                          sTemplate = sAfter(sTemplate, 1, vbNewLine)
219                          If Left$(sLine, 2) = gsSpecialLineItemDelimiter Then
220                              sChoices = sChoices & gsSC & Mid$(sLine, 4)
221                          End If
222                      Loop
223                      sChoices = Mid$(sChoices, 2)
224                      sProgID = sChoose(sChoices)
225                  End If
226                  If Len(sProgID) > 0 Then
227                      sLastClassName = .ClassName
228                      sLastChoice = sProgID

229                      If m_oUI.SliceAndDice.Categorys("Change from").Templates("All Types") Is Nothing Or InStr(1, m_oUI.InternalCurrentTemplate.memoCodeAtTop, "NoAllTypes", vbTextCompare) Then
230                          sTemplate = m_oUI.InternalCurrentTemplate.memoCodeAtTop
231                      Else
232                          With m_oUI.SliceAndDice.Categorys("Change from").Templates("All Types")
233                              If Right$(.memoCodeAtTop, 2) = vbNewLine Then
234                                  sTemplate = .memoCodeAtTop & m_oUI.InternalCurrentTemplate.memoCodeAtTop
235                              Else
236                                  sTemplate = .memoCodeAtTop & vbNewLine & m_oUI.InternalCurrentTemplate.memoCodeAtTop
237                              End If
238                          End With
239                      End If
240                      bInTemplate = False
241                      Do Until Len(sTemplate) = 0
242                          sLine = sGetToken(sTemplate, 1, vbNewLine)
243                          sTemplate = sAfter(sTemplate, 1, vbNewLine)
244                          If sLine = gsSpecialLineItemDelimiter & gsS & sProgID Then
245                              bInTemplate = True
246                              sTemplate = sGetToken(sTemplate, 1, vbNewLine & gsSpecialLineItemDelimiter & gsS)
247                              asaVar("ClassName") = CurControl.ClassName
248                              asaVar("ProgID") = CurControl.ProgId
249                              If Not m_oUI.FillTemplateWithUserInput(asaVar, sTemplate, sCodeToInsert, "GROUP MODIFY") Then
250                                  Exit Sub
251                              End If
252                              sTemplate = sCodeToInsert
253                          ElseIf Left$(sLine, 2) = gsSpecialLineItemDelimiter And bInTemplate = True Then
254                              bInTemplate = False
255                              sTemplate = vbNullString
256                          ElseIf bInTemplate = True Then
                            Select Case UCase$(sGetToken(sLine))
                                Case "ADDCONTROL"
257                                      Set NewControl = .Container.ContainedVBControls.Add(sAfter(sLine), CurControl, True)
258                                      If NewControl Is Nothing Then
259                                          MsgBox "That control has not been referenced yet. Please add a reference first.", vbInformation
260                                          sTemplate = vbNullString
261                                      End If

262                                  Case "SETNAME"
263                                      NewControl.Properties("Name") = sAfter(sLine) & Mid$(CurControl.Properties("Name"), Len(sAfter(sLine)))

264                                  Case "COPYPROPERTY"
265                                      NewControl.Properties(sAfter(sLine)) = CurControl.Properties(sAfter(sLine))

266                                  Case "SETPROPERTY"
267                                      sPropertyName = sGetToken(sAfter(sLine), 1, gsE)
268                                      If InStr(sPropertyName, gsP) > 0 Then
269                                          NewControl.Properties(sGetToken(sPropertyName, 1, gsP)).Value(sGetToken(sPropertyName, 2, gsP)).Value = sAfter(sLine, 1, gsE)
270                                      Else
271                                          NewControl.Properties(sPropertyName) = sAfter(sLine, 1, gsE)
272                                      End If
273                                      If Err.Number <> 0 Then
274                                          MsgBox "Warning: Failed to set the property '" & sPropertyName & "' to '" & sAfter(sLine, 1, gsE) & gsEolTab & Err.Description
275                                      End If
                                    'NewControl.Properties(sGetToken(sAfter(sLine), 1, gsE)) = sAfter(sLine, 1, gsE)

276                                  Case "SETCURRENTPROPERTY"
277                                      Err.Clear
278                                      sPropertyName = sGetToken(sAfter(sLine), 1, gsE)
279                                      If InStr(sPropertyName, gsP) > 0 Then
280                                          CurControl.Properties(sGetToken(sPropertyName, 1, gsP)).Value(sGetToken(sPropertyName, 2, gsP)).Value = sAfter(sLine, 1, gsE)
281                                      Else
282                                          CurControl.Properties(sPropertyName) = sAfter(sLine, 1, gsE)
283                                      End If
284                                      If Err.Number <> 0 Then
285                                          MsgBox "Warning: Failed to set the property '" & sPropertyName & "' to '" & sAfter(sLine, 1, gsE) & gsEolTab & Err.Description
286                                      End If

287                                  Case "DELETECURRENTCONTROL"
                                         If HostedByVB Then
288                                         m_oVBInst.SelectedVBComponent.Designer.VBControls.Remove CurControl
289                                         Set CurControl = Nothing
                                         End If

290                                  Case "AUTOSPACECURRENT"
291                                      CurControl.Properties(sGetToken(sAfter(sLine), 1, gsE)) = sInsertSpaces(CurControl.Properties(sGetToken(sAfter(sLine), 1, gsE)))

292                                  Case "AUTOSPACE"
293                                      NewControl.Properties(sGetToken(sAfter(sLine), 1, gsE)) = sInsertSpaces(NewControl.Properties(sGetToken(sAfter(sLine), 1, gsE)))

294                                  Case "OFFSETFROMCURRENTLEFT", "OFFSET", "OFFSETLEFT"
295                                      NewControl.Properties("Left") = CurControl.Properties("Left") + CurControl.Properties("Width") + IIf(Val(sGetToken(sLine, 2)) = 0, 100, Val(sGetToken(sLine, 2)))

296                                  Case "OFFSETFROMCURRENTTOP", "OFFSETTOP"
297                                      NewControl.Properties("Top") = CurControl.Properties("Top") + CurControl.Properties("Height") + IIf(Val(sGetToken(sLine, 2)) = 0, 100, Val(sGetToken(sLine, 2)))

298                                  Case "QUEUE", "QUEUEINSERTION", "ADDTOQUEUE"
299                                      asaVar("Queued").Value = asaVar("Queued").Value & sAfter(sLine) & gsQueuedInsertionDelimiter

300                                  Case "COPYLOCATIONCURRENT", "COPYLOCATION"
301                                      If PropertyStack Is Nothing Then Set PropertyStack = New CAssocArray
302                                      PropertyStack("Left") = CurControl.Properties("Left")
303                                      PropertyStack("Top") = CurControl.Properties("Top")
304                                      PropertyStack("Width") = CurControl.Properties("Width")
305                                      PropertyStack("Height") = CurControl.Properties("Height")

306                                  Case "PASTELOCATIONCURRENT", "PASTELOCATION"
307                                      If PropertyStack Is Nothing Then
308                                          MsgBox "No location to paste." & vbNewLine & vbTab & "Use 'CopyLocation' before 'PastLocation'."
309                                      Else
310                                          If PropertyStack("Left") < 0 Then PropertyStack("Left") = 0
311                                          If PropertyStack("Top") < 0 Then PropertyStack("Top") = 0
312                                          If PropertyStack("Width") < 30 Then PropertyStack("Width") = 30
313                                          If PropertyStack("Height") < 30 Then PropertyStack("Height") = 30

314                                          CurControl.Properties("Left") = PropertyStack("Left")
315                                          CurControl.Properties("Top") = PropertyStack("Top")
316                                          CurControl.Properties("Width") = PropertyStack("Width")
317                                          CurControl.Properties("Height") = PropertyStack("Height")
318                                      End If

319                                  Case "CLEARPROPERTYSTACK", "CLEARSTACK"
320                                      If Not PropertyStack Is Nothing Then
321                                          PropertyStack.Clear
322                                          Set PropertyStack = Nothing
323                                      End If

324                                  Case "PUSHPROPERTY", "PUSH"
325                                      If PropertyStack Is Nothing Then Set PropertyStack = New CAssocArray
326                                      PropertyStack(sAfter(sLine)) = CurControl.Properties(sAfter(sLine))

327                                  Case "POPPORPERTYCURRENT", "POPPROPERTY", "POP"
328                                      If Not PropertyStack Is Nothing Then
329                                          CurControl.Properties(sAfter(sLine)) = PropertyStack(sAfter(sLine))
330                                      Else
331                                          MsgBox "No properties to pop off the stack." & gsEolTab & "Use 'PushProperty' before 'PopProperty'."
332                                      End If

333                                  Case "SETVALUE", "SET"
334                                      asaVar.Item(sGetToken(sAfter(sLine), 1, gsE)) = sAfter(sLine, 1, gsE)

335                                  Case "INSERTTEMPLATE", "INSERT"
336                                      gbCancelInsertion = False
337                                      m_oUI.DoInsertion asaVar, sAfter(sLine)
338                                      If gbCancelInsertion Then
339                                          sTemplate = vbNullString
340                                      End If

341                                  Case "GETPROPERTY", "GET"
342                                      asaVar.Item(sGetToken(sAfter(sLine), 1, gsE)) = CurControl.Properties(sAfter(sAfter(sLine), 1, gsE))
343                              End Select
344                          End If
345                      Loop
346                  End If
347              End With
348          Next CurControl
349      End With

350      If Not gbCancelInsertion And Len(asaVar("Queued").Value) Then
351          m_oUI.QueuedInsertions = asaVar("Queued").Value
352      End If

353      handled = True

354      Set asaVar = Nothing
End Sub

Private Sub ExternalsHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
355      On Error Resume Next
356      m_oUI.ShowExternalsMenu

357      handled = True
End Sub

Private Sub FavoritesHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
358      On Error Resume Next
359      m_oUI.FavoriteCalledFromIDE = True
360      m_oUI.ShowFavMenu

361      handled = True
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
' Comment to prevent procedure from being deleted on compilation.
End Sub

Private Sub IDTExtensibility_OnConnection(ByVal vbInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
362      On Error Resume Next
363      Dim sTemplateDatabasePath As String
364      Dim CurWindow As Window
365      Dim bFound As Boolean
366      Dim bShown As Boolean
367      Dim lFirstButton As Long
         Dim sLoadList As String

368      frmSplash.lblDLLsLoaded(1).Caption = "0"
369      If GetSetting(App.ProductName, gsLast, "Show Splash", True) Then
370          bShown = True
371          frmSplash.Show
372          frmSplash.Refresh
373      Else
374          bShown = False
375      End If

376      Set m_oVBInst = vbInst                            ' Attach to the current VB IDE session

         HostedByVB = (Not m_oVBInst Is Nothing)
         
377      Set m_oWindow = New frmMain
378      Set m_oUI = m_oWindow

379      Set m_oUI.Parent = Me
380      Set m_oUI.DBClassGen = New frmDBClassGen
381      Set m_oUI.DBClassGen.Parent = m_oUI

382      DoEvents
383      sTemplateDatabasePath = TemplateDatabasePath
384      DoEvents: DoEvents: DoEvents

385     ' If HostedByVB Then
            m_oUI.InitializeAddinDLLs modGeneral.EnumFiles(App.Path) & modGeneral.EnumFiles(TemplateDatabasePath)
        ' Else
        '    sLoadList = custom(0)
        '    If Right$(sLoadList, 1) <> "|" Then sLoadList = sLoadList & "|"
        '    m_oUI.InitializeAddinDLLs Replace(custom(0), "|", ".NewCommands=Load" & vbNewLine)
        ' End If

386      Set MenuHandler = AddButton(mcbAddinButton, "Add-Ins", 4, "&" & gsSliceAndDice, 30, True, False)

387      If GetSetting(App.ProductName, gsLast, "Show Paitbrush Icon", True) Then
388          Set BarHandler = AddButton(mcbEditButton, "Standard", 1, gsSliceAndDice, 30, False, True)
389      End If

         If HostedByVB Then
390         If GetSetting(App.ProductName, gsLast, "Show On Module Right Click", True) Then
391              If Not m_oVBInst.CommandBars("Code Window") Is Nothing Then
392                 lFirstButton = IIf(m_oVBInst.CommandBars("Code Window").Controls.Count > 4, 4, 1)
393                 Set FavoritesHandler = AddButton(mcbFavoritesButton, "Code Window", lFirstButton, "Insert a &Favorite", 30, True, False)
394                 Set ExternalsHandler = AddButton(mcbExternalsButton, "Code Window", lFirstButton + 1, "Call an &External", 30, False, False)
395                 Set ShortcutHandler = AddButton(mcbShortcutButton, "Code Window", lFirstButton + 2, "&" & gsSliceAndDice, 30, False, True)
396             End If
397         End If
         End If

398      Set ChangeToHandler = AddButton(mcbChangeToButton, "Controls", 1, gsSliceAndDice & " Modify", 30, False, False)
399      Set AltChangeToHandler = AddButton(mcbAltChangeToButton, "Format", 8, "&" & gsSliceAndDice & " Modify", 30, False, False)

400      On Error Resume Next
401      If m_oUI.mHotKeyOpenWindow Is Nothing Then
402          Set m_oUI.mHotKeyOpenWindow = New cRegHotKey
403          If Not m_oUI.mHotKeyOpenWindow Is Nothing Then
404              With m_oUI.mHotKeyOpenWindow
405                  .Attach m_oUI.hwnd
406                  .RegisterKey "Sandy Activate", vbKeyS, MOD_CONTROL + MOD_SHIFT
407                  .RegisterKey "Sandy Repeat Insertion", vbKeyR, MOD_CONTROL + MOD_SHIFT
408                  .RegisterKey "Sandy Favorites", vbKeyF, MOD_CONTROL + MOD_SHIFT
409                  .RegisterKey "Sandy Externals", vbKeyE, MOD_CONTROL + MOD_SHIFT

410                  .RegisterKey "Sandy Cancel Insertion", vbKeyC, MOD_CONTROL + MOD_SHIFT

                ' Fast Insert (Favorites of the favorites)
411                  .RegisterKey "Sandy Fast Insert 1", vbKey1, MOD_CONTROL + MOD_SHIFT
412                  .RegisterKey "Sandy Fast Insert 2", vbKey2, MOD_CONTROL + MOD_SHIFT
413                  .RegisterKey "Sandy Fast Insert 3", vbKey3, MOD_CONTROL + MOD_SHIFT
414                  .RegisterKey "Sandy Fast Insert 4", vbKey4, MOD_CONTROL + MOD_SHIFT
415              End With
416          End If
417      End If
418      If bShown Then frmSplash.Hide
End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)
419      On Error Resume Next
420      Dim Cancel As Integer

    ' Make sure the edited Template (if one) is saved
421      m_oUI.SaveTemplate
422      m_oWindow.Visible = False

    ' Remove buttons from VB5 ToolBars
423      mcbAddinButton.Delete
424      mcbEditButton.Delete
425      mcbShortcutButton.Delete
426      mcbChangeToButton.Delete
427      mcbAltChangeToButton.Delete
428      mcbFavoritesButton.Delete
429      mcbExternalsButton.Delete

    ' Insure all external object references are released correctly
430      Set mcbAddinButton = Nothing
431      Set mcbEditButton = Nothing
432      Set mcbShortcutButton = Nothing
433      Set mcbChangeToButton = Nothing
434      Set mcbAltChangeToButton = Nothing
435      Set mcbFavoritesButton = Nothing

436      Set MenuHandler = Nothing
437      Set BarHandler = Nothing
438      Set ShortcutHandler = Nothing
439      Set ChangeToHandler = Nothing
440      Set AltChangeToHandler = Nothing
441      Set FavoritesHandler = Nothing
442      Set ExternalsHandler = Nothing

    ' Save settings for next time
443      SaveSetting App.ProductName, "Settings", "Exit after insert", IIf(m_oUI.ExitAfterInsert, "True", "False")
444      SaveSetting App.ProductName, "Settings", "Last " & gsTemplate, m_oUI.CurrentTemplateNameAndCategory
445      SaveFormPosition m_oWindow

    ' Destroy object references
446      m_oWindow.HideAllWindows True
447      m_oWindow.ShutdownDLLs
448      Set m_oWindow = Nothing
449      Set m_oUI.Parent = Nothing
450      Set m_oUI.DBClassGen = Nothing
451      m_oUI.Form_Unload Cancel
452      Unload m_oUI
453      Set m_oUI = Nothing

    ' Disassociate external objects referenced in this object
454      Set m_oVBInst = Nothing
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
' Comment to prevent procedure from being deleted on compilation.
End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
455      On Error Resume Next
456      m_oWindow.Visible = Not m_oWindow.Visible
457      If m_oWindow.Visible Then
458          m_oWindow.SetFocus
459          m_oWindow.ZOrder
460      Else
461          m_oWindow.HideAllWindows
462      End If

463      handled = True
End Sub

Private Sub ShortcutHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
464      MenuHandler_Click CommandBarControl, handled, CancelDefault
End Sub

Public Property Get SoftCommands() As CSadCommands
On Error Resume Next
    Set SoftCommands = m_oUI.Complete
End Property

Public Function SaveToFile(ByVal sFilename As String, ByVal sContents As String) As Boolean
    SaveToFile = modGeneral.SaveToFile(sFilename, sContents)
End Function

Public Sub BrowseTo(ByVal sURL As String)
    modGeneral.BrowseTo sURL
End Sub

Public Sub ShowExternalsMenu()
On Error Resume Next
    If Not m_oUI Is Nothing Then
       m_oUI.ShowExternalsMenu
    End If
End Sub

Public Sub ShowFavoritesMenu()
On Error Resume Next
    If Not m_oUI Is Nothing Then
       m_oUI.ShowFavMenu
    End If
End Sub

Public Sub ShowMainWindow()
    On Error Resume Next
    If m_oWindow Is Nothing Then Exit Sub
    
    m_oWindow.WindowState = 0
    
    m_oWindow.Visible = True
    If m_oWindow.Visible Then
       m_oWindow.SetFocus
       m_oWindow.ZOrder
    End If
End Sub

Public Sub HideWindows()
    On Error Resume Next
    If m_oWindow Is Nothing Then Exit Sub

    m_oWindow.Visible = False
    If Not m_oWindow.Visible Then
       m_oWindow.HideAllWindows
    End If
End Sub

Public Property Get FavoriteCalledFromIDE() As Boolean
    If m_oWindow Is Nothing Then Exit Property
    
    FavoriteCalledFromIDE = m_oWindow.FavoriteCalledFromIDE
End Property

Public Property Let FavoriteCalledFromIDE(RHS As Boolean)
    If m_oWindow Is Nothing Then Exit Property
    
    m_oWindow.FavoriteCalledFromIDE = RHS
End Property


