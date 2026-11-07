VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCommandHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Soft Command Reference"
   ClientHeight    =   6885
   ClientLeft      =   5280
   ClientTop       =   2655
   ClientWidth     =   7605
   Icon            =   "frmCommandHelp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   7605
   Begin MSComDlg.CommonDialog cdgHelp 
      Left            =   180
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSoftCommandName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   250
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   300
      Width           =   6600
   End
   Begin VB.TextBox txtAliases 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   6600
   End
   Begin VB.TextBox txtSyntax 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   900
      Width           =   6600
   End
   Begin VB.TextBox txtOneLineDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   6600
   End
   Begin VB.TextBox txtHelpFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1500
      Width           =   6600
   End
   Begin VB.TextBox txtHelpTopic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   6600
   End
   Begin VB.TextBox txtLongDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1200
      Left            =   1050
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2100
      Width           =   6600
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1200
      Left            =   1050
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3300
      Width           =   6600
   End
   Begin VB.TextBox txtSeeAlso 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1200
      Left            =   1050
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4500
      Width           =   6600
   End
   Begin VB.TextBox txtExamples 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1200
      Left            =   1050
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5700
      Width           =   6600
   End
   Begin VB.CheckBox chkIsInline 
      Appearance      =   0  'Flat
      Caption         =   "This is an Inline command if checked ( ie. %%Command::Parameters%% )"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1050
      TabIndex        =   0
      Top             =   0
      Width           =   6870
   End
   Begin VB.PictureBox picTH 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   765
      TabIndex        =   22
      Top             =   2460
      Width           =   765
   End
   Begin VB.Label lblSoftCommandName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   21
      Top             =   330
      Width           =   450
   End
   Begin VB.Label lblAliases 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aliases"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   20
      Top             =   630
      Width           =   570
   End
   Begin VB.Label lblSyntax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Syntax"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   19
      Top             =   930
      Width           =   540
   End
   Begin VB.Label lblOneLineDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   18
      Top             =   1230
      Width           =   765
   End
   Begin VB.Label lblHelpFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help File"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   17
      Top             =   1530
      Width           =   720
   End
   Begin VB.Label lblHelpTopic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help Topic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   16
      Top             =   1830
      Width           =   825
   End
   Begin VB.Label lblLongDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   15
      Top             =   2130
      Width           =   900
   End
   Begin VB.Label lblComments 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   14
      Top             =   3330
      Width           =   855
   End
   Begin VB.Label lblSeeAlso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "See Also"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   13
      Top             =   4530
      Width           =   675
   End
   Begin VB.Label lblExamples 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Examples"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   12
      Top             =   5730
      Width           =   735
   End
   Begin VB.Label lblIsInline 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Is Inline"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   11
      Top             =   60
      Width           =   660
   End
   Begin VB.Menu mnuFileExit 
      Caption         =   "&X"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFirst 
      Caption         =   "Fi&rst"
   End
   Begin VB.Menu mnuPrevious 
      Caption         =   "&<<"
   End
   Begin VB.Menu mnuNext 
      Caption         =   "&>>"
   End
   Begin VB.Menu mnuLast 
      Caption         =   "&Last"
   End
   Begin VB.Menu mnuFileFind 
      Caption         =   "&Find"
   End
   Begin VB.Menu mnuChangeCommandSets 
      Caption         =   "&Command Set"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmCommandHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SadCommandSet As CSadCommands
Public CurrCommand As CSadCommand
Public vCurrCommandKey As Variant

Public Property Let CurrCommandKey(ByVal vKey As Variant)
1        On Error Resume Next
2        If Not SadCommandSet.Item(vKey) Is Nothing Then
3            Set CurrCommand = SadCommandSet.Item(vKey)
4            vCurrCommandKey = CurrCommand.Index
5            Populate
6        End If
7        Err.Clear
End Property


Public Sub Populate()
8        On Error Resume Next
9        Dim NextTop As Long

10       If SadCommandSet Is Nothing Then Exit Sub
11       If CurrCommand Is Nothing Then Exit Sub

12       With CurrCommand
13           Caption = "SAD Soft Command Reference - " & SadCommandSet.Attributes("Name") & " ( " & SadCommandSet.Count & " commands)"
14           chkIsInline.Enabled = True
15           chkIsInline.Value = Abs(.IsInline)
16           If .IsInline Then
17               chkIsInline.Caption = "This is an INLINE Soft Command (ie. " & gsSoftVarDelimiter & .SoftCommandName & gsInlineCmdDelimiter & "Parameters" & gsSoftVarDelimiter & gsS & gsPC
18           Else
19               chkIsInline.Caption = "This is a REGULAR Soft Command (ie. " & gsSoftCmdDelimiter & .SoftCommandName & " )"
20           End If

21           NextTop = 390
22           If Len(.SoftCommandName) Then
23               txtSoftCommandName.Text = .SoftCommandName
24               txtSoftCommandName.Visible = True
25               txtSoftCommandName.Top = NextTop
26               lblSoftCommandName.Visible = True
27               lblSoftCommandName.Top = NextTop
28               txtSoftCommandName.Height = picTH.TextHeight(txtSoftCommandName.Text) + 100
29               NextTop = NextTop + txtSoftCommandName.Height
30           Else
31               txtSoftCommandName.Visible = False
32               lblSoftCommandName.Visible = False
33           End If

34           If Len(.Aliases) Then
35               txtAliases.Text = Mid$(.Aliases, 3, Len(.Aliases) - 4)
36               txtAliases.Visible = True
37               txtAliases.Top = NextTop
38               lblAliases.Visible = True
39               lblAliases.Top = NextTop
40               txtAliases.Height = picTH.TextHeight(txtAliases.Text) + 100
41               NextTop = NextTop + txtAliases.Height
42           Else
43               txtAliases.Visible = False
44               lblAliases.Visible = False
45           End If

46           If Len(.Syntax) Then
47               txtSyntax.Text = .Syntax
48               txtSyntax.Visible = True
49               txtSyntax.Top = NextTop
50               lblSyntax.Visible = True
51               lblSyntax.Top = NextTop
52               txtSyntax.Height = picTH.TextHeight(txtSyntax.Text) + 100
53               NextTop = NextTop + txtSyntax.Height
54           Else
55               txtSyntax.Visible = False
56               lblSyntax.Visible = False
57           End If

58           If Len(.OneLineDescription) Then
59               txtOneLineDescription.Text = .OneLineDescription
60               txtOneLineDescription.Visible = True
61               txtOneLineDescription.Top = NextTop
62               lblOneLineDescription.Visible = True
63               lblOneLineDescription.Top = NextTop
64               txtOneLineDescription.Height = picTH.TextHeight(txtOneLineDescription.Text) + 100
65               NextTop = NextTop + txtOneLineDescription.Height
66           Else
67               txtOneLineDescription.Visible = False
68               lblOneLineDescription.Visible = False
69           End If

70           If Len(.LongDescription) Then
71               txtLongDescription.Text = .LongDescription
72               txtLongDescription.Visible = True
73               txtLongDescription.Top = NextTop
74               lblLongDescription.Visible = True
75               lblLongDescription.Top = NextTop
76               txtOneLineDescription.Height = picTH.TextHeight(txtOneLineDescription.Text) + 100
77               NextTop = NextTop + txtLongDescription.Height
78           Else
79               txtLongDescription.Visible = False
80               lblLongDescription.Visible = False
81           End If

82           If Len(.Comments) Then
83               txtComments.Text = .Comments
84               txtComments.Visible = True
85               txtComments.Top = NextTop
86               lblComments.Visible = True
87               lblComments.Top = NextTop
88               txtComments.Height = picTH.TextHeight(txtComments.Text) + 100
89               NextTop = NextTop + txtComments.Height
90           Else
91               txtComments.Visible = False
92               lblComments.Visible = False
93           End If

94           If Len(.SeeAlso) Then
95               txtSeeAlso.Text = .SeeAlso
96               txtSeeAlso.Visible = True
97               txtSeeAlso.Top = NextTop
98               lblSeeAlso.Visible = True
99               lblSeeAlso.Top = NextTop
100              txtSeeAlso.Height = picTH.TextHeight(txtSeeAlso.Text) + 100
101              NextTop = NextTop + txtSeeAlso.Height
102          Else
103              txtSeeAlso.Visible = False
104              lblSeeAlso.Visible = False
105          End If

106          If Len(.Examples) Then
107              txtExamples.Text = .Examples
108              txtExamples.Visible = True
109              txtExamples.Top = NextTop
110              lblExamples.Visible = True
111              lblExamples.Top = NextTop
112              txtExamples.Height = picTH.TextHeight(txtExamples.Text) + 100
113              NextTop = NextTop + txtExamples.Height
114              lblExamples.ZOrder
115          Else
116              txtExamples.Visible = False
117              lblExamples.Visible = False
118          End If

119          If Len(.HelpFile) Then
120              txtHelpFile.Text = .HelpFile
121              txtHelpFile.Visible = True
122              txtHelpFile.Top = NextTop
123              lblHelpFile.Visible = True
124              lblHelpFile.Top = NextTop
125              txtHelpFile.Height = picTH.TextHeight(txtHelpFile.Text) + 100
126              NextTop = NextTop + txtHelpFile.Height
127          Else
128              txtHelpFile.Visible = False
129              lblHelpFile.Visible = False
130          End If

131          If .HelpTopic Then
132              txtHelpTopic.Text = .HelpTopic
133              txtHelpTopic.Visible = True
134              txtHelpTopic.Top = NextTop
135              lblHelpTopic.Visible = True
136              lblHelpTopic.Top = NextTop
137              txtHelpTopic.Height = picTH.TextHeight(txtHelpTopic.Text) + 100
138              NextTop = NextTop + txtHelpTopic.Height
139          Else
140              txtHelpTopic.Visible = False
141              lblHelpTopic.Visible = False
142          End If
143          Me.Height = NextTop + 600

144          picTH.Move txtHelpFile.Left, NextTop
145      End With
End Sub

Private Sub chkIsInline_Click()
146      On Error Resume Next
147      chkIsInline.Value = Abs(CurrCommand.IsInline)
End Sub

Private Sub Form_Initialize()

' LogEvent "frmCommandHelp: Initialize"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
148      On Error Resume Next
149      If (Shift And vbCtrlMask) > 0 Then
        Select Case KeyCode
            Case vbKeyPageUp: mnuPrevious_Click
150              Case vbKeyPageDown: mnuNext_Click
151              Case vbKeyF: mnuFileFind_Click
152          End Select
153      ElseIf Shift = 0 Then
        Select Case KeyCode
            Case vbKeyEscape: mnuFileExit_Click
154          End Select
155      End If
End Sub

Private Sub Form_Load()
156      On Error Resume Next
157      LoadFormPosition Me, , False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
158      If UnloadMode = vbFormControlMenu Then
159          Cancel = True
160      End If
161      mnuFileExit_Click
End Sub

Private Sub Form_Terminate()

' LogEvent "frmCommandHelp: Terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
162      On Error Resume Next
163      SaveFormPosition Me
End Sub


Public Sub mnuFileExit_Click()
164      On Error Resume Next
165      SaveFormPosition Me
166      Set CurrCommand = Nothing
167      Set SadCommandSet = Nothing
168      Hide
End Sub

Private Sub mnuFileFind_Click()
169      On Error Resume Next
170      Dim CurrMember As CSadCommand
171      Dim sChoices As String
172      Dim sChoice As String
173      Dim asaOrdered As CAssocArray

174      sChoices = vbNullString
175      For Each CurrMember In SadCommandSet
176          sChoices = sChoices & CurrMember.SoftCommandName & IIf(CurrMember.IsInline, " (Inline)" & gsE & CurrMember.OneLineDescription & gsSC, gsE & CurrMember.OneLineDescription & gsSC)
177          If Len(CurrMember.Aliases) Then
178              sChoices = sChoices & Replace(CurrMember.Aliases, ", ", IIf(CurrMember.IsInline, " (Inline)" & "=See " & CurrMember.SoftCommandName & gsSC, "=See " & CurrMember.SoftCommandName & gsSC))
179          End If
180      Next CurrMember
181      If Len(sChoices) Then
182          Set asaOrdered = New CAssocArray
183          With asaOrdered
184              .ItemDelimiter = gsSC
185              .KeyValueDelimiter = gsE
186              .AddInOrder = True
187              .All = Replace(Replace(sChoices, "; (Inline)", gsSC), ";;", gsSC)
188              sChoices = .All
189          End With
190          Set asaOrdered = Nothing
191          sChoice = sChoose(sChoices, , txtSoftCommandName.Text)

192          If Len(sChoice) Then
193              If InStr(sChoice, "(Inline)") Then
194                  CurrCommandKey = SadCommandSet.Item(Trim$(sGetToken(UCase$(sChoice), 1, " (INLINE)")) & "*I").Index
195              Else
196                  CurrCommandKey = SadCommandSet.Item(UCase$(sChoice) & "*C").Index
197              End If
198          End If
199      End If
End Sub

Private Sub mnuFirst_Click()
200      On Error Resume Next
201      CurrCommandKey = 1
End Sub

Private Sub mnuLast_Click()
202      On Error Resume Next
203      CurrCommandKey = SadCommandSet.Count
End Sub

Private Sub mnuNext_Click()
204      On Error Resume Next
205      If vCurrCommandKey + 1 < SadCommandSet.Count Then
206          CurrCommandKey = vCurrCommandKey + 1
207      End If
End Sub

Private Sub mnuPrevious_Click()
208      On Error Resume Next
209      If vCurrCommandKey - 1 > 0 Then
210          CurrCommandKey = vCurrCommandKey - 1
211      End If
End Sub

Private Sub txtHelpFile_Click()
212      txtHelpTopic_Click
End Sub


Private Sub txtHelpFile_DblClick()
213      txtHelpTopic_Click
End Sub


Private Sub txtHelpTopic_Click()
214      If Len(txtHelpFile) > 0 And Len(txtHelpTopic) = 0 Then
215          With cdgHelp
216              .HelpFile = txtHelpFile
            ' Go to the Click Event topic in the Help file.
            ' The number is determined in the [MAP] section
            ' of the .HPJ file for the .chm file. You can
            ' edit this number only if you are using the
            ' Microsoft Help Workshop to build your
            ' own Help file.
217              .HelpContext = txtHelpTopic
218              .HelpCommand = cdlHelpContext
219              .ShowHelp
220          End With
221      ElseIf Len(txtHelpTopic) > 0 Then
222          BrowseTo txtHelpTopic
223      End If
End Sub

Private Sub txtHelpTopic_DblClick()
224      txtHelpTopic_Click
End Sub


