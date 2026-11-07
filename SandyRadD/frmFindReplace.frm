VERSION 5.00
Begin VB.Form frmFindReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find/Replace"
   ClientHeight    =   2235
   ClientLeft      =   1800
   ClientTop       =   2370
   ClientWidth     =   6660
   Icon            =   "frmFindReplace.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1320.512
   ScaleMode       =   0  'User
   ScaleWidth      =   6253.378
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace &All"
      Height          =   390
      Left            =   5415
      TabIndex        =   13
      Top             =   1470
      Width           =   1140
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   390
      Left            =   5415
      TabIndex        =   12
      Top             =   1050
      Width           =   1140
   End
   Begin VB.CheckBox chkUsePatternMatching 
      Caption         =   "&Use Pattern Matching"
      Height          =   285
      Left            =   1815
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Match Ca&se"
      Height          =   285
      Left            =   1815
      TabIndex        =   8
      Top             =   1650
      Value           =   1  'Checked
      Width           =   2070
   End
   Begin VB.CheckBox chkFindWholeWordOnly 
      Caption         =   "Find Whole World &Only"
      Height          =   285
      Left            =   1815
      TabIndex        =   7
      Top             =   1410
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.ComboBox cboDirection 
      Height          =   315
      ItemData        =   "frmFindReplace.frx":0442
      Left            =   2730
      List            =   "frmFindReplace.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1020
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtFind 
      Height          =   345
      Left            =   1245
      TabIndex        =   1
      Top             =   75
      Width           =   4047
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Find &Next"
      Default         =   -1  'True
      Height          =   390
      Left            =   5415
      TabIndex        =   10
      Top             =   75
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   5415
      TabIndex        =   11
      Top             =   495
      Width           =   1140
   End
   Begin VB.TextBox txtReplace 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1245
      TabIndex        =   2
      Top             =   495
      Width           =   4047
   End
   Begin VB.Frame Frame1 
      Caption         =   " Search "
      Height          =   1365
      Left            =   30
      TabIndex        =   16
      Top             =   825
      Width           =   1710
      Begin VB.OptionButton optSearchArea 
         Caption         =   "Current &Database"
         Height          =   195
         Index           =   3
         Left            =   75
         TabIndex        =   17
         Top             =   1065
         Width           =   1590
      End
      Begin VB.OptionButton optSearchArea 
         Caption         =   "Current &Category"
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   5
         Top             =   800
         Width           =   1590
      End
      Begin VB.OptionButton optSearchArea 
         Caption         =   "Current &Template"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   535
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton optSearchArea 
         Caption         =   "Current &Pane"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   3
         Top             =   270
         Width           =   1275
      End
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Direction:"
      Height          =   195
      Index           =   2
      Left            =   1860
      TabIndex        =   15
      Top             =   1050
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Find What:"
      Height          =   270
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   135
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "Replace &With:"
      Height          =   270
      Index           =   1
      Left            =   150
      TabIndex        =   14
      Top             =   555
      Width           =   1080
   End
End
Attribute VB_Name = "frmFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DoFindNext As Boolean
Public DoReplace As Boolean
Public DoReplaceAll As Boolean
Public Canceled As Boolean

Public Enum FindReplaceSearchArea
    SearchAreaCurrentPane = 0
    SearchAreaCurrentTemplate = 1
    SearchAreaCurrentCategory = 2
    SearchAreaCurrentDatabase = 3
End Enum
Public SearchArea As FindReplaceSearchArea

Public Enum FindReplaceDirection
    DirectionAll = 0
    DirectionDown = 1
    DirectionUp = 2
End Enum
Public Direction As FindReplaceDirection

Public FindWholeWordOnly As Boolean
Public MatchCase As Boolean
Public UsePatternMatching As Boolean

Public Property Let IsReplace(New_IsReplace As Boolean)
1        lblLabels(1).Enabled = New_IsReplace
2        txtReplace.Visible = New_IsReplace
3        cmdReplaceAll.Visible = New_IsReplace
4        cmdReplace.Enabled = True                         'New_IsReplace

    'txtReplace.Enabled = New_IsReplace
    'cmdReplace.Enabled = New_IsReplace
    'cmdReplaceAll.Enabled = New_IsReplace

5        Me.Show vbModal
End Property

Public Property Get IsReplace() As Boolean
6        IsReplace = txtReplace.Enabled
End Property

Private Sub cmdCancel_Click()
7        DoFindNext = False
8        DoReplace = False
9        DoReplaceAll = False
10       Canceled = True
11       Me.Hide
End Sub

Private Sub cmdOK_Click()
12       DoFindNext = True
13       DoReplace = False
14       DoReplaceAll = False
15       Canceled = False
16       Me.Hide
End Sub

Private Sub cmdReplace_Click()
17       If cmdReplace.Visible Then
18           DoFindNext = False
19           DoReplace = True
20           DoReplaceAll = False
21           Canceled = False
22           Me.Hide
23       Else
24           DoFindNext = False
25           DoReplace = True
26           DoReplaceAll = False
27           Canceled = False
28           Me.Hide
29       End If

End Sub


Private Sub cmdReplaceAll_Click()
30       DoFindNext = False
31       DoReplace = False
32       DoReplaceAll = True
33       Canceled = False
34       Me.Hide
End Sub


Private Sub Form_Activate()
35       If Len(txtFind) <> 0 And Len(txtReplace) = 0 And txtReplace.Visible Then
36           txtReplace.SetFocus
37       Else
38           txtFind.SetFocus
39       End If
End Sub

Private Sub Form_Load()
40       txtFind = GetSetting$(App.ProductName, "Last", "Find Text", vbNullString)
41       txtReplace = GetSetting$(App.ProductName, "Last", "Replace Text", vbNullString)
42       optSearchArea(GetSetting(App.ProductName, "Last", "Search Area", 0)).Value = True
43       cboDirection.ListIndex = GetSetting(App.ProductName, "Last", "Search Direction", 0)
44       chkFindWholeWordOnly.Value = GetSetting(App.ProductName, "Last", "Find Whole Word Only", 0)
45       chkMatchCase.Value = GetSetting(App.ProductName, "Last", "Match Case", 0)
46       chkUsePatternMatching.Value = GetSetting(App.ProductName, "Last", "Use Pattern Matching", 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
47       SaveSetting App.ProductName, "Last", "Find Text", txtFind.Text
48       SaveSetting App.ProductName, "Last", "Replace Text", txtReplace
49       SaveSetting App.ProductName, "Last", "Search Direction", cboDirection.ListIndex
50       SaveSetting App.ProductName, "Last", "Find Whole Word Only", chkFindWholeWordOnly.Value
51       SaveSetting App.ProductName, "Last", "Match Case", chkMatchCase.Value
52       SaveSetting App.ProductName, "Last", "Use Pattern Matching", chkUsePatternMatching.Value

53       If optSearchArea(0).Value Then
54           SaveSetting App.ProductName, "Last", "Search Area", 0
55       ElseIf optSearchArea(1).Value Then
56           SaveSetting App.ProductName, "Last", "Search Area", 1
57       ElseIf optSearchArea(2).Value Then
58           SaveSetting App.ProductName, "Last", "Search Area", 2
59       Else
60           SaveSetting App.ProductName, "Last", "Search Area", 3
61       End If

End Sub


Private Sub optSearchArea_Click(Index As Integer)
62       If optSearchArea(0).Value Then
63           SearchArea = SearchAreaCurrentPane
64       ElseIf optSearchArea(1).Value Then
65           SearchArea = SearchAreaCurrentTemplate
66       ElseIf optSearchArea(2).Value Then
67           SearchArea = SearchAreaCurrentCategory
68       Else
69           SearchArea = SearchAreaCurrentDatabase
70       End If

End Sub


