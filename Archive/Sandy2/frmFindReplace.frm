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

Public Property Let IReplace(New_IReplace As Boolean)
    lblLabels(1).Enabled = New_IReplace
    txtReplace.Visible = New_IReplace
    cmdReplaceAll.Visible = New_IReplace
    cmdReplace.Enabled = True 'New_IReplace
    
   'txtReplace.Enabled = New_IReplace
   'cmdReplace.Enabled = New_IReplace
   'cmdReplaceAll.Enabled = New_IReplace
    
    Me.Show vbModal
End Property

Public Property Get IReplace() As Boolean
    IReplace = txtReplace.Enabled
End Property

Private Sub cmdCancel_Click()
    DoFindNext = False
    DoReplace = False
    DoReplaceAll = False
    Canceled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    DoFindNext = True
    DoReplace = False
    DoReplaceAll = False
    Canceled = False
    Me.Hide
End Sub

Private Sub cmdReplace_Click()
    If cmdReplace.Visible Then
       DoFindNext = False
       DoReplace = True
       DoReplaceAll = False
       Canceled = False
       Me.Hide
    Else
       DoFindNext = False
       DoReplace = True
       DoReplaceAll = False
       Canceled = False
       Me.Hide
    End If

End Sub


Private Sub cmdReplaceAll_Click()
    DoFindNext = False
    DoReplace = False
    DoReplaceAll = True
    Canceled = False
    Me.Hide
End Sub


Private Sub Form_Activate()
    If txtFind <> vbNullString And txtReplace = vbNullString And txtReplace.Visible Then
       txtReplace.SetFocus
    Else
       txtFind.SetFocus
    End If
End Sub

Private Sub Form_Initialize()

    ' LogEvent "frmFindReplace: Initialize"
End Sub

Private Sub Form_Load()
    txtFind = GetSetting("SliceAndDice", "Last", "Find Text", vbNullString)
    txtReplace = GetSetting("SliceAndDice", "Last", "Replace Text", vbNullString)
    optSearchArea(GetSetting("SliceAndDice", "Last", "Search Area", 0)).Value = True
    cboDirection.ListIndex = GetSetting("SliceAndDice", "Last", "Search Direction", 0)
    chkFindWholeWordOnly.Value = GetSetting("SliceAndDice", "Last", "Find Whole Word Only", 0)
    chkMatchCase.Value = GetSetting("SliceAndDice", "Last", "Match Case", 0)
    chkUsePatternMatching.Value = GetSetting("SliceAndDice", "Last", "Use Pattern Matching", 0)

End Sub


Private Sub Form_Terminate()

    ' LogEvent "frmFindReplace: Terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "SliceAndDice", "Last", "Find Text", txtFind.Text
    SaveSetting "SliceAndDice", "Last", "Replace Text", txtReplace
    SaveSetting "SliceAndDice", "Last", "Search Direction", cboDirection.ListIndex
    SaveSetting "SliceAndDice", "Last", "Find Whole Word Only", chkFindWholeWordOnly.Value
    SaveSetting "SliceAndDice", "Last", "Match Case", chkMatchCase.Value
    SaveSetting "SliceAndDice", "Last", "Use Pattern Matching", chkUsePatternMatching.Value

    If optSearchArea(0).Value Then
       SaveSetting "SliceAndDice", "Last", "Search Area", 0
    ElseIf optSearchArea(1).Value Then
       SaveSetting "SliceAndDice", "Last", "Search Area", 1
    ElseIf optSearchArea(2).Value Then
       SaveSetting "SliceAndDice", "Last", "Search Area", 2
    Else
       SaveSetting "SliceAndDice", "Last", "Search Area", 3
    End If
    
End Sub


Private Sub optSearchArea_Click(Index As Integer)
    If optSearchArea(0).Value Then
       SearchArea = SearchAreaCurrentPane
    ElseIf optSearchArea(1).Value Then
       SearchArea = SearchAreaCurrentTemplate
    ElseIf optSearchArea(2).Value Then
       SearchArea = SearchAreaCurrentCategory
    Else
       SearchArea = SearchAreaCurrentDatabase
    End If

End Sub


