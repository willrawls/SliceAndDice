VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{E60B3BB8-E409-11D2-BA4F-0080C8C222EC}#15.1#0"; "FirmSolutions.ocx"
Begin VB.Form frmMain 
   Caption         =   "Slice and Dice"
   ClientHeight    =   7770
   ClientLeft      =   1890
   ClientTop       =   2925
   ClientWidth     =   11115
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11115
   Begin VB.Frame frmTemplateInfo 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   5250
      TabIndex        =   1
      Top             =   1275
      Visible         =   0   'False
      Width           =   4035
      Begin VB.CommandButton cmdRecalc 
         Caption         =   "Recalc"
         Height          =   435
         Left            =   120
         TabIndex        =   5
         Top             =   75
         Width           =   810
      End
      Begin VB.CheckBox chkAutoRecalc 
         Caption         =   "Auto Recalc when tab selected."
         Height          =   300
         Left            =   990
         TabIndex        =   4
         Top             =   165
         Value           =   1  'Checked
         Width           =   2610
      End
      Begin VB.ListBox lstSoftVariables 
         BackColor       =   &H80000018&
         Height          =   1008
         IntegralHeight  =   0   'False
         Left            =   195
         TabIndex        =   3
         Top             =   840
         Width           =   1830
      End
      Begin VB.ListBox lstSoftCommands 
         BackColor       =   &H80000018&
         Height          =   1008
         IntegralHeight  =   0   'False
         Left            =   2130
         TabIndex        =   2
         Top             =   855
         Width           =   1860
      End
      Begin VB.Label lblTemplateInfo 
         AutoSize        =   -1  'True
         Caption         =   "Soft Variables In Use"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   570
         Width           =   1485
      End
      Begin VB.Label lblTemplateInfo 
         AutoSize        =   -1  'True
         Caption         =   "Soft Commands In Use"
         Height          =   195
         Index           =   1
         Left            =   2145
         TabIndex        =   6
         Top             =   585
         Width           =   1620
      End
   End
   Begin VB.Frame frmOptions 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3705
      Left            =   4380
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame Frame2 
         Caption         =   " Statistics "
         Height          =   1365
         Left            =   70
         TabIndex        =   26
         Top             =   2115
         Width           =   4935
         Begin VB.Label lblDelta 
            AutoSize        =   -1  'True
            Caption         =   "Delta Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   29
            Top             =   930
            Width           =   1110
         End
         Begin VB.Label lblAlpha 
            AutoSize        =   -1  'True
            Caption         =   "Alpha Date: September 15, 2000 12:00:00 PM"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   28
            Top             =   570
            Width           =   4620
         End
         Begin VB.Label lblRevision 
            AutoSize        =   -1  'True
            Caption         =   "Revision # "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   27
            Top             =   210
            Width           =   1125
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Basic "
         Height          =   1965
         Left            =   70
         TabIndex        =   21
         Top             =   0
         Width           =   1725
         Begin VB.CheckBox chkSelected 
            Caption         =   "Selected"
            Enabled         =   0   'False
            Height          =   375
            Left            =   150
            TabIndex        =   25
            ToolTipText     =   "When checked, the Template will be available for direct insertion on the code window's right click menu under ""Insert a Favorite""."
            Top             =   1515
            Width           =   1500
         End
         Begin VB.CheckBox chkFavorite 
            Caption         =   "Favorite"
            Height          =   375
            Left            =   150
            TabIndex        =   24
            ToolTipText     =   "When checked, the Template will be available for direct insertion on the code window's right click menu under ""Insert a Favorite""."
            Top             =   1110
            Width           =   1500
         End
         Begin VB.CheckBox chkUndeletable 
            Caption         =   "Undeletable"
            Height          =   375
            Left            =   150
            TabIndex        =   23
            ToolTipText     =   "When checked, this Template will not allow users to delete  it."
            Top             =   300
            Width           =   1500
         End
         Begin VB.CheckBox chkLocked 
            Caption         =   "Code Locked"
            Height          =   375
            Left            =   150
            TabIndex        =   22
            ToolTipText     =   "When checked, this Template will not allow users to modify its code contents."
            Top             =   705
            Width           =   1500
         End
      End
   End
   Begin FirmSolutions.FSListBar lsbJumpTo 
      Align           =   3  'Align Left
      Height          =   7770
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   13705
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483624
      Arrange         =   1
      LabelEdit       =   1
      View            =   3
   End
   Begin MSComctlLib.ImageList imlTabs 
      Left            =   3660
      Top             =   5940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483638
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "Template Info"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0896
            Key             =   "DocumentAlternate"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CEA
            Key             =   "OptionSet"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1142
            Key             =   "Category"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A2
            Key             =   "Document"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19F6
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E4A
            Key             =   "OptionNotSet"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrDoAction 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   1020
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   4590
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   16
      ToolTipText     =   "Enter the name of the Template here."
      Top             =   30
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   4020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Tag             =   "Code Area 2"
      Top             =   1020
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   1
      Left            =   4020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Tag             =   "Code Area 1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   0
      Left            =   4020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Tag             =   "Code Area 0"
      Top             =   3060
      Width           =   3285
   End
   Begin VB.Frame frmFile 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   4365
      TabIndex        =   9
      Top             =   990
      Visible         =   0   'False
      Width           =   4035
      Begin VB.TextBox txtCodeToFile 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   100
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Tag             =   "Code Area File"
         Top             =   750
         Width           =   3255
      End
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   100
         MaxLength       =   255
         TabIndex        =   10
         ToolTipText     =   "This can include template variables."
         Top             =   240
         Width           =   3345
      End
      Begin VB.Label Label2 
         Caption         =   "Filename to send output to:"
         Height          =   345
         Left            =   100
         TabIndex        =   12
         Top             =   30
         Width           =   2865
      End
   End
   Begin VB.Timer tmrActivateDBClassGen 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   540
   End
   Begin VB.TextBox txtShortName 
      Height          =   300
      Left            =   4590
      TabIndex        =   0
      ToolTipText     =   "Enter the name of the Snippet here."
      Top             =   30
      Width           =   4140
   End
   Begin MSComDlg.CommonDialog cdgSelect 
      Left            =   4380
      Top             =   6210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".mdb"
      DialogTitle     =   "Select Access97 DB to work on"
      Filter          =   "*.mdb"
   End
   Begin MSComctlLib.TabStrip tabCode 
      Height          =   5448
      Left            =   3528
      TabIndex        =   19
      Top             =   420
      Width           =   7512
      _ExtentX        =   13256
      _ExtentY        =   9604
      HotTracking     =   -1  'True
      ImageList       =   "imlTabs"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "(Declarations)"
            ImageVarType    =   8
            ImageKey        =   "Document"
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "At Cursor"
            ImageVarType    =   8
            ImageKey        =   "Document"
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "At Bottom"
            ImageVarType    =   8
            ImageKey        =   "Document"
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "In a file"
            ImageVarType    =   8
            ImageKey        =   "Document"
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            ImageVarType    =   8
            ImageKey        =   "OptionNotSet"
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Template Info"
            ImageVarType    =   8
            ImageKey        =   "Template Info"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCode 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name:"
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
      Index           =   3
      Left            =   4005
      TabIndex        =   18
      Top             =   75
      Width           =   510
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   345
      Left            =   4500
      TabIndex        =   17
      Top             =   960
      Width           =   1155
   End
   Begin VB.Menu mnuX 
      Caption         =   "X"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSpecialOpenDatabase 
         Caption         =   "&Open Slice and Dice database"
      End
      Begin VB.Menu mnuSpecialNewDatabase 
         Caption         =   "&New Slice and Dice database"
      End
      Begin VB.Menu mnuSpecialExportSnippet 
         Caption         =   "Export current Template"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileList 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "&Replace"
      End
   End
   Begin VB.Menu mnuTemplate 
      Caption         =   "&Template"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileCopy 
         Caption         =   "&Copy current"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIsFavorite 
         Caption         =   "On &Favorites Menu"
      End
      Begin VB.Menu mnuIsUndeletable 
         Caption         =   "Cannot be &Deleted"
      End
      Begin VB.Menu mnuIsCodeLocked 
         Caption         =   "Cannot be &Edited"
      End
      Begin VB.Menu mnuSep25 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertTemplate 
         Caption         =   "&Insert Template into VB"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "I&mport selected code from VB"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuCategories 
      Caption         =   "&Categories"
      Begin VB.Menu mnuCategoriesNewMethod 
         Caption         =   "&New Category"
         Index           =   0
      End
      Begin VB.Menu mnuCategoriesNewMethod 
         Caption         =   "&Duplicate a Category. Template names and code"
         Index           =   1
      End
      Begin VB.Menu mnuCategoriesNewMethod 
         Caption         =   "Duplicate a Category. Template names only"
         Index           =   2
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRefresh 
         Caption         =   "&Refresh Category and Template List"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCategoriesDeleteCurrent 
         Caption         =   "Delete current Category"
      End
   End
   Begin VB.Menu mnuSpecial 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuSpecialViewLog 
         Caption         =   "&View Insertion Log"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuProjectProcessor 
         Caption         =   "&Project Processor (Future)"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuSep7 
      Caption         =   "&Options"
      Begin VB.Menu mnuExitAfterInsert 
         Caption         =   "Exit after insert ?"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowSplash 
         Caption         =   "Show splash screen at startup"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowPaintbrushIcon 
         Caption         =   "Show Paintbrush icon on ""Standard"" menu"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowOnModuleRightClick 
         Caption         =   "Show ""Slice and Dice"" on Module right click"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSwitchTabsAutomatically 
         Caption         =   "Switch to first tab with code when switching templates"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPasswordProtection 
         Caption         =   "Password Protection"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOLEDragDrop 
         Caption         =   "Use OLE Text Editing - Drag && Drop"
      End
      Begin VB.Menu mnuTakeOverKeys 
         Caption         =   "Take over CTRL-SHIFT-1234567890"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeBackgroundColors 
         Caption         =   "Change Background Colors"
      End
      Begin VB.Menu mnuChangeForegroundColor 
         Caption         =   "Change Foreground Color"
      End
   End
   Begin VB.Menu mnuHistory 
      Caption         =   "&History"
      Begin VB.Menu mnuBack 
         Caption         =   "&Back        (Alt+Left Arrow)"
      End
      Begin VB.Menu mnuForward 
         Caption         =   "&Forward    (Alt+Right Arrow)"
      End
      Begin VB.Menu mnuHistorySep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHistoryList 
         Caption         =   "&List"
      End
   End
   Begin VB.Menu mnuFav 
      Caption         =   "Fa&vorites"
      Begin VB.Menu mnuFavorite 
         Caption         =   "-Empty-"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuExternalFunctions 
      Caption         =   "Exte&rnals"
      Begin VB.Menu mnuDBClassGen 
         Caption         =   "&Database to Code Generator"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExternals 
         Caption         =   "-Empty-"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Hel&p"
      Begin VB.Menu mnuHelpSoftCommandReference 
         Caption         =   "Soft &Command Reference"
      End
      Begin VB.Menu mnuHelpCodeGenSoftVarRef 
         Caption         =   "Code &Gen Soft Variable Reference"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpOnlineDocumentation 
         Caption         =   "Online &Documentation"
      End
      Begin VB.Menu mnuHelpReportIssue 
         Caption         =   "Report an &Issue"
      End
      Begin VB.Menu mnuHelpEmailWilliamRawls 
         Caption         =   "&Email William Rawls"
      End
      Begin VB.Menu mnuHelpVisitHomePage 
         Caption         =   "Visit the &Home Page"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Co&ntents"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "&Index"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpSep0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ================================================================================
' Class Module      frmMain
'
' Filename          frmMain.dob
'
' Author            William M. Rawls
'
' Created On        9/3/1997 8:00 pm
'
' Description
'
' The real workhorse of this Add-in.
'
' Revisions
'
' <Date>, William M. Rawls
' <Description of Revision>
'
' ================================================================================
Option Explicit

Private m_sTemplateDatabaseName         As String
Private m_sCurrentEventResponseCategory As String

Public CurrentCodeArea                  As Integer
Public Parent                           As Wizard

Private m_oDBClassGen                   As frmDBClassGen
Private m_asaHistory                    As New CAssocArray
Private m_asaAttributes                 As New CAssocArray

Public SliceAndDice                     As CSliceAndDice
Public CurrentTemplate                  As CTemplate
Public InternalCurrentTemplate          As CTemplate

Public Complete                         As CSadCommands

Private SadCommands()                   As ISadAddin
Public SadCommandSetCount               As Long
Private FavoriteCount                   As Long
Private ExternalCount                   As Long
Private CurrentHistoryEntry             As String

Private ActionToDo                      As String
Private ActionParam                     As String

Private mbScramFormKey                  As Boolean
Private mbFillingAddInScreen            As Boolean
Private mbIgnoreBlanks                  As Boolean
Private mbIgnoreReadOnly                As Boolean
Public OkayToDoAction                   As Boolean
Public FavoriteCalledFromIDE            As Boolean

Public OkayToUnload                     As Boolean

Public WithEvents mHotKeyOpenWindow     As cRegHotKey
Attribute mHotKeyOpenWindow.VB_VarHelpID = -1

Public Function AddSadCommandSet(ByRef oCommands As ISadAddin) As Boolean
1        On Error Resume Next
2        Dim Externals      As CAssocArray
3        Dim CurrExternal   As CAssocItem
 
4        Err.Clear
5        SadCommandSetCount = SadCommandSetCount + 1
6        ReDim Preserve SadCommands(1 To SadCommandSetCount)
7        Set SadCommands(SadCommandSetCount) = oCommands
8        If SadCommands(SadCommandSetCount).Startup(Parent, Parent.vbInst) Then
9            AddSadCommandSet = True
10           frmSplash.lblDLLsLoaded(1).Caption = vbNullString & SadCommandSetCount
11           frmSplash.lblDLLsLoaded(1).Refresh
12           Set Externals = oCommands.Externals
13           If Not Externals Is Nothing Then
14               If ExternalCount > 0 Then
15                   Load mnuExternals(ExternalCount)
16                   mnuExternals(ExternalCount).Caption = "-"
17                   mnuExternals(ExternalCount).Tag = vbNullString
18                   mnuExternals(ExternalCount).Enabled = True
19                   mnuExternals(ExternalCount).Visible = True
20                   ExternalCount = ExternalCount + 1
21               End If
22               For Each CurrExternal In Externals
23                   If ExternalCount > 0 Then
24                       Load mnuExternals(ExternalCount)
25                   End If
26                   mnuExternals(ExternalCount).Caption = CurrExternal.Key
27                   mnuExternals(ExternalCount).Tag = SadCommandSetCount & "|" & CurrExternal.Value
28                   mnuExternals(ExternalCount).Enabled = True
29                   mnuExternals(ExternalCount).Visible = True
30                   ExternalCount = ExternalCount + 1
31               Next CurrExternal
32           End If
33           Set Externals = Nothing
34           DoEvents
35       End If
36       Err.Clear
End Function

Public Function ChangeFocusOfInsertion(ByRef II As CInsertionInfo, ByVal sNewFocusOfInsertion As String) As Boolean
    Dim fh       As Long
    Dim vLines   As Variant
    Dim CurrLine As Long

  ' "Flush" current focus buffer if there is one
    If Len(II.ExternalFilename) > 0 And Len(II.TextToSendToFile) > 0 Then
       If Left$(II.ExternalFilename, 10) = "**KEYBOARD" Then
          If Right$(II.ExternalFilename, 5) = "RAW**" Then
             SendKeys Replace$(II.TextToSendToFile, vbNewLine, vbNullString)
          Else
             SendKeys Replace$(II.TextToSendToFile, vbNewLine, "{ENTER}")
          End If

       ElseIf II.ExternalFilename = "**CLIPBOARD**" Then
          StringToClipboard II.TextToSendToFile

       ElseIf II.ExternalFilename = "**MESSAGEWINDOW**" Then
          ShowMessage II.TextToSendToFile, , vbNullString

       ElseIf Left$(II.ExternalFilename, 16) = "**SOFTVARIABLE**" Then
          II.SoftVars(Mid$(II.ExternalFilename, 18)).Value = II.TextToSendToFile

       ElseIf II.ExternalFilename = "**SOFTCODE**" Then
          vLines = Split(II.TextToSendToFile, vbNewLine)
          For CurrLine = 0 To UBound(vLines)
              If Len(Trim$(vLines(CurrLine))) Then
                 vLines(CurrLine) = gsSoftCmdDelimiter & Trim$(vLines(CurrLine))
              End If
          Next CurrLine
          II.LinesLeftToProcess = Join$(vLines, vbNewLine)

       ElseIf Left$(II.ExternalFilename, 13) = "**OVERWRITE**" Then
          On Error Resume Next
          fh = FreeFile
          Open Trim$(Mid$(II.ExternalFilename, 15)) For Output Access Write As #fh
               Print #fh, II.TextToSendToFile
          Close #fh
          If Err.Number <> 0 Then
             MsgBox "ChangeFocusOfInsertion" & vbNewLine & vbTab & "Error overwriting file:" & vbNewLine & vbTab & vbTab & Trim$(Mid$(II.ExternalFilename, 15))
             gbCancelInsertion = True
             Err.Clear
          End If

       Else
          On Error Resume Next
          fh = FreeFile
          Open II.ExternalFilename For Append Access Write As #fh
               Print #fh, II.TextToSendToFile
          Close #fh
          If Err.Number <> 0 Then
             MsgBox "ChangeFocusOfInsertion" & vbNewLine & vbTab & "Error appending to file:" & vbNewLine & vbTab & vbTab & Trim$(Mid$(II.ExternalFilename, 15))
             gbCancelInsertion = True
             Err.Clear
          End If
       End If

       II.TextToSendToFile = vbNullString
    End If

    II.ExternalFilename = sNewFocusOfInsertion
    ChangeFocusOfInsertion = True
End Function

Public Property Let CurrentEventResponseCategory(sNewCategory As String)
37       m_sCurrentEventResponseCategory = sNewCategory
End Property

Public Property Get CurrentEventResponseCategory() As String
38       If Len(m_sCurrentEventResponseCategory) = 0 Then m_sCurrentEventResponseCategory = "Event Response"
39       CurrentEventResponseCategory = m_sCurrentEventResponseCategory
End Property

Public Property Get CurrentTemplateNameAndCategory() As String
40       CurrentTemplateNameAndCategory = txtName.Text
End Property

Public Property Set DBClassGen(New_DBClassGen As frmDBClassGen)
41       On Error Resume Next
42       Set m_oDBClassGen = New_DBClassGen
43       SetColors GetSetting$(App.ProductName, "Last", "Background Color", "&H80000018&"), GetSetting$(App.ProductName, "Last", "Foreground Color", "&H80000008&")
End Property

Public Property Get DBClassGen() As frmDBClassGen
44       Set DBClassGen = m_oDBClassGen
End Property

Public Sub DeleteTemplate(Optional ByVal bAutoDelete As Boolean = False)
45       On Error GoTo EH_frmMain_DeleteTemplate
46       Static bInHereAlready As Boolean
47       If bInHereAlready Then Exit Sub
48       bInHereAlready = True

49       If CurrentTemplate Is Nothing Then
50           If bAutoDelete Then
51               MsgBox "DeleteTemplate failed because nothing is selected."
52           Else
53               MsgBox "Please select a " & gsTemplate & " to delete first."
54           End If
55           bInHereAlready = False
56           Exit Sub
57       ElseIf chkUndeletable.Value <> 0 Then
58           MsgBox "This " & gsTemplate & " cannot be deleted (undeletable turned on). Turn off before continuing."
59           bInHereAlready = False
60           Exit Sub
61       End If

62       If Not bAutoDelete Then
63           If Not bUserSure("This will permanently delete the " & gsTemplate & gsS & gsQ & CurrentTemplate.Key & gsQ & gsP & gs2EOLTab & "Are you sure this is what you want to do ?") Then
64               bInHereAlready = False
65               Exit Sub
66           End If
67       End If

68       CurrentTemplate.Deleted = True
69       CurrentTemplate.Modified = True
70       SaveTemplate
71       RefillList
         If Not SliceAndDice(CurrentTemplate.ParentKey) Is Nothing Then
            If Not SliceAndDice(CurrentTemplate.ParentKey).Templates(1) Is Nothing Then
               JumpTo SliceAndDice(CurrentTemplate.ParentKey).Templates(1).Key
               lsbJumpTo.BarAndItem SliceAndDice(CurrentTemplate.ParentKey).Key, SliceAndDice(CurrentTemplate.ParentKey).Templates(1).ShortTemplateName
            Else
               GoTo FirstCategoryChecker
            End If
72       ElseIf Not SliceAndDice(1) Is Nothing Then
FirstCategoryChecker:
73           If Not SliceAndDice(1).Templates(1) Is Nothing Then
74               JumpTo SliceAndDice(1).Templates(1).Key
75               lsbJumpTo.BarAndItem SliceAndDice(1).Key, SliceAndDice(1).Templates(1).ShortTemplateName
76           ElseIf Not SliceAndDice(2) Is Nothing Then
77               If Not SliceAndDice(2).Templates(1) Is Nothing Then
78                   JumpTo SliceAndDice(2).Templates(1).Key
79                   lsbJumpTo.BarAndItem SliceAndDice(2).Key, SliceAndDice(2).Templates(1).ShortTemplateName
80               End If
81           End If
82       End If

83 EH_frmMain_DeleteTemplate_Continue:
84       bInHereAlready = False
85       Exit Sub

86 EH_frmMain_DeleteTemplate:
87       MsgBox "Error occured in:" & gsEolTab & "Module: frmMain" & gsEolTab & "Procedure: DeleteTemplate" & gs2EOL & Err.Description

88       Resume EH_frmMain_DeleteTemplate_Continue

89       Resume
End Sub

Public Function Evaluate(ByVal sExpression As String, ByRef asaVar As CAssocArray) As String
90       On Error Resume Next
91       Dim sResult As String
92       Dim sOperator As String
93       Dim sLeftOfOp As String
94       Dim sRightOfOp As String

95       Dim vLeftOfOp As Variant
96       Dim vRightOfOp As Variant

97       If InStr(UCase$(sExpression), " AND ") Then
98           vLeftOfOp = InStr(UCase$(sExpression), " AND ")
99           sLeftOfOp = Left$(sExpression, vLeftOfOp - 1)
100          sRightOfOp = Mid$(sExpression, vLeftOfOp + 5)
101          sLeftOfOp = Evaluate(Trim$(sLeftOfOp), asaVar)
102          sRightOfOp = Evaluate(Trim$(sRightOfOp), asaVar)
103          Evaluate = IIf((Val(sLeftOfOp) <> 0) And (Val(sRightOfOp) <> 0), "1", "0")
104          Exit Function

105      ElseIf InStr(UCase$(sExpression), " OR ") Then
106          vLeftOfOp = InStr(UCase$(sExpression), " OR ")
107          sLeftOfOp = Left$(sExpression, vLeftOfOp - 1)
108          sRightOfOp = Mid$(sExpression, vLeftOfOp + 5)
109          sLeftOfOp = Evaluate(Trim$(sLeftOfOp), asaVar)
110          sRightOfOp = Evaluate(Trim$(sRightOfOp), asaVar)
111          Evaluate = IIf((Val(sLeftOfOp) <> 0) Or (Val(sRightOfOp) <> 0), "1", "0")
112          Exit Function

113      ElseIf InStr(UCase$(sExpression), " XOR ") Then
114          vLeftOfOp = InStr(UCase$(sExpression), " XOR ")
115          sLeftOfOp = Left$(sExpression, vLeftOfOp - 1)
116          sRightOfOp = Mid$(sExpression, vLeftOfOp + 5)
117          sLeftOfOp = Evaluate(Trim$(sLeftOfOp), asaVar)
118          sRightOfOp = Evaluate(Trim$(sRightOfOp), asaVar)
119          Evaluate = IIf((Val(sLeftOfOp) <> 0) Xor (Val(sRightOfOp) <> 0), "1", "0")
120          Exit Function

121      End If

122      If Len(sExpression) > 0 Then
123          If InStr(sExpression, "+") > 0 Then
124              sOperator = "+"
125          ElseIf InStr(sExpression, "-") > 0 Then
126              sOperator = "-"
127          ElseIf InStr(sExpression, "*") > 0 Then
128              sOperator = "*"
129          ElseIf InStr(sExpression, "/") > 0 Then
130              sOperator = "/"
131          ElseIf InStr(sExpression, gsBS) > 0 Then
132              sOperator = gsBS
133          ElseIf InStr(sExpression, "^") > 0 Then
134              sOperator = "^"
135          ElseIf InStr(sExpression, " MOD ") > 0 Then
136              sOperator = " MOD "
137          ElseIf InStr(sExpression, "<=") > 0 Then
138              sOperator = "<="
139          ElseIf InStr(sExpression, ">=") > 0 Then
140              sOperator = ">="
141          ElseIf InStr(sExpression, "<>") > 0 Or InStr(sExpression, "!=") > 0 Then
142              sOperator = "<>"
143          ElseIf InStr(sExpression, "<") > 0 Then
144              sOperator = "<"
145          ElseIf InStr(sExpression, ">") > 0 Then
146              sOperator = ">"
147          ElseIf InStr(sExpression, "==") > 0 Then
148              sOperator = gsE
149          ElseIf InStr(sExpression, gsE) > 0 Then
150              sOperator = gsE
151          Else
152              sOperator = vbNullString
153          End If

154          If Len(sOperator) > 0 Then
155              sLeftOfOp = Trim$(sGetToken(sExpression, 1, sOperator))
156              If Len(asaVar(sLeftOfOp)) = 0 And Len(sLeftOfOp) > 0 Then
157                  asaVar(sLeftOfOp) = sLeftOfOp
158              End If
159              sRightOfOp = Trim$(sGetToken(sExpression, 2, sOperator))
160              If Len(asaVar(sRightOfOp)) = 0 And Len(sRightOfOp) > 0 Then
161                  asaVar(sRightOfOp) = sRightOfOp
162              End If

163              If IsNumeric(asaVar(sLeftOfOp)) Then
                'If Val(asaVar(sLeftOfOp)) <> 0 Then
164                  vLeftOfOp = Val(asaVar(sLeftOfOp))
165              Else
166                  vLeftOfOp = asaVar(sLeftOfOp)
167              End If

168              If IsNumeric(asaVar(sRightOfOp)) Then
                'If Val(asaVar(sRightOfOp)) <> 0 Then
169                  vRightOfOp = Val(asaVar(sRightOfOp))
170              Else
171                  vRightOfOp = asaVar(sRightOfOp)
172              End If

173              If vRightOfOp = "0" Then
                Select Case sOperator
                    Case "/", gsBS, " MOD ", "*": sResult = "0"
174                      Case "<": sResult = IIf(vLeftOfOp < 0, "1", "0")
175                      Case "<=": sResult = IIf(vLeftOfOp <= 0, "1", "0")
176                      Case ">": sResult = IIf(vLeftOfOp > 0, "1", "0")
177                      Case ">=": sResult = IIf(vLeftOfOp >= 0, "1", "0")
178                      Case "<>": sResult = IIf(vLeftOfOp <> 0, "1", "0")
179                      Case gsE: sResult = IIf(vLeftOfOp = 0, "1", "0")
180                      Case Else: sResult = vLeftOfOp
181                  End Select
182              Else
                Select Case sOperator
                    Case "+": sResult = vLeftOfOp + vRightOfOp
183                      Case "-": sResult = vLeftOfOp - vRightOfOp
184                      Case "*": sResult = vLeftOfOp * vRightOfOp
185                      Case "/": sResult = vLeftOfOp / vRightOfOp
186                      Case gsBS: sResult = vLeftOfOp \ vRightOfOp
187                      Case "^": sResult = vLeftOfOp ^ vRightOfOp
188                      Case " MOD ": sResult = vLeftOfOp Mod vRightOfOp
189                      Case "<": sResult = IIf(vLeftOfOp < vRightOfOp, "1", "0")
190                      Case "<=": sResult = IIf(vLeftOfOp <= vRightOfOp, "1", "0")
191                      Case ">": sResult = IIf(vLeftOfOp > vRightOfOp, "1", "0")
192                      Case ">=": sResult = IIf(vLeftOfOp >= vRightOfOp, "1", "0")
193                      Case "<>": sResult = IIf(vLeftOfOp <> vRightOfOp, "1", "0")
194                      Case gsE: sResult = IIf(vLeftOfOp = vRightOfOp, "1", "0")
195                      Case Else: sResult = vLeftOfOp
196                  End Select
197              End If
198          Else
199              If Len(asaVar(sExpression)) > 0 Then
200                  sResult = asaVar(sExpression)
201              Else
202                  sResult = sExpression
203              End If
204          End If
205      End If

206      Evaluate = sResult
End Function

Public Property Get ExitAfterInsert() As Boolean
207      ExitAfterInsert = mnuExitAfterInsert.Checked
End Property

Public Sub FillAddInScreen()
On Error GoTo EH_frmMain_FillAddInScreen
209      Static bInHereAlready As Boolean

         If CurrentTemplate Is Nothing Then Exit Sub
210      If bInHereAlready Then Exit Sub

211      bInHereAlready = True
212      mbFillingAddInScreen = True

213      With CurrentTemplate
214          txtName = .Key
215          txtShortName = .ShortTemplateName

216          txtCode(0) = .memoCodeAtTop
217          txtCode(1) = .memoCodeAtCursor
218          txtCode(2) = .memoCodeAtBottom

219          txtFilename = .FileName
220          txtCodeToFile = .memoCodeToFile

221          chkUndeletable = Abs(.Undeletable)
224          mnuIsUndeletable.Checked = .Undeletable

             chkLocked = Abs(.Locked)
             mnuIsCodeLocked.Checked = .Locked

             chkFavorite = Abs(.Favorite)
             mnuIsFavorite.Checked = .Favorite

225          chkSelected = Abs(.Selected)

226          lblRevision.Caption = "Revision #: " & .Revision

227          lblAlpha.Caption = "Alpha Date: " & Format$(.DateCreated, "Mmmm D, YYYY H:NN:SS AM/PM")
228          lblDelta.Caption = "Delta Date: " & Format$(.DateModified, "Mmmm D, YYYY H:NN:SS AM/PM")
229      End With

230 EH_frmMain_FillAddInScreen_Continue:
231      bInHereAlready = False
232      mbFillingAddInScreen = False
233      Exit Sub

234 EH_frmMain_FillAddInScreen:
235      MsgBox "Error occured in:" & gsEolTab & "Module: frmMain" & gsEolTab & "Procedure: FillAddInScreen" & gs2EOL & Err.Description

236      Resume EH_frmMain_FillAddInScreen_Continue

237      Resume
End Sub

Public Sub GetCategoryAndName(ByVal sCategoryAndName As String, ByRef sCategory As String, ByRef sShortName As String)
238      If lTokenCount(sCategoryAndName, gsCategoryTemplateDelimiter) < 2 Then
239          sCategory = "Unknown"
240          sShortName = sCategoryAndName
241      Else
242          sCategory = sGetToken(sCategoryAndName, 1, gsCategoryTemplateDelimiter)
243          sShortName = sAfter(sCategoryAndName, 1, gsCategoryTemplateDelimiter)
244          If Len(sShortName) = 0 Then
245              sCategory = "Unknown"
246              sShortName = sCategoryAndName
247          End If
248      End If
End Sub

Public Sub HandleIDEEvents(ByVal sTemplateName As String, Optional ByVal VBProject As VBIDE.VBProject, Optional ByVal VBComponent As VBIDE.VBComponent)
'On Error GoTo EH_frmMain_HandleIDEEvents
'    Static bInHereAlready As Boolean
'    If bInHereAlready Then Exit Sub
'    bInHereAlready = True
'
'On Error Resume Next
'
'  ' These features currently disabled
'    Exit Sub
'
'    If SliceAndDice(CurrentEventResponseCategory).Templates(sTemplateName) Is Nothing Then Exit Sub
'
'    m_asaIDEEvents.Clear
'    If Not VBProject Is Nothing Then
'       With VBProject
'            m_asaIDEEvents("Project Name") = .Name
'            m_asaIDEEvents("Project Filename") = .FileName
'            m_asaIDEEvents("Project Build Filename") = .BuildFileName
'            m_asaIDEEvents("Project Description") = .Description
'            Select Case .Type
'                   Case vbext_pt_StandardExe:    m_asaIDEEvents("Project Type") = "Standard EXE"
'                   Case vbext_pt_ActiveXExe:     m_asaIDEEvents("Project Type") = "ActiveX EXE"
'                   Case vbext_pt_ActiveXDll:     m_asaIDEEvents("Project Type") = "ActiveX DLL"
'                   Case vbext_pt_ActiveXControl: m_asaIDEEvents("Project Type") = "ActiveX Control"
'                   Case Else:                    m_asaIDEEvents("Project Type") = "Unknown"
'            End Select
'       End With
'    End If
'
'    If Not VBComponent Is Nothing Then
'       With VBComponent
'            m_asaIDEEvents("Component Name") = .Name
'            m_asaIDEEvents("Component Description") = .Description
'            Select Case .Type
'                   Case vbext_ct_ClassModule:    m_asaIDEEvents("Component Type") = "Class"
'                   Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm:
'                                                 m_asaIDEEvents("Component Type") = "Form"
'                   Case vbext_ct_StdModule:      m_asaIDEEvents("Component Type") = "Module"
'                   Case Else:                    m_asaIDEEvents("Component Type") = "Other"
'            End Select
'       End With
'    End If
'
'    DoInsertion m_asaIDEEvents, CurrentEventResponseCategory & gsCategoryTemplateDelimiter & sTemplateName
'
'    m_asaIDEEvents.Clear
'    Set m_asaIDEEvents = Nothing
'
'EH_frmMain_HandleIDEEvents_Continue:
'    bInHereAlready = False
'    Exit Sub
'
'EH_frmMain_HandleIDEEvents:
'    MsgBox "Error occured in:" & gsEolTab & "Module: frmMain" & gsEolTab & "Procedure: HandleIDEEvents" & gs2EOL & Err.Description
'
'    Resume EH_frmMain_HandleIDEEvents_Continue
'
'    Resume
End Sub




Public Sub HideAllWindows(Optional ByVal bUnloadAsWell = False)
249      On Error Resume Next
250      Dim CurrSet As Long

251      If Not m_oDBClassGen Is Nothing Then
252          m_oDBClassGen.Hide
253      End If

254      If SadCommandSetCount > 0 Then
255          If SadCommandSetCount = 1 Then
256              SadCommands(1).CommandSet.HideWindow bUnloadAsWell
257              Exit Sub
258          Else
259              For CurrSet = 1 To SadCommandSetCount
260                  SadCommands(CurrSet).CommandSet.HideWindow bUnloadAsWell
261                  SadCommands(CurrSet).ExecuteExternal "HIDE ALL WINDOWS", "HIDE ALL WINDOWS"
262              Next CurrSet
263          End If
264      End If
End Sub

Public Function InitializeAddinDLLs(ByVal sAddinList As String) As Boolean
265      Dim asaTemp As CAssocArray
266      Dim CurrAssocItem As CAssocItem
267      Dim CurrDLL As ISadAddin

268      ShutdownDLLs

269      If Len(sAddinList) = 0 Then
270          InitializeAddinDLLs = True
271          Exit Function
272      End If

273      Set asaTemp = New CAssocArray
274      asaTemp.All = sAddinList
275      For Each CurrAssocItem In asaTemp
276          If StrComp(UCase$(Trim$(CurrAssocItem.Value)), "LOAD") = 0 Then
277              On Error Resume Next
278              Err.Clear
279              Set CurrDLL = CreateObject(Trim$(CurrAssocItem.Key))
280              If Err.Number = 0 Then
281                  If Not AddSadCommandSet(CurrDLL) Then
282                      CurrAssocItem.Value = "Error in 'AddSadCommandSet'"
283                  Else
284                      SadCommands(SadCommandSetCount).CommandSet.Attributes("Name").Value = Trim$(CurrAssocItem.Key)
285                      If CurrDLL = "sadRegister.NewCommands" Then
286                          SadCommands(SadCommandSetCount).CommandSet.Attributes("Registered").Value = IIf(frmSplash.DetermineRegistration, "True", "False")
287                      End If
288                  End If
289              Else
290                  MsgBox "Failed to create the SAD Addin object: " & gsEolTab & "Name:" & Trim$(CurrAssocItem.Key) & gsEolTab & "Err #" & Err.Number & ": " & Err.Description
291              End If
292              Err.Clear
293          End If
294      Next CurrAssocItem

         UpdateCompleteListOfSoftCommands

    ' Future: Store results of loads back for next time.
295      sAddinList = asaTemp.All

296      Set asaTemp = Nothing
End Function

Public Sub NewTemplate(Optional ByVal bAutoCreate As Boolean = False, Optional ByVal sTitle As String, Optional ByVal sDefaultShortName As String, Optional ByVal bJumpToAfterCreate As Boolean = True)
297      On Error GoTo EH_frmMain_NewTemplate
298      Dim sCategory As String
299      Dim sShortName As String

300      If Len(sTitle) = 0 Then
301          sCategory = lsbJumpTo.BarKey
302          If Len(sDefaultShortName) = 0 Then
303              sDefaultShortName = Abs(NextNegativeUnique())
304          End If
305          sTitle = InputBox("What should the name of this " & gsTemplate & " be ?" & gsEolTab & "(Blank to cancel)" & gs2EOL & "Format of name MUST be:" & gsEolTab & gsCategory & " Name - " & gsTemplate & " Name", "NEW " & gsTemplate, sCategory & gsCategoryTemplateDelimiter & sDefaultShortName)
306      End If
307      If Len(sTitle) = 0 Then Exit Sub

308      GetCategoryAndName sTitle, sCategory, sShortName
309      If Len(sCategory) = 0 Or Len(sShortName) = 0 Then
310          MsgBox "New " & gsTemplate & " name must be in the form: " & gsEolTab & "<CategoryName> & ' - ' & <ShortTemplateName>"
311          Exit Sub
312      End If

313      If SliceAndDice(sCategory) Is Nothing Then
314          If Not bAutoCreate Then
315              If Not bUserSure("The " & gsCategory & " '" & sCategory & "' does not exist. Would you like to create it ?") Then
316                  Exit Sub
317              End If
318          End If
319          SliceAndDice.Categorys.Add sCategory
320      ElseIf Not (SliceAndDice(sCategory).Templates(sShortName) Is Nothing) Then
321          MsgBox "There is a " & gsTemplate & " by that name in that " & gsCategory & " already.", vbInformation
322          Exit Sub
323      ElseIf Not SliceAndDice(sCategory).Templates(sTitle) Is Nothing Then
324          MsgBox "There is a " & gsTemplate & " by that name in that " & gsCategory & " already.", vbInformation
325          Exit Sub
326      End If

327      SaveTemplate

328      With SliceAndDice(sCategory).Templates.Add(sTitle)
329          .ShortTemplateName = sShortName
330          .ParentKey = sCategory
331          .OriginalShortName = sShortName
332      End With

333      SliceAndDice.Save

334      RefillList

335      If bJumpToAfterCreate Then
336          JumpTo sTitle, False, True

337          txtName.Text = sTitle
338          txtShortName.Text = sShortName
339      End If

340 EH_frmMain_NewTemplate_Continue:
341      Exit Sub

342 EH_frmMain_NewTemplate:
343      MsgBox "Error occured in:" & gsEolTab & "Module: frmMain" & gsEolTab & "Procedure: NewTemplate" & gs2EOL & Err.Description

344      Resume EH_frmMain_NewTemplate_Continue

345      Resume
End Sub

Public Sub QueueAction(ByVal sAction As String, Optional ByVal sParam As String, Optional ByVal Interval As Integer = 150)
346      OkayToDoAction = False
347      ActionToDo = sAction
348      ActionParam = sParam
349      tmrDoAction.Interval = IIf(Interval > 65535, 65535, IIf(Interval < 100, 100, Interval))
350      tmrDoAction.Enabled = True
End Sub

Public Property Let QueuedInsertions(New_QueuedInsertions As String)
351      On Error GoTo EH_frmMain_QueuedInsertions
352      Static bInHereAlready As Boolean
353      If bInHereAlready Then Exit Property
354      bInHereAlready = True

355      Dim asaVar As New CAssocArray
356      Dim asaV As New CAssocArray
357      Dim CurItem As CAssocItem

358      asaVar.ItemDelimiter = "~"
359      asaVar.All = New_QueuedInsertions
360      For Each CurItem In asaVar.mCol
361          DoInsertion asaV, CurItem.Key
362          If gbCancelInsertion Then Exit Property
363      Next CurItem

364 EH_frmMain_QueuedInsertions_Continue:
365      bInHereAlready = False
366      Exit Property

367 EH_frmMain_QueuedInsertions:
368      MsgBox "Error occured in:" & gsEolTab & "Module: frmMain" & gsEolTab & "Procedure: QueuedInsertions" & gs2EOL & Err.Description

369      Resume EH_frmMain_QueuedInsertions_Continue

370      Resume
End Property

Public Function RefreshDatabaseConnection() As Boolean
371      On Error GoTo EH_frmMain_RefreshDatabaseConnection

372      Call NextNegativeUnique

373      Set CurrentTemplate = Nothing
374      Set InternalCurrentTemplate = Nothing
375      Set SliceAndDice = Nothing

376      Set SliceAndDice = New CSliceAndDice
377      If Not SliceAndDice.Load(m_sTemplateDatabaseName) Then
378          RefreshDatabaseConnection = False
379          Exit Function
380      End If

381      RefillList
382      On Error Resume Next
383      lsbJumpTo.HideCategories
    'lsbJumpTo.DisplayCategories

384      Caption = gsSliceAndDice & gsCategoryTemplateDelimiter & m_sTemplateDatabaseName
385      RefreshDatabaseConnection = True

386 EH_frmMain_RefreshDatabaseConnection_Continue:
387      Exit Function

388 EH_frmMain_RefreshDatabaseConnection:
389      LogError "frmMain", "RefreshDatabaseConnection", Err.Number, Err.Description, Erl
390      RefreshDatabaseConnection = False
391      Resume EH_frmMain_RefreshDatabaseConnection_Continue

392      Resume
End Function

Public Sub DoInsertion(asaV As CAssocArray, sTemplateToInsert As String, Optional ByVal bSkipDeclarations As Boolean = False)
393      On Error GoTo EH_frmMain_DoInsertion
394      Static bInHereAlready As Boolean
395      If bInHereAlready Then Exit Sub
396      bInHereAlready = True

397      lsbJumpTo.Enabled = False

398      gbCancelInsertion = False
399      mbIgnoreBlanks = False
400      mbIgnoreReadOnly = False

401      Dim lLine As Long
402      Dim lTemp As Long
403      Dim sCodeToInsert As String
404      Dim sProcName As String
405      Dim lProcType As Long
406      Dim sProcTypeLong As String

407      Dim asaVar As CAssocArray                         ' Associative Array used when filling in values to a code template when being inserted
408      Dim CurItem As CAssocItem

    'If txtName <> sTemplateToInsert Then
409      If Not SetInternalCurrentTemplate(sTemplateToInsert) Then
410          LogError "frmMain", "DoInsertion", vbObjectError + 100, "Can't find the " & gsTemplate & gsS & gsA & sTemplateToInsert & "' to insert." & gsEolTab & "Aborting this insertion.", Erl
411          GoTo EH_frmMain_DoInsertion_Continue
412      End If
    'End If

    ' Begin Log
    'frmLog.tvwLog.Nodes.Add , , (frmLog.tvwLog.Nodes.Count + 1) & " Inserting " & sTemplateToInsert, "Inserting " & sTemplateToInsert
    'If Not asaV Is Nothing Then
    '   For Each CurItem In asaV
    '       frmLog.tvwLog.Nodes.Add (frmLog.tvwLog.Nodes.Count + 1) & " Inserting " & sTemplateToInsert, tvwChild, , CurItem.Key & " = " & CurItem.Value
    '   Next CurItem
    'End If
    ' End Log

      If Parent.HostedByVB Then  ' Shell App override
413      If Parent.vbInst.ActiveCodePane Is Nothing Then
414          If Parent.vbInst.SelectedVBComponent Is Nothing Then
415              If Parent.vbInst.ActiveVBProject.VBComponents.Count > 0 Then
416                  Parent.vbInst.ActiveVBProject.VBComponents(1).CodeModule.CodePane.Show
417              Else
418                  Parent.vbInst.ActiveVBProject.VBComponents.Add(vbext_ct_StdModule).CodeModule.CodePane.Show
419              End If
420          ElseIf Not Parent.vbInst.SelectedVBComponent.CodeModule.CodePane Is Nothing Then
421              Parent.vbInst.SelectedVBComponent.CodeModule.CodePane.Show
422          Else
423              MsgBox "Can't do an insertion since no code pane is active.", vbInformation
424              GoTo EH_frmMain_DoInsertion_Continue
425          End If
426      End If
      End If

427      If asaV Is Nothing Then
428          Set asaVar = New CAssocArray                  ' Use a new associative array
429      Else
430          Set asaVar = asaV                             ' Use the supplied assiciative array
431      End If

      If Parent.HostedByVB Then  ' Shell App override
432      With Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane
433          .GetSelection lLine, lTemp, lTemp, lTemp      ' Determine where the cursor is
434          asaVar("Project Name") = Parent.vbInst.ActiveVBProject.Name    ' Add the build in soft variables
435          asaVar("Module Name") = .CodeModule.Parent.Name
436          asaVar("Module Lines") = .CodeModule.CountOfLines
437          asaVar("Module End of Declarations") = .CodeModule.CountOfDeclarationLines + 1

438          GetProcAtLine lLine, sProcName, lProcType
439          If Len(sProcName) Then
440              asaVar.Add "Proc Name", sProcName
441              asaVar.Add "Proc Type", Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
442              sProcTypeLong = .CodeModule.Lines(.CodeModule.ProcBodyLine(sProcName, lProcType), 1)
443              If InStr(sProcTypeLong, "Function") > 0 Then
444                  sProcTypeLong = "Function"
445              ElseIf InStr(sProcTypeLong, "Property") > 0 Then
446                  sProcTypeLong = "Property"
447              Else
448                  sProcTypeLong = "Sub"
449              End If
450              asaVar.Add "Proc Type Long", sProcTypeLong
451              sProcName = vbNullString
452              sProcTypeLong = vbNullString
453          End If

454          If Parent.InsertTemplate(InternalCurrentTemplate.memoCodeAtCursor, lLine, asaVar, sTemplateToInsert) Then
455              If Not bSkipDeclarations Then
456                  If Parent.InsertTemplate(InternalCurrentTemplate.memoCodeAtTop, .CodeModule.CountOfDeclarationLines + 1, asaVar, sTemplateToInsert) Then
457                      If Not Parent.vbInst.ActiveCodePane Is Nothing Then
458                          If Parent.InsertTemplate(InternalCurrentTemplate.memoCodeAtBottom, .CodeModule.CountOfLines + 1, asaVar, sTemplateToInsert) Then
459                              If Not Parent.vbInst.ActiveCodePane Is Nothing Then
460                                  Call Parent.InsertTemplate(InternalCurrentTemplate.memoCodeToFile, 1, asaVar, sTemplateToInsert, txtFilename)
461                              End If
462                          End If
463                      End If
464                  End If
465              Else
466                  If Parent.InsertTemplate(InternalCurrentTemplate.memoCodeAtBottom, .CodeModule.CountOfLines + 1, asaVar, sTemplateToInsert) Then
467                      If Not Parent.vbInst.ActiveCodePane Is Nothing Then
468                          Call Parent.InsertTemplate(InternalCurrentTemplate.memoCodeToFile, 1, asaVar, sTemplateToInsert, txtFilename)
469                      End If
470                  End If
471              End If
472          End If
473      End With
      Else
        ' We don't have an IDE to go into, so just pass zero values and keep going
          If Parent.InsertTemplate(InternalCurrentTemplate.memoCodeAtCursor, 1, asaVar, sTemplateToInsert) Then
              If Not bSkipDeclarations Then
                 If Parent.InsertTemplate(InternalCurrentTemplate.memoCodeAtTop, 1, asaVar, sTemplateToInsert) Then
                    If Parent.InsertTemplate(InternalCurrentTemplate.memoCodeAtBottom, 1, asaVar, sTemplateToInsert) Then
                       Call Parent.InsertTemplate(InternalCurrentTemplate.memoCodeToFile, 1, asaVar, sTemplateToInsert, txtFilename)
                    End If
                 End If
              Else
                 If Parent.InsertTemplate(InternalCurrentTemplate.memoCodeAtBottom, 1, asaVar, sTemplateToInsert) Then
                    Call Parent.InsertTemplate(InternalCurrentTemplate.memoCodeToFile, 1, asaVar, sTemplateToInsert, txtFilename)
                 End If
              End If
          End If
      End If

474   Set asaVar = Nothing                              ' Destroy the associative array

475   sCodeToInsert = vbNullString

476      If mnuExitAfterInsert.Checked = True Then
477          mnuFileExit_Click
478      End If


479 EH_frmMain_DoInsertion_Continue:
480      bInHereAlready = False
481      lsbJumpTo.Enabled = True
482      Exit Sub

483 EH_frmMain_DoInsertion:
484      LogError "frmMain", "DoInsertion", Err.Number, Err.Description, Erl

485      gbCancelInsertion = bUserSure("Cancel processing ?")
486      Resume EH_frmMain_DoInsertion_Continue

487      Resume
End Sub

' ================================================================================
' Name              frmMain_FillTemplateWithUserInput
'
' Parameters
'      asaX                          (O)  CAssocArray   The array to work off of
'      sToParse                      (I)  String        The input string containing soft code
'      sCodeToInsert                 (O)  String        The hard code generated by this routines
'
' Description
'
' This procedure scans the code about to be inserted, and replaces all questions
' with the user supplied responses.
'
' ================================================================================
Public Function FillTemplateWithUserInput(ByRef asaX As CAssocArray, ByVal sToParse As String, ByRef sCodeToInsert As String, ByVal sMsgBoxTitle As String) As Boolean
488      Static sVarName As String
489      Static sVarPhrase As String
490      Static sDefault As String
491      Static sT As String
492      Static sVar1 As String
493      Static sVar2 As String
494      Static sVar3 As String
495      Static sNow As String
496      Static lParamCount As Long
497      Static CurrSet As Long
498      Static bInlineCommandExecuted As Boolean

499      Do While InStr(sGetToken(sToParse, 1, vbNewLine), gsSoftVarDelimiter) > 0    ' For each soft variable found
500          sVarPhrase = sGetToken(sToParse, 2, gsSoftVarDelimiter)    ' Get the Variable name and default if provided
501          sVarName = sGetToken(sVarPhrase, 1, gsInlineCmdDelimiter)    ' Extract just the variable name
502          sNow = vbNullString
503          bInlineCommandExecuted = False
504          If SadCommandSetCount > 0 Then
505              sVar1 = sAfter(sVarPhrase, 1, gsInlineCmdDelimiter)
506              For CurrSet = 1 To SadCommandSetCount
507                  If SadCommands(CurrSet).ExecuteSoftCommandInline(asaX, UCase$(sVarName), sVar1, sNow) Then
508                      bInlineCommandExecuted = True
509                      Exit For
510                  End If
511              Next CurrSet
512          End If
513          If bInlineCommandExecuted Then
514              sT = sToParse
515              sToParse = sBefore(sT, 2, gsSoftVarDelimiter) & sNow & sAfter(sT, 2, gsSoftVarDelimiter)
516          Else
517              With asaX.Item(sVarName)                  ' With the Association for the Soft variable
518                  If Len(.Value) = 0 Then               ' If there is currently no value
519                      If mbIgnoreBlanks Then
520                      ElseIf InStr(sVarPhrase, gsInlineCmdDelimiter) Then    ' Use default provided
521                          sDefault = sGetToken(sVarPhrase, 2, gsInlineCmdDelimiter)    ' Extract the default
522                          If Left$(sDefault, 1) = "@" Then    ' See if the default is to be drawn from another Association's value
523                              sDefault = asaX.Item(Mid$(sDefault, 2)).Value    ' Lookup another value in the array as the default
524                          End If
525                          .Value = InputBox(sVarName, sMsgBoxTitle, sDefault)    ' Ask the user to enter a value and then set the Association's value to it
526                          If Len(.Value) = 0 Then gbCancelInsertion = bUserSure("Cancel processing ?")
527                      Else                              ' No default
528                          .Value = InputBox(sVarName, sMsgBoxTitle)    ' Ask the user to enter a value and then set the Association's value to it
529                          If Len(.Value) = 0 Then gbCancelInsertion = bUserSure("Cancel processing ?")
530                      End If
531                  End If                                ' At this point the Association's value is set one way or the other
532                  If gbCancelInsertion Then
533                      sVarName = vbNullString
534                      sVarPhrase = vbNullString
535                      sDefault = vbNullString
536                      sT = vbNullString
537                      FillTemplateWithUserInput = False
538                      Exit Function                     ' User canceled
539                  End If
540                  sT = sToParse                         ' Save the string so far into a temporary area
541                  sToParse = sBefore(sT, 2, gsSoftVarDelimiter) & .Value & sAfter(sT, 2, gsSoftVarDelimiter)    ' Replace the Soft variable with the user's entry
542              End With
543          End If
544      Loop

545      FillTemplateWithUserInput = True                  ' Returned the final parsed string

546      If InStr(sToParse, vbNewLine) Then
547          If Right$(sToParse, 2) <> vbNewLine Then      ' If the code to insert is more than a line long
548              sToParse = sToParse & vbNewLine           ' Insure it has an EOL at the end to be parsed properly
549          End If
550      End If

551      sCodeToInsert = sToParse

552      sVarName = vbNullString
553      sVarPhrase = vbNullString
554      sDefault = vbNullString
555      sT = vbNullString
556      sVar1 = vbNullString
557      sVar2 = vbNullString
558      sVar3 = vbNullString

End Function



' ================================================================================
' Name              frmMain_InternalInsertTemplate
'
' Parameters
'
' Description
'
' This actually causes the code indicated to get inserted correctly. Soft
' commands are handled here.
'
' ================================================================================
Public Function InternalInsertTemplate(II As CInsertionInfo) As Boolean
559      Dim CurFrame        As VBControl
560      Dim CurForm         As VBForm
561      Dim CurReference    As Reference
562      Dim CurControl      As VBControl
563      Dim CurModule       As CodeModule
564      Dim ControlVars     As CAssocArray
565      Dim tTemplate       As CTemplate

566      Dim tProject        As VBProject
567      Dim tModule         As CodeModule
568      Dim tPane           As CodePane
569      Dim tWindow         As Window
570      Dim tWindows        As Windows

571      Dim tComponent      As VBComponent

572      Dim fh As Long
573      Dim CurrSet As Long
574      Dim CurrParam As Long
575      Dim lParamCount As Long
576      Dim lStartLine As Long
577      Dim lEndLine As Long
578      Dim lStartColFound As Long
579      Dim lEndColFound As Long
580      Dim lProcType As Long
581      Dim lMouseState As MousePointerConstants

582      Dim lIfLoops As Long
583      Dim CodaIterations As Long
584      Dim NextElse As Long
585      Dim NextElseIf As Long
586      Dim NextEndIf As Long
587      Dim CmdIterations As Long

588      Dim bFunction As Boolean
589      Dim bFoundReference As Boolean
590      Dim bDoCoda As Boolean
591      Dim bT As Boolean

592      Dim sT As String
593      Dim sHold1 As String
594      Dim sProcName As String
595      Dim sProcType As String
596      Dim sHold2 As String
597      Dim sHold3 As String
598      Dim sHold4 As String
599      Dim sCurParam As String
600      Dim sCurType As String
601      Dim CommandReference As String
602      Dim sDelim1 As String
603      Dim sDelim2 As String

    ' For use by internal commands only (not outside select case statment)
604      Dim sT2 As String
605      Dim scT2 As String
606      Dim scT3 As String
607      Dim scT1 As String
608      Dim scT4 As String


609      On Error GoTo EH_InsertTemplate

610      If II Is Nothing Then
611          InternalInsertTemplate = True
612          GoTo EH_InsertTemplate_Continue
613      End If

         If Parent.HostedByVB Then  ' Shell App override
614         Set CurModule = Parent.vbInst.ActiveCodePane.CodeModule
         Else
            Set CurModule = Nothing
         End If

    ' Preprocess this Template including any referenced Template Fragments
615      lStartLine = InStr(II.OriginalCodeToInsert, "~##~Include ")
616      Do While lStartLine > 0
617          sCurType = sGetToken(Mid$(II.OriginalCodeToInsert, lStartLine + 12), 1, vbNewLine)
618          sHold2 = sCurType
619          If Val(sGetToken(sHold2, 1, gsC)) > 0 Then
620              sHold1 = Val(sGetToken(sHold2, 1, gsC))
621              sHold2 = sAfter(sHold2, 1, gsC)
622          Else
623              sHold1 = 0
624          End If

625          If InStr(sHold2, gsCategoryTemplateDelimiter) = 0 Then
626              sHold3 = sGetToken(InternalCurrentTemplate, 1, gsCategoryTemplateDelimiter)
627              sHold4 = sHold2
628          Else
629              sHold3 = sGetToken(sHold2, 1, gsCategoryTemplateDelimiter)
630              sHold4 = sAfter(sHold2, 1, gsCategoryTemplateDelimiter)
631          End If

632          sCurParam = vbNullString
             If SliceAndDice.Categorys(sHold3) Is Nothing Then
                gbCancelInsertion = bUserSure("Preprocessor directive:" & vbNewLine & vbTab & "~##~Include" & vbNewLine & vbNewLine & "Cannot find Category for Template:" & vbNewLine & vbTab & sHold3 & " - " & sHold4 & vbNewLine & vbNewLine & vbNewLine & "Cancel processing ?", "CANCEL PROCESSING ?")
                Set tTemplate = Nothing
             Else
633             Set tTemplate = SliceAndDice.Categorys(sHold3).Templates(sHold4)
             End If

634          If Not tTemplate Is Nothing Then
635              With tTemplate
                Select Case sHold1
                    Case 1: sCurParam = .memoCodeAtTop
636                      Case 2: sCurParam = .memoCodeAtCursor
637                      Case 3: sCurParam = .memoCodeAtBottom
638                      Case 0
639                          sCurParam = .memoCodeAtTop
                        'sCurParam = vbNullString
                        'If Len(.memoCodeAtCursor) Then sCurParam = sCurParam & .memoCodeAtCursor
                        'If Len(.memoCodeAtTop) Then sCurParam = sCurParam & IIf(Len(sCurParam) And Right$(sCurParam, 2) <> vbNewLine, vbNewLine, vbNullString) & "~~GotoDec End" & vbNewLine & .memoCodeAtTop
                        'If Len(.memoCodeAtBottom) Then sCurParam = sCurParam & IIf(Len(sCurParam) And Right$(sCurParam, 2) <> vbNewLine, vbNewLine, vbNullString) & "~~GotoEnd" & vbNewLine & .memoCodeAtTop
640                      Case Else: sCurParam = .memoCodeToFile
641                  End Select
642              End With
643              Set tTemplate = Nothing
644          Else
645              gbCancelInsertion = bUserSure("Preprocessor directive:" & vbNewLine & vbTab & "~##~Include" & vbNewLine & vbNewLine & "Cannot find Template:" & vbNewLine & vbTab & sHold3 & " - " & sHold4 & vbNewLine & vbNewLine & vbNewLine & "Cancel processing ?", "CANCEL PROCESSING ?")
646              If gbCancelInsertion Then GoTo EH_InsertTemplate_Continue
647          End If

648          II.OriginalCodeToInsert = Left$(II.OriginalCodeToInsert, lStartLine - 1) & sCurParam & Mid$(II.OriginalCodeToInsert, lStartLine + 12 + Len(sCurType))

649          lStartLine = InStr(II.OriginalCodeToInsert, "~##~Include ")
650      Loop

651      II.LinesLeftToProcess = II.OriginalCodeToInsert

652      On Error Resume Next
653      II.SoftVars("FirstColumn").Value = DetermineFirstColumnInSelection
654      II.SoftVars("LastColumn").Value = DetermineLastColumnInSelection
655      II.SoftVars("FirstLine").Value = DetermineFirstLineInSelection
656      II.SoftVars("LastLine").Value = DetermineLastLineInSelection

657      On Error GoTo EH_InsertTemplate

658 CODA_RESTART:
659      If Len(II.LinesLeftToProcess) > 100000 Then
660          If Not bUserSure(gsTemplate & " to insert " & gsA & II.TemplateName & "' has become is very large" & vbNewLine & vbTab & gsPO & Len(II.LinesLeftToProcess) & " bytes, started at " & Len(II.OriginalCodeToInsert) & " bytes)." & vbNewLine & vbTab & "Continue inserting anyway ?") Then
661              II.LinesLeftToProcess = vbNullString
662              sT = vbNullString
663              gbCancelInsertion = True
664              GoTo EH_InsertTemplate_Continue
665          End If
666      End If
667      If InStr(II.LinesLeftToProcess, gsSoftCmdDelimiter) = 0 And InStr(II.LinesLeftToProcess, gsSoftVarDelimiter) = 0 Then
668          If Len(II.ExternalFilename) = 0 Then
669              If Len(II.LinesLeftToProcess) > 0 Then
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
670                    CurModule.InsertLines II.PointOfInsertion, II.LinesLeftToProcess    ' No embedded commands, no embedded variables. Simple insertion
                    End If
671              End If
672          Else
673              II.TextToSendToFile = II.LinesLeftToProcess
674          End If
675      Else
676          Do Until Len(II.LinesLeftToProcess) = 0       ' More complicated line by line with embedded commands (and/or variables) insertion
677              DoEvents
678              If Not FillTemplateWithUserInput(II.SoftVars, sGetToken(II.LinesLeftToProcess, 1, vbNewLine), sT, II.TemplateName) Then
679                  InternalInsertTemplate = False
680                  GoTo EH_InsertTemplate_Continue
681              End If
682              If InStr(sT, vbNewLine) Then              ' Inserting caused more lines to appear, push the extra lines into the buffer for later insertion
683                  II.LinesLeftToProcess = sGetToken(II.LinesLeftToProcess, 1, vbNewLine) & vbNewLine & sAfter(sT, 1, vbNewLine) & IIf(Right$(sT, 2) = vbNewLine, vbNullString, vbNewLine) & sAfter(II.LinesLeftToProcess, 1, vbNewLine)
684                  sT = sGetToken(sT, 1, vbNewLine)
685              End If
686              II.CurrentLineToProcess = sT
687              If Left$(LTrim$(II.CurrentLineToProcess), 2) = gsSoftCmdDelimiter Then    ' Process an imbedded command
688                  II.CurrentLineToProcess = sGetToken(II.CurrentLineToProcess, 2, gsSoftCmdDelimiter)    ' Get the command with parameter(s)
689                  II.SoftCommandName = sGetToken(II.CurrentLineToProcess)    ' Get just the command string (Case insensitive)
690                  II.sParam = Trim$(sAfter(II.CurrentLineToProcess))    ' Get just the parameters
691                  II.AllParameters = Replace(Replace(II.sParam, "$SP$", gsS), "$TAB$", vbTab)

692                  If lTokenCount(II.sParam) = 1 Then
693                      If Val(II.sParam) <> 0 Then       ' One parameter passed
694                          sProcName = vbNullString      ' Parameter is a number (line offset)
695                          II.ParamLineOffset = Val(II.sParam)
696                          II.sParam = vbNullString
697                      Else
698                          sProcName = II.sParam         ' Parameter is a procedure name or a real parameter
699                          II.ParamLineOffset = 0
700                          II.sParam = vbNullString
701                      End If
702                  ElseIf lTokenCount(II.sParam) = 2 Then ' Two parameters passed
703                      sProcName = sGetToken(II.sParam)  ' Get the procedure name to work on
704                      II.sParam = sAfter(II.sParam)     ' Strip out the procedure name
705                      sProcType = UCase$(II.sParam)     ' Get the procedure type (first of two parameters)                  [ Default: PROC ]
706                      If Val(sProcType) <> 0 Then
707                          II.ParamLineOffset = Val(sProcType)
708                          sProcType = vbNullString
709                      Else
710                          II.ParamLineOffset = 0
711                      End If
712                  ElseIf IsNumeric(sAfter(II.sParam)) Then ' Three or more parameters passed
713                      sProcName = sGetToken(II.sParam)  ' Get the procedure name to work on
714                      II.sParam = sAfter(II.sParam)     ' Strip out the procedure name
715                      II.ParamLineOffset = "" & sAfter(II.sParam)    ' Get the line offset (usually negative) (second of two parameters) [ Default: 0 ]
716                      sProcType = UCase$(sGetToken(II.sParam))    ' Get the procedure type (first of two parameters)                  [ Default: PROC ]
                     ElseIf IsNumeric(sGetToken(II.sParam, 3)) Then ' Three or more parameters passed
                         sProcName = sGetToken(II.sParam)  ' Get the procedure name to work on
                         II.sParam = sAfter(II.sParam)     ' Strip out the procedure name
                         II.ParamLineOffset = Val("" & sAfter(II.sParam))    ' Get the line offset (usually negative) (second of two parameters) [ Default: 0 ]
                         sProcType = UCase$(sGetToken(II.sParam))    ' Get the procedure type (first of two parameters)                  [ Default: PROC ]
                     Else
                         sProcName = sGetToken(II.sParam)  ' Get the procedure name to work on
                         II.sParam = sAfter(II.sParam)     ' Strip out the procedure name
                         II.ParamLineOffset = 0 ' Get the line offset (usually negative) (second of two parameters) [ Default: 0 ]
                         sProcType = "" 'UCase$(sGetToken(II.sParam))    ' Get the procedure type (first of two parameters)                  [ Default: PROC ]
                     End If
                ' Determine which constant to use for the passed Procedure Type
                ' Note: Nothing passed ? Assume a sub or function
718                  lProcType = Switch(sProcType = "GET", vbext_pk_Get, sProcType = "SET", vbext_pk_Set, sProcType = "LET", vbext_pk_Let, sProcType = "PROC", vbext_pk_Proc, True, vbext_pk_Proc)
                ' Execute the command specified
719                  If InStr(II.AllParameters, gsE) > 0 Then
720                      II.Result = Trim$(sGetToken(II.AllParameters, 1, gsE))
721                      II.Expression = Trim$(sAfter(II.AllParameters, 1, gsE))
                    'II.Expression = sAfter(II.AllParameters, 1, gsE)
722                  End If

                ' Determine indentation level of subcode
723                  If StrComp(Left$(II.SoftCommandName, 1), "_") = 0 Then
724                      CommandReference = "_" & sGetToken(II.SoftCommandName, 2, "_") & "_"
725                      II.SoftCommandName = sAfter(II.SoftCommandName, 2, "_")
726                      If Len(II.SoftCommandName) = 0 Then
727                          If bUserSure(gsSliceAndDice & " has detected a dangling Command Reference in line:" & gs2EOL & II.CurrentLineToProcess & gs2EOLTab & "Would you like to cancel insertion ?") Then
728                              gbCancelInsertion = True
729                          End If
730                      End If
731                  Else
732                      CommandReference = vbNullString
733                  End If

734                  II.SoftCommandName = UCase$(II.SoftCommandName)

                Select Case II.SoftCommandName
                    Case "BLOCK"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
735                          sProcName = II.SoftVars(sProcName)
736                          If II.PointOfInsertion < 1 Then II.PointOfInsertion = 1
737                          If Len(II.ExternalFilename) = 0 Then
738                              CurModule.InsertLines II.PointOfInsertion, sProcName
739                          Else
740                              II.TextToSendToFile = II.TextToSendToFile & sProcName & vbNewLine
741                          End If
742                          II.PointOfInsertion = II.PointOfInsertion + lTokenCount(sProcName, vbNewLine)
                    End If

743                      Case "DEBUG"
744                          On Error Resume Next
745                          II.SoftVars("DEBUG").Value = _
                                "LinesLeftToProcess" & vbNewLine & II.LinesLeftToProcess & vbNewLine & String$(80, "-") & vbNewLine & _
                                "OriginalCodeToInsert" & vbTab & vbTab & II.OriginalCodeToInsert & vbNewLine & String$(80, "-") & vbNewLine & _
                                "PointOfInsertion" & vbTab & vbTab & II.PointOfInsertion & vbNewLine & String$(80, "-") & vbNewLine & _
                                "SoftVars.All" & vbTab & vbTab & II.SoftVars.All & vbNewLine & String$(80, "-") & vbNewLine & _
                                "TemplateName" & vbTab & vbTab & II.TemplateName
746                          On Error GoTo EH_InsertTemplate

747                      Case "ELSE"
748                          If lIfLoops > 0 Then
749                              NextEndIf = InStr(UCase$(II.LinesLeftToProcess), gsSoftCmdDelimiter & UCase$(CommandReference) & "ENDIF")
750                              If NextEndIf > 0 Then
751                                  II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, NextEndIf)
752                              Else
753                                  II.LinesLeftToProcess = vbNullString
754                              End If
755                              lIfLoops = lIfLoops - 1
756                          End If

757                      Case "IF"
758                          If Val(Evaluate(II.AllParameters, II.SoftVars)) <> 0 Then
759                              lIfLoops = lIfLoops + 1
760                          Else
761                              NextEndIf = InStr(UCase$(II.LinesLeftToProcess), gsSoftCmdDelimiter & UCase$(CommandReference) & "ENDIF")
762                              NextElse = InStr(UCase$(II.LinesLeftToProcess), gsSoftCmdDelimiter & UCase$(CommandReference) & "ELSE" & vbNewLine)
763                              If NextElse > 0 And NextElse < NextEndIf Then
764                                  II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, NextElse)
765                              ElseIf NextEndIf <> 0 Then
766                                  II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, NextEndIf)
767                              Else
768                                  II.LinesLeftToProcess = vbNullString
769                              End If
770                          End If

771                      Case "ENDIF"
772                          If lIfLoops > 0 Then lIfLoops = lIfLoops - 1

773                      Case gsA, "STARTCODA", "ENDIF", "STARTLOOPWHILE", "STARTLOOPUNTIL"

774                      Case "ABORT", "ABORTINSERTION"
775                          On Error Resume Next
776                          If Val(II.AllParameters) <> 0 Then
777                              MsgBox "Insertion aborted by the " & gsSoftCmdDelimiter & "AbortInsertion command."
778                              II.LinesLeftToProcess = vbNullString
779                              gbCancelInsertion = True
780                          End If
781                          On Error GoTo EH_InsertTemplate

782                      Case "CANCEL", "CANCELINSERTION"
783                          On Error Resume Next
784                          If Evaluate(II.AllParameters, II.SoftVars) <> 0 Then
785                              II.LinesLeftToProcess = vbNullString
786                              gbCancelInsertion = True
787                          End If
788                          On Error GoTo EH_InsertTemplate

                        '                         Case "INSERTTEMPLATE", "INSERT"
                        '                              DoInsertion asaX, II.AllParameters

                        ' ================================================
                        ' Soft Commands that control the flow of insertion
                        ' ================================================
789                      Case "CODA", "LOOPWHILE", "LOOPUNTIL"
790                          CodaIterations = CodaIterations + 1
791                          If CodaIterations > 10000 Then
792                              If bUserSure(gsSliceAndDice & " has found what appears to be an endless loop via " & gsSoftCmdDelimiter & "Coda, " & gsSoftCmdDelimiter & "LoopWhile, or " & gsSoftCmdDelimiter & "LoopUntil." & vbNewLine & vbTab & "Would you like to cancel processing ?") Then
793                                  gbCancelInsertion = True
794                              Else
795                                  CodaIterations = 0
796                              End If
797                          End If
798                          If II.SoftCommandName = "LOOPUNTIL" Then
799                              bDoCoda = (Val(sGetToken(II.AllParameters)) = 0)
800                          Else
801                              bDoCoda = (Val(sGetToken(II.AllParameters)) <> 0)
802                          End If
803                          If bDoCoda Then
                            Select Case II.SoftCommandName
                                Case "CODA"
804                                      If InStr(II.OriginalCodeToInsert, CommandReference & "STARTCODA") Then
805                                          II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "STARTCODA")
806                                      ElseIf InStr(II.OriginalCodeToInsert, CommandReference & "StartCoda") Then
807                                          II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "StartCoda")
808                                      ElseIf InStr(II.OriginalCodeToInsert, CommandReference & "startcoda") Then
809                                          II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "startcoda")
810                                      End If
811                                  Case "LOOPWHILE"
812                                      If InStr(II.OriginalCodeToInsert, CommandReference & "STARTLOOPWHILE") Then
813                                          II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "STARTLOOPWHILE")
814                                      ElseIf InStr(II.OriginalCodeToInsert, CommandReference & "StartLoopWhile") Then
815                                          II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "StartLoopWhile")
816                                      ElseIf InStr(II.OriginalCodeToInsert, CommandReference & "startloopwhile") Then
817                                          II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "startloopwhile")
818                                      End If
819                                  Case "LOOPUNTIL"
820                                      If InStr(II.OriginalCodeToInsert, CommandReference & "STARTLOOPUNTIL") Then
821                                          II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "STARTLOOPUNTIL")
822                                      ElseIf InStr(II.OriginalCodeToInsert, CommandReference & "StartLOOPUNTIL") Then
823                                          II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "StartLOOPUNTIL")
824                                      ElseIf InStr(II.OriginalCodeToInsert, CommandReference & "startloopuntil") Then
825                                          II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "startloopuntil")
826                                      End If
827                              End Select

828                              If Left$(II.LinesLeftToProcess, 2) = vbNewLine Then
829                                  II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, 3)
830                              End If
831                              GoTo CODA_RESTART
832                          Else
833                          End If

834                      Case "NOINSERT", "STOPCODEINSERTION", "STOP", "STOPINSERTION"
835                          InternalInsertTemplate = True
836                          GoTo EH_InsertTemplate_Continue    ' Prematurely stop processing of this template

837                      Case "RESUMEINSERTION", "RESUME", "FLUSH", "FLUSHBUFFER"  ' Clears file insertion and resume code insertion
                              ChangeFocusOfInsertion II, ""

                        ' ========================================================
                        ' Soft Commands that specially process areas of code/forms
                        ' ========================================================
849                      Case "COMMENTEDPARAMETERS"        ' Parse a function's parameters into readable comments
                    If Not CurModule Is Nothing Then
850                          GetProcAtLine II.PointOfInsertion, sProcName, lProcType
851                          If sProcName <> vbNullString Then
852                              If Len(II.AllParameters) Then
853                                  If lTokenCount(II.AllParameters) = 3 Then
854                                      sDelim1 = sGetToken(II.AllParameters, 2)
855                                      sDelim2 = sGetToken(II.AllParameters, 3)
856                                  ElseIf lTokenCount(II.AllParameters) = 2 Then
857                                      sDelim1 = sGetToken(II.AllParameters, 2)
858                                      sDelim2 = "||"
859                                  ElseIf lTokenCount(II.AllParameters) = 1 Then
860                                      sDelim1 = "$$"
861                                      sDelim2 = "||"
862                                  Else
863                                      II.AllParameters = vbNullString
864                                  End If
865                              Else
866                                  sDelim1 = vbNullString
867                                  sDelim2 = vbNullString
868                              End If
869                              lStartLine = CurModule.ProcBodyLine(sProcName, lProcType)
870                              II.sParam = CurModule.Lines(lStartLine, 1)    ' Get the procedure's header
871                              Do While Right$(Trim$(II.sParam), 2) = " _"
872                                  lStartLine = lStartLine + 1
873                                  II.sParam = Trim$(sBefore(II.sParam, lTokenCount(II.sParam, " _"), " _")) & gsS & Trim$(CurModule.Lines(lStartLine, 1))    ' Get the next procedure's header line
874                              Loop
875                              bFunction = (InStr(II.sParam, "Function") > 0) Or (InStr(II.sParam, "Property Get") > 0)
876                              II.sParam = sAfter(II.sParam, 1, gsPO)    ' Get just the parameters
877                              If bFunction Then
878                                  If lTokenCount(II.sParam, ") As ") > 1 Then
879                                      sHold1 = sAfter(II.sParam, lTokenCount(II.sParam, ") As ") - 1, ") As ")
880                                  Else
881                                      sHold1 = "Variant"
882                                  End If
883                              Else
884                                  sHold1 = vbNullString
885                              End If
886                              If lTokenCount(II.sParam, gsPC) > 1 Then
887                                  II.sParam = sBefore(II.sParam, lTokenCount(II.sParam, gsPC), gsPC)    ' Strip out the return type (FUTURE: Possibly use this later)
888                              End If
889                              lParamCount = lTokenCount(II.sParam, gsC)    ' Find out how many parameters there are
890                              If Len(II.sParam) = 0 Then
891                                  If Len(II.AllParameters) = 0 Then
892                                      sCurType = "'      None" & vbNewLine
893                                  Else
894                                      sCurType = vbNullString
895                                  End If
896                              Else
897                                  For CurrParam = 1 To lParamCount    ' For each parameter
898                                      sCurParam = Trim$(sGetToken(II.sParam, 1, gsC))    ' Get the next parameter
899                                      II.sParam = sAfter(II.sParam, 1, gsC)    ' Chop off the parameter
900                                      If InStr(sCurParam, "As") > 0 Then
901                                          sHold2 = sGetToken(sCurParam, 2, " As ") & gsP
902                                          If InStr(sHold2, gsE) > 0 Then
903                                              sHold2 = Trim$(sGetToken(sHold2, 1, gsE)) & " Defaults to " & Trim$(sAfter(sHold2, 1, gsE))    ' & gsP
904                                          End If
905                                          sCurParam = sGetToken(sCurParam, 1, " As ")
906                                          If InStr(sCurParam, "Optional") > 0 Then
907                                              sHold2 = "Opt. " & sHold2
908                                              sCurParam = Trim$(Replace(sCurParam, "Optional", vbNullString))
909                                          End If
910                                          If InStr(sCurParam, "ByVal ") > 0 Then
911                                              sHold2 = "(I)  " & sHold2
912                                              sCurParam = Replace(sCurParam, "ByVal ", vbNullString)
913                                          ElseIf InStr(sCurParam, "ByRef ") > 0 Then
914                                              sHold2 = "(O)  " & sHold2
915                                              sCurParam = Replace(sCurParam, "ByRef ", vbNullString)
916                                          Else
917                                              sHold2 = "(IO) " & sHold2
918                                          End If

919                                          If Len(sCurParam) > 2 Then
920                                              If Right$(sCurParam, 2) = "()" Then    ' Array
921                                                  sCurParam = Left$(sCurParam, Len(sCurParam) - 2)
922                                                  If Left$(sHold2, 4) = "(IO)" Then
923                                                      sHold2 = "(O) " & Mid$(Left$(sHold2, Len(sHold2) - 1), 5) & " Array."
924                                                  Else
925                                                      sHold2 = Left$(sHold2, Len(sHold2) - 1) & " Array."
926                                                  End If
927                                              End If
928                                          End If
929                                      Else
930                                          sHold2 = "Variant."
931                                          If InStr(sCurParam, "Optional") > 0 Then
932                                              sHold2 = "Opt. " & sHold2
933                                              sCurParam = Trim$(Replace(sCurParam, "Optional", vbNullString))
934                                          End If
935                                          If InStr(sCurParam, "ByVal ") > 0 Then
936                                              sHold2 = "(I)  " & sHold2
937                                              sCurParam = Replace(sCurParam, "ByVal ", vbNullString)
938                                          ElseIf InStr(sCurParam, "ByRef ") > 0 Then
939                                              sHold2 = "(O)  " & sHold2
940                                              sCurParam = Replace(sCurParam, "ByRef ", vbNullString)
941                                          Else
942                                              sHold2 = "(IO) " & sHold2
943                                          End If
                                        'sCurType = sCurType & "'      " & sCurParam & Space$(30 - Len(sCurParam)) & gsP & vbNewLine
944                                      End If
945                                      If Len(II.AllParameters) = 0 Then
946                                          sCurType = sCurType & "'      " & sCurParam & Space$(30 - Len(sCurParam)) & sHold2 & vbNewLine
947                                      Else
948                                          sHold2 = Replace(Replace(sHold2, gsPO, vbNullString), ") ", sDelim2)
949                                          sCurType = sCurType & sCurParam & sDelim2 & Left$(sHold2, Len(sHold2) - 1) & sDelim1
950                                      End If
951                                  Next CurrParam        ' Add it to the growing commented parameters string
952                                  If Len(II.AllParameters) = 0 Then
953                                      If Len(sHold1) > 0 Then
954                                          sCurType = sCurType & gsA & vbNewLine
955                                          sCurType = sCurType & "' Returns" & vbNewLine
956                                          sCurType = sCurType & "'      " & sHold1 & Space$(30 - Len(sHold1)) & gsP & vbNewLine
957                                      End If
958                                  Else
959                                      II.SoftVars("Function Returns").Value = sHold1
960                                  End If
961                              End If
962                          End If
963                          If Len(II.AllParameters) = 0 Then
964                              II.LinesLeftToProcess = vbNewLine & sCurType & sAfter(II.LinesLeftToProcess, 1, vbNewLine)    ' Insert the commented parameters into the insertion stream
965                          Else
966                              II.SoftVars(sGetToken(II.AllParameters)).Value = sCurType
967                          End If
968                          On Error GoTo EH_InsertTemplate
                    End If

969                      Case "FOREACHCONTROL"
         If Parent.HostedByVB Then  ' Shell App override
970                          Set ControlVars = New CAssocArray    ' Collect what to do for each type of control encountered
971                          Do Until Left$(II.CurrentLineToProcess, 2) = gsSoftCmdDelimiter
972                              If Left$(II.LinesLeftToProcess, 2) = vbNewLine Then    ' Strip off the line just parsed
973                                  II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, 3)
974                              Else
975                                  II.LinesLeftToProcess = sAfter(II.LinesLeftToProcess, 1, vbNewLine)
976                              End If
977                              II.CurrentLineToProcess = sGetToken(II.LinesLeftToProcess, 1, vbNewLine)
978                              If Left$(II.CurrentLineToProcess, 2) = gsSoftCmdDelimiter Then
979                              ElseIf Left$(II.CurrentLineToProcess, 2) = gsSpecialLineItemDelimiter Then
980                                  sCurType = UCase$(Trim$(Mid$(II.CurrentLineToProcess, 3)))
981                                  ControlVars.Add sCurType
982                              ElseIf Len(sCurType) > 0 Then
983                                  ControlVars(sCurType) = ControlVars(sCurType) & II.CurrentLineToProcess & vbNewLine
984                              End If
985                          Loop
986                          For Each CurControl In Parent.vbInst.SelectedVBComponent.Designer.VBControls
987                              If Len(ControlVars(UCase$(CurControl.ClassName))) > 0 Then
988                                  sCurType = ControlVars(UCase$(CurControl.ClassName))
989                                  Do Until InStr(sCurType, gsSpecialLineItemDelimiter) = 0
990                                      II.sParam = sGetToken(sCurType, 2, gsSpecialLineItemDelimiter)
991                                      If UCase$(Left$(II.sParam, 8)) = "CONTROL." Then
992                                          II.sParam = Mid$(II.sParam, 9)
993                                      End If
994                                      If II.sParam = "sName" Then
995                                          II.sParam = Mid$(CurControl.Properties("Name"), 4)
996                                      Else
997                                          II.sParam = CurControl.Properties(II.sParam)
998                                      End If
999                                      sCurType = sGetToken(sCurType, 1, gsSpecialLineItemDelimiter) & II.sParam & sAfter(sCurType, 2, gsSpecialLineItemDelimiter)
1000                                 Loop
1001                                 If Right$(sCurType, 2) = vbNewLine Then
1002                                     sCurType = Left$(sCurType, Len(sCurType) - 2)
1003                                 End If
1004                                 CurModule.InsertLines II.PointOfInsertion, sCurType
1005                                 II.PointOfInsertion = II.PointOfInsertion + lTokenCount(sCurType, vbNewLine)
1006                             End If
1007                         Next CurControl

1008                         Set ControlVars = Nothing
         End If
1009                     Case "FOREACHCONTROLBYFRAME"
         If Parent.HostedByVB Then  ' Shell App override
1010                         Set ControlVars = New CAssocArray    ' Collect what to do for each type of control encountered
1011                         Do Until Left$(II.CurrentLineToProcess, 2) = gsSoftCmdDelimiter
1012                             If Left$(II.LinesLeftToProcess, 2) = vbNewLine Then    ' Strip off the line just parsed
1013                                 II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, 3)
1014                             Else
1015                                 II.LinesLeftToProcess = sAfter(II.LinesLeftToProcess, 1, vbNewLine)
1016                             End If
1017                             II.CurrentLineToProcess = sGetToken(II.LinesLeftToProcess, 1, vbNewLine)
1018                             If Left$(II.CurrentLineToProcess, 2) = gsSoftCmdDelimiter Then
1019                             ElseIf Left$(II.CurrentLineToProcess, 2) = gsSpecialLineItemDelimiter Then
1020                                 sCurType = UCase$(Trim$(Mid$(II.CurrentLineToProcess, 3)))
1021                                 ControlVars.Add sCurType
1022                             ElseIf Len(sCurType) > 0 Then
1023                                 ControlVars(sCurType) = ControlVars(sCurType) & II.CurrentLineToProcess & vbNewLine
1024                             End If
1025                         Loop
1026                         For Each CurFrame In Parent.vbInst.SelectedVBComponent.Designer.VBControls
1027                             If CurFrame.ClassName = "Frame" Then
1028                                 sCurType = ControlVars(UCase$(CurFrame.ClassName))
1029                                 Do Until InStr(sCurType, gsSpecialLineItemDelimiter) = 0
1030                                     II.sParam = sGetToken(sCurType, 2, gsSpecialLineItemDelimiter)
1031                                     If UCase$(Left$(II.sParam, 8)) = "CONTROL." Then
1032                                         II.sParam = Mid$(II.sParam, 9)
1033                                     End If
1034                                     If II.sParam = "sName" Then
1035                                         II.sParam = Mid$(CurFrame.Properties("Name"), 4)
1036                                     Else
1037                                         II.sParam = CurFrame.Properties(II.sParam)
1038                                     End If
1039                                     sCurType = sGetToken(sCurType, 1, gsSpecialLineItemDelimiter) & II.sParam & sAfter(sCurType, 2, gsSpecialLineItemDelimiter)
1040                                 Loop
1041                                 If Len(sCurType) > 0 Then
1042                                     If Right$(sCurType, 2) = vbNewLine Then
1043                                         sCurType = Left$(sCurType, Len(sCurType) - 2)
1044                                     End If
1045                                     CurModule.InsertLines II.PointOfInsertion, sCurType
1046                                     II.PointOfInsertion = II.PointOfInsertion + lTokenCount(sCurType, vbNewLine)
1047                                 End If
1048                                 For Each CurControl In CurFrame.ContainedVBControls
1049                                     If CurControl.ClassName <> "Frame" And Len(ControlVars(UCase$(CurControl.ClassName))) > 0 Then
1050                                         sCurType = ControlVars(UCase$(CurControl.ClassName))
1051                                         Do Until InStr(sCurType, gsSpecialLineItemDelimiter) = 0
1052                                             II.sParam = sGetToken(sCurType, 2, gsSpecialLineItemDelimiter)
1053                                             If UCase$(Left$(II.sParam, 8)) = "CONTROL." Then
1054                                                 II.sParam = Mid$(II.sParam, 9)
1055                                             End If
1056                                             If II.sParam = "sName" Then
1057                                                 II.sParam = Mid$(CurControl.Properties("Name"), 4)
1058                                             Else
1059                                                 II.sParam = CurControl.Properties(II.sParam)
1060                                             End If
1061                                             sCurType = sGetToken(sCurType, 1, gsSpecialLineItemDelimiter) & II.sParam & sAfter(sCurType, 2, gsSpecialLineItemDelimiter)
1062                                         Loop
1063                                         If Right$(sCurType, 2) = vbNewLine Then
1064                                             sCurType = Left$(sCurType, Len(sCurType) - 2)
1065                                         End If
1066                                         CurModule.InsertLines II.PointOfInsertion, sCurType
1067                                         II.PointOfInsertion = II.PointOfInsertion + lTokenCount(sCurType, vbNewLine)
                                        'Else
                                        '   LogEvent  CurControl.ClassName
1068                                     End If
1069                                 Next CurControl
1070                             End If
1071                         Next CurFrame

1072                         Set ControlVars = Nothing
         End If

                        ' ======================================================================
                        ' Soft commands that directly manipulate VB module(s)/Forms/Controls/etc.
                        ' ======================================================================
1073                     Case "CLOSE", "CLOSECODE", "CLOSEWINDOW"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1074                         On Error Resume Next
1075                         CurModule.CodePane.Window.Close
1076                         On Error GoTo EH_InsertTemplate
                    End If

1077                     Case "HIDE", "HIDECODE", "HIDEWINDOW"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1078                         On Error Resume Next
1079                         CurModule.CodePane.Window.Visible = False
1080                         On Error GoTo EH_InsertTemplate
                    End If

1081                     Case "CLOSEALL", "CLOSEWINDOWS", "CLOSEALLWINDOWS"
         If Parent.HostedByVB Then  ' Shell App override
1082                         lMouseState = Screen.MousePointer
1083                         Screen.MousePointer = vbHourglass

1084                         Set tWindows = Parent.vbInst.Windows
1085                         For Each tWindow In tWindows
1086                             If tWindow.Type = vbext_wt_CodeWindow Or tWindow.Type = vbext_wt_Designer Then
1087                                 If tWindow.Visible Then
1088                                     tWindow.Close
1089                                 End If
1090                             End If
1091                         Next tWindow
1092                         Set tWindows = Nothing
         End If

1093                         Screen.MousePointer = lMouseState

1094                         gbCancelInsertion = True
1095                         GoTo EH_InsertTemplate_Continue
1096                         On Error GoTo EH_InsertTemplate

1097                     Case "FIND", "LOCATE", "SEARCH"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1098                         lStartLine = II.PointOfInsertion
1099                         lStartColFound = 1
1100                         lEndLine = CurModule.CountOfLines    '- II.PointOfInsertion + 1
1101                         lEndColFound = -1
1102                         If CurModule.Find(II.AllParameters, lStartLine, lStartColFound, lEndLine, lEndColFound) Then
1103                             II.SoftVars("Found").Value = lStartLine & vbNullString
1104                         Else
1105                             II.SoftVars("Found").Value = "0"
1106                         End If
                    End If

1107                     Case "FINDINPROC", "PROCFIND", "PFIND", "PROCLOCATE", "PLOCATE", "PROCSEARCH", "PSEARCH"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1108                         On Error Resume Next          ' Prevent illegal values from causing an error
1109                         Err.Clear
1110                         GetProcAtLine II.PointOfInsertion, sProcName, lProcType
1111                         lStartLine = CurModule.ProcStartLine(sProcName, lProcType)
1112                         lEndLine = CurModule.ProcCountLines(sProcName, lProcType) + lStartLine - 1
1113                         lStartColFound = 1
1114                         lEndColFound = -1
1115                         If CurModule.Find(II.AllParameters, lStartLine, lStartColFound, lEndLine, lEndColFound) Then
1116                             II.SoftVars("Found").Value = lStartLine & vbNullString
1117                         Else
1118                             II.SoftVars("Found").Value = "0"
1119                         End If
                    End If

1120                     Case "DELETEPROC"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1121                         On Error Resume Next          ' Prevent illegal values from causing an error
1122                         Err.Clear
1123                         lStartLine = CurModule.ProcStartLine(sProcName, lProcType)
1124                         If Err.Number = 0 Then
1125                             lEndLine = CurModule.ProcCountLines(sProcName, lProcType)
1126                             If lStartLine > CurModule.CountOfDeclarationLines And lEndLine > 0 Then
1127                                 CurModule.DeleteLines lStartLine, lEndLine
1128                             End If
1129                         End If
                    End If

1130                     Case "COPYPROC", "CUTPROC"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1131                         On Error Resume Next          ' Prevent illegal values from causing an error
1132                         Err.Clear
1133                         lStartLine = CurModule.ProcStartLine(sProcName, lProcType)
1134                         If Err.Number = 0 Then
1135                             lEndLine = CurModule.ProcCountLines(sProcName, lProcType)
1136                             If lStartLine > CurModule.CountOfDeclarationLines And lEndLine > 0 Then
1137                                 sT = sAfter(II.sParam)
1138                                 If Len(sT) > 0 Then
1139                                     II.SoftVars(sT).Value = CurModule.Lines(lStartLine, lEndLine)
1140                                     If II.SoftCommandName = "CUTPROC" Then
1141                                         CurModule.DeleteLines lStartLine, lEndLine
1142                                     End If
1143                                 End If
1144                             End If
1145                         End If
                    End If

1146                     Case "DELETELINES"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1147                         CurModule.DeleteLines II.PointOfInsertion, II.ParamLineOffset
                    End If

1148                     Case "DELETELINE"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1149                         CurModule.DeleteLines II.PointOfInsertion
                    End If

1150                     Case "PROCATTR"                   ' Modify the current procedure's attributes (Ouch ! That's cool !)
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1151                         On Error Resume Next          ' Prevent illegal values from causing an error
                        Select Case UCase$(sProcName)
                            Case "ID"                 ' Set a default or NewEnum property
1152                                 GetProcAtLine II.PointOfInsertion, sProcName, lProcType
1153                                 If InStr(UCase$(II.sParam), "DEFAULT") > 0 Then
1154                                     CurModule.Members(sProcName).StandardMethod = 0
1155                                 ElseIf InStr(UCase$(II.sParam), "NEWENUM") > 0 Then
1156                                     CurModule.Members(sProcName).StandardMethod = -4
1157                                 Else
1158                                     CurModule.Members(sProcName).StandardMethod = Val(II.sParam)
1159                                 End If
1160                             Case "HIDDEN"             ' Hide/unhide the property
1161                                 GetProcAtLine II.PointOfInsertion, sProcName, lProcType
1162                                 CurModule.Members(sProcName).Hidden = IIf(UCase$(II.sParam) = "TRUE" Or UCase$(II.sParam) = "T", True, False)
1163                             Case "DESC"               ' Add a description to the property
1164                                 GetProcAtLine II.PointOfInsertion, sProcName, lProcType
1165                                 CurModule.Members(sProcName).Description = II.sParam
1166                         End Select
1167                         On Error GoTo EH_InsertTemplate    ' Resume normal error processing
                    End If

1168                     Case "READLINE"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1169                         On Error Resume Next
1170                         II.SoftVars(II.AllParameters).Value = CurModule.Lines(II.PointOfInsertion, 1)
1171                         On Error GoTo EH_InsertTemplate    ' Resume normal error processing
                    End If

1172                     Case "NEXTLINE"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1173                         On Error Resume Next
1174                         If II.PointOfInsertion <= CurModule.CountOfLines Then
1175                             II.SoftVars(II.AllParameters).Value = CurModule.Lines(II.PointOfInsertion, 1)
1176                             II.PointOfInsertion = II.PointOfInsertion + 1
1177                         Else
1178                             II.SoftVars(sGetToken(II.AllParameters)).Value = vbNullString
1179                         End If
1180                         On Error GoTo EH_InsertTemplate    ' Resume normal error processing
                    End If

1181                     Case "POSTFIXLINE", "POSTFIX"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1182                         If Len(II.AllParameters) Then
1183                             sT = CurModule.Lines(II.PointOfInsertion, 1)
1184                             CurModule.DeleteLines II.PointOfInsertion, 1
1185                             CurModule.InsertLines II.PointOfInsertion, sT & II.AllParameters
1186                         End If
                    End If

1187                     Case "PREFIXLINE", "PREFIX"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1188                         If Len(II.AllParameters) Then
1189                             sT = CurModule.Lines(II.PointOfInsertion, 1)
1190                             CurModule.DeleteLines II.PointOfInsertion, 1
1191                             CurModule.InsertLines II.PointOfInsertion, II.AllParameters & sT
1192                             II.PointOfInsertion = II.PointOfInsertion + lTokenCount(II.AllParameters, vbNewLine) - 1
1193                         End If
                    End If

1194                     Case "GETTEXTSELECTION", "GETTEXT", "GETSELECTION"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1195                         II.SoftVars(sGetToken(II.AllParameters)).Value = GetCurrentTextSelection
1196                         On Error GoTo EH_InsertTemplate
                    End If

1197                     Case "GETCLIPBOARDTEXT", "GETCLIPBOARD", "GETCLIP"
1198                         On Error Resume Next
1199                         II.SoftVars(sGetToken(II.AllParameters)).Value = Clipboard.GetText(vbCFText)
1200                         On Error GoTo EH_InsertTemplate

1201                     Case "SETCLIPBOARDTEXT", "SETCLIPBOARD", "SETCLIP"
1202                         On Error Resume Next
1203                         StringToClipboard II.SoftVars(sGetToken(II.AllParameters))
1204                         On Error GoTo EH_InsertTemplate

1205                     Case "REPLACELINE"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1206                         CurModule.ReplaceLine II.PointOfInsertion, II.AllParameters
                    End If


1207                     Case "REPLACEINMODULE", "MODULEREPLACE"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
                        ' FUTURE: Correct and use code commented out to specifically remove instances instead of en' mass
                        'lStartLine = 1
                        'lEndLine = CurModule.CountOfLines
                        'lStartColFound = 1
                        'lEndColFound = -1
                        '
                        'CmdIterations = 0
                        'Do
                        '   CmdIterations = CmdIterations + 1
                        '   bT = CurModule.Find(II.AllParameters, lStartLine, lStartColFound, lEndLine, lEndColFound)
                        '   If bT Then
                        '
                        '   End If
                        'Loop While bT And CmdIterations < 10000

1208                         lEndLine = CurModule.CountOfLines
1209                         If lEndLine > 1 Then
1210                             scT1 = CurModule.Lines(1, lEndLine)    ' Get the entire module
1211                             scT2 = II.SoftVars("ToFind")
1212                             If Len(scT2) > 0 Then
1213                                 scT3 = II.SoftVars("ReplaceWith")
1214                                 scT4 = Replace(scT1, scT2, scT3)
1215                                 If StrComp(scT1, scT4) <> 0 Then
1216                                     CurModule.DeleteLines 1, lEndLine
1217                                     CurModule.AddFromString scT4
1218                                 End If
1219                             End If
1220                             scT1 = vbNullString
1221                             scT4 = vbNullString
1222                         End If
                    End If

1223                     Case "DELETESELECTION"
1224                         DeleteCurrentTextSelection

1225                     Case "LASTSELECTIONLINE"
1226                         II.PointOfInsertion = DetermineLastLineInSelection

1227                     Case "FIRSTSELECTIONLINE"
1228                         II.PointOfInsertion = DetermineFirstLineInSelection

1229                     Case "SELECTCONTROL"
         If Parent.HostedByVB Then  ' Shell App override
1230                         On Error Resume Next
1231                         Set CurForm = Parent.vbInst.SelectedVBComponent.Designer
1232                         Set II.CurrControl = CurForm.VBControls(sGetToken(II.AllParameters, 1))
1233                         On Error GoTo EH_InsertTemplate
         End If

1234                     Case "ADDCONTROL"
         If Parent.HostedByVB Then  ' Shell App override
1235                         On Error Resume Next
1236                         If InStr(sGetToken(II.AllParameters), gsP) > 0 And InStr(sGetToken(II.AllParameters, 2), gsP) = 0 Then
1237                             II.AllParameters = sGetToken(II.AllParameters, 2) & gsS & sGetToken(II.AllParameters)
1238                         End If
1239                         Set CurForm = Parent.vbInst.SelectedVBComponent.Designer
1240                         Set II.CurrControl = CurForm.VBControls.Add(sGetToken(II.AllParameters, 2))
1241                         If II.CurrControl Is Nothing Then
1242                             MsgBox "The '" & II.sParam & "' control has not been referenced yet. Please add a reference first.", vbInformation
1243                             gbCancelInsertion = bUserSure("Cancel processing ?")
1244                             If gbCancelInsertion Then GoTo EH_InsertTemplate_Continue
1245                         Else
1246                             II.CurrControl.Properties("Name") = sGetToken(II.AllParameters)
1247                         End If
1248                         On Error GoTo EH_InsertTemplate
         End If

1249                     Case "SETPROPERTY"
1250                         On Error Resume Next
1251                         If Not II.CurrControl Is Nothing Then
1252                             II.CurrControl.Properties(II.Result) = II.Expression
1253                         End If
1254                         On Error GoTo EH_InsertTemplate

1255                     Case "ADDFILEREFERENCE", "ADDFILEREF"
         If Parent.HostedByVB Then  ' Shell App override
1256                         On Error Resume Next
1257                         Err.Clear
1258                         bFoundReference = False
1259                         For Each CurReference In Parent.vbInst.ActiveVBProject.References
1260                             If UCase$(CurReference.FullPath) = UCase$(II.AllParameters) Then
1261                                 bFoundReference = True
1262                             End If
1263                         Next CurReference
1264                         If Not bFoundReference Then
1265                             Parent.vbInst.ActiveVBProject.References.AddFromFile II.AllParameters
1266                         End If
                        'Parent.vbInst.ActiveVBProject.AddToolboxProgID "FirmSolutionsDV.DataView" ', II.AllParameters
1267                         If Err.Number <> 0 Then
1268                             MsgBox "Failed to add a reference/component by Filename '" & II.AllParameters & gsA
1269                             Err.Clear
1270                             gbCancelInsertion = bUserSure("Cancel processing ?")
1271                             If gbCancelInsertion Then GoTo EH_InsertTemplate_Continue
1272                         End If
1273                         On Error GoTo EH_InsertTemplate
         End If

1274                     Case "ADDFILE", "INCLUDEFILE", "INCLUDE"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1275                         On Error Resume Next
1276                         Err.Clear
1277                         CurModule.AddFromFile II.AllParameters
1278                         If Err.Number <> 0 Then
1279                             MsgBox "Failed to include the File '" & II.AllParameters & gsA
1280                             Err.Clear
1281                             gbCancelInsertion = bUserSure("Cancel processing ?")
1282                             If gbCancelInsertion Then GoTo EH_InsertTemplate_Continue
1283                         End If
1284                         On Error GoTo EH_InsertTemplate
                    End If

1285                     Case "ADDREFERENCE", "ADDCOMPONENT", "ADDREF"
         If Parent.HostedByVB Then  ' Shell App override
1286                         On Error Resume Next
1287                         If Left$(II.AllParameters, 1) = gsBO Then
                            ' GUID Passed
1288                             Err.Clear
1289                             Parent.vbInst.ActiveVBProject.References.AddFromGuid sGetToken(II.AllParameters), CLng(sGetToken(II.AllParameters, 2)), CLng(sGetToken(II.AllParameters, 3))
1290                             If Err.Number <> 0 Then
1291                                 MsgBox "Failed to add a reference/component by GUID '" & II.AllParameters & gsA
1292                                 Err.Clear
1293                                 gbCancelInsertion = bUserSure("Cancel processing ?")
1294                                 If gbCancelInsertion Then GoTo EH_InsertTemplate_Continue
1295                             End If
1296                         Else
                            '   ' ProgID Passed
1297                             Err.Clear
1298                             Parent.vbInst.ActiveVBProject.AddToolboxProgID sGetToken(II.AllParameters)
1299                             If Err.Number <> 0 Then
                                ' Attempt to add a reference by looking up the GUID for the ProgID passed
1300                                 Err.Clear
1301                                 sHold1 = sGetGUID(sGetToken(II.AllParameters))
1302                                 If Len(sHold1) > 0 Then
1303                                     Err.Clear
1304                                     Parent.vbInst.ActiveVBProject.References.AddFromGuid sHold1, 0, 0    'CLng(sGetToken(II.AllParameters, 2)), CLng(sGetToken(II.AllParameters, 3))
1305                                 End If
1306                                 If Err.Number <> 0 Then
1307                                     MsgBox "Failed to add a reference/component by ProgID '" & II.AllParameters & gsA
1308                                     Err.Clear
1309                                     gbCancelInsertion = bUserSure("Cancel processing ?")
1310                                     If gbCancelInsertion Then GoTo EH_InsertTemplate_Continue
1311                                 End If
1312                             End If
1313                         End If
1314                         On Error GoTo EH_InsertTemplate
         End If

1315                     Case "SETFORMPROPERTY", "FORMPROPERTY"
         If Parent.HostedByVB Then  ' Shell App override
1316                         On Error Resume Next
1317                         Parent.vbInst.SelectedVBComponent.Properties(II.Result) = II.Expression
1318                         On Error GoTo EH_InsertTemplate
         End If

                        ' ================================================
                        ' Soft commands that change the point of insertion
                        ' ================================================
1319                     Case "GOTOPROJECT"
         If Parent.HostedByVB Then  ' Shell App override
1320                         On Error Resume Next
1321                         If FindInCollection(Parent.vbInst.VBProjects, sProcName) Is Nothing Then
                            Select Case UCase$(II.sParam)
                                Case "CONTROL"
1322                                     With Parent.vbInst.VBProjects.Add(vbext_pt_ActiveXControl, False)
1323                                         .Name = sProcName
1324                                         .VBComponents(1).Activate
1325                                         .VBComponents(1).CodeModule.CodePane.Show
1326                                     End With

1327                                 Case "EXE"
1328                                     With Parent.vbInst.VBProjects.Add(vbext_pt_StandardExe, False)
1329                                         .Name = sProcName
1330                                         .VBComponents(1).Activate
1331                                         .VBComponents(1).CodeModule.CodePane.Show
1332                                     End With

1333                                 Case "ACTIVEXEXE"
1334                                     With Parent.vbInst.VBProjects.Add(vbext_pt_ActiveXExe, False)
1335                                         .Name = sProcName
1336                                         .VBComponents(1).Activate
1337                                         .VBComponents(1).CodeModule.CodePane.Show
1338                                     End With

1339                                 Case Else             ' "DLL"
1340                                     With Parent.vbInst.VBProjects.Add(vbext_pt_ActiveXDll, False)
1341                                         .Name = sProcName
1342                                         .VBComponents(1).Activate
1343                                         .VBComponents(1).CodeModule.CodePane.Show
1344                                     End With
1345                             End Select

1346                         Else
1347                             With Parent.vbInst.VBProjects(sProcName)
1348                                 .VBComponents(1).CodeModule.CodePane.Show
1349                             End With
1350                         End If

1351                         II.SoftVars("Project Name").Value = Parent.vbInst.ActiveVBProject.Name    ' Add the build in soft variables
1352                         II.SoftVars("Module Name").Value = Parent.vbInst.ActiveCodePane.CodeModule.Parent.Name
1353                         II.SoftVars("Module Lines").Value = Parent.vbInst.ActiveCodePane.CodeModule.CountOfLines
1354                         II.SoftVars("Module End of Declarations").Value = Parent.vbInst.ActiveCodePane.CodeModule.CountOfDeclarationLines + 1

1355                         II.SoftVars("Proc Name").Value = vbNullString
1356                         II.SoftVars("Proc Type").Value = vbNullString
1357                         II.SoftVars("Proc Type Long").Value = vbNullString

1358                         Set CurModule = Parent.vbInst.ActiveCodePane.CodeModule
1359                         On Error GoTo EH_InsertTemplate
         End If

1360                     Case "GOTOMODULE", "GOTOCLASS", "GOTOFORM"
         If Parent.HostedByVB Then  ' Shell App override
1361                         If FindInCollection(Parent.vbInst.ActiveVBProject.VBComponents, sProcName) Is Nothing Then
1362                             II.sParam = UCase$(II.sParam)
1363                             If Len(II.sParam) = 0 Then
1364                                 II.sParam = UCase$(Mid$(sGetToken(II.CurrentLineToProcess), 5))
1365                             End If
                            Select Case UCase$(II.sParam)
                                Case "CLASS", "CLASSMODULE"
1366                                     With Parent.vbInst.ActiveVBProject.VBComponents.Add(vbext_ct_ClassModule)
1367                                         .Name = sProcName
1368                                         .Activate
1369                                     End With
1370                                 Case "FORM"
1371                                     With Parent.vbInst.ActiveVBProject.VBComponents.Add(vbext_ct_VBForm)
1372                                         .Name = sProcName
1373                                         .Activate
1374                                         .CodeModule.CodePane.Show
1375                                     End With
1376                                 Case Else             '"MODULE"
1377                                     With Parent.vbInst.ActiveVBProject.VBComponents.Add(vbext_ct_StdModule)
1378                                         .Name = sProcName
1379                                         .Activate
1380                                     End With
1381                             End Select
1382                         Else
1383                             With Parent.vbInst.ActiveVBProject.VBComponents(sProcName)
1384                                 .Activate
1385                                 .CodeModule.CodePane.Show
1386                             End With
1387                         End If
1388                         II.SoftVars("Module Name").Value = sProcName

1389                         Set CurModule = Parent.vbInst.ActiveCodePane.CodeModule
1390                         II.SoftVars("Module Lines").Value = Parent.vbInst.ActiveCodePane.CodeModule.CountOfLines
1391                         II.SoftVars("Module End of Declarations").Value = Parent.vbInst.ActiveCodePane.CodeModule.CountOfDeclarationLines + 1
         End If

1392                     Case "GOTOPROC"                   ' Set the current line to insert to the indicated line in the indicated procedure
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1393                         If II.ParamLineOffset = 0 Then II.ParamLineOffset = 1
On Error Resume Next
                             Err.Clear
                             II.PointOfInsertion = CurModule.ProcBodyLine(sProcName, lProcType) + II.ParamLineOffset
                             If Err.Number = 35 Then
                                If LogError("frmMain", "InternalInsertTemplate", 35, "Unable to find procedure '" & sProcName & "', procedure type '" & lProcType & "', in '" & sGetToken(CurModule.CodePane.Window.Caption, 1, " (") & "'", Erl, "Cancel Insertion ?") Then
                                   gbCancelInsertion = True
                                End If
                             End If
On Error GoTo EH_InsertTemplate

                        'GetProcAtLine II.PointOfInsertion, sProcName, lProcType
1395                         If Len(sProcName) Then
1396                             II.SoftVars("Proc Name").Value = sProcName
1397                             II.SoftVars("Proc Type").Value = Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
1398                             sHold2 = CurModule.Lines(CurModule.ProcBodyLine(sProcName, lProcType), 1)
1399                             If InStr(sHold2, "Function") > 0 Then
1400                                 sHold2 = "Function"
1401                             ElseIf InStr(sHold2, "Property") > 0 Then
1402                                 sHold2 = "Property"
1403                             Else
1404                                 sHold2 = "Sub"
1405                             End If
1406                             II.SoftVars("Proc Type Long").Value = sHold2
1407                         Else
1408                             II.SoftVars("Proc Name").Value = vbNullString
1409                             II.SoftVars("Proc Type").Value = vbNullString
1410                             II.SoftVars("Proc Type Long").Value = vbNullString
1411                         End If
                    End If

                         Case "FILENAME", "GOTOFILE", "OUTPUTTOFILE", "SENDTOFILE", "FILEOUT"      ' Set an external filename to output to
                              ChangeFocusOfInsertion II, II.AllParameters

                         Case "GOTOMESSAGE", "GOTOSHOWMESSAGE", "OUTPUTTOMESSAGE", "MESSAGEOUT", "MSGOUT"
                              ChangeFocusOfInsertion II, "**MESSAGEWINDOW**"

                         Case "GOTOKB", "GOTOKEYBOARD", "OUTPUTTOKEYBOARD", "OUTPUTTOKB", "OUTTOKB", "OUTTOKEYBOARD", "KEYOUT"
                              If UCase$(II.AllParameters) = "RAW" Or II.AllParameters = "1" Then
                                 ChangeFocusOfInsertion II, "**KEYBOARDRAW**"
                              Else
                                 ChangeFocusOfInsertion II, "**KEYBOARD**"
                              End If

                         Case "GOTOCLIPBOARD", "GOTOCLIP", "OUTPUTTOCLIPBOARD", "OUTTOCLIP", "OUTTOCLIPBOARD", "CLIPOUT"
                              ChangeFocusOfInsertion II, "**CLIPBOARD**"

                         Case "GOTOSOFTVARIABLE", "GOTOSOFTVAR", "GOTOVAR"
                              ChangeFocusOfInsertion II, "**SOFTVARARIABLE** " & II.AllParameters

                         Case "GOTOSOFTCODE", "SOFTCODE"
                              ChangeFocusOfInsertion II, "**SOFTCODE**"

                         Case "OVERWRITEFILE", "GOTONEWFILE", "NEWFILE", "OVERWRITE"
                              ChangeFocusOfInsertion II, "**OVERWRITE** " & II.AllParameters

1412                     Case "GOTOPROCEND"                ' Set the current line to the last line before "End Sub/Function/Property" in the indicated procedure
                          If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1413                         II.PointOfInsertion = FindLastProcLine(sProcName, lProcType) + II.ParamLineOffset
1414                         If Len(sProcName) Then
1415                             II.SoftVars("Proc Name").Value = sProcName
1416                             II.SoftVars("Proc Type").Value = Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
1417                             sHold2 = CurModule.Lines(CurModule.ProcBodyLine(sProcName, lProcType), 1)
1418                             If InStr(sHold2, "Function") > 0 Then
1419                                 sHold2 = "Function"
1420                             ElseIf InStr(sHold2, "Property") > 0 Then
1421                                 sHold2 = "Property"
1422                             Else
1423                                 sHold2 = "Sub"
1424                             End If
1425                             II.SoftVars("Proc Type Long").Value = sHold2
1426                         Else
1427                             II.SoftVars("Proc Name").Value = vbNullString
1428                             II.SoftVars("Proc Type").Value = vbNullString
1429                             II.SoftVars("Proc Type Long").Value = vbNullString
1430                         End If
                    End If

1431                     Case "ABSLINE"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1432                         II.PointOfInsertion = Abs(II.ParamLineOffset)    ' Set the current line to the absolute line number specified
1433                         GetProcAtLine II.PointOfInsertion, sProcName, lProcType
1434                         If Len(sProcName) Then
1435                             II.SoftVars("Proc Name").Value = sProcName
1436                             II.SoftVars("Proc Type").Value = Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
1437                             sHold2 = CurModule.Lines(CurModule.ProcBodyLine(sProcName, lProcType), 1)
1438                             If InStr(sHold2, "Function") > 0 Then
1439                                 sHold2 = "Function"
1440                             ElseIf InStr(sHold2, "Property") > 0 Then
1441                                 sHold2 = "Property"
1442                             Else
1443                                 sHold2 = "Sub"
1444                             End If
1445                             II.SoftVars("Proc Type Long").Value = sHold2
1446                         Else
1447                             II.SoftVars("Proc Name").Value = vbNullString
1448                             II.SoftVars("Proc Type").Value = vbNullString
1449                             II.SoftVars("Proc Type Long").Value = vbNullString
1450                         End If
                    End If

1451                     Case "LINEOFFSET", "OFFSET"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1452                         II.PointOfInsertion = II.PointOfInsertion + II.ParamLineOffset    ' Set the current line to the relative line offset specified
1453                         GetProcAtLine II.PointOfInsertion, sProcName, lProcType
1454                         If Len(sProcName) Then
1455                             II.SoftVars("Proc Name").Value = sProcName
1456                             II.SoftVars("Proc Type").Value = Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
1457                             sHold2 = CurModule.Lines(CurModule.ProcBodyLine(sProcName, lProcType), 1)
1458                             If InStr(sHold2, "Function") > 0 Then
1459                                 sHold2 = "Function"
1460                             ElseIf InStr(sHold2, "Property") > 0 Then
1461                                 sHold2 = "Property"
1462                             Else
1463                                 sHold2 = "Sub"
1464                             End If
1465                             II.SoftVars("Proc Type Long").Value = sHold2
1466                         Else
1467                             II.SoftVars("Proc Name").Value = vbNullString
1468                             II.SoftVars("Proc Type").Value = vbNullString
1469                             II.SoftVars("Proc Type Long").Value = vbNullString
1470                         End If
                    End If

1471                     Case "PROCTOP"                    ' Move to the top of the current procedure
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1472                         GetProcAtLine II.PointOfInsertion, sProcName, lProcType
1473                         If Len(sProcName) Then
1474                             II.PointOfInsertion = CurModule.ProcBodyLine(sProcName, lProcType)
1475                             II.SoftVars("Proc Name").Value = sProcName
1476                             II.SoftVars("Proc Type").Value = Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
1477                             sHold2 = CurModule.Lines(CurModule.ProcBodyLine(sProcName, lProcType), 1)
1478                             If InStr(sHold2, "Function") > 0 Then
1479                                 sHold2 = "Function"
1480                             ElseIf InStr(sHold2, "Property") > 0 Then
1481                                 sHold2 = "Property"
1482                             Else
1483                                 sHold2 = "Sub"
1484                             End If
1485                             II.SoftVars("Proc Type Long").Value = sHold2
1486                         Else
1487                             II.SoftVars("Proc Name").Value = vbNullString
1488                             II.SoftVars("Proc Type").Value = vbNullString
1489                             II.SoftVars("Proc Type Long").Value = vbNullString
1490                         End If
                    End If

1491                     Case "PROCEND"                    ' Move to the end of the current procedure
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1492                         GetProcAtLine II.PointOfInsertion, sProcName, lProcType
1493                         If Len(sProcName) Then
1494                             II.PointOfInsertion = FindLastProcLine(sProcName, lProcType)
1495                             II.SoftVars("Proc Name").Value = sProcName
1496                             II.SoftVars("Proc Type").Value = Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
1497                             sHold2 = CurModule.Lines(CurModule.ProcBodyLine(sProcName, lProcType), 1)
1498                             If InStr(sHold2, "Function") > 0 Then
1499                                 sHold2 = "Function"
1500                             ElseIf InStr(sHold2, "Property") > 0 Then
1501                                 sHold2 = "Property"
1502                             Else
1503                                 sHold2 = "Sub"
1504                             End If
1505                             II.SoftVars("Proc Type Long").Value = sHold2
1506                         Else
1507                             II.SoftVars("Proc Name").Value = vbNullString
1508                             II.SoftVars("Proc Type").Value = vbNullString
1509                             II.SoftVars("Proc Type Long").Value = vbNullString
1510                         End If
                    End If

1511                     Case "GOTODECLARATIONS", "GOTODEC"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1512                         If UCase$(II.AllParameters) = "END" Then
1513                             II.PointOfInsertion = CurModule.CountOfDeclarationLines + 1
1514                         Else                          ' Line 1
1515                             II.PointOfInsertion = 1
1516                         End If
1517                         II.SoftVars("Proc Name").Value = vbNullString
1518                         II.SoftVars("Proc Type").Value = vbNullString
1519                         II.SoftVars("Proc Type Long").Value = vbNullString
                    End If

1520                     Case "GOTOENDOFFILE", "GOTOENDOFMODULE", "GOTOEND"
                    If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1521                         II.PointOfInsertion = CurModule.CountOfLines + 1
1522                         II.SoftVars("Proc Name").Value = vbNullString
1523                         II.SoftVars("Proc Type").Value = vbNullString
1524                         II.SoftVars("Proc Type Long").Value = vbNullString
                    End If

                        ' ========================================*
                        ' Soft commands that affect the File System
                        ' ========================================*
                        '                        Case "DELETEFILE"                                               ' Causes a file in the operating system to be erased.
                        ' On Error Resume Next
                        '                             Kill II.AllParameters
                        '                             Err.Clear
                        ' On Error GoTo EH_InsertTemplate

                         Case "SHOWMESSAGE"
                              If Len(II.AllParameters) < 100 Then
                                 If Len(II.SoftVars(II.AllParameters)) Then
                                    ShowMessage II.SoftVars(II.AllParameters)
                                 Else
                                    ShowMessage II.AllParameters
                                 End If
                              Else
                                 ShowMessage II.AllParameters
                              End If

1537                     Case "IGNOREBLANKS", "BLANKSOKAY", "NOBLANKS"
1538                          mbIgnoreBlanks = True

1539                     Case "WATCHBLANKS", "BLANKSNOTOKAY", "YESBLANKS"
1540                          mbIgnoreBlanks = False

1541                     Case "IGNOREREADONLY"
1542                          mbIgnoreReadOnly = True

1543                     Case "WATCHREADONLY"
1544                          mbIgnoreReadOnly = False

1545                     Case Else
1546                          If SadCommandSetCount > 0 Then
1547                             For CurrSet = 1 To SadCommandSetCount
1548                                 If SadCommands(CurrSet).ExecuteSoftCommand(II) Then Exit For
1551                             Next CurrSet
1552                          End If
1553                 End Select

                '
                ' ======== End of soft command processing
                '

1554             Else
1555                 If II.PointOfInsertion < 1 Then II.PointOfInsertion = 1
1556                 If Len(II.ExternalFilename) = 0 Then
                        If (Not CurModule Is Nothing) And Parent.HostedByVB Then
1557                       CurModule.InsertLines II.PointOfInsertion, II.CurrentLineToProcess
                        End If
1558                 Else
1559                    II.TextToSendToFile = II.TextToSendToFile & II.CurrentLineToProcess & vbNewLine
1560                 End If
1561                 II.PointOfInsertion = II.PointOfInsertion + 1
1562             End If

1563             If StrComp(II.LinesLeftToProcess, vbNewLine) = 0 Or StrComp(II.LinesLeftToProcess, gs2EOL) = 0 Then
1564                 II.LinesLeftToProcess = vbNullString
1565             End If

1566             If Left$(II.LinesLeftToProcess, 2) = vbNewLine Then    ' Strip off the line just parsed
1567                 II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, 3)
1568             Else
1569                 II.LinesLeftToProcess = sAfter(II.LinesLeftToProcess, 1, vbNewLine)
1570             End If
1571         Loop
1572     End If
    'End With

       ' Flush any remaining Focus of Insertion buffer,
         ChangeFocusOfInsertion II, ""

       ' Continue if flushing the Focus of Insertion buffer caused new lines
       ' to be put onto the process buffer
         If Len(II.LinesLeftToProcess) Then
            GoTo CODA_RESTART
         End If

1582     InternalInsertTemplate = True

1583 EH_InsertTemplate_Continue:
1584     On Error Resume Next
1585     Set II.CurrControl = Nothing
    'Set II = Nothing
1586     Exit Function

1587 EH_InsertTemplate:
1588     If Err.Number = 40198 And mbIgnoreReadOnly Then
1589         Resume Next
1590     Else
1591         If LogError("frmMain", "InsertTemplate", Err.Number, Err.Description, Erl, "Cancel Insertion ?") Then
1592             gbCancelInsertion = True
1593             GoTo EH_InsertTemplate_Continue
1594         Else
1595             Resume Next
1596         End If
1597     End If
1598     Err.Clear
1599     Resume EH_InsertTemplate_Continue

1600     Resume
End Function

Public Function ShowMessage(ByVal sMessageToShow As String, Optional ByVal sTitle As String = "Slice and Dice - Message", Optional ByVal sToolTip As String) As String
On Error Resume Next
    Dim xWidth  As Single
    Dim xHeight As Single

    With frmMessage
        .Caption = sTitle
        With .txtMessage
             .Text = sMessageToShow
             .ToolTipText = sToolTip
             .SelStart = 0
             .SelLength = Len(.Text) + 1
        End With

        xWidth = .TextWidth(sMessageToShow) + 500
        xHeight = .TextHeight(sMessageToShow) + 500
        
        xWidth = xWidth + .TextWidth("Q") * 3
        xHeight = xHeight + .TextHeight("QWERTY") * 2

        If xWidth > Screen.Width - 1000 Then xWidth = Screen.Width - 1000
        If xWidth < 1000 Then xWidth = 1000

        If xHeight > Screen.Height - 1000 Then xHeight = Screen.Height - 1000
        If xHeight < 1000 Then xHeight = 1000

        .Width = xWidth
        .Height = xHeight
        
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
        
        If .Top < 1000 Then .Top = 50
        If .Left < 1000 Then .Left = 50

        .Show vbModal, Me
    End With
End Function


' ================================================================================
' Name              frmMain_GetProcAtLine
'
' Parameters
'      lCurrentLine                  (I)  Long
'      sProcName                     (O)  String
'      lProcType                     (O)  Integer
'
' Description
'
' Returns the Procedure name and type the given line is included in (If any).
'
' ================================================================================
Public Sub GetProcAtLine(ByRef lCurrentLine As Long, ByRef sProcName As String, ByRef lProcType As Long)
1601     Dim ProcType As vbext_ProcKind

         If Not Parent.HostedByVB Then Exit Sub   ' Shell override

1602     If lCurrentLine < 1 Or lCurrentLine > Parent.vbInst.ActiveCodePane.CodeModule.CountOfLines Then
1603         sProcName = vbNullString
1604         lProcType = 0
1605         Exit Sub
1606     End If

1607     With Parent.vbInst.ActiveCodePane.CodeModule
1608         sProcName = .ProcOfLine(lCurrentLine, ProcType)
1609         lProcType = ProcType
1610     End With
End Sub

' ================================================================================
' Name              frmMain_FindLastProcLine
'
' Parameters
'      sProcName                     (O)  String
'      lProcType                     (O)  Integer
'
' Description
'
' Determines what the last line number is for a given procedure.
'
' ================================================================================
Public Function FindLastProcLine(sProcName As String, lProcType As Long) As Long
1611     Static lLine           As Long
1612     Static lCurLine        As Long
1613     Static lLastLine       As Long
1614     Static sFindString     As String
1615     Static sFunctionHeader As String

         If Not Parent.HostedByVB Then Exit Function   ' Shell override

1616     On Error Resume Next
1617     Err.Clear
1618     With Parent.vbInst.ActiveCodePane.CodeModule
1619         lLine = .ProcStartLine(sProcName, lProcType)  ' Get the first line number of the procedure
1620         If Err.Number <> 0 Then
1621             If InStr(sProcName, "_") Then
1622                 Err.Clear
1623                 lLine = .CreateEventProc(sGetToken(sProcName, 2, "_"), sGetToken(sProcName, 1, "_"))
1624                 If Err.Number <> 0 Then
1625                     MsgBox "FindLastProcLine" & gsEolTab & "Can't find:" & gsEolTab & vbTab & sProcName & gsEolTab & "In module:" & gsEolTab & vbTab & .Parent.Name
1626                     gbCancelInsertion = bUserSure("Cancel processing ?")
1627                     FindLastProcLine = 0
1628                     Err.Clear
1629                     Exit Function
1630                 End If
1631             Else
1632                 MsgBox "FindLastProcLine" & gsEolTab & "Can't find:" & gsEolTab & vbTab & sProcName & gsEolTab & "In module:" & gsEolTab & vbTab & .Parent.Name
1633                 gbCancelInsertion = bUserSure("Cancel processing ?")
1634                 FindLastProcLine = 0
1635                 Err.Clear
1636                 Exit Function
1637             End If
1638         End If
1639         lLastLine = lLine + .ProcCountLines(sProcName, lProcType)    ' Get the last line number of the procedure
1640         sFunctionHeader = .Lines(.ProcBodyLine(sProcName, lProcType), 1)    ' Get the procedure's header
1641         If InStr(sFunctionHeader, "Function") > 0 Then ' Based on it's type,
1642             sFindString = "End Function"              '   we can determine what string to look for
1643         ElseIf InStr(sFunctionHeader, "Sub") > 0 Then
1644             sFindString = "End Sub"
1645         Else
1646             sFindString = "End Property"
1647         End If

1648         For lCurLine = lLastLine To lLine Step -1     ' Move backwards from the end of function
1649             If InStr(.Lines(lCurLine, 1), sFindString) > 0 Then    '   until we find the line containing
1650                 FindLastProcLine = lCurLine           '   "End Function/Sub/Property"
1651                 sFindString = vbNullString
1652                 sFunctionHeader = vbNullString
1653                 Exit Function                         ' At this point, we found it, return THAT line #
1654             End If
1655         Next lCurLine
1656     End With

1657     sFindString = vbNullString
1658     sFunctionHeader = vbNullString
1659     FindLastProcLine = lLastLine                      ' Something wrong: Pass back last line # found

End Function


Public Function JumpTo(ByVal sTemplateName As String, Optional ByVal bRecordInHistory As Boolean = True, Optional ByVal bSyncCategoryList As Boolean = False) As Boolean
1660     On Error GoTo frmMain_EH_JumpTo
1661     Static sCategoryName As String
1662     Static sShortTemplateName As String
1663     Static CurrHE As Long

1664     Err.Clear
1665     SaveTemplate

1666     sCategoryName = sGetToken(sTemplateName, 1, gsCategoryTemplateDelimiter)
1667     sShortTemplateName = sAfter(sTemplateName, 1, gsCategoryTemplateDelimiter)
1668     On Error Resume Next
1669     If SliceAndDice.Categorys(sCategoryName).Templates(sTemplateName) Is Nothing Then
1670         JumpTo = False
1671         Exit Function
1672     End If
1673     Set CurrentTemplate = SliceAndDice(sCategoryName).Templates(sTemplateName)
1674     FillAddInScreen

1675     On Error GoTo frmMain_EH_JumpTo

    'CodeAtTop = txtCode(0).Text
1676     tabCode.Tabs(1).Image = IIf(Len(txtCode(0)) = 0, "Document", gsCategory)
1677     tabCode.Tabs(2).Image = IIf(Len(txtCode(1)) = 0, "Document", gsCategory)
1678     tabCode.Tabs(3).Image = IIf(Len(txtCode(2)) = 0, "Document", gsCategory)
1679     tabCode.Tabs(4).Image = IIf(Len(txtCodeToFile) = 0, "Document", gsCategory)
1680     tabCode.Tabs(5).Image = IIf(CurrentTemplate.Undeletable Or CurrentTemplate.Locked Or CurrentTemplate.Selected, "OptionSet", "OptionNotSet")

1681     If mnuSwitchTabsAutomatically.Checked Then
1682         If tabCode.Tabs(1).Image = gsCategory Then
1683             tabCode.Tabs(1).Selected = True
1684         ElseIf tabCode.Tabs(2).Image = gsCategory Then
1685             tabCode.Tabs(2).Selected = True
1686         ElseIf tabCode.Tabs(3).Image = gsCategory Then
1687             tabCode.Tabs(3).Selected = True
1688         ElseIf tabCode.Tabs(4).Image = gsCategory Then
1689             tabCode.Tabs(4).Selected = True
1690         Else
1691             tabCode.Tabs(1).Selected = True
1692         End If

1693         chkLocked_Click
1694         If chkLocked.Value = 0 Then
1695             txtShortName.Enabled = (sCategoryName & gsCategoryTemplateDelimiter & sShortTemplateName <> "Change from - All Types")
1696         Else
1697             txtShortName.Enabled = False
1698         End If

1699         tabCode_MouseUp 0, 0, 0, 0
1700     Else
1701         If tabCode.Tabs(6).Selected Then
1702             If chkAutoRecalc.Value <> 0 Then
1703                 cmdRecalc_Click
1704             End If
1705         End If
1706     End If

1707     If bRecordInHistory Then
1708         For CurrHE = Val(CurrentHistoryEntry) + 1 To m_asaHistory.Count
1709             m_asaHistory.Remove vbNullString & CurrHE
1710         Next CurrHE
1711         CurrentHistoryEntry = vbNullString & m_asaHistory.Count + 1
1712         m_asaHistory.Add vbNullString & m_asaHistory.Count + 1, sTemplateName
1713         mnuBack.Enabled = True
1714         mnuForward.Enabled = False
1715     End If

1716     If bSyncCategoryList Then
1717         lsbJumpTo.BarAndItem sCategoryName, sTemplateName
1718     End If

1719     JumpTo = True

1720 frmMain_EH_JumpTo_Continue:
1721     Exit Function

1722 frmMain_EH_JumpTo:
1723     MsgBox "Error occured in:" & gsEolTab & "Module: frmMain" & gsEolTab & "Procedure: JumpTo" & gs2EOL & Err.Description
1724     JumpTo = False
1725     Resume frmMain_EH_JumpTo_Continue

1726     Resume
End Function

Public Sub RefillList()
1727     On Error GoTo EH_frmMain_RefillList
1728     Static bInHereAlready As Boolean
1729     If bInHereAlready Then Exit Sub
1730     bInHereAlready = True

1731     Dim lvwX As Object
1732     Dim tvwX As Object
1733     Dim tvwY As Object                                'TreeView
1734     Dim CurrCategory As CCategory
1735     Dim CurrTemplate As CTemplate
1736     Dim sOpened As String
1737     Dim sClosed As String

1738     lsbJumpTo.Visible = False
1739     lsbJumpTo.Clear
1740     For Each CurrCategory In SliceAndDice.Categorys
1741         With CurrCategory
1742             If .Deleted Then
                ' Ignore this one
                'Else
1743             ElseIf CurrCategory.CategoryType = 0 Then
1744                 If lsbJumpTo.BarKey = "Bar 1" Then
1745                     lsbJumpTo.CurBar = 0
1746                     lsbJumpTo.BarName = .Key & " (" & Format$(CurrCategory.Templates.Count, "00") & gsPC
1747                     lsbJumpTo.BarKey = .Key
1748                     lsbJumpTo.View = 3
1749                     lsbJumpTo.Arrange = .Arrange
1750                     lsbJumpTo.BarType = "List"
1751                     On Error Resume Next
1752                     lsbJumpTo.Bars(0).ColumnHeaders(1).Width = 3400
1753                 Else
1754                     lsbJumpTo.AddBar(.Key & " (" & Format$(CurrCategory.Templates.Count, "00") & gsPC, .Key).ColumnHeaders(1).Width = 3400
1755                 End If
1756             Else
1757                 If lsbJumpTo.BarKey = "Bar 1" Then
1758                     lsbJumpTo.CurBar = 0
1759                     lsbJumpTo.BarType = "Tree"
1760                     lsbJumpTo.BarName = .Key & " (Code Gen)"
1761                     lsbJumpTo.BarKey = .Key
1762                 Else
1763                     lsbJumpTo.AddBar "[" & CurrCategory.Templates.Count & "] " & .Key, .Key, False
1764                 End If
1765             End If
1766         End With
1767     Next CurrCategory

1768     For Each CurrCategory In SliceAndDice.Categorys
1769         If Not CurrCategory.Deleted Then
1770             lsbJumpTo.CurBar = CurrCategory.Key
1771             If CurrCategory.CategoryType = 0 Then
1772                 For Each CurrTemplate In CurrCategory.Templates
1773                     With CurrTemplate
1774                         If .Deleted Then
                            ' Ignore this one
1775                         ElseIf .Locked Or .Undeletable Then
1776                             lsbJumpTo.AddBarItem .ShortTemplateName, .Key, "Key"
1777                             .OriginalShortName = .ShortTemplateName
1778                         ElseIf Len(.memoCodeAtBottom & .memoCodeAtCursor & .memoCodeAtTop & .memoCodeToFile) Then
1779                             lsbJumpTo.AddBarItem .ShortTemplateName, .Key, "DocumentAlternate"
1780                             .OriginalShortName = .ShortTemplateName
1781                         Else
1782                             lsbJumpTo.AddBarItem .ShortTemplateName, .Key, "Document"
1783                             .OriginalShortName = .ShortTemplateName
1784                         End If
1785                     End With
1786                 Next CurrTemplate
1787             Else
1788                 Set tvwX = lsbJumpTo.Bars(CurrCategory.Key)
1789                 Set tvwY = tvwX
1790                 With tvwY.Nodes
1791                     sOpened = sTemplateIcon(CurrCategory.Templates("Settings"))
1792                     With .Add(, , CurrCategory.Key & " - Settings", "Settings", sOpened, sOpened)
1793                         .ExpandedImage = sOpened
1794                         .Expanded = True
                             .ToolTip = "'Entire Database' only. Inserted first."
1795                     End With
1796                     sOpened = sTemplateIcon(CurrCategory.Templates("Routines"))
1797                     With .Add(, , CurrCategory.Key & " - Routines", "Routines", sOpened, sOpened)
1798                         .ExpandedImage = sOpened
1799                         .Expanded = True
                             .ToolTip = "'Entire Database' only. Inserted second."
1800                     End With
1801                     sOpened = sTemplateIcon(CurrCategory.Templates("Wrapper Class"))
1802                     With .Add(, , CurrCategory.Key & " - Wrapper Class", "Wrapper Class", sOpened, sOpened)
1803                         .ExpandedImage = sOpened
1804                         .Expanded = True
                             .ToolTip = "'Entire Database' only. Inserted third. Represents the database and contains SoftVars to reflect this."
1805                     End With
1806                     sOpened = sTemplateIcon(CurrCategory.Templates("Wrapper Class - Add collection"))
1807                     With .Add(CurrCategory.Key & " - Wrapper Class", tvwChild, CurrCategory.Key & " - Wrapper class - Add collection", "Wrapper class - Add collection", sOpened, sOpened)
1808                         .ExpandedImage = sOpened
1809                         .Expanded = True
                             .ToolTip = "'Entire Database' only. Inserted forth, once for each Table that does not have a parent."
1810                     End With
                         sOpened = sTemplateIcon(CurrCategory.Templates("Wrapper Class - Finalize"))
                         With .Add(CurrCategory.Key & " - Wrapper Class", tvwChild, CurrCategory.Key & " - Wrapper class - Finalize", "Wrapper class - Finalize", sOpened, sOpened)
                             .ExpandedImage = sOpened
                             .Expanded = True
                             .ToolTip = "Inserted after all insertions of 'Wrapper Class - Add collection"
                         End With
1811                     sOpened = sTemplateIcon(CurrCategory.Templates("Collection, No Parent"))
1812                     With .Add(, , CurrCategory.Key & " - Collection, No Parent", "Collection, No Parent", sOpened, sOpened)
1813                         .ExpandedImage = sOpened
1814                         .Expanded = True
                             .ToolTip = "Inserted once for each Table that does not have a parent but does have at least one child table."
1815                     End With
1816                     sOpened = sTemplateIcon(CurrCategory.Templates("Collection, No Parent, No Child"))
1817                     With .Add(, , CurrCategory.Key & " - Collection, No Parent, No Child", "Collection, No Parent, No Child", sOpened, sOpened)
1818                         .ExpandedImage = sOpened
1819                         .Expanded = True
                             .ToolTip = "Inserted once for each Table that has neither parent nor child Tables."
1820                     End With
1821                     sOpened = sTemplateIcon(CurrCategory.Templates("Collection, No Child"))
1822                     With .Add(, , CurrCategory.Key & " - Collection, No Child", "Collection, No Child", sOpened, sOpened)
1823                         .ExpandedImage = sOpened
1824                         .Expanded = True
                             .ToolTip = "Inserted once for each Table that has a parent Table, but no child tables."
1825                     End With
1826                     sOpened = sTemplateIcon(CurrCategory.Templates("Collection Member, Terminal"))
1827                     With .Add(CurrCategory.Key & " - Collection, No Child", tvwChild, CurrCategory.Key & " - Collection Member, Terminal", "Collection Member, Terminal", sOpened, sOpened)
1828                         .ExpandedImage = sOpened
1829                         .Expanded = True
                             .ToolTip = "Inserted once for each Table that does not have a child Table regardless if it has a parent Table or not."
1830                     End With
1831                     sOpened = sTemplateIcon(CurrCategory.Templates("Collection"))
1832                     With .Add(, , CurrCategory.Key & " - Collection", "Collection", sOpened, sOpened)
1833                         .ExpandedImage = sOpened
1834                         .Expanded = True
                             .ToolTip = "Inserted once for each Table that has both a parent and at least one child Table."
1835                     End With
1836                     sOpened = sTemplateIcon(CurrCategory.Templates("Collection Member"))
1837                     With .Add(CurrCategory.Key & " - Collection", tvwChild, CurrCategory.Key & " - Collection Member", "Collection Member", sOpened, sOpened)
1838                         .ExpandedImage = sOpened
1839                         .Expanded = True
                             .ToolTip = "Inserted once for each Table that has at least one child Table regardless if it has a parent Table or not."
1840                     End With
1841                     sOpened = sTemplateIcon(CurrCategory.Templates("Collection Member - New Subcollection"))
1842                     With .Add(CurrCategory.Key & " - Collection", tvwChild, CurrCategory.Key & " - Collection Member - New Subcollection", "Collection Member - New Subcollection", sOpened, sOpened)
1843                         .ExpandedImage = sOpened
1844                         .Expanded = True
                             .ToolTip = "Inserted once for each child Table after the parent's ""Collection Member"" is inserted."
1845                     End With
1846                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - 3D Link"))
1847                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - 3D Link", "Property - 3D Link", sOpened, sOpened)
1848                         .ExpandedImage = sOpened
1849                         .Expanded = True
                             .ToolTip = "Inserted after ""Collection..."" and ""Collection Member..."" templates for the field in a child table that links to it's logical parent."
1850                     End With
1851                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - BLOB"))
1852                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - BLOB", "Property - BLOB", sOpened, sOpened)
1853                         .ExpandedImage = sOpened
1854                         .Expanded = True
                             .ToolTip = "Inserted for each ""BLOB"" (binary/varbinary) field."
1855                     End With
1856                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - Boolean"))
1857                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Boolean", "Property - Boolean", sOpened, sOpened)
1858                         .ExpandedImage = sOpened
1859                         .Expanded = True
                             .ToolTip = "Inserted for each boolean (bit) field."
1860                     End With
1861                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - Byte"))
1862                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Byte", "Property - Byte", sOpened, sOpened)
1863                         .ExpandedImage = sOpened
1864                         .Expanded = True
                             .ToolTip = "Inserted for each Byte (tinyint) field."
1865                     End With
1866                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - Currency"))
1867                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Currency", "Property - Currency", sOpened, sOpened)
1868                         .ExpandedImage = sOpened
1869                         .Expanded = True
                             .ToolTip = "Inserted for each Currency field."
1870                     End With
1871                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - Date"))
1872                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Date", "Property - Date", sOpened, sOpened)
1873                         .ExpandedImage = sOpened
1874                         .Expanded = True
                             .ToolTip = "Inserted for each Date field."
1875                     End With
1876                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - Double"))
1877                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Double", "Property - Double", sOpened, sOpened)
1878                         .ExpandedImage = sOpened
1879                         .Expanded = True
                             .ToolTip = "Inserted for each Double field."
1880                     End With
1881                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - Integer"))
1882                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Integer", "Property - Integer", sOpened, sOpened)
1883                         .ExpandedImage = sOpened
1884                         .Expanded = True
                             .ToolTip = "Inserted for each Integer field."
1885                     End With
1886                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - Long"))
1887                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Long", "Property - Long", sOpened, sOpened)
1888                         .ExpandedImage = sOpened
1889                         .Expanded = True
                             .ToolTip = "Inserted for each Long field."
1890                     End With
1891                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - Memo"))
1892                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Memo", "Property - Memo", sOpened, sOpened)
1893                         .ExpandedImage = sOpened
1894                         .Expanded = True
                             .ToolTip = "Inserted for each Memo (varchar > 255) field."
1895                     End With
1896                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - OLE_COLOR"))
1897                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - OLE_COLOR", "Property - OLE_COLOR", sOpened, sOpened)
1898                         .ExpandedImage = sOpened
1899                         .Expanded = True
                             .ToolTip = "Inserted for each OLE_COLOR field."
1900                     End With
1901                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - Single"))
1902                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Single", "Property - Single", sOpened, sOpened)
1903                         .ExpandedImage = sOpened
1904                         .Expanded = True
                             .ToolTip = "Inserted for each Single precision floating point number field."
1905                     End With
1906                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - String"))
1907                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - String", "Property - String", sOpened, sOpened)
1908                         .ExpandedImage = sOpened
1909                         .Expanded = True
                             .ToolTip = "Inserted for each Text (char/varchar) field."
1910                     End With
1911                     sOpened = sTemplateIcon(CurrCategory.Templates("Property - Variant"))
1912                     With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Variant", "Property - Variant", sOpened, sOpened)
1913                         .ExpandedImage = sOpened
1914                         .Expanded = True
                             .ToolTip = "Inserted for any field not fitting the other ""Property..."" templates."
1915                     End With
1916                     sOpened = sTemplateIcon(CurrCategory.Templates("Finalize"))
1917                     With .Add(, , CurrCategory.Key & " - Finalize", "Finalize", sOpened, sOpened)
1918                         .ExpandedImage = sOpened
1919                         .Expanded = True
                             .ToolTip = "Inserted last."
1920                     End With
                         sOpened = sTemplateIcon(CurrCategory.Templates("Collection - Finalize"))
                         With .Add(CurrCategory.Key & " - Finalize", tvwChild, CurrCategory.Key & " - Collection - Finalize", "Collection - Finalize", sOpened, sOpened)
                             .ExpandedImage = sOpened
                             .Expanded = True
                             .ToolTip = "Inserted after all fields of a Table have been process and after Colleciton Member - Finalize."
                         End With
                         sOpened = sTemplateIcon(CurrCategory.Templates("Collection Member - Finalize"))
                         With .Add(CurrCategory.Key & " - Collection - Finalize", tvwChild, CurrCategory.Key & " - Collection Member - Finalize", "Collection Member - Finalize", sOpened, sOpened)
                             .ExpandedImage = sOpened
                             .Expanded = True
                             .ToolTip = "Inserted after all fields of a Table have been process, but before ""Collection - Finalize""."
                         End With

                    ' Fix up what's missing and what's added
1921                     For Each CurrTemplate In CurrCategory.Templates
1922                         With CurrTemplate
1923                             If Not .Deleted Then
1924                                 .OriginalShortName = .ShortTemplateName
1925                                 sOpened = UCase$(CurrTemplate.ShortTemplateName)
                                Select Case sOpened
                                    Case "SETTINGS"
1926                                     Case "ROUTINES"
1927                                     Case "WRAPPER CLASS"
1928                                     Case "WRAPPER CLASS - ADD COLLECTION"
1929                                     Case "FINALIZE"
1930                                     Case "COLLECTION"
1931                                     Case Else
1932                                         If InStr(sOpened, "COLLECTION ") Or InStr(sOpened, "COLLECTION,") Or InStr(UCase$(CurrTemplate.ShortTemplateName), "PROPERTY - ") Then
1933                                         Else
1934                                             sOpened = sTemplateIcon(CurrTemplate)
1935                                             With tvwX.Nodes.Add(, , CurrCategory.Key & gsCategoryTemplateDelimiter & CurrTemplate.ShortTemplateName, CurrTemplate.ShortTemplateName, sOpened, sOpened)
1936                                                 .Expanded = True
1937                                                 .ExpandedImage = sOpened
                                                     .ToolTip = "User template. Probably used by ""~##~Include"" from one of the standard data-pasting templates."
1938                                             End With
1939                                         End If
1940                                 End Select
1941                             End If
1942                         End With
1943                     Next CurrTemplate
1944                 End With
1945                 Set tvwY = Nothing
1946                 Set tvwX = Nothing
1947             End If
1948         End If
1949     Next CurrCategory
1950     lsbJumpTo.Visible = True

1951     UpdateFavorites

1952     UpdateHotKeys

1953 EH_frmMain_RefillList_Continue:
1954     bInHereAlready = False
1955     Exit Sub

1956 EH_frmMain_RefillList:
1957     MsgBox "Error occured in:" & gsEolTab & "Module: frmMain" & gsEolTab & "Procedure: RefillList" & gs2EOL & Err.Description

1958     Resume EH_frmMain_RefillList_Continue

1959     Resume
End Sub

Private Sub SetColors(ByVal BackColor As String, ByVal ForeColor As String)
1960     On Error Resume Next
1961     If Right$(BackColor, 1) = "&" Then BackColor = Left$(BackColor, Len(BackColor) - 1)
1962     If Right$(ForeColor, 1) = "&" Then ForeColor = Left$(ForeColor, Len(ForeColor) - 1)

1963     If Left$(BackColor, 2) <> "&H" Then BackColor = "&H" & BackColor
1964     If Left$(ForeColor, 2) <> "&H" Then ForeColor = "&H" & ForeColor

1965     lsbJumpTo.BackColor = BackColor
1966     txtCode(0).BackColor = BackColor
1967     txtCode(1).BackColor = BackColor
1968     txtCode(2).BackColor = BackColor
1969     txtCodeToFile.BackColor = BackColor
1970     lstSoftCommands.BackColor = BackColor
1971     lstSoftVariables.BackColor = BackColor

1972     lsbJumpTo.ForeColor = ForeColor
1973     txtCode(0).ForeColor = ForeColor
1974     txtCode(1).ForeColor = ForeColor
1975     txtCode(2).ForeColor = ForeColor
1976     txtCodeToFile.ForeColor = ForeColor
1977     lstSoftCommands.ForeColor = ForeColor
1978     lstSoftVariables.ForeColor = ForeColor

1979     If Not m_oDBClassGen Is Nothing Then
1980         With m_oDBClassGen
1981             .lvwFields.BackColor = BackColor
1982             .lvwFields.ForeColor = ForeColor
1983             .dvwTable.BackColor = BackColor
1984             .dvwTable.ForeColor = ForeColor
1985         End With
1986     End If
End Sub

Public Function SetInternalCurrentTemplate(ByVal sTemplateName As String) As Boolean
1987     On Error Resume Next
1988     Static sCategoryName As String
1989     Static sShortTemplateName As String

1990     Err.Clear
1991     SaveTemplate

1992     sCategoryName = sGetToken(sTemplateName, 1, gsCategoryTemplateDelimiter)
1993     sShortTemplateName = sAfter(sTemplateName, 1, gsCategoryTemplateDelimiter)

1994     If SliceAndDice(sCategoryName).Templates(sTemplateName) Is Nothing Then
1995         SetInternalCurrentTemplate = False
1996     Else
1997         Set InternalCurrentTemplate = SliceAndDice(sCategoryName).Templates(sTemplateName)
1998         SetInternalCurrentTemplate = True
1999     End If
End Function

Public Function sGetCurrentLineAtCharacter(ByVal sTextToSearch As String, ByVal lCharToStart As Long) As String
2000     Dim lCount As Long
2001     lCount = lTokenCount(Left$(sTextToSearch, lCharToStart), vbNewLine)
2002     If lCount > 0 Then
2003         sGetCurrentLineAtCharacter = sGetToken(sTextToSearch, lCount, vbNewLine)
2004     Else
2005         sGetCurrentLineAtCharacter = sTextToSearch
2006     End If
End Function

Public Sub ShowExternalsMenu()
2007     On Error Resume Next
2008     PopupMenu mnuExternalFunctions
End Sub

Public Sub ShowFavMenu()
2009     On Error Resume Next
2010     PopupMenu mnuFav
End Sub

Public Function ShutdownDLLs() As Boolean
2011     On Error Resume Next
2012     Dim CurrSet As Long

2013     For CurrSet = 1 To SadCommandSetCount
2014         Call SadCommands(CurrSet).Shutdown
2015         Set SadCommands(CurrSet) = Nothing
2016     Next CurrSet
2017     ReDim SadCommands(1 To 1)
2018     SadCommandSetCount = 0

2019     ShutdownDLLs = True
End Function

Public Property Get TemplateDatabaseName() As String
2020     TemplateDatabaseName = m_sTemplateDatabaseName
End Property

Public Sub SaveTemplate()
2021     On Error GoTo EH_frmMain_SaveTemplate
2022     Static bInHereAlready As Boolean
2023     If bInHereAlready Then Exit Sub
2024     bInHereAlready = True

2025     Static CurrBar As MSComctlLib.ListView

2026     If Not CurrentTemplate Is Nothing Then
2027         With CurrentTemplate
2028             If Not .Deleted Then
2029                 If .ParentKey = vbNullString Then
2030                     MsgBox "frmMain.SaveTemplate : Error found. ParentKey blank"
2031                     GoTo EH_frmMain_SaveTemplate_Continue
2032                 End If
2033                 txtName.Text = .ParentKey & gsCategoryTemplateDelimiter & txtShortName
2034                 .Key = txtName
2035                 .ShortTemplateName = txtShortName

2036                 .memoCodeAtTop = txtCode(0)
2037                 .memoCodeAtCursor = txtCode(1)
2038                 .memoCodeAtBottom = txtCode(2)

2039                 .FileName = txtFilename
2040                 .memoCodeToFile = txtCodeToFile

2041                 .Undeletable = chkUndeletable <> 0
2042                 .Locked = chkLocked <> 0
                '.IncludeInMenu = chkIncludeInMenu <> 0
2043                 .Favorite = chkFavorite <> 0
2044                 .Selected = chkSelected <> 0

                '               With SliceAndDice.SystemInfo("Hotkey Templates").Item(.Key)
                '                    If hkyInstantInsert.HotKeyAndModifier <> 0 Then
                '                       .Value = hkyInstantInsert.HotKey & gsC & hkyInstantInsert.HotKeyModifier
                '                    Else
                '                       .Value = "0,8"
                '                    End If
                '
                '
                '               End With

                '.Modified = True
2045             End If

2046             On Error Resume Next
2047             If .Modified Then
2048                 SliceAndDice.Save
2049                 Err.Clear
2050                 If (Not lsbJumpTo.Bars(.ParentKey) Is Nothing) And Len(.OriginalShortName) > 0 Then
2051                     Set CurrBar = lsbJumpTo.Bars(.ParentKey)
2052                     If Not CurrBar Is Nothing Then
2053                         If CurrBar.ListItems(.ParentKey & gsCategoryTemplateDelimiter & .OriginalShortName).Text <> .ShortTemplateName Then
2054                             CurrBar.ListItems(.ParentKey & gsCategoryTemplateDelimiter & .OriginalShortName).Text = .ShortTemplateName
2055                             Set CurrBar = Nothing

2056                             mnuFileRefresh_Click
2057                         End If
2058                     End If
2059                 End If
2060                 Err.Clear
2061             End If
2062         End With
2063     End If

2064 EH_frmMain_SaveTemplate_Continue:
2065     bInHereAlready = False
2066     Exit Sub

2067 EH_frmMain_SaveTemplate:
2068     MsgBox "Error occured in:" & gsEolTab & "Module: frmMain" & gsEolTab & "Procedure: SaveTemplate" & gs2EOL & Err.Description
2069     Resume EH_frmMain_SaveTemplate_Continue

2070     Resume
End Sub

Public Sub UpdateFavorites()
2071     On Error Resume Next
2072     Dim CurrFav As Long
2073     Dim CurrCategory As CCategory
2074     Dim CurrTemplate As CTemplate

2075     DoEvents: DoEvents: DoEvents

2076     If FavoriteCount > 0 Then                         ' Clear out previous entries
2077         For CurrFav = FavoriteCount To 1 Step -1
2078             Unload mnuFavorite(CurrFav)
2079         Next CurrFav
2080         mnuFavorite(0).Caption = "-Empty-"
2081         mnuFavorite(0).Enabled = False
2082         FavoriteCount = 0
2083     End If

2084     For Each CurrCategory In SliceAndDice.Categorys
2085         For Each CurrTemplate In CurrCategory.Templates
2086             If CurrTemplate.Favorite Then
2087                 If FavoriteCount > 0 Then
2088                     Load mnuFavorite(FavoriteCount)
2089                 End If
2090                 mnuFavorite(FavoriteCount).Caption = CurrTemplate.Key
2091                 mnuFavorite(FavoriteCount).Enabled = True
2092                 FavoriteCount = FavoriteCount + 1
2093             End If
2094         Next CurrTemplate
2095     Next CurrCategory
End Sub

Private Function sTemplateIcon(ByVal CurrTemplate As Object) As String
2096     If CurrTemplate Is Nothing Then
2097         sTemplateIcon = "!"
2098     ElseIf Len(CurrTemplate.memoCodeAtBottom & CurrTemplate.memoCodeAtCursor & CurrTemplate.memoCodeAtTop & CurrTemplate.memoCodeToFile) > 0 Then
2099         sTemplateIcon = gsCategory
2100     ElseIf CurrTemplate.Selected Then
2101         sTemplateIcon = "Check"
2102     ElseIf CurrTemplate.Undeletable Or CurrTemplate.Locked Then
2103         sTemplateIcon = "Key"
2104     Else
2105         sTemplateIcon = "Document"
2106     End If
End Function

Public Sub UpdateHotKeys()
'On Error GoTo EH_UpdateHotKeys
'    Dim asaTaken As CAssocArray
'    Dim CurrItem As CAssocItem
'
'    Exit Sub
'
'    If SliceAndDice.SystemInfo Is Nothing Then Exit Sub
'
'    Set asaTaken = New CAssocArray
'    asaTaken.Clear
'    asaTaken(vbKeyR & gsC & (MOD_CONTROL + MOD_SHIFT)) = "Sandy Repeat Insertion"
'    asaTaken(vbKeyS & gsC & (MOD_CONTROL + MOD_SHIFT)) = "Sandy Activate"
'
'    If mHotKeyOpenWindow Is Nothing Then
'       Set mHotKeyOpenWindow = New cRegHotKey
'    End If
'
'    With mHotKeyOpenWindow
'        .Attach Me.hwnd
'        .RegisterKey "Sandy Activate", vbKeyS, MOD_CONTROL + MOD_SHIFT
'        .RegisterKey "Sandy Repeat Insertion", vbKeyR, MOD_CONTROL + MOD_SHIFT
'    End With
'
'    With SliceAndDice.SystemInfo("Hotkey Templates")
'         For Each CurrItem In .mCol
'             If CurrItem.Value <> "0,8" Then
'                If Len(asaTaken(CurrItem.Value)) = 0 Then
'                   asaTaken(CurrItem.Value) = CurrItem.Key
'                   mHotKeyOpenWindow.RegisterKey  gsTemplate & gsS & CurrItem.Key, Val(sGetToken(CurrItem.Value, 1, gsC)), Val(sGetToken(CurrItem.Value, 2, gsC))
'                End If
'             End If
'         Next CurrItem
'    End With
'
'EH_UpdateHotKeys_Continue:
'    Set CurrItem = Nothing
'    Set asaTaken = Nothing
'    Exit Sub
'
'EH_UpdateHotKeys:
'    LogError "frmMain", "UpdateHotKeys", Err.Number, Err.Description, Erl
'    Resume EH_UpdateHotKeys_Continue
'
'    Resume
End Sub

Public Sub chkAutoRecalc_Click()
2107     SaveSetting App.ProductName, "Last", "Auto Recalc", chkAutoRecalc.Value
End Sub

Public Sub chkFavorite_Click()
2108     If mbFillingAddInScreen Then Exit Sub
2109     CurrentTemplate.Favorite = (chkFavorite.Value <> 0)
2110     mnuIsFavorite.Checked = CurrentTemplate.Favorite
2111     UpdateFavorites
End Sub

Public Sub chkUndeletable_Click()
2112     Dim sPasswordCheck As String
2113     Static bInHereAlready As Boolean

2114     If mbFillingAddInScreen Then Exit Sub
2115     If bInHereAlready Then Exit Sub
2116     If Not mnuPasswordProtection.Checked Then
2117         lsbJumpTo.BarItemIcon = IIf(chkLocked.Value = 0 And chkUndeletable.Value = 0, "Document" & IIf(lsbJumpTo.BarItemIcon = "DocumentAlternate", "Alternate", vbNullString), "Key")
2118         Exit Sub
2119     End If

2120     bInHereAlready = True

2121     If chkUndeletable.Value = 0 Then
2122         If Len(CurrentTemplate.memoAttributes) Then
2123             m_asaAttributes.All = CurrentTemplate.memoAttributes
2124             If Len(m_asaAttributes("Undeletable Password")) Then
2125                 sPasswordCheck = InputBox("Enter password to unlock.", "ENTER PASSWORD")
2126                 If StrComp(sPasswordCheck, m_asaAttributes("Undeletable Password")) <> 0 Then
2127                     Beep
2128                     chkUndeletable.Value = 1
2129                     lsbJumpTo.BarItemIcon = "Key"
2130                 Else
2131                     chkUndeletable.Value = 0
2132                     lsbJumpTo.BarItemIcon = IIf(chkLocked.Value = 0 And chkUndeletable.Value = 0, "Document" & IIf(lsbJumpTo.BarItemIcon = "DocumentAlternate", "Alternate", vbNullString), "Key")
2133                 End If
2134             End If
2135         Else
2136         End If
2137     Else
2138         m_asaAttributes.All = CurrentTemplate.memoAttributes
2139         m_asaAttributes("Undeletable Password") = InputBox("Enter a password for unlocking later.", "ENTER NEW PASSWORD", m_asaAttributes("Undeletable Password"))
2140         CurrentTemplate.memoAttributes = m_asaAttributes.All
2141         chkUndeletable.Value = 1
2142         lsbJumpTo.BarItemIcon = "Key"
2143     End If
2144     bInHereAlready = False
End Sub

Public Sub chkUndeletable_Validate(Cancel As Boolean)
2145     Dim sPasswordCheck As String
End Sub


Public Sub cmdRecalc_Click()
2146     On Error GoTo EH_cmdRecalc_Click

2147     Dim sCodeToCheck(0 To 4) As String
2148     Dim CurCodeWindow As Long
2149     Dim lTokens As Long
2150     Dim CurToken As Long
2151     Dim CurListItem As Long
2152     Dim sCurToken As String

2153     Screen.MousePointer = vbHourglass

2154     lstSoftVariables.Clear
2155     lstSoftCommands.Clear
    'MsgBox "Recalc to occur here."

2156     sCodeToCheck(0) = txtCode(1)
2157     sCodeToCheck(1) = txtCode(0)
2158     sCodeToCheck(2) = txtCode(2)

2159     With lstSoftVariables
        Select Case UCase$(txtShortName)
            Case "COLLECTION", "COLLECTION, NO CHILD", "COLLECTION, NO PARENT", "COLLECTION, NO PARENT, NO CHILD"
2160                 If InStr(UCase$(txtShortName), "NO PARENT") = 0 Then
2161                     .AddItem "* Parent AutoNumber Field Name"
2162                     .AddItem "* Parent AutoNumber Property Name"
2163                 End If
                '.AddItem "* Collection Member Subcollection Property Name = Child Table Name"
2164                 .AddItem "* Property Name"
2165                 .AddItem "* Singular Property Name = Child Table Name"
2166                 .AddItem "* Child Table Name"
                '.AddItem "* Primary AutoNumber Field for Collection Member = AutoNumber Field"
2167                 .AddItem "* AutoNumber Field Name"
2168                 .AddItem "* AutoNumber Property Name"
                '.AddItem "* Table that stores this collection = Table Name"
2169                 .AddItem "* Object Name = Table Name"
2170                 .AddItem "* Table Name"
                '.AddItem "* Object Name of Collection Member = Object Name"
2171                 .AddItem "* Spaced Table Name"
2172                 .AddItem "* Spaced Object Name"
                '.AddItem "* Label Name of Collection Member = Label Name"
2173                 .AddItem "* Label Name"
                '.AddItem "* Field to use as Key = Key Field"
2174                 .AddItem "* Key Field Name"
2175                 .AddItem "* Key Property Name"

2176             Case "COLLECTION MEMBER", "COLLECTION MEMBER - TERMINAL"
                '.AddItem "* Object Name of Collection Member = Object Name"
2177                 .AddItem "* Object Name = Table Name"
2178                 .AddItem "* Table Name"
                '.AddItem "* Label Name of Collection Member = Label Name"
2179                 .AddItem "* Label Name = Spaced Table Name"
2180                 .AddItem "* Spaced Table Name"
                '.AddItem "* Property name of Class to collect"
2181                 .AddItem "* Class to collect = Property Name"
2182                 .AddItem "* Property Name"

2183             Case "COLLECTION MEMBER - NEW SUBCOLLECTION"
2184                 .AddItem "* Property Name"
2185                 .AddItem "* Singular Property Name = Child Table Name"
2186                 .AddItem "* Child Table Name"
2187                 .AddItem "* Table Name"
2188                 .AddItem "* Spaced Table Name"

2189             Case "PROPERTY - BLOB", "PROPERTY - BOOLEAN", "PROPERTY - BYTE", "PROPERTY - CURRENCY", "PROPERTY - DATE", "PROPERTY - DOUBLE", "PROPERTY - INTEGER", "PROPERTY - LONG", "PROPERTY - OLE_COLOR", "PROPERTY - SINGLE", "PROPERTY - STRING", "PROPERTY - VARIANT", "PROPERTY - 3D LINK"
                '.AddItem "* Field Name of Property"
2190                 .AddItem "* Property Name"
                '.AddItem "* Pure Field Name"
2191                 .AddItem "* Field Name"
2192                 .AddItem "* Table Name"
2193                 .AddItem "* Spaced Field Name"
2194                 .AddItem "* Spaced Table Name"
                     .AddItem "* Property Type"
                     .AddItem "* Property Size"
                     .AddItem "* Property Length"
                     .AddItem "* Field Type"
                     .AddItem "* Field Length"

2195             Case "WRAPPER CLASS", "ROUTINES"
2196                 .AddItem "* DSN = Database Name"
2197                 .AddItem "* Database Name"
2198                 .AddItem "* Database Path"
2199                 .AddItem "* Spaced Database Name"

2200             Case "WRAPPER CLASS - ADD COLLECTION"
                '.AddItem "* Property name of Class to collect"
2201                 .AddItem "* Property Name"
                '.AddItem "* Plural Table Name"
2202                 .AddItem "* Table Name"
2203                 .AddItem "* Spaced Table Name"
2204         End Select
2205     End With

2206     For CurCodeWindow = 0 To 2
        ' Scan for Soft Variables
2207         lTokens = lTokenCount(sCodeToCheck(CurCodeWindow), gsSoftVarDelimiter)
2208         If (lTokens Mod 2) = 0 And lTokens > 0 Then   ' Even
            ' Token theory clearly states that
            '   If you're using one delimiter for delimiting both
            '      the beginning and ending of a token, then there must be
            '      an ODD number of tokens or the string isn't valid.
            Select Case CurCodeWindow
                Case 0: MsgBox "You are missing at least one Soft Variable Delimiter (" & gsQ & gsSoftVarDelimiter & gsQ & ") in the 'At cursor' code area of the current " & gsTemplate & ". Discontinuing analysis."
2209                 Case 1: MsgBox "You are missing at least one Soft Variable Delimiter (" & gsQ & gsSoftVarDelimiter & gsQ & ") in the '(Declarations)' code area of the current " & gsTemplate & ". Discontinuing analysis."
2210                 Case 2: MsgBox "You are missing at least one Soft Variable Delimiter (" & gsQ & gsSoftVarDelimiter & gsQ & ") in the 'End of Module' code area of the current " & gsTemplate & ". Discontinuing analysis."
2211                 Case 3: MsgBox "You are missing at least one Soft Variable Delimiter (" & gsQ & gsSoftVarDelimiter & gsQ & ") in the 'In a file' code area of the current " & gsTemplate & ". Discontinuing analysis."
2212             End Select
2213             Screen.MousePointer = vbDefault
2214             Exit Sub
2215         ElseIf lTokens > 0 Then                       ' Odd
            ' Okay, keep going
2216             For CurToken = 2 To lTokens Step 2
2217                 sCurToken = sGetToken(sCodeToCheck(CurCodeWindow), CurToken, gsSoftVarDelimiter)
2218                 If lstSoftVariables.ListCount > 0 Then
2219                     For CurListItem = 0 To lstSoftVariables.ListCount - 1
                        Select Case StrComp(UCase$(lstSoftVariables.List(CurListItem)), UCase$(sCurToken))
                            Case 0
2220                                 CurListItem = lstSoftVariables.ListCount
2221                                 Exit For
2222                             Case Is > 0
2223                                 lstSoftVariables.AddItem sCurToken, CurListItem
2224                                 CurListItem = CurListItem + 1
2225                                 Exit For
2226                         End Select
2227                     Next CurListItem
2228                     If StrComp(UCase$(lstSoftVariables.List(CurListItem - 1)), UCase$(sCurToken)) < 0 Then
2229                         lstSoftVariables.AddItem sCurToken
2230                     End If
2231                 Else
2232                     lstSoftVariables.AddItem sCurToken
2233                 End If
2234             Next CurToken
2235         End If

        ' Repeat for Soft Commands
2236         lTokens = lTokenCount(sCodeToCheck(CurCodeWindow), vbNewLine)

        ' Okay, keep going
2237         sCurToken = vbNullString
2238         For CurToken = 1 To lTokens Step 1
2239             sCurToken = sGetToken(sGetToken(sCodeToCheck(CurCodeWindow), CurToken, vbNewLine), 2, gsSoftCmdDelimiter)
2240             If sCurToken <> gsA And Len(sCurToken) > 0 Then
2241                 lstSoftCommands.AddItem sCurToken
2242             End If
2243         Next CurToken
2244         If StrComp(UCase$(lstSoftCommands.List(CurListItem - 1)), UCase$(sCurToken)) < 0 Then
2245             lstSoftCommands.AddItem sCurToken
2246         End If
        'End If
2247     Next CurCodeWindow

2248 EH_cmdRecalc_Click_Continue:
2249     Screen.MousePointer = vbDefault
2250     Exit Sub

2251 EH_cmdRecalc_Click:
2252     Resume EH_cmdRecalc_Click_Continue

2253     Resume
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If mbScramFormKey Then KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If mbScramFormKey Then KeyCode = 0: Shift = 0
End Sub

Private Sub Form_LostFocus()
2254     On Error Resume Next
2255     lsbJumpTo.HideCategories
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
2256     If UnloadMode = vbFormControlMenu Then
2257         Cancel = True
2258         mnuFileExit_Click
2259     Else
2260         Form_Unload Cancel
2261     End If

    'If OkayToUnload Then
    '   Form_Unload Cancel
    'Else
    '   Cancel = True
    'End If
End Sub

Private Sub lsbJumpTo_AfterBarClick()
2262     On Error Resume Next
2263     JumpTo lsbJumpTo.BarKey & gsCategoryTemplateDelimiter & lsbJumpTo.BarItemName
End Sub

Public Sub lsbJumpTo_BarItemClick(ByVal BarName As String, ByVal BarKey As String, ByVal BarItemName As String, ByVal BarItemKey As String)
2264     On Error Resume Next
2265     Dim TemplateFound As CTemplate
2266     If lsbJumpTo.BarType = "List" Then
2267         JumpTo BarItemKey
2268     Else
2269         Set TemplateFound = SliceAndDice.Categorys.ItemByLongTemplateName(BarItemKey)
2270         If TemplateFound Is Nothing Then
2271             Beep
2272             If MsgBox("That " & gsTemplate & " does not exist (yet)." & gsEolTab & "Create " & gsTemplate & " now ?", vbYesNo, "NO " & gsTemplate & ": " & BarItemKey) = vbYes Then
2273                 QueueAction "NewTemplate", BarItemKey
2274                 OkayToDoAction = True
2275             Else
2276                 If Val(CurrentHistoryEntry) > 0 Then
2277                     JumpTo m_asaHistory(CurrentHistoryEntry), False, True
2278                 ElseIf SliceAndDice.Categorys(sGetToken(BarItemKey, 1, gsCategoryTemplateDelimiter)).Templates.Count > 1 Then
2279                     JumpTo SliceAndDice.Categorys(sGetToken(BarItemKey, 1, gsCategoryTemplateDelimiter)).Templates(1)
2280                 End If
2281             End If
2282         Else
2283             JumpTo BarItemKey
2284         End If
2285         Set TemplateFound = Nothing
2286     End If
End Sub

Public Sub lsbJumpTo_BarItemDblClick(ByVal BarName As String, ByVal BarKey As String, ByVal BarItemName As String, ByVal BarItemKey As String)
2287     If Len(BarItemKey) = 0 Then Exit Sub

2288     mnuInsertTemplate_Click
End Sub

Private Sub lsbJumpTo_KeyDown(KeyCode As Integer, Shift As Integer)
       If Not mbScramFormKey Then
          Form_KeyDown KeyCode, Shift
          If mbScramFormKey Then KeyCode = 0: Shift = 0
       End If

'2950     On Error GoTo EH_frmMain_lsbJumpTo_KeyDown
'2951     Dim sSelectedText   As String
'2952     Dim sOrigText       As String
'2953     Dim CurrSet         As Long
'
'
'2954     If (Shift And vbShiftMask) > 0 Then               ' Shift Key         *******************
'                Select Case KeyCode
'
'                    Case vbKeyInsert: KeyCode = 0: Shift = 0  ' Paste
'2955                     If TypeOf ActiveControl Is TextBox Then
'2956                        ActiveControl.SelText = Clipboard.GetText
'2957                     End If
'2958                Case vbKeyDelete: KeyCode = 0: Shift = 0  ' Cut
'2959                     If TypeOf ActiveControl Is TextBox Then
'2960                        If StringToClipboard(ActiveControl.SelText) Then
'2961                           ActiveControl.SelText = vbNullString
'2962                        End If
'2963                     End If
'2964            End Select
'
'
'
'2965     ElseIf (Shift And vbCtrlMask) > 0 Then            ' CTRL Key      ********************
'                Select Case KeyCode
'                    Case vbKeyInsert: KeyCode = 0: Shift = 0  ' Copy
'2966                     If TypeOf ActiveControl Is TextBox Then
'2967                        StringToClipboard ActiveControl.SelText
'2968                     End If
'
'2969                Case vbKeyDelete: KeyCode = 0: Shift = 0  ' Cut
'2970                     If TypeOf ActiveControl Is TextBox Then
'2971                        If StringToClipboard(ActiveControl.SelText) Then
'2972                           ActiveControl.SelText = vbNullString
'2973                        End If
'2974                     End If
'
'2975                Case vbKeyTab
'2977                     KeyCode = 0: Shift = 0
'  If Parent.HostedByVB Then  ' Shell App override
'2976                     On Error Resume Next
'2978                     Parent.vbInst.ActiveWindow.SetFocus
'  End If
'
'2979                Case vbKeyF: KeyCode = 0: Shift = 0: FindInCurrent
'2980                Case vbKeyH: KeyCode = 0: Shift = 0: FindInCurrent False, True
'
'2982                Case vbKeyI: KeyCode = 0: Shift = 0: mnuInsertTemplate_Click
'2983                Case vbKeyL: KeyCode = 0: Shift = 0: mnuFileRefresh_Click
'2984                Case vbKeyM: KeyCode = 0: Shift = 0: mnuFileImport_Click
'2985                Case vbKeyN: KeyCode = 0: Shift = 0: mnuFileNew_Click
'
'                    Case vbKey1: KeyCode = 0: Shift = 0: lsbJumpTo.SetFocus
'                    Case vbKey2: KeyCode = 0: Shift = 0: txtShortName.SetFocus
'                    Case vbKey3: KeyCode = 0: Shift = 0: tabCode.Tabs(1).Selected = True: tabCode_MouseUp 0, 0, 0, 0
'                    Case vbKey4: KeyCode = 0: Shift = 0: tabCode.Tabs(2).Selected = True: tabCode_MouseUp 0, 0, 0, 0
'                    Case vbKey5: KeyCode = 0: Shift = 0: tabCode.Tabs(3).Selected = True: tabCode_MouseUp 0, 0, 0, 0
'                    Case vbKey6: KeyCode = 0: Shift = 0: tabCode.Tabs(4).Selected = True: tabCode_MouseUp 0, 0, 0, 0
'                    Case vbKey7: KeyCode = 0: Shift = 0: tabCode.Tabs(5).Selected = True: tabCode_MouseUp 0, 0, 0, 0
'2986           End Select
'
'
'
'2987     ElseIf (Shift And vbAltMask) > 0 Then                      ' Alt Key       *********************
'                Select Case KeyCode
'                    Case vbKeyLeft: KeyCode = 0: Shift = 0: mnuBack_Click
'2988                     Case vbKeyRight: KeyCode = 0: Shift = 0: mnuForward_Click
'
'2989                     Case vbKeyX: KeyCode = 0: Shift = 0: mnuFileExit_Click
'
'2990                     Case vbKeyTab
'2991                         On Error Resume Next
'2992                         KeyCode = 0: Shift = 0
'  If Parent.HostedByVB Then  ' Shell App override
'2993                         Parent.vbInst.ActiveWindow.SetFocus
'  End If
'2994            End Select
'2995     Else                                              'If (Shift And vbShiftMask) = 0 Then     ' No special modifying keys
'        Select Case KeyCode
'            Case vbKeyTab
'2996                 If TypeOf ActiveControl Is TextBox Then
'2997                     If InStr(ActiveControl.Tag, "Code Area ") Then
'2998                         ActiveControl.SelText = vbTab
'2999                         KeyCode = 0
'3000                         Shift = 0
'3001                     End If
'3002                 End If
'
'3003             Case vbKeyF1
'3004                 If TypeOf ActiveControl Is TextBox Then
'3005                     KeyCode = 0: Shift = 0
'
'3006                     sOrigText = Trim$(ActiveControl.SelText)
'3007                     If Len(sOrigText) = 0 Then
'3008                         sOrigText = sGetCurrentLineAtCharacter(ActiveControl.Text, ActiveControl.SelStart)
'3009                     End If
'3010                     sSelectedText = sOrigText
'3011                     If Len(sSelectedText) > 0 Then
'3012                         If InStr(sSelectedText, gsSoftCmdDelimiter) Then
'3013                             If InStr(sSelectedText, gsSoftCmdDelimiter & "_") = 0 Then
'3014                                 sSelectedText = UCase$(sGetToken(sGetToken(sSelectedText, 2, gsSoftCmdDelimiter), 1)) & "*C"
'3015                             Else
'3016                                 sSelectedText = UCase$(sGetToken(sGetToken(sSelectedText, 3, "_"), 1)) & "*C"
'3017                             End If
'3018                         End If
'3019                         If InStr(sSelectedText, gsSoftVarDelimiter) Then
'3020                             sSelectedText = sGetToken(sGetToken(sSelectedText, 2, gsSoftVarDelimiter), 1, gsInlineCmdDelimiter) & "*I"
'                             ElseIf InStr(sSelectedText, "~##~") Then
'                                 sSelectedText = sGetToken(sSelectedText, 1, " ") & "*I"
'3021                         End If
'                             If Not Complete Is Nothing Then
'                                Complete.ShowHelpScreen sSelectedText
'3022                         ElseIf SadCommandSetCount > 0 Then
'3023                             For CurrSet = 1 To SadCommandSetCount
'3024                                 If Not SadCommands(CurrSet).CommandSet.Item(sSelectedText) Is Nothing Then
'3025                                     SadCommands(CurrSet).CommandSet.ShowHelpScreen sSelectedText
'3026                                     GoTo EH_frmMain_lsbJumpTo_KeyDown_Continue
'3027                                 End If
'3028                             Next CurrSet
'3029                             MsgBox "SoftCommand '" & Trim$(sOrigText) & "' not found."
'3030                         Else
'3031                             MsgBox "No command set DLLs loaded. No help available."
'3032                         End If
'3033                     Else
'3034                         MsgBox "Unable to determine which command to display help for."
'3035                     End If
'3036                 End If
'
'3037             Case vbKeyF3: KeyCode = 0: Shift = 0: FindInCurrent True
'3038         End Select
'3039     End If
'
'3040 EH_frmMain_lsbJumpTo_KeyDown_Continue:
'3041     Exit Sub
'
'3042 EH_frmMain_lsbJumpTo_KeyDown:
'3043     LogError "frmMain", "lsbJumpTo_KeyDown", Err.Number, Err.Description, Erl
'3044     Resume EH_frmMain_lsbJumpTo_KeyDown_Continue
'
'3045     Resume
End Sub

Private Sub lsbJumpTo_KeyPress(KeyAscii As Integer)
       If Not mbScramFormKey Then KeyAscii = 0
End Sub


Private Sub lsbJumpTo_KeyUp(KeyCode As Integer, Shift As Integer)
       If Not mbScramFormKey Then
          Form_KeyDown KeyCode, Shift
          If mbScramFormKey Then KeyCode = 0: Shift = 0
       End If
End Sub


Private Sub lsbJumpTo_LostFocus()
2289     On Error Resume Next
2290     lsbJumpTo.HideCategories
End Sub

Public Sub lsbJumpTo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
2291     If Button = vbRightButton And Shift = 0 Then      ' Right click, pop-up menu
2292         PopupMenu mnuTemplate
2293     End If
End Sub

Public Sub lsbJumpTo_MouseDownOnCategory(Button As Integer, Shift As Integer, X As Single, Y As Single)
2294     If Button = vbRightButton And Shift = 0 Then      ' Right click, pop-up menu
2295         PopupMenu mnuCategories
2296     End If
End Sub

Public Sub lstSoftCommands_DblClick()
2297     On Error Resume Next
2298     lstSoftVariables.ListIndex = -1
2299     txtCode(0).SelStart = 0: txtCode(0).SelLength = 0
2300     txtCode(1).SelStart = 0: txtCode(1).SelLength = 0
2301     txtCode(2).SelStart = 0: txtCode(2).SelLength = 0
2302     FindInCurrent False, False, True
End Sub

Public Sub lstSoftVariables_DblClick()
2303     On Error Resume Next
2304     lstSoftCommands.ListIndex = -1
2305     txtCode(0).SelStart = 0: txtCode(0).SelLength = 0
2306     txtCode(1).SelStart = 0: txtCode(1).SelLength = 0
2307     txtCode(2).SelStart = 0: txtCode(2).SelLength = 0
2308     FindInCurrent False, False, True
End Sub

Public Sub mnuBack_Click()
2309     On Error Resume Next
2310     If Val(CurrentHistoryEntry) < 2 Then
2311         Beep
2312         mnuBack.Enabled = False
2313         Exit Sub
2314     End If

2315     CurrentHistoryEntry = Val(CurrentHistoryEntry) - 1
2316     JumpTo m_asaHistory(CurrentHistoryEntry), False, True
2317     mnuForward.Enabled = True
2318     mnuBack.Enabled = Val(CurrentHistoryEntry) > 1
End Sub

Public Sub mnuCategoriesDeleteCurrent_Click()
2319     On Error GoTo EH_mnuCategoriesDeleteCurrent_Click
2320     Static bInHereAlready As Boolean

2321     If bInHereAlready Then Exit Sub
2322     bInHereAlready = True

2323     Dim sCurrentCategory As String
2324     SaveTemplate

2325     sCurrentCategory = lsbJumpTo.BarKey

2326    'If UCase$(sCurrentCategory) = "CHANGE FROM" Then
2327    '    MsgBox "The 'Change From' " & gsCategory & " is not removable.", vbExclamation
2328    '    GoTo EH_mnuCategoriesDeleteCurrent_Click_Continue
2329    '    Exit Sub
        'ElseIf SliceAndDice.Categorys(sCurrentCategory).CategoryType <> 0 Then
2330     If SliceAndDice.Categorys(sCurrentCategory).CategoryType <> 0 Then
2331         If Not bUserSure("This " & gsCategory & " is used by the code generators. Deleting it is unadvisable." & gs2EOLTab & "Are you sure you want to permanently erase this " & gsCategory & " ?") Then
2332             GoTo EH_mnuCategoriesDeleteCurrent_Click_Continue
2333             Exit Sub
2334         End If
2335     End If

2336     With SliceAndDice.Categorys(sCurrentCategory)
2337         If .Templates.Count > 0 Then
2338             If Not bUserSure("There " & IIf(.Templates.Count = 1, "is", "are") & gsS & .Templates.Count & gsS & gsTemplate & IIf(.Templates.Count = 1, vbNullString, "s") & " still in the '" & sCurrentCategory & "' " & gsCategory & ". Continuing will delete all templates in that " & gsCategory & gsP & gs2EOLTab & "Are you absolutely sure this is what you want to do ?") Then
2339                 GoTo EH_mnuCategoriesDeleteCurrent_Click_Continue
2340                 Exit Sub
2341             End If
2342         End If
2343         .Deleted = True
2344     End With

2345     SliceAndDice.Save
2346     RefillList
2347     If Not SliceAndDice(1) Is Nothing Then
2348         If Not SliceAndDice(1).Templates(1) Is Nothing Then
2349             JumpTo SliceAndDice(1).Templates(1).Key
2350             lsbJumpTo.BarAndItem SliceAndDice(1).Key, SliceAndDice(1).Templates(1).ShortTemplateName
2351         ElseIf Not SliceAndDice(2) Is Nothing Then
2352             If Not SliceAndDice(2).Templates(1) Is Nothing Then
2353                 JumpTo SliceAndDice(2).Templates(1).Key
2354                 lsbJumpTo.BarAndItem SliceAndDice(2).Key, SliceAndDice(2).Templates(1).ShortTemplateName
2355             End If
2356         End If
2357     End If

2358 EH_mnuCategoriesDeleteCurrent_Click_Continue:
2359     bInHereAlready = False
2360     Exit Sub

2361 EH_mnuCategoriesDeleteCurrent_Click:
2362     MsgBox "Error occured in:" & gsEolTab & "Module: frmMain" & gsEolTab & "Procedure: DeleteCategory" & gs2EOL & Err.Description

2363     Resume EH_mnuCategoriesDeleteCurrent_Click_Continue

2364     Resume
End Sub

Public Sub mnuCategoriesNewMethod_Click(Index As Integer)
2365     Dim sNewCategoryName As String
2366     Dim sCategoryToDuplicate As String

    Select Case Index
        Case 0                                        ' New, Blank Category
2367             sNewCategoryName = InputBox("What should the name of the new, blank " & gsCategory & " be ?", "NEW " & gsCategory, vbNullString)
2368             If Len(sNewCategoryName) = 0 Then Exit Sub
2369             If SliceAndDice(sNewCategoryName) Is Nothing Then
2370                 SliceAndDice.Categorys.Add sNewCategoryName
2371                 SliceAndDice.Save
2372                 RefillList
2373             Else
2374                 MsgBox "There is already a " & gsCategory & " by that name. Aborting.", vbInformation
2375             End If

2376         Case 1                                        ' New, Duplicate Current Category
2377             sCategoryToDuplicate = SliceAndDice.Categorys.Choose
2378             If Len(sCategoryToDuplicate) = 0 Then Exit Sub
2379             sNewCategoryName = InputBox("What should the name of the new, duplicated " & gsCategory & " be ?", "DUPLICATE " & gsCategory, "Copy of " & sCategoryToDuplicate)
2380             If Len(sNewCategoryName) = 0 Then Exit Sub
2381             If SliceAndDice(sNewCategoryName) Is Nothing Then
2382                 SliceAndDice.Categorys.Add sNewCategoryName, , sCategoryToDuplicate
2383                 SliceAndDice.Save
2384                 RefillList                            'RefreshDatabaseConnection 'RefillList
2385             Else
2386                 MsgBox "There is already a " & gsCategory & " by that name. Aborting.", vbInformation
2387             End If

2388         Case 2                                        ' New, Duplicate Current Category, But don't copy any information from the templates. Only copy names.
2389             sCategoryToDuplicate = SliceAndDice.Categorys.Choose
2390             If Len(sCategoryToDuplicate) = 0 Then Exit Sub
2391             sNewCategoryName = InputBox("What should the name of the new, duplicated " & gsCategory & " (names, no code) be ?", "DUPE NAMES ONLY", "Copy of " & sCategoryToDuplicate)
2392             If Len(sNewCategoryName) = 0 Then Exit Sub
2393             If SliceAndDice(sNewCategoryName) Is Nothing Then
2394                 SliceAndDice.Categorys.Add sNewCategoryName, , sCategoryToDuplicate, False
2395                 SliceAndDice.Save
2396                 RefillList                            'RefreshDatabaseConnection 'RefillList
2397             Else
2398                 MsgBox "There is already a " & gsCategory & " by that name. Aborting.", vbInformation
2399             End If
2400     End Select
End Sub


Private Sub mnuChangeBackgroundColors_Click()
2401     Dim ColorSelected As String
2402     ColorSelected = sChooseColor(lsbJumpTo.BackColor)
2403     If Len(ColorSelected) = 0 Then Exit Sub

2404     SaveSetting App.ProductName, "Last", "Background Color", ColorSelected
2405     SetColors ColorSelected, GetSetting$(App.ProductName, "Last", "Foreground Color", "&H80000008&")
End Sub

Private Sub mnuChangeForegroundColor_Click()
2406     Dim ColorSelected As String
2407     ColorSelected = sChooseColor(lsbJumpTo.ForeColor)
2408     If Len(ColorSelected) = 0 Then Exit Sub
2409     SaveSetting App.ProductName, "Last", "Foreground Color", ColorSelected

2410     SetColors GetSetting$(App.ProductName, "Last", "Background Color", "&H80000018&"), ColorSelected
End Sub

Public Sub mnuEditCopy_Click()
2411     If Not chkLocked Then
        Select Case tabCode.SelectedItem.Index
            Case 1: StringToClipboard txtCode(0).SelText
2412             Case 2: StringToClipboard txtCode(1).SelText
2413             Case 3: StringToClipboard txtCode(2).SelText
2414             Case 4: StringToClipboard txtCode(3).SelText
2415         End Select
2416     End If
End Sub

Public Sub mnuEditCut_Click()
2417     If Not chkLocked Then
        Select Case tabCode.SelectedItem.Index
            Case 1: If StringToClipboard(txtCode(0).SelText) Then txtCode(0).SelText = vbNullString
2418             Case 2: If StringToClipboard(txtCode(1).SelText) Then txtCode(1).SelText = vbNullString
2419             Case 3: If StringToClipboard(txtCode(2).SelText) Then txtCode(2).SelText = vbNullString
2420             Case 4: If StringToClipboard(txtCode(3).SelText) Then txtCode(3).SelText = vbNullString
2421         End Select
2422     End If
End Sub


Public Sub mnuEditFind_Click()
2423     FindInCurrent
End Sub


Public Sub mnuEditPaste_Click()
2424     If Not chkLocked Then
        Select Case tabCode.SelectedItem.Index
            Case 1: txtCode(0).SelText = Clipboard.GetText
2425             Case 2: txtCode(1).SelText = Clipboard.GetText
2426             Case 3: txtCode(2).SelText = Clipboard.GetText
2427             Case 4: txtCode(3).SelText = Clipboard.GetText
2428         End Select
2429     Else
2430         MsgBox "Code areas of current " & gsTemplate & " locked. Unlock " & gsTemplate & " (under the options tab) before attempting this again.", vbInformation
2431     End If
End Sub


Public Sub mnuEditReplace_Click()
2432     FindInCurrent False, True
End Sub


Private Sub mnuExternals_Click(Index As Integer)
2433     On Error Resume Next
2434     SadCommands(Val(sGetToken(mnuExternals(Index).Tag, 1, "|"))).ExecuteExternal mnuExternals(Index).Caption, sAfter(mnuExternals(Index).Tag, 1, "|")
End Sub

Public Sub mnuFavorite_Click(Index As Integer)
2435     If FavoriteCalledFromIDE Then
2436         FavoriteCalledFromIDE = False
2437         DoInsertion Nothing, mnuFavorite(Index).Caption
2438     Else
2439         JumpTo mnuFavorite(Index).Caption, , True
2440         lsbJumpTo.HideCategories
2441     End If
End Sub

Private Sub mnuFileApplyDeltaPatch_Click()
2442     Dim sFilename As String
2443     sFilename = sChooseFile(, , "Sandy Delta Patch (*.sad)|*.sad|All Files (*.*)|*.*")
2444     If Len(sFilename) Then
2445         SliceAndDice.ApplyPatch sFilename
2446     End If
End Sub

Private Sub mnuFileGenerateDeltaPatch_Click(Index As Integer)
2447     Dim sDate As String
2448     Dim PatchFilename As String

2449     sDate = SliceAndDice.sChoosePatch(Index)
2450     If Len(sDate) Then
2451         PatchFilename = App.Path & IIf(Right$(App.Path, 1) <> gsBS, gsBS, vbNullString) & "MDBPatch" & Replace(Format$(sDate, "00000.00"), gsP, "-") & ".sad"
2452         SliceAndDice.GenerateDeltaPatchFile CVDate(sDate), PatchFilename
2453         If Len(Dir$(PatchFilename)) Then
2454             If bUserSure("File created successfully." & gsEolTab & "Filename:" & PatchFilename & gs2EOL & "Would you like to view it now ?") Then
2455                 On Error Resume Next
2456                 Shell WindowsDirectory & "NOTEPAD.EXE " & gsQ & PatchFilename & gsQ
2457             End If
2458         End If
2459     End If
End Sub

Public Sub mnuForward_Click()
2460     If Val(CurrentHistoryEntry) >= m_asaHistory.Count Then
2461         Beep
2462         mnuForward.Enabled = False
2463         Exit Sub
2464     End If

2465     CurrentHistoryEntry = Val(CurrentHistoryEntry) + 1
2466     JumpTo m_asaHistory(CurrentHistoryEntry), False, True
2467     mnuBack.Enabled = True
2468     mnuForward.Enabled = Val(CurrentHistoryEntry) < m_asaHistory.Count
End Sub

Public Sub mnuHelpAbout_Click()
2469     On Error Resume Next
2470     With frmSplash
2471         .lblDLLsLoaded(1).Caption = vbNullString & SadCommandSetCount
2472         .Show
2473     End With
End Sub

Private Sub mnuHelpEmailWilliamRawls_Click()
2474     BrowseTo "mailto:wrawls@firmsolutions.com"
End Sub

Private Sub mnuHelpReportIssue_Click()
2475     BrowseTo "http://www.sliceanddice.com/sadissue.html"
End Sub

Private Sub mnuHelpSoftCommandReference_Click()
2481     If SadCommandSetCount > 0 Then
2500             Complete.ShowHelpScreen
2501         'End If
2502     Else
2503         MsgBox "No command set DLLs loaded." & gsEolTab & "No Soft Command Reference available." & gsEolTab & "Make sure S&D DLLs are in the same directory as the .MDB you have loaded."
2504     End If
End Sub

Private Sub UpdateCompleteListOfSoftCommands()
2476     On Error Resume Next
2477     Dim CurrSet As Long
2480     Dim CurrCommand As CSadCommand

2487             If Not Complete Is Nothing Then
2488                 Complete.Clear False
                     Set Complete.Parent = Nothing
2489                 Set Complete = Nothing
2490             End If
2491             Set Complete = New CSadCommands
2492             For CurrSet = 1 To SadCommandSetCount
2493                 If SadCommands(CurrSet).CommandSet.Count > 0 Then
2494                     For Each CurrCommand In SadCommands(CurrSet).CommandSet
2495                         Complete.Append CurrCommand
2496                     Next CurrCommand
2497                 End If
2498             Next CurrSet
2499             Set Complete.Parent = Parent
End Sub

Private Sub mnuHelpVisitHomePage_Click()
2505     BrowseTo "http://www.sliceanddice.com"
End Sub

Private Sub mnuHistoryList_Click()
2506     On Error Resume Next
2507     Dim sChoices As String
2508     Dim sChoice As String

2509     If m_asaHistory.Count > 0 Then
2510         m_asaHistory.ItemDelimiter = gsSC
2511         sChoices = m_asaHistory.Column
2512         sChoice = sChoose(sChoices, , m_asaHistory(CurrentHistoryEntry).Value)
2513         If Len(sChoice) Then
2514             CurrentHistoryEntry = vbNullString & m_asaHistory.FindKey(sChoice)
2515             JumpTo m_asaHistory(CurrentHistoryEntry), False, True
2516             mnuForward.Enabled = Val(CurrentHistoryEntry) < m_asaHistory.Count
2517             mnuBack.Enabled = Val(CurrentHistoryEntry) > 1
2518         End If
2519     End If

End Sub

Private Sub mnuHelpOnlineDocumentation_Click()
2520     BrowseTo "http://www.sliceanddice.com/saddoc.html"
End Sub

Private Sub mnuIsFavorite_Click()
    If mbFillingAddInScreen Then Exit Sub
    chkFavorite = Abs(Not -chkFavorite)
    chkFavorite_Click
End Sub
    
Private Sub mnuIsUndeletable_Click()
    If mbFillingAddInScreen Then Exit Sub
    chkUndeletable = Abs(Not -chkUndeletable)
    chkUndeletable_Click
End Sub
    
Private Sub mnuIsCodeLocked_Click()
    If mbFillingAddInScreen Then Exit Sub
    chkLocked = Abs(Not -chkLocked)
    chkLocked_Click
End Sub

Public Sub mnuPasswordProtection_Click()
2524     mnuPasswordProtection.Checked = Not mnuPasswordProtection.Checked
2525     SaveSetting App.ProductName, "Last", "Password Protection", mnuPasswordProtection.Checked
End Sub

Public Sub mnuShowOnModuleRightClick_Click()
2526     mnuShowOnModuleRightClick.Checked = Not mnuShowOnModuleRightClick.Checked
2527     SaveSetting App.ProductName, "Last", "Show On Module Right Click", mnuShowOnModuleRightClick.Checked
2528     MsgBox "This will take effect the next time Visual Basic or " & gsSliceAndDice & " is restarted.", vbInformation
End Sub

Public Sub mnuShowPaintbrushIcon_Click()
2529     mnuShowPaintbrushIcon.Checked = Not mnuShowPaintbrushIcon.Checked
2530     SaveSetting App.ProductName, "Last", "Show Paitbrush Icon", mnuShowPaintbrushIcon.Checked
2531     MsgBox "This will take effect the next time Visual Basic is restarted.", vbInformation
End Sub


Public Sub mnuOLEDragDrop_Click()
2532     On Error Resume Next
2533     mnuOLEDragDrop.Checked = Not mnuOLEDragDrop.Checked
2534     SaveSetting App.ProductName, gsLast, "OLEDragDrop", mnuOLEDragDrop.Checked

2535     txtCode(0).OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0)
2536     txtCode(0).OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0)
2537     txtCode(1).OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0)
2538     txtCode(1).OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0)
2539     txtCode(2).OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0)
2540     txtCode(2).OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0)
         txtCodeToFile.OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0)
         txtCodeToFile.OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0)
End Sub

Public Sub mnuShowSplash_Click()
2543     mnuShowSplash.Checked = Not mnuShowSplash.Checked
2544     SaveSetting App.ProductName, "Last", "Show Splash", mnuShowSplash.Checked
End Sub

'Public Sub mnuSpecialExportTemplate_Click()
'    Dim sExportToDB As String
'    Dim db       As Database
'    Dim rst      As Recordset
'
'    sExportToDB = sChooseDatabase()
'
'    If sExportToDB = m_sTemplateDatabaseName Then
'       MsgBox "Can't export to the same database you are currently using. Use 'Copy Current' instead."
'       Exit Sub
'    End If
'
'On Error Resume Next
'    Set db = OpenDatabase(sExportToDB,False,False)
'        If Err.Number <> 0 Then
'           MsgBox "Error opening database '" & sExportToDB & "'. Aborting export."
'           Err.Clear
'           Exit Sub
'        End If
'        Set rst = db.OpenRecordset( gsSelectFrom & "Templates")
'            If Err.Number <> 0 Then
'               MsgBox "Error opening table 'Templates' in export to database. Aborting export."
'               Err.Clear
'               Exit Sub
'            End If
'            With rst
'                 .AddNew
'                 If Err.Number <> 0 Then
'                    MsgBox "Error adding a new record to the export to database. Aborting export."
'                    Err.Clear
'                    rst.Close
'                    db.Close
'                    Exit Sub
'                 End If
'                 !sName = txtName
'                 !ShortTemplateName = txtShortName
'                 !memoCodeAtTop = zn(txtCode(0))
'                 !memoCodeAtCursor = zn(txtCode(1))
'                 !memoCodeAtBottom = zn(txtCode(2))
'                 !memoCodeToFile = zn(txtCodeToFile)
'                 !Filename = zn(txtFilename)
'                 !Undeletable = chkUndeletable
'                 !Locked = chkLocked
'                 !IncludeInMenu = chkIncludeInMenu
'                 .Update
'                 If Err.Number <> 0 Then
'                    MsgBox "Can't export that record for some reason. Probably a " & gsTemplate & " with that name already exists in the export to database."
'                    Err.Clear
'                 End If
'            End With
'        rst.Close
'    db.Close
'
'End Sub

'Public Sub mnuSpecialViewLog_Click()
'    frmLog.Show vbModal
'End Sub

Public Sub mnuSwitchTabsAutomatically_Click()
2545     mnuSwitchTabsAutomatically.Checked = Not mnuSwitchTabsAutomatically.Checked
2546     SaveSetting App.ProductName, "Last", "Switch tabs automatically", mnuSwitchTabsAutomatically.Checked
End Sub

Public Sub mnuX_Click()
2547     mnuFileExit_Click
End Sub

Public Sub tmrActivateDBClassGen_Timer()
2548     If gbProcessing Then Exit Sub
2549     tmrActivateDBClassGen.Enabled = False

2550     DBClassGen.RefreshCategories
2551     DBClassGen.Show , Me
End Sub

Public Sub chkLocked_Click()
2552     Dim sPasswordCheck As String
2553     Static bInHereAlready As Boolean

2554     If mbFillingAddInScreen Then Exit Sub
2555     If bInHereAlready Then Exit Sub
2556     If Not mnuPasswordProtection.Checked Then
2557         lsbJumpTo.BarItemIcon = IIf(chkLocked.Value = 0 And chkUndeletable.Value = 0, "Document" & IIf(lsbJumpTo.BarItemIcon = "DocumentAlternate", "Alternate", vbNullString), "Key")
2558         Exit Sub
2559     End If

2560     bInHereAlready = True

2561     If chkLocked.Value = 0 Then
2562         If Len(CurrentTemplate.memoAttributes) Then
2563             m_asaAttributes.All = CurrentTemplate.memoAttributes
2564             If Len(m_asaAttributes("Locked Password")) Then
2565                 sPasswordCheck = InputBox("Enter password to unlock.", "ENTER PASSWORD")
2566                 If StrComp(sPasswordCheck, m_asaAttributes("Locked Password")) <> 0 Then
2567                     Beep
2568                     chkLocked.Value = 1
2569                     lsbJumpTo.BarItemIcon = "Key"
2570                 Else
2571                     chkLocked.Value = 0
2572                     lsbJumpTo.BarItemIcon = IIf(chkLocked.Value = 0 And chkUndeletable.Value = 0, "Document" & IIf(lsbJumpTo.BarItemIcon = "DocumentAlternate", "Alternate", vbNullString), "Key")
2573                 End If
2574             End If
2575         Else
2576         End If
2577     Else
2578         m_asaAttributes.All = CurrentTemplate.memoAttributes
2579         m_asaAttributes("Locked Password") = InputBox("Enter a password for unlocking later.", "ENTER NEW PASSWORD", m_asaAttributes("Locked Password"))
2580         CurrentTemplate.memoAttributes = m_asaAttributes.All
2581         chkLocked.Value = 1
2582         lsbJumpTo.BarItemIcon = "Key"
2583     End If


2584     txtCode(0).Enabled = (chkLocked.Value = 0)
2585     txtCode(1).Enabled = (chkLocked.Value = 0)
2586     txtCode(2).Enabled = (chkLocked.Value = 0)
2587     txtName.Enabled = (chkLocked.Value = 0)
2588     txtShortName.Enabled = (chkLocked.Value = 0)
2589     txtFilename.Enabled = (chkLocked.Value = 0)
2590     frmFile.Enabled = (chkLocked.Value = 0)

2591     bInHereAlready = False
End Sub

Public Sub mnuSpecialNewDatabase_Click()
2592     Dim sDatabasePath    As String
2593     Dim sNewDatabaseName As String
2594     Dim db               As Database
2595     Dim tblTemplates     As TableDef
2596     Dim fldTemplates     As Field
2597     Dim ndxTemplates     As Index
2598     Dim rstCategory      As Recordset

2599     sDatabasePath = Trim$(BrowseForFolder(DBClassGen.hwnd, "Where should database go ?"))
2600     If Len(sDatabasePath) = 0 Then Exit Sub

2601     sNewDatabaseName = Trim$(InputBox("What should the name of the new " & gsTemplate & " database be ?", "CREATE " & gsTemplate & " DATABASE"))
2602     If Len(sNewDatabaseName) = 0 Then Exit Sub

2603     If Right$(sDatabasePath, 1) <> gsBS Then sDatabasePath = sDatabasePath & gsBS
2604     If Right$(LCase$(sNewDatabaseName), 4) <> ".mdb" Then sNewDatabaseName = sDatabasePath & sNewDatabaseName & ".mdb"

2605     On Error GoTo mnuSpecialNewDatabase_Click
2606     Err.Clear
2607     Set db = CreateDatabase(sNewDatabaseName, dbLangGeneral, dbVersion30)
2608     If Err.Number <> 0 Then
2609         MsgBox "Error creating " & gsTemplate & " database. Aborting."
2610         Exit Sub
2611     End If

2612     Set tblTemplates = db.CreateTableDef("Category")
2613     With tblTemplates
2614         Set fldTemplates = .CreateField("CategoryID", dbLong)
2615         fldTemplates.Attributes = dbAutoIncrField
2616         .Fields.Append fldTemplates
2617         .Fields.Append .CreateField("CategoryName", dbText, 255)
2618         .Fields.Append .CreateField("CategoryType", dbLong)
2619         .Fields.Append .CreateField("ColumnWidth", dbSingle)
2620         .Fields.Append .CreateField("View", dbInteger)
2621         .Fields.Append .CreateField("Arrange", dbInteger)
2622         .Fields.Append .CreateField("DateCreated", dbDate)
2623         .Fields.Append .CreateField("DateModified", dbDate)
2624         .Fields.Append .CreateField("memoAttributes", dbMemo)

2625         Set ndxTemplates = .CreateIndex("PrimaryKey")
2626         With ndxTemplates
2627             .Fields.Append .CreateField("CategoryID")
2628             .Primary = True
2629             .Unique = True
2630             .Required = True
2631         End With
2632         .Indexes.Append ndxTemplates

2633         Set ndxTemplates = .CreateIndex("CategoryName")
2634         With ndxTemplates
2635             .Fields.Append .CreateField("CategoryName")
2636             .Primary = False
2637             .Unique = True
2638             .Required = True
2639         End With
2640         .Indexes.Append ndxTemplates

2641         Set ndxTemplates = Nothing

2642         db.TableDefs.Append tblTemplates
2643     End With

2644     Set tblTemplates = db.CreateTableDef("Template")
2645     With tblTemplates
2646         Set fldTemplates = .CreateField("TemplateID", dbLong)
2647         fldTemplates.Attributes = dbAutoIncrField
2648         .Fields.Append fldTemplates
2649         .Fields.Append .CreateField("CategoryID", dbLong)
2650         .Fields.Append .CreateField("TemplateName", dbText, 255)
2651         .Fields.Append .CreateField("ShortTemplateName", dbText, 255)
2652         .Fields.Append .CreateField("Filename", dbText, 255)
2653         .Fields.Append .CreateField("Undeletable", dbBoolean)
2654         .Fields.Append .CreateField("Locked", dbBoolean)
2655         .Fields.Append .CreateField("IncludeInMenu", dbBoolean)
2656         .Fields.Append .CreateField("memoCodeAtCursor", dbMemo)
2657         .Fields.Append .CreateField("memoCodeAtTop", dbMemo)
2658         .Fields.Append .CreateField("memoCodeAtBottom", dbMemo)
2659         .Fields.Append .CreateField("memoCodeToFile", dbMemo)
2660         .Fields.Append .CreateField("DateCreated", dbDate)
2661         .Fields.Append .CreateField("DateModified", dbDate)
2662         .Fields.Append .CreateField("memoAttributes", dbMemo)
2663         .Fields.Append .CreateField("Favorite", dbBoolean)
2664         .Fields.Append .CreateField("RevisionCount", dbLong)
2665         .Fields.Append .CreateField("TimerInsertion", dbText, 255)

2666         Set ndxTemplates = .CreateIndex("PrimaryKey")
2667         With ndxTemplates
2668             .Fields.Append .CreateField("TemplateID")
2669             .Primary = True
2670             .Unique = True
2671             .Required = True
2672         End With
2673         .Indexes.Append ndxTemplates

2674         Set ndxTemplates = .CreateIndex("CategoryID")
2675         With ndxTemplates
2676             .Fields.Append .CreateField("CategoryID")
2677             .Primary = False
2678             .Unique = False
2679             .Required = True
2680         End With
2681         .Indexes.Append ndxTemplates

2682         Set ndxTemplates = .CreateIndex("ShortTemplateName")
2683         With ndxTemplates
2684             .Fields.Append .CreateField("ShortTemplateName")
2685             .Primary = False
2686             .Unique = False
2687             .Required = True
2688         End With
2689         .Indexes.Append ndxTemplates

2690         Set ndxTemplates = .CreateIndex("TemplateName")
2691         With ndxTemplates
2692             .Fields.Append .CreateField("TemplateName")
2693             .Primary = False
2694             .Unique = False
2695             .Required = True
2696         End With
2697         .Indexes.Append ndxTemplates
2698         Set ndxTemplates = Nothing

2699         db.TableDefs.Append tblTemplates
2700     End With

2701     Set tblTemplates = db.CreateTableDef("SystemInfo")
2702     With tblTemplates
2703         Set fldTemplates = .CreateField("SystemInfoID", dbLong)
2704         fldTemplates.Attributes = dbAutoIncrField
2705         .Fields.Append fldTemplates
2706         .Fields.Append .CreateField("SystemInfoName", dbText, 255)
2707         .Fields.Append .CreateField("DateCreated", dbDate)
2708         .Fields.Append .CreateField("DateModified", dbDate)
2709         .Fields.Append .CreateField("memoAttributes", dbMemo)

2710         Set ndxTemplates = .CreateIndex("PrimaryKey")
2711         With ndxTemplates
2712             .Fields.Append .CreateField("SystemInfoID")
2713             .Primary = True
2714             .Unique = True
2715             .Required = True
2716         End With
2717         .Indexes.Append ndxTemplates

2718         Set ndxTemplates = .CreateIndex("SystemInfoName")
2719         With ndxTemplates
2720             .Fields.Append .CreateField("SystemInfoName")
2721             .Primary = False
2722             .Unique = True
2723             .Required = True
2724         End With
2725         .Indexes.Append ndxTemplates

2726         Set ndxTemplates = Nothing

2727         db.TableDefs.Append tblTemplates
2728     End With

2729     Set rstCategory = db.OpenRecordset("Category")
2730     With rstCategory
2731         .AddNew
2732         !CategoryName = "Basic"
2733         !CategoryType = 0
2734         !View = 3
2735         .Update
2736         .AddNew
2737         !CategoryName = "Change from"
2738         !CategoryType = 0
2739         !View = 3
2740         .Update
2741         .AddNew
2742         !CategoryName = "From the Internet"
2743         !CategoryType = 0
2744         !View = 3
2745         .Update
2746     End With

2747     db.Close

2748     m_sTemplateDatabaseName = sNewDatabaseName
2749     RefreshDatabaseConnection

2750 mnuSpecialNewDatabase_Click_Continue:
2751     Exit Sub

2752 mnuSpecialNewDatabase_Click:
2753     LogError "frmMain", "mnuSpecialNewDatabase_Click", Err.Number, Err.Description, Erl
2754     Resume mnuSpecialNewDatabase_Click_Continue:

2755     Resume
End Sub

Public Sub mnuSpecialOpenDatabase_Click()
2756     Dim sTemplateDatabaseName  As String
2757     Dim sOldDatabaseName       As String

2758     sTemplateDatabaseName = sChooseDatabase(App.Path)
2759     If Len(sTemplateDatabaseName) Then
2760         SaveTemplate
2761         sOldDatabaseName = m_sTemplateDatabaseName

             UpdateRecentFileList

2762         m_sTemplateDatabaseName = sTemplateDatabaseName
2763         If RefreshDatabaseConnection Then
2764             SaveSetting App.ProductName, "Settings", "Current Database", sTemplateDatabaseName
2765             If Not SliceAndDice(1) Is Nothing Then
2766                 If Not SliceAndDice(1).Templates(1) Is Nothing Then
2767                     JumpTo SliceAndDice(1).Templates(1).Key
2768                     lsbJumpTo.BarAndItem SliceAndDice(1).Key, SliceAndDice(1).Templates(1).ShortTemplateName
2769                 ElseIf Not SliceAndDice(2) Is Nothing Then
2770                     If Not SliceAndDice(2).Templates(1) Is Nothing Then
2771                         JumpTo SliceAndDice(2).Templates(1).Key
2772                         lsbJumpTo.BarAndItem SliceAndDice(2).Key, SliceAndDice(2).Templates(1).ShortTemplateName
2773                     End If
2774                 End If
2775             End If
2776         Else
2777             m_sTemplateDatabaseName = sOldDatabaseName
2778         End If
2779     End If
End Sub

Public Sub tabCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
2780     On Error GoTo EH_frmMain_tabCode_MouseUp
2781     Static bInHereAlready As Boolean

    Select Case tabCode.SelectedItem.Index
           Case 1
2782             txtCode(0).Visible = True
2783             txtCode(1).Visible = False
2784             txtCode(2).Visible = False
2785             frmFile.Visible = False
2786             frmOptions.Visible = False
2787             frmTemplateInfo.Visible = False
On Error Resume Next
                 txtCode(0).SetFocus

2788       Case 2
2789             txtCode(0).Visible = False
2790             txtCode(1).Visible = True
2791             txtCode(2).Visible = False
2792             frmFile.Visible = False
2793             frmOptions.Visible = False
2794             frmTemplateInfo.Visible = False
On Error Resume Next
                 txtCode(1).SetFocus

2795       Case 3
2796             txtCode(0).Visible = False
2797             txtCode(1).Visible = False
2798             txtCode(2).Visible = True
2799             frmFile.Visible = False
2800             frmOptions.Visible = False
2801             frmTemplateInfo.Visible = False
On Error Resume Next
                 txtCode(2).SetFocus

2802       Case 4
2803             txtCode(0).Visible = False
2804             txtCode(1).Visible = False
2805             txtCode(2).Visible = False
2806             frmFile.Visible = True
2807             frmOptions.Visible = False
2808             frmTemplateInfo.Visible = False
On Error Resume Next
                 txtCodeToFile.SetFocus

2809       Case 5
2810             txtCode(0).Visible = False
2811             txtCode(1).Visible = False
2812             txtCode(2).Visible = False
2813             frmFile.Visible = False
2814             frmOptions.Visible = True
2815             frmTemplateInfo.Visible = False
On Error Resume Next
                 chkFavorite.SetFocus

2816       Case 6
2817             txtCode(0).Visible = False
2818             txtCode(1).Visible = False
2819             txtCode(2).Visible = False
2820             frmFile.Visible = False
2821             frmOptions.Visible = False
2822             frmTemplateInfo.Visible = True
On Error Resume Next
                 cmdRecalc.SetFocus

2823             If chkAutoRecalc.Value <> 0 Then
2824                 cmdRecalc_Click
2825             End If
2826     End Select

2827 EH_frmMain_tabCode_MouseUp_Continue:
2828     bInHereAlready = False
2829     Exit Sub

2830 EH_frmMain_tabCode_MouseUp:
2831     LogError "frmMain", "tabCode_MouseUp", Err.Number, Err.Description, Erl
2832     Resume EH_frmMain_tabCode_MouseUp_Continue

2833     Resume
End Sub

Private Sub tmrDoAction_Timer()
2834     On Error Resume Next
2835     If Not OkayToDoAction Then Exit Sub

2836     tmrDoAction.Enabled = False

    Select Case UCase$(ActionToDo)
        Case "NEWTEMPLATE"
2837             NewTemplate True, ActionParam

            'Case "DELTACHECK", "DELTA CHECK"
            '     If Len(Dir$(Parent.TemplateDatabasePath & "MDBPatch*.sad", vbNormal)) Then
            '        If bUserSure("A Delta Patch file has been found. Would you like to apply it now ?") Then
            '           SliceAndDice.ApplyPatch Dir$(Parent.TemplateDatabasePath & "MDBPatch*.sad", vbNormal)
            '        End If
            '     End If
            '     QueueAction "DeltaCheck", vbNullString, 65535
            '     OkayToDoAction = True

2838         Case "DOINSERTION"
2839             DoInsertion Nothing, ActionParam
2840     End Select
End Sub

Public Sub txtCode_GotFocus(Index As Integer)
2841     CurrentCodeArea = Index
End Sub

Public Sub txtCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
       Form_KeyDown KeyCode, Shift
       If mbScramFormKey Then KeyCode = 0: Shift = 0
End Sub


' ================================================================================
' Name              Form_GotFocus
'
' Parameters
'      None
'
' Description
'
' When the form first shows, insure the list of Templates is selected since the
' user is most likely going to insert a Template
'
' ================================================================================
Public Sub Form_GotFocus()
2843     On Error Resume Next
2844     lsbJumpTo.SetFocus                                ' More than likely the user is going to want to insert a pre-existing Template.
End Sub

' ================================================================================
' Name              Form_Initialize
'
' Parameters
'      None
'
' Description
'
' Takes care of positioning the form to its previous location and redrawing
' everything to make it look good when it is first seen.
'
' ================================================================================
Public Sub Form_Initialize()
2845     Dim sLastTemplate As String
2846     Dim sCategory As String
2847     Dim sShortName As String

2848     InitPublic

2849     mnuExitAfterInsert.Checked = GetSetting(App.ProductName, "Settings", "Exit after insert", True)

2850     mnuShowPaintbrushIcon.Checked = GetSetting(App.ProductName, "Last", "Show Paitbrush Icon", True)
2851     mnuShowOnModuleRightClick.Checked = GetSetting(App.ProductName, "Last", "Show On Module Right Click", True)

2852     mnuSwitchTabsAutomatically.Checked = GetSetting(App.ProductName, "Last", "Switch tabs automatically", True)
2853     mnuPasswordProtection.Checked = GetSetting(App.ProductName, "Last", "Password Protection", False)
2854     mnuShowSplash.Checked = GetSetting(App.ProductName, "Last", "Show Splash", True)
2855     mnuOLEDragDrop.Checked = GetSetting(App.ProductName, "Last", "OLEDragDrop", False)

            txtCode(0).OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0):     txtCode(0).OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0)
            txtCode(1).OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0):     txtCode(1).OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0)
            txtCode(2).OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0):     txtCode(2).OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0)
            txtCodeToFile.OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0):  txtCodeToFile.OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0)
         
2856     chkAutoRecalc.Value = GetSetting(App.ProductName, "Last", "Auto Recalc", 0)

2857     lsbJumpTo.Arrange = GetSetting(App.ProductName, "Settings", "Bar Arrange", "1")
2858     lsbJumpTo.View = GetSetting(App.ProductName, "Settings", "Bar View", "1")

2859     If Len(Dir$(App.Path & gsBS & "SliceAndDice.mdb")) = 0 And Len(Dir$(App.Path & gsBS & "SliceAndDiceNew.mdb")) <> 0 Then
2860         Name App.Path & gsBS & "SliceAndDiceNew.mdb" As App.Path & gsBS & "SliceAndDice.mdb"
2861     End If
2862     m_sTemplateDatabaseName = GetSetting$(App.ProductName, "Settings", "Current Database")
2863     If Len(Dir$(m_sTemplateDatabaseName)) = 0 Then
2864         m_sTemplateDatabaseName = 0
2865     End If
2866 UserDoc_Init_Try_Again:
2867     If Len(m_sTemplateDatabaseName) = 0 Then
2868         If Dir$(App.Path & gsBS & "SliceAndDice.mdb") = vbNullString Then
2869             m_sTemplateDatabaseName = sChooseDatabase(App.Path, "SliceAndDice.mdb")
2870         Else
2871             m_sTemplateDatabaseName = App.Path & gsBS & "SliceAndDice.mdb"
2872         End If

2873         If Len(m_sTemplateDatabaseName) = 0 Then
2874             MsgBox "A " & gsTemplate & " database must be chosen. Please try again."
2875             GoTo UserDoc_Init_Try_Again
2876         Else
2877             SaveSetting App.ProductName, "Settings", "Current Database", m_sTemplateDatabaseName
2878         End If
2879     End If

2880     If Not RefreshDatabaseConnection Then
2881         m_sTemplateDatabaseName = vbNullString
2882         GoTo UserDoc_Init_Try_Again
2883     End If

2884     Form_Resize                                       ' Force redraw to make sure everything looks good

2885     sLastTemplate = GetSetting$(App.ProductName, "Settings", "Last " & gsTemplate)
2886     If Len(sLastTemplate) = 0 Then
2887         sLastTemplate = "Release Notes - Welcome"
2888     End If

2889     On Error Resume Next
2890     GetCategoryAndName sLastTemplate, sCategory, sShortName
2891     JumpTo sLastTemplate, False, True
2892     DoEvents
    'lsbJumpTo.DisplayCategories

2893     Err.Clear

    'lsbJumpTo.DisplayCategories

    ' LogEvent "frmMain: Initialize"
End Sub

Public Function sChooseDatabase(Optional ByVal sPath As String, Optional ByVal sFilename As String) As String
2894     On Error Resume Next
2895     Err.Clear
2896     With cdgSelect
2897         .Filter = "Access Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
2898         .FilterIndex = 0
2899         If Len(sPath) > 0 Then .InitDir = sPath
2900         If Len(sFilename) > 0 Then .FileName = sFilename
2901         .ShowOpen
2902         If Err <> 0 Then
2903             Err.Clear
2904             Exit Function
2905         End If
2906         sChooseDatabase = .FileName
2907     End With
End Function

Public Function sChooseFile(Optional ByVal sPath As String, Optional ByVal sFilename As String, Optional ByVal sFilter As String = vbNullString) As String
2908     On Error Resume Next
2909     Err.Clear
2910     With cdgSelect
2911         .Filter = IIf(Len(sFilter) And InStr(sFilter, "|"), sFilter, "All Files (*.*)|*.*")
2912         .FilterIndex = 0
2913         If Len(sPath) > 0 Then .InitDir = sPath
2914         If Len(sFilename) > 0 Then .FileName = sFilename
2915         .ShowOpen
2916         If Err <> 0 Then
2917             Err.Clear
2918             Exit Function
2919         End If
2920         sChooseFile = .FileName
2921     End With
End Function

Public Function sChooseColor(Optional ByVal sInitialColor As String) As String
2922     On Error Resume Next
2923     Dim Red As Integer
2924     Dim Green As Integer
2925     Dim Blue As Integer

2926     Err.Clear
2927     With cdgSelect
2928         .CancelError = True
2929         If lTokenCount(sInitialColor, gsSC) = 3 Then
2930             Red = sGetToken(sInitialColor, 1, gsSC)
2931             Green = sGetToken(sInitialColor, 1, gsSC)
2932             Blue = sGetToken(sInitialColor, 1, gsSC)
2933             If Red > 255 Then Red = 255
2934             If Red < 0 Then Red = 0
2935             If Green > 255 Then Green = 255
2936             If Green < 0 Then Green = 0
2937             If Blue > 255 Then Blue = 255
2938             If Blue < 0 Then Blue = 0
2939             .Color = RGB(Red, Green, Blue)
2940         ElseIf Val(sInitialColor) > 0 Then
2941             .Color = Val(sInitialColor)
2942         End If
2943         On Error GoTo ErrHandler
2944         .Flags = cdlCCRGBInit
2945         .ShowColor
2946         sChooseColor = Hex$(.Color)
2947     End With
2948 ErrHandler:
    ' User pressed Cancel button.
2949     Exit Function
End Function

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
2950     On Error GoTo EH_frmMain_Form_KeyDown
2951     Dim sSelectedText   As String
2952     Dim sOrigText       As String
2953     Dim CurrSet         As Long

         mbScramFormKey = False

2954     If (Shift And vbShiftMask) > 0 Then               ' Shift Key         *******************
                Select Case KeyCode
                    
                    Case vbKeyInsert: KeyCode = 0: Shift = 0: mbScramFormKey = True   ' Paste
2955                     If TypeOf ActiveControl Is TextBox Then
2956                        ActiveControl.SelText = Clipboard.GetText
2957                     End If
                         mbScramFormKey = True
2958                Case vbKeyDelete: KeyCode = 0: Shift = 0: mbScramFormKey = True   ' Cut
2959                     If TypeOf ActiveControl Is TextBox Then
2960                        If StringToClipboard(ActiveControl.SelText) Then
2961                           ActiveControl.SelText = vbNullString
2962                        End If
2963                     End If
                         mbScramFormKey = True
2964            End Select



2965     ElseIf (Shift And vbCtrlMask) > 0 Then            ' CTRL Key      ********************
                Select Case KeyCode
                    Case vbKeyInsert: KeyCode = 0: Shift = 0: mbScramFormKey = True   ' Copy
2966                     If TypeOf ActiveControl Is TextBox Then
2967                        StringToClipboard ActiveControl.SelText
2968                     End If
        
2969                Case vbKeyDelete: KeyCode = 0: Shift = 0: mbScramFormKey = True   ' Cut
2970                     If TypeOf ActiveControl Is TextBox Then
2971                        If StringToClipboard(ActiveControl.SelText) Then
2972                           ActiveControl.SelText = vbNullString
2973                        End If
2974                     End If
        
2975                Case vbKeyTab
2977                     KeyCode = 0: Shift = 0: mbScramFormKey = True
  If Parent.HostedByVB Then  ' Shell App override
2976                     On Error Resume Next
2978                     Parent.vbInst.ActiveWindow.SetFocus
  End If

2979                Case vbKeyF:    KeyCode = 0: Shift = 0: mbScramFormKey = True: FindInCurrent
2980                Case vbKeyH:    KeyCode = 0: Shift = 0: mbScramFormKey = True: FindInCurrent False, True

2982                Case vbKeyI:    KeyCode = 0: Shift = 0: mbScramFormKey = True: mnuInsertTemplate_Click
2983                Case vbKeyL:    KeyCode = 0: Shift = 0: mbScramFormKey = True: mnuFileRefresh_Click
2984                Case vbKeyM:    KeyCode = 0: Shift = 0: mbScramFormKey = True: mnuFileImport_Click
2985                Case vbKeyN:    KeyCode = 0: Shift = 0: mbScramFormKey = True: mnuFileNew_Click
        
                    Case vbKey1:    KeyCode = 0: Shift = 0: mbScramFormKey = True: lsbJumpTo.SetFocus
                    Case vbKey2:    KeyCode = 0: Shift = 0: mbScramFormKey = True: txtShortName.SetFocus
                    Case vbKey3:    KeyCode = 0: Shift = 0: mbScramFormKey = True: tabCode.Tabs(1).Selected = True: tabCode_MouseUp 0, 0, 0, 0
                    Case vbKey4:    KeyCode = 0: Shift = 0: mbScramFormKey = True: tabCode.Tabs(2).Selected = True: tabCode_MouseUp 0, 0, 0, 0
                    Case vbKey5:    KeyCode = 0: Shift = 0: mbScramFormKey = True: tabCode.Tabs(3).Selected = True: tabCode_MouseUp 0, 0, 0, 0
                    Case vbKey6:    KeyCode = 0: Shift = 0: mbScramFormKey = True: tabCode.Tabs(4).Selected = True: tabCode_MouseUp 0, 0, 0, 0
                    Case vbKey7:    KeyCode = 0: Shift = 0: mbScramFormKey = True: tabCode.Tabs(5).Selected = True: tabCode_MouseUp 0, 0, 0, 0
                    Case vbKeyUp:
                        MsgBox "'Jump to previous template still to be written"
'                    KeyCode = 0: Shift = 0: mbScramFormKey = True
'                         Dim CurrItem As ListItem
'                         Dim CurrNode As Node
'                         Dim CurrListBar As FSListBar
'
'                         If Not lsbJumpTo.SelectedItem Is Nothing Then
'                            Set CurrListBar = lsbJumpTo.CurBar
'                            lsbJumpTo.SelectedItem.Index
'                         Else
'                         End If
                         
2986           End Select



2987     ElseIf (Shift And vbAltMask) > 0 Then                      ' Alt Key       *********************
                Select Case KeyCode
                    Case vbKeyLeft: KeyCode = 0: Shift = 0: mbScramFormKey = True: mnuBack_Click
2988                     Case vbKeyRight: KeyCode = 0: Shift = 0: mbScramFormKey = True: mnuForward_Click
        
2989                     Case vbKeyX: KeyCode = 0: Shift = 0: mbScramFormKey = True: mnuFileExit_Click
        
2990                     Case vbKeyTab
2991                         On Error Resume Next
2992                         KeyCode = 0: Shift = 0: mbScramFormKey = True
  If Parent.HostedByVB Then  ' Shell App override
2993                         Parent.vbInst.ActiveWindow.SetFocus
  End If
2994            End Select
2995     Else                                              'If (Shift And vbShiftMask) = 0 Then     ' No special modifying keys
        Select Case KeyCode
            Case vbKeyTab
2996                 If TypeOf ActiveControl Is TextBox Then
2997                     If InStr(ActiveControl.Tag, "Code Area ") Then
2998                         ActiveControl.SelText = vbTab
2999                         KeyCode = 0
3000                         Shift = 0
3001                     End If
3002                 End If

3003             Case vbKeyF1
3004                 If TypeOf ActiveControl Is TextBox Then
3005                     KeyCode = 0: Shift = 0: mbScramFormKey = True

3006                     sOrigText = Trim$(ActiveControl.SelText)
3007                     If Len(sOrigText) = 0 Then
3008                         sOrigText = sGetCurrentLineAtCharacter(ActiveControl.Text, ActiveControl.SelStart)
3009                     End If
3010                     sSelectedText = sOrigText
3011                     If Len(sSelectedText) > 0 Then
3012                         If InStr(sSelectedText, gsSoftCmdDelimiter) Then
3013                             If InStr(sSelectedText, gsSoftCmdDelimiter & "_") = 0 Then
3014                                 sSelectedText = UCase$(sGetToken(sGetToken(sSelectedText, 2, gsSoftCmdDelimiter), 1)) & "*C"
3015                             Else
3016                                 sSelectedText = UCase$(sGetToken(sGetToken(sSelectedText, 3, "_"), 1)) & "*C"
3017                             End If
3018                         End If
3019                         If InStr(sSelectedText, gsSoftVarDelimiter) Then
3020                             sSelectedText = sGetToken(sGetToken(sSelectedText, 2, gsSoftVarDelimiter), 1, gsInlineCmdDelimiter) & "*I"
                             ElseIf InStr(sSelectedText, "~##~") Then
                                 sSelectedText = sGetToken(sSelectedText, 1, " ") & "*I"
3021                         End If
                             If Not Complete Is Nothing Then
                                Complete.ShowHelpScreen sSelectedText
3022                         ElseIf SadCommandSetCount > 0 Then
3023                             For CurrSet = 1 To SadCommandSetCount
3024                                 If Not SadCommands(CurrSet).CommandSet.Item(sSelectedText) Is Nothing Then
3025                                     SadCommands(CurrSet).CommandSet.ShowHelpScreen sSelectedText
3026                                     GoTo EH_frmMain_Form_KeyDown_Continue
3027                                 End If
3028                             Next CurrSet
3029                             MsgBox "SoftCommand '" & Trim$(sOrigText) & "' not found."
3030                         Else
3031                             MsgBox "No command set DLLs loaded. No help available."
3032                         End If
3033                     Else
3034                         MsgBox "Unable to determine which command to display help for."
3035                     End If
3036                 End If

3037             Case vbKeyF3: KeyCode = 0: Shift = 0: mbScramFormKey = True: FindInCurrent True
3038         End Select
3039     End If

3040 EH_frmMain_Form_KeyDown_Continue:
3041     Exit Sub

3042 EH_frmMain_Form_KeyDown:
3043     LogError "frmMain", "Form_KeyDown", Err.Number, Err.Description, Erl
3044     Resume EH_frmMain_Form_KeyDown_Continue

3045     Resume
End Sub

Public Sub FindInCurrent(Optional ByVal bRepeatLastSearch As Boolean = False, Optional ByVal bReplace As Boolean = False, Optional ByVal bAuto As Boolean = False)
3046     On Error GoTo EH_frmMain_FindInCurrent
3047     Static bInHereAlready As Boolean
3048     If bInHereAlready Then Exit Sub
3049     bInHereAlready = True

3050     Dim CurCodeArea As Long
3051     Dim lLastFound As Long
3052     Dim bSomethingFound As Boolean

3053     If CurrentTemplate Is Nothing Then
3054         MsgBox "Please select a " & gsTemplate & " before selecting to search."
3055         Exit Sub
3056     End If

3057     With frmFindReplace
        Select Case tabCode.SelectedItem.Index
            Case 1 To 4
3058                 If Len(.txtFind.Text) = 0 Then
3059                     .txtFind = txtCode(Me.CurrentCodeArea).SelText
3060                 Else
3061                     .txtFind.SelStart = 0
3062                     .txtFind.SelLength = Len(.txtFind.Text)
3063                 End If
3064             Case 6
3065                 If lstSoftVariables.ListIndex > -1 Then
3066                     .txtFind = Replace(lstSoftVariables, "* ", vbNullString)
3067                     .txtReplace = vbNullString
3068                 ElseIf lstSoftCommands.ListIndex > -1 Then
3069                     .txtFind = lstSoftCommands
3070                     .txtReplace = vbNullString
3071                 End If
3072         End Select

3073         If bAuto Then
3074             .DoFindNext = True
3075             .DoReplace = False
3076             .DoReplaceAll = False
3077             .Canceled = False
3078         ElseIf Not bRepeatLastSearch Then
3079             .IsReplace = bReplace
3080         ElseIf bReplace Then
3081             .DoReplace = True
3082         Else
3083             .DoFindNext = True
3084         End If

3085         If Len(.txtFind) = 0 Then
            ' Avoid a nasty endless loop
3086             Beep
3087         ElseIf .DoReplaceAll Then
3088             Screen.MousePointer = vbHourglass
3089             SaveTemplate
            Select Case .SearchArea
                Case SearchAreaCurrentPane: txtCode(CurrentCodeArea).Text = Replace(txtCode(CurrentCodeArea).Text, .txtFind, .txtReplace)
3090                 Case SearchAreaCurrentTemplate: CurrentTemplate.Replace .txtFind, .txtReplace
3091                 Case SearchAreaCurrentCategory: SliceAndDice.Categorys(CurrentTemplate.ParentKey).Replace .txtFind, .txtReplace
3092                 Case SearchAreaCurrentDatabase: SliceAndDice.Categorys.Replace .txtFind, .txtReplace
3093             End Select
3094             SliceAndDice.Save
3095             FillAddInScreen
3096             Screen.MousePointer = vbDefault
3097         ElseIf .DoFindNext Or .DoReplace Then
3098             For CurCodeArea = 0 To 2
3099                 lLastFound = txtCode(CurCodeArea).SelStart + txtCode(CurCodeArea).SelLength
3100                 If lLastFound = 0 Then lLastFound = 1
3101                 If .chkMatchCase.Value <> 0 Then
3102                     lLastFound = InStr(Mid$(txtCode(CurCodeArea), lLastFound), .txtFind)
3103                 Else
3104                     lLastFound = InStr(UCase$(Mid$(txtCode(CurCodeArea), lLastFound)), UCase$(.txtFind))
3105                 End If

3106                 If lLastFound > 0 Then
3107                     bSomethingFound = True
3108                     With txtCode(CurCodeArea)
3109                         tabCode.Tabs(CurCodeArea + 1).Selected = True
3110                         tabCode_MouseUp 0, 0, 0, 0
3111                         On Error Resume Next
3112                         .SetFocus
3113                         lLastFound = lLastFound - 2 + IIf(.SelStart + .SelLength = 0, 1, .SelStart + .SelLength)
3114                         .SelStart = lLastFound
3115                         .SelLength = Len(frmFindReplace.txtFind)
3116                         If frmFindReplace.DoReplace Then
3117                             .SelText = frmFindReplace.txtReplace
3118                         End If
3119                     End With
3120                 End If
3121             Next CurCodeArea

3122             If Not bSomethingFound Then MsgBox "Search text not found."
3123         End If
3124     End With

3125 EH_frmMain_FindInCurrent_Continue:
3126     bInHereAlready = False
3127     Exit Sub

3128 EH_frmMain_FindInCurrent:
3129     MsgBox "Error occured in:" & gsEolTab & "Module: frmMain" & gsEolTab & "Procedure: FindInCurrent" & gs2EOL & Err.Description
3130     Resume EH_frmMain_FindInCurrent_Continue

3131     Resume
End Sub

' ================================================================================
' Name              Form_Resize
'
' Parameters
'      None
'
' Description
'
' This code makes sure everything looks good after a form resize.
'
' ================================================================================
Public Sub Form_Resize()
3132     On Error GoTo EH_Form_Resize
3133     With tabCode                                      ' Position the code entry areas
3134         .Height = ScaleHeight - .Top
        'lsbJumpTo.Height = ScaleHeight - 415
3135         If ScaleWidth - .Left < 0 Then Exit Sub       ' If there isn't enough display area to show the code entry areas, don't attempt to redraw it
3136         .Width = ScaleWidth - .Left
3137         txtName.Move lblCode(3).Left + lblCode(3).Width + 40, txtName.Top, .Width - (lblCode(3).Left + lblCode(3).Width + 40 - .Left), txtName.Height
3138         txtShortName.Move lblCode(3).Left + lblCode(3).Width + 40, txtName.Top, .Width - (lblCode(3).Left + lblCode(3).Width + 40 - .Left), txtName.Height

3139         txtCode(0).Move .Left + 100, .Top + 500, .Width - 200, .Height - 600
3140         txtCode(1).Move .Left + 100, .Top + 500, .Width - 200, .Height - 600
3141         txtCode(2).Move .Left + 100, .Top + 500, .Width - 200, .Height - 600
3142         frmOptions.Move .Left + 100, .Top + 500, .Width - 200, .Height - 600
3143         frmFile.Move .Left + 100, .Top + 500, .Width - 200, .Height - 600
3144         txtFilename.Width = frmFile.Width - txtFilename.Left * 2
3145         txtCodeToFile.Width = txtFilename.Width
3146         txtCodeToFile.Height = frmFile.Height - txtCodeToFile.Top - 100
3147         frmTemplateInfo.Move .Left + 100, .Top + 500, .Width - 200, .Height - 600
3148         lstSoftVariables.Width = (frmTemplateInfo.Width - lstSoftVariables.Left * 3) \ 2
3149         lstSoftCommands.Left = lstSoftVariables.Left * 2 + lstSoftVariables.Width
3150         lstSoftCommands.Width = lstSoftVariables.Width
3151         lblTemplateInfo(0).Left = lstSoftVariables.Left
3152         lblTemplateInfo(1).Left = lstSoftCommands.Left
3153         lstSoftVariables.Height = frmTemplateInfo.Height - lstSoftVariables.Top - 100
3154         lstSoftCommands.Height = frmTemplateInfo.Height - lstSoftVariables.Top - 100
3155     End With

3156 EH_Form_Resize_Continue:
3157     Exit Sub

3158 EH_Form_Resize:
3159     Resume EH_Form_Resize_Continue:

3160     Resume
End Sub

Public Sub mnuDBClassGen_Click()
3161     tmrActivateDBClassGen.Enabled = True
End Sub

Public Function sPropertyType(sFieldType As String) As String
    Select Case sFieldType
        Case "Big Integer": sPropertyType = "Long"
3162         Case "Binary": sPropertyType = "Variant"
3163         Case "Boolean": sPropertyType = "Boolean"
3164         Case "Byte": sPropertyType = "Byte"
3165         Case "Char": sPropertyType = "String"
3166         Case "Currency": sPropertyType = "Currency"
3167         Case "Date / Time": sPropertyType = "Date"
3168         Case "Decimal": sPropertyType = "Variant"
3169         Case "Double": sPropertyType = "Double"
3170         Case "Float": sPropertyType = "Double"
3171         Case "Guid": sPropertyType = "String"
3172         Case "Integer": sPropertyType = "Integer"
3173         Case "Long": sPropertyType = "Long"
3174         Case "Long Binary (OLE Object)": sPropertyType = "Variant"
3175         Case "Memo": sPropertyType = "Memo"
3176         Case "Numeric": sPropertyType = "Variant"
3177         Case "Single": sPropertyType = "Single"
3178         Case "Text": sPropertyType = "String"
3179         Case "Time": sPropertyType = "Date"
3180         Case "Time Stamp": sPropertyType = "Date"
3181         Case "VarBinary": sPropertyType = "Variant"
3182         Case Else: sPropertyType = "Variant"
3183     End Select
End Function


' ================================================================================
' Name              frmMain_mnuExitAfterInsert_Click
'
' Parameters
'      None
'
' Description
'
' Toggle the menu item's checked state
'
' ================================================================================
Public Sub mnuExitAfterInsert_Click()
3184     mnuExitAfterInsert.Checked = Not mnuExitAfterInsert.Checked
End Sub

' ================================================================================
' Name              frmMain_mnuFileCopy_Click
'
' Parameters
'      None
'
' Description
'
' This code takes care of copying the current Template and assigning it a unique
' name ala explorer.
'
' ================================================================================
Public Sub mnuFileCopy_Click()
3185     Dim sCategory As String
3186     Dim sShortName As String
3187     Dim sName As String
3188     Dim sCode(0 To 4) As String

3189     If CurrentTemplate Is Nothing Then
3190         MsgBox "Please select a " & gsTemplate & " to copy before selecting this option."
3191         Exit Sub
3192     End If

3193     sName = txtName.Text                              ' Save the contents of the Template to copy
3194     GetCategoryAndName sName, sCategory, sShortName

3195     sCode(0) = txtCode(0).Text
3196     sCode(1) = txtCode(1).Text
3197     sCode(2) = txtCode(2).Text
3198     sCode(3) = txtCodeToFile
3199     sCode(4) = txtFilename

3200     On Error Resume Next

3201     sName = sCategory & gsCategoryTemplateDelimiter & sShortName & gsS & Abs(NextNegativeUnique)
3202     sName = InputBox("What should the name of this " & gsTemplate & " be ?" & gsEolTab & "(Blank to cancel)" & gs2EOL & "Format of name MUST be:" & gsEolTab & gsCategory & " Name - " & gsTemplate & " Name", "NEW " & gsTemplate, sName)
3203     If Len(sName) > 0 Then
3204         NewTemplate True, sName
3205         txtCode(0).Text = sCode(0)                    ' Paste in the code from the template to copy
3206         txtCode(1).Text = sCode(1)
3207         txtCode(2).Text = sCode(2)
3208         txtCodeToFile = sCode(3)
3209         txtFilename = sCode(4)
3210     End If
End Sub

' ================================================================================
' Name              frmMain_mnuFileDelete_Click
'
' Parameters
'      None
'
' Description
'
' This procedure causes the current template to get deleted. The user is prompted
' to make sure this is what they want to do.
'
' ================================================================================
Public Sub mnuFileDelete_Click()
3211     DeleteTemplate
End Sub

' ================================================================================
' Name              frmMain_mnuFileExit_Click
'
' Parameters
'      None
'
' Description
'
' Whenever the add-in is exited (after an insert or on user request), this
' procedure makes sure the template record being edited (if it is) is saved.
'
' ================================================================================
Public Sub mnuFileExit_Click()
3212     SaveTemplate

3213     Hide
3214     HideAllWindows
    'VBIDEWindow.Visible = False      '   So hiding it will return control to VB
End Sub

' ================================================================================
' Name              frmMain_mnuFileImport_Click
'
' Parameters
'      None
'
' Description
'
' This takes care of importing the current code selection from the VB environment
' as a new template.
'
' ================================================================================
Public Sub mnuFileImport_Click()
3215     Dim lLine      As Long
3216     Dim lLastLine  As Long
3217     Dim lTemp      As Long
3218     Dim lFirstCol  As Long
3219     Dim lLastCol   As Long
3220     Dim sCode      As String
3221     Dim sProcName  As String
3222     Dim lProcType  As Long

         If Not Parent.HostedByVB Then Exit Sub

3223     On Error Resume Next
3224     With Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane
3225         .GetSelection lLine, lFirstCol, lLastLine, lLastCol    ' Determine where the cursor is
3226         lLastLine = lLastLine - lLine + Abs(lLastCol > 1)    ' Determine what the last line selected is (discard last line if at beginning)
3227         sCode = .CodeModule.Lines(lLine, lLastLine)   ' Grab the code selected from the active pane
3228         GetProcAtLine lLine, sProcName, lProcType
3229     End With

3230     If Len(sCode) = 0 Then
3231         MsgBox "Nothing to import.", vbInformation
3232         Exit Sub
3233     End If

3234     If Len(sProcName) Then
3235         NewTemplate , , sProcName
3236     Else
3237         NewTemplate
3238     End If

3239     txtCode(2).Text = sCode
3240     tabCode.Tabs(2).Selected = True
3241     tabCode_MouseUp 0, 0, 0, 0
End Sub

Public Function GetCurrentTextSelection() As String
3242     Dim lLine      As Long
3243     Dim lLastLine  As Long
3244     Dim lFirstCol  As Long
3245     Dim lLastCol   As Long
3246     Dim sCode      As String

         If Not Parent.HostedByVB Then Exit Function

3247     On Error Resume Next
3248     With Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane
3249         .GetSelection lLine, lFirstCol, lLastLine, lLastCol    ' Determine where the cursor is
3250         lLastLine = lLastLine - lLine + Abs(lLastCol > 1)    ' Determine what the last line selected is (discard last line if at beginning)
3251         GetCurrentTextSelection = .CodeModule.Lines(lLine, lLastLine)    ' Grab the code selected from the active pane
3252     End With
End Function

Public Sub DeleteCurrentTextSelection()
3253     Dim lLine As Long
3254     Dim lLastLine As Long
3255     Dim lFirstCol As Long
3256     Dim lLastCol As Long
3257     Dim sCode As String

         If Not Parent.HostedByVB Then Exit Sub

3258     On Error Resume Next
3259     With Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane
3260         .GetSelection lLine, lFirstCol, lLastLine, lLastCol
3261         lLastLine = lLastLine - lLine + Abs(lLastCol > 1)
3262         .CodeModule.DeleteLines lLine, lLastLine
3263     End With
End Sub

Public Function DetermineLastLineInSelection() As Long
3264     Dim lLine As Long
3265     Dim lLastLine As Long
3266     Dim lFirstCol As Long
3267     Dim lLastCol As Long
3268     Dim sCode As String

         If Not Parent.HostedByVB Then Exit Function

3269     On Error Resume Next
3270     With Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane
3271         .GetSelection lLine, lFirstCol, lLastLine, lLastCol
3272         DetermineLastLineInSelection = lLastLine
3273     End With
End Function

Public Function DetermineFirstLineInSelection() As Long
3274     Dim lLine As Long
3275     Dim lLastLine As Long
3276     Dim lFirstCol As Long
3277     Dim lLastCol As Long
3278     Dim sCode As String

         If Not Parent.HostedByVB Then Exit Function

3279     On Error Resume Next
3280     With Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane
3281         .GetSelection lLine, lFirstCol, lLastLine, lLastCol
3282         DetermineFirstLineInSelection = lLine
3283     End With
End Function

Public Function DetermineFirstColumnInSelection() As Long
3284     Dim lLine As Long
3285     Dim lLastLine As Long
3286     Dim lFirstCol As Long
3287     Dim lLastCol As Long
3288     Dim sCode As String

         If Not Parent.HostedByVB Then Exit Function

3289     On Error Resume Next
3290     With Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane
3291         .GetSelection lLine, lFirstCol, lLastLine, lLastCol
3292         DetermineFirstColumnInSelection = lFirstCol
3293     End With
End Function

Public Function DetermineLastColumnInSelection() As Long
3294     Dim lLine      As Long
3295     Dim lLastLine  As Long
3296     Dim lFirstCol  As Long
3297     Dim lLastCol   As Long
3298     Dim sCode      As String

         If Not Parent.HostedByVB Then Exit Function

3299     On Error Resume Next
3300     With Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane
3301         .GetSelection lLine, lFirstCol, lLastLine, lLastCol
3302         DetermineLastColumnInSelection = lLastCol
3303     End With
End Function

' ================================================================================
' Name              frmMain_mnuFileNew_Click
'
' Parameters
'      None
'
' Description
'
' Inserts a new template record.
'
' ================================================================================
Public Sub mnuFileNew_Click()
3304     NewTemplate
End Sub

' ================================================================================
' Name              frmMain_mnuFileRefresh_Click
'
' Parameters
'      None
'
' Description
'
' Refreshes the list of templates
'
' ================================================================================
Public Sub mnuFileRefresh_Click()
3305     On Error Resume Next
3306     Dim sTitle As String

3307     sTitle = lsbJumpTo.BarKey & gsCategoryTemplateDelimiter & lsbJumpTo.BarItemName
3308     RefillList
3309     JumpTo sTitle, False, True
End Sub


Public Sub mnuInsertTemplate_Click()
3310     DoInsertion Nothing, txtName
End Sub

Public Sub Form_Terminate()
3311     Dim Cancel As Integer
3312     Form_Unload Cancel
    ' LogEvent "frmMain: Terminate"
End Sub

Public Sub Form_Load()
3313     m_asaHistory.Clear
3314     LoadFormPosition Me
3315     SetColors GetSetting$(App.ProductName, "Last", "Background Color", "&H80000018&"), GetSetting$(App.ProductName, "Last", "Foreground Color", "&H80000008&")
End Sub

Public Sub Form_Unload(Cancel As Integer)
3316     If Not mHotKeyOpenWindow Is Nothing Then
3317         mHotKeyOpenWindow.Clear
3318         Set mHotKeyOpenWindow = Nothing
3319     End If

3320     ShutdownDLLs
3321     Set CurrentTemplate = Nothing
3322     SaveFormPosition Me
End Sub

Private Sub mHotKeyOpenWindow_HotKeyPress(ByVal sName As String, ByVal eModifiers As EHKModifiers, ByVal eKey As KeyCodeConstants)
3323     Dim sKey      As String
3324     Dim sRegValue As String

3325     If sName = "Sandy Cancel Insertion" Then
3326         gbCancelInsertion = True
3327     ElseIf sName = "Sandy Activate" Then
3328         mHotKeyOpenWindow.RestoreAndActivate Me.hwnd
3329     ElseIf sName = "Sandy Repeat Insertion" Then
3330         If Not InternalCurrentTemplate Is Nothing Then
3331             sKey = InternalCurrentTemplate.Key
3332         End If
3333     ElseIf sName = "Sandy Favorites" Then
3334         FavoriteCalledFromIDE = True
3335         ShowFavMenu
3336     ElseIf sName = "Sandy Externals" Then
3337         ShowExternalsMenu
3338     ElseIf Left$(sName, 18) = "Sandy Fast Insert " Then
3339         sKey = Mid$(sName, 19)
3340         sRegValue = GetSetting$(gsSliceAndDice, "Fast Insert", sKey)
3341         If Len(sRegValue) = 0 Then
3342             ' MsgBox "No " & gsSliceAndDice & gsS & gsTemplate & " has been associated with CTRL-SHIFT-" & sKey, vbOKOnly, "NO " & gsTemplate & " TO INSERT"
                 If bUserSure("No " & gsSliceAndDice & gsS & gsTemplate & " has been associated with CTRL-SHIFT-" & sKey & vbNewLine & vbTab & "Would you like to select one now ?", "NO " & gsTemplate & " TO INSERT") Then
                    sRegValue = InputBox$("Template to run when CTRL-SHIFT-" & sKey & " is pressed ?", "TEMPLATE TO ASSOCIATE", txtName)
                    If Len(sRegValue) > 0 Then
                       SaveSetting gsSliceAndDice, "Fast Insert", sKey, sRegValue
                    End If
                    'sKey = sRegValue
                 'Else
                 End If
3343             sKey = vbNullString
3344         Else
3345             sKey = sRegValue
3346         End If
3347     End If

3348     If Len(sKey) Then
3349         QueueAction "DoInsertion", sKey
3350         OkayToDoAction = True
3351     End If
End Sub

Private Sub txtCodeToFile_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
    If mbScramFormKey Then KeyCode = 0: Shift = 0
End Sub

Private Sub txtFilename_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
    If mbScramFormKey Then KeyCode = 0: Shift = 0
End Sub


Private Sub txtShortName_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
    If mbScramFormKey Then KeyCode = 0: Shift = 0
End Sub

Public Sub UpdateRecentFileList(Optional ByVal sFileToAdd As String)
    If Len(sFileToAdd) > 0 Then
       If mnuFileList.Count = 1 Then
          Load mnuFileList(mnuFileList.Count)
          With mnuFileList(mnuFileList.UBound)
               .Caption = sFileToAdd
          End With
       ElseIf mnuFileList.Count < 5 Then
       Else
       End If
    End If

End Sub
