VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{E60B3BB8-E409-11D2-BA4F-0080C8C222EC}#15.1#0"; "FirmSolutions.ocx"
Begin VB.Form frmMain 
   Caption         =   "Slice and Dice"
   ClientHeight    =   7920
   ClientLeft      =   2085
   ClientTop       =   3495
   ClientWidth     =   11115
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmOptions 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3705
      Left            =   4380
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame Frame3 
         Caption         =   " Press this Hot Key to Instantly Insert "
         Height          =   765
         Left            =   1890
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   3135
      End
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
      Height          =   7920
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   13970
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
         Width           =   2610
      End
      Begin VB.ListBox lstSoftVariables 
         BackColor       =   &H80000018&
         Height          =   1035
         Left            =   195
         TabIndex        =   3
         Top             =   840
         Width           =   1830
      End
      Begin VB.ListBox lstSoftCommands 
         BackColor       =   &H80000018&
         Height          =   1035
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
      Height          =   5445
      Left            =   3540
      TabIndex        =   19
      Top             =   420
      Width           =   7515
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
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
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
         Caption         =   "&New category"
         Index           =   0
      End
      Begin VB.Menu mnuCategoriesNewMethod 
         Caption         =   "&Duplicate a category. Template names and code"
         Index           =   1
      End
      Begin VB.Menu mnuCategoriesNewMethod 
         Caption         =   "Duplicate a category. Template names only"
         Index           =   2
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRefresh 
         Caption         =   "Refresh Category and Template List"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCategoriesDeleteCurrent 
         Caption         =   "Delete current category"
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
         Caption         =   "&Back"
      End
      Begin VB.Menu mnuForward 
         Caption         =   "&Forward"
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
Option Explicit

Private m_sTemplateDatabaseName         As String
Private m_sCurrentEventResponseCategory As String
'Private m_oDBClassGen                   As SandySupport.ISandyWindowGen
Private m_asaHistory                    As New CAssocArray
Private m_asaAttributes                 As New CAssocArray
Private SadCommands()                   As SandySupport.ISadAddin
Private FavoriteCount                   As Long
Private ExternalCount                   As Long
Private CurrentHistoryEntry             As String
Private ActionToDo                      As String
Private ActionParam                     As String
Private mbFillingAddInScreen            As Boolean
Private mbIgnoreBlanks                  As Boolean
Private mbIgnoreReadOnly                As Boolean

Public CurrentCodeArea                  As Integer
Public Parent                           As SandySupport.ISandyWizard
Public SliceAndDice                     As SandySupport.CSliceAndDice
Public CurrentTemplate                  As CTemplate
Public InternalCurrentTemplate          As CTemplate
Public Complete                         As CSadCommands
Public SadCommandSetCount               As Long
Public OkayToDoAction                   As Boolean
Public FavoriteCalledFromIDE            As Boolean
Public OkayToUnload                     As Boolean

Public WithEvents mHotKeyOpenWindow     As SandySupport.cRegHotKey
Attribute mHotKeyOpenWindow.VB_VarHelpID = -1

Private Const vbext_ct_StdModule = 1
Private Const vbext_ct_ClassModule = 2
Private Const vbext_ct_VBForm = 5
Private Const vbext_ct_UserControl = 8
Private Const vbext_ct_VBMDIForm = 6

Private Const vbext_pk_Proc = 0
Private Const vbext_pk_Let = 1
Private Const vbext_pk_Set = 2
Private Const vbext_pk_Get = 3

Private Const vbext_pt_StandardExe = 0
Private Const vbext_pt_ActiveXExe = 1
Private Const vbext_pt_ActiveXDll = 2
Private Const vbext_pt_ActiveXControl = 3

Implements SandySupport.ISandyWindowMain

Public Function ISandyWindowMain_AddSadCommandSet(ByRef oCommands As ISadAddin) As Boolean
On Error Resume Next
    Dim Externals       As SandySupport.CAssocArray
    Dim CurrExternal    As SandySupport.CAssocItem

    Err.Clear
        SadCommandSetCount = SadCommandSetCount + 1
        ReDim Preserve SadCommands(1 To SadCommandSetCount)
        Set SadCommands(SadCommandSetCount) = oCommands
        If SadCommands(SadCommandSetCount).Startup(Parent, Parent.SandyIDE) Then
           ISandyWindowMain_AddSadCommandSet = True
           frmSplash.lblDLLsLoaded(1).Caption = vbNullString & SadCommandSetCount
           frmSplash.lblDLLsLoaded(1).Refresh
           Set Externals = oCommands.Externals
               If Not Externals Is Nothing Then
                  If ExternalCount > 0 Then
                     Load mnuExternals(ExternalCount)
                     mnuExternals(ExternalCount).Caption = "-"
                     mnuExternals(ExternalCount).Tag = vbNullString
                     mnuExternals(ExternalCount).Enabled = True
                     mnuExternals(ExternalCount).Visible = True
                     ExternalCount = ExternalCount + 1
                  End If
                  For Each CurrExternal In Externals
                      If ExternalCount > 0 Then
                         Load mnuExternals(ExternalCount)
                      End If
                      mnuExternals(ExternalCount).Caption = CurrExternal.Key
                      mnuExternals(ExternalCount).Tag = SadCommandSetCount & "|" & CurrExternal.Value
                      mnuExternals(ExternalCount).Enabled = True
                      mnuExternals(ExternalCount).Visible = True
                      ExternalCount = ExternalCount + 1
                  Next CurrExternal
               End If
           Set Externals = Nothing
           DoEvents
        End If
    Err.Clear
End Function

Private Property Set ISandyWindowMain_Complete(ByVal RHS As SandySupport.CSadCommands)
    Set Complete = RHS
End Property

Private Property Get ISandyWindowMain_Complete() As SandySupport.CSadCommands
    Set ISandyWindowMain_Complete = Complete
End Property

Private Property Let ISandyWindowMain_CurrentCodeArea(ByVal RHS As Integer)
    CurrentCodeArea = RHS
End Property

Private Property Get ISandyWindowMain_CurrentCodeArea() As Integer
    ISandyWindowMain_CurrentCodeArea = CurrentCodeArea
End Property

Public Property Let ISandyWindowMain_CurrentEventResponseCategory(ByVal sNewCategory As String)
    m_sCurrentEventResponseCategory = sNewCategory
End Property

Public Property Get ISandyWindowMain_CurrentEventResponseCategory() As String
    If Len(m_sCurrentEventResponseCategory) = 0 Then m_sCurrentEventResponseCategory = "Event Response"
    ISandyWindowMain_CurrentEventResponseCategory = m_sCurrentEventResponseCategory
End Property

Private Property Set ISandyWindowMain_CurrentTemplate(ByVal RHS As SandySupport.CTemplate)
    Set CurrentTemplate = RHS
End Property

Private Property Get ISandyWindowMain_CurrentTemplate() As SandySupport.CTemplate
    Set ISandyWindowMain_CurrentTemplate = CurrentTemplate
End Property

Public Property Get ISandyWindowMain_CurrentTemplateNameAndCategory() As String
    ISandyWindowMain_CurrentTemplateNameAndCategory = txtName.Text
End Property

'Private Property Set ISandyWindowMain_DBClassGen(ByVal RHS As SandySupport.ISandyWindowGen)
'On Error Resume Next
'    Set m_oDBClassGen = RHS
'    SetColors GetSetting("SliceAndDice", "Last", "Background Color", "&H80000018&"), GetSetting("SliceAndDice", "Last", "Foreground Color", "&H80000008&")
'End Property
'
'Public Property Get ISandyWindowMain_DBClassGen() As SandySupport.ISandyWindowGen
'    Set ISandyWindowMain_DBClassGen = m_oDBClassGen
'End Property
'
Public Sub ISandyWindowMain_DeleteTemplate(Optional ByVal bAutoDelete As Boolean = False)
On Error GoTo EH_frmMain_DeleteTemplate
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Sub
    bInHereAlready = True

    If CurrentTemplate Is Nothing Then
       If bAutoDelete Then
          MsgBox "DeleteTemplate failed because nothing is selected."
       Else
          MsgBox "Please select a template to delete first."
       End If
       bInHereAlready = False
       Exit Sub
    ElseIf chkUndeletable.Value <> 0 Then
       MsgBox "This template cannot be deleted (undeletable turned on). Turn off before continuing."
       bInHereAlready = False
       Exit Sub
    End If

    If Not bAutoDelete Then
       If Not bUserSure("This will permanently delete the template """ & CurrentTemplate.Key & """." & gs2EOLTab & "Are you sure this is what you want to do ?") Then
          bInHereAlready = False
          Exit Sub
       End If
    End If

    CurrentTemplate.Deleted = True
    CurrentTemplate.Modified = True

    ISandyWindowMain_SaveTemplate
    ISandyWindowMain_RefillList
    
'    ISandyWindowMain_JumpTo SliceAndDice(1).Templates(1).Key
'    lsbJumpTo.BarAndItem SliceAndDice(1).Key, SliceAndDice(1).Templates(1).ShortTemplateName
    If Not SliceAndDice(1) Is Nothing Then
       If Not SliceAndDice(1).Templates(1) Is Nothing Then
          ISandyWindowMain_JumpTo SliceAndDice(1).Templates(1).Key
          lsbJumpTo.BarAndItem SliceAndDice(1).Key, SliceAndDice(1).Templates(1).ShortTemplateName
       ElseIf Not SliceAndDice(2) Is Nothing Then
          If Not SliceAndDice(2).Templates(1) Is Nothing Then
             ISandyWindowMain_JumpTo SliceAndDice(2).Templates(1).Key
             lsbJumpTo.BarAndItem SliceAndDice(2).Key, SliceAndDice(2).Templates(1).ShortTemplateName
          End If
       End If
    End If

EH_frmMain_DeleteTemplate_Continue:
    bInHereAlready = False
    Exit Sub

EH_frmMain_DeleteTemplate:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: DeleteTemplate" & vbCr & vbCr & Err.Description
    
    Resume EH_frmMain_DeleteTemplate_Continue

    Resume
End Sub

Public Property Get ISandyWindowMain_ExitAfterInsert() As Boolean
    ISandyWindowMain_ExitAfterInsert = mnuExitAfterInsert.Checked
End Property

Private Property Let ISandyWindowMain_FavoriteCalledFromIDE(ByVal RHS As Boolean)
    FavoriteCalledFromIDE = RHS
End Property

Private Property Get ISandyWindowMain_FavoriteCalledFromIDE() As Boolean
    ISandyWindowMain_FavoriteCalledFromIDE = FavoriteCalledFromIDE
End Property

Public Sub ISandyWindowMain_FillAddInScreen()
On Error GoTo EH_frmMain_FillAddInScreen
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Sub
    bInHereAlready = True
    mbFillingAddInScreen = True
    With CurrentTemplate
         txtName = .Key
         txtShortName = .ShortTemplateName

         txtCode(0) = .memoCodeAtTop
         txtCode(1) = .memoCodeAtCursor
         txtCode(2) = .memoCodeAtBottom

         txtFilename = .FileName
         txtCodeToFile = .memoCodeToFile
         
         chkUndeletable = Abs(.Undeletable)
         chkLocked = Abs(.Locked)
         chkFavorite = Abs(.Favorite)
         chkSelected = Abs(.Selected)
         lblRevision.Caption = "Revision #: " & .Revision
         lblAlpha.Caption = "Alpha Date: " & Format(.DateCreated, "Mmmm D, YYYY H:NN:SS AM/PM")
         lblDelta.Caption = "Delta Date: " & Format(.DateModified, "Mmmm D, YYYY H:NN:SS AM/PM")
'         With SliceAndDice.SystemInfo("Hotkey Templates").Item(.Key)
'              If Len(.Value) Then
'                 hkyInstantInsert.HotKeyModifier = Val(sGetToken(.Value, 2, ","))
'                 hkyInstantInsert.HotKey = Val(sGetToken(.Value, 1, ","))
'              Else
'                 hkyInstantInsert.HotKeyModifier = HOTKEYF_EXT
'                 hkyInstantInsert.HotKey = 0
'              End If
'         End With
    End With

EH_frmMain_FillAddInScreen_Continue:
    bInHereAlready = False
    mbFillingAddInScreen = False
    Exit Sub

EH_frmMain_FillAddInScreen:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: FillAddInScreen" & vbCr & vbCr & Err.Description
    
    Resume EH_frmMain_FillAddInScreen_Continue

    Resume
End Sub

Private Sub ISandyWindowMain_FormUnload()
    Dim Cancel As Integer
    Form_Unload Cancel
End Sub

Public Sub ISandyWindowMain_GetCategoryAndName(ByVal sCategoryAndName As String, ByRef sCategory As String, ByRef sShortName As String)
    If lTokenCount(sCategoryAndName, " - ") < 2 Then
       sCategory = "Unknown"
       sShortName = sCategoryAndName
    Else
       sCategory = sGetToken(sCategoryAndName, 1, " - ")
       sShortName = sAfter(sCategoryAndName, 1, " - ")
       If Len(sShortName) = 0 Then
          sCategory = "Unknown"
          sShortName = sCategoryAndName
       End If
    End If
End Sub

Private Sub ISandyWindowMain_Hide()
    Me.Hide
End Sub

'Public Sub HandleIDEEvents(ByVal sTemplateName As String, Optional ByVal VBProject As VBIDE.VBProject, Optional ByVal VBComponent As VBIDE.VBComponent)
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
'    DoInsertion m_asaIDEEvents, CurrentEventResponseCategory & " - " & sTemplateName
'
'    m_asaIDEEvents.Clear
'    Set m_asaIDEEvents = Nothing
'
'EH_frmMain_HandleIDEEvents_Continue:
'    bInHereAlready = False
'    Exit Sub
'
'EH_frmMain_HandleIDEEvents:
'    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: HandleIDEEvents" & vbCr & vbCr & Err.Description
'
'    Resume EH_frmMain_HandleIDEEvents_Continue
'
'    Resume
'End Sub

Public Sub ISandyWindowMain_HideAllWindows(Optional ByVal bUnloadAsWell = False)
On Error Resume Next
    Dim CurrSet As Long

    'If Not m_oDBClassGen Is Nothing Then
    '   m_oDBClassGen.Hide
    'End If

    If SadCommandSetCount > 0 Then
       If SadCommandSetCount = 1 Then
          SadCommands(1).CommandSet.HideWindow bUnloadAsWell
          Exit Sub
       Else
          For CurrSet = 1 To SadCommandSetCount
              SadCommands(CurrSet).CommandSet.HideWindow bUnloadAsWell
              SadCommands(CurrSet).ExecuteExternal "HIDE ALL WINDOWS", "HIDE ALL WINDOWS"
          Next CurrSet
       End If
    End If
End Sub

Private Property Let ISandyWindowMain_hWnd(ByVal RHS As Long)
    ' Can't do this
End Property

Private Property Get ISandyWindowMain_hWnd() As Long
    ISandyWindowMain_hWnd = Me.hWnd
End Property

Public Function ISandyWindowMain_InitializeAddinDLLs(ByVal sAddinList As String) As Boolean
    Dim asaTemp As SandySupport.CAssocArray
    Dim CurrAssocItem As SandySupport.CAssocItem
    Dim CurrDLL As ISadAddin

    ISandyWindowMain_ShutdownDLLs
    
    If Len(sAddinList) = 0 Then
       ISandyWindowMain_InitializeAddinDLLs = True
       Exit Function
    End If

    Set asaTemp = CreateObject("SandySupport.CAssocArray")
        asaTemp.All = sAddinList
        For Each CurrAssocItem In asaTemp
            If StrComp(UCase$(Trim$(CurrAssocItem.Value)), "LOAD") = 0 Then
On Error Resume Next
               Err.Clear
               Set CurrDLL = CreateObject(Trim$(CurrAssocItem.Key))
               If Err.Number = 0 Then
                  If Not ISandyWindowMain_AddSadCommandSet(CurrDLL) Then
                     CurrAssocItem.Value = "Error in 'AddSadCommandSet'"
                  Else
                     SadCommands(SadCommandSetCount).CommandSet.Attributes("Name").Value = Trim$(CurrAssocItem.Key)
                     If UCase$(Trim$(CurrAssocItem.Key)) = "SADREGISTER.NEWCOMMANDS" Then
                        SadCommands(SadCommandSetCount).CommandSet.Attributes("Registered").Value = IIf(frmSplash.DetermineRegistration, "True", "False")
                     End If
                  End If
               Else
                  MsgBox "Failed to create the SAD Addin object: " & vbCr & vbTab & "Name:" & Trim$(CurrAssocItem.Key) & vbCr & vbTab & "Err #" & Err.Number & ": " & Err.Description
               End If
               Err.Clear
            End If
        Next CurrAssocItem

      ' Future: Store results of loads back for next time.
        sAddinList = asaTemp.All

    Set asaTemp = Nothing
End Function

Private Property Set ISandyWindowMain_InternalCurrentTemplate(ByVal RHS As SandySupport.CTemplate)
    Set InternalCurrentTemplate = RHS
End Property

Private Property Get ISandyWindowMain_InternalCurrentTemplate() As SandySupport.CTemplate
    Set ISandyWindowMain_InternalCurrentTemplate = InternalCurrentTemplate
End Property

Private Property Set ISandyWindowMain_mHotKeyOpenWindow(ByVal RHS As SandySupport.cRegHotKey)
    Set mHotKeyOpenWindow = RHS
End Property

Private Property Get ISandyWindowMain_mHotKeyOpenWindow() As SandySupport.cRegHotKey
    Set ISandyWindowMain_mHotKeyOpenWindow = mHotKeyOpenWindow
End Property

Public Sub ISandyWindowMain_NewTemplate(Optional ByVal bAutoCreate As Boolean = False, Optional ByVal sTitle As String, Optional ByVal sDefaultShortName As String, Optional ByVal bJumpToAfterCreate As Boolean = True)
On Error GoTo EH_frmMain_NewTemplate
    Dim sCategory As String
    Dim sShortName As String

    If Len(sTitle) = 0 Then
       sCategory = lsbJumpTo.BarKey
       If Len(sDefaultShortName) = 0 Then
          sDefaultShortName = Abs(NextNegativeUnique())
       End If
       sTitle = InputBox("What should the name of this template be ?" & gsEolTab & "(Blank to cancel)" & gs2EOL & "Format of name MUST be:" & gsEolTab & "Category Name - Template Name", "NEW TEMPLATE", sCategory & " - " & sDefaultShortName)
    End If
    If Len(sTitle) = 0 Then Exit Sub

    ISandyWindowMain_GetCategoryAndName sTitle, sCategory, sShortName
    If Len(sCategory) = 0 Or Len(sShortName) = 0 Then
       MsgBox "New Template name must be in the form: " & vbCr & vbTab & "<CategoryName> & ' - ' & <ShortTemplateName>"
       Exit Sub
    End If

    If SliceAndDice(sCategory) Is Nothing Then
       If Not bAutoCreate Then
          If Not bUserSure("The category '" & sCategory & "' does not exist. Would you like to create it ?") Then
             Exit Sub
          End If
       End If
       SliceAndDice.Categorys.Add sCategory
    ElseIf Not (SliceAndDice(sCategory).Templates(sShortName) Is Nothing) Then
       MsgBox "There is a template by that name in that category already.", vbInformation
       Exit Sub
    End If

    ISandyWindowMain_SaveTemplate

    With SliceAndDice(sCategory).Templates.Add(sTitle)
         .ShortTemplateName = sShortName
         .ParentKey = sCategory
         .OriginalShortName = sShortName
    End With

    SliceAndDice.Save

    ISandyWindowMain_RefillList
    
    If bJumpToAfterCreate Then
       ISandyWindowMain_JumpTo sTitle, False, True
    
       txtName.Text = sTitle
       txtShortName.Text = sShortName
    End If

EH_frmMain_NewTemplate_Continue:
    Exit Sub

EH_frmMain_NewTemplate:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: NewTemplate" & vbCr & vbCr & Err.Description
    
    Resume EH_frmMain_NewTemplate_Continue

    Resume
End Sub

Private Property Let ISandyWindowMain_OkayToDoAction(ByVal RHS As Boolean)
    OkayToDoAction = RHS
End Property

Private Property Get ISandyWindowMain_OkayToDoAction() As Boolean
    ISandyWindowMain_OkayToDoAction = OkayToDoAction
End Property

Private Property Let ISandyWindowMain_OkayToUnload(ByVal RHS As Boolean)
    OkayToUnload = RHS
End Property

Private Property Get ISandyWindowMain_OkayToUnload() As Boolean
    ISandyWindowMain_OkayToUnload = OkayToUnload
End Property

Private Property Set ISandyWindowMain_Parent(ByVal RHS As SandySupport.ISandyWizard)
    Set Parent = RHS
End Property

Private Property Get ISandyWindowMain_Parent() As SandySupport.ISandyWizard
    Set ISandyWindowMain_Parent = Parent
End Property

Public Sub ISandyWindowMain_QueueAction(ByVal sAction As String, Optional ByVal sParam As String, Optional ByVal Interval As Integer = 150)
    OkayToDoAction = False
    ActionToDo = sAction
    ActionParam = sParam
    tmrDoAction.Interval = IIf(Interval > 65535, 65535, IIf(Interval < 100, 100, Interval))
    tmrDoAction.Enabled = True
End Sub

Public Property Let ISandyWindowMain_QueuedInsertions(New_QueuedInsertions As String)
On Error GoTo EH_frmMain_QueuedInsertions
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Property
    bInHereAlready = True

    Dim asaVar As New CAssocArray
    Dim asaV As New CAssocArray
    Dim CurItem As SandySupport.CAssocItem

    asaVar.ItemDelimiter = "~"
    asaVar.All = New_QueuedInsertions
    For Each CurItem In asaVar.mCol
        ISandyWindowMain_DoInsertion asaV, CurItem.Key
        If CancelInsertion Then Exit Property
    Next CurItem

EH_frmMain_QueuedInsertions_Continue:
    bInHereAlready = False
    Exit Property

EH_frmMain_QueuedInsertions:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: QueuedInsertions" & vbCr & vbCr & Err.Description
    
    Resume EH_frmMain_QueuedInsertions_Continue

    Resume
End Property

Public Function ISandyWindowMain_RefreshDatabaseConnection() As Boolean
On Error GoTo EH_frmMain_RefreshDatabaseConnection

    Call NextNegativeUnique
    
    Set CurrentTemplate = Nothing
    Set InternalCurrentTemplate = Nothing
    Set SliceAndDice = Nothing
    
    Set SliceAndDice = CreateObject("SandySupport.CSliceAndDice")
    If Not SliceAndDice.Load(m_sTemplateDatabaseName) Then
       ISandyWindowMain_RefreshDatabaseConnection = False
       Exit Function
    End If

    ISandyWindowMain_RefillList
On Error Resume Next
    lsbJumpTo.HideCategories
    lsbJumpTo.DisplayCategories

    Caption = "Sandy " & App.Major & "." & App.Minor & "." & App.Revision & " - " & m_sTemplateDatabaseName
    ISandyWindowMain_RefreshDatabaseConnection = True

EH_frmMain_RefreshDatabaseConnection_Continue:
    Exit Function

EH_frmMain_RefreshDatabaseConnection:
    MsgBox "Error during RefreshDatabaseConnection." & gsEolTab & Err.Description
    ISandyWindowMain_RefreshDatabaseConnection = False
    Resume EH_frmMain_RefreshDatabaseConnection_Continue
End Function

Public Sub ISandyWindowMain_DoInsertion(asaV As SandySupport.CAssocArray, sTemplateToInsert As String, Optional ByVal bSkipDeclarations As Boolean = False)
On Error GoTo EH_frmMain_DoInsertion
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Sub
    bInHereAlready = True
    lsbJumpTo.Enabled = False

    CancelInsertion = False
    mbIgnoreBlanks = False

    Dim lLine As Long
    Dim lTemp As Long
    Dim sCodeToInsert As String
    Dim sProcName As String
    Dim lProcType As Long
    Dim sProcTypeLong As String

    Dim asaVar As SandySupport.CAssocArray                                                                ' Associative Array used when filling in values to a code template when being inserted
    Dim CurItem As SandySupport.CAssocItem

   'If txtName <> sTemplateToInsert Then
       If Not ISandyWindowMain_SetInternalCurrentTemplate(sTemplateToInsert) Then
          MsgBox "Can't find the template '" & sTemplateToInsert & "' to insert." & gsEolTab & "Aborting this insertion.", , "DoInsertion Error"
          GoTo EH_frmMain_DoInsertion_Continue
       End If
   'End If

    ' Begin Log
      'frmLog.tvwLog.Nodes.Add , , (frmLog.tvwLog.Nodes.Count + 1) & " Inserting " & sTemplateToInsert, "Inserting " & sTemplateToInsert
      'If Not asaV Is Nothing Then
      '   For Each CurItem In asaV
      '       frmLog.tvwLog.Nodes.Add (frmLog.tvwLog.Nodes.Count + 1) & " Inserting " & sTemplateToInsert, tvwChild, , CurItem.Key & " = " & CurItem.Value
      '   Next CurItem
      'End If
    ' End Log

    If Parent.SandyIDE.ActiveCodePane Is Nothing Then
       If Parent.SandyIDE.SelectedComponent Is Nothing Then
          If Parent.SandyIDE.Components.Count > 0 Then
             Parent.SandyIDE.Components(1).CodeModule.CodePane.Show
          Else
             Parent.SandyIDE.AddComponent(vbext_ct_StdModule).CodeModule.CodePane.Show
          End If
       ElseIf Not Parent.SandyIDE.SelectedComponent.CodeModule.CodePane Is Nothing Then
          Parent.SandyIDE.SelectedComponent.CodeModule.CodePane.Show
       Else
          MsgBox "Can't do an insertion since no code pane is active.", vbInformation
          GoTo EH_frmMain_DoInsertion_Continue
       End If
    End If

    If asaV Is Nothing Then
       Set asaVar = CreateObject("SandySupport.CAssocArray")
    Else
       Set asaVar = asaV                                                                        ' Use the supplied assiciative array
    End If

        With Parent.SandyIDE.ActiveCodePane                                                                 ' Send all output to the active code pane
             .GetSelection lLine, lTemp, lTemp, lTemp                                           ' Determine where the cursor is
             asaVar.Add "Project Name", Parent.SandyIDE.ActiveProject.Name                                 ' Add the build in soft variables
             asaVar.Add "Module Name", .CodeModule.Parent.Name
             ISandyWindowMain_GetProcAtLine lLine, sProcName, lProcType
             If sProcName <> vbNullString Then
                asaVar.Add "Proc Name", sProcName
                asaVar.Add "Proc Type", Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
                sProcTypeLong = .CodeModule.Lines(.CodeModule.ProcBodyLine(sProcName, lProcType), 1)
                If InStr(sProcTypeLong, "Function") > 0 Then
                   sProcTypeLong = "Function"
                ElseIf InStr(sProcTypeLong, "Property") > 0 Then
                   sProcTypeLong = "Property"
                Else
                   sProcTypeLong = "Sub"
                End If
                asaVar.Add "Proc Type Long", sProcTypeLong
                sProcName = vbNullString
                sProcTypeLong = vbNullString
             End If
             If Parent.InsertTemplate(InternalCurrentTemplate.memoCodeAtCursor, lLine, asaVar, sTemplateToInsert) Then
                If Not bSkipDeclarations Then
                   If Parent.InsertTemplate(InternalCurrentTemplate.memoCodeAtTop, .CodeModule.CountOfDeclarationLines + 1, asaVar, sTemplateToInsert) Then
                      If Parent.InsertTemplate(InternalCurrentTemplate.memoCodeAtBottom, .CodeModule.CountOfLines + 1, asaVar, sTemplateToInsert) Then
                         Call Parent.InsertTemplate(InternalCurrentTemplate.memoCodeToFile, 1, asaVar, sTemplateToInsert, txtFilename)
                      End If
                   End If
                Else
                   If Parent.InsertTemplate(InternalCurrentTemplate.memoCodeAtBottom, .CodeModule.CountOfLines + 1, asaVar, sTemplateToInsert) Then
                      Call Parent.InsertTemplate(InternalCurrentTemplate.memoCodeToFile, 1, asaVar, sTemplateToInsert, txtFilename)
                   End If
                End If
             End If
        End With
    Set asaVar = Nothing                                                                        ' Destroy the associative array

    sCodeToInsert = vbNullString

    If mnuExitAfterInsert.Checked = True Then
       mnuFileExit_Click
    End If


EH_frmMain_DoInsertion_Continue:
    bInHereAlready = False
    lsbJumpTo.Enabled = True
    Exit Sub

EH_frmMain_DoInsertion:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: DoInsertion" & vbCr & vbCr & Err.Description
    
    CancelInsertion = bUserSure("Cancel processing ?")
    Resume EH_frmMain_DoInsertion_Continue

    Resume
End Sub

' ********************************************************************************
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
' ********************************************************************************
Public Function ISandyWindowMain_FillTemplateWithUserInput(ByRef asaX As SandySupport.CAssocArray, ByVal sToParse As String, ByRef sCodeToInsert As String, ByVal sMsgBoxTitle As String) As Boolean
    Static sVarName As String
    Static sVarPhrase As String
    Static sDefault As String
    Static sT As String
    Static sVar1 As String
    Static sVar2 As String
    Static sVar3 As String
    Static sNow As String
    Static lParamCount As Long
    Static CurrSet As Long
    Static bInlineCommandExecuted As Boolean

    Do While InStr(sGetToken(sToParse, 1, gsEOL), gsSoftVarDelimiter) > 0                                             ' For each soft variable found
       sVarPhrase = sGetToken(sToParse, 2, gsSoftVarDelimiter)                                     ' Get the Variable name and default if provided
       sVarName = sGetToken(sVarPhrase, 1, "::")                                     ' Extract just the variable name
       sNow = vbNullString
       bInlineCommandExecuted = False
       If SadCommandSetCount > 0 Then
          sVar1 = sAfter(sVarPhrase, 1, "::")
          For CurrSet = 1 To SadCommandSetCount
              If SadCommands(CurrSet).ExecuteSoftCommandInline(asaX, UCase$(sVarName), sVar1, sNow) Then
                 bInlineCommandExecuted = True
                 Exit For
              End If
          Next CurrSet
       End If
       If bInlineCommandExecuted Then
          sT = sToParse
          sToParse = sBefore(sT, 2, gsSoftVarDelimiter) & sNow & sAfter(sT, 2, gsSoftVarDelimiter)
       Else
                      With asaX.Item(sVarName)                                                 ' With the Association for the Soft variable
                           If Len(.Value) = 0 Then                                             ' If there is currently no value
                              If mbIgnoreBlanks Then
                              ElseIf InStr(sVarPhrase, "::") Then                              ' Use default provided
                                 sDefault = sGetToken(sVarPhrase, 2, "::")                     ' Extract the default
                                 If Left$(sDefault, 1) = "@" Then                               ' See if the default is to be drawn from another Association's value
                                    sDefault = asaX.Item(Mid$(sDefault, 2)).Value               ' Lookup another value in the array as the default
                                 End If
                                 .Value = InputBox(sVarName, sMsgBoxTitle, sDefault)       ' Ask the user to enter a value and then set the Association's value to it
                                 If Len(.Value) = 0 Then CancelInsertion = bUserSure("Cancel processing ?")
                              Else                                                             ' No default
                                 .Value = InputBox(sVarName, sMsgBoxTitle)                ' Ask the user to enter a value and then set the Association's value to it
                                 If Len(.Value) = 0 Then CancelInsertion = bUserSure("Cancel processing ?")
                              End If
                           End If                                                              ' At this point the Association's value is set one way or the other
                           If CancelInsertion Then
                              sVarName = vbNullString
                              sVarPhrase = vbNullString
                              sDefault = vbNullString
                              sT = vbNullString
                              ISandyWindowMain_FillTemplateWithUserInput = False
                              Exit Function                                                    ' User canceled
                           End If
                           sT = sToParse                                                       ' Save the string so far into a temporary area
                           sToParse = sBefore(sT, 2, gsSoftVarDelimiter) & .Value & sAfter(sT, 2, gsSoftVarDelimiter)      ' Replace the Soft variable with the user's entry
                      End With
       End If
    Loop

    ISandyWindowMain_FillTemplateWithUserInput = True                                                 ' Returned the final parsed string

    If InStr(sToParse, gsEOL) Then
       If Right$(sToParse, 2) <> gsEOL Then                                          ' If the code to insert is more than a line long
          sToParse = sToParse & gsEOL                                                ' Insure it has an EOL at the end to be parsed properly
       End If
    End If

    sCodeToInsert = sToParse

    sVarName = vbNullString
    sVarPhrase = vbNullString
    sDefault = vbNullString
    sT = vbNullString
    sVar1 = vbNullString
    sVar2 = vbNullString
    sVar3 = vbNullString

End Function



' ********************************************************************************
' Name              frmMain_InternalInsertTemplate
'
' Parameters
'
' Description
'
' This actually causes the code indicated to get inserted correctly. Soft
' commands are handled here.
'
' ********************************************************************************
Public Function ISandyWindowMain_InternalInsertTemplate(II As CInsertionInfo) As Boolean
    Dim CurDesigner      As SandySupport.IDesigner
    Dim CurForm          As SandySupport.IForm
    Dim CurFrame         As SandySupport.IControl
    Dim CurControl       As SandySupport.IControl
    Dim CurReference     As SandySupport.IReference
    Dim CurModule        As SandySupport.ICodeModule
    Dim ControlVars      As SandySupport.CAssocArray

    Dim lCurControl      As Long
    Dim lCurFrame        As Long
    Dim fh               As Long
    Dim CurrSet          As Long
    Dim CurrParam        As Long
    Dim lParamCount      As Long
    Dim lStartLine       As Long
    Dim lEndLine         As Long
    Dim lStartColFound   As Long
    Dim lEndColFound     As Long
    Dim lProcType        As Long
    
    Dim lIfLoops         As Long
    Dim CodaIterations   As Long
    Dim NextElse         As Long
    Dim NextElseIf       As Long
    Dim NextEndIf        As Long
    Dim IndentationLevel As Long

    Dim bFunction        As Boolean
    Dim bFoundReference  As Boolean
    Dim bDoCoda          As Boolean

    Dim sT               As String
    Dim sHold1           As String
    Dim sProcName        As String
    Dim sProcType        As String
    Dim sHold2           As String
    Dim sCurParam        As String
    Dim sCurType         As String
    Dim CommandReference As String
    Dim sDelim1          As String
    Dim sDelim2          As String

On Error GoTo EH_InsertTemplate
    If II Is Nothing Then
       ISandyWindowMain_InternalInsertTemplate = True
       GoTo EH_InsertTemplate_Continue
    End If

   'Set II = II
    Set CurModule = Parent.SandyIDE.ActiveCodePane.CodeModule
    II.LinesLeftToProcess = II.OriginalCodeToInsert

On Error Resume Next
    II.SoftVars("FIRSTCOLUMN").Value = ISandyWindowMain_DetermineFirstColumnInSelection
    II.SoftVars("LASTCOLUMN").Value = ISandyWindowMain_DetermineLastColumnInSelection
    II.SoftVars("FIRSTLINE").Value = ISandyWindowMain_DetermineFirstLineInSelection
    II.SoftVars("LASTLINE").Value = ISandyWindowMain_DetermineLastLineInSelection
                             
On Error GoTo EH_InsertTemplate

CODA_RESTART:
    'With CurModule
         If Len(II.LinesLeftToProcess) > 100000 Then
            If Not bUserSure("Template to insert '" & II.TemplateName & "' has become is very large" & vbCr & vbTab & "(" & Len(II.LinesLeftToProcess) & " bytes, started at " & Len(II.OriginalCodeToInsert) & " bytes)." & vbCr & vbTab & "Continue inserting anyway ?") Then
               II.LinesLeftToProcess = vbNullString
               sT = vbNullString
               CancelInsertion = True
               GoTo EH_InsertTemplate_Continue
            End If
         End If
         If InStr(II.LinesLeftToProcess, gsSoftCmdDelimiter) = 0 And InStr(II.LinesLeftToProcess, gsSoftVarDelimiter) = 0 Then
            If Len(II.ExternalFilename) = 0 Then
               If Len(II.LinesLeftToProcess) > 0 Then
                  CurModule.InsertLines II.PointOfInsertion, II.LinesLeftToProcess                 ' No embedded commands, no embedded variables. Simple insertion
               End If
            Else
               II.TextToSendToFile = II.LinesLeftToProcess
            End If
        'ElseIf InStr(II.LinesLeftToProcess, gsSoftCmdDelimiter) = 0 Then
        '   sT = vbNullString
        '   If FillTemplateWithUserInput(II.SoftVars, II.LinesLeftToProcess, sT, II.TemplateName) Then
        '      II.LinesLeftToProcess = sT
        '      If InStr(sGetToken(sT, 1, gsEOL), gsSoftCmdDelimiter) Then
        '       ' Inserting caused more lines to appear, push the extra lines into the buffer for later insertion
        '        'II.LinesLeftToProcess = sGetToken(II.LinesLeftToProcess, 1, gsEOL) & gsEOL & sAfter(sT, 1, gsEOL) & IIf(Right$(sT, 2) = gsEOL, vbNullString, gsEOL) & sAfter(II.LinesLeftToProcess, 1, gsEOL)
        '         GoTo CODA_RESTART
        '      End If
        '      If Len(II.ExternalFilename) = 0 Then
        '         CurModule.InsertLines II.PointOfInsertion, II.LinesLeftToProcess                 ' No embedded commands. Simple insertion
        '      Else
        '         II.TextToSendToFile = II.LinesLeftToProcess
        '      End If
        '   Else
        '      II.TextToSendToFile = vbNullString
        '      II.LinesLeftToProcess = sT
        '   End If
         Else
            Do Until Len(II.LinesLeftToProcess) = 0                         ' More complicated line by line with embedded commands (and/or variables) insertion
               If CancelInsertion Then GoTo EH_InsertTemplate_Continue
               
               DoEvents
               If Not ISandyWindowMain_FillTemplateWithUserInput(II.SoftVars, sGetToken(II.LinesLeftToProcess, 1, gsEOL), sT, II.TemplateName) Then
                  ISandyWindowMain_InternalInsertTemplate = False
                  GoTo EH_InsertTemplate_Continue
               End If
               If InStr(sT, gsEOL) Then ' Inserting caused more lines to appear, push the extra lines into the buffer for later insertion
                  II.LinesLeftToProcess = sGetToken(II.LinesLeftToProcess, 1, gsEOL) & gsEOL & sAfter(sT, 1, gsEOL) & IIf(Right$(sT, 2) = gsEOL, vbNullString, gsEOL) & sAfter(II.LinesLeftToProcess, 1, gsEOL)
                  sT = sGetToken(sT, 1, gsEOL)
               End If
               II.CurrentLineToProcess = sT
               If Left$(II.CurrentLineToProcess, 2) = gsSoftCmdDelimiter Then            ' Process an imbedded command
                  II.CurrentLineToProcess = sGetToken(II.CurrentLineToProcess, 2, gsSoftCmdDelimiter)   ' Get the command with parameter(s)
                  II.SoftCommandName = sGetToken(II.CurrentLineToProcess)     ' Get just the command string (Case insensitive)
                  II.sParam = sAfter(II.CurrentLineToProcess)                 ' Get just the parameters
                  II.AllParameters = Replace(Replace(II.sParam, "$SP$", " "), "$TAB$", vbTab)

                  If lTokenCount(II.sParam) = 1 Then
                     If Val(II.sParam) <> 0 Then                ' One parameter passed
                        sProcName = vbNullString                      ' Parameter is a number (line offset)
                        II.ParamLineOffset = Val(II.sParam)
                        II.sParam = vbNullString
                     Else
                        sProcName = II.sParam                  ' Parameter is a procedure name or a real parameter
                        II.ParamLineOffset = 0
                        II.sParam = vbNullString
                     End If
                  ElseIf lTokenCount(II.sParam) = 2 Then       ' Two parameters passed
                     sProcName = sGetToken(II.sParam)          ' Get the procedure name to work on
                     II.sParam = sAfter(II.sParam)                ' Strip out the procedure name
                     sProcType = UCase$(II.sParam)             ' Get the procedure type (first of two parameters)                  [ Default: PROC ]
                     If Val(sProcType) <> 0 Then
                        II.ParamLineOffset = Val(sProcType)
                        sProcType = vbNullString
                     Else
                        II.ParamLineOffset = 0
                     End If
                  Else                                      ' Three or more parameters passed
                     sProcName = sGetToken(II.sParam)          ' Get the procedure name to work on
                     II.sParam = sAfter(II.sParam)                ' Strip out the procedure name
                     II.ParamLineOffset = Val(sAfter(II.sParam))      ' Get the line offset (usually negative) (second of two parameters) [ Default: 0 ]
                     sProcType = UCase$(sGetToken(II.sParam))  ' Get the procedure type (first of two parameters)                  [ Default: PROC ]
                  End If
                                                            ' Determine which constant to use for the passed Procedure Type
                                                            ' Note: Nothing passed ? Assume a sub or function
                  lProcType = Switch(sProcType = "GET", vbext_pk_Get, sProcType = "SET", vbext_pk_Set, sProcType = "LET", vbext_pk_Let, sProcType = "PROC", vbext_pk_Proc, True, vbext_pk_Proc)
                                                            ' Execute the command specified
                  If InStr(II.AllParameters, "=") > 0 Then
                     II.Result = Trim$(sGetToken(II.AllParameters, 1, "="))
                     II.Expression = Trim$(sAfter(II.AllParameters, 1, "="))
                    'II.Expression = sAfter(II.AllParameters, 1, "=")
                  End If

                ' Determine indentation level of subcode
                  If StrComp(Left$(II.SoftCommandName, 1), "_") = 0 Then
                     CommandReference = "_" & sGetToken(II.SoftCommandName, 2, "_") & "_"
                     II.SoftCommandName = sAfter(II.SoftCommandName, 2, "_")
                     If Len(II.SoftCommandName) = 0 Then
                        If bUserSure("Slice and Dice has detected a dangling Command Reference in line:" & vbNewLine & vbNewLine & II.CurrentLineToProcess & vbNewLine & vbNewLine & vbTab & "Would you like to cancel insertion ?") Then
                           CancelInsertion = True
                        End If
                     End If
                  Else
                     CommandReference = vbNullString
                  End If

                  II.SoftCommandName = UCase$(II.SoftCommandName)

                  Select Case II.SoftCommandName
                         Case "BLOCK"
                               sProcName = II.SoftVars(sProcName)
                               If II.PointOfInsertion < 1 Then II.PointOfInsertion = 1
                               If Len(II.ExternalFilename) = 0 Then
                                  CurModule.InsertLines II.PointOfInsertion, sProcName
                               Else
                                  II.TextToSendToFile = II.TextToSendToFile & sProcName & gsEOL
                               End If
                               II.PointOfInsertion = II.PointOfInsertion + lTokenCount(sProcName, vbNewLine)
                         
                         Case "DEBUG"
                               II.SoftCommandName = II.SoftCommandName
                               
                         Case "ELSE"
                              If lIfLoops > 0 Then
                                 NextEndIf = InStr(UCase$(II.LinesLeftToProcess), "~~" & UCase$(CommandReference) & "ENDIF")
                                 If NextEndIf > 0 Then
                                    II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, NextEndIf)
                                 Else
                                    II.LinesLeftToProcess = vbNullString
                                 End If
                                 lIfLoops = lIfLoops - 1
                              End If

                         Case "IF"
                              If Val(Evaluate(II.AllParameters, II.SoftVars)) <> 0 Then
                                 lIfLoops = lIfLoops + 1
                              Else
                                 NextEndIf = InStr(UCase$(II.LinesLeftToProcess), "~~" & UCase$(CommandReference) & "ENDIF")
                                 NextElse = InStr(UCase$(II.LinesLeftToProcess), "~~" & UCase$(CommandReference) & "ELSE" & vbCr)
                                 If NextElse > 0 And NextElse < NextEndIf Then
                                    II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, NextElse)
                                 ElseIf NextEndIf <> 0 Then
                                    II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, NextEndIf)
                                 Else
                                    II.LinesLeftToProcess = vbNullString
                                 End If
                              End If
                         
'                         Case "ELSE"
'                              If lIfLoops > 0 Then
'                                 NextEndIf = InStr(UCase$(II.LinesLeftToProcess), "~~ENDIF")
'                                 If NextEndIf > 0 Then
'                                    II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, NextEndIf)
'                                 Else
'                                    II.LinesLeftToProcess = vbNullString
'                                 End If
'                                 lIfLoops = lIfLoops - 1
'                              End If
'
'                         Case "IF"
'                              If Val(II.AllParameters) <> 0 Then
'                                 lIfLoops = lIfLoops + 1
'                              Else
'                                 NextEndIf = InStr(UCase$(II.LinesLeftToProcess), "~~ENDIF")
'                                 NextElse = InStr(UCase$(II.LinesLeftToProcess), "~~ELSE" & vbCr)
'                                 If NextElse > 0 And NextElse < NextEndIf Then
'                                    II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, NextElse)
'                                 ElseIf NextEndIf <> 0 Then
'                                    II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, NextEndIf)
'                                 Else
'                                    II.LinesLeftToProcess = vbNullString
'                                 End If
'                              End If
'
                        'Case "ELSEIF"
                               
                         
                         Case "ENDIF"
                               If lIfLoops > 0 Then lIfLoops = lIfLoops - 1

                         Case "'", "STARTCODA", "ENDIF"
                         
                         
                         Case "ABORT", "ABORTINSERTION"
On Error Resume Next
                              If Val(II.AllParameters) <> 0 Then
                                 MsgBox "Insertion aborted by the ~~AbortInsertion command."
                                 II.LinesLeftToProcess = vbNullString
                                 CancelInsertion = True
                              End If
On Error GoTo EH_InsertTemplate

                         Case "CANCEL", "CANCELINSERTION"
On Error Resume Next
                              If Val(II.AllParameters) <> 0 Then
                                 II.LinesLeftToProcess = vbNullString
                              End If
On Error GoTo EH_InsertTemplate

'                         Case "INSERTTEMPLATE", "INSERT"
'                              DoInsertion asaX, II.AllParameters

' ************************************************
' Soft Commands that control the flow of insertion
' ************************************************
                         Case "CODA", "LOOPWHILE", "LOOPUNTIL"
                              CodaIterations = CodaIterations + 1
                              If CodaIterations > 10000 Then
                                 If bUserSure("Slice and Dice has found what appears to be an endless loop via ~~Coda, ~~LoopWhile, or ~~LoopUntil." & vbNewLine & vbTab & "Would you like to cancel processing ?") Then
                                    CancelInsertion = True
                                 Else
                                    CodaIterations = 0
                                 End If
                              End If
                              If II.SoftCommandName = "LOOPUNTIL" Then
                                 bDoCoda = (Val(sGetToken(II.AllParameters)) = 0)
                              Else
                                 bDoCoda = (Val(sGetToken(II.AllParameters)) <> 0)
                              End If
                              If bDoCoda Then
                                 Select Case II.SoftCommandName
                                        Case "CODA"
                                              If InStr(II.OriginalCodeToInsert, CommandReference & "STARTCODA") Then
                                                 II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "STARTCODA")
                                              ElseIf InStr(II.OriginalCodeToInsert, CommandReference & "StartCoda") Then
                                                 II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "StartCoda")
                                              ElseIf InStr(II.OriginalCodeToInsert, CommandReference & "startcoda") Then
                                                 II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "startcoda")
                                              End If
                                        Case "LOOPWHILE"
                                              If InStr(II.OriginalCodeToInsert, CommandReference & "STARTLOOPWHILE") Then
                                                 II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "STARTLOOPWHILE")
                                              ElseIf InStr(II.OriginalCodeToInsert, CommandReference & "StartLoopWhile") Then
                                                 II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "StartLoopWhile")
                                              ElseIf InStr(II.OriginalCodeToInsert, CommandReference & "startloopwhile") Then
                                                 II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "startloopwhile")
                                              End If
                                        Case "LOOPUNTIL"
                                              If InStr(II.OriginalCodeToInsert, CommandReference & "STARTLOOPUNTIL") Then
                                                 II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "STARTLOOPUNTIL")
                                              ElseIf InStr(II.OriginalCodeToInsert, CommandReference & "StartLOOPUNTIL") Then
                                                 II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "StartLOOPUNTIL")
                                              ElseIf InStr(II.OriginalCodeToInsert, CommandReference & "startloopuntil") Then
                                                 II.LinesLeftToProcess = sAfter(II.OriginalCodeToInsert, 1, CommandReference & "startloopuntil")
                                              End If
                                 End Select
                                 
                                 If Left$(II.LinesLeftToProcess, 2) = gsEOL Then
                                    II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, 3)
                                 End If
                                 GoTo CODA_RESTART
                              End If

                         Case "NOINSERT", "STOPCODEINSERTION"
                               ISandyWindowMain_InternalInsertTemplate = True
                               GoTo EH_InsertTemplate_Continue             ' Prematurely stop processing of this template

                        Case "RESUMEINSERTION", "RESUME"                                          ' Clears file insertion and resume code insertion
                             If Len(II.ExternalFilename) > 0 And Len(II.TextToSendToFile) > 0 Then
 On Error Resume Next                                                                   ' First save any results to a previously mentioned file.
                                fh = FreeFile
                                Open II.ExternalFilename For Append Access Write As #fh
                                     Print #fh, II.TextToSendToFile
                                Close #fh
                                II.TextToSendToFile = vbNullString
                                Err.Clear
 On Error GoTo EH_InsertTemplate
                             End If
                             II.ExternalFilename = vbNullString

' ********************************************************
' Soft Commands that specially process areas of code/forms
' ********************************************************
                        Case "COMMENTEDPARAMETERS"          ' Parse a function's parameters into readable comments
                             ISandyWindowMain_GetProcAtLine II.PointOfInsertion, sProcName, lProcType
                             If sProcName <> vbNullString Then
                                If Len(II.AllParameters) Then
                                   If lTokenCount(II.AllParameters) = 3 Then
                                      sDelim1 = sGetToken(II.AllParameters, 2)
                                      sDelim2 = sGetToken(II.AllParameters, 3)
                                   ElseIf lTokenCount(II.AllParameters) = 2 Then
                                      sDelim1 = sGetToken(II.AllParameters, 2)
                                      sDelim2 = "||"
                                   ElseIf lTokenCount(II.AllParameters) = 1 Then
                                      sDelim1 = "$$"
                                      sDelim2 = "||"
                                   Else
                                      II.AllParameters = ""
                                   End If
                                Else
                                   sDelim1 = ""
                                   sDelim2 = ""
                                End If
                                lStartLine = CurModule.ProcBodyLine(sProcName, lProcType)
                                II.sParam = CurModule.Lines(lStartLine, 1) ' Get the procedure's header
                                Do While Right$(Trim$(II.sParam), 2) = " _"
                                   lStartLine = lStartLine + 1
                                   II.sParam = Trim$(sBefore(II.sParam, lTokenCount(II.sParam, " _"), " _")) & " " & Trim$(CurModule.Lines(lStartLine, 1)) ' Get the next procedure's header line
                                Loop
                                bFunction = (InStr(II.sParam, "Function") > 0) Or (InStr(II.sParam, "Property Get") > 0)
                                II.sParam = sAfter(II.sParam, 1, "(")                         ' Get just the parameters
                                If bFunction Then
                                   If lTokenCount(II.sParam, ") As ") > 1 Then
                                      sHold1 = sAfter(II.sParam, lTokenCount(II.sParam, ") As ") - 1, ") As ")
                                   Else
                                      sHold1 = "Variant"
                                   End If
                                Else
                                   sHold1 = vbNullString
                                End If
                                If lTokenCount(II.sParam, ")") > 1 Then
                                   II.sParam = sBefore(II.sParam, lTokenCount(II.sParam, ")"), ")")                        ' Strip out the return type (FUTURE: Possibly use this later)
                                End If
                                lParamCount = lTokenCount(II.sParam, ",")                  ' Find out how many parameters there are
                                If Len(II.sParam) = 0 Then
                                   If Len(II.AllParameters) = 0 Then
                                      sCurType = "'      None" & gsEOL
                                   Else
                                      sCurType = ""
                                   End If
                                Else
                                   For CurrParam = 1 To lParamCount                        ' For each parameter
                                        sCurParam = Trim$(sGetToken(II.sParam, 1, ","))        ' Get the next parameter
                                        II.sParam = sAfter(II.sParam, 1, ",")                     ' Chop off the parameter
                                        If InStr(sCurParam, "As") > 0 Then
                                           sHold2 = sGetToken(sCurParam, 2, " As ") & "."
                                           If InStr(sHold2, "=") > 0 Then
                                              sHold2 = Trim$(sGetToken(sHold2, 1, "=")) & " Defaults to " & Trim$(sAfter(sHold2, 1, "=")) ' & "."
                                           End If
                                           sCurParam = sGetToken(sCurParam, 1, " As ")
                                           If InStr(sCurParam, "Optional") > 0 Then
                                              sHold2 = "Opt. " & sHold2
                                              sCurParam = Trim$(Replace(sCurParam, "Optional", vbNullString))
                                           End If
                                           If InStr(sCurParam, "ByVal ") > 0 Then
                                              sHold2 = "(I)  " & sHold2
                                              sCurParam = Replace(sCurParam, "ByVal ", vbNullString)
                                           ElseIf InStr(sCurParam, "ByRef ") > 0 Then
                                              sHold2 = "(O)  " & sHold2
                                              sCurParam = Replace(sCurParam, "ByRef ", vbNullString)
                                           Else
                                              sHold2 = "(IO) " & sHold2
                                           End If
                                           
                                           If Len(sCurParam) > 2 Then
                                              If Right$(sCurParam, 2) = "()" Then ' Array
                                                 sCurParam = Left$(sCurParam, Len(sCurParam) - 2)
                                                 If Left$(sHold2, 4) = "(IO)" Then
                                                    sHold2 = "(O) " & Mid$(Left$(sHold2, Len(sHold2) - 1), 5) & " Array."
                                                 Else
                                                    sHold2 = Left$(sHold2, Len(sHold2) - 1) & " Array."
                                                 End If
                                              End If
                                           End If
                                        Else
                                           sHold2 = "Variant."
                                           If InStr(sCurParam, "Optional") > 0 Then
                                              sHold2 = "Opt. " & sHold2
                                              sCurParam = Trim$(Replace(sCurParam, "Optional", vbNullString))
                                           End If
                                           If InStr(sCurParam, "ByVal ") > 0 Then
                                              sHold2 = "(I)  " & sHold2
                                              sCurParam = Replace(sCurParam, "ByVal ", vbNullString)
                                           ElseIf InStr(sCurParam, "ByRef ") > 0 Then
                                              sHold2 = "(O)  " & sHold2
                                              sCurParam = Replace(sCurParam, "ByRef ", vbNullString)
                                           Else
                                              sHold2 = "(IO) " & sHold2
                                           End If
                                           'sCurType = sCurType & "'      " & sCurParam & Space(30 - Len(sCurParam)) & "." & gsEOL
                                        End If
                                        If Len(II.AllParameters) = 0 Then
                                           sCurType = sCurType & "'      " & sCurParam & Space(30 - Len(sCurParam)) & sHold2 & gsEOL
                                        Else
                                           sHold2 = Replace(Replace(sHold2, "(", ""), ") ", sDelim2)
                                           sCurType = sCurType & sCurParam & sDelim2 & Left$(sHold2, Len(sHold2) - 1) & sDelim1
                                        End If
                                    Next CurrParam                                          ' Add it to the growing commented parameters string
                                    If Len(II.AllParameters) = 0 Then
                                       If Len(sHold1) > 0 Then
                                          sCurType = sCurType & "'" & gsEOL
                                          sCurType = sCurType & "' Returns" & gsEOL
                                          sCurType = sCurType & "'      " & sHold1 & Space(30 - Len(sHold1)) & "." & gsEOL
                                       End If
                                    Else
                                       II.SoftVars("Function Returns") = sHold1
                                    End If
                                End If
                             End If
                             If Len(II.AllParameters) = 0 Then
                                II.LinesLeftToProcess = gsEOL & sCurType & sAfter(II.LinesLeftToProcess, 1, gsEOL)              ' Insert the commented parameters into the insertion stream
                             Else
                                II.SoftVars(sGetToken(II.AllParameters)) = sCurType
                             End If
 On Error GoTo EH_InsertTemplate

                         Case "FOREACHCONTROL"
                              Set ControlVars = CreateObject("SandySupport.CAssocArray") ' Collect what to do for each type of control encountered
                              Do Until Left$(II.CurrentLineToProcess, 2) = gsSoftCmdDelimiter
                                 If Left$(II.LinesLeftToProcess, 2) = gsEOL Then                                          ' Strip off the line just parsed
                                    II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, 3)
                                 Else
                                    II.LinesLeftToProcess = sAfter(II.LinesLeftToProcess, 1, gsEOL)
                                 End If
                                 II.CurrentLineToProcess = sGetToken(II.LinesLeftToProcess, 1, gsEOL)
                                 If Left$(II.CurrentLineToProcess, 2) = gsSoftCmdDelimiter Then
                                 ElseIf Left$(II.CurrentLineToProcess, 2) = "**" Then
                                    sCurType = UCase$(Trim$(Mid$(II.CurrentLineToProcess, 3)))
                                    ControlVars.Add sCurType
                                 ElseIf Len(sCurType) > 0 Then
                                    ControlVars(sCurType) = ControlVars(sCurType) & II.CurrentLineToProcess & gsEOL
                                 End If
                              Loop
                              For lCurControl = 1 To Parent.SandyIDE.SelectedComponent.Designer.Controls.Count
                                  Set CurControl = Parent.SandyIDE.SelectedComponent.Designer.Controls(lCurControl)
                                  If Len(ControlVars(UCase$(CurControl.ClassName))) > 0 Then
                                     sCurType = ControlVars(UCase$(CurControl.ClassName))
                                     Do Until InStr(sCurType, "**") = 0
                                        II.sParam = sGetToken(sCurType, 2, "**")
                                        If UCase$(Left$(II.sParam, 8)) = "CONTROL." Then
                                           II.sParam = Mid$(II.sParam, 9)
                                        End If
                                        If II.sParam = "sName" Then
                                           II.sParam = Mid$(CurControl.Properties("Name"), 4)
                                        Else
                                           II.sParam = CurControl.Properties(II.sParam)
                                       End If
                                        sCurType = sGetToken(sCurType, 1, "**") & II.sParam & sAfter(sCurType, 2, "**")
                                     Loop
                                     If Right$(sCurType, 2) = gsEOL Then
                                        sCurType = Left$(sCurType, Len(sCurType) - 2)
                                     End If
                                     CurModule.InsertLines II.PointOfInsertion, sCurType
                                     II.PointOfInsertion = II.PointOfInsertion + lTokenCount(sCurType, gsEOL)
                                  End If
                              Next lCurControl

                              Set ControlVars = Nothing

                         Case "FOREACHCONTROLBYFRAME"
                              Set ControlVars = CreateObject("SandySupport.CAssocArray") ' Collect what to do for each type of control encountered
                              Do Until Left$(II.CurrentLineToProcess, 2) = gsSoftCmdDelimiter
                                 If Left$(II.LinesLeftToProcess, 2) = gsEOL Then                                          ' Strip off the line just parsed
                                    II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, 3)
                                 Else
                                    II.LinesLeftToProcess = sAfter(II.LinesLeftToProcess, 1, gsEOL)
                                 End If
                                 II.CurrentLineToProcess = sGetToken(II.LinesLeftToProcess, 1, gsEOL)
                                 If Left$(II.CurrentLineToProcess, 2) = gsSoftCmdDelimiter Then
                                 ElseIf Left$(II.CurrentLineToProcess, 2) = "**" Then
                                    sCurType = UCase$(Trim$(Mid$(II.CurrentLineToProcess, 3)))
                                    ControlVars.Add sCurType
                                 ElseIf Len(sCurType) > 0 Then
                                    ControlVars(sCurType) = ControlVars(sCurType) & II.CurrentLineToProcess & gsEOL
                                 End If
                              Loop
                              For lCurFrame = 1 To Parent.SandyIDE.SelectedComponent.Designer.Controls.Count
                                  Set CurFrame = Parent.SandyIDE.SelectedComponent.Designer.Controls(lCurFrame)
                                  If CurFrame.ClassName = "Frame" Then
                                     sCurType = ControlVars(UCase$(CurFrame.ClassName))
                                     Do Until InStr(sCurType, "**") = 0
                                        II.sParam = sGetToken(sCurType, 2, "**")
                                        If UCase$(Left$(II.sParam, 8)) = "CONTROL." Then
                                           II.sParam = Mid$(II.sParam, 9)
                                        End If
                                        If II.sParam = "sName" Then
                                           II.sParam = Mid$(CurFrame.Properties("Name"), 4)
                                        Else
                                           II.sParam = CurFrame.Properties(II.sParam)
                                        End If
                                        sCurType = sGetToken(sCurType, 1, "**") & II.sParam & sAfter(sCurType, 2, "**")
                                     Loop
                                     If Len(sCurType) > 0 Then
                                        If Right$(sCurType, 2) = gsEOL Then
                                           sCurType = Left$(sCurType, Len(sCurType) - 2)
                                        End If
                                        CurModule.InsertLines II.PointOfInsertion, sCurType
                                        II.PointOfInsertion = II.PointOfInsertion + lTokenCount(sCurType, gsEOL)
                                     End If
                                     For lCurControl = 1 To CurFrame.ContainedControls.Count
                                           Set CurControl = CurFrame.ContainedControls(lCurControl)
                                           If CurControl.ClassName <> "Frame" And Len(ControlVars(UCase$(CurControl.ClassName))) > 0 Then
                                              sCurType = ControlVars(UCase$(CurControl.ClassName))
                                              Do Until InStr(sCurType, "**") = 0
                                                 II.sParam = sGetToken(sCurType, 2, "**")
                                                 If UCase$(Left$(II.sParam, 8)) = "CONTROL." Then
                                                    II.sParam = Mid$(II.sParam, 9)
                                                 End If
                                                 If II.sParam = "sName" Then
                                                    II.sParam = Mid$(CurControl.Properties("Name"), 4)
                                                 Else
                                                    II.sParam = CurControl.Properties(II.sParam)
                                                End If
                                                 sCurType = sGetToken(sCurType, 1, "**") & II.sParam & sAfter(sCurType, 2, "**")
                                              Loop
                                              If Right$(sCurType, 2) = gsEOL Then
                                                 sCurType = Left$(sCurType, Len(sCurType) - 2)
                                              End If
                                              CurModule.InsertLines II.PointOfInsertion, sCurType
                                              II.PointOfInsertion = II.PointOfInsertion + lTokenCount(sCurType, gsEOL)
                                           End If
                                           Set CurControl = Nothing
                                     Next lCurControl
                                  End If
                                  Set CurFrame = Nothing
                              Next lCurFrame

                              Set ControlVars = Nothing

' ***********************************************************************
' Soft commands that directly manipulate VB module(s)/Forms/Controls/etc.
' ***********************************************************************
                         Case "CLOSE", "CLOSECODE", "CLOSEWINDOW"
                               CurModule.CodePane.Window.Close

                         Case "HIDE", "HIDECODE", "HIDEWINDOW"
                               CurModule.CodePane.Window.Visible = False

                         Case "FIND", "LOCATE", "SEARCH"
                               lStartLine = II.PointOfInsertion
                               lStartColFound = 1
                               lEndLine = CurModule.CountOfLines '- II.PointOfInsertion + 1
                               lEndColFound = -1
                               If CurModule.Find(II.AllParameters, lStartLine, lStartColFound, lEndLine, lEndColFound) Then
                                  II.SoftVars("Found").Value = lStartLine & vbNullString
                               Else
                                  II.SoftVars("Found").Value = "0"
                               End If

                         Case "FINDINPROC", "PROCFIND", "PFIND", "PROCLOCATE", "PLOCATE", "PROCSEARCH", "PSEARCH"
                               ISandyWindowMain_GetProcAtLine II.PointOfInsertion, sProcName, lProcType
                               lStartLine = CurModule.ProcStartLine(sProcName, lProcType)
                               lEndLine = CurModule.ProcCountLines(sProcName, lProcType) + lStartLine - 1
                               lStartColFound = 1
                               lEndColFound = -1
                               If CurModule.Find(II.AllParameters, lStartLine, lStartColFound, lEndLine, lEndColFound) Then
                                  II.SoftVars("Found").Value = lStartLine & vbNullString
                               Else
                                  II.SoftVars("Found").Value = "0"
                               End If

                         Case "DELETEPROC"
                               lStartLine = CurModule.ProcStartLine(sProcName, lProcType)
                               lEndLine = CurModule.ProcCountLines(sProcName, lProcType)
                               CurModule.DeleteLines lStartLine, lEndLine

                         Case "DELETELINES"
                               CurModule.DeleteLines II.PointOfInsertion, II.ParamLineOffset

                         Case "DELETELINE"
                               CurModule.DeleteLines II.PointOfInsertion

                        Case "PROCATTR"                     ' Modify the current procedure's attributes (Ouch ! That's cool !)
On Error Resume Next                                        ' Prevent illegal values from causing an error
                             Select Case UCase$(sProcName)
                                    Case "ID"               ' Set a default or NewEnum property
                                          ISandyWindowMain_GetProcAtLine II.PointOfInsertion, sProcName, lProcType
                                          If InStr(UCase$(II.sParam), "DEFAULT") > 0 Then
                                             CurModule.Members(sProcName).StandardMethod = 0
                                          ElseIf InStr(UCase$(II.sParam), "NEWENUM") > 0 Then
                                             CurModule.Members(sProcName).StandardMethod = -4
                                          Else
                                             CurModule.Members(sProcName).StandardMethod = Val(II.sParam)
                                          End If
                                    Case "HIDDEN"           ' Hide/unhide the property
                                          ISandyWindowMain_GetProcAtLine II.PointOfInsertion, sProcName, lProcType
                                          CurModule.Members(sProcName).Hidden = IIf(UCase$(II.sParam) = "TRUE" Or UCase$(II.sParam) = "T", True, False)
                                    Case "DESC"             ' Add a description to the property
                                          ISandyWindowMain_GetProcAtLine II.PointOfInsertion, sProcName, lProcType
                                          CurModule.Members(sProcName).Description = II.sParam
                             End Select
On Error GoTo EH_InsertTemplate                              ' Resume normal error processing

                         Case "READLINE"
On Error Resume Next
                             II.SoftVars(sGetToken(II.AllParameters)).Value = CurModule.Lines(II.PointOfInsertion, 1)
On Error GoTo EH_InsertTemplate                              ' Resume normal error processing

                         Case "POSTFIXLINE", "POSTFIX"
                             If Len(II.AllParameters) Then
                                sT = CurModule.Lines(II.PointOfInsertion, 1)
                                CurModule.DeleteLines II.PointOfInsertion, 1
                                CurModule.InsertLines II.PointOfInsertion, sT & II.AllParameters
                             End If

                         Case "PREFIXLINE", "PREFIX"
                             If Len(II.AllParameters) Then
                                sT = CurModule.Lines(II.PointOfInsertion, 1)
                                CurModule.DeleteLines II.PointOfInsertion, 1
                                CurModule.InsertLines II.PointOfInsertion, II.AllParameters & sT
                                II.PointOfInsertion = II.PointOfInsertion + lTokenCount(II.AllParameters, vbNewLine) - 1
                             End If

                        Case "GETTEXTSELECTION", "GETTEXT", "GETSELECTION"
                             II.SoftVars(sGetToken(II.AllParameters)).Value = ISandyWindowMain_GetCurrentTextSelection
On Error GoTo EH_InsertTemplate

                        Case "GETCLIPBOARDTEXT", "GETCLIPBOARD", "GETCLIP"
On Error Resume Next
                             II.SoftVars(sGetToken(II.AllParameters)).Value = Clipboard.GetText(vbCFText)
On Error GoTo EH_InsertTemplate

                        Case "SETCLIPBOARDTEXT", "SETCLIPBOARD", "SETCLIP"
On Error Resume Next
                             Clipboard.SetText II.SoftVars(sGetToken(II.AllParameters)), vbCFText
On Error GoTo EH_InsertTemplate

                        Case "DELETESELECTION"
                             ISandyWindowMain_DeleteCurrentTextSelection

                        Case "LASTSELECTIONLINE"
                             II.PointOfInsertion = ISandyWindowMain_DetermineLastLineInSelection

                        Case "FIRSTSELECTIONLINE"
                             II.PointOfInsertion = ISandyWindowMain_DetermineFirstLineInSelection
                             
                        Case "SELECTCONTROL"
On Error Resume Next
                             Set CurDesigner = Parent.SandyIDE.SelectedComponent.Designer
                             Set II.CurrControl = CurDesigner.Controls(sGetToken(II.AllParameters, 1))
On Error GoTo EH_InsertTemplate
                             
                        Case "ADDCONTROL"
On Error Resume Next
                             If InStr(sGetToken(II.AllParameters), ".") > 0 And InStr(sGetToken(II.AllParameters, 2), ".") = 0 Then
                                II.AllParameters = sGetToken(II.AllParameters, 2) & " " & sGetToken(II.AllParameters)
                             End If
                             Set CurDesigner = Parent.SandyIDE.SelectedComponent.Designer
                             Set II.CurrControl = CurDesigner.Controls.Add(sGetToken(II.AllParameters, 2))
                             If II.CurrControl Is Nothing Then
                                MsgBox "The '" & II.sParam & "' control has not been referenced yet. Please add a reference first.", vbInformation
                                CancelInsertion = bUserSure("Cancel processing ?")
                                If CancelInsertion Then GoTo EH_InsertTemplate_Continue
                             Else
                                II.CurrControl.Properties("Name") = sGetToken(II.AllParameters)
                             End If
                            'If Len(m_asaMisc("Started")) = 0 Then
                            '   II.CurrControl.Properties("Top") = 100
                            '   II.CurrControl.Properties("Left") = 100
                            '   m_asaMisc("Started") = "True"
                            'Else
                            '   With II.CurrControl.Properties
                            '        Select Case UCase$(m_asaMisc("ProgID"))
                            '               Case "VB.LABEL"
                            '                    .Item("Left") = 1900 ' Val(m_asaMisc("Left")) + Val(m_asaMisc("Width")) + 100
                            '                    .Item("Top") = m_asaMisc("Top")
                            '
                            '               Case "VB.TEXTBOX"
                            '                    .Item("Left") = 100 ' m_asaMisc("Left")
                            '                    .Item("Top") = Val(m_asaMisc("Top")) + Val(m_asaMisc("Height")) + 100
                            '
                            '               Case Else
                            '                    .Item("Left") = m_asaMisc("Left")
                            '                    .Item("Top") = Val(m_asaMisc("Top")) + Val(m_asaMisc("Height")) + 100
                            '        End Select
                            '   End With
                            'End If
                            ' With II.CurrControl.Properties
                            '      Select Case UCase$(II.CurrControl.ProgId)
                            '             Case "VB.LABEL"
                            '                 .Item("AutoSize") = True
                            '             Case "VB.TEXTBOX"
                            '                 .Item("Height") = 300
                            '                 .Item("Width") = 1000
                            '                 .Item("Text") = vbNullString
                            '            Case "VB.CHECKBOX"
                            '                 .Item("Height") = 300
                            '                 .Item("Width") = 1500
                            '                 .Item("Left") = 1900
                            '           Case Else
                            '   End Select
                               'm_asaMisc("Top") = .Item("Top")
                               'm_asaMisc("Left") = .Item("Left")
                               'm_asaMisc("Width") = .Item("Width")
                               'm_asaMisc("Height") = .Item("Height")
                               'm_asaMisc("ProgID") = II.CurrControl.ProgId
                            ' End With
On Error GoTo EH_InsertTemplate

                        Case "SETPROPERTY"
On Error Resume Next
                             If Not II.CurrControl Is Nothing Then
                                'If UCase$(sGetToken(II.AllParameters, 1, "=")) = "DATASOURCE" Then
                                '   If Not II.CurrControl.Container.ContainedVBControls(sAfter(II.AllParameters, 1, "=")) Is Nothing Then
                                '      II.CurrControl.ControlObject.DataSource = sAfter(II.AllParameters, 1, "=") 'II.CurrControl.Container.ContainedVBControls(sAfter(II.AllParameters, 1, "="))
                                '   End If
                                'Else
                                'End If
                                II.CurrControl.Properties(II.Result) = II.Expression
                             End If
On Error GoTo EH_InsertTemplate

                        Case "ADDFILEREFERENCE", "ADDFILEREF"
On Error Resume Next
                             Err.Clear
                             bFoundReference = False
                             For Each CurReference In Parent.SandyIDE.ActiveProject.References
                                 If UCase$(CurReference.FullPath) = UCase$(II.AllParameters) Then
                                    bFoundReference = True
                                 End If
                             Next CurReference
                             If Not bFoundReference Then
                                Parent.SandyIDE.ActiveProject.References.AddFromFile II.AllParameters
                             End If
                            'Parent.SandyIDE.ActiveProject.AddToolboxProgID "FirmSolutionsDV.DataView" ', II.AllParameters
                             If Err.Number <> 0 Then
                                MsgBox "Failed to add a reference/component by Filename '" & II.AllParameters & "'"
                                Err.Clear
                                CancelInsertion = bUserSure("Cancel processing ?")
                                If CancelInsertion Then GoTo EH_InsertTemplate_Continue
                             End If
On Error GoTo EH_InsertTemplate

                        Case "ADDFILE", "INCLUDEFILE", "INCLUDE"
On Error Resume Next
                             Err.Clear
                             CurModule.AddFromFile II.AllParameters
                             If Err.Number <> 0 Then
                                MsgBox "Failed to include the File '" & II.AllParameters & "'"
                                Err.Clear
                                CancelInsertion = bUserSure("Cancel processing ?")
                                If CancelInsertion Then GoTo EH_InsertTemplate_Continue
                             End If
On Error GoTo EH_InsertTemplate

                        Case "ADDREFERENCE", "ADDCOMPONENT", "ADDREF"
On Error Resume Next
                             If Left$(II.AllParameters, 1) = "{" Then
                                ' GUID Passed
                                Err.Clear
                                Parent.SandyIDE.ActiveProject.References.AddFromGuid sGetToken(II.AllParameters), CLng(sGetToken(II.AllParameters, 2)), CLng(sGetToken(II.AllParameters, 3))
                                If Err.Number <> 0 Then
                                   MsgBox "Failed to add a reference/component by GUID '" & II.AllParameters & "'"
                                   Err.Clear
                                   CancelInsertion = bUserSure("Cancel processing ?")
                                   If CancelInsertion Then GoTo EH_InsertTemplate_Continue
                                End If
                             Else
                             '   ' ProgID Passed
                                Err.Clear
                                Parent.SandyIDE.ActiveProject.AddToolboxProgID sGetToken(II.AllParameters)
                                If Err.Number <> 0 Then
                                 ' Attempt to add a reference by looking up the GUID for the ProgID passed
                                   Err.Clear
                                   
                                   sHold1 = sGetGUID(sGetToken(II.AllParameters))
                                   If Len(sHold1) > 0 Then
                                      Err.Clear
                                      Parent.SandyIDE.ActiveProject.References.AddFromGuid sHold1, 0, 0  'CLng(sGetToken(II.AllParameters, 2)), CLng(sGetToken(II.AllParameters, 3))
                                   End If
                                   If Err.Number <> 0 Then
                                      CancelInsertion = bUserSure("Failed to add a reference/component by ProgID '" & II.AllParameters & "'" & vbNewLine & vbNewLine & vbTab & "Cancel processing ?")
                                      If CancelInsertion Then GoTo EH_InsertTemplate_Continue
                                   End If
                                End If
                             End If
On Error GoTo EH_InsertTemplate

                        Case "SETFORMPROPERTY", "FORMPROPERTY"
On Error Resume Next
                             Parent.SandyIDE.SelectedComponent.Properties(II.Result) = II.Expression
On Error GoTo EH_InsertTemplate

'                        Case "SETDEFAULT"
'On Error Resume Next
'                             If InStr(II.AllParameters, "=") > 0 Then
'                                m_asaDefaults(sGetToken(II.AllParameters, 1, "=")) = sGetToken(II.AllParameters, 2, "=")
'                             End If
'On Error GoTo EH_InsertTemplate
'
'                        Case "RESETDEFAULTS"
'                             m_asaDefaults.Clear
'

' ************************************************
' Soft commands that change the point of insertion
' ************************************************
                         Case "GOTOPROJECT"
On Error Resume Next
                               If FindInCollection(Parent.SandyIDE.Projects, sProcName) Is Nothing Then
                                  Select Case UCase$(II.sParam)
                                         Case "CONTROL"
                                              With Parent.SandyIDE.Projects.Add(vbext_pt_ActiveXControl, False)
                                                   .Name = sProcName
                                                   .Components(1).Activate
                                                   .Components(1).CodeModule.CodePane.Show
                                              End With
                                         
                                         Case "EXE"
                                              With Parent.SandyIDE.Projects.Add(vbext_pt_StandardExe, False)
                                                   .Name = sProcName
                                                   .Components(1).Activate
                                                   .Components(1).CodeModule.CodePane.Show
                                              End With

                                         Case "ACTIVEXEXE"
                                              With Parent.SandyIDE.Projects.Add(vbext_pt_ActiveXExe, False)
                                                   .Name = sProcName
                                                   .Components(1).Activate
                                                   .Components(1).CodeModule.CodePane.Show
                                              End With

                                         Case Else ' "DLL"
                                              With Parent.SandyIDE.Projects.Add(vbext_pt_ActiveXDll, False)
                                                   .Name = sProcName
                                                   .Components(1).Activate
                                                   .Components(1).CodeModule.CodePane.Show
                                              End With
                                  End Select

                               Else
                                  With Parent.SandyIDE.Projects(sProcName)
                                       .Components(1).CodeModule.CodePane.Show
                                  End With
                               End If

                               II.SoftVars("Project Name").Value = Parent.SandyIDE.ActiveProject.Name                                 ' Add the build in soft variables
                               II.SoftVars("Module Name").Value = Parent.SandyIDE.ActiveCodePane.CodeModule.Parent.Name
                               II.SoftVars("Proc Name").Value = vbNullString
                               II.SoftVars("Proc Type").Value = vbNullString
                               II.SoftVars("Proc Type Long").Value = vbNullString

                               Set CurModule = Parent.SandyIDE.ActiveCodePane.CodeModule
On Error GoTo EH_InsertTemplate
                               
                         Case "GOTOMODULE", "GOTOCLASS", "GOTOFORM"
                               'If Right$(sProcName, 2) = "ss" Then
                               '   sProcName = Left$(sProcName, Len(sProcName) - 1)
                               'End If

                               If FindInCollection(Parent.SandyIDE.Components, sProcName) Is Nothing Then
                                  II.sParam = UCase$(II.sParam)
                                  If Len(II.sParam) = 0 Then
                                     II.sParam = UCase$(Mid$(sGetToken(II.CurrentLineToProcess), 5))
                                  End If
                                  Select Case UCase$(II.sParam)
                                         Case "CLASS", "CLASSMODULE"
                                              With Parent.SandyIDE.AddComponent(vbext_ct_ClassModule)
                                                   DoEvents: DoEvents: DoEvents: DoEvents
                                                   .Name = sProcName
                                                   .Activate
                                                   .CodeModule.CodePane.Show
                                                   Set CurModule = .CodeModule
                                              End With
                                         Case "FORM"
                                              With Parent.SandyIDE.AddComponent(vbext_ct_VBForm)
                                                   DoEvents: DoEvents: DoEvents: DoEvents
                                                   .Name = sProcName
                                                   .Activate
                                                   .CodeModule.CodePane.Show
                                                   Set CurModule = .CodeModule
                                              End With
                                         Case "MDIFORM"
                                              With Parent.SandyIDE.AddComponent(vbext_ct_VBMDIForm)
                                                   DoEvents: DoEvents: DoEvents: DoEvents
                                                   .Name = sProcName
                                                   .Activate
                                                   .CodeModule.CodePane.Show
                                                   Set CurModule = .CodeModule
                                              End With
                                         Case "USERCONTROL", "CONTROL", "ACTIVEXCONTROL", "ACTIVEX"
                                              With Parent.SandyIDE.AddComponent(vbext_ct_UserControl)
                                                   DoEvents: DoEvents: DoEvents: DoEvents
                                                   .Name = sProcName
                                                   .Activate
                                                   .CodeModule.CodePane.Show
                                                   Set CurModule = .CodeModule
                                              End With
                                         Case Else '"MODULE"
                                              With Parent.SandyIDE.AddComponent(vbext_ct_StdModule)
                                                   DoEvents: DoEvents: DoEvents: DoEvents
                                                   .Name = sProcName
                                                   .Activate
                                                   .CodeModule.CodePane.Show
                                                   Set CurModule = .CodeModule
                                              End With
                                  End Select
                               Else
                                  With Parent.SandyIDE.Components(sProcName)
                                       .Activate
                                       .CodeModule.CodePane.Show
                                       Set CurModule = .CodeModule
                                  End With
                               End If
                               II.SoftVars("Module Name").Value = sProcName
                               
                              'II.LinesLeftToProcess = sAfter(II.LinesLeftToProcess, 1, gsEOL)
                               

                         Case "GOTOPROC"                    ' Set the current line to insert to the indicated line in the indicated procedure
                               If II.ParamLineOffset = 0 Then II.ParamLineOffset = 1
                               II.PointOfInsertion = CurModule.ProcBodyLine(sProcName, lProcType)
                               If Err.Number <> 0 Then
                                  If InStr(sProcName, "_") Then
                                     Err.Clear
                                     II.PointOfInsertion = CurModule.CreateEventProc(sGetToken(sProcName, 2, "_"), sGetToken(sProcName, 1, "_"))
                                     If Err.Number <> 0 Then
                                        CancelInsertion = bUserSure("FindLastProcLine" & vbCr & vbTab & "Can't find:" & vbCr & vbTab & vbTab & sProcName & vbCr & vbTab & "In module:" & vbCr & vbTab & vbTab & CurModule.Parent.Name & vbNewLine & vbNewLine & vbTab & "Cancel processing ?")
                                        II.PointOfInsertion = 0
                                        Err.Clear
                                        Exit Function
                                     End If
                                  Else
                                     CancelInsertion = bUserSure("GotoProc" & vbCr & vbTab & "Can't find:" & vbCr & vbTab & vbTab & sProcName & vbCr & vbTab & "In module:" & vbCr & vbTab & vbTab & CurModule.Parent.Name & vbNewLine & vbNewLine & vbTab & "Cancel processing ?")
                                     II.PointOfInsertion = 0
                                     Err.Clear
                                  End If
                               End If
                               II.PointOfInsertion = II.PointOfInsertion + II.ParamLineOffset
                               
                              'GetProcAtLine II.PointOfInsertion, sProcName, lProcType
                               If sProcName <> vbNullString Then
                                  II.SoftVars("Proc Name").Value = sProcName
                                  II.SoftVars("Proc Type").Value = Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
                                  sHold2 = CurModule.Lines(CurModule.ProcBodyLine(sProcName, lProcType), 1)
                                  If InStr(sHold2, "Function") > 0 Then
                                     sHold2 = "Function"
                                  ElseIf InStr(sHold2, "Property") > 0 Then
                                     sHold2 = "Property"
                                  Else
                                     sHold2 = "Sub"
                                  End If
                                  II.SoftVars("Proc Type Long").Value = sHold2
                               Else
                                  II.SoftVars("Proc Name").Value = vbNullString
                                  II.SoftVars("Proc Type").Value = vbNullString
                                  II.SoftVars("Proc Type Long").Value = vbNullString
                               End If

                         Case "GOTOPROCEND"                 ' Set the current line to the last line before "End Sub/Function/Property" in the indicated procedure
                              'II.PointOfInsertion = CurModule.ProcBodyLine(sProcName, lProcType)
                               II.PointOfInsertion = ISandyWindowMain_FindLastProcLine(sProcName, lProcType) + II.ParamLineOffset
                              'GetProcAtLine II.PointOfInsertion, sProcName, lProcType
                               If sProcName <> vbNullString Then
                                  II.SoftVars("Proc Name").Value = sProcName
                                  II.SoftVars("Proc Type").Value = Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
                                  sHold2 = CurModule.Lines(CurModule.ProcBodyLine(sProcName, lProcType), 1)
                                  If InStr(sHold2, "Function") > 0 Then
                                     sHold2 = "Function"
                                  ElseIf InStr(sHold2, "Property") > 0 Then
                                     sHold2 = "Property"
                                  Else
                                     sHold2 = "Sub"
                                  End If
                                  II.SoftVars("Proc Type Long").Value = sHold2
                               Else
                                  II.SoftVars("Proc Name").Value = vbNullString
                                  II.SoftVars("Proc Type").Value = vbNullString
                                  II.SoftVars("Proc Type Long").Value = vbNullString
                               End If

                         Case "ABSLINE"
                               II.PointOfInsertion = Abs(II.ParamLineOffset)     ' Set the current line to the absolute line number specified
                               ISandyWindowMain_GetProcAtLine II.PointOfInsertion, sProcName, lProcType
                               If sProcName <> vbNullString Then
                                  II.SoftVars("Proc Name").Value = sProcName
                                  II.SoftVars("Proc Type").Value = Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
                                  sHold2 = CurModule.Lines(CurModule.ProcBodyLine(sProcName, lProcType), 1)
                                  If InStr(sHold2, "Function") > 0 Then
                                     sHold2 = "Function"
                                  ElseIf InStr(sHold2, "Property") > 0 Then
                                     sHold2 = "Property"
                                  Else
                                     sHold2 = "Sub"
                                  End If
                                  II.SoftVars("Proc Type Long").Value = sHold2
                               Else
                                  II.SoftVars("Proc Name").Value = vbNullString
                                  II.SoftVars("Proc Type").Value = vbNullString
                                  II.SoftVars("Proc Type Long").Value = vbNullString
                               End If

                         Case "LINEOFFSET", "OFFSET"
                               II.PointOfInsertion = II.PointOfInsertion + II.ParamLineOffset  ' Set the current line to the relative line offset specified
                               ISandyWindowMain_GetProcAtLine II.PointOfInsertion, sProcName, lProcType
                               If sProcName <> vbNullString Then
                                  II.SoftVars("Proc Name").Value = sProcName
                                  II.SoftVars("Proc Type").Value = Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
                                  sHold2 = CurModule.Lines(CurModule.ProcBodyLine(sProcName, lProcType), 1)
                                  If InStr(sHold2, "Function") > 0 Then
                                     sHold2 = "Function"
                                  ElseIf InStr(sHold2, "Property") > 0 Then
                                     sHold2 = "Property"
                                  Else
                                     sHold2 = "Sub"
                                  End If
                                  II.SoftVars("Proc Type Long").Value = sHold2
                               Else
                                  II.SoftVars("Proc Name").Value = vbNullString
                                  II.SoftVars("Proc Type").Value = vbNullString
                                  II.SoftVars("Proc Type Long").Value = vbNullString
                               End If

                         Case "PROCTOP"                     ' Move to the top of the current procedure
                               ISandyWindowMain_GetProcAtLine II.PointOfInsertion, sProcName, lProcType
                               If sProcName <> vbNullString Then
                                  II.PointOfInsertion = CurModule.ProcBodyLine(sProcName, lProcType)
                                  II.SoftVars("Proc Name").Value = sProcName
                                  II.SoftVars("Proc Type").Value = Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
                                  sHold2 = CurModule.Lines(CurModule.ProcBodyLine(sProcName, lProcType), 1)
                                  If InStr(sHold2, "Function") > 0 Then
                                     sHold2 = "Function"
                                  ElseIf InStr(sHold2, "Property") > 0 Then
                                     sHold2 = "Property"
                                  Else
                                     sHold2 = "Sub"
                                  End If
                                  II.SoftVars("Proc Type Long").Value = sHold2
                               Else
                                  II.SoftVars("Proc Name").Value = vbNullString
                                  II.SoftVars("Proc Type").Value = vbNullString
                                  II.SoftVars("Proc Type Long").Value = vbNullString
                               End If

                         Case "PROCEND"                     ' Move to the end of the current procedure
                               ISandyWindowMain_GetProcAtLine II.PointOfInsertion, sProcName, lProcType
                               If sProcName <> vbNullString Then
                                  II.PointOfInsertion = ISandyWindowMain_FindLastProcLine(sProcName, lProcType)
                                  II.SoftVars("Proc Name").Value = sProcName
                                  II.SoftVars("Proc Type").Value = Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", True, vbNullString)
                                  sHold2 = CurModule.Lines(CurModule.ProcBodyLine(sProcName, lProcType), 1)
                                  If InStr(sHold2, "Function") > 0 Then
                                     sHold2 = "Function"
                                  ElseIf InStr(sHold2, "Property") > 0 Then
                                     sHold2 = "Property"
                                  Else
                                     sHold2 = "Sub"
                                  End If
                                  II.SoftVars("Proc Type Long").Value = sHold2
                               Else
                                  II.SoftVars("Proc Name").Value = vbNullString
                                  II.SoftVars("Proc Type").Value = vbNullString
                                  II.SoftVars("Proc Type Long").Value = vbNullString
                               End If
                            
                        Case "GOTODECLARATIONS", "GOTODEC"
                             If UCase$(II.AllParameters) = "END" Then
                                II.PointOfInsertion = CurModule.CountOfDeclarationLines + 1
                             Else ' Line 1
                                II.PointOfInsertion = 1
                             End If
                             II.SoftVars("Proc Name").Value = vbNullString
                             II.SoftVars("Proc Type").Value = vbNullString
                             II.SoftVars("Proc Type Long").Value = vbNullString

                        Case "GOTOENDOFFILE", "GOTOENDOFMODULE", "GOTOEND"
                             II.PointOfInsertion = CurModule.CountOfLines + 1
                             II.SoftVars("Proc Name").Value = vbNullString
                             II.SoftVars("Proc Type").Value = vbNullString
                             II.SoftVars("Proc Type Long").Value = vbNullString

' *****************************************
' Soft commands that affect the File System
' *****************************************
'                        Case "DELETEFILE"                                               ' Causes a file in the operating system to be erased.
' On Error Resume Next
'                             Kill II.AllParameters
'                             Err.Clear
' On Error GoTo EH_InsertTemplate

                        Case "FILENAME"                                                 ' Set an external filename to output to
                             If Len(II.ExternalFilename) > 0 And Len(II.TextToSendToFile) > 0 Then
 On Error Resume Next                                                                   ' First save any results to a previously mentioned file.
                                fh = FreeFile
                                Open II.ExternalFilename For Append Access Write As #fh
                                     Print #fh, II.TextToSendToFile
                                Close #fh
                                II.TextToSendToFile = vbNullString
                                Err.Clear
 On Error GoTo EH_InsertTemplate
                             End If
                             II.ExternalFilename = II.AllParameters

                        Case "IGNOREBLANKS", "BLANKSOKAY", "NOBLANKS"
                             mbIgnoreBlanks = True
                        
                        Case "WATCHBLANKS", "BLANKSNOTOKAY", "YESBLANKS"
                             mbIgnoreBlanks = False

                        Case "IGNOREREADONLY"
                             mbIgnoreReadOnly = True

                        Case "WATCHREADONLY"
                             mbIgnoreReadOnly = False

                        Case Else
                             If SadCommandSetCount > 0 Then
                                For CurrSet = 1 To SadCommandSetCount
                                    If SadCommands(CurrSet).ExecuteSoftCommand(II) Then
                                       Exit For
                                    End If
                                Next CurrSet
                             End If
                  End Select

                  '
                  ' ******** End of soft command processing
                  '

               Else
                  If II.PointOfInsertion < 1 Then II.PointOfInsertion = 1
                  If Len(II.ExternalFilename) = 0 Then
                     CurModule.InsertLines II.PointOfInsertion, II.CurrentLineToProcess
                  Else
                     II.TextToSendToFile = II.TextToSendToFile & II.CurrentLineToProcess & gsEOL
                  End If
                  II.PointOfInsertion = II.PointOfInsertion + 1
               End If

               If StrComp(II.LinesLeftToProcess, gsEOL) = 0 Or StrComp(II.LinesLeftToProcess, gs2EOL) = 0 Then
                  II.LinesLeftToProcess = vbNullString
               End If

               If Left$(II.LinesLeftToProcess, 2) = gsEOL Then                                          ' Strip off the line just parsed
                  II.LinesLeftToProcess = Mid$(II.LinesLeftToProcess, 3)
               Else
                  II.LinesLeftToProcess = sAfter(II.LinesLeftToProcess, 1, gsEOL)
               End If
            Loop
         End If
    'End With

    If Len(II.ExternalFilename) > 0 And Len(II.TextToSendToFile) > 0 Then
On Error Resume Next                                                                    ' Save to file if previously mentioned and if there is code to save.
       Kill II.ExternalFilename
       fh = FreeFile
       Open II.ExternalFilename For Output Access Write As #fh
            Print #fh, II.TextToSendToFile
       Close #fh
       Err.Clear
    End If
    
    ISandyWindowMain_InternalInsertTemplate = True

EH_InsertTemplate_Continue:
On Error Resume Next
    Set II.CurrControl = Nothing
   'Set II = Nothing
    Exit Function

EH_InsertTemplate:
    If Err.Number = 40198 And mbIgnoreReadOnly Then
       Resume Next
    Else
       LogError "frmMain", "InsertTemplate", Err.Number, Err.Description
    End If

    Err.Clear

    Resume EH_InsertTemplate_Continue
    
    Resume
End Function

' ********************************************************************************
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
' ********************************************************************************
Public Sub ISandyWindowMain_GetProcAtLine(ByRef lCurrentLine As Long, ByRef sProcName As String, ByRef lProcType As Long)
       Dim ProcType As Long

       With Parent.SandyIDE.ActiveCodePane.CodeModule
            sProcName = .ProcOfLine(lCurrentLine, ProcType)
            lProcType = ProcType
       End With
End Sub

' ********************************************************************************
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
' ********************************************************************************
Public Function ISandyWindowMain_FindLastProcLine(sProcName As String, lProcType As Long) As Long
       Static lLine As Long
       Static lCurLine As Long
       Static lLastLine As Long

       Static sFindString As String
       Static sFunctionHeader As String

On Error Resume Next
       Err.Clear
       With Parent.SandyIDE.ActiveCodePane.CodeModule
            lLine = .ProcStartLine(sProcName, lProcType)                            ' Get the first line number of the procedure
            If Err.Number <> 0 Then
               If InStr(sProcName, "_") Then
                  Err.Clear
                  lLine = .CreateEventProc(sGetToken(sProcName, 2, "_"), sGetToken(sProcName, 1, "_"))
                  If Err.Number <> 0 Then
                     MsgBox "FindLastProcLine" & vbCr & vbTab & "Can't find:" & vbCr & vbTab & vbTab & sProcName & vbCr & vbTab & "In module:" & vbCr & vbTab & vbTab & .Parent.Name
                     CancelInsertion = bUserSure("Cancel processing ?")
                     ISandyWindowMain_FindLastProcLine = 0
                     Err.Clear
                     Exit Function
                  End If
               Else
                  MsgBox "FindLastProcLine" & vbCr & vbTab & "Can't find:" & vbCr & vbTab & vbTab & sProcName & vbCr & vbTab & "In module:" & vbCr & vbTab & vbTab & .Parent.Name
                  CancelInsertion = bUserSure("Cancel processing ?")
                  ISandyWindowMain_FindLastProcLine = 0
                  Err.Clear
                  Exit Function
               End If
            End If
            lLastLine = lLine + .ProcCountLines(sProcName, lProcType)               ' Get the last line number of the procedure
            sFunctionHeader = .Lines(.ProcBodyLine(sProcName, lProcType), 1)        ' Get the procedure's header
            If InStr(sFunctionHeader, "Function") > 0 Then                          ' Based on it's type,
               sFindString = "End Function"                                         '   we can determine what string to look for
            ElseIf InStr(sFunctionHeader, "Sub") > 0 Then
               sFindString = "End Sub"
            Else
               sFindString = "End Property"
            End If

            For lCurLine = lLastLine To lLine Step -1                               ' Move backwards from the end of function
                If InStr(.Lines(lCurLine, 1), sFindString) > 0 Then                 '   until we find the line containing
                   ISandyWindowMain_FindLastProcLine = lCurLine                                      '   "End Function/Sub/Property"
                   sFindString = vbNullString
                   sFunctionHeader = vbNullString
                   Exit Function                                                    ' At this point, we found it, return THAT line #
                End If
            Next lCurLine
       End With

       sFindString = vbNullString
       sFunctionHeader = vbNullString
       ISandyWindowMain_FindLastProcLine = lLastLine                                                 ' Something wrong: Pass back last line # found

End Function


Public Function ISandyWindowMain_JumpTo(ByVal sTemplateName As String, Optional ByVal bRecordInHistory As Boolean = True, Optional ByVal bSyncCategoryList As Boolean = False) As Boolean
On Error GoTo frmMain_EH_JumpTo
    Static sCategoryName As String
    Static sShortTemplateName As String
    Static CurrHE As Long

    Err.Clear
    ISandyWindowMain_SaveTemplate

    sCategoryName = sGetToken(sTemplateName, 1, " - ")
    sShortTemplateName = sAfter(sTemplateName, 1, " - ")
On Error Resume Next
    If SliceAndDice(sCategoryName).Templates(sTemplateName) Is Nothing Then
       ISandyWindowMain_JumpTo = False
       Exit Function
    End If
    Set CurrentTemplate = SliceAndDice(sCategoryName).Templates(sTemplateName)
    ISandyWindowMain_FillAddInScreen

On Error GoTo frmMain_EH_JumpTo

    'CodeAtTop = txtCode(0).Text
    tabCode.Tabs(1).Image = IIf(Len(txtCode(0)) = 0, "Document", "Category")
    tabCode.Tabs(2).Image = IIf(Len(txtCode(1)) = 0, "Document", "Category")
    tabCode.Tabs(3).Image = IIf(Len(txtCode(2)) = 0, "Document", "Category")
    tabCode.Tabs(4).Image = IIf(Len(txtCodeToFile) = 0, "Document", "Category")
    tabCode.Tabs(5).Image = IIf(CurrentTemplate.Undeletable Or CurrentTemplate.Locked Or CurrentTemplate.Selected, "OptionSet", "OptionNotSet")

    If mnuSwitchTabsAutomatically.Checked Then
       If tabCode.Tabs(1).Image = "Category" Then
          tabCode.Tabs(1).Selected = True
       ElseIf tabCode.Tabs(2).Image = "Category" Then
          tabCode.Tabs(2).Selected = True
       ElseIf tabCode.Tabs(3).Image = "Category" Then
          tabCode.Tabs(3).Selected = True
       ElseIf tabCode.Tabs(4).Image = "Category" Then
          tabCode.Tabs(4).Selected = True
       Else
          tabCode.Tabs(1).Selected = True
       End If

       chkLocked_Click
       If chkLocked.Value = 0 Then
          txtShortName.Enabled = (sCategoryName & " - " & sShortTemplateName <> "Change from - All Types")
       Else
          txtShortName.Enabled = False
       End If

       tabCode_MouseUp 0, 0, 0, 0
    Else
       If tabCode.Tabs(6).Selected Then
          If chkAutoRecalc.Value <> 0 Then
             cmdRecalc_Click
          End If
       End If
    End If
    
    If bRecordInHistory Then
       For CurrHE = Val(CurrentHistoryEntry) + 1 To m_asaHistory.Count
           m_asaHistory.Remove vbNullString & CurrHE
       Next CurrHE
       CurrentHistoryEntry = vbNullString & m_asaHistory.Count + 1
       m_asaHistory.Add vbNullString & m_asaHistory.Count + 1, sTemplateName
       mnuBack.Enabled = True
       mnuForward.Enabled = False
    End If
    
    If bSyncCategoryList Then
       lsbJumpTo.BarAndItem sCategoryName, sTemplateName
    End If
    
    ISandyWindowMain_JumpTo = True
    
frmMain_EH_JumpTo_Continue:
    Exit Function

frmMain_EH_JumpTo:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: JumpTo" & vbCr & vbCr & Err.Description
    ISandyWindowMain_JumpTo = False
    Resume frmMain_EH_JumpTo_Continue
    
    Resume
End Function

Public Sub ISandyWindowMain_RefillList()
On Error GoTo EH_frmMain_ISandyWindowMain_RefillList
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Sub
    bInHereAlready = True

    Dim lvwX As Object
    Dim tvwX As Object
    Dim tvwY As Object 'TreeView
    Dim CurrCategory As CCategory
    Dim CurrTemplate As CTemplate
    Dim sOpened As String
    Dim sClosed As String

    lsbJumpTo.Visible = False
    lsbJumpTo.Clear
    For Each CurrCategory In SliceAndDice.Categorys
        With CurrCategory
             If .Deleted Then
               ' Ignore this one
            'Else
             ElseIf CurrCategory.CategoryType = 0 Then
                If lsbJumpTo.BarKey = "Bar 1" Then
                   lsbJumpTo.CurBar = 0
                   lsbJumpTo.BarName = .Key & " (" & Format(CurrCategory.Templates.Count, "00") & ")"
                   lsbJumpTo.BarKey = .Key
                   lsbJumpTo.View = 3
                   lsbJumpTo.Arrange = .Arrange
                   lsbJumpTo.BarType = "List"
On Error Resume Next
                   lsbJumpTo.Bars(0).ColumnHeaders(1).Width = 3400
                Else
                   lsbJumpTo.AddBar(.Key & " (" & Format(CurrCategory.Templates.Count, "00") & ")", .Key).ColumnHeaders(1).Width = 3400
                End If
             Else
                If lsbJumpTo.BarKey = "Bar 1" Then
                   lsbJumpTo.CurBar = 0
                   lsbJumpTo.BarType = "Tree"
                   lsbJumpTo.BarName = .Key & " (Code Gen)"
                   lsbJumpTo.BarKey = .Key
                Else
                   lsbJumpTo.AddBar "[" & CurrCategory.Templates.Count & "] " & .Key, .Key, False
                End If
             End If
        End With
    Next CurrCategory

    For Each CurrCategory In SliceAndDice.Categorys
        If Not CurrCategory.Deleted Then
           lsbJumpTo.CurBar = CurrCategory.Key
           If CurrCategory.CategoryType = 0 Then
              For Each CurrTemplate In CurrCategory.Templates
                  With CurrTemplate
                       If .Deleted Then
                          ' Ignore this one
                       ElseIf .Locked Or .Undeletable Then
                          lsbJumpTo.AddBarItem .ShortTemplateName, .Key, "Key"
                          .OriginalShortName = .ShortTemplateName
                       ElseIf Len(.memoCodeAtBottom & .memoCodeAtCursor & .memoCodeAtTop & .memoCodeToFile) Then
                          lsbJumpTo.AddBarItem .ShortTemplateName, .Key, "DocumentAlternate"
                          .OriginalShortName = .ShortTemplateName
                       Else
                          lsbJumpTo.AddBarItem .ShortTemplateName, .Key, "Document"
                          .OriginalShortName = .ShortTemplateName
                       End If
                  End With
              Next CurrTemplate
           Else
              Set tvwX = lsbJumpTo.Bars(CurrCategory.Key)
              Set tvwY = tvwX
              With tvwY.Nodes
                  sOpened = sTemplateIcon(CurrCategory.Templates("Settings"))
                  With .Add(, , CurrCategory.Key & " - Settings", "Settings", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Routines"))
                  With .Add(, , CurrCategory.Key & " - Routines", "Routines", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Wrapper Class"))
                  With .Add(, , CurrCategory.Key & " - Wrapper Class", "Wrapper Class", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Wrapper class - Add collection"))
                  With .Add(CurrCategory.Key & " - Wrapper Class", tvwChild, CurrCategory.Key & " - Wrapper class - Add collection", "Wrapper class - Add collection", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Collection, No Parent"))
                  With .Add(, , CurrCategory.Key & " - Collection, No Parent", "Collection, No Parent", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Collection, No Parent, No Child"))
                  With .Add(, , CurrCategory.Key & " - Collection, No Parent, No Child", "Collection, No Parent, No Child", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Collection, No Child"))
                  With .Add(, , CurrCategory.Key & " - Collection, No Child", "Collection, No Child", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Collection Member, Terminal"))
                  With .Add(CurrCategory.Key & " - Collection, No Child", tvwChild, CurrCategory.Key & " - Collection Member, Terminal", "Collection Member, Terminal", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Collection"))
                  With .Add(, , CurrCategory.Key & " - Collection", "Collection", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Collection Member"))
                  With .Add(CurrCategory.Key & " - Collection", tvwChild, CurrCategory.Key & " - Collection Member", "Collection Member", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Collection Member - New Subcollection"))
                  With .Add(CurrCategory.Key & " - Collection", tvwChild, CurrCategory.Key & " - Collection Member - New Subcollection", "Collection Member - New Subcollection", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - 3D Link"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - 3D Link", "Property - 3D Link", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - BLOB"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - BLOB", "Property - BLOB", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - Boolean"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Boolean", "Property - Boolean", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - Byte"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Byte", "Property - Byte", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - Currency"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Currency", "Property - Currency", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - Date"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Date", "Property - Date", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - Double"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Double", "Property - Double", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - Integer"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Integer", "Property - Integer", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - Long"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Long", "Property - Long", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - Memo"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Memo", "Property - Memo", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - OLE_COLOR"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - OLE_COLOR", "Property - OLE_COLOR", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - Single"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Single", "Property - Single", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - String"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - String", "Property - String", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Property - Variant"))
                  With .Add(CurrCategory.Key & " - Collection Member", tvwChild, CurrCategory.Key & " - Property - Variant", "Property - Variant", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With
                  sOpened = sTemplateIcon(CurrCategory.Templates("Finalize"))
                  With .Add(, , CurrCategory.Key & " - Finalize", "Finalize", sOpened, sOpened)
                       .ExpandedImage = sOpened
                       .Expanded = True
                  End With

                 ' Fix up what's missing and what's added
                   For Each CurrTemplate In CurrCategory.Templates
                       With CurrTemplate
                            If Not .Deleted Then
                               .OriginalShortName = .ShortTemplateName
                               sOpened = UCase$(CurrTemplate.ShortTemplateName)
                               Select Case sOpened
                                      Case "SETTINGS"
                                      Case "ROUTINES"
                                      Case "WRAPPER CLASS"
                                      Case "WRAPPER CLASS - ADD COLLECTION"
                                      Case "FINALIZE"
                                      Case "COLLECTION"
                                      Case Else
                                           If InStr(sOpened, "COLLECTION ") Or InStr(sOpened, "COLLECTION,") Or InStr(UCase$(CurrTemplate.ShortTemplateName), "PROPERTY - ") Then
                                           Else
                                              sOpened = sTemplateIcon(CurrTemplate)
                                              With tvwX.Nodes.Add(, , CurrCategory.Key & " - " & CurrTemplate.ShortTemplateName, CurrTemplate.ShortTemplateName, sOpened, sOpened)
                                                   .Expanded = True
                                                   .ExpandedImage = sOpened
                                              End With
                                           End If
                               End Select
                            End If
                       End With
                   Next CurrTemplate
              End With
              Set tvwY = Nothing
              Set tvwX = Nothing
           End If
        End If
    Next CurrCategory
    lsbJumpTo.Visible = True

    ISandyWindowMain_UpdateFavorites

    ISandyWindowMain_UpdateHotKeys

EH_frmMain_ISandyWindowMain_RefillList_Continue:
    bInHereAlready = False
    Exit Sub

EH_frmMain_ISandyWindowMain_RefillList:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: ISandyWindowMain_RefillList" & vbCr & vbCr & Err.Description
    
    Resume EH_frmMain_ISandyWindowMain_RefillList_Continue

    Resume
End Sub

Private Sub SetColors(ByVal BackColor As String, ByVal ForeColor As String)
On Error Resume Next
    If Right$(BackColor, 1) = "&" Then BackColor = Left$(BackColor, Len(BackColor) - 1)
    If Right$(ForeColor, 1) = "&" Then ForeColor = Left$(ForeColor, Len(ForeColor) - 1)
    
    If Left$(BackColor, 2) <> "&H" Then BackColor = "&H" & BackColor
    If Left$(ForeColor, 2) <> "&H" Then ForeColor = "&H" & ForeColor
    
    lsbJumpTo.BackColor = BackColor
    txtCode(0).BackColor = BackColor
    txtCode(1).BackColor = BackColor
    txtCode(2).BackColor = BackColor
    txtCodeToFile.BackColor = BackColor
    lstSoftCommands.BackColor = BackColor
    lstSoftVariables.BackColor = BackColor

    lsbJumpTo.ForeColor = ForeColor
    txtCode(0).ForeColor = ForeColor
    txtCode(1).ForeColor = ForeColor
    txtCode(2).ForeColor = ForeColor
    txtCodeToFile.ForeColor = ForeColor
    lstSoftCommands.ForeColor = ForeColor
    lstSoftVariables.ForeColor = ForeColor
    
    'If Not m_oDBClassGen Is Nothing Then
    '   m_oDBClassGen.SetColors ForeColor, BackColor
    'End If
End Sub

Private Property Let ISandyWindowMain_SadCommandSetCount(ByVal RHS As Long)
    SadCommandSetCount = RHS
End Property

Private Property Get ISandyWindowMain_SadCommandSetCount() As Long
    ISandyWindowMain_SadCommandSetCount = SadCommandSetCount
End Property

Private Sub ISandyWindowMain_SetFocus()
    Me.SetFocus
End Sub

Public Function ISandyWindowMain_SetInternalCurrentTemplate(ByVal sTemplateName As String) As Boolean
On Error Resume Next
    Static sCategoryName As String
    Static sShortTemplateName As String

    Err.Clear
    ISandyWindowMain_SaveTemplate

    sCategoryName = sGetToken(sTemplateName, 1, " - ")
    sShortTemplateName = sAfter(sTemplateName, 1, " - ")

    If SliceAndDice(sCategoryName).Templates(sTemplateName) Is Nothing Then
       ISandyWindowMain_SetInternalCurrentTemplate = False
    Else
       Set InternalCurrentTemplate = SliceAndDice(sCategoryName).Templates(sTemplateName)
       ISandyWindowMain_SetInternalCurrentTemplate = True
    End If
End Function

Public Function ISandyWindowMain_sGetCurrentLineAtCharacter(ByVal sTextToSearch As String, ByVal lCharToStart As Long) As String
    Dim lCount As Long
    lCount = lTokenCount(Left$(sTextToSearch, lCharToStart), gsEOL)
    If lCount > 0 Then
       ISandyWindowMain_sGetCurrentLineAtCharacter = sGetToken(sTextToSearch, lCount, gsEOL)
    Else
       ISandyWindowMain_sGetCurrentLineAtCharacter = sTextToSearch
    End If
End Function

Private Sub ISandyWindowMain_Show(Optional ByVal ModalSetting As Integer, Optional ParentWindow As Object)
    If ParentWindow Is Nothing Then
       Me.Show ModalSetting
    Else
       Me.Show ModalSetting, ParentWindow
    End If
End Sub

Public Sub ISandyWindowMain_ShowExternalsMenu()
On Error Resume Next
    PopupMenu mnuExternalFunctions
End Sub

Public Sub ISandyWindowMain_ShowFavMenu()
On Error Resume Next
    PopupMenu mnuFav
End Sub

Public Function ISandyWindowMain_ShutdownDLLs() As Boolean
On Error Resume Next
    Dim CurrSet As Long

    For CurrSet = 1 To SadCommandSetCount
        Call SadCommands(CurrSet).Shutdown
        Set SadCommands(CurrSet) = Nothing
    Next CurrSet
    ReDim SadCommands(1 To 1)
    SadCommandSetCount = 0

    ISandyWindowMain_ShutdownDLLs = True
End Function

Private Property Set ISandyWindowMain_SliceAndDice(ByVal RHS As SandySupport.CSliceAndDice)
    Set SliceAndDice = RHS
End Property

Private Property Get ISandyWindowMain_SliceAndDice() As SandySupport.CSliceAndDice
    Set ISandyWindowMain_SliceAndDice = SliceAndDice
End Property

Public Property Get ISandyWindowMain_TemplateDatabaseName() As String
    ISandyWindowMain_TemplateDatabaseName = m_sTemplateDatabaseName
End Property

Public Sub ISandyWindowMain_SaveTemplate()
On Error GoTo EH_frmMain_SaveTemplate
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Sub
    bInHereAlready = True

    Static CurrBar As MSComctlLib.ListView

    If Not CurrentTemplate Is Nothing Then
       With CurrentTemplate
            If Not .Deleted Then
               If .ParentKey = vbNullString Then
                  MsgBox "frmMain.SaveTemplate : Error found. ParentKey blank"
                  GoTo EH_frmMain_SaveTemplate_Continue
               End If
               txtName.Text = .ParentKey & " - " & txtShortName
               .Key = txtName
               .ShortTemplateName = txtShortName

               .memoCodeAtTop = txtCode(0)
               .memoCodeAtCursor = txtCode(1)
               .memoCodeAtBottom = txtCode(2)

               .FileName = txtFilename
               .memoCodeToFile = txtCodeToFile
    
               .Undeletable = chkUndeletable <> 0
               .Locked = chkLocked <> 0
              '.IncludeInMenu = chkIncludeInMenu <> 0
               .Favorite = chkFavorite <> 0
               .Selected = chkSelected <> 0

'               With SliceAndDice.SystemInfo("Hotkey Templates").Item(.Key)
'                    If hkyInstantInsert.HotKeyAndModifier <> 0 Then
'                       .Value = hkyInstantInsert.HotKey & "," & hkyInstantInsert.HotKeyModifier
'                    Else
'                       .Value = "0,8"
'                    End If
'
'
'               End With

              '.Modified = True
            End If

On Error Resume Next
            If .Modified Then
               SliceAndDice.Save
               Err.Clear
               If (Not lsbJumpTo.Bars(.ParentKey) Is Nothing) And Len(.OriginalShortName) > 0 Then
                  Set CurrBar = lsbJumpTo.Bars(.ParentKey)
                  If Not CurrBar Is Nothing Then
                     If CurrBar.ListItems(.ParentKey & " - " & .OriginalShortName).Text <> .ShortTemplateName Then
                        CurrBar.ListItems(.ParentKey & " - " & .OriginalShortName).Text = .ShortTemplateName
                        Set CurrBar = Nothing
                        mnuFileRefresh_Click
                     End If
                  End If
               End If
               Err.Clear
            End If
       End With
    End If

EH_frmMain_SaveTemplate_Continue:
    bInHereAlready = False
    Exit Sub

EH_frmMain_SaveTemplate:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: SaveTemplate" & vbCr & vbCr & Err.Description
    Resume EH_frmMain_SaveTemplate_Continue

    Resume
End Sub

Public Sub ISandyWindowMain_UpdateFavorites()
On Error Resume Next
    Dim CurrFav As Long
    Dim CurrCategory As CCategory
    Dim CurrTemplate As CTemplate

    DoEvents: DoEvents: DoEvents

    If FavoriteCount > 0 Then ' Clear out previous entries
       For CurrFav = FavoriteCount To 1 Step -1
           Unload mnuFavorite(CurrFav)
       Next CurrFav
       mnuFavorite(0).Caption = "-Empty-"
       mnuFavorite(0).Enabled = False
       FavoriteCount = 0
    End If
    
    For Each CurrCategory In SliceAndDice.Categorys
        For Each CurrTemplate In CurrCategory.Templates
            If CurrTemplate.Favorite Then
               If FavoriteCount > 0 Then
                  Load mnuFavorite(FavoriteCount)
               End If
               mnuFavorite(FavoriteCount).Caption = CurrTemplate.Key
               mnuFavorite(FavoriteCount).Enabled = True
               FavoriteCount = FavoriteCount + 1
            End If
        Next CurrTemplate
    Next CurrCategory
End Sub

Private Function sTemplateIcon(ByVal CurrTemplate As Object) As String
    If CurrTemplate Is Nothing Then
       sTemplateIcon = "!"
    ElseIf Len(CurrTemplate.memoCodeAtBottom & CurrTemplate.memoCodeAtCursor & CurrTemplate.memoCodeAtTop & CurrTemplate.memoCodeToFile) > 0 Then
       sTemplateIcon = "Category"
    ElseIf CurrTemplate.Selected Then
       sTemplateIcon = "Check"
    ElseIf CurrTemplate.Undeletable Or CurrTemplate.Locked Then
       sTemplateIcon = "Key"
    Else
       sTemplateIcon = "Document"
    End If
End Function

Public Sub ISandyWindowMain_UpdateHotKeys()
'On Error GoTo EH_UpdateHotKeys
'    Dim asaTaken As SandySupport.CAssocArray
'    Dim CurrItem As SandySupport.CAssocItem
'
'    Exit Sub
'
'    If SliceAndDice.SystemInfo Is Nothing Then Exit Sub
'
'    Set asaTaken = CreateObject("SandySupport.CAssocArray")
'    asaTaken.Clear
'    asaTaken(vbKeyR & "," & (MOD_CONTROL + MOD_SHIFT)) = "Sandy Repeat Insertion"
'    asaTaken(vbKeyS & "," & (MOD_CONTROL + MOD_SHIFT)) = "Sandy Activate"
'
'    If mHotKeyOpenWindow Is Nothing Then
'       Set mHotKeyOpenWindow = CreateObject("SandyInstance.cRegHotKey")
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
'                   mHotKeyOpenWindow.RegisterKey "TEMPLATE " & CurrItem.Key, Val(sGetToken(CurrItem.Value, 1, ",")), Val(sGetToken(CurrItem.Value, 2, ","))
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
'    LogError "frmMain", "UpdateHotKeys", Err.Number, Err.Description
'    Resume EH_UpdateHotKeys_Continue
'
'    Resume
End Sub

Public Sub chkAutoRecalc_Click()
    SaveSetting "SliceAndDice", "Last", "Auto Recalc", chkAutoRecalc.Value
End Sub

Public Sub chkFavorite_Click()
    If mbFillingAddInScreen Then Exit Sub
    CurrentTemplate.Favorite = (chkFavorite.Value <> 0)
    ISandyWindowMain_UpdateFavorites
End Sub

Public Sub chkUndeletable_Click()
    Dim sPasswordCheck As String
    Static bInHereAlready As Boolean

    If mbFillingAddInScreen Then Exit Sub
    If bInHereAlready Then Exit Sub
    If Not mnuPasswordProtection.Checked Then
       lsbJumpTo.BarItemIcon = IIf(chkLocked.Value = 0 And chkUndeletable.Value = 0, "Document" & IIf(lsbJumpTo.BarItemIcon = "DocumentAlternate", "Alternate", vbNullString), "Key")
       Exit Sub
    End If

    bInHereAlready = True

    If chkUndeletable.Value = 0 Then
       If Len(CurrentTemplate.memoAttributes) Then
          m_asaAttributes.All = CurrentTemplate.memoAttributes
          If Len(m_asaAttributes("Undeletable Password")) Then
             sPasswordCheck = InputBox("Enter password to unlock.", "ENTER PASSWORD")
             If StrComp(sPasswordCheck, m_asaAttributes("Undeletable Password")) <> 0 Then
                Beep
                chkUndeletable.Value = 1
                lsbJumpTo.BarItemIcon = "Key"
             Else
                chkUndeletable.Value = 0
                lsbJumpTo.BarItemIcon = IIf(chkLocked.Value = 0 And chkUndeletable.Value = 0, "Document" & IIf(lsbJumpTo.BarItemIcon = "DocumentAlternate", "Alternate", vbNullString), "Key")
             End If
          End If
       Else
       End If
    Else
       m_asaAttributes.All = CurrentTemplate.memoAttributes
       m_asaAttributes("Undeletable Password") = InputBox("Enter a password for unlocking later.", "ENTER NEW PASSWORD", m_asaAttributes("Undeletable Password"))
       CurrentTemplate.memoAttributes = m_asaAttributes.All
       chkUndeletable.Value = 1
       lsbJumpTo.BarItemIcon = "Key"
    End If
    bInHereAlready = False
End Sub

Public Sub chkUndeletable_Validate(Cancel As Boolean)
    Dim sPasswordCheck As String
End Sub


Public Sub cmdRecalc_Click()
On Error GoTo EH_cmdRecalc_Click

    Dim sCodeToCheck(0 To 4) As String
    Dim CurCodeWindow As Long
    Dim lTokens As Long
    Dim CurToken As Long
    Dim CurListItem As Long
    Dim sCurToken As String

    Screen.MousePointer = vbHourglass
    
    lstSoftVariables.Clear
    lstSoftCommands.Clear
   'MsgBox "Recalc to occur here."
    
    sCodeToCheck(0) = txtCode(1)
    sCodeToCheck(1) = txtCode(0)
    sCodeToCheck(2) = txtCode(2)
    
    With lstSoftVariables
        Select Case UCase$(txtShortName)
               Case "COLLECTION", "COLLECTION, NO CHILD", "COLLECTION, NO PARENT", "COLLECTION, NO PARENT, NO CHILD"
                    If InStr(UCase$(txtShortName), "NO PARENT") = 0 Then
                       .AddItem "* Parent AutoNumber Field Name"
                       .AddItem "* Parent AutoNumber Property Name"
                    End If
                   '.AddItem "* Collection Member Subcollection Property Name = Child Table Name"
                    .AddItem "* Property Name"
                    .AddItem "* Singular Property Name = Child Table Name"
                    .AddItem "* Child Table Name"
                   '.AddItem "* Primary AutoNumber Field for Collection Member = AutoNumber Field"
                    .AddItem "* AutoNumber Field Name"
                    .AddItem "* AutoNumber Property Name"
                   '.AddItem "* Table that stores this collection = Table Name"
                    .AddItem "* Object Name = Table Name"
                    .AddItem "* Table Name"
                   '.AddItem "* Object Name of Collection Member = Object Name"
                    .AddItem "* Spaced Table Name"
                    .AddItem "* Spaced Object Name"
                   '.AddItem "* Label Name of Collection Member = Label Name"
                    .AddItem "* Label Name"
                   '.AddItem "* Field to use as Key = Key Field"
                    .AddItem "* Key Field Name"
                    .AddItem "* Key Property Name"
    
               Case "COLLECTION MEMBER", "COLLECTION MEMBER - TERMINAL"
                   '.AddItem "* Object Name of Collection Member = Object Name"
                    .AddItem "* Object Name = Table Name"
                    .AddItem "* Table Name"
                   '.AddItem "* Label Name of Collection Member = Label Name"
                    .AddItem "* Label Name = Spaced Table Name"
                    .AddItem "* Spaced Table Name"
                   '.AddItem "* Property name of Class to collect"
                    .AddItem "* Class to collect = Property Name"
                    .AddItem "* Property Name"

               Case "COLLECTION MEMBER - NEW SUBCOLLECTION"
                    .AddItem "* Property Name"
                    .AddItem "* Singular Property Name = Child Table Name"
                    .AddItem "* Child Table Name"
                    .AddItem "* Table Name"
                    .AddItem "* Spaced Table Name"
               
               Case "PROPERTY - BLOB", "PROPERTY - BOOLEAN", "PROPERTY - BYTE", "PROPERTY - CURRENCY", "PROPERTY - DATE", "PROPERTY - DOUBLE", "PROPERTY - INTEGER", "PROPERTY - LONG", "PROPERTY - OLE_COLOR", "PROPERTY - SINGLE", "PROPERTY - STRING", "PROPERTY - VARIANT", "PROPERTY - 3D LINK"
                   '.AddItem "* Field Name of Property"
                    .AddItem "* Property Name"
                   '.AddItem "* Pure Field Name"
                    .AddItem "* Field Name"
                    .AddItem "* Table Name"
                    .AddItem "* Spaced Field Name"
                    .AddItem "* Spaced Table Name"
                    
               Case "WRAPPER CLASS", "ROUTINES"
                    .AddItem "* DSN = Database Name"
                    .AddItem "* Database Name"
                    .AddItem "* Database Path"
                    .AddItem "* Spaced Database Name"
                    
               Case "WRAPPER CLASS - ADD COLLECTION"
                   '.AddItem "* Property name of Class to collect"
                    .AddItem "* Property Name"
                   '.AddItem "* Plural Table Name"
                    .AddItem "* Table Name"
                    .AddItem "* Spaced Table Name"
        End Select
    End With
    
    For CurCodeWindow = 0 To 2
      ' Scan for Soft Variables
        lTokens = lTokenCount(sCodeToCheck(CurCodeWindow), gsSoftVarDelimiter)
        If (lTokens Mod 2) = 0 And lTokens > 0 Then ' Even
         ' Token theory clearly states that
         '   If you're using one delimiter for delimiting both
         '      the beginning and ending of a token, then there must be
         '      an ODD number of tokens or the string isn't valid.
           Select Case CurCodeWindow
                  Case 0: MsgBox "You are missing at least one Soft Variable Delimiter (""" & gsSoftVarDelimiter & """) in the 'At cursor' code area of the current template. Discontinuing analysis."
                  Case 1: MsgBox "You are missing at least one Soft Variable Delimiter (""" & gsSoftVarDelimiter & """) in the '(Declarations)' code area of the current template. Discontinuing analysis."
                  Case 2: MsgBox "You are missing at least one Soft Variable Delimiter (""" & gsSoftVarDelimiter & """) in the 'End of Module' code area of the current template. Discontinuing analysis."
                  Case 3: MsgBox "You are missing at least one Soft Variable Delimiter (""" & gsSoftVarDelimiter & """) in the 'In a file' code area of the current template. Discontinuing analysis."
           End Select
           Screen.MousePointer = vbDefault
           Exit Sub
        ElseIf lTokens > 0 Then ' Odd
         ' Okay, keep going
           For CurToken = 2 To lTokens Step 2
              sCurToken = sGetToken(sCodeToCheck(CurCodeWindow), CurToken, gsSoftVarDelimiter)
              If lstSoftVariables.ListCount > 0 Then
                For CurListItem = 0 To lstSoftVariables.ListCount - 1
                    Select Case StrComp(UCase$(lstSoftVariables.List(CurListItem)), UCase$(sCurToken))
                           Case 0
                                CurListItem = lstSoftVariables.ListCount
                                Exit For
                           Case Is > 0
                                lstSoftVariables.AddItem sCurToken, CurListItem
                                CurListItem = CurListItem + 1
                                Exit For
                    End Select
                Next CurListItem
                If StrComp(UCase$(lstSoftVariables.List(CurListItem - 1)), UCase$(sCurToken)) < 0 Then
                   lstSoftVariables.AddItem sCurToken
                End If
              Else
                lstSoftVariables.AddItem sCurToken
              End If
           Next CurToken
        End If

      ' Repeat for Soft Commands
        lTokens = lTokenCount(sCodeToCheck(CurCodeWindow), gsEOL)
        'If (lTokens Mod 2) = 0 And lTokens > 0 Then  ' Even
        ' ' Token theory clearly states that
        ' '   If you're using one delimiter for delimiting both
        ' '      the beginning and ending of a token, then there must be
        ' '      an ODD number of tokens or the string isn't valid.
        '   Select Case CurCodeWindow
        '          Case 0: MsgBox "You are missing at least one Soft Command Delimiter (""" & gsSoftCmdDelimiter & """) in the 'At cursor' code area of the current template. Discontinuing analysis."
        '          Case 1: MsgBox "You are missing at least one Soft Command Delimiter (""" & gsSoftCmdDelimiter & """) in the '(Declarations)' code area of the current template. Discontinuing analysis."
        '          Case 2: MsgBox "You are missing at least one Soft Command Delimiter (""" & gsSoftCmdDelimiter & """) in the 'End of Module' code area of the current template. Discontinuing analysis."
        '          Case 3: MsgBox "You are missing at least one Soft Command Delimiter (""" & gsSoftCmdDelimiter & """) in the 'In a file' code area of the current template. Discontinuing analysis."
        '   End Select
        '   Screen.MousePointer = vbDefault
        '   Exit Sub
        'ElseIf lTokens > 0 Then ' Odd
         ' Okay, keep going
           sCurToken = vbNullString
           For CurToken = 1 To lTokens Step 1
               sCurToken = sGetToken(sGetToken(sCodeToCheck(CurCodeWindow), CurToken, gsEOL), 2, gsSoftCmdDelimiter)
               If sCurToken <> "'" And Len(sCurToken) > 0 Then
                  'If lstSoftCommands.ListCount > 0 Then
                  '   For CurListItem = 0 To lstSoftCommands.ListCount - 1
                  '       Select Case StrComp(UCase$(lstSoftCommands.List(CurListItem)), UCase$(sCurToken))
                  '              Case 0
                  '                   Exit For
                  '              Case Is > 0
                  '                   lstSoftCommands.AddItem sCurToken, CurListItem
                  '                   CurListItem = CurListItem + 1
                  '                   Exit For
                  '       End Select
                  '   Next CurListItem
                  'Else
                     lstSoftCommands.AddItem sCurToken
                  'End If
               End If
           Next CurToken
           If StrComp(UCase$(lstSoftCommands.List(CurListItem - 1)), UCase$(sCurToken)) < 0 Then
              lstSoftCommands.AddItem sCurToken
           End If
        'End If
    Next CurCodeWindow
    
EH_cmdRecalc_Click_Continue:
    Screen.MousePointer = vbDefault
    Exit Sub
    
EH_cmdRecalc_Click:
    Resume EH_cmdRecalc_Click_Continue
    
    Resume
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
       Cancel = True
       mnuFileExit_Click
    Else
       Form_Unload Cancel
    End If

    'If OkayToUnload Then
    '   Form_Unload Cancel
    'Else
    '   Cancel = True
    'End If
End Sub

Private Property Let ISandyWindowMain_Visible(ByVal RHS As Boolean)
    Me.Visible = RHS
End Property

Private Property Get ISandyWindowMain_Visible() As Boolean
    ISandyWindowMain_Visible = Me.Visible
End Property

Private Sub ISandyWindowMain_ZOrder()
    Me.ZOrder
End Sub

Private Sub lsbJumpTo_AfterBarClick()
On Error Resume Next
    ISandyWindowMain_JumpTo lsbJumpTo.BarKey & " - " & lsbJumpTo.BarItemName
End Sub

Public Sub lsbJumpTo_BarItemClick(ByVal BarName As String, ByVal BarKey As String, ByVal BarItemName As String, ByVal BarItemKey As String)
On Error Resume Next
    Dim TemplateFound As CTemplate
    If lsbJumpTo.BarType = "List" Then
       ISandyWindowMain_JumpTo BarItemKey
    Else
       Set TemplateFound = SliceAndDice.Categorys.ItemByLongTemplateName(BarItemKey)
       If TemplateFound Is Nothing Then
          Beep
          If MsgBox("That template does not exist (yet)." & vbCr & vbTab & "Create template now ?", vbYesNo, "NO TEMPLATE: " & BarItemKey) = vbYes Then
             ISandyWindowMain_QueueAction "NewTemplate", BarItemKey
             OkayToDoAction = True
          Else
             If Val(CurrentHistoryEntry) > 0 Then
                ISandyWindowMain_JumpTo m_asaHistory(CurrentHistoryEntry), False, True
             ElseIf SliceAndDice.Categorys(sGetToken(BarItemKey, 1, " - ")).Templates.Count > 1 Then
                ISandyWindowMain_JumpTo SliceAndDice.Categorys(sGetToken(BarItemKey, 1, " - ")).Templates(1)
             End If
          End If
       Else
          ISandyWindowMain_JumpTo BarItemKey
       End If
       Set TemplateFound = Nothing
    End If
End Sub

Public Sub lsbJumpTo_BarItemDblClick(ByVal BarName As String, ByVal BarKey As String, ByVal BarItemName As String, ByVal BarItemKey As String)
    If Len(BarItemKey) = 0 Then Exit Sub

    mnuInsertTemplate_Click
End Sub

Private Sub lsbJumpTo_ItemClick(Item As MSComctlLib.IListItem)
   'If Button = vbRightButton And Shift = 0 Then ' Right click, pop-up menu
       'lsbJumpTo_BarItemClick lsbJumpTo.BarName, Item.Key, Item.Text, Item.Key
       'PopupMenu mnuTemplate
   'End If
End Sub

Public Sub lsbJumpTo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Shift = 0 Then ' Right click, pop-up menu
       PopupMenu mnuTemplate
    End If
End Sub

Public Sub lsbJumpTo_MouseDownOnCategory(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Shift = 0 Then ' Right click, pop-up menu
       PopupMenu mnuCategories
    End If
End Sub

Public Sub lstSoftCommands_DblClick()
On Error Resume Next
    lstSoftVariables.ListIndex = -1
    txtCode(0).SelStart = 0:    txtCode(0).SelLength = 0
    txtCode(1).SelStart = 0:    txtCode(1).SelLength = 0
    txtCode(2).SelStart = 0:    txtCode(2).SelLength = 0
    ISandyWindowMain_FindInCurrent False, False, True
End Sub

Public Sub lstSoftVariables_DblClick()
On Error Resume Next
    lstSoftCommands.ListIndex = -1
    txtCode(0).SelStart = 0:    txtCode(0).SelLength = 0
    txtCode(1).SelStart = 0:    txtCode(1).SelLength = 0
    txtCode(2).SelStart = 0:    txtCode(2).SelLength = 0
    ISandyWindowMain_FindInCurrent False, False, True
End Sub

Public Sub mnuBack_Click()
On Error Resume Next
    If Val(CurrentHistoryEntry) < 2 Then
       Beep
       mnuBack.Enabled = False
       Exit Sub
    End If

    CurrentHistoryEntry = Val(CurrentHistoryEntry) - 1
    ISandyWindowMain_JumpTo m_asaHistory(CurrentHistoryEntry), False, True
    mnuForward.Enabled = True
    mnuBack.Enabled = Val(CurrentHistoryEntry) > 1
End Sub

Public Sub mnuCategoriesDeleteCurrent_Click()
On Error GoTo EH_mnuCategoriesDeleteCurrent_Click
    Static bInHereAlready As Boolean

    If bInHereAlready Then Exit Sub
    bInHereAlready = True

    Dim sCurrentCategory As String
    ISandyWindowMain_SaveTemplate

    sCurrentCategory = lsbJumpTo.BarKey

    If UCase$(sCurrentCategory) = "CHANGE FROM" Then
       MsgBox "The 'Change From' category is not removable.", vbExclamation
       GoTo EH_mnuCategoriesDeleteCurrent_Click_Continue
       Exit Sub
    ElseIf SliceAndDice.Categorys(sCurrentCategory).CategoryType <> 0 Then
       If Not bUserSure("This category is used by the code generators. Deleting it is unadvisable." & gs2EOLTab & "Are you sure you want to permanently erase this category ?") Then
          GoTo EH_mnuCategoriesDeleteCurrent_Click_Continue
          Exit Sub
       End If
    End If

    With SliceAndDice.Categorys(sCurrentCategory)
         If .Templates.Count > 0 Then
            If Not bUserSure("There " & IIf(.Templates.Count = 1, "is", "are") & " " & .Templates.Count & " template" & IIf(.Templates.Count = 1, vbNullString, "s") & " still in the '" & sCurrentCategory & "' category. Continuing will delete all templates in that category." & gs2EOLTab & "Are you absolutely sure this is what you want to do ?") Then
               GoTo EH_mnuCategoriesDeleteCurrent_Click_Continue
               Exit Sub
            End If
         End If
         .Deleted = True
    End With

    SliceAndDice.Save
    ISandyWindowMain_RefillList
    If Not SliceAndDice(1) Is Nothing Then
       If Not SliceAndDice(1).Templates(1) Is Nothing Then
          ISandyWindowMain_JumpTo SliceAndDice(1).Templates(1).Key
          lsbJumpTo.BarAndItem SliceAndDice(1).Key, SliceAndDice(1).Templates(1).ShortTemplateName
       ElseIf Not SliceAndDice(2) Is Nothing Then
          If Not SliceAndDice(2).Templates(1) Is Nothing Then
             ISandyWindowMain_JumpTo SliceAndDice(2).Templates(1).Key
             lsbJumpTo.BarAndItem SliceAndDice(2).Key, SliceAndDice(2).Templates(1).ShortTemplateName
          End If
       End If
    End If

EH_mnuCategoriesDeleteCurrent_Click_Continue:
    bInHereAlready = False
    Exit Sub

EH_mnuCategoriesDeleteCurrent_Click:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: DeleteTemplate" & vbCr & vbCr & Err.Description
    
    Resume EH_mnuCategoriesDeleteCurrent_Click_Continue

    Resume
End Sub

Public Sub mnuCategoriesNewMethod_Click(Index As Integer)
    Dim sNewCategoryName As String
    Dim sCategoryToDuplicate As String

    Select Case Index
           Case 0       ' New, Blank Category
                sNewCategoryName = InputBox("What should the name of the new, blank category be ?", "NEW CATEGORY", vbNullString)
                If Len(sNewCategoryName) = 0 Then Exit Sub
                If SliceAndDice(sNewCategoryName) Is Nothing Then
                   SliceAndDice.Categorys.Add sNewCategoryName
                   SliceAndDice.Save
                   ISandyWindowMain_RefillList
                Else
                   MsgBox "There is already a Category by that name. Aborting.", vbInformation
                End If

           Case 1       ' New, Duplicate Current Category
                sCategoryToDuplicate = SliceAndDice.Categorys.Choose
                If Len(sCategoryToDuplicate) = 0 Then Exit Sub
                sNewCategoryName = InputBox("What should the name of the new, duplicated category be ?", "DUPLICATE CATEGORY", "Copy of " & sCategoryToDuplicate)
                If Len(sNewCategoryName) = 0 Then Exit Sub
                If SliceAndDice(sNewCategoryName) Is Nothing Then
                   SliceAndDice.Categorys.Add sNewCategoryName, , sCategoryToDuplicate
                   SliceAndDice.Save
                   ISandyWindowMain_RefillList 'RefreshDatabaseConnection 'ISandyWindowMain_RefillList
                Else
                   MsgBox "There is already a Category by that name. Aborting.", vbInformation
                End If

           Case 2       ' New, Duplicate Current Category, But don't copy any information from the templates. Only copy names.
                sCategoryToDuplicate = SliceAndDice.Categorys.Choose
                If Len(sCategoryToDuplicate) = 0 Then Exit Sub
                sNewCategoryName = InputBox("What should the name of the new, duplicated category (names, no code) be ?", "DUPE NAMES ONLY", "Copy of " & sCategoryToDuplicate)
                If Len(sNewCategoryName) = 0 Then Exit Sub
                If SliceAndDice(sNewCategoryName) Is Nothing Then
                   SliceAndDice.Categorys.Add sNewCategoryName, , sCategoryToDuplicate, False
                   SliceAndDice.Save
                   ISandyWindowMain_RefillList 'RefreshDatabaseConnection 'ISandyWindowMain_RefillList
                Else
                   MsgBox "There is already a Category by that name. Aborting.", vbInformation
                End If
    End Select
End Sub


Private Sub mnuChangeBackgroundColors_Click()
    Dim ColorSelected As String
    ColorSelected = ISandyWindowMain_sChooseColor(lsbJumpTo.BackColor)
    If Len(ColorSelected) = 0 Then Exit Sub
    
    SaveSetting "SliceAndDice", "Last", "Background Color", ColorSelected
    SetColors ColorSelected, GetSetting("SliceAndDice", "Last", "Foreground Color", "&H80000008&")
End Sub

Private Sub mnuChangeForegroundColor_Click()
    Dim ColorSelected As String
    ColorSelected = ISandyWindowMain_sChooseColor(lsbJumpTo.ForeColor)
    If Len(ColorSelected) = 0 Then Exit Sub
    SaveSetting "SliceAndDice", "Last", "Foreground Color", ColorSelected
    
    SetColors GetSetting("SliceAndDice", "Last", "Background Color", "&H80000018&"), ColorSelected
End Sub

Public Sub mnuEditCopy_Click()
    If Not chkLocked Then
       Select Case tabCode.SelectedItem.Index
              Case 1: Clipboard.SetText txtCode(0).SelText
              Case 2: Clipboard.SetText txtCode(1).SelText
              Case 3: Clipboard.SetText txtCode(2).SelText
              Case 4: Clipboard.SetText txtCode(3).SelText
       End Select
    End If
End Sub

Public Sub mnuEditCut_Click()
    If Not chkLocked Then
       Select Case tabCode.SelectedItem.Index
              Case 1: Clipboard.SetText txtCode(0).SelText: txtCode(0).SelText = vbNullString
              Case 2: Clipboard.SetText txtCode(1).SelText: txtCode(1).SelText = vbNullString
              Case 3: Clipboard.SetText txtCode(2).SelText: txtCode(2).SelText = vbNullString
              Case 4: Clipboard.SetText txtCode(3).SelText: txtCode(3).SelText = vbNullString
       End Select
    End If
End Sub


Public Sub mnuEditFind_Click()
    ISandyWindowMain_FindInCurrent
End Sub


Public Sub mnuEditPaste_Click()
    If Not chkLocked Then
       Select Case tabCode.SelectedItem.Index
              Case 1: txtCode(0).SelText = Clipboard.GetText
              Case 2: txtCode(1).SelText = Clipboard.GetText
              Case 3: txtCode(2).SelText = Clipboard.GetText
              Case 4: txtCode(3).SelText = Clipboard.GetText
       End Select
    Else
       MsgBox "Code areas of current template locked. Unlock template (under the options tab) before attempting this again.", vbInformation
    End If
End Sub


Public Sub mnuEditReplace_Click()
    ISandyWindowMain_FindInCurrent False, True
End Sub


Private Sub mnuExternals_Click(Index As Integer)
On Error Resume Next
    SadCommands(Val(sGetToken(mnuExternals(Index).Tag, 1, "|"))).ExecuteExternal mnuExternals(Index).Caption, sAfter(mnuExternals(Index).Tag, 1, "|")
End Sub

Public Sub mnuFavorite_Click(Index As Integer)
    If FavoriteCalledFromIDE Then
       FavoriteCalledFromIDE = False
       ISandyWindowMain_DoInsertion Nothing, mnuFavorite(Index).Caption
    Else
       ISandyWindowMain_JumpTo mnuFavorite(Index).Caption, , True
       lsbJumpTo.HideCategories
    End If
End Sub

Private Sub mnuFileApplyDeltaPatch_Click()
    Dim sFilename As String
    sFilename = ISandyWindowMain_sChooseFile(, , "Sandy Delta Patch (*.sad)|*.sad|All Files (*.*)|*.*")
    If Len(sFilename) Then
       SliceAndDice.ApplyPatch sFilename
    End If
End Sub

Private Sub mnuFileGenerateDeltaPatch_Click(Index As Integer)
    Dim sDate As String
    Dim PatchFilename As String

    sDate = SliceAndDice.sChoosePatch(Index)
    If Len(sDate) Then
       PatchFilename = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString) & "MDBPatch" & Replace(Format(sDate, "00000.00"), ".", "-") & ".sad"
       SliceAndDice.GenerateDeltaPatchFile CVDate(sDate), PatchFilename
       If Len(Dir$(PatchFilename)) Then
          If bUserSure("File created successfully." & gsEolTab & "Filename:" & PatchFilename & gs2EOL & "Would you like to view it now ?") Then
On Error Resume Next
             Shell sGetWindowsDir & "NOTEPAD.EXE """ & PatchFilename & """"
          End If
       End If
    End If
End Sub

Public Sub mnuForward_Click()
    If Val(CurrentHistoryEntry) >= m_asaHistory.Count Then
       Beep
       mnuForward.Enabled = False
       Exit Sub
    End If
    
    CurrentHistoryEntry = Val(CurrentHistoryEntry) + 1
    ISandyWindowMain_JumpTo m_asaHistory(CurrentHistoryEntry), False, True
    mnuBack.Enabled = True
    mnuForward.Enabled = Val(CurrentHistoryEntry) < m_asaHistory.Count
End Sub

Public Sub mnuHelpAbout_Click()
On Error Resume Next
    With frmSplash
         .lblDLLsLoaded(1).Caption = vbNullString & SadCommandSetCount
         .Show
    End With
End Sub

Private Sub mnuHelpEmailWilliamRawls_Click()
    BrowseTo "mailto:wrawls@firmsolutions.com"
End Sub

Private Sub mnuHelpReportIssue_Click()
    BrowseTo "http://www.sliceanddice.com/sadissue.html"
End Sub

Private Sub mnuHelpSoftCommandReference_Click()
On Error Resume Next
    Dim CurrSet As Long
    Dim sChoices As String
    Dim sChoice As String
    Dim CurrCommand As CSadCommand

    If SadCommandSetCount > 0 Then
       If SadCommandSetCount = 1 Then
          SadCommands(1).CommandSet.ShowHelpScreen
          Exit Sub
       Else
          sChoices = vbNullString
          If Not Complete Is Nothing Then
             Complete.Clear False
             Set Complete = Nothing
          End If
          Set Complete = CreateObject("SandySupport.CSadCommands")
          For CurrSet = 1 To SadCommandSetCount
              If SadCommands(CurrSet).CommandSet.Count > 0 Then
                 For Each CurrCommand In SadCommands(CurrSet).CommandSet
                     Complete.Append CurrCommand
                 Next CurrCommand
              End If
          Next CurrSet
          Set Complete.Parent = Parent
          Complete.ShowHelpScreen
         'Set Complete = Nothing
       End If
    Else
       MsgBox "No command set DLLs loaded." & vbCr & vbTab & "No Soft Command Reference available." & vbCr & vbTab & "Make sure S&D DLLs are in the same directory as the .MDB you have loaded."
    End If
End Sub

Private Sub mnuHelpVisitHomePage_Click()
    BrowseTo "http://www.sliceanddice.com"
End Sub

Private Sub mnuHistoryList_Click()
On Error Resume Next
    Dim sChoices As String
    Dim sChoice As String
    
    If m_asaHistory.Count > 0 Then
       m_asaHistory.ItemDelimiter = ";"
       sChoices = m_asaHistory.Column
       sChoice = sChoose(sChoices, , m_asaHistory(CurrentHistoryEntry).Value)
       If Len(sChoice) Then
          CurrentHistoryEntry = vbNullString & m_asaHistory.FindKey(sChoice)
          ISandyWindowMain_JumpTo m_asaHistory(CurrentHistoryEntry), False, True
          mnuForward.Enabled = Val(CurrentHistoryEntry) < m_asaHistory.Count
          mnuBack.Enabled = Val(CurrentHistoryEntry) > 1
       End If
    End If
    
End Sub

Private Sub mnuHelpOnlineDocumentation_Click()
    BrowseTo "http://www.sliceanddice.com/saddoc.html"
End Sub

Public Sub mnuPasswordProtection_Click()
    mnuPasswordProtection.Checked = Not mnuPasswordProtection.Checked
    SaveSetting "SliceAndDice", "Last", "Password Protection", mnuPasswordProtection.Checked
End Sub

Public Sub mnuShowOnModuleRightClick_Click()
    mnuShowOnModuleRightClick.Checked = Not mnuShowOnModuleRightClick.Checked
    SaveSetting "SliceAndDice", "Last", "Show On Module Right Click", mnuShowOnModuleRightClick.Checked
    MsgBox "This will take effect the next time Visual Basic or Slice and Dice is restarted.", vbInformation
End Sub

Public Sub mnuShowPaintbrushIcon_Click()
    mnuShowPaintbrushIcon.Checked = Not mnuShowPaintbrushIcon.Checked
    SaveSetting "SliceAndDice", "Last", "Show Paitbrush Icon", mnuShowPaintbrushIcon.Checked
    MsgBox "This will take effect the next time Visual Basic is restarted.", vbInformation
End Sub


Public Sub mnuShowSplash_Click()
    mnuShowSplash.Checked = Not mnuShowSplash.Checked
    SaveSetting "SliceAndDice", "Last", "Show Splash", mnuShowSplash.Checked
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
'        Set rst = db.OpenRecordset("SELECT * FROM Templates")
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
'                    MsgBox "Can't export that record for some reason. Probably a Template with that name already exists in the export to database."
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
    mnuSwitchTabsAutomatically.Checked = Not mnuSwitchTabsAutomatically.Checked
    SaveSetting "SliceAndDice", "Last", "Switch tabs automatically", mnuSwitchTabsAutomatically.Checked
End Sub

Public Sub mnuX_Click()
    mnuFileExit_Click
End Sub

'Public Sub tmrActivateDBClassGen_Timer()
'   'If gbProcessing Then Exit Sub
'    tmrActivateDBClassGen.Enabled = False
'
'    m_oDBClassGen.RefreshCategories
'    m_oDBClassGen.Show , Me
'End Sub

Public Sub chkLocked_Click()
    Dim sPasswordCheck As String
    Static bInHereAlready As Boolean

    If mbFillingAddInScreen Then Exit Sub
    If bInHereAlready Then Exit Sub
    If Not mnuPasswordProtection.Checked Then
       lsbJumpTo.BarItemIcon = IIf(chkLocked.Value = 0 And chkUndeletable.Value = 0, "Document" & IIf(lsbJumpTo.BarItemIcon = "DocumentAlternate", "Alternate", vbNullString), "Key")
       Exit Sub
    End If

    bInHereAlready = True

    If chkLocked.Value = 0 Then
       If Len(CurrentTemplate.memoAttributes) Then
          m_asaAttributes.All = CurrentTemplate.memoAttributes
          If Len(m_asaAttributes("Locked Password")) Then
             sPasswordCheck = InputBox("Enter password to unlock.", "ENTER PASSWORD")
             If StrComp(sPasswordCheck, m_asaAttributes("Locked Password")) <> 0 Then
                Beep
                chkLocked.Value = 1
                lsbJumpTo.BarItemIcon = "Key"
             Else
                chkLocked.Value = 0
                lsbJumpTo.BarItemIcon = IIf(chkLocked.Value = 0 And chkUndeletable.Value = 0, "Document" & IIf(lsbJumpTo.BarItemIcon = "DocumentAlternate", "Alternate", vbNullString), "Key")
             End If
          End If
       Else
       End If
    Else
       m_asaAttributes.All = CurrentTemplate.memoAttributes
       m_asaAttributes("Locked Password") = InputBox("Enter a password for unlocking later.", "ENTER NEW PASSWORD", m_asaAttributes("Locked Password"))
       CurrentTemplate.memoAttributes = m_asaAttributes.All
       chkLocked.Value = 1
       lsbJumpTo.BarItemIcon = "Key"
    End If
    
    
    txtCode(0).Enabled = (chkLocked.Value = 0)
    txtCode(1).Enabled = (chkLocked.Value = 0)
    txtCode(2).Enabled = (chkLocked.Value = 0)
    txtName.Enabled = (chkLocked.Value = 0)
    txtShortName.Enabled = (chkLocked.Value = 0)
    txtFilename.Enabled = (chkLocked.Value = 0)
    frmFile.Enabled = (chkLocked.Value = 0)

    bInHereAlready = False
End Sub

Public Sub mnuSpecialNewDatabase_Click()
    Dim sNewDatabaseName As String

    sNewDatabaseName = CreateSandyDatabase(Me.hWnd)
    If Len(sNewDatabaseName) Then
       m_sTemplateDatabaseName = sNewDatabaseName
       ISandyWindowMain_RefreshDatabaseConnection
    End If
End Sub

Public Sub mnuSpecialOpenDatabase_Click()
    Dim sTemplateDatabaseName As String
    Dim sOldDatabaseName As String
    
    sTemplateDatabaseName = ISandyWindowMain_sChooseDatabase(App.Path)
    If Len(sTemplateDatabaseName) Then
       ISandyWindowMain_SaveTemplate
       sOldDatabaseName = m_sTemplateDatabaseName
       m_sTemplateDatabaseName = sTemplateDatabaseName
       If ISandyWindowMain_RefreshDatabaseConnection Then
          SaveSetting App.ProductName, "Settings", "Current Database", sTemplateDatabaseName
          If Not SliceAndDice(1) Is Nothing Then
             If Not SliceAndDice(1).Templates(1) Is Nothing Then
                ISandyWindowMain_JumpTo SliceAndDice(1).Templates(1).Key
                lsbJumpTo.BarAndItem SliceAndDice(1).Key, SliceAndDice(1).Templates(1).ShortTemplateName
             ElseIf Not SliceAndDice(2) Is Nothing Then
                If Not SliceAndDice(2).Templates(1) Is Nothing Then
                   ISandyWindowMain_JumpTo SliceAndDice(2).Templates(1).Key
                   lsbJumpTo.BarAndItem SliceAndDice(2).Key, SliceAndDice(2).Templates(1).ShortTemplateName
                End If
             End If
          End If
       Else
          m_sTemplateDatabaseName = sOldDatabaseName
       End If
    End If
End Sub

Public Sub tabCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo EH_frmMain_tabCode_MouseUp
    Static bInHereAlready As Boolean
    'If bInHereAlready Then Exit Sub
    'bInHereAlready = True

    Select Case tabCode.SelectedItem.Index
           Case 1
                txtCode(0).Visible = True
                txtCode(1).Visible = False
                txtCode(2).Visible = False
                frmFile.Visible = False
                frmOptions.Visible = False
                frmTemplateInfo.Visible = False
                'If Not VBIDEWindow Is Nothing Then
                '   If VBIDEWindow.Visible Then
                '      If txtCode(0).Enabled Then
                '         txtCode(0).SetFocus
                '      End If
                '   End If
                'End If
           Case 2
                txtCode(0).Visible = False
                txtCode(1).Visible = True
                txtCode(2).Visible = False
                frmFile.Visible = False
                frmOptions.Visible = False
                frmTemplateInfo.Visible = False
                'If Not VBIDEWindow Is Nothing Then
                '   If VBIDEWindow.Visible Then
                '      If txtCode(1).Enabled Then
                '         txtCode(1).SetFocus
                '      End If
                '   End If
                'End If
           Case 3
                txtCode(0).Visible = False
                txtCode(1).Visible = False
                txtCode(2).Visible = True
                frmFile.Visible = False
                frmOptions.Visible = False
                frmTemplateInfo.Visible = False
                'If Not VBIDEWindow Is Nothing Then
                '   If VBIDEWindow.Visible Then
                '      If txtCode(2).Enabled Then
                '         txtCode(2).SetFocus
                '      End If
                '   End If
                'End If
           Case 4
                txtCode(0).Visible = False
                txtCode(1).Visible = False
                txtCode(2).Visible = False
                frmFile.Visible = True
                frmOptions.Visible = False
                frmTemplateInfo.Visible = False
                'If Not VBIDEWindow Is Nothing Then
                '   If VBIDEWindow.Visible Then
                '      If txtCodeToFile.Enabled Then
                '         txtCodeToFile.SetFocus
                '      End If
                '   End If
                'End If
           Case 5
                txtCode(0).Visible = False
                txtCode(1).Visible = False
                txtCode(2).Visible = False
                frmFile.Visible = False
                frmOptions.Visible = True
                frmTemplateInfo.Visible = False
                'If Not VBIDEWindow Is Nothing Then
                '   If VBIDEWindow.Visible Then
                '      If txtShortName.Enabled Then
                '         txtShortName.SetFocus
                '      End If
                '   End If
                'End If
           Case 6
                txtCode(0).Visible = False
                txtCode(1).Visible = False
                txtCode(2).Visible = False
                frmFile.Visible = False
                frmOptions.Visible = False
                frmTemplateInfo.Visible = True
                'If Not VBIDEWindow Is Nothing Then
                '   If VBIDEWindow.Visible Then
                '      If cmdRecalc.Enabled Then
                '         cmdRecalc.SetFocus
                '      End If
                '   End If
                'End If
                If chkAutoRecalc.Value <> 0 Then
                   cmdRecalc_Click
                End If
    End Select

EH_frmMain_tabCode_MouseUp_Continue:
    bInHereAlready = False
    Exit Sub

EH_frmMain_tabCode_MouseUp:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: tabCode_MouseUp" & vbCr & vbCr & Err.Description
    Resume EH_frmMain_tabCode_MouseUp_Continue

    Resume
End Sub

Private Sub tmrDoAction_Timer()
On Error Resume Next
    If Not OkayToDoAction Then Exit Sub

    tmrDoAction.Enabled = False

    Select Case UCase$(ActionToDo)
           Case "NEWTEMPLATE"
                ISandyWindowMain_NewTemplate True, ActionParam

           Case "DELTACHECK", "DELTA CHECK"
                If Len(Dir$(Parent.TemplateDatabasePath & "MDBPatch*.sad", vbNormal)) Then
                   If bUserSure("A Delta Patch file has been found. Would you like to apply it now ?") Then
                      SliceAndDice.ApplyPatch Dir$(Parent.TemplateDatabasePath & "MDBPatch*.sad", vbNormal)
                   End If
                End If
    End Select

    ISandyWindowMain_QueueAction "DeltaCheck", vbNullString, 65535
    OkayToDoAction = True
End Sub

Public Sub txtCode_GotFocus(Index As Integer)
    CurrentCodeArea = Index
End Sub

Public Sub txtCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     Form_KeyDown KeyCode, Shift
End Sub


' ********************************************************************************
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
' ********************************************************************************
Public Sub Form_GotFocus()
On Error Resume Next
    lsbJumpTo.SetFocus      ' More than likely the user is going to want to insert a pre-existing Template.
End Sub

' ********************************************************************************
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
' ********************************************************************************
Public Sub Form_Initialize()
    Dim sLastTemplate As String
    Dim sCategory As String
    Dim sShortName As String

    mnuExitAfterInsert.Checked = GetSetting("SliceAndDice", "Settings", "Exit after insert", True)
    mnuShowPaintbrushIcon.Checked = GetSetting("SliceAndDice", "Last", "Show Paitbrush Icon", True)
    mnuShowOnModuleRightClick.Checked = GetSetting("SliceAndDice", "Last", "Show On Module Right Click", True)
    mnuSwitchTabsAutomatically.Checked = GetSetting("SliceAndDice", "Last", "Switch tabs automatically", True)
    mnuPasswordProtection.Checked = GetSetting("SliceAndDice", "Last", "Password Protection", False)
    
    chkAutoRecalc.Value = GetSetting("SliceAndDice", "Last", "Auto Recalc", 0)

    lsbJumpTo.Arrange = GetSetting("SliceAndDice", "Settings", "Bar Arrange", "1")
    lsbJumpTo.View = GetSetting("SliceAndDice", "Settings", "Bar View", "1")

    If Dir$(App.Path & "\SliceAndDice.mdb") = vbNullString And Dir$(App.Path & "\SliceAndDiceNew.mdb") <> vbNullString Then
       Name App.Path & "\SliceAndDiceNew.mdb" As App.Path & "\SliceAndDice.mdb"
    End If
    m_sTemplateDatabaseName = GetSetting("SliceAndDice", "Settings", "Current Database")
    If Len(Dir$(m_sTemplateDatabaseName)) = 0 Then
       m_sTemplateDatabaseName = vbNullString
    End If
UserDoc_Init_Try_Again:
    If Len(m_sTemplateDatabaseName) = 0 Then
       If Dir$(App.Path & "\SliceAndDice.mdb") = vbNullString Then
          m_sTemplateDatabaseName = ISandyWindowMain_sChooseDatabase(App.Path, "SliceAndDice.mdb")
       Else
          m_sTemplateDatabaseName = App.Path & "\SliceAndDice.mdb"
       End If

       If Len(m_sTemplateDatabaseName) = 0 Then
          MsgBox "A Template database must be chosen. Please try again."
          GoTo UserDoc_Init_Try_Again
       Else
          SaveSetting "SliceAndDice", "Settings", "Current Database", m_sTemplateDatabaseName
       End If
    End If

    If Not ISandyWindowMain_RefreshDatabaseConnection Then
       m_sTemplateDatabaseName = vbNullString
       GoTo UserDoc_Init_Try_Again
    End If

    Form_Resize                                                 ' Force redraw to make sure everything looks good
    
    sLastTemplate = GetSetting("SliceAndDice", "Settings", "Last Template")
    If Len(sLastTemplate) = 0 Then
       sLastTemplate = "Release Notes - Welcome"
    End If

On Error Resume Next
    ISandyWindowMain_GetCategoryAndName sLastTemplate, sCategory, sShortName
    ISandyWindowMain_JumpTo sLastTemplate, False, True
    DoEvents
    lsbJumpTo.DisplayCategories

    Err.Clear

    lsbJumpTo.DisplayCategories
    
    ' LogEvent "frmMain: Initialize"
End Sub

Public Function ISandyWindowMain_sChooseDatabase(Optional ByVal sPath As String, Optional ByVal sFilename As String) As String
On Error Resume Next
    Err.Clear
    With cdgSelect
         .Filter = "Access Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
         .FilterIndex = 0
         If Len(sPath) > 0 Then .InitDir = sPath
         If Len(sFilename) > 0 Then .FileName = sFilename
         .ShowOpen
         If Err <> 0 Then
            Err.Clear
            Exit Function
         End If
         ISandyWindowMain_sChooseDatabase = .FileName
    End With
End Function

Public Function ISandyWindowMain_sChooseFile(Optional ByVal sPath As String, Optional ByVal sFilename As String, Optional ByVal sFilter As String = vbNullString) As String
On Error Resume Next
    Err.Clear
    With cdgSelect
         .Filter = IIf(Len(sFilter) And InStr(sFilter, "|"), sFilter, "All Files (*.*)|*.*")
         .FilterIndex = 0
         If Len(sPath) > 0 Then .InitDir = sPath
         If Len(sFilename) > 0 Then .FileName = sFilename
         .ShowOpen
         If Err <> 0 Then
            Err.Clear
            Exit Function
         End If
         ISandyWindowMain_sChooseFile = .FileName
    End With
End Function

Public Function ISandyWindowMain_sChooseColor(Optional ByVal sInitialColor As String) As String
On Error Resume Next
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer

    Err.Clear
    With cdgSelect
         .CancelError = True
         If lTokenCount(sInitialColor, ";") = 3 Then
            Red = sGetToken(sInitialColor, 1, ";")
            Green = sGetToken(sInitialColor, 1, ";")
            Blue = sGetToken(sInitialColor, 1, ";")
            If Red > 255 Then Red = 255
            If Red < 0 Then Red = 0
            If Green > 255 Then Green = 255
            If Green < 0 Then Green = 0
            If Blue > 255 Then Blue = 255
            If Blue < 0 Then Blue = 0
            .Color = RGB(Red, Green, Blue)
        ElseIf Val(sInitialColor) > 0 Then
            .Color = Val(sInitialColor)
        End If
On Error GoTo ErrHandler
         .Flags = cdlCCRGBInit
         .ShowColor
         ISandyWindowMain_sChooseColor = Hex(.Color)
    End With
ErrHandler:
   ' User pressed Cancel button.
   Exit Function
End Function

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo EH_frmMain_Form_KeyDown
    Dim sSelectedText As String
    Dim sOrigText As String
    Dim CurrSet As Long
    
    If (Shift And vbShiftMask) > 0 Then     ' Shift Key
        Select Case KeyCode
               Case vbKeyInsert: KeyCode = 0: Shift = 0         ' Paste
                    ActiveControl.SelText = Clipboard.GetText
               Case vbKeyDelete: KeyCode = 0: Shift = 0         ' Cut
                    Err.Clear
                    Clipboard.SetText ActiveControl.SelText
                    If Err.Number = 0 Then
                       ActiveControl.SelText = vbNullString
                    End If
        End Select
    ElseIf (Shift And vbCtrlMask) > 0 Then   ' Control Key
        Select Case KeyCode
               Case vbKeyInsert: KeyCode = 0: Shift = 0         ' Copy
                    Clipboard.SetText ActiveControl.SelText
               Case vbKeyDelete: KeyCode = 0: Shift = 0         ' Cut
                    Err.Clear
                    Clipboard.SetText ActiveControl.SelText
                    If Err.Number = 0 Then
                       ActiveControl.SelText = vbNullString
                    End If
               Case vbKeyTab
On Error Resume Next
                    KeyCode = 0: Shift = 0
                    Parent.SandyIDE.ActiveWindow.SetFocus

              'Case vbKeyC: KeyCode = 0: Shift = 0: mnuFileCopy_Click
               Case vbKeyF: KeyCode = 0: Shift = 0: ISandyWindowMain_FindInCurrent
               Case vbKeyH: KeyCode = 0: Shift = 0: ISandyWindowMain_FindInCurrent False, True
               Case vbKeyI: KeyCode = 0: Shift = 0: mnuInsertTemplate_Click
               Case vbKeyL: KeyCode = 0: Shift = 0: mnuFileRefresh_Click
               Case vbKeyM: KeyCode = 0: Shift = 0: mnuFileImport_Click
               Case vbKeyN: KeyCode = 0: Shift = 0: mnuFileNew_Click
               
              'Case vbKey1: KeyCode = 0: Shift = 0: txtShortName.SetFocus
              'Case vbKey2: KeyCode = 0: Shift = 0: tabCode.Tabs(1).Selected = True: tabCode_MouseUp 0, 0, 0, 0
              'Case vbKey3: KeyCode = 0: Shift = 0: tabCode.Tabs(2).Selected = True: tabCode_MouseUp 0, 0, 0, 0
              'Case vbKey4: KeyCode = 0: Shift = 0: tabCode.Tabs(3).Selected = True: tabCode_MouseUp 0, 0, 0, 0
              'Case vbKey5: KeyCode = 0: Shift = 0: tabCode.Tabs(4).Selected = True: tabCode_MouseUp 0, 0, 0, 0
              'Case vbKey6: KeyCode = 0: Shift = 0: tabCode.Tabs(5).Selected = True: tabCode_MouseUp 0, 0, 0, 0
              'Case vbKey7: KeyCode = 0: Shift = 0: lsbJumpTo.SetFocus
        End Select
    ElseIf (Shift And vbAltMask) > 0 Then     ' Alt Key
        Select Case KeyCode
               Case vbKeyX: KeyCode = 0: Shift = 0: mnuFileExit_Click
               Case vbKeyPageUp
               Case vbKeyPageDown
               Case vbKeyTab
On Error Resume Next
                    KeyCode = 0: Shift = 0
                    Parent.SandyIDE.ActiveWindow.SetFocus
                    'MsgBox "How to get back to VB IDE from here ?"
        End Select
    Else 'If (Shift And vbShiftMask) = 0 Then     ' No special modifying keys
        Select Case KeyCode
               Case vbKeyTab
                    If InStr(ActiveControl.Tag, "Code Area ") Then
                       ActiveControl.SelText = vbTab
                       KeyCode = 0
                       Shift = 0
                    End If

               Case vbKeyF1
                    KeyCode = 0
                    Shift = 0
                    sOrigText = Trim$(ActiveControl.SelText)
                    If Len(sOrigText) = 0 Then
                       sOrigText = ISandyWindowMain_sGetCurrentLineAtCharacter(ActiveControl.Text, ActiveControl.SelStart)
                    End If
                    sSelectedText = sOrigText
                    If Len(sSelectedText) > 0 Then
                       If InStr(sSelectedText, "~~") Then
                          sSelectedText = UCase$(sGetToken(sGetToken(sSelectedText, 2, "~~"), 1)) & "*C"
                       End If
                       If InStr(sSelectedText, "%%") Then
                          sSelectedText = sGetToken(sGetToken(sSelectedText, 2, "%%"), 1, "::") & "*I"
                       End If
                       If SadCommandSetCount > 0 Then
                          For CurrSet = 1 To SadCommandSetCount
                              If Not SadCommands(CurrSet).CommandSet.Item(sSelectedText) Is Nothing Then
                                 SadCommands(CurrSet).CommandSet.ShowHelpScreen sSelectedText
                                'With frmCommandHelp
                                '     Set .SadCommandSet = SadCommands(CurrSet).CommandSet
                                '     .CurrCommandKey = .SadCommandSet(sSelectedText).Index
                                '     .Show 0, Me
                                'End With
                                 GoTo EH_frmMain_Form_KeyDown_Continue
                              End If
                          Next CurrSet
                          MsgBox "SoftCommand '" & Trim$(sOrigText) & "' not found."
                       Else
                          MsgBox "No command set DLLs loaded. No help available."
                       End If
                    Else
                       MsgBox "Unable to determine which command to display help for."
                    End If
                             
               Case vbKeyF3: KeyCode = 0: Shift = 0: ISandyWindowMain_FindInCurrent True
        End Select
    End If

EH_frmMain_Form_KeyDown_Continue:
    Exit Sub

EH_frmMain_Form_KeyDown:
    LogError "frmMain", "Form_KeyDown", Err.Number, Err.Description
    Resume EH_frmMain_Form_KeyDown_Continue

    Resume
End Sub

Public Sub ISandyWindowMain_FindInCurrent(Optional ByVal bRepeatLastSearch As Boolean = False, Optional ByVal bReplace As Boolean = False, Optional ByVal bAuto As Boolean = False)
On Error GoTo EH_frmMain_FindInCurrent
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Sub
    bInHereAlready = True

    Dim CurCodeArea As Long
    Dim lLastFound As Long

    If CurrentTemplate Is Nothing Then
       MsgBox "Please select a template before selecting to search."
       Exit Sub
    End If

    With frmFindReplace
         Select Case tabCode.SelectedItem.Index
                Case 1 To 4
                     .txtFind = txtCode(Me.CurrentCodeArea).SelText
                     .txtReplace = vbNullString
                Case 6
                     If lstSoftVariables.ListIndex > -1 Then
                        .txtFind = Replace(lstSoftVariables, "* ", vbNullString)
                        .txtReplace = vbNullString
                     ElseIf lstSoftCommands.ListIndex > -1 Then
                        .txtFind = lstSoftCommands
                        .txtReplace = vbNullString
                     End If
         End Select

         If bAuto Then
            .DoFindNext = True
            .DoReplace = False
            .DoReplaceAll = False
            .Canceled = False
         ElseIf Not bRepeatLastSearch Then
            .IReplace = bReplace
         ElseIf bReplace Then
            .DoReplace = True
         Else
            .DoFindNext = True
         End If

         If .DoReplaceAll Then
            Screen.MousePointer = vbHourglass
                ISandyWindowMain_SaveTemplate
                    Select Case .SearchArea
                           Case SearchAreaCurrentPane:      txtCode(CurrentCodeArea).Text = Replace(txtCode(CurrentCodeArea).Text, .txtFind, .txtReplace)
                           Case SearchAreaCurrentTemplate:  CurrentTemplate.Replace .txtFind, .txtReplace
                           Case SearchAreaCurrentCategory:  SliceAndDice.Categorys(CurrentTemplate.ParentKey).Replace .txtFind, .txtReplace
                           Case SearchAreaCurrentDatabase:  SliceAndDice.Categorys.Replace .txtFind, .txtReplace
                    End Select
                SliceAndDice.Save
                ISandyWindowMain_FillAddInScreen
            Screen.MousePointer = vbDefault
         ElseIf .DoFindNext Or .DoReplace Then
            For CurCodeArea = 0 To 2
                lLastFound = txtCode(CurCodeArea).SelStart + txtCode(CurCodeArea).SelLength
                If lLastFound = 0 Then lLastFound = 1
                If .chkMatchCase.Value <> 0 Then
                   lLastFound = InStr(Mid$(txtCode(CurCodeArea), lLastFound), .txtFind)
                Else
                   lLastFound = InStr(UCase$(Mid$(txtCode(CurCodeArea), lLastFound)), UCase$(.txtFind))
                End If

                If lLastFound > 0 Then
                   With txtCode(CurCodeArea)
                        tabCode.Tabs(CurCodeArea + 1).Selected = True
                        tabCode_MouseUp 0, 0, 0, 0
On Error Resume Next
                        .SetFocus
                        lLastFound = lLastFound - 2 + IIf(.SelStart + .SelLength = 0, 1, .SelStart + .SelLength)
                        .SelStart = lLastFound
                        .SelLength = Len(frmFindReplace.txtFind)
                        If frmFindReplace.DoReplace Then
                           .SelText = frmFindReplace.txtReplace
                        End If
                   End With
                End If
            Next CurCodeArea
         End If
    End With

EH_frmMain_FindInCurrent_Continue:
    bInHereAlready = False
    Exit Sub

EH_frmMain_FindInCurrent:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: FindInCurrent" & vbCr & vbCr & Err.Description
    Resume EH_frmMain_FindInCurrent_Continue

    Resume
End Sub

' ********************************************************************************
' Name              Form_Resize
'
' Parameters
'      None
'
' Description
'
' This code makes sure everything looks good after a form resize.
'
' ********************************************************************************
Public Sub Form_Resize()
On Error GoTo EH_Form_Resize
    With tabCode                                            ' Position the code entry areas
         .Height = ScaleHeight - .Top
        'lsbJumpTo.Height = ScaleHeight - 415
         If ScaleWidth - .Left < 0 Then Exit Sub            ' If there isn't enough display area to show the code entry areas, don't attempt to redraw it
         .Width = ScaleWidth - .Left
         txtName.Move lblCode(3).Left + lblCode(3).Width + 40, txtName.Top, .Width - (lblCode(3).Left + lblCode(3).Width + 40 - .Left), txtName.Height
         txtShortName.Move lblCode(3).Left + lblCode(3).Width + 40, txtName.Top, .Width - (lblCode(3).Left + lblCode(3).Width + 40 - .Left), txtName.Height

         txtCode(0).Move .Left + 100, .Top + 500, .Width - 200, .Height - 600
         txtCode(1).Move .Left + 100, .Top + 500, .Width - 200, .Height - 600
         txtCode(2).Move .Left + 100, .Top + 500, .Width - 200, .Height - 600
         frmOptions.Move .Left + 100, .Top + 500, .Width - 200, .Height - 600
         frmFile.Move .Left + 100, .Top + 500, .Width - 200, .Height - 600
            txtFilename.Width = frmFile.Width - txtFilename.Left * 2
            txtCodeToFile.Width = txtFilename.Width
            txtCodeToFile.Height = frmFile.Height - txtCodeToFile.Top - 100
         frmTemplateInfo.Move .Left + 100, .Top + 500, .Width - 200, .Height - 600
            lstSoftVariables.Width = (frmTemplateInfo.Width - lstSoftVariables.Left * 3) \ 2
            lstSoftCommands.Left = lstSoftVariables.Left * 2 + lstSoftVariables.Width
            lstSoftCommands.Width = lstSoftVariables.Width
            lblTemplateInfo(0).Left = lstSoftVariables.Left
            lblTemplateInfo(1).Left = lstSoftCommands.Left
            lstSoftVariables.Height = frmTemplateInfo.Height - lstSoftVariables.Top - 100
            lstSoftCommands.Height = frmTemplateInfo.Height - lstSoftVariables.Top - 100
    End With
    
EH_Form_Resize_Continue:
    Exit Sub
    
EH_Form_Resize:
    Resume EH_Form_Resize_Continue:
    
    Resume
End Sub

Public Sub mnuDBClassGen_Click()
    tmrActivateDBClassGen.Enabled = True
End Sub

Public Function ISandyWindowMain_sPropertyType(sFieldType As String) As String
       Select Case sFieldType
              Case "Big Integer":               ISandyWindowMain_sPropertyType = "Long"
              Case "Binary":                    ISandyWindowMain_sPropertyType = "Variant"
              Case "Boolean":                   ISandyWindowMain_sPropertyType = "Boolean"
              Case "Byte":                      ISandyWindowMain_sPropertyType = "Byte"
              Case "Char":                      ISandyWindowMain_sPropertyType = "String"
              Case "Currency":                  ISandyWindowMain_sPropertyType = "Currency"
              Case "Date / Time":               ISandyWindowMain_sPropertyType = "Date"
              Case "Decimal":                   ISandyWindowMain_sPropertyType = "Variant"
              Case "Double":                    ISandyWindowMain_sPropertyType = "Double"
              Case "Float":                     ISandyWindowMain_sPropertyType = "Double"
              Case "Guid":                      ISandyWindowMain_sPropertyType = "String"
              Case "Integer":                   ISandyWindowMain_sPropertyType = "Integer"
              Case "Long":                      ISandyWindowMain_sPropertyType = "Long"
              Case "Long Binary (OLE Object)":  ISandyWindowMain_sPropertyType = "Variant"
              Case "Memo":                      ISandyWindowMain_sPropertyType = "Memo"
              Case "Numeric":                   ISandyWindowMain_sPropertyType = "Variant"
              Case "Single":                    ISandyWindowMain_sPropertyType = "Single"
              Case "Text":                      ISandyWindowMain_sPropertyType = "String"
              Case "Time":                      ISandyWindowMain_sPropertyType = "Date"
              Case "Time Stamp":                ISandyWindowMain_sPropertyType = "Date"
              Case "VarBinary":                 ISandyWindowMain_sPropertyType = "Variant"
              Case Else:                        ISandyWindowMain_sPropertyType = "Variant"
        End Select
End Function


' ********************************************************************************
' Name              frmMain_mnuExitAfterInsert_Click
'
' Parameters
'      None
'
' Description
'
' Toggle the menu item's checked state
'
' ********************************************************************************
Public Sub mnuExitAfterInsert_Click()
    mnuExitAfterInsert.Checked = Not mnuExitAfterInsert.Checked
End Sub

' ********************************************************************************
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
' ********************************************************************************
Public Sub mnuFileCopy_Click()
    Dim sCategory As String
    Dim sShortName As String
    Dim sName As String
    Dim sCode(0 To 2) As String
    

    If CurrentTemplate Is Nothing Then
       MsgBox "Please select a template to copy before selecting this option."
       Exit Sub
    End If

    sName = txtName.Text               ' Save the contents of the Template to copy
    ISandyWindowMain_GetCategoryAndName sName, sCategory, sShortName
    sCode(0) = txtCode(0).Text
    sCode(1) = txtCode(1).Text
    sCode(2) = txtCode(2).Text

On Error Resume Next

    ISandyWindowMain_NewTemplate True, sCategory & " - Copy of " & sShortName

    txtCode(0).Text = sCode(0)         ' Paste in the code from the template to copy
    txtCode(1).Text = sCode(1)
    txtCode(2).Text = sCode(2)

End Sub

' ********************************************************************************
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
' ********************************************************************************
Public Sub mnuFileDelete_Click()
    ISandyWindowMain_DeleteTemplate
End Sub

' ********************************************************************************
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
' ********************************************************************************
Public Sub mnuFileExit_Click()
    ISandyWindowMain_SaveTemplate

    Hide
    ISandyWindowMain_HideAllWindows
   'VBIDEWindow.Visible = False      '   So hiding it will return control to VB
End Sub

' ********************************************************************************
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
' ********************************************************************************
Public Sub mnuFileImport_Click()
    Dim lLine As Long
    Dim lLastLine As Long
    Dim lTemp As Long
    Dim lFirstCol As Long
    Dim lLastCol As Long
    Dim sCode As String
    Dim sProcName As String
    Dim lProcType As Long

On Error Resume Next
    With Parent.SandyIDE.ActiveCodePane                                 ' Send all output to the active code pane
         .GetSelection lLine, lFirstCol, lLastLine, lLastCol    ' Determine where the cursor is
         lLastLine = lLastLine - lLine + Abs(lLastCol > 1)         ' Determine what the last line selected is (discard last line if at beginning)
         sCode = .CodeModule.Lines(lLine, lLastLine)            ' Grab the code selected from the active pane
         ISandyWindowMain_GetProcAtLine lLine, sProcName, lProcType
    End With

    If Len(sCode) = 0 Then
       MsgBox "Nothing to import.", vbInformation
       Exit Sub
    End If

    If sProcName <> vbNullString Then
       ISandyWindowMain_NewTemplate , , sProcName
    Else
       ISandyWindowMain_NewTemplate
    End If

    txtCode(2).Text = sCode
    tabCode.Tabs(2).Selected = True
    tabCode_MouseUp 0, 0, 0, 0
End Sub

Public Function ISandyWindowMain_GetCurrentTextSelection() As String
    Dim lLine As Long
    Dim lLastLine As Long
    Dim lFirstCol As Long
    Dim lLastCol As Long
    Dim sCode As String

On Error Resume Next
    With Parent.SandyIDE.ActiveCodePane                                 ' Send all output to the active code pane
         .GetSelection lLine, lFirstCol, lLastLine, lLastCol    ' Determine where the cursor is
         lLastLine = lLastLine - lLine + Abs(lLastCol > 1)         ' Determine what the last line selected is (discard last line if at beginning)
         ISandyWindowMain_GetCurrentTextSelection = .CodeModule.Lines(lLine, lLastLine)           ' Grab the code selected from the active pane
    End With
End Function

Public Sub ISandyWindowMain_DeleteCurrentTextSelection()
    Dim lLine As Long
    Dim lLastLine As Long
    Dim lFirstCol As Long
    Dim lLastCol As Long
    Dim sCode As String

On Error Resume Next
    With Parent.SandyIDE.ActiveCodePane                                 ' Send all output to the active code pane
         .GetSelection lLine, lFirstCol, lLastLine, lLastCol
         lLastLine = lLastLine - lLine + Abs(lLastCol > 1)
         .CodeModule.DeleteLines lLine, lLastLine
    End With
End Sub

Public Function ISandyWindowMain_DetermineLastLineInSelection() As Long
    Dim lLine As Long
    Dim lLastLine As Long
    Dim lFirstCol As Long
    Dim lLastCol As Long
    Dim sCode As String

On Error Resume Next
    With Parent.SandyIDE.ActiveCodePane                                 ' Send all output to the active code pane
         .GetSelection lLine, lFirstCol, lLastLine, lLastCol
         ISandyWindowMain_DetermineLastLineInSelection = lLastLine
    End With
End Function

Public Function ISandyWindowMain_DetermineFirstLineInSelection() As Long
    Dim lLine As Long
    Dim lLastLine As Long
    Dim lFirstCol As Long
    Dim lLastCol As Long
    Dim sCode As String

On Error Resume Next
    With Parent.SandyIDE.ActiveCodePane                                 ' Send all output to the active code pane
         .GetSelection lLine, lFirstCol, lLastLine, lLastCol
         ISandyWindowMain_DetermineFirstLineInSelection = lLine
    End With
End Function

Public Function ISandyWindowMain_DetermineFirstColumnInSelection() As Long
    Dim lLine As Long
    Dim lLastLine As Long
    Dim lFirstCol As Long
    Dim lLastCol As Long
    Dim sCode As String

On Error Resume Next
    With Parent.SandyIDE.ActiveCodePane                                 ' Send all output to the active code pane
         .GetSelection lLine, lFirstCol, lLastLine, lLastCol
         ISandyWindowMain_DetermineFirstColumnInSelection = lFirstCol
    End With
End Function

Public Function ISandyWindowMain_DetermineLastColumnInSelection() As Long
    Dim lLine As Long
    Dim lLastLine As Long
    Dim lFirstCol As Long
    Dim lLastCol As Long
    Dim sCode As String

On Error Resume Next
    With Parent.SandyIDE.ActiveCodePane                                 ' Send all output to the active code pane
         .GetSelection lLine, lFirstCol, lLastLine, lLastCol
         ISandyWindowMain_DetermineLastColumnInSelection = lLastCol
    End With
End Function

' ********************************************************************************
' Name              frmMain_mnuFileNew_Click
'
' Parameters
'      None
'
' Description
'
' Inserts a new template record.
'
' ********************************************************************************
Public Sub mnuFileNew_Click()
    ISandyWindowMain_NewTemplate
End Sub

' ********************************************************************************
' Name              frmMain_mnuFileRefresh_Click
'
' Parameters
'      None
'
' Description
'
' Refreshes the list of templates
'
' ********************************************************************************
Public Sub mnuFileRefresh_Click()
On Error Resume Next
    Dim sTitle As String

    sTitle = lsbJumpTo.BarKey & " - " & lsbJumpTo.BarItemName
    ISandyWindowMain_RefillList
    ISandyWindowMain_JumpTo sTitle, False, True
End Sub


Public Sub mnuInsertTemplate_Click()
    ISandyWindowMain_DoInsertion Nothing, txtName
End Sub

Public Sub Form_Terminate()
    Dim Cancel As Integer
    Form_Unload Cancel
    ' LogEvent "frmMain: Terminate"
End Sub

Public Sub Form_Load()
    m_asaHistory.Clear
    LoadFormPosition Me
    SetColors GetSetting("SliceAndDice", "Last", "Background Color", "&H80000018&"), GetSetting("SliceAndDice", "Last", "Foreground Color", "&H80000008&")
End Sub

Public Sub Form_Unload(Cancel As Integer)
    If Not mHotKeyOpenWindow Is Nothing Then
       mHotKeyOpenWindow.Clear
       Set mHotKeyOpenWindow = Nothing
    End If

    ISandyWindowMain_ShutdownDLLs
    Set CurrentTemplate = Nothing
    SaveFormPosition Me
End Sub

Private Sub mHotKeyOpenWindow_HotKeyPress(ByVal sName As String, ByVal eModifiers As EHKModifiers, ByVal eKey As KeyCodeConstants)
    Dim sKey As String

    If sName = "Sandy Activate" Then
       mHotKeyOpenWindow.RestoreAndActivate Me.hWnd
    ElseIf sName = "Sandy Repeat Insertion" Then
       If Not InternalCurrentTemplate Is Nothing Then
          sKey = InternalCurrentTemplate.Key
       End If
    ElseIf sName = "Sandy Favorites" Then
       FavoriteCalledFromIDE = True
       ISandyWindowMain_ShowFavMenu
    ElseIf sName = "Sandy Externals" Then
       ISandyWindowMain_ShowExternalsMenu
    ElseIf Left$(sName, 9) = "TEMPLATE " Then
          sKey = Mid$(sName, 10)
    End If

    If Len(sKey) Then
       ISandyWindowMain_DoInsertion Nothing, sKey
    End If
End Sub

