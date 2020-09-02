VERSION 5.00
Begin VB.Form Search2Form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get File List"
   ClientHeight    =   1305
   ClientLeft      =   4695
   ClientTop       =   2145
   ClientWidth     =   5610
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1305
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkIncludeSubDirs 
      Caption         =   "Include sub-directories"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1290
      TabIndex        =   5
      Top             =   900
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.TextBox txtFilePattern 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "*.*"
      ToolTipText     =   "Enter a pattern to search for (*.* finds everything, etc.)"
      Top             =   480
      Width           =   1275
   End
   Begin VB.TextBox txtStartDir 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Enter the drive and path to start the search at"
      Top             =   120
      Width           =   4155
   End
   Begin VB.CommandButton cmdSearch 
      Cancel          =   -1  'True
      Caption         =   "&Get File List"
      Default         =   -1  'True
      Height          =   495
      Left            =   4260
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "File Pattern"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Start Directory"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Search2Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FileList As String

' Initialize the file name.
Private Sub Form_Load()
    Dim sFilePath As String
    sFilePath = App.Path
    If Right$(sFilePath, 1) <> "\" Then sFilePath = sFilePath & "\"
    txtStartDir.Text = sFilePath
End Sub

' Search for subdirectories.
Private Sub cmdSearch_Click()
    Dim sFileList As String
    FileList = GetFileList(txtStartDir.Text, txtFilePattern.Text)
    Hide
End Sub
