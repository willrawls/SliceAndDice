VERSION 5.00
Object = "{699F1A81-813D-11D1-B4DC-0060979C4B57}#1.0#0"; "FirmSolutions.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   3090
   ClientTop       =   2325
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   9690
   Begin VB.Frame Frame2 
      Caption         =   " Arrange Icons "
      Height          =   1215
      Left            =   6600
      TabIndex        =   7
      Top             =   1680
      Width           =   2775
      Begin VB.OptionButton optArrange 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optArrange 
         Caption         =   "Left"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optArrange 
         Caption         =   "Top"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " View "
      Height          =   1455
      Left            =   6600
      TabIndex        =   2
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton optListStyle 
         Caption         =   "Report"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optListStyle 
         Caption         =   "List"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optListStyle 
         Caption         =   "Small Icon"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optListStyle 
         Caption         =   "Icon"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin FirmSolutions.FSListBar FSListBar1 
      Align           =   3  'Align Left
      Height          =   5145
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   9075
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelEdit       =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim CurImage As ListImage

    With FSListBar1
         .BarName = "Bar 0"
         .AddBar "Bar 1"
         .AddBar "Bar 2"
         .AddBar "Bar 3"
         .AddBar "Bar 4"
         .CurBar = 0
           .AddBarItem "Icon 1", , "Computer"
           .AddBarItem "Icon 2", , "FolderClosed"
           .AddBarItem "Icon 3", , "FolderOpen"
         .CurBar = 1
           .AddBarItem "Icon 4", , "Network"
           .AddBarItem "Icon 5", , "RecycleEmpty"
           .AddBarItem "Icon 6", , "RecycleFull"
         .CurBar = 2
           .AddBarItem "Icon 7", , "FloppyDrive"
           .AddBarItem "Icon 8", , "Drive"
           .AddBarItem "Icon 9", , "DisconnectedDrive"
         .CurBar = 3
           .AddBarItem "Icon 10", , "MailNone"
           .AddBarItem "Icon 11", , "MailSome"
           .AddBarItem "Icon 12", , "MailNew"
         .CurBar = 4
           .AddBarItem "Icon 13", , "Check"
           .AddBarItem "Icon 14", , "Question"
           .AddBarItem "Icon 15", , "Heart"
           .AddBarItem "Icon 16", , "Club"
           .AddBarItem "Icon 17", , "Diamond"
           .AddBarItem "Icon 18", , "Spade"
           
         For Each CurImage In .LargeListImages
             List1.AddItem CurImage.Key
         Next CurImage
    End With
End Sub


Private Sub FSListBar1_AfterBarClick()
    optListStyle(FSListBar1.View).Value = 1
End Sub

Private Sub FSListBar1_BarItemDblClick(BarName As String, BarItemName As String)
    MsgBox "Double click on:" & Chr$(13) & Chr$(9) & "Bar: " & BarName & Chr$(13) & Chr$(9) & "Item: " & BarItemName
End Sub

Private Sub List1_DblClick()
    FSListBar1.BarItemIcon = List1
End Sub


Private Sub optArrange_Click(Index As Integer)
    FSListBar1.Arrange = Index
End Sub


Private Sub optListStyle_Click(Index As Integer)
    FSListBar1.View = Index
End Sub


