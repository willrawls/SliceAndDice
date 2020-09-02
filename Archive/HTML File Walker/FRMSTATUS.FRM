VERSION 5.00
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "VSOCX32.OCX"
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1320
   ClientLeft      =   2325
   ClientTop       =   1890
   ClientWidth     =   3195
   ControlBox      =   0   'False
   Icon            =   "FRMSTATUS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1320
   ScaleWidth      =   3195
   StartUpPosition =   2  'CenterScreen
   Begin VsOcxLib.VideoSoftElastic vseStatus 
      Height          =   1320
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3195
      _Version        =   327680
      _ExtentX        =   5636
      _ExtentY        =   2328
      _StockProps     =   70
      Caption         =   "Walking... "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Align           =   5
      CaptionPos      =   0
      Picture         =   "FRMSTATUS.frx":0442
      PicturePos      =   8
      CaptionStyle    =   4
      MousePointer    =   1
      Begin VsOcxLib.VideoSoftElastic vseCancel 
         Height          =   450
         Left            =   315
         TabIndex        =   1
         Top             =   720
         Width           =   1245
         _Version        =   327680
         _ExtentX        =   2196
         _ExtentY        =   794
         _StockProps     =   70
         Caption         =   "Cancel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         ConvInfo        =   1418783674
         CaptionPos      =   4
         Style           =   3
      End
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub vseCancel_Click()
    glInterruptBuild = True
    Visible = False
    frmMain.Visible = True
End Sub


