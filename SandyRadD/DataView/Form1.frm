VERSION 5.00
Object = "{896BE5B9-95B3-11D2-88D0-006008AED66C}#1.0#0"; "FIRMSOLUTIONSDV.OCX"
Begin VB.Form frmMain 
   Caption         =   "Testing FirmSolutionsDV.DataView"
   ClientHeight    =   7830
   ClientLeft      =   1545
   ClientTop       =   1920
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   14205
   Begin FirmSolutionsDV.DataView DataView1 
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   9869
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      ScaleWidth      =   6390
      ScaleMode       =   0
      MouseIcon       =   "Form1.frx":0000
      FullRowSelect   =   -1  'True
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Arrange         =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With DataView1
         .DatabaseName = "D:\WMR\SliceAndDice\Copy of SliceAndDice.mdb"
         .RecordSource = "Template"
         .Requery
    End With
End Sub


Private Sub Form_Resize()
    DataView1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


