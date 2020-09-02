VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "RegisterHotKey Demo"
   ClientHeight    =   2955
   ClientLeft      =   4230
   ClientTop       =   2460
   ClientWidth     =   4185
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   4185
   Begin VB.Label Label2 
      Caption         =   "The HotKey for this demo has been set to CTRL-ALT-UP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   2160
      Width           =   4035
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "frmTest.frx":1272
      Top             =   120
      Width           =   1920
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4095
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmTest.frx":197E
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   780
      Width           =   4035
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_cHotKey  As cRegHotKey
Attribute m_cHotKey.VB_VarHelpID = -1

Private Sub Form_Load()
   Set m_cHotKey = New cRegHotKey
   m_cHotKey.Attach Me.hwnd
   m_cHotKey.RegisterKey "Activate", vbKeyUp, MOD_ALT + MOD_CONTROL
End Sub

Private Sub m_cHotKey_HotKeyPress(ByVal sName As String, ByVal eModifiers As EHKModifiers, ByVal eKey As KeyCodeConstants)
   m_cHotKey.RestoreAndActivate Me.hwnd
   MsgBox "Got HotKey: " & sName
End Sub
