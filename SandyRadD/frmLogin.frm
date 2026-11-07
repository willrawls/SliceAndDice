VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Thank you for considering registration of Slice and Dice."
   ClientHeight    =   2745
   ClientLeft      =   2340
   ClientTop       =   2370
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1621.837
   ScaleMode       =   0  'User
   ScaleWidth      =   8957.544
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      Height          =   345
      Index           =   5
      Left            =   150
      TabIndex        =   24
      Top             =   810
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   7140
      TabIndex        =   21
      Top             =   810
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Index           =   4
      Left            =   4800
      TabIndex        =   20
      Top             =   810
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   2475
      TabIndex        =   18
      Top             =   810
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Index           =   0
      Left            =   150
      TabIndex        =   12
      Text            =   "Sequence from Server"
      Top             =   150
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   7140
      TabIndex        =   9
      Text            =   "Fx(Local Generated)"
      Top             =   150
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Index           =   1
      Left            =   4800
      TabIndex        =   8
      Text            =   "Checksum"
      Top             =   150
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2475
      TabIndex        =   6
      Text            =   "Random"
      Top             =   150
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   7140
      TabIndex        =   5
      Top             =   1470
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Index           =   3
      Left            =   4800
      TabIndex        =   4
      Top             =   1470
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2475
      TabIndex        =   3
      Top             =   1470
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Index           =   2
      Left            =   150
      TabIndex        =   2
      Top             =   1470
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   2250
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   1290
      TabIndex        =   1
      Top             =   2250
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Rep Public ID"
      Height          =   195
      Index           =   11
      Left            =   150
      TabIndex        =   25
      Top             =   1200
      Width           =   990
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Rep Remote Confirmation"
      Height          =   195
      Index           =   10
      Left            =   7140
      TabIndex        =   23
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Rep Remote Generated"
      Height          =   195
      Index           =   9
      Left            =   4800
      TabIndex        =   22
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Rep Local Generated"
      Height          =   195
      Index           =   8
      Left            =   2475
      TabIndex        =   19
      Top             =   1200
      Width           =   1530
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Team Public ID"
      Height          =   195
      Index           =   7
      Left            =   150
      TabIndex        =   17
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Team Remote Confirmation"
      Height          =   195
      Index           =   6
      Left            =   7140
      TabIndex        =   16
      Top             =   1860
      Width           =   1920
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Team Remote Generated"
      Height          =   195
      Index           =   5
      Left            =   4800
      TabIndex        =   15
      Top             =   1860
      Width           =   1800
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Team Local Generated"
      Height          =   195
      Index           =   4
      Left            =   2475
      TabIndex        =   14
      Top             =   1860
      Width           =   1635
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "SHPC Public ID"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   13
      Top             =   540
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "SHPC Remote Confirmation"
      Height          =   195
      Index           =   3
      Left            =   7140
      TabIndex        =   11
      Top             =   540
      Width           =   1950
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "SHPC Remote Generated"
      Height          =   195
      Index           =   2
      Left            =   4800
      TabIndex        =   10
      Top             =   540
      Width           =   1830
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "SHPC Local Generated"
      Height          =   195
      Index           =   1
      Left            =   2475
      TabIndex        =   7
      Top             =   540
      Width           =   1665
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If txtPassword = "password" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

