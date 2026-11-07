VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   6240
   ClientTop       =   2070
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   5550
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt"
      Height          =   525
      Left            =   4290
      TabIndex        =   4
      Top             =   3120
      Width           =   1245
   End
   Begin VB.TextBox txtDec 
      Height          =   2655
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3450
      Width           =   4185
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt"
      Height          =   525
      Left            =   4290
      TabIndex        =   2
      Top             =   510
      Width           =   1245
   End
   Begin VB.TextBox txtOut 
      Height          =   2655
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   795
      Width           =   4185
   End
   Begin VB.TextBox txtIn 
      Height          =   765
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "TestCrypt.frx":0000
      Top             =   30
      Width           =   4185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDecrypt_Click()
    txtDec.Text = sadDecrypt(txtOut.Text)
End Sub

Private Sub cmdEncrypt_Click()
    txtOut.Text = sadEncrypt(txtIn.Text)
End Sub


