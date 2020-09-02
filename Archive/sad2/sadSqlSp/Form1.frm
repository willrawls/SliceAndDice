VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2556
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2556
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   732
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   1752
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cdaSql As clsSqlSp
Private Sub Command1_Click()
    Set cdaSql = New clsSqlSp
'    MsgBox cdaSql.spParamList("dsn=projectsite", "_test")
    MsgBox cdaSql.spSelectList("_test", "1, 343,43,2", "dsn=projectsite")
End Sub
