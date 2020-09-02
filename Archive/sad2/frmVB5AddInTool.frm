VERSION 5.00
Begin VB.Form frmRegisterCodeSnippet 
   Caption         =   "VB5 Add-in registration tool"
   ClientHeight    =   2055
   ClientLeft      =   5655
   ClientTop       =   3465
   ClientWidth     =   5115
   Icon            =   "frmVB5AddInTool.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtObjectName 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "SliceAndDice.Wizard"
      Top             =   480
      Width           =   2805
   End
   Begin VB.CommandButton cmdRemoveAddin 
      Caption         =   "Remove from VB5 Add-in list"
      Height          =   435
      Left            =   2190
      TabIndex        =   3
      Top             =   1530
      Width           =   2805
   End
   Begin VB.TextBox txtLocation 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Text            =   "c:\windows\vbaddin.ini"
      Top             =   120
      Width           =   2805
   End
   Begin VB.CommandButton cmdAddAddin 
      Caption         =   "Add to VB5 Add-in list"
      Height          =   435
      Left            =   2190
      TabIndex        =   2
      Top             =   930
      Width           =   2805
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Addin Object Name"
      Height          =   195
      Index           =   1
      Left            =   660
      TabIndex        =   5
      Top             =   540
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Addin list Location"
      Height          =   195
      Index           =   0
      Left            =   780
      TabIndex        =   4
      Top             =   180
      Width           =   1290
   End
End
Attribute VB_Name = "frmRegisterCodeSnippet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub RegisterIt()
    Dim fh As Integer
    Dim sLine As String

    fh = FreeFile
    
  ' Make sure it's not already there
    Open txtLocation For Input Access Read As #fh
         Do Until EOF(fh)
            Input #fh, sLine
            If InStr(UCase(sLine), UCase(txtObjectName)) > 0 Then
               MsgBox "'" & txtObjectName & "' has already been registered. Aborting action."
               Close #fh
               Exit Sub
            End If
         Loop
    Close #fh
    
    Open txtLocation For Append Access Write As #fh
         Print #fh, txtObjectName & "=0"
    Close #fh
End Sub

Private Sub cmdAddAddin_Click()
    RegisterIt
    MsgBox "Code Snippet Add-in added successfully. Use the Add-in Manager to activate it."
End Sub


Private Sub cmdRemoveAddin_Click()
    Dim fh As Integer
    Dim sLine As String
    Dim sBackOut As String

    fh = FreeFile
    Open txtLocation For Input Access Read As #fh
         Do Until EOF(fh)
            Input #fh, sLine
            If InStr(sLine, txtObjectName) = 0 Then
               sBackOut = sBackOut & sLine & Chr$(13) & Chr$(10)
            End If
         Loop
    Close #fh
    
    Open txtLocation For Output Access Write As #fh
         Print #fh, sBackOut
    Close #fh

    MsgBox "Code Snippet Add-in removed successfully."
End Sub


Private Sub Form_Load()
    txtLocation.Text = sGetWindowsDir() & "vbaddin.ini"
    
    If UCase(Command) = "REGISTER SLICE AND DICE" Then
       RegisterIt
       End
    End If
End Sub

