VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Slice and Dice"
   ClientHeight    =   4890
   ClientLeft      =   3570
   ClientTop       =   3555
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   4275
      Left            =   30
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
On Error Resume Next
    txtMessage.Move 30, 30, ScaleWidth - 30, ScaleHeight - 30
End Sub


Private Sub mnuEditCopy_Click()
On Error Resume Next
    If Len(txtMessage.SelText) > 0 Then
       StringToClipboard txtMessage.SelText
    End If
    txtMessage.SetFocus
End Sub

Private Sub mnuEditCut_Click()
On Error Resume Next
    If Len(txtMessage.SelText) > 0 Then
       If StringToClipboard(txtMessage.SelText) Then
          txtMessage.SelText = vbNullString
       End If
    End If
    txtMessage.SetFocus
End Sub

Private Sub mnuEditPaste_Click()
On Error Resume Next
    txtMessage.SelText = Clipboard.GetText
    txtMessage.SetFocus
End Sub

Private Sub mnuEditSelectAll_Click()
On Error Resume Next
    txtMessage.SelStart = 1
    txtMessage.SelLength = Len(txtMessage.Text)
    txtMessage.SetFocus
End Sub


Private Sub mnuFileExit_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
       Select Case KeyCode
              Case vbKeyEscape
                   KeyCode = 0
                   Shift = 0
                   Unload Me

              Case Else
                   txtMessage.SetFocus
       End Select
    Else
       txtMessage.SetFocus
    End If
End Sub

