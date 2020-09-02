VERSION 5.00
Begin VB.Form frmURLEntry 
   Caption         =   "Editing URL Entry"
   ClientHeight    =   3735
   ClientLeft      =   5670
   ClientTop       =   1860
   ClientWidth     =   7635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDataPaste 
      Height          =   300
      Left            =   2085
      TabIndex        =   12
      Top             =   3405
      Width           =   5500
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   1815
      TabIndex        =   11
      ToolTipText     =   "Remove the currently selected entry from the list"
      Top             =   2520
      Width           =   285
   End
   Begin VB.ListBox lstDataToPost 
      Height          =   1815
      IntegralHeight  =   0   'False
      ItemData        =   "frmAutoWebURLEntry.frx":0000
      Left            =   2085
      List            =   "frmAutoWebURLEntry.frx":0002
      TabIndex        =   10
      Top             =   1600
      Width           =   5500
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   1815
      TabIndex        =   9
      ToolTipText     =   "Add a new entry to the list"
      Top             =   1605
      Width           =   285
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   15
      TabIndex        =   8
      Top             =   2910
      Width           =   1215
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   495
      Left            =   15
      TabIndex        =   7
      Top             =   2370
      Width           =   1215
   End
   Begin VB.TextBox txtObjectToActivate 
      Height          =   300
      Left            =   2085
      TabIndex        =   5
      Top             =   1100
      Width           =   5500
   End
   Begin VB.TextBox txtAltURL 
      Height          =   300
      Left            =   2085
      TabIndex        =   3
      Top             =   600
      Width           =   5500
   End
   Begin VB.TextBox txtURL 
      Height          =   300
      Left            =   2085
      TabIndex        =   1
      Top             =   100
      Width           =   5500
   End
   Begin VB.Label lblDataToPost 
      Caption         =   "Data To Post"
      Height          =   480
      Left            =   75
      TabIndex        =   6
      Top             =   1605
      Width           =   1200
   End
   Begin VB.Label lblObjectToActivate 
      Caption         =   "Object To Activate"
      Height          =   480
      Left            =   75
      TabIndex        =   4
      Top             =   1095
      Width           =   1860
   End
   Begin VB.Label lblAltURL 
      Caption         =   "Alt URL"
      Height          =   480
      Left            =   60
      TabIndex        =   2
      Top             =   585
      Width           =   1200
   End
   Begin VB.Label lblURL 
      Caption         =   "URL"
      Height          =   480
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   1200
   End
End
Attribute VB_Name = "frmURLEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msURLEntry As String

Public Canceled As Boolean
Public Property Get URLEntry() As String
    msURLEntry = txtURL & "~~~"
    msURLEntry = msURLEntry & txtAltURL & "~~~"
    msURLEntry = msURLEntry & ListToString(lstDataToPost, False, "$$$") & "~~~"
    msURLEntry = msURLEntry & txtObjectToActivate & "~~~"
    URLEntry = msURLEntry
End Property

Public Property Let URLEntry(NewData As String)
    msURLEntry = NewData
       
    txtURL = sGetToken(NewData, 1, "~~~")
    txtAltURL = sGetToken(NewData, 2, "~~~")
    StringToList sGetToken(NewData, 3, "~~~"), lstDataToPost, True, "$$$"
    txtObjectToActivate = sGetToken(NewData, 4, "~~~")
End Property

Private Sub cmdAdd_Click()
    Dim sNewEntry As String
    sNewEntry = InputBox("What should the new entry be ?" & vbCrLf & vbTab & "Name=Value", "ADD DATA TO POST")
    If Len(sNewEntry) Then
       lstDataToPost.AddItem sNewEntry
    End If
End Sub

Private Sub cmdCancel_Click()
    Canceled = True
    Hide
End Sub

Private Sub cmdOkay_Click()
    Canceled = False
    Hide
End Sub


Private Sub cmdRemove_Click()
    If lstDataToPost.ListIndex > -1 Then
       'If MsgBox("Are you sure you want to remove that item ?", vbYesNo, "REMOVE DATA ITEM TO POST WITH THIS URL") = vbYes Then
          lstDataToPost.RemoveItem lstDataToPost.ListIndex
       'End If
    End If
End Sub


Private Sub Form_Load()

    LoadFormPosition Me
End Sub


Private Sub Form_Resize()
On Error Resume Next
    txtURL.Width = ScaleWidth - txtURL.Left - 50
    txtAltURL.Width = ScaleWidth - txtAltURL.Left - 50
    txtObjectToActivate.Width = ScaleWidth - txtObjectToActivate.Left - 50
    lstDataToPost.Width = ScaleWidth - lstDataToPost.Left - 50
    txtDataPaste.Width = ScaleWidth - txtDataPaste.Left - 50
    txtDataPaste.Top = ScaleHeight - txtDataPaste.Height - 50
    lstDataToPost.Height = ScaleHeight - lstDataToPost.Top - txtDataPaste.Height - 100
End Sub


Private Sub Form_Unload(Cancel As Integer)

    SaveFormPosition Me
End Sub


Private Sub lstDataToPost_DblClick()
    Dim sNewEntry As String
    If lstDataToPost.ListIndex > -1 Then
       sNewEntry = InputBox("Edit the value to post", "EDIT POST VALUE", lstDataToPost.List(lstDataToPost.ListIndex))
       If Len(sNewEntry) > 0 Then
          lstDataToPost.List(lstDataToPost.ListIndex) = sNewEntry
       End If
    End If
End Sub


Private Sub lstDataToPost_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Shift = 0 Then
       lstDataToPost_DblClick
    End If
End Sub


Private Sub txtDataPaste_Change()
    If Len(txtDataPaste) = 0 Then Exit Sub
    If lTokenCount(txtDataPaste, "$$$") > 1 Then
       StringToList txtDataPaste, lstDataToPost, False, "$$$"
    ElseIf lTokenCount(txtDataPaste, "&") > 1 Then
       StringToList txtDataPaste, lstDataToPost, False, "&"
    End If
    txtDataPaste = ""
End Sub

