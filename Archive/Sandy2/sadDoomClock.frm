VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "William's Doom Clock - Y2k"
   ClientHeight    =   2655
   ClientLeft      =   30
   ClientTop       =   0
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrTimeLeft 
      Interval        =   7000
      Left            =   6120
      Top             =   2040
   End
   Begin VB.PictureBox picTheBug 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2700
      Left            =   180
      Picture         =   "sadDoomClock.frx":0000
      ScaleHeight     =   2700
      ScaleWidth      =   2250
      TabIndex        =   2
      ToolTipText     =   "Click and drag to move this window where you want it."
      Top             =   -30
      Width           =   2250
   End
   Begin VB.Label lblFirmSolutions 
      Caption         =   "By Firm Solutions"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4860
      TabIndex        =   4
      ToolTipText     =   "Double click to visit Firm Solutions Latest revolutionary VB5/6 add-in beta... Please 8)"
      Top             =   0
      Width           =   1635
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Time remaining till the end of the world"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2490
      TabIndex        =   1
      ToolTipText     =   "Double click to end program."
      Top             =   360
      Width           =   3945
   End
   Begin VB.Label lblTimeLeft 
      BackColor       =   &H000000FF&
      Caption         =   "No place on Earth will protect you from Y2k..."
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   5010
      TabIndex        =   3
      Top             =   900
      Width           =   1245
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   2865
      Left            =   4860
      Top             =   -90
      Width           =   1635
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DOOM CLOCK"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1515
      Index           =   1
      Left            =   2190
      TabIndex        =   0
      ToolTipText     =   "Double click to end program."
      Top             =   750
      Width           =   2805
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2085
      Left            =   0
      Top             =   270
      Width           =   6855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Deactivate()
    Me.ZOrder 1
End Sub

Private Sub Form_Load()

    LoadFormPosition frmMain, False
End Sub


Private Sub Form_Unload(Cancel As Integer)

    SaveFormPosition frmMain
End Sub


Private Sub lblButton_DblClick(Index As Integer)
    Unload Me
End Sub


Private Sub lblFirmSolutions_DblClick()
    Shell "explorer http://www.firmsolutions.com/VB5CodeWalker/techsupport.html"
End Sub


Private Sub picTheBug_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
       Me.Move Me.Left + X, Me.Top + Y
    End If
End Sub


Private Sub tmrTimeLeft_Timer()
    Static mbShowAlready As Boolean
    Dim sRoughTime As String
    Dim sOutput As String
    
    If CLng(CVDate("1/1/2000") - Now()) > 0 Then
       sRoughTime = Int(CVDate("1/1/2000") - Now()) & " days " & Format(CVDate("1/1/2000") - Now(), "hh \h\o\u\r\s nn \m\i\n\u\t\e\s ss \s\e\c\o\n\d\s")
        
       sOutput = sGetToken(sRoughTime, 1, " ") & " " & sGetToken(sRoughTime, 2, " ") & Chr(13) & Chr(10)
       sOutput = sOutput & sGetToken(sRoughTime, 3, " ") & " " & sGetToken(sRoughTime, 4, " ") & Chr(13) & Chr(10)
       sOutput = sOutput & sGetToken(sRoughTime, 5, " ") & " " & sGetToken(sRoughTime, 6, " ") & Chr(13) & Chr(10)
       lblTimeLeft.Caption = sOutput & sGetToken(sRoughTime, 7, " ") & " " & sGetToken(sRoughTime, 8, " ")
       tmrTimeLeft.Interval = Int((20000 - 1000 + 1) * Rnd + 1000)
    ElseIf Not mbShowAlready Then
       lblTimeLeft = "If you can read this, then the worst is over."
    End If
End Sub


