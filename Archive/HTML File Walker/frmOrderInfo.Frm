VERSION 4.00
Begin VB.Form frmOrderInfo 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3972
   ClientLeft      =   2628
   ClientTop       =   1992
   ClientWidth     =   6312
   ControlBox      =   0   'False
   Height          =   4356
   Left            =   2580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3972
   ScaleWidth      =   6312
   Top             =   1656
   Width           =   6408
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   3885
      Left            =   90
      ScaleHeight     =   3840
      ScaleWidth      =   6084
      TabIndex        =   0
      Top             =   75
      Width           =   6135
      Begin VB.TextBox lblOnExitMain 
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2610
         Left            =   30
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   705
         Width           =   3390
      End
      Begin VB.CommandButton cmdOrderNow 
         Caption         =   "&Order File Walker online now!"
         Height          =   405
         Left            =   3585
         TabIndex        =   10
         Top             =   3390
         Width           =   2415
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   390
         Left            =   2670
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   3390
         Width           =   840
      End
      Begin VB.Timer tmrRotate 
         Interval        =   1000
         Left            =   2616
         Top             =   1596
      End
      Begin VB.CommandButton cmdContinue 
         Caption         =   "Enter HTML &File Walker, Jr."
         Height          =   396
         Left            =   45
         TabIndex        =   1
         Top             =   3390
         Visible         =   0   'False
         Width           =   2556
      End
      Begin VB.Label lblSecondsRemaining 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5205
         TabIndex        =   12
         Top             =   2145
         Width           =   720
      End
      Begin VB.Image imgByWmr 
         Appearance      =   0  'Flat
         Height          =   372
         Left            =   3792
         Picture         =   "frmOrderInfo.frx":0000
         Top             =   1632
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Get Free Internet updates of HTML File Walker for Life !"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   630
         Index           =   5
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   3375
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Seaside, CA 93955"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   4
         Left            =   4260
         TabIndex        =   2
         Top             =   3000
         Width           =   1755
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seconds Remaining"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   3465
         TabIndex        =   7
         Top             =   2115
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Purchase the full version via MasterCard, Visa or Money Order"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   870
         Index           =   0
         Left            =   3450
         TabIndex        =   6
         Top             =   705
         Width           =   2535
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "P.O. Box 852"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   4260
         TabIndex        =   3
         Top             =   2745
         Width           =   1755
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   3465
         Picture         =   "frmOrderInfo.frx":2DCA
         Stretch         =   -1  'True
         Top             =   2505
         Width           =   810
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Firm Solutions"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   4260
         TabIndex        =   4
         Top             =   2505
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Send $20 U.S.. (plus sales tax in California) to:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   1
         Left            =   3450
         TabIndex        =   5
         Top             =   30
         Width           =   2520
      End
   End
End
Attribute VB_Name = "frmOrderInfo"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdContinue_Click()
    Dim sT$
    
    Me.Hide

    frmMain.Show
    If Len(frmMain.txtTitle) = 0 And Len(frmMain.txtOutFilename) = 0 And Len(frmMain.txtStartPath) = 0 Then
       frmMain.cmdReset_Click
    End If
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdOrderNow_Click()
    tmrRotate.Enabled = False
        lblSecondsRemaining.Caption = "0"
        tmrRotate_Timer
        HTML_ViewWebPage (gsHTMLFileWalkerHomePage), (frmMain.txtBrowserPath)
    tmrRotate.Enabled = True
End Sub

Private Sub Form_Load()
    Dim sT$
    
    lblSecondsRemaining.Caption = gsTimeOut

    CenterForm Me
    
    sT = "What do you get for $20.00 US ?" & Chr$(13) & Chr$(10)
    sT = sT & "1. Free updates via Internet for" & Chr$(13) & Chr$(10) & "     ONE FULL YEAR !" & Chr$(13) & Chr$(10)
    sT = sT & "2. Put a page you build directly onto the Web !" & Chr$(13) & Chr$(10)
    sT = sT & "3. Break directories of images into" & Chr$(13) & Chr$(10) & "     'Gallery' style Web sites !" & Chr$(13) & Chr$(10)
    sT = sT & "4. Automatically incorporate the contents of" & Chr$(13) & Chr$(10) & "     another web document before AND after" & Chr$(13) & Chr$(10) & "     a web page built with HTML File Walker!" & Chr$(13) & Chr$(10)
    sT = sT & "5. " & Chr$(13) & Chr$(10)
    sT = sT & Chr$(13) & Chr$(10) & "Thanks for trying HTML File Walker, Jr..." & Chr$(13) & Chr$(10)
    lblOnExitMain = sT
    lblOnExitMain.Visible = True
    lblOnExitMain.Locked = True
    
    tmrRotate_Timer
    Me.Show
    DoEvents
    LoadSettings frmMain.txtFileSettings
End Sub


Private Sub tmrRotate_Timer()
    If Val(lblSecondsRemaining.Caption) > 0 Then
       lblSecondsRemaining.Caption = CStr(Val(lblSecondsRemaining.Caption) - 1)
    Else
       cmdContinue.Visible = True
       cmdExit.Visible = True
       tmrRotate.Enabled = False
       lblSecondsRemaining.Visible = False
       Label2.Visible = False
    End If
End Sub



