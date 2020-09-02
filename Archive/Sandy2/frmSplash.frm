VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4470
   ClientLeft      =   3105
   ClientTop       =   3300
   ClientWidth     =   8565
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   4620
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Click on any non-web link to close this form. Thank you for using Slice and Dice !"
      Top             =   60
      Width           =   8520
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.sliceanddice.com/register.html"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   1380
         MouseIcon       =   "frmSplash.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   37
         ToolTipText     =   "Manually submit a template or category for inclusion in the general S&D product."
         Top             =   3390
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Register Sandy"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   30
         TabIndex        =   36
         Top             =   3390
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblDaysLeft 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "??????????"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6720
         TabIndex        =   35
         Top             =   390
         Width           =   1695
      End
      Begin VB.Label lblDayLeftCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Days Left in your Free Evaluation Period:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   6720
         TabIndex        =   34
         Top             =   0
         Width           =   1620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDLLsLoaded 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   7200
         TabIndex        =   33
         ToolTipText     =   "This is the number of S&D add-in DLLs currently loaded in memory."
         Top             =   2640
         Width           =   90
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   675
         Index           =   1
         Left            =   6660
         Top             =   -30
         Width           =   1905
      End
      Begin VB.Label TeamIP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "209.196.104.22"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   7200
         TabIndex        =   32
         ToolTipText     =   "This is the TCP/IP address of your Team's S&D Server (Team Slice and Dice only)."
         Top             =   3495
         Width           =   1125
      End
      Begin VB.Label CentralIP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "209.196.104.22"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   7200
         TabIndex        =   31
         ToolTipText     =   "This is the TCP/IP address of the S&D Internet Server."
         Top             =   3300
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Team IP:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   6480
         TabIndex        =   30
         Top             =   3495
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Central IP:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   6390
         TabIndex        =   29
         Top             =   3300
         Width           =   735
      End
      Begin VB.Label ParentRepID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999-999999"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   7200
         TabIndex        =   28
         ToolTipText     =   "This is the ID of the Representative who made you a S&D representative."
         Top             =   3900
         Width           =   855
      End
      Begin VB.Label UserRepID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999-999999"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   7200
         TabIndex        =   27
         ToolTipText     =   "This is your S&D Representative ID (Only if you've registered to sell S&D)."
         Top             =   3705
         Width           =   855
      End
      Begin VB.Label TeamID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999-999999"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   7200
         TabIndex        =   26
         ToolTipText     =   "This is your Team's S&D serial number (Team Slice and Dice only)."
         Top             =   3090
         Width           =   855
      End
      Begin VB.Label UserID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999-999999"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   7200
         TabIndex        =   25
         ToolTipText     =   "This is your S&D serial number."
         Top             =   2895
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parent Rep ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   6060
         TabIndex        =   24
         Top             =   3900
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your Rep ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   6195
         TabIndex        =   23
         Top             =   3705
         Width           =   930
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Team ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   6465
         TabIndex        =   22
         Top             =   3090
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   6540
         TabIndex        =   21
         Top             =   2895
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Miscellaneous Info:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   6000
         TabIndex        =   20
         Top             =   2070
         Width           =   1665
      End
      Begin VB.Label lblDLLsLoaded 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Addins Loaded:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   6015
         TabIndex        =   19
         ToolTipText     =   "This is the number of S&D add-in DLLs currently loaded in memory."
         Top             =   2640
         Width           =   1110
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   2145
         Index           =   0
         Left            =   5940
         Top             =   2010
         Width           =   2685
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Agreement:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   30
         TabIndex        =   18
         Top             =   3660
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Submit Template:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   30
         TabIndex        =   17
         Top             =   3120
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.sliceanddice.com/submit.html"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   1380
         MouseIcon       =   "frmSplash.frx":044E
         MousePointer    =   99  'Custom
         TabIndex        =   16
         ToolTipText     =   "Manually submit a template or category for inclusion in the general S&D product."
         Top             =   3120
         Width           =   2985
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Team S&&D:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   30
         TabIndex        =   15
         Top             =   2850
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.sliceanddice.com/teamsad.html"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   1380
         MouseIcon       =   "frmSplash.frx":0890
         MousePointer    =   99  'Custom
         TabIndex        =   14
         ToolTipText     =   "Find out more about Team Slice and Dice, an enterprise solution."
         Top             =   2850
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report an Issue:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   13
         Top             =   2580
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.sliceanddice.com/sadissue.html"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   1380
         MouseIcon       =   "frmSplash.frx":0CD2
         MousePointer    =   99  'Custom
         TabIndex        =   12
         ToolTipText     =   "Submit an issue / bug / feature directly to the S&D developer."
         Top             =   2580
         Width           =   3180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Latest updates:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   11
         Top             =   2310
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Some places on the web to go for more information:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   10
         Top             =   2070
         Width           =   3630
      End
      Begin VB.Label lblPlatform 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visual Basic (SP3) - Win95"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2010
         TabIndex        =   5
         Top             =   930
         Width           =   3960
      End
      Begin VB.Image imgLogo 
         Height          =   1725
         Left            =   -30
         MouseIcon       =   "frmSplash.frx":0FDC
         MousePointer    =   99  'Custom
         Picture         =   "frmSplash.frx":141E
         Stretch         =   -1  'True
         ToolTipText     =   "Visit the Firm Solutions Web Site"
         Top             =   0
         Width           =   1725
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Use of Slice and Dice is governed by the Slice and Dice end-user agreement. Click here to view it."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   1350
         MouseIcon       =   "frmSplash.frx":4420
         MousePointer    =   99  'Custom
         TabIndex        =   9
         ToolTipText     =   "View end-user agreement."
         Top             =   3660
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.sliceanddice.com/dl.html"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   1380
         MouseIcon       =   "frmSplash.frx":4862
         MousePointer    =   99  'Custom
         TabIndex        =   8
         ToolTipText     =   "Visit the main S&D Web Site"
         Top             =   2310
         Width           =   2640
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright 1999, Firm Solutions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2010
         TabIndex        =   3
         ToolTipText     =   "Slice and Dice is copyright 1999 by Firm Solutions and William M. Rawls. Slice and Dice is a trademark of Firm Solutions."
         Top             =   1500
         Width           =   2190
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warning: Firm Solutions is not liable for any damages caused by this program. Use with care."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   330
         TabIndex        =   2
         Top             =   4170
         Width           =   7650
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2010
         TabIndex        =   4
         Top             =   1230
         Width           =   885
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Licensed to: William M. Rawls, Super Human Programmer"
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
         Left            =   2520
         TabIndex        =   1
         ToolTipText     =   "This is the name of the person who owns this copy of S&D"
         Top             =   1710
         Width           =   5835
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Firm Solutions proudly presents:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1755
         TabIndex        =   6
         Top             =   15
         Width           =   4890
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slice and Dice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1890
         TabIndex        =   7
         Top             =   330
         Width           =   4380
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   2145
         Left            =   -30
         Top             =   2010
         Width           =   5925
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Friend Function DetermineRegistration() As Boolean
On Error Resume Next
    'Dim bOkaySoFar As Boolean

    'bOkaySoFar = IsDate(CVDate(sFOfX(1))) And IsDate(CVDate(sFOfX(3)))
    'If Not bOkaySoFar Then
    '   Call sFOfX(0)
    'Else
    '   bOkaySoFar = (Format(sGetToken(sFOfX(3), 1, "."), "00000") = sGetToken(sFOfX(8), 1, "-")) And (sGetToken(Format(CVDate(sFOfX(3)), ".00000"), 2, ".") = sGetToken(sFOfX(8), 3, "-"))
    'End If

    UserID = sFOfX(8)
    TeamID = sFOfX(10)
    UserRepID = sFOfX(12)
    ParentRepID = sFOfX(14)

    CentralIP = IIf(sadGetLicenseKey("CentralTemplateLibrary", "Unknown") <> "Unknown", "Known", "Unknown")
    TeamIP = IIf(sadGetLicenseKey("TeamIP", "None") <> "None", "Known", "None")

    lblProductName = "Slice and Dice"
    lblPlatform = "VB 5 / 6 - Win95/98/NT4/2000, Y2oKay"

    lblDaysLeft = CLng(CVDate(sFOfX(1)) - Now)

    'If (Now > 36400) And (Not bOkaySoFar) Then
    '   SaveSetting "API Viewer", "Options", "AllowCopy", "True"
    'End If
    gbEvaluationHasExpired = ((((Val(lblDaysLeft) < 0) Or (Now > 36450)) And (UserID = "00000-000-00000" Or (lTokenCount(UserID, "-") <> 3) Or (lTokenCount(TeamID, "-") <> 3) Or (lTokenCount(UserRepID, "-") <> 3) Or (lTokenCount(ParentRepID, "-") <> 3)))) 'Or (GetSetting("API Viewer", "Options", "AllowCopy", vbNullString) = "True")
    If gbEvaluationHasExpired Then Exit Function

    If StrComp(UserID, "00000-000-00000") = 0 Then 'Or (lTokenCount(UserID, "-") <> 3) Then 'Or Len(UserID) <> 15 Then ' Or Len(TeamID) <> 15 Or Len(UserRepID) <> 15 Or Len(ParentRepID) <> 15 Then 'Or Len(sadGetLicenseKey("Invoice Number", vbNullString)) = 0 Then
       lblDayLeftCaption = "Days Left in your Free Evaluation Period:"
       lblLicenseTo = "Evaluation Licensed to: " & sFOfX
       lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision & " Evaluation"
       DetermineRegistration = False
    Else
       lblDayLeftCaption = "Thank you for registering Slice and Dice"
       lblDaysLeft = vbNullString
       lblLicenseTo = "Licensed to: " & sFOfX & ", Super Human Programmer"
       lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
       DetermineRegistration = True
    End If
End Function

Private Sub Form_Initialize()

    ' LogEvent "frmSplash: Initialize"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    DetermineRegistration
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
End Sub

Friend Function sFOfX(Optional ByVal Value As Long = 5)
    Dim sLicenseInfo As String
    Dim RightNow As Date

On Error Resume Next
    If Value > 0 Then
       sLicenseInfo = sadGetLicenseKey("Value" & Format(Value, "00"), vbNullString)
    End If

    If Len(sLicenseInfo) = 0 Or Value = 0 Then
       sLicenseInfo = InputBox("What's your name ?", "EVALUATION LICENSE INFO", vbNullString)
       If Len(sLicenseInfo) = 0 Then sLicenseInfo = "Unknown User"
       RightNow = Now
       CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License"
       sadSaveLicenseKey "Value01", vbNullString & CDbl(RightNow + 30)
       sadSaveLicenseKey "Value02", vbNullString & CStr(Rnd * 10000000) & "!)$*%&@("
       sadSaveLicenseKey "Value03", vbNullString & CDbl(RightNow)
       sadSaveLicenseKey "Value04", vbNullString & CStr(Rnd * 10000000)  '& "@#*%&@(@#"
       sadSaveLicenseKey "Value05", vbNullString & sLicenseInfo          ' User Name
       sadSaveLicenseKey "Value06", vbNullString & CStr(Rnd * 10000000)  '& "@#)*($^YTIcm")
       sadSaveLicenseKey "Value07", vbNullString & CStr(Rnd * 10000000)  '& "!@)^&#(KLSpsi42")
       sadSaveLicenseKey "Value08", vbNullString & "00000-000-00000"     ' User ID
       sadSaveLicenseKey "Value09", vbNullString & CStr(Rnd * 10000000)  '& ":kdofie(843")
       sadSaveLicenseKey "Value10", vbNullString & "00000-000-00000"     ' Team ID
       sadSaveLicenseKey "Value11", vbNullString & CStr(Rnd * 10000000)  '& "(840Usoifn1$")
       sadSaveLicenseKey "Value12", vbNullString & "00000-000-00000"     ' User Rep ID
       sadSaveLicenseKey "Value13", vbNullString & CStr(Rnd * 10000000)  '& "(39WU905dj!@~")
       sadSaveLicenseKey "Value14", vbNullString & "00000-000-00000"     ' Parent Rep ID
    End If
    If Value > 0 Then
       sFOfX = sadGetLicenseKey("Value" & Format(Value, "00"), vbNullString)
    End If
End Function

Private Sub Form_Terminate()

    ' LogEvent "frmSplash: Terminate"
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub imgLogo_Click()
    BrowseTo "http://www.firmsolutions.com"
End Sub

Private Sub Label1_Click(Index As Integer)
    BrowseTo vbNullString & Label1(Index).Caption
End Sub

Private Sub Label2_Click()
    BrowseTo "http://www.sliceanddice.com/agreement.html"
End Sub

Private Sub lblCompanyProduct_Click()
    Frame1_Click
End Sub

Private Sub lblCopyright_Click()
    Frame1_Click
End Sub

Private Sub lblLicenseTo_Click()
    Frame1_Click
End Sub


Private Sub lblPlatform_Click()
    Frame1_Click
End Sub

Private Sub lblProductName_Click()
    Frame1_Click
End Sub


Private Sub lblVersion_Click()
    Frame1_Click
End Sub


Private Sub lblWarning_Click()
    Frame1_Click
End Sub


