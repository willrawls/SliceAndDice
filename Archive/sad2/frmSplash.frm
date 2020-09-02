VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4470
   ClientLeft      =   1695
   ClientTop       =   3615
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
         Caption         =   "http://willrawls.ewebcity.com/coreTeam"
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
         Height          =   192
         Index           =   4
         Left            =   1380
         MouseIcon       =   "frmSplash.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   37
         ToolTipText     =   "Manually submit a template or category for inclusion in the general S&D product."
         Top             =   3348
         Width           =   2904
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opensource Dev:"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Index           =   11
         Left            =   36
         TabIndex        =   36
         Top             =   3348
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://willrawls.ewebcity.com/sndforum"
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
         Height          =   192
         Index           =   2
         Left            =   1380
         MouseIcon       =   "frmSplash.frx":044E
         MousePointer    =   99  'Custom
         TabIndex        =   35
         ToolTipText     =   "Manually submit a template or category for inclusion in the general S&D product."
         Top             =   3096
         Width           =   2844
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User forums:"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Index           =   2
         Left            =   30
         TabIndex        =   34
         Top             =   3096
         Width           =   900
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         Height          =   192
         Index           =   1
         Left            =   7452
         TabIndex        =   31
         ToolTipText     =   "This is the number of S&D add-in DLLs currently loaded in memory."
         Top             =   2388
         Width           =   96
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7200
         TabIndex        =   30
         ToolTipText     =   "This is the TCP/IP address of your Team's S&D Server (Team Slice and Dice only)."
         Top             =   3495
         Width           =   1125
      End
      Begin VB.Label CentralIP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "209.196.104.22"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7200
         TabIndex        =   29
         ToolTipText     =   "This is the TCP/IP address of the S&D Internet Server."
         Top             =   3300
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Team IP:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   6480
         TabIndex        =   28
         Top             =   3495
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Central IP:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   6390
         TabIndex        =   27
         Top             =   3300
         Width           =   735
      End
      Begin VB.Label ParentRepID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999-999999"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7200
         TabIndex        =   26
         ToolTipText     =   "This is the ID of the Representative who made you a S&D representative."
         Top             =   3900
         Width           =   855
      End
      Begin VB.Label UserRepID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999-999999"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7200
         TabIndex        =   25
         ToolTipText     =   "This is your S&D Representative ID (Only if you've registered to sell S&D)."
         Top             =   3705
         Width           =   855
      End
      Begin VB.Label TeamID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999-999999"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7200
         TabIndex        =   24
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
         TabIndex        =   23
         ToolTipText     =   "This is your S&D serial number."
         Top             =   2895
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parent Rep ID:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   6060
         TabIndex        =   22
         Top             =   3900
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your Rep ID:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   6195
         TabIndex        =   21
         Top             =   3705
         Width           =   930
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Team ID:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   6465
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   2070
         Width           =   1665
      End
      Begin VB.Label lblDLLsLoaded 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sandals Loaded:"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Index           =   0
         Left            =   6144
         TabIndex        =   17
         ToolTipText     =   "This is the number of S&D add-in DLLs currently loaded in memory."
         Top             =   2388
         Width           =   1236
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   2148
         Index           =   0
         Left            =   5940
         Top             =   2016
         Width           =   2688
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Agreement:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   30
         TabIndex        =   16
         Top             =   3660
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Submit Template:"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Index           =   3
         Left            =   30
         TabIndex        =   15
         Top             =   2832
         Width           =   1236
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
         Height          =   216
         Index           =   3
         Left            =   1380
         MouseIcon       =   "frmSplash.frx":0890
         MousePointer    =   99  'Custom
         TabIndex        =   14
         ToolTipText     =   "Manually submit a template or category for inclusion in the general S&D product."
         Top             =   2832
         Width           =   2988
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
         Left            =   540
         TabIndex        =   5
         Top             =   936
         Width           =   3960
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
         MouseIcon       =   "frmSplash.frx":0FDC
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
         MouseIcon       =   "frmSplash.frx":141E
         MousePointer    =   99  'Custom
         TabIndex        =   8
         ToolTipText     =   "Visit the main S&D Web Site"
         Top             =   2310
         Width           =   2640
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "William Rawls retains rights to everything before 4/1/2000. All public domain thereafter."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   72
         TabIndex        =   3
         ToolTipText     =   "Slice and Dice is copyright 1999 by Firm Solutions and William M. Rawls. Slice and Dice is a trademark of Firm Solutions."
         Top             =   1500
         Width           =   8496
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warning: You are responsible for any damages caused by this program. Use with care."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   336
         TabIndex        =   2
         Top             =   4176
         Width           =   7236
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
         Height          =   288
         Left            =   756
         TabIndex        =   4
         Top             =   1260
         Width           =   888
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
         Caption         =   "all freeware, all opensource"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         TabIndex        =   6
         Top             =   15
         Width           =   4245
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
         Height          =   756
         Left            =   204
         TabIndex        =   7
         Top             =   336
         Width           =   6372
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   2148
         Left            =   0
         Top             =   2016
         Width           =   5928
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
1        On Error Resume Next
    'Dim bOkaySoFar As Boolean

    'bOkaySoFar = IsDate(CVDate(sFOfX(1))) And IsDate(CVDate(sFOfX(3)))
    'If Not bOkaySoFar Then
    '   Call sFOfX(0)
    'Else
    '   bOkaySoFar = (Format$(sGetToken(sFOfX(3), 1, gsP), "00000") = sGetToken(sFOfX(8), 1, "-")) And (sGetToken(Format$(CVDate(sFOfX(3)), ".00000"), 2, gsP) = sGetToken(sFOfX(8), 3, "-"))
    'End If

2        UserID = sFOfX(8)
3        TeamID = sFOfX(10)
4        UserRepID = sFOfX(12)
5        ParentRepID = sFOfX(14)

6        CentralIP = IIf(sadGetLicenseKey("CentralTemplateLibrary", "Unknown") <> "Unknown", "Known", "Unknown")
7        TeamIP = IIf(sadGetLicenseKey("TeamIP", "None") <> "None", "Known", "None")

8        lblProductName = gsSliceAndDice
9        lblPlatform = "VB 5 / 6 - Win 95/98/NT4/2000"

10       lblDaysLeft = 99                                  ' CLng(CVDate(sFOfX(1)) - Now)

    'If (Now > 36400) And (Not bOkaySoFar) Then
    '   SaveSetting "API Viewer", "Options", "AllowCopy", "True"
    'End If
11       gbEvaluationHasExpired = False                    '((((Val(lblDaysLeft) < 0)) And (UserID = "00000-000-00000" Or (lTokenCount(UserID, "-") <> 3) Or (lTokenCount(TeamID, "-") <> 3) Or (lTokenCount(UserRepID, "-") <> 3) Or (lTokenCount(ParentRepID, "-") <> 3)))) 'Or (GetSetting$("API Viewer", "Options", "AllowCopy", vbNullString) = "True")    If gbEvaluationHasExpired Then Exit Function

    'If StrComp(UserID, "00000-000-00000") = 0 Then 'Or (lTokenCount(UserID, "-") <> 3) Then 'Or Len(UserID) <> 15 Then ' Or Len(TeamID) <> 15 Or Len(UserRepID) <> 15 Or Len(ParentRepID) <> 15 Then 'Or Len(sadGetLicenseKey("Invoice Number", vbNullString)) = 0 Then
    '   lblDayLeftCaption = "Days Left in your Free Evaluation Period:"
    '   lblLicenseTo = "Evaluation Licensed to: " & sFOfX
    '   lblVersion = "Version " & App.Major & gsP & App.Minor & gsP & App.Revision & " Evaluation"
    '   DetermineRegistration = False
    'Else
12       lblDayLeftCaption = "Thank you for using " & gsSliceAndDice
13       lblDaysLeft = vbNullString
14       lblLicenseTo = "Licensed to: " & sFOfX & ", fwSHPC"
15       lblVersion = "Version " & App.Major & gsP & App.Minor & gsP & App.Revision
16       DetermineRegistration = True
    'End If
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
17       Unload Me
End Sub

Private Sub Form_Load()
18       On Error Resume Next
19       DetermineRegistration
20       Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

End Sub

Friend Function sFOfX(Optional ByVal Value As Long = 5)
21       Dim sLicenseInfo As String
22       Dim RightNow As Date

23       On Error Resume Next
24       If Value > 0 Then
25           sLicenseInfo = sadGetLicenseKey("Value" & Format$(Value, "00"), vbNullString)
26       End If

27       If Len(sLicenseInfo) = 0 Or Value = 0 Then
28           sLicenseInfo = InputBox("What's your name ?", "FREEWARE LICENSE INFO", vbNullString)
29           If Len(sLicenseInfo) = 0 Then sLicenseInfo = "Unknown User"
30           RightNow = Now
31           CreateKey "HKEY_LOCAL_MACHINE" & gsBS & "SOFTWARE" & gsBS & "Zion Systems" & gsBS & "License"
32           sadSaveLicenseKey "Value01", vbNullString & CDbl(RightNow + 30)
33           sadSaveLicenseKey "Value02", vbNullString & CStr(Rnd * 10000000) & "!)$*%&@("
34           sadSaveLicenseKey "Value03", vbNullString & CDbl(RightNow)
35           sadSaveLicenseKey "Value04", vbNullString & CStr(Rnd * 10000000)    '& "@#*%&@(@#"
36           sadSaveLicenseKey "Value05", vbNullString & sLicenseInfo    ' User Name
37           sadSaveLicenseKey "Value06", vbNullString & CStr(Rnd * 10000000)    '& "@#)*($^YTIcm")
38           sadSaveLicenseKey "Value07", vbNullString & CStr(Rnd * 10000000)    '& "!@)^&#(KLSpsi42")
39           sadSaveLicenseKey "Value08", vbNullString & Left$(CDbl(Now) & "", 5) & "-999-" & Mid$(CDbl(Now) & "", 7, 5)    '"00000-000-00000"     ' User ID
40           sadSaveLicenseKey "Value09", vbNullString & CStr(Rnd * 10000000)    '& ":kdofie(843")
41           sadSaveLicenseKey "Value10", vbNullString & Left$(CDbl(Now) & "", 5) & "-999-" & Mid$(CDbl(Now) & "", 7, 5)    '"00000-000-00000"     ' Team ID
42           sadSaveLicenseKey "Value11", vbNullString & CStr(Rnd * 10000000)    '& "(840Usoifn1$")
43           sadSaveLicenseKey "Value12", vbNullString & Left$(CDbl(Now) & "", 5) & "-999-" & Mid$(CDbl(Now) & "", 7, 5)    '"00000-000-00000"     ' User Rep ID
44           sadSaveLicenseKey "Value13", vbNullString & CStr(Rnd * 10000000)    '& "(39WU905dj!@~")
45           sadSaveLicenseKey "Value14", vbNullString & "00001-001-00001"    ' Parent Rep ID (WMR)
46       End If
47       If Value > 0 Then
48           sFOfX = sadGetLicenseKey("Value" & Format$(Value, "00"), vbNullString)
49       End If
End Function

Private Sub Frame1_Click()
50       Unload Me
End Sub

Private Sub imgLogo_Click()
51       BrowseTo "http://www.sliceanddice.com"
End Sub

Private Sub Label1_Click(Index As Integer)
52       BrowseTo vbNullString & Label1(Index).Caption
End Sub

Private Sub Label2_Click()
53       BrowseTo "http://www.sliceanddice.com/agreement.html"
End Sub

Private Sub lblCompanyProduct_Click()
54       Frame1_Click
End Sub

Private Sub lblCopyright_Click()
55       Frame1_Click
End Sub

Private Sub lblLicenseTo_Click()
56       Frame1_Click
End Sub


Private Sub lblPlatform_Click()
57       Frame1_Click
End Sub

Private Sub lblProductName_Click()
58       Frame1_Click
End Sub


Private Sub lblVersion_Click()
59       Frame1_Click
End Sub


Private Sub lblWarning_Click()
60       Frame1_Click
End Sub


