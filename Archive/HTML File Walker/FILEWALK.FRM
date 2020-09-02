VERSION 4.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML File Walker Senior"
   ClientHeight    =   5130
   ClientLeft      =   1665
   ClientTop       =   1575
   ClientWidth     =   7050
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   5535
   Icon            =   "FILEWALK.frx":0000
   Left            =   1605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7050
   Top             =   1230
   Width           =   7170
   Begin VB.TextBox txtHostPort 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1710
      TabIndex        =   40
      Text            =   "21"
      Top             =   4395
      Width           =   720
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   300
      Index           =   2
      Left            =   6192
      TabIndex        =   39
      Top             =   1608
      Width           =   800
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   300
      Index           =   0
      Left            =   6192
      TabIndex        =   38
      Top             =   345
      Width           =   800
   End
   Begin VB.TextBox txtBrowserPath 
      Height          =   300
      Left            =   1224
      TabIndex        =   36
      Text            =   "C:\Program Files\Plus!\Microsoft Internet\iexplore.exe"
      Top             =   2628
      Width           =   4884
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   300
      Index           =   1
      Left            =   6192
      TabIndex        =   35
      Top             =   645
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4155
      TabIndex        =   33
      Text            =   "10"
      Top             =   2145
      Width           =   408
   End
   Begin VB.TextBox txtFTPOutputTo 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1725
      TabIndex        =   31
      Top             =   3990
      Width           =   4350
   End
   Begin VB.CheckBox chkMultipage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Break into pages of"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2148
      TabIndex        =   30
      Top             =   2196
      Width           =   2100
   End
   Begin VB.CheckBox chkTable 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Use Tables"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2148
      TabIndex        =   29
      Top             =   1980
      Width           =   2200
   End
   Begin VB.CommandButton cmdViewPage 
      Caption         =   "&View"
      Height          =   435
      Left            =   6192
      TabIndex        =   27
      Top             =   2556
      Width           =   800
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   435
      Left            =   5295
      TabIndex        =   14
      Top             =   4650
      Width           =   800
   End
   Begin VB.TextBox txtRemotePath 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1725
      TabIndex        =   10
      Top             =   3660
      Width           =   4350
   End
   Begin VB.TextBox txtPassword 
      Enabled         =   0   'False
      Height          =   300
      Left            =   4470
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   3390
      Width           =   1605
   End
   Begin VB.TextBox txtUsername 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1725
      TabIndex        =   8
      Top             =   3390
      Width           =   1725
   End
   Begin VB.TextBox txtHost 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1725
      TabIndex        =   7
      Top             =   3075
      Width           =   4350
   End
   Begin VB.CommandButton cmdPut 
      Caption         =   "&Put"
      Height          =   435
      Left            =   6192
      TabIndex        =   12
      Top             =   3096
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   432
      Left            =   6150
      TabIndex        =   13
      Top             =   4650
      Width           =   840
   End
   Begin VB.CheckBox chkLocal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Local file references"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   75
      TabIndex        =   11
      Top             =   2205
      Width           =   2200
   End
   Begin VB.TextBox txtStartPath 
      Height          =   300
      Left            =   2340
      TabIndex        =   3
      Top             =   990
      Width           =   3800
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2340
      TabIndex        =   0
      Top             =   36
      Width           =   3800
   End
   Begin VB.CommandButton cmdGo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Build"
      Height          =   435
      Left            =   6192
      TabIndex        =   6
      Top             =   1980
      Width           =   800
   End
   Begin VB.TextBox txtBackgroundSound 
      Height          =   300
      Left            =   2340
      TabIndex        =   2
      Top             =   645
      Width           =   3800
   End
   Begin VB.TextBox txtBackgroundPic 
      Height          =   300
      Left            =   2340
      TabIndex        =   1
      Top             =   345
      Width           =   3800
   End
   Begin VB.TextBox txtExclude 
      Height          =   300
      Left            =   2340
      TabIndex        =   4
      Top             =   1305
      Width           =   3800
   End
   Begin VB.TextBox txtOutFilename 
      Height          =   300
      Left            =   2340
      TabIndex        =   5
      Top             =   1605
      Width           =   3800
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Height          =   315
      Left            =   30
      TabIndex        =   15
      Top             =   1590
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkAsOf 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Include ""As of"" line"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   90
      TabIndex        =   28
      Top             =   1980
      Value           =   1  'Checked
      Width           =   2200
   End
   Begin VB.DirListBox dirStack 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1530
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   570
      TabIndex        =   42
      Top             =   1695
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Host Port:"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   13
      Left            =   -30
      TabIndex        =   41
      Top             =   4440
      Width           =   1695
   End
   Begin CIFTPLib.CIFTP ftpPutPage 
      Height          =   450
      Left            =   75
      Top             =   3120
      Width           =   480
      _Version        =   65537
      _ExtentX        =   847
      _ExtentY        =   794
      _StockProps     =   0
      AccessChannelConnectionWAV=   ""
      AccessChannelClosedWAV=   ""
      DataChannelConnectionWAV=   ""
      DataChannelClosedWAV=   ""
      FileClosedWAV   =   ""
      ListBoxesPopulatedWAV=   ""
      SocketClosedWAV =   ""
      WSAErrorWAV     =   ""
      HostName        =   ""
      HostAddress     =   ""
      RemoteFileName  =   ""
      LoginName       =   ""
      Password        =   ""
      RepresentationType=   ""
      WorkingDirectory=   ""
      OverwriteExistingFile=   -1  'True
   End
   Begin MSComDlg.CommonDialog cdgBrowse 
      Left            =   6375
      Top             =   1035
      _ExtentX        =   688
      _ExtentY        =   688
      DialogTitle     =   "Select the file to use"
      FontSize        =   2.54052e-29
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Browser Path:"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   10
      Left            =   -1008
      TabIndex        =   37
      Top             =   2664
      Width           =   2208
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      X1              =   -60
      X2              =   7050
      Y1              =   2445
      Y2              =   2460
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "images."
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   12
      Left            =   4620
      TabIndex        =   34
      Top             =   2190
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Remote filename:"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   11
      Left            =   15
      TabIndex        =   32
      Top             =   4035
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   -150
      X2              =   7842
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Remote path:"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   9
      Left            =   0
      TabIndex        =   26
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Password:"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   8
      Left            =   3435
      TabIndex        =   25
      Top             =   3450
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "User name:"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   7
      Left            =   -15
      TabIndex        =   24
      Top             =   3435
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "FTP Host:"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   6
      Left            =   -15
      TabIndex        =   23
      Top             =   3135
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Directory tree to convert:"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   22
      Top             =   1050
      Width           =   2205
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Web Page Title:"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   90
      Width           =   2205
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Background Sound:"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   690
      Width           =   2205
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Background Picture:"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   3
      Left            =   120
      TabIndex        =   20
      Top             =   390
      Width           =   2205
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exclude Directory:"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   660
      TabIndex        =   19
      Top             =   1350
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Output to:"
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   1260
      TabIndex        =   18
      Top             =   1650
      Width           =   1065
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Public Function bCheckCommandAndRemove%(sToCheck$, ByVal sOptionName$, sReturn$, bPauseAfter%)
    Dim sUOptionName$, sUToCheck$
    Dim nStart%, nEnd%

    sUOptionName = UCase$(sOptionName)
    sUToCheck = UCase$(sToCheck)

    If InStr(sUToCheck, sUOptionName & "(") > 0 Then
     ' Option with a parameter
       nStart = InStr(sUToCheck, sUOptionName) + Len(sOptionName) + 1
        nEnd = InStr(Mid$(sToCheck, nStart), ")") - 1
        sReturn = Mid$(sToCheck, nStart, nEnd)
       nStart = InStr(sUToCheck, sUOptionName)
        nEnd = InStr(Mid$(sToCheck, nStart), ")") + 1
        If nStart = 1 Then
           sToCheck = Mid$(sToCheck, nEnd)
        Else
           sToCheck = Left$(sToCheck, nStart - 1) & Mid$(sToCheck, nEnd + 1)
        End If
       bCheckCommandAndRemove = True
    ElseIf InStr(sUToCheck, sUOptionName) > 0 Then
     ' Option without a parameter
       sReturn = ""
       bCheckCommandAndRemove = True
    End If
End Function

Public Sub HandleCommandLineActions()
       Dim sCommandLine$
       Dim sCommandLineParam$

       sCommandLine = Command$

       gbQuiet = bCheckCommandAndRemove(sCommandLine, "-Quiet", sCommandLineParam, True)

       If bCheckCommandAndRemove(sCommandLine, "-Build", sCommandLineParam, True) Then
          cmdGo_Click
          DoEvents: DoEvents
       End If

       If bCheckCommandAndRemove(sCommandLine, "-Explore", sCommandLineParam, True) Then
        ' AutoExplore Shells out with
        '   Explorer.exe /e,/root,/idlist,<PATH>
          cmdViewPage_Click
          DoEvents: DoEvents
       End If

       If bCheckCommandAndRemove(sCommandLine, "-Exit", sCommandLineParam, True) Then
          cmdExit_Click
          DoEvents: DoEvents
       End If


End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    Select Case Index
           Case 0
                cdgBrowse.DialogTitle = "Select which PICTURE to display in the background of the web page."
                    cdgBrowse.filename = txtBackgroundPic
                    cdgBrowse.DefaultExt = ".gif"
                    cdgBrowse.Filter = "GIFs (*.gif)|*.gif|JPEGs (*.jpg;*.jpeg)|*.jpg;*.jpeg|All files (*.*)|*.*"
                    cdgBrowse.FilterIndex = 1
                    cdgBrowse.Action = 1
                If sGetPath(cdgBrowse.filename) = App.path & "\" Then
                   txtBackgroundPic = sGetFilename(cdgBrowse.filename)
                Else
                   txtBackgroundPic = cdgBrowse.filename
                End If

           Case 1
                cdgBrowse.DialogTitle = "Select which MIDI to play in the background."
                    cdgBrowse.filename = txtBackgroundSound
                    cdgBrowse.DefaultExt = ".mid"
                    cdgBrowse.Filter = "MIDI Songs (*.mid)|*.mid|WAVs  (*.wav)|*.wav|All files (*.*)|*.*"
                    cdgBrowse.FilterIndex = 1
                    cdgBrowse.Action = 1
                If sGetPath(cdgBrowse.filename) = App.path & "\" Then
                   txtBackgroundSound = sGetFilename(cdgBrowse.filename)
                Else
                   txtBackgroundSound = cdgBrowse.filename
                End If

           Case 2
                cdgBrowse.DialogTitle = "Select where you would like the web page saved to."
                    cdgBrowse.filename = txtOutFilename
                    cdgBrowse.DefaultExt = ".htm"
                    cdgBrowse.Filter = "Web pages (*.htm or *.html)|*.htm;*.html|Text file (*.txt)|*.txt|All files (*.*)|*.*"
                    cdgBrowse.FilterIndex = 1
                    cdgBrowse.Action = 1
                If sGetPath(cdgBrowse.filename) = App.path & "\" Then
                   txtOutFilename = sGetFilename(cdgBrowse.filename)
                Else
                   txtOutFilename = cdgBrowse.filename
                End If
    End Select
End Sub


Private Sub cmdExit_Click()
    SaveSettings App.path & "\HTMLFW.DAT"
    End
End Sub

Private Sub cmdGo_Click()
    Dim fh%

    Me.Hide

    If Not gbQuiet Then
       frmStatus.Show
       DoEvents
    End If

    sOut = "<HTML>" & sHTML_Enclose("TITLE", txtTitle) & gsEOL & "<BODY BACKGROUND=" & Chr$(34) & txtBackgroundPic & Chr$(34) & ">" & gs2EOL
    If Len(txtBackgroundSound) > 0 Then sOut = sOut & "<BGSOUND SRC=" & Chr$(34) & txtBackgroundSound & Chr$(34) & ">" & gs2EOL
    If Len(txtTitle) > 0 Then sOut = sOut & gsTAB & "<CENTER>" & gsEOL & gsTAB & gsTAB & sHTML_Enclose("H1", txtTitle) & gsEOL
    If chkAsOf.Value <> 0 Then
       sOut = sOut & gsTAB & gsTAB & sHTML_Enclose("H2", "As of: " & Format$(Now, "dddd, mmmm d, yyyy \a\t h:nn AM/PM")) & gsEOL & gsTAB & "</CENTER>" & gsEOL & "<BR>" & gs2EOL
    Else
       sOut = sOut & gsEOL & gsTAB & "</CENTER>" & gsEOL & "<BR>" & gs2EOL
    End If

  ' For each sub-directory in this directory
    DirProcess txtStartPath

    sOut = sOut & gsEOL & "<BR><HR><BR>" & sHTML_Enclose("CENTER", sHTML_Enclose("H4", "This page created with " & sHTML_BuildLink("HTML File Walker Junior", "http://www.redshift.com/~jon-tom/htmlfw/index.html") & " by " & sHTML_BuildLink("William M. Rawls", "http://www.redshift.com/~jon-tom/index.html")))
    sOut = sOut & gs2EOL & "</BODY>" & gsEOL & "</HTML>"

    fh = FreeFile
    Open txtOutFilename For Output Access Write As #fh
        Print #fh, sOut
    Close #fh

    If Not gbQuiet Then
       frmStatus.Hide
          Beep
       MsgBox "Web Page built successfully."
       Me.Show
    End If

End Sub

Private Sub cmdPut_Click()
    txtHost = Trim$(txtHost)
    If Len(txtHost) > 0 Then
       If Val(txtHost) > 0 Then
          ftpPutPage.HostAddress = txtHost
       Else
          ftpPutPage.HostName = txtHost
       End If
    Else
       MsgBox "Please supply the host name or address in the 'FTP Host' area." & gsEOL & gsTAB & "This will be the Internet name (such as stinger.redshift.com)" & gsEOL & gsTAB & "or IP Address (such as 123.123.123.4)" & gsEOL & gsTAB & "of the computer where you would like the page put."
    End If

    txtHostPort = Trim$(txtHostPort)
    If Val(txtHostPort) > 0 Then
       ftpPutPage.AccessPort = txtHostPort
    Else
       MsgBox "Please supply the host name or address." & gsEOL & gsTAB & "This will be the Internet name (such as stinger.redshift.com)" & gsEOL & gsTAB & "or IP Address (such as 123.123.123.4)" & gsEOL & gsTAB & "of the computer where you would like the page put."
    End If

    txtUsername = Trim$(txtUsername)
    If Len(txtUsername) > 0 Then
       ftpPutPage.LoginName = txtUsername
    Else
       MsgBox "Please supply the user to connect as in the 'User name' area."
    End If

    txtPassword = Trim$(txtPassword)
    If Val(txtPassword) > 0 Then
       ftpPutPage.Password = txtPassword
    Else
       MsgBox "Please supply the password for the user account in the 'Password' area."
    End If

    txtFTPOutputTo = Trim$(txtFTPOutputTo)
    If Val(txtFTPOutputTo) > 0 Then
       ftpPutPage.RemoveFileName = txtFTPOutputTo
    Else
       MsgBox "Please supply the directory on the Host to store the page in the 'Remote path' area."
    End If

    txtOutFilename = Trim$(txtOutFilename)
    If Val(txtOutFilename) > 0 Then
       ftpPutPage.LocalFileName = txtOutFilename
       ftpPutPage.RepresentationType = "A" ' HTML is always ASCII
    Else
       MsgBox "Please supply the filename to store the page in the 'Remote filename' area."
    End If

    ftpPutPage.PutFile

End Sub


Public Sub cmdReset_Click()
    txtTitle = "My Favorites"
    txtBackgroundPic = "aged-paper.jpg"
    txtBackgroundSound = "Cursum3.mid"
    txtStartPath = "c:\windows\favorites"
    txtExclude = ""
    txtOutFilename = "favorites.html"
    txtBrowserPath = "C:\Program Files\Plus!\Microsoft Internet\iexplore.exe"
    chkAsOf.Value = 1
End Sub

Private Sub cmdViewPage_Click()
    HTML_ViewWebPage (txtOutFilename), (txtBrowserPath)
End Sub

Private Sub dirStack_Change()
    File1.path = dirStack.path
End Sub


Private Sub Drive1_Change()
    dirStack.path = Drive1
End Sub


Private Sub Form_Load()
On Error Resume Next
    Dim sStartDrive$, sStartPath$, sCommandLine$, sCommandLineParam$

    CenterForm Me

    gsEOL = Chr$(13) & Chr$(10)
    gs2EOL = gsEOL & gsEOL
    gsTAB = Chr$(9)
    sLinePrefix = gsTAB & gsTAB
    sLineSuffix = gsEOL

    If Len(Command) > 0 Then
       HandleCommandLineVars
       DoEvents
    Else
       LoadSettings App.path & "\HTMLFW.DAT"
    End If

    Drive1 = "C:\"
    dirStack = sStartPath
    DoEvents

    If Len(Command) > 0 Then HandleCommandLineActions
End Sub

Private Sub HandleCommandLineVars()
       Dim sCommandLine$
       Dim sCommandLineParam$

       sCommandLine = Command$

       If bCheckCommandAndRemove(sCommandLine, "-Title", sCommandLineParam, True) Then
          txtTitle = sCommandLineParam
       End If

       If bCheckCommandAndRemove(sCommandLine, "-Pic", sCommandLineParam, True) Then
          txtBackgroundPic = sCommandLineParam
       End If

       If bCheckCommandAndRemove(sCommandLine, "-Sound", sCommandLineParam, True) Then
          txtBackgroundSound = sCommandLineParam
       End If

       If bCheckCommandAndRemove(sCommandLine, "-Dir", sCommandLineParam, True) Then
          txtStartPath = sCommandLineParam
       End If

       If bCheckCommandAndRemove(sCommandLine, "-Exclude", sCommandLineParam, True) Then
          txtExclude = sCommandLineParam
       End If

       If bCheckCommandAndRemove(sCommandLine, "-Output", sCommandLineParam, True) Then
          txtOutFilename = sCommandLineParam
       End If

       If bCheckCommandAndRemove(sCommandLine, "-Shell", sCommandLineParam, True) Then
          txtBrowserPath = sCommandLineParam
       End If
End Sub

Private Sub ftpPutPage_PutFile()
    MsgBox "Page successfully stored on '" & txtHost & "' as '" & txtFTPOutputTo & "'"
End Sub


Private Sub txtOutFilename_Change()
    txtFTPOutputTo = sGetFilename(txtOutFilename)
End Sub




