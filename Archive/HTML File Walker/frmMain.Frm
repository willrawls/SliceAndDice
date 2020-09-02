VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "VSOCX32.OCX"
Object = "{43559273-DABD-11D0-9E15-444553540000}#1.0#0"; "MFTPX.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   Caption         =   "HTML File Walker Senior"
   ClientHeight    =   5355
   ClientLeft      =   1680
   ClientTop       =   3975
   ClientWidth     =   9510
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5355
   ScaleWidth      =   9510
   Begin VB.FileListBox File1 
      BackColor       =   &H000080FF&
      Height          =   1065
      Left            =   3510
      TabIndex        =   70
      Top             =   4230
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.DirListBox dirStack 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Height          =   990
      Left            =   1800
      TabIndex        =   71
      Top             =   4230
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Height          =   315
      Left            =   240
      TabIndex        =   69
      Top             =   4215
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VsOcxLib.VideoSoftElastic VideoSoftElastic8 
      Height          =   720
      Left            =   8025
      TabIndex        =   68
      Top             =   3210
      Width           =   1260
      _Version        =   327680
      _ExtentX        =   2222
      _ExtentY        =   1270
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      AutoSizeChildren=   1
      BorderWidth     =   8
      BevelOuter      =   0
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   480
         Left            =   120
         TabIndex        =   57
         Top             =   120
         Width           =   1020
      End
   End
   Begin VsOcxLib.VideoSoftElastic VideoSoftElastic7 
      Height          =   5355
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   9510
      _Version        =   327680
      _ExtentX        =   16775
      _ExtentY        =   9446
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Align           =   5
      AutoSizeChildren=   1
      Begin VsOcxLib.VideoSoftIndexTab vstMain 
         Height          =   5175
         Left            =   90
         TabIndex        =   0
         Top             =   90
         Width           =   9330
         _Version        =   327680
         _ExtentX        =   16457
         _ExtentY        =   9128
         _StockProps     =   102
         Caption         =   "Basic Page|Page Options|View Page|Store on FTP|Settings|Command Line|Custom|New"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         CurrTab         =   6
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         DogEars         =   -1  'True
         New3D           =   -1  'True
         Begin VsOcxLib.VideoSoftElastic VideoSoftElastic10 
            Height          =   4800
            Left            =   75045
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   330
            Width           =   9240
            _Version        =   327680
            _ExtentX        =   16298
            _ExtentY        =   8467
            _StockProps     =   70
            ConvInfo        =   1418783674
            BevelOuter      =   0
            TagWidth        =   300
            TagPosition     =   1
            Begin VB.CommandButton cmdWalkClipbaord 
               Caption         =   "Walk the File List on Clipboard"
               Height          =   435
               Left            =   3550
               TabIndex        =   89
               Top             =   3100
               Width           =   3180
            End
            Begin VB.CommandButton cmdWalkFileList 
               Caption         =   "Walk the File List Above"
               Height          =   435
               Left            =   270
               TabIndex        =   88
               Top             =   3100
               Width           =   3180
            End
            Begin VB.TextBox Text1 
               Height          =   2295
               Left            =   270
               MultiLine       =   -1  'True
               TabIndex        =   87
               Tag             =   "File List to Walk"
               Text            =   "frmMain.frx":0442
               Top             =   480
               Width           =   3180
            End
         End
         Begin VsOcxLib.VideoSoftElastic VideoSoftElastic9 
            Height          =   4800
            Left            =   45
            TabIndex        =   72
            Top             =   330
            Width           =   9240
            _Version        =   327680
            _ExtentX        =   16298
            _ExtentY        =   8467
            _StockProps     =   70
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ConvInfo        =   1418783674
            AutoSizeChildren=   3
            BevelOuter      =   0
            Begin VsOcxLib.VideoSoftElastic vseSpecial 
               Height          =   1080
               Index           =   0
               Left            =   90
               TabIndex        =   73
               Top             =   3630
               Width           =   9060
               _Version        =   327680
               _ExtentX        =   15981
               _ExtentY        =   1905
               _StockProps     =   70
               Caption         =   "Before End of Page (at bottom)"
               ConvInfo        =   1418783674
               AutoSizeChildren=   2
               BorderWidth     =   0
               BevelOuter      =   5
               Style           =   1
               Begin VB.TextBox txtSpecial 
                  Height          =   840
                  Index           =   3
                  Left            =   0
                  MultiLine       =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   56
                  Top             =   240
                  Width           =   7710
               End
               Begin VsOcxLib.VideoSoftElastic VideoSoftElastic11 
                  Height          =   840
                  Left            =   7800
                  TabIndex        =   91
                  Top             =   240
                  Width           =   1260
                  _Version        =   327680
                  _ExtentX        =   2222
                  _ExtentY        =   1482
                  _StockProps     =   70
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ConvInfo        =   1418783674
                  AutoSizeChildren=   1
                  BorderWidth     =   8
                  BevelOuter      =   0
               End
            End
            Begin VsOcxLib.VideoSoftElastic vseSpecial 
               Height          =   1080
               Index           =   1
               Left            =   90
               TabIndex        =   74
               Top             =   2460
               Width           =   9060
               _Version        =   327680
               _ExtentX        =   15981
               _ExtentY        =   1905
               _StockProps     =   70
               ConvInfo        =   1418783674
               AutoSizeChildren=   1
               BorderWidth     =   0
               BevelOuter      =   5
               MaxChildSize    =   7600
               TagWidth        =   300
               TagPosition     =   1
               Begin VB.TextBox txtSpecial 
                  Height          =   780
                  Index           =   5
                  Left            =   4575
                  MultiLine       =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   55
                  Tag             =   "Cell Tag Template (After)"
                  ToolTipText     =   "Using this area will cause all images to NOT display their name along"
                  Top             =   300
                  Width           =   4485
               End
               Begin VB.TextBox txtSpecial 
                  Height          =   780
                  Index           =   2
                  Left            =   0
                  MultiLine       =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   54
                  Tag             =   "Cell Tag Template (Before)"
                  ToolTipText     =   "Use %%Filename%% and %%Image%% to completely override."
                  Top             =   300
                  Width           =   4485
               End
            End
            Begin VsOcxLib.VideoSoftElastic vseSpecial 
               Height          =   1095
               Index           =   2
               Left            =   90
               TabIndex        =   75
               Top             =   1275
               Width           =   9060
               _Version        =   327680
               _ExtentX        =   15981
               _ExtentY        =   1931
               _StockProps     =   70
               ConvInfo        =   1418783674
               AutoSizeChildren=   1
               BorderWidth     =   0
               BevelOuter      =   5
               MaxChildSize    =   7600
               TagWidth        =   300
               TagPosition     =   1
               Begin VB.TextBox txtSpecial 
                  Height          =   795
                  Index           =   4
                  Left            =   4575
                  MultiLine       =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   53
                  Tag             =   "Table Tag Template (After)"
                  Top             =   300
                  Width           =   4485
               End
               Begin VB.TextBox txtSpecial 
                  Height          =   795
                  Index           =   1
                  Left            =   0
                  MultiLine       =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   52
                  Tag             =   "Table Tag Template (Before)"
                  Top             =   300
                  Width           =   4485
               End
            End
            Begin VsOcxLib.VideoSoftElastic vseSpecial 
               Height          =   1095
               Index           =   3
               Left            =   90
               TabIndex        =   76
               Top             =   90
               Width           =   9060
               _Version        =   327680
               _ExtentX        =   15981
               _ExtentY        =   1931
               _StockProps     =   70
               Caption         =   "After start of page (at top)"
               ConvInfo        =   1418783674
               AutoSizeChildren=   1
               BorderWidth     =   0
               BevelOuter      =   5
               Style           =   1
               Begin VB.TextBox txtSpecial 
                  CausesValidation=   0   'False
                  Height          =   855
                  Index           =   0
                  Left            =   0
                  MultiLine       =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   51
                  Top             =   240
                  Width           =   9060
               End
            End
         End
         Begin VsOcxLib.VideoSoftElastic VideoSoftElastic6 
            Height          =   4800
            Left            =   -73455
            TabIndex        =   59
            Top             =   330
            Width           =   9240
            _Version        =   327680
            _ExtentX        =   16298
            _ExtentY        =   8467
            _StockProps     =   70
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ConvInfo        =   1418783674
            BevelOuter      =   0
            TagWidth        =   300
            TagPosition     =   1
            Begin VB.CommandButton cmdSetupFolderWalk 
               Caption         =   "Install Folder Walking"
               Height          =   435
               Left            =   6885
               TabIndex        =   50
               Tag             =   "Other:"
               Top             =   390
               Width           =   2145
            End
            Begin VB.TextBox txtCommandLine 
               Appearance      =   0  'Flat
               Height          =   2055
               Left            =   90
               MultiLine       =   -1  'True
               TabIndex        =   60
               Tag             =   "HTML File Walker Senior can take the following command line arguments:"
               Text            =   "frmMain.frx":0462
               Top             =   390
               Width           =   6585
            End
         End
         Begin VsOcxLib.VideoSoftElastic VideoSoftElastic5 
            Height          =   4800
            Left            =   -73755
            TabIndex        =   61
            Top             =   330
            Width           =   9240
            _Version        =   327680
            _ExtentX        =   16298
            _ExtentY        =   8467
            _StockProps     =   70
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ConvInfo        =   1418783674
            BevelOuter      =   0
            TagWidth        =   1500
            Begin VB.CommandButton cmdTalkToTheProgrammer 
               Caption         =   "Talk to the programmer"
               Height          =   435
               Left            =   6750
               TabIndex        =   49
               ToolTipText     =   "Click to send an email to the programmer (Bug report, suggestions, questions, etc.)."
               Top             =   945
               Width           =   2415
            End
            Begin VB.Timer tmrMoveFS 
               Enabled         =   0   'False
               Interval        =   1
               Left            =   8010
               Top             =   480
            End
            Begin VB.CommandButton cmdAndNow 
               Caption         =   "Just for the fun of it..."
               Height          =   435
               Left            =   3735
               TabIndex        =   48
               ToolTipText     =   "Click to start/stop a nice little time waster..."
               Top             =   945
               Width           =   2415
            End
            Begin VB.CommandButton cmdMakeShortcut 
               Caption         =   "Create shortcut to current settings"
               Enabled         =   0   'False
               Height          =   435
               Left            =   270
               TabIndex        =   45
               Top             =   1500
               Width           =   3180
            End
            Begin VB.TextBox txtFileSettings 
               Height          =   315
               Left            =   270
               TabIndex        =   43
               Top             =   540
               Width           =   7590
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "&Save settings"
               Height          =   435
               Index           =   4
               Left            =   270
               TabIndex        =   47
               Top             =   3100
               Width           =   3180
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "&Load settings from..."
               Height          =   435
               Index           =   3
               Left            =   270
               TabIndex        =   44
               Top             =   945
               Width           =   3180
            End
            Begin VB.CommandButton cmdReset 
               Caption         =   "Reset to defaults"
               Height          =   435
               Left            =   270
               TabIndex        =   46
               Top             =   2055
               Width           =   3180
            End
            Begin VB.PictureBox picBack 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               ForeColor       =   &H80000008&
               Height          =   2085
               Left            =   3750
               ScaleHeight     =   4
               ScaleMode       =   0  'User
               ScaleTop        =   -2
               ScaleWidth      =   20
               TabIndex        =   82
               Top             =   1425
               Width           =   5400
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Firm"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   18
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   435
                  Left            =   75
                  MouseIcon       =   "frmMain.frx":0605
                  MousePointer    =   99  'Custom
                  TabIndex        =   84
                  ToolTipText     =   "Click once to walk on over to Firm Solutions' Web Site."
                  Top             =   15
                  Width           =   915
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Solutions"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   18
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   435
                  Left            =   1065
                  MouseIcon       =   "frmMain.frx":0A47
                  MousePointer    =   99  'Custom
                  TabIndex        =   83
                  ToolTipText     =   "Click once to walk on over to Firm Solutions' Web Site."
                  Top             =   15
                  Width           =   1860
               End
            End
            Begin VB.CheckBox chkAutoLoad 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Automatically Load these settings next time."
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   4680
               TabIndex        =   42
               ToolTipText     =   "When this is on, the web page will only be viewable from YOUR hard drive."
               Top             =   270
               Width           =   4335
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Path and filename to store these settings in."
               ForeColor       =   &H00000000&
               Height          =   300
               Index           =   14
               Left            =   270
               TabIndex        =   62
               Top             =   270
               Width           =   3990
            End
         End
         Begin VsOcxLib.VideoSoftElastic VideoSoftElastic4 
            Height          =   4800
            Left            =   -74055
            TabIndex        =   63
            Top             =   330
            Width           =   9240
            _Version        =   327680
            _ExtentX        =   16298
            _ExtentY        =   8467
            _StockProps     =   70
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ConvInfo        =   1418783674
            BevelOuter      =   0
            TagWidth        =   1600
            Begin MabryFtp.mFtp ftpPutPage 
               Left            =   3645
               Top             =   3090
               _ExtentX        =   847
               _ExtentY        =   847
               DebugMode       =   1
               DisablePasv     =   0   'False
            End
            Begin VB.TextBox txtPassword 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   2265
               PasswordChar    =   "*"
               TabIndex        =   38
               Tag             =   "Password:"
               Top             =   1185
               Width           =   4170
            End
            Begin VB.TextBox txtUsername 
               Height          =   300
               Left            =   2265
               TabIndex        =   37
               Tag             =   "User name:"
               Top             =   840
               Width           =   4170
            End
            Begin VB.TextBox txtHostPort 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   2265
               MultiLine       =   -1  'True
               TabIndex        =   36
               Tag             =   "Host Port:"
               Text            =   "frmMain.frx":0E89
               Top             =   510
               Width           =   930
            End
            Begin VB.TextBox txtRemotePath 
               Height          =   300
               Left            =   1890
               TabIndex        =   39
               Tag             =   "Remote directory:"
               Top             =   1800
               Width           =   4545
            End
            Begin VB.TextBox txtHost 
               Height          =   300
               Left            =   1890
               TabIndex        =   35
               Tag             =   "FTP Host:"
               Top             =   180
               Width           =   4545
            End
            Begin VB.CommandButton cmdPut 
               Caption         =   "&Put current page on FTP Host"
               Height          =   435
               Left            =   270
               TabIndex        =   41
               Top             =   3100
               Width           =   3180
            End
            Begin VB.TextBox txtFTPOutputTo 
               Height          =   300
               Left            =   2265
               TabIndex        =   40
               Tag             =   "Remote filename:"
               Top             =   2205
               Width           =   4170
            End
         End
         Begin VsOcxLib.VideoSoftElastic VideoSoftElastic3 
            Height          =   4800
            Left            =   -74355
            TabIndex        =   64
            Top             =   330
            Width           =   9240
            _Version        =   327680
            _ExtentX        =   16298
            _ExtentY        =   8467
            _StockProps     =   70
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ConvInfo        =   1418783674
            BevelOuter      =   0
            TagWidth        =   1500
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Locate"
               Height          =   300
               Index           =   5
               Left            =   7860
               TabIndex        =   33
               Top             =   630
               Width           =   885
            End
            Begin VB.TextBox txtBrowserPath 
               Height          =   300
               Left            =   315
               TabIndex        =   32
               Text            =   "C:\Program Files\Plus!\Microsoft Internet\iexplore.exe"
               Top             =   630
               Width           =   7425
            End
            Begin VB.CommandButton cmdViewPage 
               Caption         =   "&View the current page"
               Height          =   435
               Left            =   270
               TabIndex        =   34
               Top             =   3100
               Width           =   3180
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Path and filename of the program to use to browse, view, and/or edit the web page:"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   10
               Left            =   360
               TabIndex        =   65
               Top             =   360
               Width           =   7155
            End
         End
         Begin VsOcxLib.VideoSoftElastic VideoSoftElastic2 
            Height          =   4800
            Left            =   -74655
            TabIndex        =   66
            Top             =   330
            Width           =   9240
            _Version        =   327680
            _ExtentX        =   16298
            _ExtentY        =   8467
            _StockProps     =   70
            ConvInfo        =   1418783674
            BevelOuter      =   0
            TagWidth        =   1900
            Begin VB.CheckBox chkRenameMP3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Rename MP3"
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   6435
               TabIndex        =   90
               ToolTipText     =   "Currently disabled."
               Top             =   3210
               Width           =   2055
            End
            Begin VB.CheckBox chkProcessMP3Tags 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Process MP3 Tags"
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   6435
               TabIndex        =   31
               ToolTipText     =   "Currently disabled."
               Top             =   2874
               Width           =   2055
            End
            Begin VB.CheckBox chkHideExt 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Hide file extensions"
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   6435
               TabIndex        =   28
               ToolTipText     =   "When this is on, all file extensions will be removed from the Title of each hyperlink."
               Top             =   2538
               Width           =   2055
            End
            Begin VB.CheckBox chkNoRootEntry 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "No desc. on root directory"
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   6435
               TabIndex        =   27
               ToolTipText     =   "When this is on, No description is given for the first entry walked."
               Top             =   2175
               Width           =   2595
            End
            Begin VB.CheckBox chkBuildIndex 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Build index of pages in root"
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   6435
               TabIndex        =   26
               ToolTipText     =   "When this is on, an additional index.html page is built in the Directory to Walk cateloging the sub directory pages."
               Top             =   1845
               Width           =   2685
            End
            Begin VB.CheckBox chkBreakIntoPages 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "One page per Subdirectory"
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   6435
               TabIndex        =   25
               ToolTipText     =   "When this is on, One index.html page is created in each subdirectory walked."
               Top             =   1530
               Width           =   2685
            End
            Begin VB.TextBox txtMaxFilesPerPage 
               Enabled         =   0   'False
               Height          =   288
               Left            =   2190
               TabIndex        =   19
               Text            =   "10"
               Top             =   1815
               Width           =   585
            End
            Begin VB.CheckBox chkAsOf 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Include ""As of"" line."
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   4230
               TabIndex        =   78
               ToolTipText     =   "When this is on, an ""As of"" date/time stamp line is placed under the Web Page's Title."
               Top             =   1845
               Value           =   1  'Checked
               Width           =   2115
            End
            Begin VB.CheckBox chkTable 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Use tables"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   135
               TabIndex        =   16
               ToolTipText     =   $"frmMain.frx":0E8C
               Top             =   1530
               Width           =   1260
            End
            Begin VB.CheckBox chkMultipage 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Break into pages of"
               Enabled         =   0   'False
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   135
               TabIndex        =   18
               ToolTipText     =   "When this is on, multiple pages will be created if more than the indicated number of pictures is encountered."
               Top             =   1845
               Width           =   2010
            End
            Begin VB.TextBox txtCellsWide 
               Height          =   288
               Left            =   1425
               TabIndex        =   17
               Text            =   "5"
               Top             =   1500
               Width           =   408
            End
            Begin VB.CheckBox chkIncludeCredits 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Include credits."
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   4230
               TabIndex        =   77
               ToolTipText     =   "When this is on, a credit for HTML file walker will be added to the bottom of the web page created."
               Top             =   1530
               Value           =   1  'Checked
               Width           =   1680
            End
            Begin VB.CheckBox chkLocal 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Local references."
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   135
               TabIndex        =   23
               ToolTipText     =   "When this is on, the web page will only be viewable from YOUR hard drive."
               Top             =   2520
               Value           =   1  'Checked
               Width           =   1845
            End
            Begin VB.TextBox txtFilePrefix 
               Enabled         =   0   'False
               Height          =   300
               Left            =   3060
               TabIndex        =   24
               Tag             =   "              File Prefix:"
               ToolTipText     =   "What's in here get's placed before every link created."
               Top             =   2475
               Width           =   3240
            End
            Begin VB.CommandButton cmdAdvancedTable 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "&Advanced Table Options"
               Height          =   435
               Left            =   3735
               TabIndex        =   30
               Top             =   3100
               Width           =   2460
            End
            Begin VB.CommandButton cmdGo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "&Build page using these settings"
               Height          =   435
               Index           =   1
               Left            =   270
               TabIndex        =   29
               Top             =   3100
               Width           =   3180
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Locate"
               Height          =   300
               Index           =   1
               Left            =   8130
               TabIndex        =   13
               Top             =   682
               Width           =   800
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Locate"
               Height          =   300
               Index           =   0
               Left            =   8130
               TabIndex        =   11
               Top             =   225
               Width           =   800
            End
            Begin VB.TextBox txtExclude 
               Height          =   300
               Left            =   2025
               TabIndex        =   14
               Tag             =   "Exclude Directory:"
               ToolTipText     =   "This directory and all sub-directories will be excluded when walked."
               Top             =   1140
               Width           =   5985
            End
            Begin VB.TextBox txtBackgroundPic 
               Height          =   300
               Left            =   2025
               TabIndex        =   10
               Tag             =   "Background Picture:"
               ToolTipText     =   "This is the file which appears in the background."
               Top             =   225
               Width           =   5985
            End
            Begin VB.TextBox txtBackgroundSound 
               Height          =   300
               Left            =   2025
               TabIndex        =   12
               Tag             =   "Backgound Sound:"
               ToolTipText     =   "This will be the song played in the background when the page is viewed."
               Top             =   690
               Width           =   5985
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Locate"
               Height          =   300
               Index           =   9
               Left            =   8130
               TabIndex        =   15
               Top             =   1140
               Width           =   800
            End
            Begin VB.TextBox txtMaxY 
               Height          =   288
               Left            =   3480
               TabIndex        =   22
               Text            =   "100"
               ToolTipText     =   "Maximum Image Size (Y)"
               Top             =   2145
               Width           =   585
            End
            Begin VB.TextBox txtMaxX 
               Height          =   288
               Left            =   2505
               TabIndex        =   21
               Text            =   "100"
               ToolTipText     =   "Maximum Image Size (X)"
               Top             =   2145
               Width           =   585
            End
            Begin VB.CheckBox chkLimitSize 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Limit size of pictures to"
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   135
               TabIndex        =   20
               ToolTipText     =   "When this is on, Images included will be specifically sized to the indicated pixel X and Y."
               Top             =   2175
               Width           =   2385
            End
            Begin VB.CheckBox chkSkipSubdirs 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Skip Subdirectories"
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   4230
               TabIndex        =   85
               ToolTipText     =   "Currently disabled."
               Top             =   2175
               Width           =   2175
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "images / files."
               Enabled         =   0   'False
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   12
               Left            =   2835
               TabIndex        =   81
               Top             =   1845
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "cells wide."
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   1890
               TabIndex        =   80
               Top             =   1530
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "by"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   3165
               TabIndex        =   79
               Top             =   2175
               Width           =   210
            End
         End
         Begin VsOcxLib.VideoSoftElastic VideoSoftElastic1 
            Height          =   4800
            Left            =   -74955
            TabIndex        =   67
            Top             =   330
            Width           =   9240
            _Version        =   327680
            _ExtentX        =   16298
            _ExtentY        =   8467
            _StockProps     =   70
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ConvInfo        =   1418783674
            BevelOuter      =   0
            TagWidth        =   300
            TagPosition     =   1
            Begin VB.CheckBox chkAutoSave 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Automatically Save Settings before Build"
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   3735
               TabIndex        =   8
               ToolTipText     =   "When this is on, the web page will only be viewable from YOUR hard drive."
               Top             =   3060
               Width           =   3975
            End
            Begin VB.CheckBox chkAutoView 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Automatically View page after Build"
               ForeColor       =   &H00000000&
               Height          =   228
               Left            =   3735
               TabIndex        =   9
               ToolTipText     =   "When this is on, the web page will only be viewable from YOUR hard drive."
               Top             =   3330
               Value           =   1  'Checked
               Width           =   3615
            End
            Begin VB.TextBox txtOutFilename 
               Height          =   300
               Left            =   315
               TabIndex        =   3
               Tag             =   "Output to:"
               Top             =   1215
               Width           =   6735
            End
            Begin VB.TextBox txtTitle 
               Height          =   300
               Left            =   315
               TabIndex        =   5
               Tag             =   "Web page title:"
               Top             =   2070
               Width           =   6735
            End
            Begin VB.TextBox txtStartPath 
               Height          =   300
               Left            =   315
               TabIndex        =   1
               Tag             =   "Directory to walk:"
               Top             =   450
               Width           =   6735
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Locate"
               Height          =   345
               Index           =   2
               Left            =   7155
               TabIndex        =   6
               Top             =   2040
               Width           =   840
            End
            Begin VB.CommandButton cmdGo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "&Build page using these settings"
               Height          =   435
               Index           =   0
               Left            =   270
               TabIndex        =   7
               Top             =   3100
               Width           =   3180
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Locate"
               Height          =   345
               Index           =   7
               Left            =   7155
               TabIndex        =   2
               Top             =   420
               Width           =   840
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Locate"
               Height          =   345
               Index           =   8
               Left            =   7155
               TabIndex        =   4
               Top             =   1170
               Width           =   840
            End
         End
      End
   End
   Begin MSComDlg.CommonDialog cdgBrowse 
      Left            =   1215
      Top             =   4605
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dCount As Double
Public dDirection As Double
Public dCountPrevious As Double
Public dCountP2 As Double

Public StartTab As Long

Public mp3IDTag As New CMP3IDTag
Public mp3S As Database
Public rst As Recordset

Public Dummy As Boolean

Private Sub BuildDirIndex(ByVal sStartPath As String, ByVal sOutputFilename As String, ByVal sParentDirectory As String)
On Error GoTo EH_BuildDirIndex
    Dim i           As Long
    Dim sRootIndex  As String
    Dim sTemp       As String
    Dim sRelPath    As String
    
    ' Debug.Print  "BuildDirIndex: " & vbTab & sStartPath & vbNewLine & vbTab & sOutputFilename
    
    dirStack.path = vbNullString
    dirStack.path = sStartPath

    sRootIndex = vbNullString
    
    If Right$(sStartPath, 1) <> "\" Then
       sStartPath = sStartPath & "\"
    End If

    HTML_BuildTop sRootIndex
       'If IIf(Right$(sStartPath, 1) = "\", sStartPath, sStartPath & "\") <> IIf(Right$(sParentDirectory, 1) = "\", sParentDirectory, sParentDirectory & "\") Then
        If IIf(Right$(sStartPath, 1) = "\", sStartPath, sStartPath & "\") <> IIf(Right$(sParentDirectory, 1) = "\", sParentDirectory, sParentDirectory & "\") Then
           sRelPath = ".."  '& sAfter(dirStack.List(i), 1, sStartPath)
           If Left$(sRelPath, 1) = "\" Then sRelPath = Mid$(sRelPath, 2)
           sRootIndex = sRootIndex & sHTML_BuildLink(sHTML_Enclose("B", "Go up to Parent Directory: " & sGetFilename(sParentDirectory)), sRelPath & "\" & gsDefaultHTMLFilename) & "<BR>" & vbNewLine
        End If
        For i = 0 To dirStack.ListCount - 1
            If chkBuildIndex.Value <> 0 Then
               sRelPath = sAfter(dirStack.List(i), 1, sStartPath)
               If Left$(sRelPath, 1) = "\" Then sRelPath = Mid$(sRelPath, 2)

               sRootIndex = sRootIndex & sHTML_BuildLink(sHTML_Enclose("B", "Change to Directory: " & sRelPath), sRelPath & "\" & gsDefaultHTMLFilename) & "<BR>" & vbNewLine
            End If
        Next i
        
        gsOut = vbNullString
        BuildDir sStartPath, sStartPath & gsDefaultHTMLFilename, False
        dirStack.path = sStartPath

        sRootIndex = Replace$(sRootIndex & gsOut, vbNewLine & vbNewLine, vbNewLine)
    HTML_BuildBottom sRootIndex

    HTML_SaveFile sOutputFilename, sRootIndex

    sRootIndex = vbNullString
    gsOut = vbNullString
    
    dirStack.path = sStartPath
    For i = 0 To dirStack.ListCount - 1
        BuildDirIndex dirStack.List(i), dirStack.List(i) & "\" & gsDefaultHTMLFilename, sStartPath
            dirStack.path = vbNullString
            dirStack.path = sStartPath
    Next i

EH_BuildDirIndex_Continue:
    Exit Sub
    
EH_BuildDirIndex:
    MsgBox "Error occured (BuildDirIndex):" & vbNewLine & vbTab & Err.Number & vbNewLine & vbTab & Err.Description, vbExclamation
    Resume EH_BuildDirIndex_Continue
    Resume
End Sub

Public Function FileURLPath(ByVal j As Long) As String
    If Len(gsFileURLPath) = 0 Then
       FileURLPath = "file://" & Replace$(File1.path & "\" & File1.List(j), "\", "/")
    Else
       FileURLPath = gsFileURLPath
    End If
End Function

Public Sub HTML_BuildBottom(ByRef sOutput As String)
     If Len(txtSpecial(3)) > 0 Then sOutput = sOutput & txtSpecial(3) & gs2EOL
    
     If chkIncludeCredits.Value <> 0 Then
        sOutput = sOutput & vbNewLine & "<BR><HR><BR>" & sHTML_Enclose("CENTER", sHTML_Enclose("H4", "This page created with " & sHTML_BuildLink("HTML File Walker Senior", gsHTMLFileWalkerHomePage) & " by " & sHTML_BuildLink("William M. Rawls", "http://www.redshift.com/~jon-tom/index.html")))
     End If

     sOutput = sOutput & gs2EOL & "</BODY>" & vbNewLine & "</HTML>"
End Sub

Public Sub HTML_BuildTop(ByRef sOutput As String)
    sOutput = "<HTML>" & sHTML_Enclose("TITLE", txtTitle) & vbNewLine
    If Len(txtBackgroundSound) > 0 Then
       sOutput = sOutput & "<BGSOUND SRC=" & Chr$(34) & txtBackgroundSound & Chr$(34) & ">" & gs2EOL
    End If

    sOutput = sOutput & sHTML_Body(txtBackgroundPic) & gs2EOL

    If Len(txtTitle) > 0 Then
       sOutput = sOutput & vbTab & "<CENTER>" & vbNewLine & vbTab & vbTab & sHTML_Enclose("B", sHTML_Enclose("FONT SIZE=4", txtTitle)) & "<BR>" & vbNewLine
    End If
    
    If chkAsOf.Value <> 0 Then
       sOutput = sOutput & vbTab & vbTab & sHTML_Enclose("FONT SIZE=3", "As of: " & Format(Now, "dddd, mmmm d, yyyy \a\t h:nn AM/PM"), "FONT") & vbNewLine & vbTab & "</CENTER>" & vbNewLine & "<BR>" & vbNewLine '& gs2EOL
    Else
       sOutput = sOutput & vbNewLine & vbTab & "</CENTER>" & vbNewLine & "<BR>" & vbNewLine '& gs2EOL
    End If

    If Len(txtSpecial(0)) > 0 Then
       sOutput = sOutput & txtSpecial(0) & vbNewLine '& gs2EOL
    End If
    
End Sub

Public Sub HTML_SaveFile(sFilename As String, sFileContents As String)
On Error Resume Next
    If FileExists(sFilename) Then
       Kill sFilename
       Err.Clear
    End If

On Error GoTo EH_HTML_SaveFile
    Dim fh As Long
    fh = FreeFile
    Open sFilename For Output Access Write As #fh
        Print #fh, sFileContents
    Close #fh

EH_HTML_SaveFile_Continue:

EH_HTML_SaveFile:
End Sub

Public Sub HTML_ViewWebPage(ByVal sFilename As String, sBrowserPath As String)
    Static x, sT As String

    sT = sFilename

    Screen.MousePointer = 11

    If Len(Trim$(sT)) > 0 Then
       If InStr(sFilename, "\") = 0 And InStr(LCase$(sT), "http://") = 0 Then
          sT = App.path & "\" & Trim$(sT)
       End If
HTML_ViewWebPage_TryAgain:
     If Not FileExists(sBrowserPath) Then
          If MsgBox("The browser/HTML editor does not exist at" & gsEOLTab & "'" & sBrowserPath & "'." & gs2EOL & "Would you like to locate the correct one and try again ?", vbInformation + vbYesNo, "LOCATE BROWSER ?") = vbYes Then
             ' Simulate a click of the "Locate" button for Browser Path selection
             cmdBrowse_Click 5
             GoTo HTML_ViewWebPage_TryAgain
          End If
       ElseIf InStr(LCase$(sT), "http://") > 0 Then
          x = Shell(Trim$(sBrowserPath) & " " & sT, vbNormalFocus)
       ElseIf FileExists(sT) Then
          x = Shell(Trim$(sBrowserPath) & " " & sT, vbNormalFocus)
       Else
          MsgBox "The file '" & sFilename & "' does not exist."
       End If
    End If

    Screen.MousePointer = 0

End Sub

Public Sub ProcessMP3(ByVal sPath As String, ByVal sFile As String)
On Error GoTo EH_frmMain_ProcessMP3
'On Error Resume Next

    If chkRenameMP3.Value = 1 Then RenameMP3 sPath, sFile
    
    mp3IDTag.Read sPath & "\" & sFile

    If mp3S Is Nothing Then '
       Set mp3S = OpenDatabase("c:\mp3catalog.mdb")
       mp3S.Execute "delete from mp3"
       Set rst = mp3S.OpenRecordset("mp3")
    End If

    If Not rst Is Nothing Then
       With rst
            .AddNew
                !mp3Name = mp3IDTag.Filename
                !DateCreated = Now
                !DateModified = Now
               '!Complete = mp3IDTag.Complete
                If Not mp3IDTag.HasData Then
                   sFile = sGetToken(sFile, 1, ".mp3")
                   If InStr(sFile, " - ") Then
                      !Title = sGetToken(sFile, 1, " - ")
                      !Artist = sGetToken(sFile, 2, " - ")
                   Else
                      !Title = sFile
                      !Artist = sFile
                   End If
                Else
                   !Title = mp3IDTag.Title
                   !Artist = mp3IDTag.Artist
                End If
                !Album = mp3IDTag.Album
                !Year = mp3IDTag.Year
                !Comment = mp3IDTag.Comment
                !Genre = mp3IDTag.Genre
            .Update
       End With
    End If

EH_frmMain_ProcessMP3_Continue:
    Exit Sub

EH_frmMain_ProcessMP3:
    LogError "frmMain", "ProcessMP3", Err.Number, Err.Description
    Resume EH_frmMain_ProcessMP3_Continue

    Resume
End Sub

Public Sub RenameMP3(ByVal sPath As String, ByRef sFile As String)
On Error GoTo EH_frmMain_RenameMP3
    Dim sOldFile As String
    Dim nFile As Long
    Dim nLastSpace As Long
    Dim j As Long
    
    sOldFile = sFile
    sFile = Replace$(sFile, "_", " ")
    sFile = Replace$(sFile, "(", vbNullString)
    sFile = Replace$(sFile, ")", vbNullString)
    sFile = Replace$(sFile, "[", vbNullString)
    sFile = Replace$(sFile, "]", vbNullString)
    sFile = Replace$(sFile, "{", vbNullString)
    sFile = Replace$(sFile, "}", vbNullString)
    sFile = Replace$(sFile, "--", "-")
    sFile = Replace$(sFile, "-", "~")
    sFile = Replace$(sFile, "~", " - ")
    sFile = Replace$(sFile, "  -  ", " ~ ")
    sFile = Replace$(sFile, "~", "-")
    sFile = Replace$(sFile, "  ", " - ")
    sFile = Replace$(sFile, " - - ", " - ")
    
    nFile = Len(sFile)
    nLastSpace = 0
    For j = 1 To nFile
        If Asc(Mid$(sFile, j, 1)) >= Asc("a") And Asc(Mid$(sFile, j, 1)) <= Asc("z") And nLastSpace = j - 1 Then
           Mid$(sFile, j, 1) = UCase$(Mid$(sFile, j, 1))
        ElseIf Mid$(sFile, j, 1) = " " Then
           nLastSpace = j
        End If
    Next j

    Name sPath & "\" & sOldFile As sPath & "\" & sFile

EH_frmMain_RenameMP3_Continue:
    Exit Sub

EH_frmMain_RenameMP3:
    LogError "frmMain", "RenameMP3", Err.Number, Err.Description
    Resume EH_frmMain_RenameMP3_Continue

    Resume
End Sub

Public Function sHTML_Body(Optional ByVal sBackgroundPic As String, Optional ByVal sOptions As String)
    If Len(sBackgroundPic) = 0 And Len(sOptions) = 0 Then
        sHTML_Body = "<BODY>"
    ElseIf Len(sOptions) = 0 Then
        sHTML_Body = "<BODY " & sOptions & ">"
    Else
        sHTML_Body = "<BODY BACKGROUND=""" & sBackgroundPic & """>"
    End If
End Function

Public Function sHTML_BuildImage(ByVal sTitle As String, ByVal sLink As String) As String
    If chkLimitSize.Value = 0 Then
        If Len(sTitle) > 0 And Len(txtSpecial(5)) > 0 Then
           sHTML_BuildImage = sHTML_BuildLink("<IMG SRC=""" & sLink & """>", sLink)
        ElseIf Len(sTitle) > 0 Then
           sHTML_BuildImage = sHTML_BuildLink("<IMG SRC=""" & sLink & """><BR>" & sTitle, sLink)
        Else
           sHTML_BuildImage = sHTML_BuildLink("<IMG SRC=""" & sLink & """>", sLink)
        End If
    Else
        If Len(sTitle) > 0 And Len(txtSpecial(5)) > 0 Then
           sHTML_BuildImage = sHTML_BuildLink("<IMG WIDTH=" & txtMaxX & " HEIGHT=" & txtMaxY & " SRC=""" & sLink & """>", sLink)
        ElseIf Len(sTitle) > 0 Then
           sHTML_BuildImage = sHTML_BuildLink("<IMG WIDTH=" & txtMaxX & " HEIGHT=" & txtMaxY & " SRC=""" & sLink & """><BR>" & sTitle, sLink)
        Else
           sHTML_BuildImage = sHTML_BuildLink("<IMG WIDTH=" & txtMaxX & " HEIGHT=" & txtMaxY & " SRC=""" & sLink & """>", sLink)
        End If
    End If
End Function

Public Function sHTML_BuildMP3(ByVal sTitle As String, ByVal sLink As String) As String
    Dim sT As String

    If StrComp(Right$(LCase$(sTitle), 4), ".mp3") = 0 Then sTitle = Left$(sTitle, Len(sTitle) - 4)

    If InStr(sTitle, " - ") Then
       sT = sTitle
       If StrComp(sGetToken(sT, 1, " - "), "Unknown") <> 0 Then
          sTitle = "'" & sGetToken(sT, 2, " - ") & "' by " & sGetToken(sT, 1, " - ")
       Else
          sTitle = "'" & sGetToken(sT, 2, " - ") & "' by an unknown artist"
       End If
    End If

    sHTML_BuildMP3 = sHTML_BuildLink(sTitle, sLink)
End Function


Public Function sHTML_BuildLink(ByVal sTitle As String, ByVal sLink As String)
    Static lLastFound As Long

    If Left$(sTitle, 1) = "\" Then sTitle = Mid$(sTitle, 2)

    lLastFound = InStr(sTitle, ".")

    If chkHideExt.Value = 0 Or lLastFound = 0 Then
       sHTML_BuildLink = "<A HREF = """ & sLink & """>" & sTitle & "</A>"
    Else
       If Len(sGetToken(sTitle, nTokens(sTitle, "."), ".")) < 5 Then
          sHTML_BuildLink = "<A HREF = """ & sLink & """>" & sBefore(sTitle, nTokens(sTitle, "."), ".") & "</A>"
       Else
          sHTML_BuildLink = "<A HREF = """ & sLink & """>" & sTitle & "</A>"
       End If
    End If
End Function

Public Function sHTML_Enclose(ByVal sTag As String, ByVal sText As String, Optional ByVal sEndTag As String)
    If Len(sEndTag) = 0 Then sEndTag = sTag

    If (UCase$(sTag) = "LI" Or UCase$(sEndTag) = "FONT") And chkTable.Value <> 0 Then
       iNextCellAcross = iNextCellAcross + 1
       If iNextCellAcross > Val(txtCellsWide) Then
          If Len(txtSpecial(2)) > 0 Or Len(txtSpecial(5)) > 0 Then
             sHTML_Enclose = vbNewLine & vbTab & vbTab & "</TR><TR>" & vbNewLine & vbTab & vbTab & vbTab & txtSpecial(2) & sText & txtSpecial(5)
          ElseIf Len(txtSpecial(2)) > 0 Then
             sHTML_Enclose = vbNewLine & vbTab & vbTab & "</TR><TR>" & vbNewLine & vbTab & vbTab & vbTab & txtSpecial(2) & sText & "</TD>"
          ElseIf frmAdvTable.chkSpreadEvenly.Value = 0 Then
             sHTML_Enclose = vbTab & vbTab & "</TR><TR>" & vbNewLine & vbTab & vbTab & vbTab & "<TD>" & sText & "</TD>"
          Else
             sHTML_Enclose = vbTab & vbTab & "</TR><TR>" & vbNewLine & vbTab & vbTab & vbTab & "<TD WIDTH=""" & CStr(100 \ txtCellsWide) & "%"">" & sText & "</TD>"
          End If
'          ElseIf frmAdvTable.chkSpreadEvenly.Value = 0 Then
'             sHTML_Enclose = vbNewLine & vbTab & vbTab & "</TR><TR>" & vbNewLine & vbTab & vbTab & vbTab & "<TD>" & sText & "</TD>"
'          Else
'             sHTML_Enclose = vbNewLine & vbTab & vbTab & "</TR><TR>" & vbNewLine & vbTab & vbTab & vbTab & "<TD WIDTH=""" & CStr(100 \ txtCellsWide) & "%"">" & sText & "</TD>"
'          End If
          iNextCellAcross = 1
       Else
          If Len(txtSpecial(2)) > 0 Or Len(txtSpecial(5)) > 0 Then
             sHTML_Enclose = vbNewLine & vbTab & vbTab & vbTab & txtSpecial(2) & sText & txtSpecial(5)
          ElseIf Len(txtSpecial(2)) > 0 Then
             sHTML_Enclose = vbNewLine & vbTab & vbTab & vbTab & txtSpecial(2) & sText & "</TD>"
          ElseIf frmAdvTable.chkSpreadEvenly.Value = 0 Then
             sHTML_Enclose = vbTab & vbTab & vbTab & "<TD>" & sText & "</TD>"
          Else
             sHTML_Enclose = vbTab & vbTab & vbTab & "<TD WIDTH=""" & CStr(100 \ txtCellsWide) & "%"">" & sText & "</TD>"
          End If
'          ElseIf frmAdvTable.chkSpreadEvenly.Value = 0 Then
'             sHTML_Enclose = vbNewLine & vbTab & vbTab & vbTab & "<TD>" & sText & "</TD>"
'          Else
'             sHTML_Enclose = vbNewLine & vbTab & vbTab & vbTab & "<TD WIDTH=""" & CStr(100 \ txtCellsWide) & "%"">" & sText & "</TD>"
'          End If
       End If
    ElseIf InStr(sTag, "<") Or InStr(sTag, ">") Then
       sHTML_Enclose = sTag & sText & "</" & sEndTag & ">"
    Else
       sHTML_Enclose = "<" & sTag & ">" & sText & "</" & sEndTag & ">"
    End If
End Function

Public Function sHTML_TableSettings() As String
    With frmAdvTable
         If Len(txtSpecial(1)) > 0 Then
            sHTML_TableSettings = txtSpecial(1)
         ElseIf StrComp(.vseBackground.Caption, "Clear") = 0 Then
            sHTML_TableSettings = "<TABLE WIDTH=""" & .txtWidth & "%"" BORDER=" & .txtBorderWidth & " cellspacing=" & .txtCellSpacing & " cellpadding=" & .txtCellPadding & ">"
         Else
            sHTML_TableSettings = "<TABLE BGCOLOR=""#" & Hex(.vseBackground.BackColor) & """ WIDTH=""" & .txtWidth & "%"" BORDER=" & .txtBorderWidth & " cellspacing=" & .txtCellSpacing & " cellpadding=" & .txtCellPadding & ">"
         End If
    End With
End Function

Public Sub BuildDir(ByVal sStartPath As String, ByVal sOutFilename As String, Optional ByVal bActuallyMakeFile As Boolean = True)
On Error Resume Next
    Dim i   As Long

    ' Debug.Print  "BuildDir: " & vbTab & sStartPath & vbNewLine & vbTab & sOutFilename
    
    If bActuallyMakeFile Then
       If FileExists(sOutFilename) Then
          Kill sOutFilename
       End If
    End If

On Error GoTo EH_BuildDir
    If bActuallyMakeFile Then HTML_BuildTop gsOut
        glInterruptBuild = False
        iNextCellAcross = 0
        DirProcess sStartPath, bActuallyMakeFile                ' Process this directory (and every subdirectory if appropriate)
    If bActuallyMakeFile Then HTML_BuildBottom gsOut

    If bActuallyMakeFile Then HTML_SaveFile sOutFilename, gsOut

    If bActuallyMakeFile Then
       If chkBreakIntoPages.Value <> 0 Then
          dirStack.path = vbNullString
          dirStack.path = sStartPath
          For i = 0 To dirStack.ListCount - 1
              BuildDir dirStack.List(i), dirStack.List(i) & "\" & gsDefaultHTMLFilename
              dirStack.path = sStartPath
          Next i
       End If
    End If

    If Not mp3S Is Nothing Then
       rst.Close
       Set rst = Nothing
       mp3S.Close
       Set mp3S = Nothing
    End If

EH_BuildDir_Continue:
    Exit Sub

EH_BuildDir:
    MsgBox "Error occured in 'BuildDir' while processing '" & sStartPath & "' :" & vbNewLine & vbTab & Err.Number & vbNewLine & vbTab & Err.Description, vbExclamation
    Resume EH_BuildDir_Continue
    Resume
End Sub


Public Sub DirProcess(ByVal sDir As String, ByVal bActuallyMakeFile As Boolean)
On Error GoTo EH_frmMain_DirProcess
    Dim CurrDirectory   As Long
    Dim CurrFile        As Long
    Dim fh              As Long
    Dim sStr            As String
    Dim sFile           As String
    Dim gsLinePrefix    As String
    Dim sFormatBegin    As String
    Dim sFormatEnd      As String
    Dim sCurrExtension  As String

    If LCase$(sGetFilename(sDir)) = LCase$(txtExclude) And Len(txtExclude) > 0 Then
       Exit Sub
    ElseIf glInterruptBuild = True Then
       Exit Sub
    End If

    sFormatBegin = "<H2><B>"
    sFormatEnd = "</H2></B>"

    If chkNoRootEntry.Value = 0 Then
        If chkTable.Value = 0 Then
           gsOut = gsOut & vbTab & sFormatBegin & "Contents of Directory: " & sCapitalize(sGetFilename(sDir)) & sFormatEnd & vbNewLine & vbTab & "<UL>" & vbNewLine
        Else
           iNextCellAcross = 0
           gsOut = gsOut & vbTab & sFormatBegin & "Contents of Directory: " & sCapitalize(sGetFilename(sDir)) & sFormatEnd & vbNewLine & vbTab & sHTML_TableSettings() & vbNewLine & vbTab & vbTab & "<TR>" & vbNewLine
        End If
    Else
        If chkTable.Value = 0 Then
           gsOut = gsOut & vbTab & sFormatBegin & "Table of Contents" & vbNewLine & vbTab & sFormatEnd & "<UL>" & vbNewLine
        Else
           iNextCellAcross = 0
           gsOut = gsOut & vbTab & sFormatBegin & "Table of Contents" & vbNewLine & vbTab & sHTML_TableSettings() & vbNewLine & vbTab & vbTab & sFormatEnd & "<TR>" & vbNewLine
        End If
    End If

    dirStack.path = vbNullString
    dirStack.path = sDir

    fh = FreeFile
    For CurrFile = 0 To File1.ListCount - 1
        If glInterruptBuild = True Then Exit Sub

        sCurrExtension = UCase$(sGetExtension(File1.path & "\" & File1.List(CurrFile)))

        Select Case sCurrExtension
             ' Internet Shortcut
               Case "URL"
                    Open File1.path & "\" & File1.List(CurrFile) For Input Access Read As #fh
                         Do Until EOF(fh)
                            Line Input #fh, sStr
                            If UCase$(Left$(sStr, 3)) = "URL" Then
                               sStr = Mid$(sStr, 5)
                               sFile = Left$(File1.List(CurrFile), InStr(UCase$(File1.List(CurrFile)), ".URL") - 1)
                               gsOut = gsOut & gsLinePrefix & sHTML_Enclose("LI", (sHTML_BuildLink(sFile, sStr))) & gsLineSuffix
                            End If
                         Loop
                    Close #fh

            ' Inline image
              Case "BMP", "JPG", "JPEG", "GIF", "TIF", "TIFF", "DIB", "RLE", "JPE", "JIF", "JFIF", "WMF", "PCX", "DCX", "TGA"
                    If chkLocal.Value = 1 Then
                       sFile = FileURLPath(CurrFile)
                    Else
                       sFile = Replace$(txtFilePrefix & sAfter(File1.path & "\" & File1.List(CurrFile), 1, IIf(Right$(sDir, 1) = "\", sDir, sDir & "\")), "\", "/")
                      'sFile = Replace$(txtFilePrefix & sAfter(File1.path & "\" & File1.List(CurrFile), 1, gsStartPath & "\"), "\", "/") 'sAfter(File1.path & "\" & File1.List(CurrFile), 1, sGetPath(txtOutFilename)) 'sAfter(File1.path & "\" & File1.List(CurrFile), 1, sGetPath(txtOutFilename))
                       If Left$(sFile, 1) = "/" Then sFile = Mid$(sFile, 2)
                    End If
                    gsOut = gsOut & gsLinePrefix & sHTML_Enclose("LI", (sHTML_BuildImage(File1.List(CurrFile), sFile))) & gsLineSuffix

            ' MP3s
              Case "MP3"
                    If chkProcessMP3Tags.Value = 1 Then ProcessMP3 File1.path, File1.List(CurrFile)
                    If chkLocal.Value = 1 Then
                       sFile = FileURLPath(CurrFile)
                    Else
                       sFile = Replace$(txtFilePrefix & sAfter(File1.path & "\" & File1.List(CurrFile), 1, IIf(Right$(sDir, 1) = "\", sDir, sDir & "\")), "\", "/")
                      'sFile = Replace$(txtFilePrefix & sAfter(File1.path & "\" & File1.List(CurrFile), 1, gsStartPath & "\"), "\", "/") 'sAfter(File1.path & "\" & File1.List(CurrFile), 1, sGetPath(txtOutFilename))
                       If Left$(sFile, 1) = "/" Then sFile = Mid$(sFile, 2)
                    End If
                    
                    gsOut = gsOut & gsLinePrefix & sHTML_Enclose("LI", (sHTML_BuildMP3(File1.List(CurrFile), sFile))) & gsLineSuffix

            ' File types to skip
              Case "HFW" ', "XLS"
                  ' Don't do anything

              Case "FRX", "DLL", "OCX", "EXP", "LIB", "SWT", "PDM", "VBZ", "VBR", "LIF", "SWT", "REG", "INF", "SYS", "BIN", "CFG", "INI", "OCA", "DDF", "CNT", "PDB", "ISS", "ION", "INS", "BAT", "VAS", "CTX"
                    If chkLocal.Value = 1 Then
                       sFile = FileURLPath(CurrFile)
                    Else
                       sFile = Replace$(txtFilePrefix & sAfter(File1.path & "\" & File1.List(CurrFile), 1, IIf(Right$(sDir, 1) = "\", sDir, sDir & "\")), "\", "/")
                       If Left$(sFile, 1) = "/" Then sFile = Mid$(sFile, 2)
                    End If
                    If sFile <> gsDefaultHTMLFilename Then
                       gsOut = gsOut & gsLinePrefix & sHTML_Enclose("LI", sFile) & gsLineSuffix
                    End If

            ' Any other files
              Case Else
                    If StrComp(Right$(sCurrExtension, 1), "_") <> 0 Then
                       If chkLocal.Value = 1 Then
                          sFile = FileURLPath(CurrFile)
                       Else
                          sFile = Replace$(txtFilePrefix & sAfter(File1.path & "\" & File1.List(CurrFile), 1, IIf(Right$(sDir, 1) = "\", sDir, sDir & "\")), "\", "/")
                          If Left$(sFile, 1) = "/" Then sFile = Mid$(sFile, 2)
                       End If
                       If sFile <> gsDefaultHTMLFilename Then
                          gsOut = gsOut & gsLinePrefix & sHTML_Enclose("LI", (sHTML_BuildLink(File1.List(CurrFile), sFile))) & gsLineSuffix
                       End If
                    End If
        End Select
    Next CurrFile

    If File1.ListCount < 1 Then
       gsOut = gsOut & gsLinePrefix & sHTML_Enclose("LI", "(no files)") & gsLineSuffix
    ElseIf File1.ListCount = 1 And (StrComp(LCase$(File1.List(0)), LCase$(gsDefaultHTMLFilename)) = 0) Then
       gsOut = gsOut & gsLinePrefix & sHTML_Enclose("LI", "(no files)") & gsLineSuffix
    End If

    If dirStack.ListCount > 0 And bActuallyMakeFile Then
       If chkBreakIntoPages.Value = 0 Then
          For CurrDirectory = 0 To dirStack.ListCount - 1
              DirProcess dirStack.List(CurrDirectory), bActuallyMakeFile
              dirStack.path = sDir
          Next CurrDirectory
       Else
          For CurrDirectory = 0 To dirStack.ListCount - 1
              BuildDirIndex dirStack.List(CurrDirectory), dirStack.List(CurrDirectory) & "\" & gsDefaultHTMLFilename, sDir
              dirStack.path = sDir
          Next CurrDirectory
       End If
    End If

    If chkTable.Value = 0 Then
       gsOut = gsOut & vbTab & "</UL>" & gs2EOL
    Else
       If Len(frmMain.txtSpecial(4)) > 0 Then
          gsOut = gsOut & vbTab & vbTab & "</TR>" & vbNewLine & vbTab & frmMain.txtSpecial(4) & gs2EOL
       Else
          gsOut = gsOut & vbTab & vbTab & "</TR>" & vbNewLine & vbTab & "</TABLE>" & gs2EOL
       End If
    End If

EH_frmMain_DirProcess_Continue:
    Exit Sub

EH_frmMain_DirProcess:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: frmMain" & vbCr & vbTab & "Procedure: DirProcess" & vbCr & vbCr & Err.Description
    Resume EH_frmMain_DirProcess_Continue

    Resume
End Sub


Public Function bCheckCommandAndRemove(sToCheck As String, ByVal sOptionName As String, sReturn As String, bPauseAfter As Boolean) As Boolean
    Dim sUOptionName As String, sUToCheck As String
    Dim nStart As Long, nEnd As Long

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
       sReturn = vbNullString
       bCheckCommandAndRemove = True
    End If
End Function

Public Sub HandleCommandLineActions(ByVal CommandLineToHandle As String)
       Dim sCommandLine As String
       Dim sCommandLineParam As String

       sCommandLine = CommandLineToHandle

       gbQuiet = bCheckCommandAndRemove(sCommandLine, "-Quiet", sCommandLineParam, True)

       If bCheckCommandAndRemove(sCommandLine, "-Build", sCommandLineParam, True) Then
          cmdGo_Click 0
          DoEvents: DoEvents
       End If

       If bCheckCommandAndRemove(sCommandLine, "-ExploreLocal", sCommandLineParam, True) Then
        ' AutoExplore Shells out with
        '   Explorer.exe /e,/root,/idlist,<PATH>
          cmdViewPage_Click
          DoEvents: DoEvents
       End If

       If bCheckCommandAndRemove(sCommandLine, "-ExploreRemote", sCommandLineParam, True) Then
          HTML_ViewWebPage (sCommandLineParam), (txtBrowserPath)
          DoEvents: DoEvents
       End If

       If bCheckCommandAndRemove(sCommandLine, "-Exit", sCommandLineParam, True) Then
         'MsgBox "About to exit"
          cmdExit_Click
          DoEvents: DoEvents
       End If

End Sub

Private Sub chkLimitSize_Click()
    txtMaxX.Enabled = (chkLimitSize.Value <> 0)
    txtMaxY.Enabled = (chkLimitSize.Value <> 0)
    VideoSoftElastic2.Refresh
End Sub

Private Sub chkLocal_Click()
    txtFilePrefix.Enabled = True '(chkLocal.Value = 0)
    VideoSoftElastic2.Refresh
End Sub

Private Sub chkMultipage_Click()
    txtMaxFilesPerPage.Enabled = (chkMultipage.Value <> 0)
End Sub

Private Sub chkTable_Click()
    txtCellsWide.Enabled = (chkTable.Value <> 0)
    cmdAdvancedTable.Enabled = (chkTable.Value <> 0)
End Sub

Private Sub cmdAdvancedTable_Click()
    frmAdvTable.Show vbModal, Me
End Sub

Private Sub cmdAndNow_Click()
    tmrMoveFS.Enabled = Not tmrMoveFS.Enabled
    If dDirection = 0 Then dDirection = 1
End Sub

Public Sub cmdBrowse_Click(Index As Integer)
On Error Resume Next
    With cdgBrowse
        Select Case Index
               Case 0 ' Select Background Picture
                    .DialogTitle = "Select which PICTURE to display in the background of the web page."
                        .Filename = txtBackgroundPic
                        .DefaultExt = ".gif"
                        .Filter = "GIFs (*.gif)|*.gif|JPEGs (*.jpg;*.jpeg)|*.jpg;*.jpeg|All files (*.*)|*.*"
                        .FilterIndex = 1
                        Err.Clear
                        .ShowOpen
                        If Err.Number = 20477 Then
                           .Filename = vbNullString
                           Err.Clear
                           .ShowOpen
                        End If
                    If sGetPath(.Filename) = App.path & "\" Then
                       txtBackgroundPic = sGetFilename(.Filename)
                    Else
                       txtBackgroundPic = .Filename
                    End If
    
               Case 1 ' Select Background Sound
                    .DialogTitle = "Select the MIDI or WAV to play in the background."
                        .Filename = txtBackgroundSound
                        .DefaultExt = ".mid"
                        .Filter = "MIDI Songs (*.mid)|*.mid|WAVs  (*.wav)|*.wav|All files (*.*)|*.*"
                        .FilterIndex = 1
                        .ShowOpen
                    If sGetPath(.Filename) = App.path & "\" Then
                       txtBackgroundSound = sGetFilename(.Filename)
                    Else
                       txtBackgroundSound = .Filename
                    End If
    
               Case 2 ' Select web page location
                    .DialogTitle = "Select where you would like the web page saved to."
                        .Filename = txtOutFilename
                        .DefaultExt = ".htm"
                        .Filter = "Web pages (*.htm or *.html)|*.htm;*.html|Text file (*.txt)|*.txt|All files (*.*)|*.*"
                        .FilterIndex = 1
                        .ShowSave
                        If Err.Number = 20477 Then
                           .Filename = vbNullString
                           Err.Clear
                           .ShowOpen
                        End If
                    If sGetPath(.Filename) = App.path & "\" Then
                       txtOutFilename = sGetFilename(.Filename)
                    Else
                       txtOutFilename = .Filename
                    End If
    
               Case 3 ' Load HTML File Walker Settings
                    .DialogTitle = "Select where you would like these HTML File Walker settings loaded from."
                        .Filename = txtFileSettings
                        .DefaultExt = ".hfw"
                        .Filter = "HTML File Walker Data Settings (*.hfw)|*.hfw|Text file (*.txt)|*.txt|All files (*.*)|*.*"
                        .FilterIndex = 1
                        
                        .ShowOpen
                        If Err.Number = 20477 Then
                           .Filename = vbNullString
                           Err.Clear
                           .ShowOpen
                        End If
                    If sGetPath(.Filename) = App.path & "\" Then
                       txtFileSettings = sGetFilename(.Filename)
                    Else
                       txtFileSettings = .Filename
                    End If
                    LoadSettings txtFileSettings
                    vstMain.CurrTab = 1

               Case 4 ' Save HTML File Walker Settings
                    .DialogTitle = "Select where you would like these HTML File Walker settings saved to."
                    .Filename = txtFileSettings
                    .DefaultExt = ".hfw"
                    .Filter = "HTML File Walker Data Settings (*.hfw)|*.hfw|Text file (*.txt)|*.txt|All files (*.*)|*.*"
                    .FilterIndex = 1
                    .ShowSave
                    If sGetPath(.Filename) = App.path & "\" Then
                       txtFileSettings = sGetFilename(.Filename)
                    Else
                       txtFileSettings = .Filename
                    End If
                    SaveSettings txtFileSettings
                    vstMain.CurrTab = 1
    
               Case 5 ' Select browser/editor location
                    .DialogTitle = "Please locate the browser or HTML editor to view web pages with."
                        .Filename = txtBrowserPath
                        .DefaultExt = ".exe"
                        .Filter = "Executable (*.exe or *.bat)|*.exe;*.bat|All files (*.*)|*.*"
                        .FilterIndex = 1
                        .ShowOpen
                        If Err.Number = 20477 Then
                           .Filename = vbNullString
                           Err.Clear
                           .ShowOpen
                        End If
                        txtBrowserPath = .Filename
                        SaveSetting App.ProductName, "Settings", "Browser Path", txtBrowserPath
    
               Case 7 ' Choose a Directory to walk
                    txtStartPath = BrowseForFolder(hWnd, "Plese locate the directory to 'Walk'." & vbNewLine & vbTab & "All files will be processed.")
    
               Case 8 ' Choose a file to output to
                         .DialogTitle = "Please locate the file to write the web page to."
                         .Filename = txtOutFilename
                         .DefaultExt = ".htm"
                         .Filter = "Web Page (*.html or *.htm)|*.htm;*.html|All files (*.*)|*.*"
                         .FilterIndex = 1
                         .ShowSave
                         txtOutFilename = .Filename
    
               Case 9 ' Choose a Directory to Exclude
                    txtExclude = BrowseForFolder(hWnd, "Plese locate the directory to EXCLUDE from 'Walking'.")
    
        End Select
    End With
    
    Err.Clear

End Sub


Private Sub cmdExit_Click()
    Dim Cancel As Integer
    SaveSetting App.ProductName, "Settings", "AutoLoad", IIf(chkAutoLoad.Value <> 0, "True", "False")
    If chkAutoLoad.Value <> 0 Then
       SaveSetting App.ProductName, "Settings", "AutoLoadFilename", txtFileSettings
    End If
    SaveSetting App.ProductName, "Settings", "Browser Path", txtBrowserPath
    
    Form_Unload Cancel
    If Cancel = 0 Then
       End 'Hide
    End If
End Sub

Private Sub cmdGo_Click(Index As Integer)
On Error GoTo EH_cmdGo_Click
    Dim i           As Long
    Dim sRootIndex  As String
    Dim sTemp       As String
    Dim sRelPath    As String
    
    sTemp = sGetToken(txtOutFilename, lTokenCount(txtOutFilename, "\"), "\")
    If Len(sTemp) = 0 Then sTemp = "index.html"

    If chkBreakIntoPages.Value <> 0 Then
       gsDefaultHTMLFilename = InputBox$("For each sub-directory under '" & txtStartPath & "'," & vbNewLine & vbTab & "What should the name of the index HTML file be in each directory ?", "CHOOSE SUB-DIR FILENAME", sTemp)
       If Len(gsDefaultHTMLFilename) = 0 Then Exit Sub
    Else
       gsDefaultHTMLFilename = sTemp
    End If

    If chkLocal.Value = 1 Then
       gsFileURLPath = txtFilePrefix
    Else
       gsFileURLPath = vbNullString
    End If
    
    For i = 0 To 5
        txtSpecial(i) = Trim$(txtSpecial(i))
    Next i

    Hide

    If Not gbQuiet Then
       frmStatus.Show
       DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
    End If

    If chkAutoSave.Value <> 0 And Len(txtFileSettings) > 0 Then
       SaveSettings txtFileSettings
    End If

    If Right$(txtStartPath, 1) <> "\" Then
       gsStartPath = txtStartPath
    Else
       gsStartPath = Left$(txtStartPath, Len(txtStartPath) - 1)
    End If

    If chkBreakIntoPages.Value = 0 Then
       BuildDir txtStartPath, txtOutFilename
    Else
       BuildDirIndex txtStartPath, txtOutFilename, txtStartPath
    End If

    If Not gbQuiet Then
       frmStatus.Hide
       Show
       vstMain.CurrTab = 2
    End If

    If chkAutoView.Value <> 0 Then
       cmdViewPage_Click
    End If

EH_cmdGo_Click_Continue:
    Exit Sub
    
EH_cmdGo_Click:
    MsgBox "Error occured:" & vbNewLine & vbTab & Err.Number & vbNewLine & vbTab & Err.Description, vbExclamation
    Resume EH_cmdGo_Click_Continue
    Resume
End Sub

Private Sub cmdPut_Click()
    With ftpPutPage
        txtHost = Trim$(txtHost)
        If Len(txtHost) > 0 Then
           If Val(txtHost) > 0 Then
              .Host = txtHost
           Else
              .Host = txtHost
           End If
        Else
           MsgBox "Please supply the host name or address in the 'FTP Host' area." & vbNewLine & vbTab & "This will be the Internet name (such as stinger.redshift.com)" & vbNewLine & vbTab & "or IP Address (such as 123.123.123.4)" & vbNewLine & vbTab & "of the computer where you would like the page put."
           Exit Sub
        End If
    
        txtHostPort = Trim$(txtHostPort)
        If Val(txtHostPort) > 0 Then
           .Port = txtHostPort
        Else
           MsgBox "Please supply the host name or address." & vbNewLine & vbTab & "This will be the Internet name (such as stinger.redshift.com)" & vbNewLine & vbTab & "or IP Address (such as 123.123.123.4)" & vbNewLine & vbTab & "of the computer where you would like the page put."
           Exit Sub
        End If
    
        txtUsername = Trim$(txtUsername)
        If Len(txtUsername) > 0 Then
           .LogonName = txtUsername
        Else
           MsgBox "Please supply the user to connect as in the 'User name' area."
           Exit Sub
        End If
    
        txtPassword = Trim$(txtPassword)
        If Len(txtPassword) > 0 Then
           .LogonPassword = txtPassword
        Else
           MsgBox "Please supply the password for the user account in the 'Password' area."
           Exit Sub
        End If
    
        txtRemotePath = Trim$(txtRemotePath)
        If Len(txtRemotePath) > 0 Then
            .Directory = txtRemotePath
        Else
           MsgBox "Please supply the directory on the Host to store the page in the 'Remote path' area."
           Exit Sub
        End If
    
        txtFTPOutputTo = Trim$(txtFTPOutputTo)
        If Len(txtFTPOutputTo) > 0 Then
           .DstFilename = txtFTPOutputTo
        Else
           MsgBox "Please supply the filename on the Host to store the page in the 'Remote filename' area."
           Exit Sub
        End If
    
        txtOutFilename = Trim$(txtOutFilename)
        If Len(txtOutFilename) > 0 Then
           .SrcFilename = txtOutFilename
        Else
           MsgBox "Please supply the filename to store the page in the 'Remote filename' area."
           Exit Sub
        End If

On Error Resume Next
        frmMain.MousePointer = vbHourglass
            .Connect
                If Err.Number <> 0 Then
                   frmMain.MousePointer = vbDefault
                   MsgBox "Error connecting to remote host. Please check your settings and try again.", vbCritical
                   Exit Sub
                End If
            .ChangeDir
                If Err.Number <> 0 Then
                   frmMain.MousePointer = vbDefault
                   MsgBox "Error changing to the remote directory. Please check your settings and try again.", vbCritical
                   .Disconnect
                   Exit Sub
                End If
            .PutFile
                If Err.Number <> 0 Then
                   frmMain.MousePointer = vbDefault
                   MsgBox "Error putting the file. Please check your settings and try again.", vbCritical
                   .Disconnect
                   Exit Sub
                End If
            .Disconnect
        frmMain.MousePointer = vbDefault
        MsgBox "Page stored on remote site.", vbInformation
    End With
End Sub


Public Sub cmdReset_Click()
    txtTitle = "My Favorites"
    txtBackgroundPic = "aged-paper.jpg"
    txtBackgroundSound = "Cursum3.mid"
    txtStartPath = "c:\windows\favorites"
    txtExclude = vbNullString
    txtOutFilename = "favorites.html"
    txtBrowserPath = "C:\Program Files\Plus!\Microsoft Internet\iexplore.exe"
    chkAsOf.Value = 1
End Sub


Private Sub cmdSetupFolderWalk_Click()
On Error Resume Next
    Dim fh As Long
    fh = FreeFile
    
    Open App.path & "\walkit.reg" For Output Access Write As #fh
         Print #fh, "REGEDIT4"
         Print #fh, vbNullString
         Print #fh, "[HKEY_CLASSES_ROOT\Directory\shell\Walk It]"
         Print #fh, vbNullString
         Print #fh, "[HKEY_CLASSES_ROOT\Directory\shell\Walk It\command]"
         Print #fh, "@=""\""" & Replace$(Replace$(App.path, "\", "/"), "/", "\\") & "\\HTML File Walker Senior.exe\"" -Dir(%1)"""
         Print #fh, vbNullString
    Close #fh

    Err.Clear
    Shell sGetWindowsSysDir() & "regsvr32.exe /s """ & App.path & "\walkit.reg"""
    If Err.Number = 0 Then
       MsgBox "Directory Walking installed. Right click a folder from Explorer and select 'Walk It' to test this functionality."
    Else
       MsgBox "The registry file could not be incorporated. If the executable in on a network drive, try copying it to your local hard drive." & gs2EOL & "Alternatively, locate and double click on the file named:" & gsEOLTab & vbTab & App.path & "\walkit.reg"
    End If
End Sub

Private Sub cmdTalkToTheProgrammer_Click()
    Dim sMAILTOCommandLine As String
    Dim ER As New CSetting

    ER.KeyPath = "HKEY_CLASSES_ROOT\mailto\shell\open\command\"
    sMAILTOCommandLine = ER.Value

    If Len(sMAILTOCommandLine) > 0 Then
       sMAILTOCommandLine = Replace$(sMAILTOCommandLine, "%1", "wrawls@firmsolutions.com")
       Shell sMAILTOCommandLine, vbNormalFocus
    ElseIf bUserSure("I cannot determine what your default mailto handler is." & vbCr & "Please manually open your e-mail software and send e-mail to:" & vbCr & vbCr & vbTab & "wrawls@firmsolutions.com" & vbCr & vbCr & "Would you like me to copy that e-mail address to the clipboard now ?", "COPY EMAIL ADDRESS ?") Then
       Clipboard.Clear
       Clipboard.SetText "wrawls@firmsolutions.com"
    End If
End Sub

Private Sub cmdViewPage_Click()

    If Not gbPageGenerated Then
    End If
    
    HTML_ViewWebPage (txtOutFilename), (txtBrowserPath)
    vstMain.CurrTab = 1
End Sub

Private Sub dirStack_Change()
    File1.path = dirStack.path
End Sub


Private Sub Drive1_Change()
    dirStack.path = Drive1
End Sub


Private Sub Form_Load()
On Error Resume Next

   'StartTab = 4
    StartTab = 0

    gs2EOL = vbNewLine & vbNewLine
    gsEOLTab = vbNewLine & vbTab
    gsLinePrefix = vbTab & vbTab
    gsLineSuffix = vbNewLine

    Caption = App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    txtBrowserPath.Text = GetSetting(App.ProductName, "Settings", "Browser Path", vbNullString)

    If Len(Command) > 0 Then
       HandleCommandLineVars
       DoEvents
    Else
       chkAutoLoad.Value = IIf(GetSetting(App.ProductName, "Settings", "AutoLoad", "False") = "True", 1, 0)
       If chkAutoLoad.Value = 0 Then
          txtFileSettings = App.path & "\Default.HFW"
       Else
          txtFileSettings = GetSetting(App.ProductName, "Settings", "AutoLoadFilename", App.path & "\Default.HFW")
       End If
       LoadSettings txtFileSettings
    End If

    Drive1 = "C:\"
    dirStack = vbNullString
    
    DoEvents

    If Len(Command) > 0 Then HandleCommandLineActions Command

    gbQuiet = False

    vstMain.CurrTab = StartTab
    
    txtFilePrefix.Enabled = True
    LoadFormPosition Me

End Sub

Private Sub HandleCommandLineVars()
       Dim sCommandLine As String
       Dim sCommandLineParam As String
       Dim sT As String
       Dim bSaveToDirectory As Boolean

       bSaveToDirectory = False
       sCommandLine = Command()

       If bCheckCommandAndRemove(sCommandLine, "-Load", sCommandLineParam, True) Then
          txtFileSettings = sCommandLineParam
       Else
          sT = sCommandLine
          If bCheckCommandAndRemove(sT, "-Dir", sCommandLineParam, True) Then
             bSaveToDirectory = True
             If FileExists(sCommandLineParam & "\WalkIt.HFW") Then
                txtFileSettings = sCommandLineParam & "\WalkIt.HFW"
             Else
                txtFileSettings = App.path & "\WalkIt.HFW"
             End If
          Else
             txtFileSettings = App.path & "\WalkIt.HFW"
          End If
       End If
       LoadSettings txtFileSettings

       If bSaveToDirectory Then
          txtFileSettings = sCommandLineParam & "\WalkIt.HFW"
          txtOutFilename = sCommandLineParam & ".html"
          StartTab = 0
       End If

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
         'MsgBox "Command Line Processing Directory" & gsEolTab & sCommandLineParam
          txtStartPath = sCommandLineParam
       End If

       If bCheckCommandAndRemove(sCommandLine, "-Exclude", sCommandLineParam, True) Then
          txtExclude = sCommandLineParam
       End If

       If bCheckCommandAndRemove(sCommandLine, "-Output", sCommandLineParam, True) Then
         'MsgBox "Command Processing Output To" & gsEolTab & sCommandLineParam
          txtOutFilename = sCommandLineParam
       End If

       If bCheckCommandAndRemove(sCommandLine, "-Shell", sCommandLineParam, True) Then
          txtBrowserPath = sCommandLineParam
       End If
       
End Sub

Public Sub LoadSettings(ByVal sFilename As String)
On Error GoTo EH_LoadSettings
    Static sT As String

    If Len(sFilename) = 0 Then sFilename = App.path & "\Default.HFW"
    Open sFilename For Input Access Read As #1
         Do Until EOF(1)
            Line Input #1, sT
            Select Case UCase$(Trim$(Left$(sT, InStr(sT, "=") - 1)))
                   Case "INPUTPATH":        txtStartPath = Trim$(sAfter(sT, 1, "="))
                   Case "OUTPUTFILE":       txtOutFilename = Trim$(sAfter(sT, 1, "="))
                   Case "TITLE":            txtTitle = Trim$(sAfter(sT, 1, "="))
                   Case "EXCLUDE":          txtExclude = Trim$(sAfter(sT, 1, "="))
                   Case "BACKGROUNDPIC":    txtBackgroundPic = Trim$(sAfter(sT, 1, "="))
                   Case "BACKGROUNDSOUND":  txtBackgroundSound = Trim$(sAfter(sT, 1, "="))
                  'Case "BROWSERPATH":      txtBrowserPath = Trim$(sAfter(sT, 1, "="))
                   Case "FTPHOST":          txtHost = Trim$(sAfter(sT, 1, "="))
                   Case "USERNAME":         txtUsername = Trim$(sAfter(sT, 1, "="))
                   Case "PASSWORD":         txtPassword = Trim$(sAfter(sT, 1, "="))
                   Case "REMOTEPATH":       txtRemotePath = Trim$(sAfter(sT, 1, "="))
                   Case "HOSTPORT":         txtHostPort = Trim$(sAfter(sT, 1, "="))
                   Case "USELOCAL":         chkLocal.Value = Trim$(sAfter(sT, 1, "="))
                   Case "FILESPERPAGE":     txtMaxFilesPerPage = Trim$(sAfter(sT, 1, "="))
                                            chkMultipage.Value = Abs(Val(txtMaxFilesPerPage) > 0)
                   Case "CELLSACROSS":      txtCellsWide = Trim$(sAfter(sT, 1, "="))
                                            If Val(txtCellsWide) = 0 Then txtCellsWide = "1"
                   Case "USETABLES":        chkTable.Value = IIf(Trim$(sAfter(sT, 1, "=")) = "True", 1, 0)
                   Case "SPREADEVENLY":     frmAdvTable.chkSpreadEvenly.Value = Val(Trim$(sAfter(sT, 1, "=")))
                   Case "TABLEBACKCOLOR"
                         With frmAdvTable.vseBackground
                              If Trim$(sAfter(sT, 1, "=")) = "Clear" Then
                                 .Caption = "Clear"
                                 .BackColor = &H8000000F
                              Else
                                 .Caption = vbNullString
                                 .BackColor = "&H" & Trim$(sAfter(sT, 1, "="))
                              End If
                         End With
                   Case "TABLEWIDTH":       frmAdvTable.txtWidth = Trim$(sAfter(sT, 1, "="))
                   Case "BORDERWIDTH":      frmAdvTable.txtBorderWidth = Trim$(sAfter(sT, 1, "="))
                   Case "CELLPADDING":      frmAdvTable.txtCellPadding = Trim$(sAfter(sT, 1, "="))
                   Case "CELLSPACING":      frmAdvTable.txtCellSpacing = Trim$(sAfter(sT, 1, "="))
                   Case "SPECIAL 0":        txtSpecial(0) = sDenormalize(Trim$(sAfter(sT, 1, "=")))
                   Case "SPECIAL 1":        txtSpecial(1) = sDenormalize(Trim$(sAfter(sT, 1, "=")))
                   Case "SPECIAL 2":        txtSpecial(2) = sDenormalize(Trim$(sAfter(sT, 1, "=")))
                   Case "SPECIAL 3":        txtSpecial(3) = sDenormalize(Trim$(sAfter(sT, 1, "=")))
                   Case "SPECIAL 4":        txtSpecial(4) = sDenormalize(Trim$(sAfter(sT, 1, "=")))
                   Case "SPECIAL 5":        txtSpecial(5) = sDenormalize(Trim$(sAfter(sT, 1, "=")))
                   Case "LIMITSIZE":        chkLimitSize.Value = IIf(Trim$(sAfter(sT, 1, "=")) = "True", 1, 0)
                   Case "MAXX":             txtMaxX = Trim$(sAfter(sT, 1, "="))
                   Case "MAXY":             txtMaxY = Trim$(sAfter(sT, 1, "="))
                   Case "AUTOVIEW":         chkAutoView.Value = IIf(Trim$(sAfter(sT, 1, "=")) = "True", 1, 0)
                   Case "FILEPREFIX":       txtFilePrefix = Trim$(sAfter(sT, 1, "="))
                   Case "BREAKINTOPAGES":   chkBreakIntoPages.Value = IIf(Trim$(sAfter(sT, 1, "=")) = "True", 1, 0)
                   Case "BUILDINDEX":       chkBuildIndex.Value = IIf(Trim$(sAfter(sT, 1, "=")) = "True", 1, 0)
                   Case "NOROOTENTRY":      chkNoRootEntry.Value = IIf(Trim$(sAfter(sT, 1, "=")) = "True", 1, 0)
                   Case "NOFILEEXT":        chkHideExt.Value = IIf(Trim$(sAfter(sT, 1, "=")) = "True", 1, 0)
                   Case "AUTOSAVE":         chkAutoSave.Value = IIf(Trim$(sAfter(sT, 1, "=")) = "True", 1, 0)
                   Case "SKIPSUBDIRS":      chkSkipSubdirs.Value = IIf(Trim$(sAfter(sT, 1, "=")) = "True", 1, 0)
                   Case "PROCESSMP3":       chkProcessMP3Tags.Value = IIf(Trim$(sAfter(sT, 1, "=")) = "True", 1, 0)
                   Case "RENAMEMP3":        chkRenameMP3.Value = IIf(Trim$(sAfter(sT, 1, "=")) = "True", 1, 0)
            End Select
         Loop
    Close #1

EH_LoadSettings_Exit:
    Exit Sub

EH_LoadSettings:
    GoTo EH_LoadSettings_Exit
End Sub


Public Sub SaveSettings(ByVal sFilename As String)
    Dim ff As Long
On Error Resume Next

    If Len(sFilename) = 0 Then Exit Sub

    ff = FreeFile
    Open sFilename For Output Access Write As #ff
            Print #ff, "InputPath=" & txtStartPath
            Print #ff, "OutputFile=" & txtOutFilename
            Print #ff, "Title=" & txtTitle
            Print #ff, "Exclude=" & txtExclude
            Print #ff, "BackgroundPic=" & txtBackgroundPic
            Print #ff, "BackgroundSound=" & txtBackgroundSound
            Print #ff, "FTPHost=" & txtHost
            Print #ff, "Username=" & txtUsername
            Print #ff, "Password=" & txtPassword
            Print #ff, "RemotePath=" & txtRemotePath
            Print #ff, "HostPort=" & txtHostPort
            Print #ff, "UseLocal=" & chkLocal.Value
            Print #ff, "FilesPerPage=" & IIf(chkMultipage.Value <> 0, txtMaxFilesPerPage, 0)
            Print #ff, "CellsAcross=" & txtCellsWide
            Print #ff, "UseTables=" & IIf(chkTable.Value <> 0, "True", "False")
            Print #ff, "SpreadEvenly=" & frmAdvTable.chkSpreadEvenly.Value
            Print #ff, "TableBackColor=" & IIf(frmAdvTable.vseBackground.Caption = "Clear", "Clear", Hex(frmAdvTable.vseBackground.BackColor))
            Print #ff, "TableWidth=" & frmAdvTable.txtWidth
            Print #ff, "BorderWidth=" & frmAdvTable.txtBorderWidth
            Print #ff, "CellPadding=" & frmAdvTable.txtCellPadding
            Print #ff, "CellSpacing=" & frmAdvTable.txtCellSpacing
            Print #ff, "Special 0=" & sNormalize(txtSpecial(0))
            Print #ff, "Special 1=" & sNormalize(txtSpecial(1))
            Print #ff, "Special 2=" & sNormalize(txtSpecial(2))
            Print #ff, "Special 3=" & sNormalize(txtSpecial(3))
            Print #ff, "Special 4=" & sNormalize(txtSpecial(4))
            Print #ff, "Special 5=" & sNormalize(txtSpecial(5))
            Print #ff, "LimitSize=" & IIf(chkLimitSize.Value <> 0, "True", "False")
            Print #ff, "MaxX=" & txtMaxX
            Print #ff, "MaxY=" & txtMaxY
            Print #ff, "AutoView=" & IIf(chkAutoView.Value <> 0, "True", "False")
            Print #ff, "FilePrefix=" & txtFilePrefix
            Print #ff, "BreakIntoPages=" & IIf(chkBreakIntoPages.Value <> 0, "True", "False")
            Print #ff, "BuildIndex=" & IIf(chkBuildIndex.Value <> 0, "True", "False")
            Print #ff, "NoRootEntry=" & IIf(chkNoRootEntry.Value <> 0, "True", "False")
            Print #ff, "NoFileExt=" & IIf(chkHideExt.Value <> 0, "True", "False")
            Print #ff, "AutoSave=" & IIf(chkAutoSave.Value <> 0, "True", "False")
            Print #ff, "SkipSubDirs=" & IIf(chkSkipSubdirs.Value <> 0, "True", "False")
            Print #ff, "ProcessMP3=" & IIf(chkProcessMP3Tags.Value <> 0, "True", "False")
            Print #ff, "RenameMP3=" & IIf(chkRenameMP3.Value <> 0, "True", "False")
    Close #1
End Sub


Private Sub Form_Resize()
On Error Resume Next
    VideoSoftElastic8.Move ScaleWidth - 1485, ScaleHeight - 945
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPosition Me
End Sub


Private Sub ftpPutPage_AsyncError(ByVal ErrorNum As Long, ByVal ErrorMsg As String)
    MsgBox "ftpPutPage: AsyncError ( # " & ErrorNum & ") = " & ErrorMsg
End Sub

Private Sub ftpPutPage_Done()
    'MsgBox "ftpPutPage: Done"
End Sub


Private Sub Label2_Click()
    Label3_Click
End Sub

Private Sub Label3_Click()
    tmrMoveFS.Enabled = False
    HTML_ViewWebPage ("http://www.firmsolutions.com"), (txtBrowserPath)
End Sub

Private Sub tmrMoveFS_Timer()
    Static SinDCount As Double
    Static SinDCount2 As Double
    Static SinDCOunt3 As Double
    Static Color As Long
    Static ColorSet As Long
    Static YOffset As Long
    Dim iDepth As Long

tmrMoveFS_Timer_Again:
    If iDepth > 20 Then
       Exit Sub
    End If

    If dCount > 20 Or dCount < 0 Then
       dDirection = -dDirection
      'YOffset = YOffset + 0.8
      'If dDirection = 1 Then
      '   dDirection = -1
      'Else
      '   dDirection = 1
      'End If
    End If

    dCountP2 = dCountPrevious
    dCountPrevious = dCount
    dCount = dCount + (0.01 * dDirection)

    'picMoveFS.Move dCount, Sin(dCount * dDirection)

    SinDCount = Sin((dCountP2 + YOffset) * dDirection)
    SinDCount2 = SinDCount + 0.8 'dCount 'picMoveFS.Height
    SinDCOunt3 = dCountP2 + 0.8 'dCount 'picMoveFS.Height
    
    Color = (Color + 1) Mod 128
    ColorSet = Color + (Color Or &HFF) + &HFF     '(Color Or &H1111)

    picBack.PSet (dCountP2, SinDCount), ColorSet
    picBack.PSet (dCountP2, SinDCount2), ColorSet + 16
    picBack.PSet (SinDCOunt3, SinDCount), ColorSet + 64
    picBack.PSet (SinDCOunt3, SinDCount2), ColorSet + 128

    iDepth = iDepth + 1
    GoTo tmrMoveFS_Timer_Again
End Sub


Private Sub txtCellsWide_LostFocus()
    txtCellsWide = Val(txtCellsWide)
    If Val(txtCellsWide) = 0 Then txtCellsWide = "1"
End Sub
Private Sub txtCommandLine_KeyPress(KeyAscii As Integer)
    Beep
    KeyAscii = 0
End Sub


Private Sub txtMaxFilesPerPage_LostFocus()
    txtMaxFilesPerPage = Val(txtMaxFilesPerPage)
End Sub


Private Sub txtOutFilename_Change()
    'txtFTPOutputTo = sGetFilename(txtOutFilename)
End Sub

Private Sub vstMain_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    tmrMoveFS.Enabled = False
End Sub


