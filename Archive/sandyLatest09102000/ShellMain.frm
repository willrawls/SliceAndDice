VERSION 5.00
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "SysTray.ocx"
Begin VB.Form frmMain 
   Caption         =   "Slice and Dice Shell"
   ClientHeight    =   525
   ClientLeft      =   270
   ClientTop       =   1890
   ClientWidth     =   2955
   ClipControls    =   0   'False
   HasDC           =   0   'False
   Icon            =   "ShellMain.frx":0000
   LinkTopic       =   "SandyShell"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin SysTrayCtl.cSysTray trayMain 
      Left            =   0
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "ShellMain.frx":014A
      TrayTip         =   "Slice and Dice Shell"
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuTraySep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuTraySep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowExternals 
         Caption         =   "&Externals..."
      End
      Begin VB.Menu mnuFavorites 
         Caption         =   "&Favorites..."
      End
      Begin VB.Menu mnuTraySep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainWindow 
         Caption         =   "&Main Window"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myIDTExtender       As VBIDE.IDTExtensibility
Private QueuedPopupMenu     As VB.Menu

Private CustomVars(0 To 1)  As Variant
Private sQueuedPopupMenu    As String

Private IsMenuDisplayed     As Boolean

Public Sub ShowQueuedPopupMenu()
On Error Resume Next
    Dim MenuToShow      As VB.Menu
    Dim mySandyWizard   As SliceAndDice.Wizard

TryAgain:
    If Not QueuedPopupMenu Is Nothing Then
       Set MenuToShow = QueuedPopupMenu
       Set QueuedPopupMenu = Nothing
       IsMenuDisplayed = True
           PopupMenu MenuToShow, , , , mnuMainWindow
       IsMenuDisplayed = False
       GoTo TryAgain
    ElseIf Len(sQueuedPopupMenu) > 0 Then
       Set mySandyWizard = myIDTExtender
           Select Case UCase$(sQueuedPopupMenu)
                  Case "FAVORITES": mySandyWizard.FavoriteCalledFromIDE = True: mySandyWizard.ShowFavoritesMenu
                  Case "EXTERNALS": mySandyWizard.ShowExternalsMenu
                  Case "ABOUT":     mySandyWizard.ShowSplashScreen
           End Select
       Set mySandyWizard = Nothing
       sQueuedPopupMenu = vbNullString
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Set myIDTExtender = New SliceAndDice.Wizard

    CustomVars(0) = "sadAddin|sadFile|sadRegister|sadSoftCodeWmr"
    CustomVars(1) = ""
    CustomVars(2) = ""

    myIDTExtender.OnConnection Nothing, vbext_cm_External, Nothing, CustomVars
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    myIDTExtender.OnDisconnection vbext_dm_HostShutdown, CustomVars
    Set myIDTExtender = Nothing
End Sub

Private Sub mnuAbout_Click()
    sQueuedPopupMenu = "About"
End Sub

Private Sub mnuExit_Click()
On Error Resume Next
    Dim mySandyWizard As SliceAndDice.Wizard
    Set mySandyWizard = myIDTExtender
        mySandyWizard.HideWindows
    Set mySandyWizard = Nothing

    Unload Me
End Sub

Private Sub mnuFavorites_Click()
    sQueuedPopupMenu = "Favorites"
End Sub

Private Sub mnuMainWindow_Click()
On Error Resume Next
    Dim mySandyWizard As SliceAndDice.Wizard
    Set mySandyWizard = myIDTExtender
        mySandyWizard.ShowMainWindow
    Set mySandyWizard = Nothing
End Sub

Private Sub mnuShowExternals_Click()
    sQueuedPopupMenu = "Externals"
End Sub

Private Sub trayMain_MouseDblClick(Button As Integer, Id As Long)
On Error Resume Next
    mnuMainWindow_Click
End Sub

Private Sub trayMain_MouseUp(Button As Integer, Id As Long)
On Error Resume Next
    Set QueuedPopupMenu = mnuTray
    ShowQueuedPopupMenu
End Sub

