VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   15135
   ClientLeft      =   5385
   ClientTop       =   1530
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   15135
   ScaleWidth      =   11145
   Begin VB.CommandButton cmdMachineName 
      Caption         =   "Machine Name"
      Height          =   285
      Left            =   7860
      TabIndex        =   9
      Top             =   60
      Width           =   1395
   End
   Begin VB.CommandButton cmdAllAccounts 
      Caption         =   "&All Accounts"
      Height          =   285
      Left            =   6635
      TabIndex        =   8
      Top             =   60
      Width           =   1125
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   285
      Left            =   5410
      TabIndex        =   7
      Top             =   60
      Width           =   1125
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10515
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoWebMain.frx":0000
            Key             =   "Success"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoWebMain.frx":0452
            Key             =   "Failure"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoWebMain.frx":0BA4
            Key             =   "Timeout"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwResults 
      Height          =   2205
      Left            =   75
      TabIndex        =   5
      Top             =   13065
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   3889
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   8454143
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Time"
         Text            =   "Time"
         Object.Width           =   3757
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Success"
         Text            =   "Success"
         Object.Width           =   1640
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "Process"
         Text            =   "Process"
         Object.Width           =   1640
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "URL"
         Text            =   "URL"
         Object.Width           =   23098
      EndProperty
   End
   Begin VB.TextBox txtPrefix 
      Height          =   300
      Left            =   60
      TabIndex        =   4
      Text            =   "http://www.vbcode.com"
      Top             =   30
      Width           =   3180
   End
   Begin VB.ListBox lstURLs 
      Height          =   1815
      ItemData        =   "frmAutoWebMain.frx":0FF6
      Left            =   60
      List            =   "frmAutoWebMain.frx":0FF8
      TabIndex        =   3
      Top             =   390
      Width           =   11040
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   285
      Left            =   4770
      TabIndex        =   2
      Top             =   60
      Width           =   540
   End
   Begin VB.TextBox txtInstances 
      Height          =   285
      Left            =   4130
      TabIndex        =   0
      Text            =   "10"
      Top             =   60
      Width           =   540
   End
   Begin InetCtlsObjects.Inet xinetCofee 
      Index           =   0
      Left            =   9900
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin MSComctlLib.ListView lvwStep 
      Height          =   10800
      Left            =   60
      TabIndex        =   6
      Top             =   2250
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   19050
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   8454143
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Process"
         Text            =   "Process"
         Object.Width           =   1640
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Step"
         Text            =   "Step"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "URL"
         Text            =   "URL"
         Object.Width           =   22931
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Instances"
      Height          =   195
      Left            =   3340
      TabIndex        =   1
      Top             =   75
      Width           =   690
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mlInstances As Long
Public InstancesRunning As Long
Public IEBrowsers As CAutoWebBrowsers

Public MachineName As String

Private bLoadingForm As Boolean
Public Sub LoadInstances()
On Error GoTo EH_frmMain_LoadInstances
    Dim CurrInstance As Long
    
    If Len(MachineName) = 0 Then
       MachineName = InputBox("Please entery your Machine's name", "MACHINE NAME")
       If Len(MachineName) = 0 Then Exit Sub
    End If
    
    For CurrInstance = 0 To mlInstances - 1
        With IEBrowsers.Add("Instance " & CurrInstance)
             InstancesRunning = InstancesRunning + 1
             .StartingAddress = txtPrefix & sGetToken(lstURLs.List(0), 1, "~~~")
             .MachineName = MachineName
             .Tag = 1
             With lvwStep.ListItems.Add(, "Instance " & CurrInstance + 1, "" & CurrInstance + 1)
                  .SubItems(1) = 1
                  .SubItems(2) = txtPrefix & lstURLs.List(0)
             End With
             .Visible = True
             .ControledBrowsing = True
             .brwWebBrowser.Navigate .StartingAddress
        End With
    Next CurrInstance

EH_frmMain_LoadInstances_Continue:
    Exit Sub

EH_frmMain_LoadInstances:
    LogError "frmMain", "LoadInstances", Err.Number, Err.Description
    Resume EH_frmMain_LoadInstances_Continue

    Resume
End Sub

Public Sub UnloadInstances()
    IEBrowsers.Clear
End Sub

Private Sub cmdAllAccounts_Click()
On Error GoTo EH_frmMain_cmdAllAccounts_Click
    frmBrowser.Show
    frmBrowser.StartingAddress = "http://cofeeweb/cofee_web_ins/GTECofeeAllAccounts.asp"
    frmBrowser.brwWebBrowser.Navigate frmBrowser.StartingAddress

EH_frmMain_cmdAllAccounts_Click_Continue:
    Exit Sub

EH_frmMain_cmdAllAccounts_Click:
    LogError "frmMain", "cmdAllAccounts_Click", Err.Number, Err.Description
    Resume EH_frmMain_cmdAllAccounts_Click_Continue

    Resume
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo EH_frmMain_cmdBrowse_Click
    frmBrowser.StartingAddress = txtPrefix
    frmBrowser.brwWebBrowser.Navigate frmBrowser.StartingAddress
    frmBrowser.Show

EH_frmMain_cmdBrowse_Click_Continue:
    Exit Sub

EH_frmMain_cmdBrowse_Click:
    LogError "frmMain", "cmdBrowse_Click", Err.Number, Err.Description
    Resume EH_frmMain_cmdBrowse_Click_Continue

    Resume
End Sub


Private Sub cmdGo_Click()
On Error GoTo EH_frmMain_cmdGo_Click
    Dim CurrInstance As Long
    Dim MaxCount As Long
    mlInstances = Val(txtInstances)
    If mlInstances < 1 Then mlInstances = 1
    If mlInstances > 150 Then mlInstances = 150
    txtInstances = mlInstances

    UnloadInstances
    
    lvwResults.ListItems.Clear
    lvwStep.ListItems.Clear
    Screen.MousePointer = vbHourglass
    InstancesRunning = 0

    LoadInstances

EH_frmMain_cmdGo_Click_Continue:
    Exit Sub

EH_frmMain_cmdGo_Click:
    LogError "frmMain", "cmdGo_Click", Err.Number, Err.Description
    Resume EH_frmMain_cmdGo_Click_Continue

    Resume
End Sub

Private Sub cmdMachineName_Click()
    MachineName = InputBox("Please entery your Machine's name", "MACHINE NAME")
    SaveSetting App.ProductName, "Last", "Machine Name", MachineName
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If bLoadingForm Then Exit Sub
    lstURLs.Width = ScaleWidth
    lvwResults.Move 50, lvwResults.Top, ScaleWidth - 100, ScaleHeight - lvwResults.Top - 50
    lvwStep.Move 50, lvwStep.Top, ScaleWidth - 100, ScaleHeight - lvwStep.Top - lvwResults.Height - 50
End Sub

Private Sub Form_Load()
On Error GoTo EH_frmMain_Form_Load
    bLoadingForm = True
       LoadFormPosition Me
    bLoadingForm = False
    lvwResults.Move 50, ScaleHeight - lvwResults.Height - 50
    txtInstances.Text = GetSetting(App.ProductName, "Last Value", "txtInstances", "")
    MachineName = GetSetting(App.ProductName, "Last Value", "Machine Name", "")

    Set IEBrowsers = New CAutoWebBrowsers
    Set IEBrowsers.Parent = Me

  ' Read the list contents of the 'lstURLs' control from the registry
    Dim CurrItem As Long
    Dim sEntry As String
    Dim sContents As String

    lstURLs.Clear
    sContents = GetSetting(App.ProductName, "Last Value", "lstURLs", "")
    Do While InStr(1, sContents, vbCrLf)
       sEntry = Left(sContents, InStr(1, sContents, vbCrLf) - 1)
       lstURLs.AddItem sEntry
       sContents = Mid(sContents, InStr(1, sContents, vbCrLf) + 2)
    Loop
    ExceptionStartup

EH_frmMain_Form_Load_Continue:
    Exit Sub

EH_frmMain_Form_Load:
    LogError "frmMain", "Form_Load", Err.Number, Err.Description
    Resume EH_frmMain_Form_Load_Continue

    Resume
End Sub

Private Sub Form_Terminate()
On Error Resume Next
    IEBrowsers.Clear False
    Set IEBrowsers.Parent = Nothing
    Set IEBrowsers = Nothing
    Unload frmBrowser
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo EH_frmMain_Form_Unload
    SaveFormPosition Me
    SaveSetting App.ProductName, "Last Value", "txtInstances", txtInstances.Text

  ' Save the list contents of the 'lstURLs' control to the registry
    Dim sList As String
    Dim CurrItem As Long
    For CurrItem = 0 To lstURLs.ListCount - 1
        sList = sList & lstURLs.List(CurrItem) & vbCrLf
    Next CurrItem
    SaveSetting App.ProductName, "Last Value", "lstURLs", sList

    UnloadInstances
    Form_Terminate
    ExceptionShutdown

EH_frmMain_Form_Unload_Continue:
    Exit Sub

EH_frmMain_Form_Unload:
    LogError "frmMain", "Form_Unload", Err.Number, Err.Description
    Resume EH_frmMain_Form_Unload_Continue

    Resume
End Sub

Private Sub lstURLs_DblClick()
    lstURLs.RemoveItem lstURLs.ListIndex
End Sub


Private Sub lstURLs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo EH_frmMain_lstURLs_MouseUp
    If Button = vbRightButton And Shift = 0 Then
       With frmURLEntry
            .URLEntry = lstURLs.List(lstURLs.ListIndex)
            .Show vbModal, Me
            If Not .Canceled Then
               lstURLs.List(lstURLs.ListIndex) = .URLEntry
            End If
       End With
    End If

EH_frmMain_lstURLs_MouseUp_Continue:
    Exit Sub

EH_frmMain_lstURLs_MouseUp:
    LogError "frmMain", "lstURLs_MouseUp", Err.Number, Err.Description
    Resume EH_frmMain_lstURLs_MouseUp_Continue

    Resume
End Sub


