VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowser 
   ClientHeight    =   7350
   ClientLeft      =   4410
   ClientTop       =   3195
   ClientWidth     =   10425
   Icon            =   "frmAutoWebBrowser.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin InetCtlsObjects.Inet inetGet 
      Left            =   90
      Top             =   5070
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3495
      Left            =   30
      TabIndex        =   0
      Top             =   1455
      Width           =   5400
      ExtentX         =   9525
      ExtentY         =   6165
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3930
      Top             =   990
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   10425
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   10425
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   30
         TabIndex        =   2
         Top             =   315
         Width           =   3795
      End
      Begin VB.Label lblAddress 
         Caption         =   "&Address:"
         Height          =   255
         Left            =   45
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   60
         Width           =   3075
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2670
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoWebBrowser.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoWebBrowser.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoWebBrowser.frx":0A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoWebBrowser.frx":0CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoWebBrowser.frx":0FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutoWebBrowser.frx":12AC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DontNavigateNow As Boolean
Public ControledBrowsing As Boolean
Public SequenceFinished As Boolean

Public StartingAddress As String
Public AltAccepted As String
Public DataToPost As String
Public ButtonToPress As String
Public QueryString As String

Public Index As Long
Public MachineName As String

Public strData As String
Public sAtTop As String
Public sAtBottom As String
Public sTemplateName As String
Public sCategoryName As String

Private AlreadyAsked As Boolean
Public NavigationComplete As Boolean

Private CurrentStage As Long
Public Parent As NewCommands

Public Sub SubmitTemplate()
1    On Error Resume Next
2        Dim sCategory As String
3        Dim sTemplate As String

4        With Parent.Parent
5             If Len(.CurrentTemplateNameAndCategory) > 0 Then
6                .GetCategoryAndName .CurrentTemplateNameAndCategory, sCategory, sTemplate
7                If Not .SliceAndDice(sCategory).Templates(sTemplate) Is Nothing Then
8                   With .SliceAndDice(sCategory).Templates(sTemplate)
                    
9                   End With
10               End If
11            End If
12       End With
End Sub


'Public Function GetCentralUpdateInfo(Optional ByVal bFetchNewFiles As Boolean = False) As Boolean
'13       Dim sResponse As String
'14       Dim asaX As CAssocArray
'15       Dim CurrItem As CAssocItem
'16   On Error Resume Next
'17       Screen.MousePointer = vbHourglass
'18           sResponse = GetURL("http://www.sliceanddice.com/central.update")
'19       Screen.MousePointer = vbDefault
'20       If Len(sResponse) = 0 Then
'21          If bUserSure("The Central Server Update Information cannot be acceessed right now." & vbCr & vbTab & "Continue with current settings ?") Then
'22             GetCentralUpdateInfo = True
'23          End If
'24       Else
'25          sResponse = Replace$(sResponse, vbCrLf, "")
'26          If InStr(sResponse, "$$$$") = 0 Then
'27             If bUserSure("The Central Server Update Information cannot be acceessed right now." & vbCr & vbTab & "Continue with current settings ?") Then
'28                GetCentralUpdateInfo = True
'29             End If
'30          End If
'31          Set asaX = New CAssocArray
'32              asaX.ItemDelimiter = "$$$$"
'33              asaX.All = sResponse
'34              For Each CurrItem In asaX
'35                  sadSaveLicenseKey CurrItem.Key, CurrItem.Value
'36                  If bFetchNewFiles Then
'
'37                  End If
'38              Next CurrItem
'39          Set asaX = Nothing
'40          GetCentralUpdateInfo = True
'41       End If
'End Function

Public Function GetFile(ByVal sURL As String, ByVal sFilename As String) As Boolean
42   On Error Resume Next
43       Dim fh As Long
44       Dim b() As Byte
45       Dim strURL As String

46       If Len(sURL) = 0 Or Len(sFilename) = 0 Then Exit Function

47       If InStr(sFilename, "\") = 0 Then sFilename = App.Path & "\" & sFilename

48       If Len(Dir(sFilename)) > 0 Then
49          Kill sFilename
50       End If

   'Load inetGet
51           inetGet.RequestTimeout = 60
52           b() = inetGet.OpenURL(sURL, icByteArray)
   'Unload inetGet

53       fh = FreeFile
54       Open sFilename For Binary Access Write As #fh
55            Put #fh, , b()
56       Close #fh
End Function

Public Function GetURL(ByVal sURL As String) As String
57   On Error Resume Next
   'Load inetGet
58           inetGet.RequestTimeout = 60
59           GetURL = inetGet.OpenURL(sURL)
   'Unload inetGet
End Function

Public Sub PostURL(ByVal sURL As String, ByVal sData)
60   On Error Resume Next
   'Load inetGet
61           inetGet.RequestTimeout = 60
62           inetGet.Execute sURL, "POST", sData
   'Unload inetGet
End Sub

Private Sub brwWebBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim bOkaySoFar As Boolean

    DoEvents
On Error Resume Next
    strData = vbNullString
    strData = brwWebBrowser.Document.body.innerhtml
On Error GoTo EH_frmBrowser_brwWebBrowser_DocumentComplete
    
    sAtBottom = sAfter(sGetToken(sGetToken(strData, 3, "<textarea"), 1, "</textarea"), 1, ">")
    sAtTop = sAfter(sGetToken(sGetToken(strData, 2, "<textarea"), 1, "</textarea"), 1, ">")
    
    If Len(sAtBottom) > 0 Or Len(sAtTop) > 0 Then
       If InStr(1, strData, "<!--") > 0 Then
          sTemplateName = Left$(Trim(sGetToken(sGetToken(strData, 2, "Task</STRONG>:"), 1, "</FONT")), 254)
          If Len(sTemplateName) > 0 And Len(timTimer.Tag) = 0 And (Not AlreadyAsked) Then
             AlreadyAsked = True
             If bUserSure("This appears to be a Code Snippet from www.vbcode.com" & vbCrLf & vbTab & "Would you like to import this as a Slice and Dice ?" & vbCrLf & vbCrLf & vbTab & vbTab & "Cagetory Name = " & sCategoryName & vbCrLf & vbTab & vbTab & "Template Name = " & sTemplateName) Then
                timTimer.Tag = "Import Template": timTimer.Enabled = True
             End If
          End If
       End If
    End If

EH_frmBrowser_brwWebBrowser_DocumentComplete_Continue:
    NavigationComplete = True
    Exit Sub

EH_frmBrowser_brwWebBrowser_DocumentComplete:
    LogError "frmBrowser", "brwWebBrowser_DocumentComplete", Err.Number, Err.Description
    Resume EH_frmBrowser_brwWebBrowser_DocumentComplete_Continue

    Resume
End Sub

Private Sub brwWebBrowser_DownloadBegin()
    AlreadyAsked = False
    NavigationComplete = False
    Me.ZOrder
End Sub

Private Sub brwWebBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
On Error GoTo EH_frmBrowser_brwWebBrowser_NavigateComplete2
    Dim i As Integer
    Dim bFound As Boolean
    Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    DontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    DontNavigateNow = False

EH_frmBrowser_brwWebBrowser_NavigateComplete2_Continue:
    Me.ZOrder
    Exit Sub

EH_frmBrowser_brwWebBrowser_NavigateComplete2:
    LogError "frmBrowser", "brwWebBrowser_NavigateComplete2", Err.Number, Err.Description
    Resume EH_frmBrowser_brwWebBrowser_NavigateComplete2_Continue

    Resume
End Sub

Private Sub Form_Load()
On Error Resume Next
    tbToolBar.Refresh
    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15
    LoadFormPosition Me
    sCategoryName = "From VBCodeDotCom"
End Sub

Private Sub brwWebBrowser_DownloadComplete()
On Error Resume Next
    Caption = brwWebBrowser.LocationName
    Me.ZOrder
End Sub

Private Sub cboAddress_Click()
    If DontNavigateNow Then Exit Sub
    timTimer.Enabled = True: timTimer.Tag = vbNullString

    brwWebBrowser.Navigate cboAddress.Text
    Me.ZOrder
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
    Me.ZOrder
End Sub

Private Sub Form_Resize()
On Error Resume Next
    cboAddress.Width = ScaleWidth - 100
    brwWebBrowser.Width = ScaleWidth - 100
    brwWebBrowser.Height = ScaleHeight - (picAddress.Top + picAddress.Height) - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPosition Me
End Sub

Private Sub timTimer_Timer()
On Error GoTo EH_frmBrowser_timTimer_Timer
    Dim sVarName As String
    Dim lToken As Long

    Me.ZOrder
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        If Len(timTimer.Tag) = 0 Then
           Caption = brwWebBrowser.LocationName
        ElseIf timTimer.Tag = "Import Template" Then
           timTimer.Tag = vbNullString
           If Not Parent.Parent.SliceAndDice.Categorys(sCategoryName) Is Nothing Then
              If Not Parent.Parent.SliceAndDice.Categorys(sCategoryName).Templates(sTemplateName) Is Nothing Then
                 If Not bUserSure("That template already exists." & vbCrLf & vbCrLf & vbTab & "Overwrite it ?") Then
                    Exit Sub
                 End If
              Else
                 Parent.Parent.NewTemplate True, sCategoryName & " - " & sTemplateName, sTemplateName, False
              End If
           Else
              Parent.Parent.NewTemplate True, sCategoryName & " - " & sTemplateName, sTemplateName, False
           End If
           
           If Not Parent.Parent.SliceAndDice.Categorys(sCategoryName) Is Nothing Then
              If Not Parent.Parent.SliceAndDice.Categorys(sCategoryName).Templates(sTemplateName) Is Nothing Then
                 With Parent.Parent.SliceAndDice.Categorys(sCategoryName).Templates(sTemplateName)
                      .memoCodeAtTop = "~~GotoModule modGeneral" & vbCrLf & sAtTop
                      .memoCodeAtBottom = sAtBottom
                 End With
                 Parent.Parent.SliceAndDice.Save
                 Parent.Parent.JumpTo sCategoryName & " - " & sTemplateName
              End If
           End If
        Else
           timTimer.Tag = vbNullString
           
           Do While InStr(1, StartingAddress, "%%")
              sVarName = sGetToken(StartingAddress, 2, "%%")
              lToken = lFindToken(QueryString, sVarName, "&")
              If lToken > 0 Then
                 sVarName = sAfter(sGetToken(QueryString, lToken, "&"), 1, "=")
              Else
                 sVarName = sVarName & "="
              End If
              StartingAddress = sBefore(StartingAddress, 2, "%%") & sVarName & sAfter(StartingAddress, 2, "%%")
           Loop

           If Len(StartingAddress) Then
              brwWebBrowser.Navigate StartingAddress
           Else
On Error Resume Next
              Do While Len(DataToPost)
                 sVarName = sGetToken(DataToPost, 1, "$$$")
                 DataToPost = sAfter(DataToPost, 1, "$$$")
                 If Len(sVarName) = 0 Then
                 ElseIf InStr(1, sGetToken(sVarName, 1, "="), ".") Then
                    If StrComp(sGetToken(sVarName, 1, "="), "txtFirstName") <> 0 Then
                       brwWebBrowser.Document.Forms(CInt(Left$(sVarName, 1))).Item(sGetToken(Mid$(sVarName, 3), 1, "=")).Value = sAfter(sVarName, 1, "=")
                    Else
                       brwWebBrowser.Document.Forms(CInt(Left$(sVarName, 1))).Item(sGetToken(Mid$(sVarName, 3), 1, "=")).Value = MachineName & " " & Index
                    End If
                 Else
                    If StrComp(sGetToken(sVarName, 1, "="), "txtFirstName") <> 0 Then
                       brwWebBrowser.Document.Forms(0).Item(sGetToken(sVarName, 1, "=")).Value = sAfter(sVarName, 1, "=")
                    Else
                       brwWebBrowser.Document.Forms(0).Item(sGetToken(sVarName, 1, "=")).Value = MachineName & " " & Index
                    End If
                 End If
              Loop
              
              If Len(ButtonToPress) Then
                 Select Case UCase$(sGetToken(ButtonToPress))
                        Case "SUBMIT"
                             Select Case UCase$(sGetToken(ButtonToPress, 2))
                                    Case "FORM"
                                         sVarName = sGetToken(ButtonToPress, 3)
                                         Call brwWebBrowser.Document.Forms(CInt(sVarName)).submit
                             End Select
                 End Select
              End If
           End If
        End If
    Else
        Caption = "Working..."
    End If

EH_frmBrowser_timTimer_Timer_Continue:
    Exit Sub

EH_frmBrowser_timTimer_Timer:
    LogError "frmBrowser", "timTimer_Timer", Err.Number, Err.Description
    Resume EH_frmBrowser_timTimer_Timer_Continue

    Resume
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
On Error Resume Next
    Dim sTemp As String
     
    timTimer.Enabled = True: timTimer.Tag = vbNullString
     
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            brwWebBrowser.GoHome
        Case "Search"
            brwWebBrowser.GoSearch
        Case "Stop"
            timTimer.Enabled = False: timTimer.Tag = vbNullString
            brwWebBrowser.Stop
            Caption = brwWebBrowser.LocationName

'        Case "Get List"
'             sTemp = Replace(ListToString(cboAddress), frmMain.txtPrefix, vbNullString)
'             StringToList sTemp, frmMain.lstURLs
'        Case "Append"
'             sTemp = Replace(ListToString(cboAddress), frmMain.txtPrefix, vbNullString)
'             StringToList sTemp, frmMain.lstURLs, False
'        Case "Form Data"
'             Clipboard.SetText CollectFormData(brwWebBrowser)
    End Select
    Me.ZOrder
End Sub

