VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmCommandHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Soft Command Reference"
   ClientHeight    =   6885
   ClientLeft      =   5280
   ClientTop       =   2655
   ClientWidth     =   7605
   Icon            =   "frmCommandHelp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   7605
   Begin MSComDlg.CommonDialog cdgHelp 
      Left            =   180
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSoftCommandName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   250
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   300
      Width           =   6600
   End
   Begin VB.TextBox txtAliases 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   6600
   End
   Begin VB.TextBox txtSyntax 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   900
      Width           =   6600
   End
   Begin VB.TextBox txtOneLineDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   6600
   End
   Begin VB.TextBox txtHelpFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1500
      Width           =   6600
   End
   Begin VB.TextBox txtHelpTopic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   6600
   End
   Begin VB.TextBox txtLongDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1200
      Left            =   1050
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2100
      Width           =   6600
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1200
      Left            =   1050
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3300
      Width           =   6600
   End
   Begin VB.TextBox txtSeeAlso 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1200
      Left            =   1050
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4500
      Width           =   6600
   End
   Begin VB.TextBox txtExamples 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1200
      Left            =   1050
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5700
      Width           =   6600
   End
   Begin VB.CheckBox chkIsInline 
      Appearance      =   0  'Flat
      Caption         =   "This is an Inline command if checked ( ie. %%Command::Parameters%% )"
      CausesValidation=   0   'False
      Enabled         =   0   'False
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
      Height          =   300
      Left            =   1050
      TabIndex        =   0
      Top             =   0
      Width           =   6870
   End
   Begin VB.PictureBox picTH 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   60
      ScaleHeight     =   645
      ScaleWidth      =   765
      TabIndex        =   22
      Top             =   2460
      Width           =   765
   End
   Begin VB.Label lblSoftCommandName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   21
      Top             =   330
      Width           =   450
   End
   Begin VB.Label lblAliases 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aliases"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   20
      Top             =   630
      Width           =   570
   End
   Begin VB.Label lblSyntax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Syntax"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   19
      Top             =   930
      Width           =   540
   End
   Begin VB.Label lblOneLineDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   18
      Top             =   1230
      Width           =   765
   End
   Begin VB.Label lblHelpFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help File"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   17
      Top             =   1530
      Width           =   720
   End
   Begin VB.Label lblHelpTopic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help Topic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   16
      Top             =   1830
      Width           =   825
   End
   Begin VB.Label lblLongDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   15
      Top             =   2130
      Width           =   900
   End
   Begin VB.Label lblComments 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   14
      Top             =   3330
      Width           =   855
   End
   Begin VB.Label lblSeeAlso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "See Also"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   13
      Top             =   4530
      Width           =   675
   End
   Begin VB.Label lblExamples 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Examples"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   12
      Top             =   5730
      Width           =   735
   End
   Begin VB.Label lblIsInline 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Is Inline"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   11
      Top             =   60
      Width           =   660
   End
   Begin VB.Menu mnuFileExit 
      Caption         =   "&X"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFirst 
      Caption         =   "Fi&rst"
   End
   Begin VB.Menu mnuPrevious 
      Caption         =   "&<<"
   End
   Begin VB.Menu mnuNext 
      Caption         =   "&>>"
   End
   Begin VB.Menu mnuLast 
      Caption         =   "&Last"
   End
   Begin VB.Menu mnuFileFind 
      Caption         =   "&Find"
   End
   Begin VB.Menu mnuChangeCommandSets 
      Caption         =   "&Command Set"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmCommandHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SadCommandSet As CSadCommands
Public CurrCommand As CSadCommand
Public vCurrCommandKey As Variant

Implements SandySupport.ISandyWindowHelp
Public Property Let CurrCommandKey(ByVal vKey As Variant)
On Error Resume Next
    If Not SadCommandSet.Item(vKey) Is Nothing Then
       Set CurrCommand = SadCommandSet.Item(vKey)
       vCurrCommandKey = CurrCommand.Index
       Populate
    End If
    Err.Clear
End Property


Public Sub Populate()
On Error Resume Next
    Dim NextTop As Long
    
    If SadCommandSet Is Nothing Then Exit Sub
    If CurrCommand Is Nothing Then Exit Sub
    
    With CurrCommand
         Caption = "SAD Soft Command Reference - " & SadCommandSet.Attributes("Name") & " ( " & SadCommandSet.Count & " commands)"
         chkIsInline.Enabled = True
         chkIsInline.Value = Abs(.IsInline)
         If .IsInline Then
            chkIsInline.Caption = "This is an INLINE Soft Command (ie. %%" & .SoftCommandName & "::Parameters%% )"
         Else
            chkIsInline.Caption = "This is a REGULAR Soft Command (ie. ~~" & .SoftCommandName & " )"
         End If

         NextTop = 390
         If Len(.SoftCommandName) Then
            txtSoftCommandName.Text = .SoftCommandName
            txtSoftCommandName.Visible = True
            txtSoftCommandName.Top = NextTop
            lblSoftCommandName.Visible = True
            lblSoftCommandName.Top = NextTop
            txtSoftCommandName.Height = picTH.TextHeight(txtSoftCommandName.Text) + 100
            NextTop = NextTop + txtSoftCommandName.Height
         Else
            txtSoftCommandName.Visible = False
            lblSoftCommandName.Visible = False
         End If

         If Len(.Aliases) Then
            txtAliases.Text = Mid$(.Aliases, 3, Len(.Aliases) - 4)
            txtAliases.Visible = True
            txtAliases.Top = NextTop
            lblAliases.Visible = True
            lblAliases.Top = NextTop
            txtAliases.Height = picTH.TextHeight(txtAliases.Text) + 100
            NextTop = NextTop + txtAliases.Height
         Else
            txtAliases.Visible = False
            lblAliases.Visible = False
         End If

         If Len(.Syntax) Then
            txtSyntax.Text = .Syntax
            txtSyntax.Visible = True
            txtSyntax.Top = NextTop
            lblSyntax.Visible = True
            lblSyntax.Top = NextTop
            txtSyntax.Height = picTH.TextHeight(txtSyntax.Text) + 100
            NextTop = NextTop + txtSyntax.Height
         Else
            txtSyntax.Visible = False
            lblSyntax.Visible = False
         End If

         If Len(.OneLineDescription) Then
            txtOneLineDescription.Text = .OneLineDescription
            txtOneLineDescription.Visible = True
            txtOneLineDescription.Top = NextTop
            lblOneLineDescription.Visible = True
            lblOneLineDescription.Top = NextTop
            txtOneLineDescription.Height = picTH.TextHeight(txtOneLineDescription.Text) + 100
            NextTop = NextTop + txtOneLineDescription.Height
         Else
            txtOneLineDescription.Visible = False
            lblOneLineDescription.Visible = False
         End If

         If Len(.LongDescription) Then
            txtLongDescription.Text = .LongDescription
            txtLongDescription.Visible = True
            txtLongDescription.Top = NextTop
            lblLongDescription.Visible = True
            lblLongDescription.Top = NextTop
            txtOneLineDescription.Height = picTH.TextHeight(txtOneLineDescription.Text) + 100
            NextTop = NextTop + txtLongDescription.Height
         Else
            txtLongDescription.Visible = False
            lblLongDescription.Visible = False
         End If

         If Len(.Comments) Then
            txtComments.Text = .Comments
            txtComments.Visible = True
            txtComments.Top = NextTop
            lblComments.Visible = True
            lblComments.Top = NextTop
            txtComments.Height = picTH.TextHeight(txtComments.Text) + 100
            NextTop = NextTop + txtComments.Height
         Else
            txtComments.Visible = False
            lblComments.Visible = False
         End If

         If Len(.SeeAlso) Then
            txtSeeAlso.Text = .SeeAlso
            txtSeeAlso.Visible = True
            txtSeeAlso.Top = NextTop
            lblSeeAlso.Visible = True
            lblSeeAlso.Top = NextTop
            txtSeeAlso.Height = picTH.TextHeight(txtSeeAlso.Text) + 100
            NextTop = NextTop + txtSeeAlso.Height
         Else
            txtSeeAlso.Visible = False
            lblSeeAlso.Visible = False
         End If

         If Len(.Examples) Then
            txtExamples.Text = .Examples
            txtExamples.Visible = True
            txtExamples.Top = NextTop
            lblExamples.Visible = True
            lblExamples.Top = NextTop
            txtExamples.Height = picTH.TextHeight(txtExamples.Text) + 100
            NextTop = NextTop + txtExamples.Height
            lblExamples.ZOrder
         Else
            txtExamples.Visible = False
            lblExamples.Visible = False
         End If

         If Len(.HelpFile) Then
            txtHelpFile.Text = .HelpFile
            txtHelpFile.Visible = True
            txtHelpFile.Top = NextTop
            lblHelpFile.Visible = True
            lblHelpFile.Top = NextTop
            txtHelpFile.Height = picTH.TextHeight(txtHelpFile.Text) + 100
            NextTop = NextTop + txtHelpFile.Height
         Else
            txtHelpFile.Visible = False
            lblHelpFile.Visible = False
         End If

         If .HelpTopic Then
            txtHelpTopic.Text = .HelpTopic
            txtHelpTopic.Visible = True
            txtHelpTopic.Top = NextTop
            lblHelpTopic.Visible = True
            lblHelpTopic.Top = NextTop
            txtHelpTopic.Height = picTH.TextHeight(txtHelpTopic.Text) + 100
            NextTop = NextTop + txtHelpTopic.Height
         Else
            txtHelpTopic.Visible = False
            lblHelpTopic.Visible = False
         End If
         Me.Height = NextTop + 600
         picTH.Move txtHelpFile.Left, NextTop
    End With
End Sub

Private Sub chkIsInline_Click()
On Error Resume Next
    chkIsInline.Value = Abs(CurrCommand.IsInline)
End Sub

Private Sub Form_Initialize()

    ' LogEvent "frmCommandHelp: Initialize"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If (Shift And vbCtrlMask) > 0 Then
       Select Case KeyCode
              Case vbKeyPageUp:    mnuPrevious_Click
              Case vbKeyPageDown:  mnuNext_Click
              Case vbKeyF:         mnuFileFind_Click
       End Select
    ElseIf Shift = 0 Then
       Select Case KeyCode
              Case vbKeyEscape:       mnuFileExit_Click
       End Select
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    LoadFormPosition Me, , False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
       Cancel = True
    End If
    mnuFileExit_Click
End Sub

Private Sub Form_Terminate()

    ' LogEvent "frmCommandHelp: Terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    SaveFormPosition Me
End Sub


Private Property Let ISandyWindowHelp_BackColor(ByVal RHS As stdole.OLE_COLOR)
    BackColor = RHS
End Property

Private Property Get ISandyWindowHelp_BackColor() As stdole.OLE_COLOR
    ISandyWindowHelp_BackColor = BackColor
End Property

Private Property Set ISandyWindowHelp_CurrCommand(ByVal RHS As SandySupport.CSadCommand)
    Set CurrCommand = RHS
End Property

Private Property Get ISandyWindowHelp_CurrCommand() As SandySupport.CSadCommand
    Set ISandyWindowHelp_CurrCommand = CurrCommand
End Property

Private Property Let ISandyWindowHelp_CurrCommandKey(ByVal RHS As Variant)
    CurrCommandKey = RHS
End Property

Private Sub ISandyWindowHelp_FileExit()
    mnuFileExit_Click
End Sub


Private Property Let ISandyWindowHelp_ForeColor(ByVal RHS As stdole.OLE_COLOR)
    ForeColor = RHS
End Property

Private Property Get ISandyWindowHelp_ForeColor() As stdole.OLE_COLOR
    ISandyWindowHelp_ForeColor = ForeColor
End Property


Private Sub ISandyWindowHelp_Hide()
    Hide
End Sub

Private Sub ISandyWindowHelp_Populate()
    Populate
End Sub


Private Property Set ISandyWindowHelp_SadCommandSet(ByVal RHS As SandySupport.CSadCommands)
    Set SadCommandSet = RHS
End Property


Private Property Get ISandyWindowHelp_SadCommandSet() As SandySupport.CSadCommands
    Set ISandyWindowHelp_SadCommandSet = SadCommandSet
End Property

Private Sub ISandyWindowHelp_Show(Optional ByVal ModalSetting As Integer, Optional ParentWindow As Object)
    If IsMissing(ParentWindow) Then
       Show ModalSetting
    Else
       Show ModalSetting, ParentWindow
    End If
End Sub

Private Property Set ISandyWindowHelp_vCurrCommandKey(RHS As Variant)
    Set vCurrCommandKey = RHS
End Property

Private Property Let ISandyWindowHelp_vCurrCommandKey(RHS As Variant)
    vCurrCommandKey = RHS
End Property

Private Property Get ISandyWindowHelp_vCurrCommandKey() As Variant
    ISandyWindowHelp_vCurrCommandKey = vCurrCommandKey
End Property


Private Sub ISandyWindowHelp_ZOrder()
    ZOrder
End Sub

Public Sub mnuFileExit_Click()
On Error Resume Next
    SaveFormPosition Me
    Set CurrCommand = Nothing
    Set SadCommandSet = Nothing
    Hide
End Sub

Private Sub mnuFileFind_Click()
On Error Resume Next
    Dim CurrMember As CSadCommand
    Dim sChoices As String
    Dim sChoice As String
    Dim asaOrdered As SandySupport.CAssocArray

    sChoices = vbNullString
    For Each CurrMember In SadCommandSet
        sChoices = sChoices & CurrMember.SoftCommandName & IIf(CurrMember.IsInline, " (Inline)" & "=" & CurrMember.OneLineDescription & ";", "=" & CurrMember.OneLineDescription & ";")
        If Len(CurrMember.Aliases) Then
           sChoices = sChoices & Replace(CurrMember.Aliases, ", ", IIf(CurrMember.IsInline, " (Inline)" & "=See " & CurrMember.SoftCommandName & ";", "=See " & CurrMember.SoftCommandName & ";"))
        End If
    Next CurrMember
    If Len(sChoices) Then
       Set asaOrdered = CreateObject("SandySupport.CAssocArray")
       With asaOrdered
            .ItemDelimiter = ";"
            .KeyValueDelimiter = "="
            .AddInOrder = True
            .All = Replace(Replace(sChoices, "; (Inline)", ";"), ";;", ";")
            sChoices = .All
       End With
       Set asaOrdered = Nothing
       sChoice = sChoose(sChoices, , txtSoftCommandName.Text)

       If Len(sChoice) Then
          If InStr(sChoice, "(Inline)") Then
             CurrCommandKey = SadCommandSet.Item(Trim$(sGetToken(UCase$(sChoice), 1, " (INLINE)")) & "*I").Index
          Else
             CurrCommandKey = SadCommandSet.Item(UCase$(sChoice) & "*C").Index
          End If
       End If
    End If
End Sub

Private Sub mnuFirst_Click()
On Error Resume Next
    CurrCommandKey = 1
End Sub

Private Sub mnuLast_Click()
On Error Resume Next
    CurrCommandKey = SadCommandSet.Count
End Sub

Private Sub mnuNext_Click()
On Error Resume Next
    If vCurrCommandKey + 1 < SadCommandSet.Count Then
       CurrCommandKey = vCurrCommandKey + 1
    End If
End Sub

Private Sub mnuPrevious_Click()
On Error Resume Next
    If vCurrCommandKey - 1 > 0 Then
       CurrCommandKey = vCurrCommandKey - 1
    End If
End Sub

Private Sub txtHelpFile_Click()
    txtHelpTopic_Click
End Sub


Private Sub txtHelpFile_DblClick()
    txtHelpTopic_Click
End Sub


Private Sub txtHelpTopic_Click()
    If Len(txtHelpFile) > 0 And Len(txtHelpTopic) = 0 Then
       With cdgHelp
          .HelpFile = txtHelpFile
          ' Go to the Click Event topic in the Help file.
          ' The number is determined in the [MAP] section
          ' of the .HPJ file for the .chm file. You can
          ' edit this number only if you are using the
          ' Microsoft Help Workshop to build your
          ' own Help file.
          .HelpContext = txtHelpTopic
          .HelpCommand = cdlHelpContext
          .ShowHelp
        End With
    ElseIf Len(txtHelpTopic) > 0 Then
        BrowseTo txtHelpTopic
    End If
End Sub

Private Sub txtHelpTopic_DblClick()
    txtHelpTopic_Click
End Sub


