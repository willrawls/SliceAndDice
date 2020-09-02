VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmListSelect 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Select one"
   ClientHeight    =   5610
   ClientLeft      =   2175
   ClientTop       =   1875
   ClientWidth     =   7110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lstChoose 
      Height          =   5115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   9022
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imlSmallIcons"
      SmallIcons      =   "imlSmallIcons"
      ColHdrIcons     =   "imlSmallIcons"
      ForeColor       =   0
      BackColor       =   -2147483624
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   5130
      Width           =   1095
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   5130
      Width           =   1095
   End
   Begin MSComctlLib.ImageList imlSmallIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":0000
            Key             =   "Timer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":0454
            Key             =   "Category"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":08B4
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":0D08
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":115C
            Key             =   "!"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":18B0
            Key             =   "LightOff"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":1D04
            Key             =   "LightOn"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":2158
            Key             =   "Plus"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":25AC
            Key             =   "Minus"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":2A00
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":2E54
            Key             =   "Binoculars"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":32A8
            Key             =   "DocumentAlternate2"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":36FC
            Key             =   "BookOpenAngled"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":3B50
            Key             =   "BookOpen"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":3FA4
            Key             =   "BookClosed"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":43F8
            Key             =   "IndexCard"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":484C
            Key             =   "DocumentAlternate"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSelect.frx":4CA0
            Key             =   "Document"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmListSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Choices As CAssocArray

Private m_sChoice As String

Public Key As String

Implements SandySupport.ISandyWindowSelect

Public Property Get ISandyWindowSelect_Choice() As String
    ISandyWindowSelect_Choice = m_sChoice
End Property


Private Sub cmdCancel_Click()
    m_sChoice = vbNullString
    Hide
End Sub

Private Sub cmdOkay_Click()
    If lstChoose.SelectedItem Is Nothing Then
   'If lstChoose.ListIndex < 0 Then
       MsgBox "Please choose one before pressing OK.", vbInformation
       Exit Sub
    End If
    
    m_sChoice = lstChoose.SelectedItem.Text

    Form_Unload 0
    Hide
End Sub

Private Property Set ISandyWindowSelect_Choices(ByVal RHS As SandySupport.CAssocArray)
    Set Choices = RHS
End Property

Private Property Get ISandyWindowSelect_Choices() As SandySupport.CAssocArray
    Set ISandyWindowSelect_Choices = Choices
End Property

Private Sub ISandyWindowSelect_Hide()
    Me.Hide
End Sub

Public Sub ISandyWindowSelect_Initialize(sChoices As String, Optional ByVal sDelimiter As String = ";", Optional ByVal sDefault As String)
On Error Resume Next
    With Choices
         .ItemDelimiter = sDelimiter
         .All = sChoices
         Key = Left$(sChoices, 25)
         If Len(Key) > 0 Then
            LoadFormPosition Me, , , Key
         Else
            LoadFormPosition Me
         End If
         If InStr(sChoices, .KeyValueDelimiter) And Len(sDelimiter) > 0 Then
            lstChoose.ColumnHeaders.Clear
            lstChoose.ColumnHeaders.Add , , , 2800
            lstChoose.ColumnHeaders.Add , , , 9000
         Else
            lstChoose.ColumnHeaders.Clear
            lstChoose.ColumnHeaders.Add , , , 7000
         End If
         .FillListView lstChoose
         SetListViewIndex lstChoose, sDefault
    End With
End Sub

Private Sub Form_Initialize()
    Set Choices = CreateObject("SandySupport.CAssocArray")
    ' LogEvent "frmListSelect: Initialize"
End Sub

Private Sub Form_Load()
    With lstChoose.ColumnHeaders
         .Clear
         .Add , , "Title", 2800
         .Add , , "Description", 9000
    End With
    Form_Resize
End Sub

Private Sub Form_Resize()
    cmdOkay.Move 0, ScaleHeight - cmdOkay.Height
    cmdCancel.Move ScaleWidth - cmdCancel.Width, cmdOkay.Top
    lstChoose.Move 0, 0, ScaleWidth, ScaleHeight - cmdOkay.Height
End Sub


Private Sub Form_Terminate()
    Set Choices = Nothing
    ' LogEvent "frmListSelect: Terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Len(Key) Then
       SaveFormPosition Me, Key
    Else
       SaveFormPosition Me
    End If
End Sub

Private Property Let ISandyWindowSelect_Key(ByVal RHS As String)
    Key = RHS
End Property

Private Property Get ISandyWindowSelect_Key() As String
    ISandyWindowSelect_Key = Key
End Property

Private Sub ISandyWindowSelect_Show(Optional ByVal ModalSetting As Integer, Optional ParentWindow As Object)
    If ParentWindow Is Nothing Then
       Me.Show ModalSetting
    Else
       Me.Show ModalSetting, ParentWindow
    End If
End Sub

Private Sub ISandyWindowSelect_ZOrder()
    Me.ZOrder
End Sub

Private Sub lstChoose_DblClick()
    cmdOkay_Click
End Sub

