VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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

Public Choices As New CAssocArray

Private m_sChoice As String

Public Key As String
Public Property Get Choice() As String
1        Choice = m_sChoice
End Property


Private Sub cmdCancel_Click()
2        m_sChoice = vbNullString
3        Hide
End Sub

Private Sub cmdOkay_Click()
4        If lstChoose.SelectedItem Is Nothing Then
        'If lstChoose.ListIndex < 0 Then
5            MsgBox "Please choose one before pressing OK.", vbInformation
6            Exit Sub
7        End If

8        m_sChoice = lstChoose.SelectedItem.Text

9        Form_Unload 0
10       Hide
End Sub

Public Sub Initialize(sChoices As String, Optional ByVal sDelimiter As String = gsSC, Optional ByVal sDefault As String)
11       On Error Resume Next
12       With Choices
13           .ItemDelimiter = sDelimiter
14           .All = sChoices
15           Key = Left$(sChoices, 25)
16           If Len(Key) > 0 Then
17               LoadFormPosition Me, , , Key
18           Else
19               LoadFormPosition Me
20           End If
21           If InStr(sChoices, .KeyValueDelimiter) And Len(sDelimiter) > 0 Then
22               lstChoose.ColumnHeaders.Clear
23               lstChoose.ColumnHeaders.Add , , , 2800
24               lstChoose.ColumnHeaders.Add , , , 9000
25           Else
26               lstChoose.ColumnHeaders.Clear
27               lstChoose.ColumnHeaders.Add , , , 7000
28           End If
29           .FillListView lstChoose
30           SetListViewIndex lstChoose, sDefault
31       End With
End Sub

Private Sub Form_Load()
32       With lstChoose.ColumnHeaders
33           .Clear
34           .Add , , "Title", 2800
35           .Add , , "Description", 9000
36       End With
37       Form_Resize
End Sub

Private Sub Form_Resize()
38       cmdOkay.Move 0, ScaleHeight - cmdOkay.Height
39       cmdCancel.Move ScaleWidth - cmdCancel.Width, cmdOkay.Top
40       lstChoose.Move 0, 0, ScaleWidth, ScaleHeight - cmdOkay.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
41       If Len(Key) Then
42           SaveFormPosition Me, Key
43       Else
44           SaveFormPosition Me
45       End If
End Sub

Private Sub lstChoose_DblClick()
46       cmdOkay_Click
End Sub

