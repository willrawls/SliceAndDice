VERSION 5.00
Begin VB.Form frmFieldType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Field Name, type, size"
   ClientHeight    =   3705
   ClientLeft      =   5205
   ClientTop       =   4005
   ClientWidth     =   2670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox FieldName 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   270
      Width           =   2565
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   1770
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   1770
      TabIndex        =   3
      Top             =   900
      Width           =   855
   End
   Begin VB.TextBox txtLength 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmFieldType.frx":0000
      Top             =   3330
      Width           =   1605
   End
   Begin VB.ListBox lstType 
      Height          =   2070
      IntegralHeight  =   0   'False
      ItemData        =   "frmFieldType.frx":0004
      Left            =   60
      List            =   "frmFieldType.frx":0026
      TabIndex        =   1
      Top             =   900
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Field Name:"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Field Length:"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   3060
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Field Type:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   660
      Width           =   780
   End
End
Attribute VB_Name = "frmFieldType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bCanceled As Boolean

Public Property Get Canceled() As Boolean
1        Canceled = m_bCanceled
End Property

Public Property Get dbFieldType() As DataTypeEnum
    Select Case lstType
        Case "Text": dbFieldType = dbText
2            Case "Long Integer": dbFieldType = dbLong
3            Case "Boolean": dbFieldType = dbBoolean
4            Case "Currency": dbFieldType = dbCurrency
5            Case "Date/Time": dbFieldType = dbDate
6            Case "Double": dbFieldType = dbDouble
7            Case "Integer": dbFieldType = dbInteger
8            Case "Binary": dbFieldType = dbLongBinary
9            Case "Memo": dbFieldType = dbMemo
10           Case "Single": dbFieldType = dbSingle
11       End Select
End Property

Public Property Get Length() As Long
12       If lstType = "Text" Then
13           Length = Val(txtLength)
14       Else
15           Length = 0
16       End If
End Property

Private Sub cmdCancel_Click()
17       m_bCanceled = True
18       Hide
End Sub

Private Sub cmdOkay_Click()
19       FieldName = Trim$(FieldName)

20       If Len(FieldName) = 0 Then
21           MsgBox "Please enter a name before pressing 'OK'", vbInformation
22           Exit Sub
23       ElseIf lstType.ListIndex < 0 Then
24           MsgBox "Please select a field type before pressing 'OK'", vbInformation
25           Exit Sub
26       End If

27       m_bCanceled = False
28       Hide
End Sub

Private Sub Form_Activate()
29       On Error Resume Next
30       FieldName.SetFocus
End Sub

Private Sub Form_Load()
31       On Error Resume Next
32       If lstType.ListIndex = -1 Then
33           lstType.ListIndex = 0
34       End If
End Sub

Private Sub lstType_Click()
35       txtLength.Enabled = (lstType = "Text")
End Sub


Private Sub lstType_DblClick()
36       cmdOkay_Click
End Sub


Private Sub txtLength_Change()
37       If CStr(Val(txtLength)) <> txtLength Then
38           If Val(txtLength) < 1 Then
39               txtLength = "1"
40           ElseIf Val(txtLength) > 255 Then
41               txtLength = "255"
42           Else
43               txtLength = CStr(Val(txtLength))
44           End If
45       End If
End Sub


