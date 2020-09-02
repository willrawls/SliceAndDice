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
    Canceled = m_bCanceled
End Property

Public Property Get dbFieldType() As DataTypeEnum
       Select Case lstType
              Case "Text":          dbFieldType = dbText
              Case "Long Integer":  dbFieldType = dbLong
              Case "Boolean":       dbFieldType = dbBoolean
              Case "Currency":      dbFieldType = dbCurrency
              Case "Date/Time":     dbFieldType = dbDate
              Case "Double":        dbFieldType = dbDouble
              Case "Integer":       dbFieldType = dbInteger
              Case "Binary":        dbFieldType = dbLongBinary
              Case "Memo":          dbFieldType = dbMemo
              Case "Single":        dbFieldType = dbSingle
        End Select
End Property

Public Property Get Length() As Long
    If lstType = "Text" Then
       Length = Val(txtLength)
    Else
       Length = 0
    End If
End Property

Private Sub cmdCancel_Click()
    m_bCanceled = True
    Hide
End Sub

Private Sub cmdOkay_Click()
    FieldName = Trim$(FieldName)

    If Len(FieldName) = 0 Then
       MsgBox "Please enter a name before pressing 'OK'", vbInformation
       Exit Sub
    ElseIf lstType.ListIndex < 0 Then
       MsgBox "Please select a field type before pressing 'OK'", vbInformation
       Exit Sub
    End If

    m_bCanceled = False
    Hide
End Sub

Private Sub Form_Activate()
On Error Resume Next
    FieldName.SetFocus
End Sub

Private Sub Form_Initialize()

    ' LogEvent "frmFieldType: Initialize"
End Sub

Private Sub Form_Load()
On Error Resume Next
   If lstType.ListIndex = -1 Then
      lstType.ListIndex = 0
   End If
End Sub

Private Sub Form_Terminate()

    ' LogEvent "frmFieldType: Terminate"
End Sub

Private Sub lstType_Click()
    txtLength.Enabled = (lstType = "Text")
End Sub


Private Sub lstType_DblClick()
    cmdOkay_Click
End Sub


Private Sub txtLength_Change()
    If CStr(Val(txtLength)) <> txtLength Then
       If Val(txtLength) < 1 Then
          txtLength = "1"
       ElseIf Val(txtLength) > 255 Then
          txtLength = "255"
       Else
          txtLength = CStr(Val(txtLength))
       End If
    End If
End Sub


