VERSION 5.00
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "VSOCX32.OCX"
Begin VB.Form frmAdvTable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Table Settings"
   ClientHeight    =   2205
   ClientLeft      =   1155
   ClientTop       =   2640
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBorderWidth 
      Height          =   315
      Left            =   2175
      TabIndex        =   10
      Text            =   "3"
      Top             =   1170
      Width           =   435
   End
   Begin VB.TextBox txtCellSpacing 
      Height          =   300
      Left            =   2175
      TabIndex        =   8
      Text            =   "3"
      Top             =   1815
      Width           =   435
   End
   Begin VB.TextBox txtCellPadding 
      Height          =   315
      Left            =   2175
      TabIndex        =   6
      Text            =   "3"
      Top             =   1485
      Width           =   435
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   495
      Left            =   3465
      TabIndex        =   3
      Top             =   1605
      Width           =   1035
   End
   Begin VB.TextBox txtWidth 
      Height          =   300
      Left            =   4200
      TabIndex        =   2
      Text            =   "100"
      Top             =   570
      Width           =   435
   End
   Begin VB.CheckBox chkSpreadEvenly 
      Caption         =   "Spread cells evenly across the width"
      Height          =   255
      Left            =   345
      TabIndex        =   1
      Top             =   840
      Value           =   1  'Checked
      Width           =   2985
   End
   Begin VsOcxLib.VideoSoftElastic vseBackground 
      Height          =   375
      Left            =   2085
      TabIndex        =   0
      Top             =   75
      Width           =   765
      _Version        =   327680
      _ExtentX        =   1349
      _ExtentY        =   661
      _StockProps     =   70
      Caption         =   "Clear"
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
      BorderWidth     =   4
      CaptionPos      =   4
      Style           =   3
      PicturePos      =   3
      CornerColor     =   12632256
      ShowFocusRect   =   -1  'True
      ShowOutline     =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Border width:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   735
      TabIndex        =   11
      Top             =   1200
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cell spacing:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   720
      TabIndex        =   9
      Top             =   1845
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cell padding:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   720
      TabIndex        =   7
      Top             =   1515
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Table width (in percent of screen width):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   5
      Top             =   570
      Width           =   4110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Background Color:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   75
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmAdvTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOkay_Click()
    txtWidth = CStr(Val(txtWidth))
    If Val(txtWidth) < 1 Or Val(txtWidth) > 100 Then
       MsgBox "Please enter a percentage for the table width (1-100).", vbInformation
       Exit Sub
    End If
    
    Hide
End Sub


Private Sub txtBorderWidth_LostFocus()
    txtBorderWidth = Val(txtBorderWidth)
End Sub


Private Sub txtCellPadding_LostFocus()
    txtCellPadding = Val(txtCellPadding)
End Sub


Private Sub txtCellSpacing_LostFocus()
    txtCellSpacing = Val(txtCellSpacing)
End Sub


Private Sub txtWidth_LostFocus()
    txtWidth = Val(txtWidth)
End Sub


Private Sub vseBackground_Click()
    If vseBackground.Caption = vbNullString Then
       If MsgBox("Set to clear ?", vbYesNo) = vbYes Then
          vseBackground.BackColor = &H8000000F
          vseBackground.Caption = "Clear"
          Exit Sub
       End If
    End If
    
    frmMain.cdgBrowse.ShowColor
    vseBackground.BackColor = frmMain.cdgBrowse.Color
    vseBackground.Caption = vbNullString
End Sub

