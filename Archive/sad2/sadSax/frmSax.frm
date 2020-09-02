VERSION 5.00
Object = "{B20ABC70-3855-11D3-8F7F-0000861EF01D}#1.0#0"; "SB6ENT.OCX"
Begin VB.Form frmSax 
   Caption         =   "Sandy Sax"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTest 
      Left            =   8820
      Top             =   90
   End
   Begin Sb6entCtl.BasicIdeCtl saxScript 
      CausesValidation=   0   'False
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   9315
      _cx             =   16431
      _cy             =   11668
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sandy Sax"
      EditTools       =   -1  'True
      FileDesc        =   "Sandy Sax Script"
      FileExt         =   "sss"
      FileTools       =   -1  'True
      FileMenu        =   -1  'True
      HiddenCode      =   ""
      StatusVisible   =   -1  'True
      ToolbarVisible  =   -1  'True
      ReservedColor   =   16711680
      ExtensionColor  =   8388608
      BuiltinColor    =   8421376
      CommentColor    =   32768
      ErrorColor      =   255
      AlwaysSplit     =   -1  'True
      EventMode       =   0   'False
      FileChangeDir   =   -1  'True
      FullPopupMenu   =   -1  'True
      Locked          =   0   'False
      MultiSheet      =   -1  'True
      NegotiateMenus  =   -1  'True
      BlockedKeywords =   ""
      DefaultMacroName=   "Macro"
      DefaultObjectName=   "Object.obm|Object"
      ProcDisplayMode =   2
      DefaultDataType =   ""
      DebugHeight     =   0
      HelpMenu        =   -1  'True
      BreakColor      =   128
      ExecColor       =   65535
      TabAsSpaces     =   0   'False
      TabWidth        =   4
      DesignModeVisible=   -1  'True
      FileMRULimit    =   8
      TaskbarIconMode =   7
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      LargeIcon       =   "frmSax.frx":0000
      SmallIcon       =   "frmSax.frx":031A
      TaskbarIcon     =   "frmSax.frx":04F4
      HideSelection   =   -1  'True
   End
End
Attribute VB_Name = "frmSax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Parent As NewCommands

Private Sub Form_Load()

    LoadFormPosition Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    Select Case saxScript.Shutdown
    Case -1
        tmrTest.Interval = 100
        Cancel = True
    Case 0
    Case 1
        Cancel = True
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
    saxScript.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Cancel = Not saxScript.Disconnect
    SaveFormPosition Me
End Sub

Private Sub tmrTest_Timer()
On Error Resume Next
    tmrTest.Interval = 0
    If Not saxScript.Run Then Unload Me
End Sub
