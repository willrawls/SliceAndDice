VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl FSListBar 
   Alignable       =   -1  'True
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3570
   ScaleWidth      =   3120
   ToolboxBitmap   =   "FSListBar.ctx":0000
   Begin VB.PictureBox picChoose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   210
      ScaleHeight     =   3135
      ScaleWidth      =   2895
      TabIndex        =   2
      Top             =   300
      Width           =   2895
      Begin VB.ListBox lstChoose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   3075
         IntegralHeight  =   0   'False
         ItemData        =   "FSListBar.ctx":0532
         Left            =   60
         List            =   "FSListBar.ctx":0534
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   30
         Width           =   2835
      End
   End
   Begin VB.CommandButton cmdChangeCategory 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2790
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Click once to select a different category."
      Top             =   0
      Width           =   375
   End
   Begin MSComctlLib.TreeView tvwArray 
      Height          =   3135
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   300
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   5530
      _Version        =   393217
      Indentation     =   265
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imlSmallIcons"
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwArray 
      Height          =   3255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Tag             =   "List"
      Top             =   300
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5741
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imlSmallIcons"
      SmallIcons      =   "imlSmallIcons"
      ColHdrIcons     =   "imlSmallIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imlSmallIcons 
      Left            =   2460
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":0536
            Key             =   "Timer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":098A
            Key             =   "Category"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":0DDE
            Key             =   "Categoryx"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":123E
            Key             =   "Keyx"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":1692
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":1AE6
            Key             =   "!"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":223A
            Key             =   "LightOff"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":268E
            Key             =   "LightOn"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":2AE2
            Key             =   "DocumentAlternate"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":2F36
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":338A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":37DE
            Key             =   "Binoculars"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":3C32
            Key             =   "DocumentAlternate2"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":4086
            Key             =   "BookOpenAngled"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":44DA
            Key             =   "BookOpen"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":492E
            Key             =   "BookClosed"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":4D82
            Key             =   "IndexCard"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":51D6
            Key             =   "DocumentAlternatex"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSListBar.ctx":562A
            Key             =   "Document"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.UpDown updCategory 
      Height          =   285
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Up to go to previous category, Down to go to next category"
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      Min             =   1
      Wrap            =   -1  'True
      Enabled         =   -1  'True
   End
   Begin VB.Label lblArray 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bar 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "FSListBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Firm Solutions List Bar"
Option Explicit

' Default Property Values:
  Private Const m_def_CurBar = 0
  Private Const m_def_CurBarItem = -1
  Private Const m_def_CaptionHeight As Long = 375

' Property Variables:
  Private m_Value         As Variant
  Private m_CurBar        As Long
  Private m_CurBarItem    As Long
  Private m_BarCount      As Long
  Private m_CaptionHeight As Long

' Event Declarations:
  Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=lvwArray(0),lvwArray,0,KeyDown
  Event KeyPress(KeyAscii As Integer) 'MappingInfo=lvwArray(0),lvwArray,0,KeyPress
  Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=lvwArray(0),lvwArray,0,KeyUp
  Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvwArray(0),lvwArray,0,MouseDown
  Event MouseDownOnCategory(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvwArray(0),lvwArray,0,MouseDown
  Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvwArray(0),lvwArray,0,MouseMove
  Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvwArray(0),lvwArray,0,MouseUp
  Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=lvwArray(0),lvwArray,0,OLEDragOver
  Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvwArray(0),lvwArray,0,OLEDragDrop
  Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=lvwArray(0),lvwArray,0,OLEGiveFeedback
  Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=lvwArray(0),lvwArray,0,OLEStartDrag
  Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=lvwArray(0),lvwArray,0,OLESetData
  Event OLECompleteDrag(Effect As Long) 'MappingInfo=lvwArray(0),lvwArray,0,OLECompleteDrag
  Event BeforeLabelEdit(Cancel As Integer) 'MappingInfo=lvwArray(0),lvwArray,0,BeforeLabelEdit
  Event AfterLabelEdit(Cancel As Integer, NewString As String) 'MappingInfo=lvwArray(0),lvwArray,0,AfterLabelEdit
  Event ItemClick(Item As ListItem) 'MappingInfo=lvwArray(0),lvwArray,0,ItemClick
  Event NodeClick(NodeClicked As Node)
  Event BeforeBarClick()
  Event AfterBarClick()
  Event BarItemDblClick(ByVal BarName As String, ByVal BarKey As String, ByVal BarItemName As String, ByVal BarItemKey As String)
  Event BarItemClick(ByVal BarName As String, ByVal BarKey As String, ByVal BarItemName As String, ByVal BarItemKey As String)

  Event BeforeDropDown()
  Event AfterDropDown()
  Event DropDownClick(Index As Long)

Private ItemJustChecked As Long

Public Property Get BarCount() As Long
Attribute BarCount.VB_MemberFlags = "400"
    BarCount = m_BarCount + 1
End Property

Public Property Let BarItemIcon(New_BarItemIcon As String)
On Error Resume Next
    If Not lvwArray(m_CurBar).SelectedItem Is Nothing Then
       lvwArray(m_CurBar).SelectedItem.Icon = New_BarItemIcon
       lvwArray(m_CurBar).SelectedItem.SmallIcon = New_BarItemIcon
    End If
End Property

Public Property Get BarItemIcon() As String
    If lvwArray(m_CurBar).Tag = "List" Then
       If Not lvwArray(m_CurBar).SelectedItem Is Nothing Then
          BarItemIcon = lvwArray(m_CurBar).SelectedItem.Icon
       End If
    Else
       If Not tvwArray(m_CurBar).SelectedItem Is Nothing Then
          BarItemIcon = tvwArray(m_CurBar).SelectedItem.Image
       End If
    End If
End Property

Public Property Get BarKey() As String
    BarKey = lblArray(m_CurBar).Tag
End Property

Public Property Let BarKey(New_BarKey As String)
    lblArray(m_CurBar).Tag = New_BarKey
    PropertyChanged "BarKey"
End Property

Public Property Get Bars(ByVal Index As Variant) As Object
    Static nCurBar As Long

    If VarType(Index) = vbString Then
       For nCurBar = 0 To m_BarCount
           If lblArray(nCurBar).Caption = Index Or lblArray(nCurBar).Tag = Index Then
              If lvwArray(nCurBar).Tag = "List" Then
                 Set Bars = lvwArray(nCurBar)
              Else
On Error Resume Next
                 Set Bars = tvwArray(nCurBar)
              End If
              Exit Property
           End If
       Next nCurBar
    ElseIf VarType(Index) = vbLong Or VarType(Index) = vbInteger Then
       If Index >= 0 And Index <= m_BarCount Then
          If lvwArray(Index).Tag = "List" Then
             Set Bars = lvwArray(Index)
          Else
On Error Resume Next
             Set Bars = tvwArray(Index)
          End If
          Exit Property
       End If
    End If
    Set Bars = Nothing
End Property

Public Property Let CaptionHeight(New_CaptionHeight As Long)
    If New_CaptionHeight < 0 Then Exit Property

    If New_CaptionHeight > 5000 Then
       New_CaptionHeight = 5000
    End If

    m_CaptionHeight = New_CaptionHeight
    
    UserControl_Resize
    PropertyChanged "CaptionHeight"
End Property

Public Property Get CaptionHeight() As Long
    CaptionHeight = m_CaptionHeight
End Property

Public Sub Clear()
    Dim nCurBar As Long

On Error Resume Next
    For nCurBar = 1 To m_BarCount
        Unload lblArray(nCurBar)
        Unload lvwArray(nCurBar)
        Unload tvwArray(nCurBar)
    Next nCurBar

    With lblArray(0)
         .Tag = "Bar 1"
         .Caption = "Bar 1"
    End With

    With lvwArray(0)
         .ListItems.Clear
         .ColumnHeaders.Clear
         .ColumnHeaders.Add , "Name"
    End With

    tvwArray(0).Nodes.Clear

    m_CurBar = 0
    m_BarCount = 0

End Sub

Public Sub BarAndItem(ByVal Bar As Variant, ByVal Item As Variant)
    CurBar = Bar
    CurBarItem = Item
End Sub

Public Sub DisplayCategories()
    Static nCurBar As Long

    lstChoose.Clear
    For nCurBar = 0 To m_BarCount
        lstChoose.AddItem lblArray(nCurBar).Caption
        lstChoose.ItemData(lstChoose.NewIndex) = nCurBar
    Next nCurBar

    lstChoose.Height = (picChoose.TextHeight("W") + 30) * (m_BarCount + 1)
    picChoose.Height = lstChoose.Height + lstChoose.Top * 2
    picChoose.Visible = True
    picChoose.ZOrder
End Sub

Public Sub HideCategories()
On Error Resume Next
    picChoose.Visible = False
End Sub

Public Sub ResizeControls()
    Static nCurBar As Long
    Static nCurPosition As Long
    Static nFirstPosition As Long

On Error Resume Next

    With lblArray(m_CurBar)
         .Top = 0
         .ZOrder 0
         .Height = m_CaptionHeight
         .Visible = True
    End With

    With lvwArray(m_CurBar)
         .Top = lblArray(m_CurBar).Height
         .ZOrder 1
         .Height = ScaleHeight - lblArray(0).Height                 ' ((m_BarCount + 1) * lblArray(0).Height)
         .Visible = (.Tag = "List")
         nFirstPosition = lblArray(0).Height + .Height
    End With

    tvwArray(m_CurBar).Top = lblArray(m_CurBar).Height
    tvwArray(m_CurBar).ZOrder 1
    tvwArray(m_CurBar).Height = ScaleHeight - lblArray(0).Height                 ' ((m_BarCount + 1) * lblArray(0).Height)
    tvwArray(m_CurBar).Visible = (lvwArray(m_CurBar).Tag = "Tree")
    nFirstPosition = lblArray(0).Height + lvwArray(m_CurBar).Height

    nCurPosition = 0
    For nCurBar = 0 To m_BarCount
        lblArray(nCurBar).Width = ScaleWidth
        lblArray(nCurBar).Height = m_CaptionHeight
        lblArray(nCurBar).ZOrder 0
        lvwArray(nCurBar).Width = ScaleWidth
        tvwArray(nCurBar).Width = ScaleWidth
        If nCurBar <> m_CurBar Then
           lblArray(nCurBar).Visible = False
           lblArray(nCurBar).Top = nFirstPosition + lblArray(0).Height * nCurPosition
           lvwArray(nCurBar).Visible = False
           tvwArray(nCurBar).Visible = False
           nCurPosition = nCurPosition + 1
        End If
    Next nCurBar

    cmdChangeCategory.Top = 0
    cmdChangeCategory.Left = ScaleWidth - cmdChangeCategory.Width
    cmdChangeCategory.Height = m_CaptionHeight - 25

    picChoose.Move 250, cmdChangeCategory.Top + cmdChangeCategory.Height, ScaleWidth - 250 ', ScaleHeight - lstChoose.Top - 250
    lstChoose.Move 30, 30, picChoose.ScaleWidth
    
    updCategory.ZOrder 1
    updCategory.Height = m_CaptionHeight - 25
End Sub

Public Property Get Version() As String
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Sub cmdChangeCategory_Click()
On Error GoTo EH_cmdChangeCategory_Click
    UserControl.SetFocus
    If picChoose.Visible Then
       picChoose.Visible = False
       Exit Sub
    End If

    DisplayCategories
    
    lstChoose.SetFocus

EH_cmdChangeCategory_Click_Continue:
    Exit Sub
    
EH_cmdChangeCategory_Click:
    MsgBox "Something happened during cmdChangeCategory_Click"
    Resume EH_cmdChangeCategory_Click_Continue
    
    Resume
End Sub

Private Sub lblArray_Click(Index As Integer)
    picChoose.Visible = False
End Sub

Private Sub lblArray_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDownOnCategory(Button, Shift, x, y)
End Sub

Private Sub lstChoose_Click()
    picChoose.Visible = False
    
    If Len(lstChoose) = 0 Then Exit Sub

    SwitchToBar lstChoose.ListIndex
End Sub

Private Sub lstChoose_ItemCheck(Item As Integer)
    ItemJustChecked = Item
End Sub

Private Sub lvwArray_OLEDragOver(Index As Integer, Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub lvwArray_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    RaiseEvent ItemClick(Item)
End Sub

Private Sub lvwArray_OLEDragDrop(Index As Integer, Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub lvwArray_OLEStartDrag(Index As Integer, Data As MSComctlLib.DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub tvwArray_AfterLabelEdit(Index As Integer, Cancel As Integer, NewString As String)
    picChoose.Visible = False
    RaiseEvent AfterLabelEdit(Cancel, NewString)
End Sub

Private Sub tvwArray_BeforeLabelEdit(Index As Integer, Cancel As Integer)
    RaiseEvent BeforeLabelEdit(Cancel)
End Sub

Private Sub tvwArray_Click(Index As Integer)
On Error Resume Next
    Dim sText As String
    Dim sKey As String
    If Not tvwArray(Index).SelectedItem Is Nothing Then
       picChoose.Visible = False
       sText = tvwArray(Index).SelectedItem.Text
       sKey = tvwArray(Index).SelectedItem.Key
       RaiseEvent BarItemClick(lblArray(Index).Caption, lblArray(Index).Tag, sText, sKey)
    End If
End Sub

Private Sub tvwArray_DblClick(Index As Integer)
On Error Resume Next
    Dim sText As String
    Dim sKey As String
    If Not tvwArray(Index).SelectedItem Is Nothing Then
       picChoose.Visible = False
       sText = tvwArray(Index).SelectedItem.Text
       sKey = tvwArray(Index).SelectedItem.Key
       RaiseEvent BarItemDblClick(lblArray(Index).Caption, lblArray(Index).Tag, sText, sKey)
    End If
End Sub

Private Sub tvwArray_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub tvwArray_KeyPress(Index As Integer, KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub tvwArray_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim sText As String
    Dim sKey As String

    RaiseEvent KeyUp(KeyCode, Shift)

    If (KeyCode = 38 Or KeyCode = 40) And Shift = 0 Then
       If Not tvwArray(Index).SelectedItem Is Nothing Then
          sText = tvwArray(Index).SelectedItem.Text
          sKey = tvwArray(Index).SelectedItem.Key
          RaiseEvent BarItemClick(lblArray(Index).Caption, lblArray(Index).Tag, sText, sKey)
       End If
    End If
End Sub

Private Sub tvwArray_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If Not tvwArray(Index).HitTest(x, y) Is Nothing Then
       RaiseEvent MouseDown(Button, Shift, x, y)
    End If
End Sub

Private Sub tvwArray_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub tvwArray_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub tvwArray_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
    RaiseEvent NodeClick(Node)
End Sub

Private Sub tvwArray_OLECompleteDrag(Index As Integer, Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub tvwArray_OLEDragDrop(Index As Integer, Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub tvwArray_OLEDragOver(Index As Integer, Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub tvwArray_OLEGiveFeedback(Index As Integer, Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub tvwArray_OLESetData(Index As Integer, Data As MSComctlLib.DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub


Private Sub tvwArray_OLEStartDrag(Index As Integer, Data As MSComctlLib.DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub updCategory_DownClick()
    CurBar = m_CurBar + 1
    If picChoose.Visible Then
       picChoose.Visible = False
    End If
End Sub


Private Sub updCategory_UpClick()
    CurBar = m_CurBar - 1
End Sub


Private Sub UserControl_Initialize()
    UserControl_InitProperties
End Sub

Private Sub UserControl_Resize()
    ResizeControls
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get BackStyle() As Long
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Long)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Long
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Long)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
    UserControl.Refresh
End Sub

Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

Public Property Get LargeListImages() As IImages
Attribute LargeListImages.VB_Description = "Returns a reference to a collection of ListImage objects in an ImageList control."
    Set LargeListImages = imlSmallIcons.ListImages
End Property

Public Property Get SmallListImages() As IImages
    Set SmallListImages = imlSmallIcons.ListImages
End Property

Public Property Set LargeListImages(ByVal New_ListImages As IImages)
    Set imlSmallIcons.ListImages = New_ListImages
    PropertyChanged "LargeListImages"
End Property

Public Property Set SmallListImages(ByVal New_ListImages As IImages)
    Set imlSmallIcons.ListImages = New_ListImages
    PropertyChanged "SmallListImages"
End Property

Public Sub PopupMenu(Menu As Object, Optional Flags As Variant, Optional x As Variant, Optional y As Variant, Optional DefaultMenu As Variant)
Attribute PopupMenu.VB_Description = "Displays a pop-up menu on an MDIForm or Form object."
    UserControl.PopupMenu Menu, Flags, x, y, DefaultMenu
End Sub

Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Public Function AddBar(Optional ByVal BarName As String, Optional ByVal BarKey As String, Optional ByVal bBarList As Boolean = True) As Object
Attribute AddBar.VB_Description = "Add a ListBar to the control."
    m_BarCount = m_BarCount + 1
    Load lblArray(m_BarCount)
    Load lvwArray(m_BarCount)
    If Not bBarList Then
       Load tvwArray(m_BarCount)
    End If

    If Len(BarName) = 0 And Len(BarKey) = 0 Then
       BarName = "Bar " & m_BarCount
       BarKey = BarName
    ElseIf Len(BarName) = 0 Then
       BarName = BarKey
    ElseIf Len(BarKey) = 0 Then
       BarKey = BarName
    End If
    
    With lblArray(m_BarCount)
         .Caption = BarName
         .Tag = BarKey
         .Height = m_CaptionHeight
         .Width = lblArray(0).Width
         Set .Font = UserControl.Font
         .Visible = True
    End With

    With lvwArray(m_BarCount)
         .BackColor = lvwArray(0).BackColor
         .ForeColor = lvwArray(0).ForeColor
         .Arrange = lvwArray(0).Arrange
         .LabelEdit = lvwArray(0).LabelEdit
         .LabelWrap = lvwArray(0).LabelWrap
         .MultiSelect = lvwArray(0).MultiSelect
         .OLEDragMode = lvwArray(0).OLEDragMode
         .OLEDropMode = lvwArray(0).OLEDropMode
         .Sorted = lvwArray(0).Sorted
         .SortOrder = lvwArray(0).SortOrder
         .ToolTipText = lvwArray(0).ToolTipText
         .View = 3 'lvwArray(0).View
         .WhatsThisHelpID = lvwArray(0).WhatsThisHelpID
         Set .Font = UserControl.Font
         .Tag = IIf(bBarList, "List", "Tree")
    End With

On Error Resume Next

    If bBarList Then
       Set AddBar = lvwArray(m_BarCount)
    Else
       tvwArray(m_BarCount).LabelEdit = tvwArray(0).LabelEdit
       tvwArray(m_BarCount).OLEDragMode = tvwArray(0).OLEDragMode
       tvwArray(m_BarCount).OLEDropMode = tvwArray(0).OLEDropMode
       tvwArray(m_BarCount).Sorted = tvwArray(0).Sorted
       tvwArray(m_BarCount).ToolTipText = tvwArray(0).ToolTipText
       tvwArray(m_BarCount).WhatsThisHelpID = tvwArray(0).WhatsThisHelpID
       Set tvwArray(m_BarCount).Font = UserControl.Font

       Set AddBar = tvwArray(m_BarCount)
    End If
End Function

Public Sub AddBarItem(ByVal BarItemName As String, Optional ByVal BarItemKey As String, Optional BarItemIcon As String)
Attribute AddBarItem.VB_Description = "Add an Item on the current Bar"
On Error Resume Next
    If Len(BarItemKey) And Len(BarItemName) = 0 Then
       Err.Raise vbObjectError + 0, "FSListBar_AddBarItem", "No BarItemKey or BarItemName passed. At least one required."
    ElseIf Len(BarItemKey) = 0 Then
       BarItemKey = BarItemName
    ElseIf Len(BarItemName) = 0 Then
       BarItemName = BarItemKey
    End If

    If Len(BarItemIcon) = 0 Then
       lvwArray(m_CurBar).ListItems.Add , BarItemKey, BarItemName, "Document", "Document"
    ElseIf Len(imlSmallIcons.ListImages(BarItemIcon).Key) = 0 Then
       lvwArray(m_CurBar).ListItems.Add , BarItemKey, BarItemName, "Document", "Document"
    ElseIf Len(imlSmallIcons.ListImages(BarItemIcon).Key) = 0 Then
       lvwArray(m_CurBar).ListItems.Add , BarItemKey, BarItemName, "Document", "Document"
    Else
       lvwArray(m_CurBar).ListItems.Add , BarItemKey, BarItemName, BarItemIcon, BarItemIcon
    End If
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next
    m_CurBar = m_def_CurBar
    m_CurBarItem = m_def_CurBarItem
    m_CaptionHeight = m_def_CaptionHeight

    With lvwArray(0)
         .ColumnHeaders.Add , "Name"
         .View = 3
         .Tag = "List"
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    UserControl_InitProperties

   'm_CaptionHeight = PropBag.ReadProperty("CaptionHeight", m_def_CaptionHeight)

    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)

    With UserControl
         .Enabled = PropBag.ReadProperty("Enabled", True)
         .BackStyle = PropBag.ReadProperty("BackStyle", 1)
         .BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
         .AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    End With
    
    With lvwArray(0)
         .BackColor = PropBag.ReadProperty("BackColor", &H80000005)
         .ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
         .Arrange = PropBag.ReadProperty("Arrange", 0)
         .LabelEdit = PropBag.ReadProperty("LabelEdit", 1)
         .LabelWrap = PropBag.ReadProperty("LabelWrap", True)
         .MultiSelect = PropBag.ReadProperty("MultiSelect", False)
         .OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
         .OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
         .Sorted = PropBag.ReadProperty("Sorted", False)
        '.SortKey = PropBag.ReadProperty("SortKey", 0)
         .SortOrder = PropBag.ReadProperty("SortOrder", 0)
         .ToolTipText = PropBag.ReadProperty("ToolTipText", "")
         .View = 3 'PropBag.ReadProperty("View", 0)
         .WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
    End With

    tvwArray(0).LabelEdit = PropBag.ReadProperty("LabelEdit", 1)
    tvwArray(0).OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    tvwArray(0).OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    tvwArray(0).Sorted = PropBag.ReadProperty("Sorted", False)
    tvwArray(0).ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    tvwArray(0).WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
    Dim nCurrBar As Long
    For nCurrBar = m_BarCount To 1 Step -1
        Unload lblArray(nCurrBar)
        Unload lvwArray(nCurrBar)
        Unload tvwArray(nCurrBar)
    Next nCurrBar
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CaptionHeight", m_CaptionHeight, m_def_CaptionHeight)

    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)

    With UserControl
         Call PropBag.WriteProperty("Enabled", .Enabled, True)
         Call PropBag.WriteProperty("BackStyle", .BackStyle, 1)
         Call PropBag.WriteProperty("BorderStyle", .BorderStyle, 1)
         Call PropBag.WriteProperty("AutoRedraw", .AutoRedraw, False)
    End With

    With lvwArray(0)
         Call PropBag.WriteProperty("BackColor", .BackColor, &H80000005)
         Call PropBag.WriteProperty("ForeColor", .ForeColor, &H80000008)
         Call PropBag.WriteProperty("Arrange", .Arrange, 0)
         Call PropBag.WriteProperty("LabelEdit", .LabelEdit, 0)
         Call PropBag.WriteProperty("LabelWrap", .LabelWrap, True)
         Call PropBag.WriteProperty("MultiSelect", .MultiSelect, False)
         Call PropBag.WriteProperty("OLEDragMode", .OLEDragMode, 0)
         Call PropBag.WriteProperty("OLEDropMode", .OLEDropMode, 0)
         Call PropBag.WriteProperty("Sorted", .Sorted, False)
         Call PropBag.WriteProperty("SortOrder", .SortOrder, 0)
         Call PropBag.WriteProperty("ToolTipText", .ToolTipText, "")
         Call PropBag.WriteProperty("View", .View, 0)
         Call PropBag.WriteProperty("WhatsThisHelpID", .WhatsThisHelpID, 0)
    End With

End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = lvwArray(m_CurBar).BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Dim CurrBar As Long
On Error Resume Next
    For CurrBar = 0 To lvwArray.UBound
        lvwArray(CurrBar).BackColor = New_BackColor
       'tvwArray(CurrBar).BackColor = New_BackColor
    Next CurrBar
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    ForeColor = lvwArray(m_CurBar).ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Dim CurrBar As Long
On Error Resume Next
    For CurrBar = 0 To lvwArray.UBound
        lvwArray(CurrBar).ForeColor = New_ForeColor
       'tvwArray(CurrBar).ForeColor = New_ForeColor
    Next CurrBar
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lvwArray(m_CurBar).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Static nCurBar As Long
On Error Resume Next
    For nCurBar = 0 To m_BarCount
        Set lblArray(nCurBar).Font = New_Font
        Set lvwArray(nCurBar).Font = New_Font
        Set tvwArray(nCurBar).Font = New_Font
    Next nCurBar
    PropertyChanged "Font"
End Property

Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this control can act as an OLE drop target."
    OLEDropMode = lvwArray(m_CurBar).OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    lvwArray(m_CurBar).OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
On Error Resume Next
    If lvwArray(m_CurBar).Tag = "List" Then
       lvwArray(m_CurBar).OLEDrag
    Else
       tvwArray(m_CurBar).OLEDrag
    End If
End Sub

Private Sub lvwArray_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub lvwArray_KeyPress(Index As Integer, KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub lvwArray_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
On Error Resume Next
    If (KeyCode = 38 Or KeyCode = 40) And Shift = 0 Then
       If Not lvwArray(Index).SelectedItem Is Nothing Then
          With lvwArray(Index).SelectedItem
               RaiseEvent BarItemClick(lblArray(Index).Caption, lblArray(Index).Tag, .Text, .Key)
          End With
       End If
    End If
End Sub

Private Sub lvwArray_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If Not lvwArray(Index).HitTest(x, y) Is Nothing Then
       RaiseEvent MouseDown(Button, Shift, x, y)
    End If
End Sub

Private Sub lvwArray_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub lvwArray_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub lvwArray_OLEGiveFeedback(Index As Integer, Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub lvwArray_OLESetData(Index As Integer, Data As MSComctlLib.DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub lvwArray_OLECompleteDrag(Index As Integer, Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Public Property Get Arrange() As ListArrangeConstants
Attribute Arrange.VB_Description = "Returns/sets how the icons in a ListView control's Icon or SmallIcon view are arranged."
    Arrange = lvwArray(m_CurBar).Arrange
End Property

Public Property Let Arrange(ByVal New_Arrange As ListArrangeConstants)
    lvwArray(m_CurBar).Arrange() = New_Arrange
    PropertyChanged "Arrange"
End Property

Public Property Get Icons() As Object
Attribute Icons.VB_Description = "Returns/sets the images associated with the Icon properties of a ListView control."
    Set Icons = lvwArray(m_CurBar).Icons
End Property

Public Property Set Icons(ByVal New_Icons As Object)
    Set lvwArray(m_CurBar).Icons = New_Icons
    PropertyChanged "Icons"
End Property

Public Property Get ListItems() As IListItems
    Set ListItems = lvwArray(m_CurBar).ListItems
End Property

Public Property Get Nodes() As INodes
Attribute Nodes.VB_Description = "Returns a reference to a collection of ListItem objects in a ListView control."
On Error Resume Next
    Set Nodes = tvwArray(m_CurBar).Nodes
End Property

Public Property Get BarType() As String
    BarType = lvwArray(m_CurBar).Tag
End Property

Public Property Let BarType(NewType As String)
    If NewType = "List" Or NewType = "Tree" Then
       lvwArray(m_CurBar).Tag = NewType
    End If
End Property

Public Property Get LabelEdit() As ListLabelEditConstants
Attribute LabelEdit.VB_Description = "Returns/sets a value that determines if a user can edit the label of a ListItem or Node object."
On Error Resume Next
    If lvwArray(m_CurBar).Tag = "List" Then
       LabelEdit = lvwArray(m_CurBar).LabelEdit
    Else
       LabelEdit = tvwArray(m_CurBar).LabelEdit
    End If
End Property

Public Property Let LabelEdit(ByVal New_LabelEdit As ListLabelEditConstants)
On Error Resume Next
    If lvwArray(m_CurBar).Tag = "List" Then
       lvwArray(m_CurBar).LabelEdit() = New_LabelEdit
    Else
       tvwArray(m_CurBar).LabelEdit() = New_LabelEdit
    End If
    PropertyChanged "LabelEdit"
End Property

Public Property Get LabelWrap() As Boolean
Attribute LabelWrap.VB_Description = "Returns or sets a value that determines if labels are wrapped when the ListView is in Icon view."
    LabelWrap = lvwArray(m_CurBar).LabelWrap
End Property

Public Property Let LabelWrap(ByVal New_LabelWrap As Boolean)
    lvwArray(m_CurBar).LabelWrap() = New_LabelWrap
    PropertyChanged "LabelWrap"
End Property

Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Returns/sets a value indicating whether a user can make multiple selections in the ListView control and how the multiple selections can be made."
    MultiSelect = lvwArray(m_CurBar).MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
On Error Resume Next
    lvwArray(m_CurBar).MultiSelect() = New_MultiSelect
    PropertyChanged "MultiSelect"
End Property

Public Property Get SelectedNode() As INode
On Error Resume Next
    Set SelectedNode = tvwArray(m_CurBar).SelectedItem
End Property

Public Property Set SelectedNode(ByVal New_SelectedItem As INode)
On Error Resume Next
    Set tvwArray(m_CurBar).SelectedItem = New_SelectedItem
    PropertyChanged "SelectedItem"
End Property

Public Property Get SelectedItem() As IListItem
Attribute SelectedItem.VB_Description = "Returns a reference to the currently selected ListItem or Node object."
    Set SelectedItem = lvwArray(m_CurBar).SelectedItem
End Property

Public Property Set SelectedItem(ByVal New_SelectedItem As IListItem)
On Error Resume Next
    Set lvwArray(m_CurBar).SelectedItem = New_SelectedItem
    PropertyChanged "SelectedItem"
End Property

Public Property Get SmallIcons() As Object
Attribute SmallIcons.VB_Description = "Returns/sets the images associated with the SmallIcons property of a ListView control."
    Set SmallIcons = lvwArray(m_CurBar).SmallIcons
End Property

Public Property Set SmallIcons(ByVal New_SmallIcons As Object)
On Error Resume Next
    Set lvwArray(m_CurBar).SmallIcons = New_SmallIcons
    Set tvwArray(m_CurBar).ImageList = New_SmallIcons
    PropertyChanged "SmallIcons"
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = lvwArray(m_CurBar).Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
    lvwArray(m_CurBar).Sorted() = New_Sorted
    PropertyChanged "Sorted"
End Property

Public Property Get SortKey() As Long
Attribute SortKey.VB_Description = "Returns/sets the current sort key."
    SortKey = lvwArray(m_CurBar).SortKey
End Property

Public Property Let SortKey(ByVal New_SortKey As Long)
    lvwArray(m_CurBar).SortKey() = New_SortKey
    PropertyChanged "SortKey"
End Property

Public Property Get SortOrder() As ListSortOrderConstants
Attribute SortOrder.VB_Description = "Returns/sets whether or not the ListItems will be sorted in ascending or descending order."
    SortOrder = lvwArray(m_CurBar).SortOrder
End Property

Public Property Let SortOrder(ByVal New_SortOrder As ListSortOrderConstants)
    lvwArray(m_CurBar).SortOrder() = New_SortOrder
    PropertyChanged "SortOrder"
End Property

Public Property Get View() As ListViewConstants
Attribute View.VB_Description = "Returns/sets the current view of the ListView control."
    View = lvwArray(m_CurBar).View
End Property

Public Property Let View(ByVal New_View As ListViewConstants)
    lvwArray(m_CurBar).View() = New_View
    PropertyChanged "View"
End Property

Public Property Get OLEDragMode() As OLEDragConstants
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this control can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
    OLEDragMode = lvwArray(m_CurBar).OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As OLEDragConstants)
    lvwArray(m_CurBar).OLEDragMode() = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

Public Function FindItem(sz As String, Optional Where As Variant, Optional Index As Variant, Optional fPartial As Variant) As IListItem
Attribute FindItem.VB_Description = "Finds an item in the list and returns a reference to that item."
    FindItem = lvwArray(m_CurBar).FindItem(sz, Where, Index, fPartial)
End Function

Public Function TreeHitTest(x As Single, y As Single) As INode
On Error Resume Next
    TreeHitTest = tvwArray(m_CurBar).HitTest(x, y)
End Function

Public Function ListHitTest(x As Single, y As Single) As IListItem
Attribute ListHitTest.VB_Description = "Returns a reference to the ListItem object or Node object located at the coordinates of x and y. Used with drag and drop operations."
    ListHitTest = lvwArray(m_CurBar).HitTest(x, y)
End Function

Public Sub StartLabelEdit()
Attribute StartLabelEdit.VB_Description = "Begins a label editing operation on a ListItem or Node object."
On Error Resume Next
    If lvwArray(m_CurBar).Tag = "List" Then
       lvwArray(m_CurBar).StartLabelEdit
    Else
       tvwArray(m_CurBar).StartLabelEdit
    End If
End Sub

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
    WhatsThisHelpID = lvwArray(m_CurBar).WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
On Error Resume Next
    lvwArray(m_CurBar).WhatsThisHelpID() = New_WhatsThisHelpID
    tvwArray(m_CurBar).WhatsThisHelpID() = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = lvwArray(m_CurBar).ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
On Error Resume Next
    lvwArray(m_CurBar).ToolTipText() = New_ToolTipText
    tvwArray(m_CurBar).ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Private Sub lvwArray_BeforeLabelEdit(Index As Integer, Cancel As Integer)
    RaiseEvent BeforeLabelEdit(Cancel)
End Sub

Private Sub lvwArray_AfterLabelEdit(Index As Integer, Cancel As Integer, NewString As String)
    picChoose.Visible = False
    RaiseEvent AfterLabelEdit(Cancel, NewString)
End Sub

Public Property Get Value() As Variant
Attribute Value.VB_Description = "The value associated with the current Item on the current Bar"
    If Ambient.UserMode Then Err.Raise 393
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Variant)
    If Ambient.UserMode Then Err.Raise 393
    m_Value = New_Value
    PropertyChanged "Value"
End Property

Public Property Get BarName() As String
Attribute BarName.VB_Description = "The name of the current Bar"
Attribute BarName.VB_MemberFlags = "400"
    BarName = lblArray(m_CurBar).Caption
End Property

Public Property Let BarName(ByVal New_BarName As String)
    lblArray(m_CurBar).Caption = New_BarName
    PropertyChanged "BarName"
End Property

Public Property Get BarItemName() As String
Attribute BarItemName.VB_Description = "The name of the current Item in the current Bar"
Attribute BarItemName.VB_MemberFlags = "400"
On Error Resume Next
    If lvwArray(m_CurBar).Tag = "List" Then
       If Not lvwArray(m_CurBar).SelectedItem Is Nothing Then
          BarItemName = lvwArray(m_CurBar).SelectedItem.Text
       End If
    Else
       If Not tvwArray(m_CurBar).SelectedItem Is Nothing Then
          BarItemName = tvwArray(m_CurBar).SelectedItem.Text
       End If
    End If
End Property

Public Property Let BarItemName(ByVal New_BarItemName As String)
On Error Resume Next
    If lvwArray(m_CurBar).Tag = "List" Then
       If Not lvwArray(m_CurBar).SelectedItem Is Nothing Then
          lvwArray(m_CurBar).SelectedItem.Text = BarItemName
       End If
    Else
       If Not tvwArray(m_CurBar).SelectedItem Is Nothing Then
          tvwArray(m_CurBar).SelectedItem.Text = BarItemName
       End If
    End If

    PropertyChanged "BarItemName"
End Property

'Public Sub RemoveBar(Index As Variant)
'
'End Sub
'
'Public Sub RemoveBarItem(Index As Variant)
'
'End Sub

Public Sub SwitchToBar(Index As Long)
On Error GoTo EH_SwitchToBar
    RaiseEvent BeforeBarClick

    m_CurBar = lstChoose.ItemData(Index)
    UserControl_Resize

    UserControl.SetFocus
    If lvwArray(m_CurBar).Tag = "List" Then
       lvwArray(m_CurBar).ZOrder 0
       lvwArray(m_CurBar).Visible = True
       lvwArray(m_CurBar).SetFocus
    Else
On Error Resume Next
       tvwArray(m_CurBar).ZOrder 0
       tvwArray(m_CurBar).Visible = True
       tvwArray(m_CurBar).SetFocus
    End If
    
    RaiseEvent AfterBarClick

EH_SwitchToBar_Continue:
    Exit Sub

EH_SwitchToBar:
    MsgBox "Error" & Chr(13) & Chr(13) & Chr(9) & Err.Description
    Resume EH_SwitchToBar_Continue
    
    Resume
End Sub

Private Sub lvwArray_DblClick(Index As Integer)
On Error Resume Next
    If Not lvwArray(Index).SelectedItem Is Nothing Then
       picChoose.Visible = False
       With lvwArray(Index).SelectedItem
            RaiseEvent BarItemDblClick(lblArray(Index).Caption, lblArray(Index).Tag, .Text, .Key)
       End With
    End If
End Sub

Private Sub lvwArray_Click(Index As Integer)
On Error Resume Next
    If Not lvwArray(Index).SelectedItem Is Nothing Then
       picChoose.Visible = False
       With lvwArray(Index).SelectedItem
            RaiseEvent BarItemClick(lblArray(Index).Caption, lblArray(Index).Tag, .Text, .Key)
       End With
    End If
End Sub

Public Property Get CurBar() As Variant
Attribute CurBar.VB_Description = "The index of the current bar"
Attribute CurBar.VB_MemberFlags = "400"
    If Ambient.UserMode Then Err.Raise 393
    CurBar = m_CurBar
End Property

Public Property Let CurBar(ByVal New_CurBar As Variant)
    Static nCurBar As Long

    If VarType(New_CurBar) = vbString Then
       For nCurBar = 0 To m_BarCount
           If lblArray(nCurBar).Caption = New_CurBar Or lblArray(nCurBar).Tag = New_CurBar Then
              m_CurBar = nCurBar
              UserControl_Resize
              PropertyChanged "CurBar"
              Exit Property
           End If
       Next nCurBar
    ElseIf VarType(New_CurBar) = vbLong Or VarType(New_CurBar) = vbInteger Then
       If New_CurBar >= 0 And New_CurBar <= m_BarCount Then
          m_CurBar = New_CurBar
          UserControl_Resize
          PropertyChanged "CurBar"
          Exit Property
       End If
    End If
End Property

Public Property Get CurBarItem() As Variant
Attribute CurBarItem.VB_Description = "The index of the current bar Item"
Attribute CurBarItem.VB_MemberFlags = "400"
    CurBarItem = m_CurBarItem
End Property

Public Property Let CurBarItem(ByVal New_CurBarItem As Variant)
    Static nCurBarItem As Long
    Static CurItem As ListItem
    Static CurNode As Node

On Error Resume Next

    If lvwArray(m_CurBar).Tag = "List" Then
       If VarType(New_CurBarItem) = vbString Then
          For Each CurItem In lvwArray(m_CurBar).ListItems
              If CurItem.Key = New_CurBarItem Or CurItem.Text = New_CurBarItem Then
                 m_CurBarItem = CurItem.Index
                 CurItem.Selected = True
                 Set CurItem = Nothing
                 PropertyChanged "CurBarItem"
                 Exit Property
              End If
          Next CurItem
       ElseIf VarType(New_CurBarItem) = vbLong Or VarType(New_CurBarItem) = vbInteger Then
          If New_CurBarItem >= 0 And New_CurBarItem <= lvwArray(m_CurBar).ListItems.Count - 1 Then
             If Not lvwArray(m_CurBar).ListItems(New_CurBarItem) Is Nothing Then
                m_CurBarItem = New_CurBarItem
                PropertyChanged "CurBarItem"
                Exit Property
             End If
          End If
       End If
    Else
On Error Resume Next
       If VarType(New_CurBarItem) = vbString Then
          For Each CurNode In tvwArray(m_CurBar).Nodes
              If CurNode.Key = New_CurBarItem Or CurNode.Text = New_CurBarItem Then
                 m_CurBarItem = CurNode.Index
                 CurNode.Selected = True
                 Set CurNode = Nothing
                 PropertyChanged "CurBarItem"
                 Exit Property
              End If
          Next CurNode
       ElseIf VarType(New_CurBarItem) = vbLong Or VarType(New_CurBarItem) = vbInteger Then
          If New_CurBarItem >= 0 And New_CurBarItem <= tvwArray(m_CurBar).Nodes.Count - 1 Then
             If Not tvwArray(m_CurBar).Nodes(New_CurBarItem) Is Nothing Then
                m_CurBarItem = New_CurBarItem
                PropertyChanged "CurBarItem"
                Exit Property
             End If
          End If
       End If
    End If
End Property

Public Property Get DropDownToolTipText() As String
    DropDownToolTipText = cmdChangeCategory.ToolTipText
End Property

Public Property Let DropDownToolTipText(ByVal sNewValue As String)
    cmdChangeCategory.ToolTipText = sNewValue
End Property

Public Property Get DropDownListBox() As Object
    Set DropDownListBox = lstChoose
End Property

