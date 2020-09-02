VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl DataView 
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4800
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2115
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   3731
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "DataView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Firm Solutions DataView"
Option Explicit

'Default Property Values:
Const m_def_DatabaseName = ""
Const m_def_ColumnNames = ""
Const m_def_ColumnWidths = ""
Const m_def_RecordSource = ""

'Property Variables:
Dim m_DatabaseName As String
Dim m_ColumnNames As String
Dim m_ColumnWidths As String
Dim m_RecordSource As String

Private rmxData As CAssocArray

'Event Declarations:
Event Click() 'MappingInfo=lvwMain,lvwMain,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=lvwMain,lvwMain,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=lvwMain,lvwMain,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=lvwMain,lvwMain,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=lvwMain,lvwMain,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvwMain,lvwMain,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvwMain,lvwMain,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvwMain,lvwMain,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event WriteProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,WriteProperties
Attribute WriteProperties.VB_Description = "Occurs when a user control or user document is asked to write its data to a file."
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show
Attribute Show.VB_Description = "Occurs when the control's Visible property changes to True."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event ReadProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,ReadProperties
Attribute ReadProperties.VB_Description = "Occurs when a user control or user document is asked to read its data from a file."
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Occurs when any part of a form or PictureBox control is moved, enlarged, or exposed."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=lvwMain,lvwMain,-1,OLEStartDrag
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=lvwMain,lvwMain,-1,OLESetData
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=lvwMain,lvwMain,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "OLEGiveFeedback event"
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=lvwMain,lvwMain,-1,OLEDragOver
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvwMain,lvwMain,-1,OLEDragDrop
Event OLECompleteDrag(Effect As Long) 'MappingInfo=lvwMain,lvwMain,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "OLECompleteDrag event"
Event ItemClick(Item As ListItem) 'MappingInfo=lvwMain,lvwMain,-1,ItemClick
Attribute ItemClick.VB_Description = "Occurs when a ListItem object is clicked or selected"
Event ItemCheck(Item As ListItem) 'MappingInfo=lvwMain,lvwMain,-1,ItemCheck
Attribute ItemCheck.VB_Description = "Occurs when a ListSubItem object is checked"
Event InitProperties() 'MappingInfo=UserControl,UserControl,-1,InitProperties
Attribute InitProperties.VB_Description = "Occurs the first time a user control or user document is created."
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Attribute Hide.VB_Description = "Occurs when the control's Visible property changes to False."
Event ColumnClick(ColumnHeader As ColumnHeader) 'MappingInfo=lvwMain,lvwMain,-1,ColumnClick
Attribute ColumnClick.VB_Description = "Occurs when a ColumnHeader object in a ListView control is clicked."
Event BeforeLabelEdit(Cancel As Integer) 'MappingInfo=lvwMain,lvwMain,-1,BeforeLabelEdit
Attribute BeforeLabelEdit.VB_Description = "Occurs when a user attempts to edit the label of the currently selected ListItem or Node object."
Event AsyncReadComplete(AsyncProp As AsyncProperty) 'MappingInfo=UserControl,UserControl,-1,AsyncReadComplete
Attribute AsyncReadComplete.VB_Description = "Occurs when the data that is requested by the AsyncRead method is available."
Event AfterLabelEdit(Cancel As Integer, NewString As String) 'MappingInfo=lvwMain,lvwMain,-1,AfterLabelEdit
Attribute AfterLabelEdit.VB_Description = "Occurs after a user edits the label of the currently selected Node or ListItem object."

Public Property Get SelectedItem() As Object
    Set SelectedItem = lvwMain.SelectedItem
End Property

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    RaiseEvent ColumnClick(ColumnHeader)
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As ListItem)
    RaiseEvent ItemClick(Item)
End Sub

Private Sub UserControl_Initialize()
    Set rmxData = New CAssocArray
    ExtendListView lvwMain.hWnd
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = lvwMain.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lvwMain.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    ForeColor = lvwMain.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lvwMain.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lvwMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lvwMain.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
    lvwMain.Refresh
End Sub

Private Sub lvwMain_Click()
    RaiseEvent Click
End Sub

Private Sub lvwMain_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Terminate()
    Set rmxData = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    RaiseEvent WriteProperties(PropBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", lvwMain.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", lvwMain.BorderStyle, 1)
    Call PropBag.WriteProperty("WhatsThisHelpID", lvwMain.WhatsThisHelpID, 0)
    Call PropBag.WriteProperty("View", lvwMain.View, 0)
    Call PropBag.WriteProperty("ToolTipText", lvwMain.ToolTipText, "")
    Call PropBag.WriteProperty("TextBackground", lvwMain.TextBackground, 0)
    Call PropBag.WriteProperty("SortOrder", lvwMain.SortOrder, 0)
    Call PropBag.WriteProperty("SortKey", lvwMain.SortKey, 0)
    Call PropBag.WriteProperty("Sorted", lvwMain.Sorted, False)
    Call PropBag.WriteProperty("SmallIcons", SmallIcons, Nothing)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 4800)
    Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
    Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 3600)
    Call PropBag.WriteProperty("RightToLeft", UserControl.RightToLeft, False)
    Call PropBag.WriteProperty("PictureAlignment", lvwMain.PictureAlignment, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("PaletteMode", UserControl.PaletteMode, 3)
    Call PropBag.WriteProperty("Palette", Palette, Nothing)
    Call PropBag.WriteProperty("OLEDropMode", lvwMain.OLEDropMode, 0)
    Call PropBag.WriteProperty("OLEDragMode", lvwMain.OLEDragMode, 0)
    Call PropBag.WriteProperty("MultiSelect", lvwMain.MultiSelect, False)
    Call PropBag.WriteProperty("MousePointer", lvwMain.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MaskPicture", MaskPicture, Nothing)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, -2147483633)
    Call PropBag.WriteProperty("LabelWrap", lvwMain.LabelWrap, True)
    Call PropBag.WriteProperty("LabelEdit", lvwMain.LabelEdit, 0)
    
    Call PropBag.WriteProperty("HoverSelection", lvwMain.HoverSelection, False)
    Call PropBag.WriteProperty("HotTracking", lvwMain.HotTracking, False)
    Call PropBag.WriteProperty("HideSelection", lvwMain.HideSelection, True)
    Call PropBag.WriteProperty("HideColumnHeaders", lvwMain.HideColumnHeaders, False)
    Call PropBag.WriteProperty("GridLines", lvwMain.GridLines, False)
    Call PropBag.WriteProperty("FullRowSelect", lvwMain.FullRowSelect, False)
    Call PropBag.WriteProperty("FontUnderline", UserControl.FontUnderline, 0)
    Call PropBag.WriteProperty("FontTransparent", UserControl.FontTransparent, True)
    Call PropBag.WriteProperty("FontStrikethru", UserControl.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 0)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, "")
    Call PropBag.WriteProperty("FontItalic", UserControl.FontItalic, 0)
    Call PropBag.WriteProperty("FontBold", UserControl.FontBold, 0)
    Call PropBag.WriteProperty("FlatScrollBar", lvwMain.FlatScrollBar, False)
    Call PropBag.WriteProperty("FillStyle", UserControl.FillStyle, 1)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("DrawWidth", UserControl.DrawWidth, 1)
    Call PropBag.WriteProperty("DrawStyle", UserControl.DrawStyle, 0)
    Call PropBag.WriteProperty("DrawMode", UserControl.DrawMode, 13)
    Call PropBag.WriteProperty("CurrentY", UserControl.CurrentY, 0)
    Call PropBag.WriteProperty("CurrentX", UserControl.CurrentX, 0)
    
    Call PropBag.WriteProperty("Checkboxes", lvwMain.Checkboxes, False)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("Arrange", lvwMain.Arrange, 0)
    Call PropBag.WriteProperty("Appearance", lvwMain.Appearance, 1)
    Call PropBag.WriteProperty("AllowColumnReorder", lvwMain.AllowColumnReorder, False)
    Call PropBag.WriteProperty("DatabaseName", m_DatabaseName, m_def_DatabaseName)
    Call PropBag.WriteProperty("ColumnNames", m_ColumnNames, m_def_ColumnNames)
    Call PropBag.WriteProperty("ColumnWidths", m_ColumnWidths, m_def_ColumnWidths)
    Call PropBag.WriteProperty("RecordSource", m_RecordSource, m_def_RecordSource)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,View
Public Property Get View() As ListViewConstants
Attribute View.VB_Description = "Returns/sets the current view of the ListView control."
    View = lvwMain.View
End Property

Public Property Let View(ByVal New_View As ListViewConstants)
    lvwMain.View() = New_View
    PropertyChanged "View"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,SortOrder
Public Property Get SortOrder() As ListSortOrderConstants
Attribute SortOrder.VB_Description = "Returns/sets whether or not the ListItems will be sorted in ascending or descending order."
    SortOrder = lvwMain.SortOrder
End Property

Public Property Let SortOrder(ByVal New_SortOrder As ListSortOrderConstants)
    lvwMain.SortOrder() = New_SortOrder
    PropertyChanged "SortOrder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,SortKey
Public Property Get SortKey() As Integer
Attribute SortKey.VB_Description = "Returns/sets the current sort key."
    SortKey = lvwMain.SortKey
End Property

Public Property Let SortKey(ByVal New_SortKey As Integer)
    lvwMain.SortKey() = New_SortKey
    PropertyChanged "SortKey"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,Sorted
Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = lvwMain.Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
    lvwMain.Sorted() = New_Sorted
    PropertyChanged "Sorted"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,SmallIcons
Public Property Get SmallIcons() As Object
Attribute SmallIcons.VB_Description = "Returns/sets the images associated with the SmallIcons property of a ListView control."
    Set SmallIcons = lvwMain.SmallIcons
End Property

Public Property Set SmallIcons(ByVal New_SmallIcons As Object)
    Set lvwMain.SmallIcons = New_SmallIcons
    PropertyChanged "SmallIcons"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Size
Public Sub Size(Width As Single, Height As Single)
Attribute Size.VB_Description = "Changes the width and height of a User Control."
    UserControl.Size Width, Height
End Sub

Private Sub UserControl_Show()
    RaiseEvent Show
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
    ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    UserControl.ScaleTop() = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As Integer
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
    UserControl.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
    ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    UserControl.ScaleLeft() = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'The Underscore following "Scale" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Scale
Public Sub Scale_(Optional X1 As Variant, Optional Y1 As Variant, Optional X2 As Variant, Optional Y2 As Variant)
    UserControl.Scale (X1, Y1)-(X2, Y2)
End Sub

Private Sub UserControl_Resize()
    lvwMain.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    'RaiseEvent Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    RaiseEvent ReadProperties(PropBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    lvwMain.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    lvwMain.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    lvwMain.WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
    lvwMain.View = PropBag.ReadProperty("View", 0)
    lvwMain.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    lvwMain.TextBackground = PropBag.ReadProperty("TextBackground", 0)
    lvwMain.SortOrder = PropBag.ReadProperty("SortOrder", 0)
    lvwMain.SortKey = PropBag.ReadProperty("SortKey", 0)
    lvwMain.Sorted = PropBag.ReadProperty("Sorted", False)
    Set SmallIcons = PropBag.ReadProperty("SmallIcons", Nothing)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 4800)
    UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
    UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 3600)
    UserControl.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    lvwMain.PictureAlignment = PropBag.ReadProperty("PictureAlignment", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.PaletteMode = PropBag.ReadProperty("PaletteMode", 3)
    Set Palette = PropBag.ReadProperty("Palette", Nothing)
    lvwMain.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    lvwMain.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    lvwMain.MultiSelect = PropBag.ReadProperty("MultiSelect", False)
    lvwMain.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", -2147483633)
    lvwMain.LabelWrap = PropBag.ReadProperty("LabelWrap", True)
    lvwMain.LabelEdit = PropBag.ReadProperty("LabelEdit", 0)
    
    lvwMain.HoverSelection = PropBag.ReadProperty("HoverSelection", False)
    lvwMain.HotTracking = PropBag.ReadProperty("HotTracking", False)
    lvwMain.HideSelection = PropBag.ReadProperty("HideSelection", True)
    lvwMain.HideColumnHeaders = PropBag.ReadProperty("HideColumnHeaders", False)
    lvwMain.GridLines = PropBag.ReadProperty("GridLines", False)
    lvwMain.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
    UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    UserControl.FontTransparent = PropBag.ReadProperty("FontTransparent", True)
    UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    UserControl.FontSize = PropBag.ReadProperty("FontSize", 8.25)
    UserControl.FontName = PropBag.ReadProperty("FontName", "MS Sans Serif")
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    UserControl.FontBold = PropBag.ReadProperty("FontBold", 0)
    lvwMain.FlatScrollBar = PropBag.ReadProperty("FlatScrollBar", False)
    UserControl.FillStyle = PropBag.ReadProperty("FillStyle", 1)
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    UserControl.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
    UserControl.DrawStyle = PropBag.ReadProperty("DrawStyle", 0)
    UserControl.DrawMode = PropBag.ReadProperty("DrawMode", 13)
    UserControl.CurrentY = PropBag.ReadProperty("CurrentY", 0)
    UserControl.CurrentX = PropBag.ReadProperty("CurrentX", 0)
    
    lvwMain.Checkboxes = PropBag.ReadProperty("Checkboxes", False)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    lvwMain.Arrange = PropBag.ReadProperty("Arrange", 0)
    lvwMain.Appearance = PropBag.ReadProperty("Appearance", 1)
    lvwMain.AllowColumnReorder = PropBag.ReadProperty("AllowColumnReorder", False)
    m_DatabaseName = PropBag.ReadProperty("DatabaseName", m_def_DatabaseName)
    m_ColumnNames = PropBag.ReadProperty("ColumnNames", m_def_ColumnNames)
    m_ColumnWidths = PropBag.ReadProperty("ColumnWidths", m_def_ColumnWidths)
    m_RecordSource = PropBag.ReadProperty("RecordSource", m_def_RecordSource)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,MultiSelect
Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Returns/sets a value indicating whether a user can make multiple selections in the ListView control and how the multiple selections can be made."
    MultiSelect = lvwMain.MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
    lvwMain.MultiSelect() = New_MultiSelect
    PropertyChanged "MultiSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,LabelWrap
Public Property Get LabelWrap() As Boolean
Attribute LabelWrap.VB_Description = "Returns or sets a value that determines if labels are wrapped when the ListView is in Icon view."
    LabelWrap = lvwMain.LabelWrap
End Property

Public Property Let LabelWrap(ByVal New_LabelWrap As Boolean)
    lvwMain.LabelWrap() = New_LabelWrap
    PropertyChanged "LabelWrap"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,LabelEdit
Public Property Get LabelEdit() As ListLabelEditConstants
Attribute LabelEdit.VB_Description = "Returns/sets a value that determines if a user can edit the label of a ListItem or Node object."
    LabelEdit = lvwMain.LabelEdit
End Property

Public Property Let LabelEdit(ByVal New_LabelEdit As ListLabelEditConstants)
    lvwMain.LabelEdit() = New_LabelEdit
    PropertyChanged "LabelEdit"
End Property

Private Sub UserControl_InitProperties()
    RaiseEvent InitProperties
    m_DatabaseName = m_def_DatabaseName
    m_ColumnNames = m_def_ColumnNames
    m_ColumnWidths = m_def_ColumnWidths
    m_RecordSource = m_def_RecordSource
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,HoverSelection
Public Property Get HoverSelection() As Boolean
Attribute HoverSelection.VB_Description = "Returns/sets whether hover selection is enabled."
     HoverSelection = lvwMain.HoverSelection
End Property

Public Property Let HoverSelection(ByVal New_HoverSelection As Boolean)
    lvwMain.HoverSelection() = New_HoverSelection
    PropertyChanged "HoverSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,HotTracking
Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets whether hot tracking is enabled."
    HotTracking = lvwMain.HotTracking
End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
    lvwMain.HotTracking() = New_HotTracking
    PropertyChanged "HotTracking"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Determines whether the selected item will display as selected when the ListView loses focus"
    HideSelection = lvwMain.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    lvwMain.HideSelection() = New_HideSelection
    PropertyChanged "HideSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,HideColumnHeaders
Public Property Get HideColumnHeaders() As Boolean
Attribute HideColumnHeaders.VB_Description = "Returns/sets whether or not a ListView control's column headers are hidden in Report view."
    HideColumnHeaders = lvwMain.HideColumnHeaders
End Property

Public Property Let HideColumnHeaders(ByVal New_HideColumnHeaders As Boolean)
    lvwMain.HideColumnHeaders() = New_HideColumnHeaders
    PropertyChanged "HideColumnHeaders"
End Property

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,GridLines
Public Property Get GridLines() As Boolean
Attribute GridLines.VB_Description = "Returns/sets whether grid lines appear between rows and columns"
    GridLines = lvwMain.GridLines
End Property

Public Property Let GridLines(ByVal New_GridLines As Boolean)
    lvwMain.GridLines() = New_GridLines
    PropertyChanged "GridLines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,GetFirstVisible
Public Function GetFirstVisible() As IListItem
Attribute GetFirstVisible.VB_Description = "Retrieves a reference of the first item visible in the client area."
    GetFirstVisible = lvwMain.GetFirstVisible()
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,FullRowSelect
Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Returns/sets whether selecting a column highlights the entire row."
    FullRowSelect = lvwMain.FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
    lvwMain.FullRowSelect() = New_FullRowSelect
    PropertyChanged "FullRowSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,FindItem
Public Function FindItem(sz As String, Optional Where As Variant, Optional Index As Variant, Optional fPartial As Variant) As IListItem
Attribute FindItem.VB_Description = "Finds an item in the list and returns a reference to that item."
    FindItem = lvwMain.FindItem(sz, Where, Index, fPartial)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
Attribute Controls.VB_Description = "A collection whose elements represent each control on a form, including elements of control arrays. "
    Set Controls = UserControl.Controls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,ColumnHeaders
Public Property Get ColumnHeaders() As IColumnHeaders
Attribute ColumnHeaders.VB_Description = "Returns a reference to a collection of ColumnHeader objects."
    Set ColumnHeaders = lvwMain.ColumnHeaders
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,Arrange
Public Property Get Arrange() As ListArrangeConstants
Attribute Arrange.VB_Description = "Returns/sets how the icons in a ListView control's Icon or SmallIcon view are arranged."
    Arrange = lvwMain.Arrange
End Property

Public Property Let Arrange(ByVal New_Arrange As ListArrangeConstants)
    lvwMain.Arrange() = New_Arrange
    PropertyChanged "Arrange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvwMain,lvwMain,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not controls, Forms or an MDIForm are painted at run time with 3-D effects."
    Appearance = lvwMain.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    lvwMain.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

Public Property Get DatabaseName() As String
    DatabaseName = m_DatabaseName
End Property

Public Property Let DatabaseName(ByVal New_DatabaseName As String)
    m_DatabaseName = New_DatabaseName
    PropertyChanged "DatabaseName"
End Property

Public Sub Requery(Optional ByVal dbOutside As Object, Optional ByVal lLimitNumberOfRecordsTo As Long = 100)
Attribute Requery.VB_Description = "Takes DatabaseName, Recordsouce, ColumnNames, and ColumnWidths and (re)fills this control's contents."
On Error GoTo EH_DataView_Requery
    Dim db As Database
    Dim rst As Recordset
    Dim CurField As Long

    If Len(Me.DatabaseName) = 0 Or Len(Me.RecordSource) = 0 Then Exit Sub

    Screen.MousePointer = vbHourglass
    lvwMain.Visible = False
        If dbOutside Is Nothing Then
           Set db = OpenDatabase(DatabaseName, , True)
        Else
           Set db = dbOutside
           ColumnNames = ""
        End If
            Set rst = db.OpenRecordset(RecordSource)
                With rmxData
                     .Clear
                     If Len(ColumnNames) = 0 Then
                        For CurField = 0 To rst.Fields.Count - 1
                            ColumnNames = ColumnNames & rst.Fields(CurField).Name & .KeyValueDelimiter & "1500" & .ItemDelimiter
                        Next CurField
                     End If
                     .All = ColumnNames
                     .FillListViewColumns lvwMain
                End With
                FillListView rst, lvwMain, False, 20
            rst.Close
            Set rst = Nothing
        If dbOutside Is Nothing Then db.Close
        Set db = Nothing

EH_DataView_Requery_Continue:
    Screen.MousePointer = vbDefault
    lvwMain.Visible = True
    Exit Sub

EH_DataView_Requery:
    MsgBox "Error occured in:" & Chr(13) & Chr(9) & "Module: DataView" & Chr(13) & Chr(9) & "Procedure: Requery" & Chr(13) & Chr(13) & Err.Description
    Resume EH_DataView_Requery_Continue

    Resume
End Sub

Public Sub FillListView(rst As Recordset, lvwCtrl As Object, Optional bFullLine As Boolean = True, Optional ByVal lLimitNumberOfRecordsTo As Long = 0)
    Static lvwX As ListView
    Static NewItem As ListItem
    Static SubItems As Integer
    Static CurSubItem As Integer
    Static sT As String
    Static lCurrent As Long

On Error Resume Next
    If lvwCtrl Is Nothing Then
       Set lvwX = lvwCtrl
    Else
       Set lvwX = lvwMain
    End If
        lvwX.ListItems.Clear
        If bFullLine Then ExtendListView lvwX.hWnd
        lCurrent = 0
        Do Until rst.EOF Or (lCurrent > lLimitNumberOfRecordsTo And lLimitNumberOfRecordsTo > 0)
           Set NewItem = lvwX.ListItems.Add(, , "" & rst.Fields(0))
           For CurSubItem = 1 To rst.Fields.Count - 1
               NewItem.SubItems(CurSubItem) = "" & rst.Fields(CurSubItem)
           Next CurSubItem
           rst.MoveNext
           lCurrent = lCurrent + 1
        Loop
        Set NewItem = Nothing
    Set lvwX = Nothing
    sT = ""
End Sub

Public Property Get ColumnNames() As String
    ColumnNames = m_ColumnNames
End Property

Public Property Let ColumnNames(ByVal New_ColumnNames As String)
    m_ColumnNames = New_ColumnNames
    PropertyChanged "ColumnNames"
End Property

Public Property Get RecordSource() As String
    RecordSource = m_RecordSource
End Property

Public Property Let RecordSource(ByVal New_RecordSource As String)
    m_RecordSource = New_RecordSource
    PropertyChanged "RecordSource"
End Property

