VERSION 5.00
Begin VB.UserControl usrSplitter 
   BackColor       =   &H80000009&
   CanGetFocus     =   0   'False
   ClientHeight    =   264
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   216
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Splitter.ctx":0000
   ScaleHeight     =   264
   ScaleWidth      =   216
   ToolboxBitmap   =   "Splitter.ctx":056A
End
Attribute VB_Name = "usrSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
'   These variables are used for storing often used calculations
Private m_intTotalXPadding As Integer
Private m_intTotalYPadding As Integer
Private m_intTopPosition As Integer
Private m_intBorderSize As Long
'
'   This control is made to work with both a form or a PictureBox as the container,
'       therefore I cannot use the ScaleHeight and ScaleWidth Properties since they do not exist on a Picture
'       To reserve space for a ToolBar, StatusBar or other controls in a Container you
'       can set these two properties to the height of the controls which are taking
'       up a portion of the window space
Private m_intReservedSpaceTop As Integer
Private m_intReservedSpaceBottom As Integer
'
'   property vaiables
Private m_intTopPad As Integer
Private m_intLeftPad As Integer
Private m_intCenterPad As Integer
Private m_intRightPad As Integer
Private m_intBottomPad As Integer
Private m_intStyle As Style
Private m_intBarWidth As Integer
Private m_intMinHeight As Integer
Private m_intMinBottomHeight As Integer
Private m_Cursor As Cursor
'
'   Default values
Const m_cnstTopPadding = 100
Const m_cnstLeftPadding = 100
Const m_cnstRightPadding = 100
Const m_cnstCenterPadding = 50
Const m_cnstBottomPadding = 100
Const m_cnstStyle = 1
Const m_cnstBarWidth = 50
Const m_cnstMinHeight = 1000
Const m_cnstMinBottomHeight = 1000

Public Enum Style
    [frstVertical] = 1
    [frstHorizontal] = 2
    [frstBoth] = 4
End Enum
Public Enum Cursor
    [spltDefault] = 0
    [spltCustom] = 1
End Enum
'
'   This flag is used to know of the form is being initially sized
Private m_bInitialResize As Boolean

Private WithEvents m_picVSplitter As PictureBox
Attribute m_picVSplitter.VB_VarHelpID = -1
Private WithEvents m_imgVSplitter As Image
Attribute m_imgVSplitter.VB_VarHelpID = -1
Private WithEvents m_picHSplitter As PictureBox
Attribute m_picHSplitter.VB_VarHelpID = -1
Private WithEvents m_imgHSplitter As Image
Attribute m_imgHSplitter.VB_VarHelpID = -1

Private m_objLeftSide As Object
Private m_objRightSide As Object
Private m_objBottom As Object
Private m_objContainer As Object

Private m_bMoving As Boolean
Const m_sglSplitLimit = 300

'
'   API calls and required defs
Const SM_CXBORDER = 5
Const SM_CYBORDER = 6
Const SM_CXDLGFRAME = 7
Const SM_CYDLGFRAME = 8
Const SM_CXFRAME = 32
Const SM_CYFRAME = 33

Private Type RECT
     left As Long
     top As Long
     right As Long
     bottom As Long
End Type

Private Declare Function GetClientRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Private Declare Function GetWindowRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Private Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)

Private Function GetWindowHeight() As Long
    Dim lpRect As RECT
    Dim lRet As Long
    
    lRet = GetClientRect(ContainerHWnd, lpRect)
    GetWindowHeight = lpRect.bottom * Screen.TwipsPerPixelX
End Function

Private Sub GetWindowSize(ByRef lWindowWidth As Long, ByRef lWindowHeight As Long)
    Dim lpRect As RECT
    Dim lRet As Long
    
    lRet = GetClientRect(ContainerHWnd, lpRect)
    lWindowHeight = (lpRect.bottom - lpRect.top) * Screen.TwipsPerPixelX
    lWindowWidth = lpRect.right * Screen.TwipsPerPixelY
End Sub

Private Function GetWindowWidth() As Long
    Dim lpRect As RECT
    Dim lRet As Long
    
    lRet = GetClientRect(ContainerHWnd, lpRect)
    GetWindowWidth = lpRect.right * Screen.TwipsPerPixelY
End Function

Private Sub HorizontalSplitterResize()
    Dim intBarPos As Integer
    Dim lHeight As Long
    
    '
    '   Call a private routine to use windows API to get the window sizes
    lHeight = GetWindowHeight
    '
    '   Store the current top position of the splitter bar
    intBarPos = m_picHSplitter.top
    '
    '   The height of the top pane is the bar position minus the top poition
    '
    '   But don't let the user drop the bar ABOVE the top of the top pane
    If intBarPos <= m_intTopPosition Then
        '
        '   make the top window 0
        m_objLeftSide.Height = 0
        '
        '   and set the splitter at the top position
        intBarPos = m_intTopPosition
        m_imgHSplitter.top = intBarPos
        
    Else
        If intBarPos > lHeight - m_intTopPosition Then
            intBarPos = lHeight - m_intTopPosition
        End If
        m_objLeftSide.Height = intBarPos - m_intTopPosition
    End If
    '
    '   Verify that a right side pne exists and if it does, set it's height to the
    '       same as the left pane
    If Not m_objRightSide Is Nothing Then
        m_objRightSide.Height = m_objLeftSide.Height
    End If
    '
    '
    m_imgHSplitter.top = intBarPos
    '
    '   If there is no Bottom object, don't try to size it!
    If Not m_objBottom Is Nothing Then
        m_objBottom.top = intBarPos + m_intCenterPad
        If 0 > lHeight - m_objLeftSide.Height - m_intCenterPad - m_intTotalYPadding Then
            '
            '   make the bottom pane zero
            m_objBottom.Height = 0
        Else
            m_objBottom.Height = lHeight - m_objLeftSide.Height - m_intCenterPad - m_intTotalYPadding
        End If
    End If
    '
    '   If a vertical splitter is also defined then set it's new size to the height
    '       of the top pane.  If we don't do this then the Vertical splitter will
    '       extebd beyond the pane or window
    If m_intStyle And frstVertical Then
        m_imgVSplitter.Height = m_objLeftSide.Height
    End If
End Sub

Public Sub Initialize(objLeft As Object, objRight As Object, objBottom As Object, Optional objContainer As Object = Nothing)
    Set m_objLeftSide = objLeft
    Set m_objRightSide = objRight
    Set m_objBottom = objBottom
    Set m_objContainer = objContainer
    
    m_intTotalXPadding = m_intLeftPad + m_intRightPad
    m_intTotalYPadding = m_intReservedSpaceTop + m_intTopPad + m_intBottomPad + m_intReservedSpaceBottom
    m_intTopPosition = m_intReservedSpaceTop + m_intTopPad
    '
    '   When the objects are smallest they still have a size of 2 * border.
    '   So we calcualte this value and use it in the code to make the windows go away when needed
    m_intBorderSize = GetSystemMetrics(SM_CYBORDER) * Screen.TwipsPerPixelY * 2

    '
    '   This code is here to whatch for a mis match of Stryle flags
    '    the and actual objects passed to Initialize
    If m_objBottom Is Nothing Then
        '
        '   Remove the Horizontal flag from the style flag
        m_intStyle = Not (m_intStyle Imp 2)
    End If
    If m_objRightSide Is Nothing Then
        '
        '   Remove the Vertical flag from the style flag
        m_intStyle = Not (m_intStyle Imp 1)
    End If
    
    If m_intStyle And frstVertical Then
        '
        '   If an optinal Container was passed in, host the images in that container
        If m_objContainer Is Nothing Then
            Set m_imgVSplitter = Parent.Controls.Add("VB.Image", "m_imgVSplitter")
            Set m_picVSplitter = Parent.Controls.Add("VB.Picturebox", "m_picVSplitter")
        Else
            '
            '   Otherwise host them on the parent form
            Set m_imgVSplitter = Parent.Controls.Add("VB.Image", "m_imgVSplitter", m_objContainer)
            Set m_picVSplitter = Parent.Controls.Add("VB.Picturebox", "m_picVSplitter", m_objContainer)
        End If
        With m_imgVSplitter
            If m_Cursor = spltDefault Then
                .MousePointer = vbSizeWE
            Else
                .MousePointer = vbCustom
                .MouseIcon = LoadResPicture(102, vbResCursor)
            End If
            '
            '   Set up as a vertical splitter bar
            .top = m_objLeftSide.top
            .Height = m_objLeftSide.Height
            .Visible = True
            .left = m_objLeftSide.Width + m_intLeftPad
            .Width = m_intCenterPad
            .Appearance = 0
            .BorderStyle = 0
        End With
        
        With m_picVSplitter
            .Visible = False
            .DrawStyle = 0
            .BackColor = &H808080
            .Width = m_intBarWidth
        End With
    End If

    If m_intStyle And frstHorizontal Then
        If m_objContainer Is Nothing Then
            Set m_imgHSplitter = Parent.Controls.Add("VB.Image", "m_imgHSplitter")
            Set m_picHSplitter = Parent.Controls.Add("VB.Picturebox", "m_picHSplitter")
        Else
            Set m_imgHSplitter = Parent.Controls.Add("VB.Image", "m_imgHSplitter", m_objContainer)
            Set m_picHSplitter = Parent.Controls.Add("VB.Picturebox", "m_picHSplitter", m_objContainer)
        End If
        With m_imgHSplitter
            If m_Cursor = spltDefault Then
                .MousePointer = vbSizeNS
            Else
                .MousePointer = vbCustom
                .MouseIcon = LoadResPicture(101, vbResCursor)
            End If
            '
            '   This is for the horizontal spliter
            .top = m_objLeftSide.Height + m_intTopPosition
            .Height = m_intCenterPad
            .BorderStyle = 0
            .Visible = True
            .left = m_objLeftSide.left
            If Not m_objRightSide Is Nothing Then
                .Width = m_objLeftSide.Width + m_intCenterPad + m_objRightSide.Width
            Else
                .Width = m_objLeftSide.Width
            End If
        End With
        
        With m_picHSplitter
            .Visible = False
            .DrawStyle = 0
            .BackColor = &H808080
            .Width = m_intBarWidth
        End With
    End If

End Sub



Public Sub Resize()

    Dim lpRect As RECT
    Dim lWidth As Long
    Dim lHeight As Long
    
On Error GoTo Error_Exit
    
    GetWindowSize lWidth, lHeight
    
    If m_bInitialResize = False Then
        m_objLeftSide.left = m_intLeftPad
        m_objLeftSide.top = m_intReservedSpaceTop + m_intTopPad
        m_objLeftSide.Height = lHeight - m_intTotalYPadding
        If Not m_objRightSide Is Nothing Then
            m_objRightSide.top = m_objLeftSide.top
            m_objRightSide.Height = m_objLeftSide.Height
            m_objRightSide.left = m_intLeftPad + m_objLeftSide.Width + m_intCenterPad
            m_objRightSide.Width = lWidth - m_intCenterPad - m_intLeftPad - m_objLeftSide.Width - m_intRightPad
        End If
        If Not m_objBottom Is Nothing Then
            m_objLeftSide.Height = m_objLeftSide.Height * 0.6
            If Not m_objRightSide Is Nothing Then
                m_objRightSide.Height = m_objLeftSide.Height
            End If
            m_objBottom.left = m_intLeftPad
            m_objBottom.top = m_intTopPosition + m_objLeftSide.Height + m_intCenterPad
            m_objBottom.Height = lHeight - m_objLeftSide.Height - m_intTotalYPadding - m_intCenterPad
            m_objBottom.Width = lWidth - m_intTotalXPadding
            '
        End If

        '   Now set the splitter bar(s)
        If m_intStyle And frstHorizontal Then
            m_imgHSplitter.left = m_intLeftPad
            m_imgHSplitter.top = m_objLeftSide.Height + m_intTopPosition
            m_imgHSplitter.Width = lWidth - m_intLeftPad - m_intRightPad
            m_imgHSplitter.Height = m_intCenterPad
            m_imgHSplitter.Visible = True
        End If
        If m_intStyle And frstVertical Then
            m_imgVSplitter.left = m_intLeftPad + m_objLeftSide.Width
            m_imgVSplitter.top = m_objLeftSide.top
            m_imgVSplitter.Width = m_intCenterPad
            m_imgVSplitter.Height = m_objRightSide.Height
            m_imgVSplitter.Visible = True
        End If
        m_bInitialResize = True
    Else
        ResizeObjects
    End If
    
Error_Exit:
'    MsgBox Err.Description
End Sub

Private Sub ResizeBottomPane(lYSpaceAllowed As Long, lXSpaceAllowed As Long)
    '
    '   Now position the bottom pane
    On Error GoTo BottomTooSmall
    If m_intBorderSize = m_objLeftSide.Height Then
        m_objBottom.top = m_intTopPosition
    Else
        m_objBottom.top = m_objLeftSide.Height + m_intCenterPad + m_intTopPosition
    End If
    m_objBottom.Height = lYSpaceAllowed - m_objLeftSide.Height - m_intCenterPad - m_intTotalYPadding
    m_objBottom.Width = lXSpaceAllowed - m_intTotalXPadding
    
    Exit Sub

BottomTooSmall:
    m_objBottom.Height = 0

End Sub

Private Sub ResizeObjects()
    Dim intBarPos As Integer
    Dim lHeight As Long
    Dim lWidth As Long
    
    GetWindowSize lWidth, lHeight
    
    On Error GoTo ResizeObjects_Error
    
    '
    '   If there is no bottom panee ...
    If m_objBottom Is Nothing Then
        ResizeRightAndLeft lHeight, lWidth
    Else
        '
        '   If there is no right pane ...
        If m_objRightSide Is Nothing Then
            ResizeSingleTopAndBottom lHeight, lWidth
        Else
            ResizeDoubleTopAndBottom lHeight, lWidth
        End If
    End If
    Exit Sub
    
ResizeObjects_Error:
    MsgBox Err.Description
    Exit Sub
    
End Sub

Private Sub ResizeDoubleTopAndBottom(lYSpaceAllowed As Long, lXSpaceAllowed As Long)
    Dim lVerticalSpace As Long
    
    lVerticalSpace = lYSpaceAllowed - m_objBottom.Height - m_intCenterPad
    ResizeRightAndLeft lVerticalSpace, lXSpaceAllowed
    ResizeBottomPane lYSpaceAllowed, lXSpaceAllowed
    m_imgHSplitter.top = m_objLeftSide.Height + m_intTopPosition
    m_imgHSplitter.Width = m_objBottom.Width

End Sub


Private Sub ResizeRightAndLeft(lYSpaceAllowed As Long, lXSpaceAllowed As Long)
    Static intLeftSideSize As Integer
    
    '
    '   Do the heights first, it's easier
    On Error GoTo PanesTooSmall
    m_objLeftSide.Height = lYSpaceAllowed - m_intTotalYPadding
    m_objRightSide.Height = m_objLeftSide.Height
    '
    '   Now the widths, keep the left pane the same and adjust the right
    
    '
    '   Use the error handler to catch when the window gets to an illegal size
    On Error GoTo RightSideTooSmall
    
    '   if the right side is not hidden, calcualte the width of the right side holdeing the right side constant
    If intLeftSideSize = 0 Then
        m_objRightSide.Width = lXSpaceAllowed - m_intTotalXPadding - m_intCenterPad - m_objLeftSide.Width
        m_objLeftSide.Width = lXSpaceAllowed - m_intTotalXPadding - m_objRightSide.Width - m_intCenterPad
    Else
        m_objLeftSide.Width = lXSpaceAllowed - m_intTotalXPadding + m_intBorderSize
        If m_objLeftSide.Width > (intLeftSideSize + m_intCenterPad) Then
            m_objLeftSide.Width = intLeftSideSize
            intLeftSideSize = 0
            m_objRightSide.left = m_intLeftPad + m_objLeftSide.Width + m_intCenterPad
            m_objRightSide.Width = lXSpaceAllowed - m_intTotalXPadding - m_intCenterPad - m_objLeftSide.Width
        End If
    End If
    
    '
    '   Reposition the vertical splitter image
    m_imgVSplitter.left = m_objLeftSide.Width + m_intLeftPad
    m_imgVSplitter.Height = m_objLeftSide.Height
    Exit Sub

RightSideTooSmall:
    m_objRightSide.Width = m_intBorderSize
    If intLeftSideSize = 0 Then intLeftSideSize = m_objLeftSide.Width
    Resume Next
    
PanesTooSmall:
    m_objLeftSide.Height = 0
    m_objRightSide.Height = 0
    '
    '   by NOT making them invisible I let the error handler handle the error everytime
    '   and then exit.
    Exit Sub
    
    
End Sub

Private Sub ResizeSingleTopAndBottom(lYSpaceAllowed As Long, lXSpaceAllowed As Long)
    '
    '   A horizontal splitter with only a single top pane
    '
    '   The width of both panes is easy
    m_objLeftSide.Width = lXSpaceAllowed - m_intTotalXPadding
    m_objBottom.Width = m_objLeftSide.Width
    '
    '   Now lets look at the heights
    '       We keep the bottom pane the same size and adjust the top
    On Error GoTo TopTooSmall
    m_objLeftSide.Height = lYSpaceAllowed - m_intTotalYPadding - m_intCenterPad - m_objBottom.Height
    '
    '   Now position the bottom pane
    On Error GoTo BottomTooSmall
    If m_intBorderSize = m_objLeftSide.Height Then
        m_objBottom.top = m_intTopPad + m_intReservedSpaceTop
    Else
        m_objBottom.top = m_objLeftSide.Height + m_intCenterPad + m_intTopPosition
    End If
    m_objBottom.Height = lYSpaceAllowed - m_objLeftSide.Height - m_intCenterPad - m_intTotalYPadding
    
BottomSizeHandled:
    '   Reset the splitter image
    m_imgHSplitter.top = m_objLeftSide.Height + m_intTopPosition
    m_imgHSplitter.Width = m_objLeftSide.Width
    Exit Sub
    
TopTooSmall:
'    Stop
    m_objLeftSide.Height = m_intBorderSize
    Resume Next
    
BottomTooSmall:
    m_objBottom.Height = 0
    Resume BottomSizeHandled
    
End Sub


Private Sub VerticalSplitterResize()
    Dim intBarPos As Integer
    Dim lWindowWidth As Long
    
    lWindowWidth = GetWindowWidth
    
    intBarPos = m_picVSplitter.left
    m_objLeftSide.Width = intBarPos - m_intLeftPad - (m_intCenterPad / 2)
    m_objRightSide.left = intBarPos + m_intCenterPad
    m_objRightSide.Width = lWindowWidth - m_intLeftPad - m_intCenterPad - m_intRightPad - m_objLeftSide.Width
    m_imgVSplitter.left = intBarPos - (m_intCenterPad / 2)

End Sub

Private Sub m_imgHSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "Horizontal Splitter Height:  " & m_imgHSplitter.Height
    With m_imgHSplitter
        m_picHSplitter.Move .left, .top, .Width, m_intBarWidth
    End With
    m_picHSplitter.Visible = True
    m_picHSplitter.ZOrder 0
    m_bMoving = True
End Sub


Private Sub m_imgHSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    Dim lHeight As Long
    
    lHeight = GetWindowHeight
    '
    '   There are, as you can see, several conditios that we need to account for
    If m_bMoving Then
        sglPos = Y + m_imgHSplitter.top
        If sglPos < (m_intMinHeight + m_intReservedSpaceTop + m_intTopPad) Then
            '
            '   Don't move the splitter bar above the minimum height
            m_picHSplitter.top = m_intMinHeight + m_intReservedSpaceTop + m_intTopPad
        ElseIf sglPos <= 0 Or sglPos <= m_intTopPosition Then
            '
            '   Don't move the splitter bar above the top of the container
            m_picHSplitter.top = m_intTopPosition
        ElseIf sglPos > (m_objBottom.Height + m_objBottom.top) Then
            '
            '   Don't move the splitter bar below the bottom of the bottom pane
            sglPos = m_objBottom.top + m_objBottom.Height
        ElseIf sglPos > (lHeight - m_intMinBottomHeight - m_intReservedSpaceBottom - m_intBottomPad) Then
            '
            '   Don't resize the bottom to a size smaller than the minimum
            sglPos = m_objBottom.top + m_intMinBottomHeight
        Else
            m_picHSplitter.top = sglPos
        End If
    End If
End Sub


Private Sub m_imgHSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_picHSplitter.Visible = False
    m_bMoving = False
    HorizontalSplitterResize
End Sub


Private Sub m_imgVSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Debug.Print "m_imgVSplitter.Width = " & m_imgVSplitter.Width
    With m_imgVSplitter
        m_picVSplitter.Move .left, .top, .Width \ 2, .Height - 20
    End With
    m_picVSplitter.Visible = True
    m_picVSplitter.ZOrder 0
    m_bMoving = True
End Sub


Private Sub m_imgVSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    Dim lWidth As Long
    Dim lHeight As Long
    
    GetWindowSize lWidth, lHeight
    
    If m_bMoving Then
        sglPos = X + m_imgVSplitter.left
        If sglPos < m_sglSplitLimit Then
            m_picVSplitter.left = m_sglSplitLimit
        ElseIf sglPos > lWidth - m_sglSplitLimit Then
            m_picVSplitter.left = lWidth - m_sglSplitLimit
        Else
            m_picVSplitter.left = sglPos
        End If
    End If
End Sub


Private Sub m_imgVSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_picVSplitter.Visible = False
    m_bMoving = False
    VerticalSplitterResize
End Sub


Public Property Let SpaceOnLeft(ByVal vNewValue As Integer)
Attribute SpaceOnLeft.VB_Description = "The number of pixels between the left side of a pane and the control to the left of the pane."
Attribute SpaceOnLeft.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    m_intLeftPad = vNewValue
    PropertyChanged "SpaceOnLeft"
End Property

Public Property Get SpaceOnLeft() As Integer
    SpaceOnLeft = m_intLeftPad
End Property

Public Property Let SpaceInCenter(ByVal vNewValue As Integer)
Attribute SpaceInCenter.VB_Description = "The space in pizels that is to be between two panes."
Attribute SpaceInCenter.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    m_intCenterPad = vNewValue
    PropertyChanged "SpaceInCenter"
End Property

Public Property Get SpaceInCenter() As Integer
    SpaceInCenter = m_intCenterPad
End Property

Public Property Let SpaceOnRight(ByVal vNewValue As Integer)
Attribute SpaceOnRight.VB_Description = "The number of pixels between the right side of a pane and the control to the right of the pane."
Attribute SpaceOnRight.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    m_intRightPad = vNewValue
    PropertyChanged "SpaceOnRight"
End Property

Public Property Get SpaceOnRight() As Integer
    SpaceOnRight = m_intRightPad
End Property

Public Property Let SpaceOnBottom(ByVal vNewValue As Integer)
Attribute SpaceOnBottom.VB_Description = "The nnumber of pixels between the bottom of a pane and the control below the pane"
Attribute SpaceOnBottom.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    m_intBottomPad = vNewValue
    PropertyChanged "SpaceOnBottom"
End Property

Public Property Get SpaceOnBottom() As Integer
    SpaceOnBottom = m_intBottomPad
End Property

Private Sub UserControl_InitProperties()
    m_intTopPad = m_cnstTopPadding
    m_intLeftPad = m_cnstLeftPadding
    m_intCenterPad = m_cnstCenterPadding
    m_intRightPad = m_cnstRightPadding
    m_intBottomPad = m_cnstBottomPadding
    m_intStyle = m_cnstStyle
    m_intBarWidth = m_cnstBarWidth
    m_intMinHeight = m_cnstMinHeight
    m_intMinBottomHeight = m_cnstMinBottomHeight
    m_Cursor = spltDefault
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_intTopPad = PropBag.ReadProperty("SpaceOnTop", m_cnstTopPadding)
    m_intLeftPad = PropBag.ReadProperty("SpaceOnLeft", m_cnstLeftPadding)
    m_intCenterPad = PropBag.ReadProperty("SpaceInCenter", m_cnstCenterPadding)
    m_intRightPad = PropBag.ReadProperty("SpaceOnRight", m_cnstRightPadding)
    m_intBottomPad = PropBag.ReadProperty("SpaceOnBottom", m_cnstBottomPadding)
    m_intStyle = PropBag.ReadProperty("Style", m_cnstStyle)
    m_intBarWidth = PropBag.ReadProperty("BarWidth", m_cnstBarWidth)
    m_intMinHeight = PropBag.ReadProperty("MinimumHeightTopPane", m_cnstMinHeight)
    m_intMinBottomHeight = PropBag.ReadProperty("MinimumHeightBottomPane", m_cnstMinBottomHeight)
    m_Cursor = PropBag.ReadProperty("MouseCursor", spltDefault)
    m_intReservedSpaceTop = PropBag.ReadProperty("ReservedSpaceTop", 0)
    m_intReservedSpaceBottom = PropBag.ReadProperty("ReservedSpaceBottom", 0)
End Sub


Private Sub UserControl_Terminate()
'
'   At this point the parent is already gone so we canot clear out the controls here
'
'    If m_intStyle And frstvertical Then
'        Parent.Controls.Remove "m_imgVSplitter"
'        Parent.Controls.Remove "m_picVSplitter"
'    End If

    If Not m_objLeftSide Is Nothing Then
        Set m_objLeftSide = Nothing
    End If
    If Not m_objRightSide Is Nothing Then
        Set m_objRightSide = Nothing
    End If
    If Not m_objBottom Is Nothing Then
        Set m_objBottom = Nothing
    End If
    If Not m_objContainer Is Nothing Then
        Set m_objContainer = Nothing
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SpaceOnTop", m_intTopPad, m_cnstTopPadding)
    Call PropBag.WriteProperty("SpaceOnLeft", m_intLeftPad, m_cnstLeftPadding)
    Call PropBag.WriteProperty("SpaceOnRight", m_intRightPad, m_cnstRightPadding)
    Call PropBag.WriteProperty("SpaceOnBottom", m_intBottomPad, m_cnstBottomPadding)
    Call PropBag.WriteProperty("SpaceInCenter", m_intCenterPad, m_cnstCenterPadding)
    Call PropBag.WriteProperty("Style", m_intStyle, m_cnstStyle)
    Call PropBag.WriteProperty("BarWidth", m_intBarWidth, m_cnstBarWidth)
    Call PropBag.WriteProperty("MinimumHeightTopPane", m_intMinHeight, m_cnstMinHeight)
    Call PropBag.WriteProperty("MinimumHeightBottomPane", m_intMinBottomHeight, m_cnstMinBottomHeight)
    Call PropBag.WriteProperty("MouseCursor", m_Cursor, spltDefault)
    Call PropBag.WriteProperty("ReservedSpaceTop", m_intReservedSpaceTop, 0)
    Call PropBag.WriteProperty("ReservedSpaceBottom", m_intReservedSpaceBottom, 0)
End Sub



Public Property Let BarWidth(ByVal vNewValue As Integer)
Attribute BarWidth.VB_Description = "The width of the Bar that is shown as the splitter is moved in the Container"
Attribute BarWidth.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    m_intBarWidth = vNewValue
    PropertyChanged "BarWidth"
End Property

Public Property Get BarWidth() As Integer
    BarWidth = m_intBarWidth
End Property

'Public Property Let MinimumHeight(ByVal vNewValue As Integer)
'    m_intMinHeight = vNewValue
'End Property
'
'Public Property Get MinimumHeight() As Integer
'    MinimumHeight = m_intMinHeight
'End Property
'
'Public Property Let MinimumWidth(ByVal vNewValue As Integer)
'    m_intMinWidth = vNewValue
'End Property
'
'Public Property Get MinimumWidth() As Integer
'    MinimumWidth = m_intMinWidth
'End Property

Public Property Let Style(ByRef vNewValue As Style)
Attribute Style.VB_Description = "The 'Style' is how the spliter bars show up in the Container."
Attribute Style.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    m_intStyle = vNewValue
    PropertyChanged "Style"
End Property

Public Property Get Style() As Style
    Style = m_intStyle
End Property

'Public Property Let BottomHeightMininum(ByVal vNewValue As Integer)
'    m_intMinBottomHeight = vNewValue
'End Property

'Public Property Get BottomHeightMininum() As Integer
'    BottomHeightMininum = m_intMinBottomHeight
'End Property

Public Property Let ReservedSpaceBottom(ByVal vNewValue As Integer)
Attribute ReservedSpaceBottom.VB_Description = "The space to reserve for a StatusBar or other control"
Attribute ReservedSpaceBottom.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    m_intReservedSpaceBottom = vNewValue
    PropertyChanged "ReservedSpaceBottom"
End Property

Public Property Get ReservedSpaceBottom() As Integer
    ReservedSpaceBottom = m_intReservedSpaceBottom
End Property

Public Property Let ReservedSpaceTop(ByVal vNewValue As Integer)
Attribute ReservedSpaceTop.VB_Description = "The space to reserve for a StatusBar or other control"
Attribute ReservedSpaceTop.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    m_intReservedSpaceTop = vNewValue
    PropertyChanged "ReservedSpaceTop"
End Property

Public Property Get ReservedSpaceTop() As Integer
    ReservedSpaceTop = m_intReservedSpaceTop
End Property

Public Property Let MinimumHeighTopPane(ByRef vNewValue As Integer)
Attribute MinimumHeighTopPane.VB_Description = "The minimum height of the top pane.  Used only when moving a SplitterBar, not resizing the window."
Attribute MinimumHeighTopPane.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    m_intMinHeight = vNewValue
    PropertyChanged "MinimumHeightTopPane"
End Property

Public Property Get MinimumHeighTopPane() As Integer
    MinimumHeighTopPane = m_intMinHeight
End Property

Public Property Let MinimumHeightBottomPane(ByRef vNewValue As Integer)
Attribute MinimumHeightBottomPane.VB_Description = "The minimum height of the bottom pane.  Used only when moving a SplitterBar, not resizing the window."
Attribute MinimumHeightBottomPane.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    m_intMinBottomHeight = vNewValue
    PropertyChanged "MinimumHeightBottomPane"
End Property

Public Property Get MinimumHeightBottomPane() As Integer
    MinimumHeightBottomPane = m_intMinBottomHeight
End Property

Public Property Let MouseCursor(ByRef vNewValue As Cursor)
Attribute MouseCursor.VB_Description = "spltDefault is the windows default cursor.  spltCustom requires an RES file with oindex 101 and 102."
Attribute MouseCursor.VB_ProcData.VB_Invoke_PropertyPut = ";Misc"
    m_Cursor = vNewValue
    PropertyChanged "MouseCursor"
End Property

Public Property Get MouseCursor() As Cursor
    MouseCursor = m_Cursor
End Property
