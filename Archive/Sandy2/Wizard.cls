VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_oVBInst                       As Object
Private m_oUI                           As SandySupport.ISandyWindowMain
Private PropertyStack                   As SandySupport.CAssocArray

Implements SandySupport.ISandyWizard

Public Property Get Version() As String
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Sub OnConnection(ByVal VBInst As Object)
On Error Resume Next
   'Dim CurWindow As Window
    Dim sTemplateDatabasePath   As String
    Dim bFound                  As Boolean
    Dim bShown                  As Boolean
    Dim lFirstButton            As Long

    frmSplash.lblDLLsLoaded(1).Caption = "0"
    If GetSetting(App.ProductName, "Last", "Show Splash", True) Then
       bShown = True
       frmSplash.Show
       frmSplash.Refresh
    Else
       bShown = False
    End If

    If gbEvaluationHasExpired Then
       If MsgBox("Thank you for evaluating Slice and Dice. Your 30 day evaluation period has expired." & vbCr & "Continued, unregistered use of Slice and Dice violates the End User Agreement." & vbCr & vbCr & "If you have enjoyed using Slice and Dice, please consider perchasing it." & vbNewLine & vbTab & "Would you like to purchase it now ?", "EVALUATION HAS EXPIRED", vbYesNo) = vbYes Then
          BrowseTo "http://www.sliceanddice.com/register.html"
       End If
    End If

        Set m_oVBInst = VBInst                                                          ' Attach to the current VB IDE session

        Set m_oUI = CreateObject("SandyInstance.frmMain")

       'If m_oUI.lsbJumpTo.Version <> Version Then
       '   MsgBox "Version mismatch:" & vbCr & vbTab & "FirmSolutions.FSListBar = " & m_oUI.lsbJumpTo.Version & vbCr & vbTab & "Slice and Dice = " & Version, , "*** WARNING - OCX Version mismatch **"
       'End If
    
      ' Cause the VB IDE to create a dockable window for the add-in
'        Set addinSliceAndDice = m_oVBInst.Addins("SliceAndDice.Wizard")
'            sDescription = addinSliceAndDice.Description
'            guidMYTOOL = addinSliceAndDice.Guid
'            Set m_oUI = Nothing
'            Set m_oUI = m_oVBInst.Windows.CreateToolWindow(addinSliceAndDice, "SliceAndDice.uiAddInWindow", sDescription, guidMYTOOL, m_oUI)
'            If m_oUI Is Null Then
'               bFound = False
'               For Each CurWindow In m_oVBInst.Windows
'                   If CurWindow.Caption = sDescription Then
'                      bFound = True
'                      Set m_oUI = CurWindow
'                   End If
'               Next CurWindow
'               If Not bFound Then
'                  MsgBox "Failure to initialize VB Addin Window. Disconnect this addin now."
'               End If
'            End If
'        Set addinSliceAndDice = Nothing

        Set m_oUI.Parent = Me
        Set m_oUI.DBClassGen = CreateObject("SandyInstance.frmDBClassGen")
        Set m_oUI.DBClassGen.Parent = m_oUI

        DoEvents
          sTemplateDatabasePath = TemplateDatabasePath
        DoEvents: DoEvents: DoEvents
       'MsgBox "Template Database Path = " & TemplateDatabasePath
        Call m_oUI.InitializeAddinDLLs(modGeneral.EnumFiles(App.Path) & modGeneral.EnumFiles(TemplateDatabasePath))

        Set MenuHandler = AddButton(mcbAddinButton, "Add-Ins", 4, "&Slice and Dice", 30, True, False)

        If GetSetting(App.ProductName, "Last", "Show Paitbrush Icon", True) Then
           Set BarHandler = AddButton(mcbEditButton, "Standard", 1, "Slice and Dice", 30, False, True)
        End If

       'Set ImportHandler = AddButton(mcbImportButton, "Code Window", 1, "I&mport selection as new Template", 30, False, False)

        If GetSetting(App.ProductName, "Last", "Show On Module Right Click", True) Then
           If Not m_oVBInst.CommandBars("Code Window") Is Nothing Then
              lFirstButton = IIf(m_oVBInst.CommandBars("Code Window").Controls.Count > 4, 4, 1)
              Set FavoritesHandler = AddButton(mcbFavoritesButton, "Code Window", lFirstButton, "Insert a &Favorite", 30, True, False)
              Set ExternalsHandler = AddButton(mcbExternalsButton, "Code Window", lFirstButton + 1, "Call an &External", 30, False, False)
              Set ShortcutHandler = AddButton(mcbShortcutButton, "Code Window", lFirstButton + 2, "&Slice and Dice", 30, False, True)
           End If
        End If

       'ChangeToName = AddPopup(mcpChangeTo, "Controls", 1, "Modify Selected Controls", False, True)

        Set ChangeToHandler = AddButton(mcbChangeToButton, "Controls", 1, "Slice and Dice Modify", 30, False, False)
        Set AltChangeToHandler = AddButton(mcbAltChangeToButton, "Format", 8, "&Slice and Dice Modify", 30, False, False)
       'Set FormChangeToHandler = AddButton(mcbFormChangeToButton, "Forms", 1, "&Slice and Dice Modify", 0, False, False)

       'Set IDEEvents = New CIDEEvents
       'IDEEvents.Init m_oVBInst, m_oUI

    If m_oUI.mHotKeyOpenWindow Is Nothing Then
    Set m_oUI.mHotKeyOpenWindow = New cRegHotKey
    If Not m_oUI.mHotKeyOpenWindow Is Nothing Then
       With m_oUI.mHotKeyOpenWindow
            .Attach m_oUI.hwnd
            .RegisterKey "Sandy Activate", vbKeyS, MOD_CONTROL + MOD_SHIFT
            .RegisterKey "Sandy Repeat Insertion", vbKeyR, MOD_CONTROL + MOD_SHIFT
            .RegisterKey "Sandy Favorites", vbKeyF, MOD_CONTROL + MOD_SHIFT
            .RegisterKey "Sandy Externals", vbKeyE, MOD_CONTROL + MOD_SHIFT
       End With
    End If
    End If
    If bShown Then frmSplash.Hide
End Sub

Public Sub OnDisconnection()
On Error Resume Next
    Dim Cancel As Integer

  ' Make sure the edited Template (if one) is saved
    m_oUI.SaveTemplate
    m_oUI.Visible = False
    
  ' Remove buttons from VB5 ToolBars
    mcbAddinButton.Delete
    mcbEditButton.Delete
    mcbShortcutButton.Delete
   'mcbImportButton.Delete
    mcbChangeToButton.Delete
    mcbAltChangeToButton.Delete
    mcbFavoritesButton.Delete
    mcbExternalsButton.Delete
   'mcpChangeTo.Delete

  ' Insure all external object references are released correctly
    Set mcbAddinButton = Nothing
    Set mcbEditButton = Nothing
    Set mcbShortcutButton = Nothing
    Set mcbChangeToButton = Nothing
    Set mcbAltChangeToButton = Nothing
    Set mcbFavoritesButton = Nothing

    Set MenuHandler = Nothing
    Set BarHandler = Nothing
    Set ShortcutHandler = Nothing
    Set ChangeToHandler = Nothing
    Set AltChangeToHandler = Nothing
    Set FavoritesHandler = Nothing
    Set ExternalsHandler = Nothing

   'Set mcbImportButton = Nothing
   'Set mcpChangeTo = Nothing
   'Set ImportHandler = Nothing
   
   'Set IDEEvents = Nothing

  ' Save settings for next time
    SaveSetting App.ProductName, "Settings", "Exit after insert", IIf(m_oUI.ExitAfterInsert, "True", "False")
    SaveSetting App.ProductName, "Settings", "Last Template", m_oUI.CurrentTemplateNameAndCategory
    SaveFormPosition m_oUI
    
   'SaveSetting App.ProductName, "Settings", "Bar Arrange", lsbJumpTo.Arrange
   'SaveSetting App.ProductName, "Settings", "Bar View", lsbJumpTo.View

  ' Destroy object references
   'If frmCommandHelp.Visible Then frmCommandHelp.Hide
   'SaveFormPosition frmCommandHelp
   'Set frmCommandHelp.CurrCommand = Nothing
   'Set frmCommandHelp.SadCommandSet = Nothing
   'Unload frmCommandHelp

    m_oUI.HideAllWindows True
    m_oUI.ShutdownDLLs
    Set m_oUI = Nothing

    'Set SubClassIDE = Nothing
   'Set m_oUI.VBIDEWindow = Nothing
    Set m_oUI.Parent = Nothing
    Set m_oUI.DBClassGen = Nothing
    m_oUI.Form_Unload Cancel
    Unload m_oUI
    Set m_oUI = Nothing

  ' Disassociate external objects referenced in this object
    Set m_oVBInst = Nothing
End Sub
