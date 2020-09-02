VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NewCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements SliceAndDice.ISadAddin
Private mcbToolbar As Office.CommandBarControl
Private mcbMenuBar As Office.CommandBarControl
Private Indent As Office.CommandBarControl
Private BlankLine As Office.CommandBarControl
Public WithEvents IndentHandler As CommandBarEvents
Attribute IndentHandler.VB_VarHelpID = -1
Public WithEvents IndentHandlerPop As CommandBarEvents
Attribute IndentHandlerPop.VB_VarHelpID = -1

Public WithEvents BlanklinesHandler As CommandBarEvents
Attribute BlanklinesHandler.VB_VarHelpID = -1
Public WithEvents BlankLinesHandlerPop As CommandBarEvents
Attribute BlankLinesHandlerPop.VB_VarHelpID = -1

Public WithEvents ProcHeadHandler As CommandBarEvents
Attribute ProcHeadHandler.VB_VarHelpID = -1
Public WithEvents CleardebugHandler As CommandBarEvents
Attribute CleardebugHandler.VB_VarHelpID = -1
Public WithEvents CloseUnusedWindowsHandler As CommandBarEvents
Attribute CloseUnusedWindowsHandler.VB_VarHelpID = -1
Public WithEvents SettingsHandler As CommandBarEvents
Attribute SettingsHandler.VB_VarHelpID = -1
Public WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
Private CodeReview As New CodeReviewer
Private vbME As VBIDE.VBE
Private Parent As SliceAndDice.Wizard
Private MySadCommands As SliceAndDice.CSadCommands

Private Sub BlanklinesHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo ERR_HANDLER
CodeReview.ReduceBlanks vbME
Exit Sub
ERR_HANDLER:
    Err.Clear
    
End Sub

Private Sub BlankLinesHandlerPop_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo ERR_HANDLER
CodeReview.ReduceBlanks vbME
Exit Sub
ERR_HANDLER:
    Err.Clear
    
End Sub

Private Sub Class_Initialize()
On Error Resume Next
    Set MySadCommands = New SliceAndDice.CSadCommands
    With MySadCommands
         .ParameterDelimiter = ","
         .ParameterTypeDelimiter = ":"
         .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision & " Beta"
    End With
End Sub

Private Sub Class_Terminate()
    Set MySadCommands = Nothing
End Sub

Private Sub CleardebugHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo ERR_HANDLER
Dim w As Window
    Set w = vbME.Windows("Immediate")
    If Not w Is Nothing Then
        If w.Visible = True Then
            w.SetFocus
            VBA.Interaction.SendKeys ("^({Home})"), True
            VBA.Interaction.SendKeys ("^(+({End}))"), True
            VBA.Interaction.SendKeys ("{Del}"), True
        End If
    End If
    Exit Sub
ERR_HANDLER:
    Err.Clear
    
End Sub

Private Sub CloseUnusedWindowsHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo ERR_HANDLER
Dim win As Window
    For Each win In vbME.Windows
        If Not win Is vbME.ActiveWindow Then
            If win.Type = 0 Or win.Type = 1 Then
         
                win.Close
            End If
        End If
    Next
Exit Sub
ERR_HANDLER:
    Err.Clear
    
End Sub

Private Sub IndentHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo ERR_HANDLER

CodeReview.IndentLines vbME
Exit Sub
ERR_HANDLER:
    Err.Clear
    
End Sub

Private Sub IndentHandlerPop_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo ERR_HANDLER

CodeReview.IndentLines vbME
Exit Sub
ERR_HANDLER:
    Err.Clear
    
End Sub

Private Property Get ISadAddin_CommandSet() As SliceAndDice.CSadCommands
On Error Resume Next
    Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
    '
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommand
    Dim bEOLAtEndOfLine As Boolean

    Dim lWrapLength As Long
    Dim lThisWrap As Long
    Dim lLineOffset As Long

    Dim sWordWrapped As String
    Dim sToWrap As String
    Dim Token1 As String
    Dim Token2 As String
    Dim SoftVar1 As String
    Dim SoftVar2 As String
    Dim sOperator As String

    With II
         If Not MySadCommands(.SoftCommandName & "*C") Is Nothing Then
            If MySadCommands(.SoftCommandName & "*C").IsInline Then Exit Function
            
'            Select Case UCase(.SoftCommandName)
'                    Case "TESTING2"
'                        Dim softVars As New SliceAndDice.CAssocArray
'                        softVars.Add "test1", "valtest1"
'                        softVars.Add "test2", "valtest2"
'                        softVars.Add "test3", "valtest3"
'                        softVars.Add "test4", "valtest4"
'
'                        Parent.DoInsertion softVars, "Lists - Test"
'                        ISadAddin_ExecuteSoftCommand = True
'
'                    'Case "TESTING"
''                         MsgBox "Soft command 'Testing' executed."
''                         ISadAddin_ExecuteSoftCommand = True
''
''                   Case "ANOTHERONE"
''                         MsgBox "Soft command 'AnotherOne' executed."
''                         ISadAddin_ExecuteSoftCommand = True
'
'                  'Case "X"
'                  'Case "Y"
'                  'Case "Z"
'            End Select
         End If
    End With

EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:
    Exit Function

EH_SADAddin_ISadAddin_ExecuteSoftCommand:
    Parent.LogError "SADAddin", "ISadAddin_ExecuteSoftCommand", Err.Number, Err.Description
    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue

    Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef softVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline
    Dim Template As CTemplate

    Dim Area As Long
    Dim CurrSet As Long
    Dim bInlineCommandExecuted As Boolean
    Dim lParamCount As Long

    Dim sDefault As String
    Dim sT As String
    Dim sVar1 As String
    Dim sVar2 As String
    Dim sVar3 As String

    If Not MySadCommands(sInlineSoftCommandName & "*I") Is Nothing Then
       If Not MySadCommands(sInlineSoftCommandName & "*I").IsInline Then Exit Function
'       Select Case sInlineSoftCommandName
'              Case "INLINE12"
'                   sResults = InputBox("What do you want in here ?")
'                   ISadAddin_ExecuteSoftCommandInline = True
'
'             Case "INLINEX"
'                   sResults = ""
'                   ISadAddin_ExecuteSoftCommandInline = True
'
'             Case "INLINEY"
'                   sResults = ""
'                   ISadAddin_ExecuteSoftCommandInline = True
'
'             Case "INLINEZ"
'                   sResults = ""
'                   ISadAddin_ExecuteSoftCommandInline = True
'
'        End Select
    End If

EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue:
    Exit Function

EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline:
    Parent.LogError "SADAddin", "ISadAddin_ExecuteSoftCommandInline", Err.Number, Err.Description
    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue

    Resume
End Function

Private Property Get ISadAddin_Externals() As SliceAndDice.CAssocArray
    '
End Property

Private Function ISadAddin_Shutdown() As Boolean
On Error Resume Next
    MySadCommands.Clear
    Set MySadCommands.Parent = Nothing
    Set MySadCommands = Nothing
    Set Parent = Nothing
mcbMenuBar.Delete
BlankLine.Delete
Indent.Delete
    ISadAddin_Shutdown = True
    
    
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard, vbInst As VBIDE.VBE) As Boolean
Dim cbToolbar As Object
Dim cbSubMenubar As Office.CommandBarControl
Dim cbMenu As Object
On Error Resume Next
    Err.Clear

    If UCase(oParent.Version) <> UCase(MySadCommands.Attributes("Version")) Then
       If MsgBox("Version mismatch:" & Chr(13) & Chr(9) & "SADAddin.NewCommands = " & MySadCommands.Attributes("Version") & Chr(13) & Chr(9) & "Slice and Dice = " & oParent.Version & Chr(13) & Chr(13) & "Continue loading DLL ?", vbYesNo, "*** WARNING - DLL Version mismatch **") = vbNo Then
          Exit Function
       End If
    End If

    Set vbME = vbInst
    Err.Clear
        Set Parent = oParent
        Set MySadCommands.Parent = oParent
        MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath & "sadCodeReview.txt")

        Set cbMenu = vbME.CommandBars("Menu Bar")
        Set mcbMenuBar = cbMenu.Controls.Add(msoControlPopup, , , 1)
        mcbMenuBar.Caption = "E&xtra Tools"
        Set Me.MenuHandler = vbME.Events.CommandBarEvents(mcbMenuBar)
        
        Set cbSubMenubar = mcbMenuBar.Controls.Add(msoControlButton)
        cbSubMenubar.Caption = "Indent &Lines"
        Clipboard.SetData LoadResPicture("INDENT", vbResBitmap)
        cbSubMenubar.PasteFace
        Set IndentHandler = vbME.Events.CommandBarEvents(cbSubMenubar)
        Set cbSubMenubar = Nothing
        
        Set cbSubMenubar = mcbMenuBar.Controls.Add(msoControlButton)
        cbSubMenubar.Caption = "&Reduce blank lines"
        Clipboard.SetData LoadResPicture("BLANKS", vbResBitmap)
        cbSubMenubar.PasteFace
        Set BlanklinesHandler = vbME.Events.CommandBarEvents(cbSubMenubar)
        Set cbSubMenubar = Nothing
        
        Set cbSubMenubar = mcbMenuBar.Controls.Add(msoControlButton)
        cbSubMenubar.Caption = "Add &Proc Header"
        cbSubMenubar.ToolTipText = "Add a quick and easy Procedure header"
        Clipboard.SetData LoadResPicture("PROCHEADER", vbResBitmap)
        cbSubMenubar.PasteFace
        Set ProcHeadHandler = vbME.Events.CommandBarEvents(cbSubMenubar)
        Set cbSubMenubar = Nothing
        
        Set cbSubMenubar = mcbMenuBar.Controls.Add(msoControlButton)
        cbSubMenubar.Caption = "Close &Unused Windows"
        cbSubMenubar.ToolTipText = "Close all unused windows in a flash"
        Clipboard.SetData LoadResPicture("CLOSEWINDOWS", vbResBitmap)
        cbSubMenubar.PasteFace
        Set CloseUnusedWindowsHandler = vbME.Events.CommandBarEvents(cbSubMenubar)
        Set cbSubMenubar = Nothing
        
        Set cbSubMenubar = mcbMenuBar.Controls.Add(msoControlButton)
        cbSubMenubar.Caption = "C&lear Debug Window"
        cbSubMenubar.ToolTipText = "Clear your immediate window"
        Clipboard.SetData LoadResPicture("CLEARDEBUG", vbResBitmap)
        cbSubMenubar.PasteFace
        Set CleardebugHandler = vbME.Events.CommandBarEvents(cbSubMenubar)
        Set cbSubMenubar = Nothing
        
        Set cbSubMenubar = mcbMenuBar.Controls.Add(msoControlButton)
        cbSubMenubar.Caption = "CodeReviewer Settings"
        cbSubMenubar.ToolTipText = "Clear your immediate window"
        cbSubMenubar.PasteFace
        Set SettingsHandler = vbME.Events.CommandBarEvents(cbSubMenubar)
        Set cbSubMenubar = Nothing
        
        Set cbToolbar = vbME.CommandBars("Code Window")
        Set Indent = cbToolbar.Controls.Add(msoControlButton, , , 1)
        Indent.Caption = "Indent"
        Clipboard.SetData LoadResPicture("INDENT", vbResBitmap)
        Indent.PasteFace
        Set IndentHandlerPop = vbME.Events.CommandBarEvents(Indent)
       
        Set BlankLine = cbToolbar.Controls.Add(msoControlButton)
        BlankLine.Caption = "&Reduce blank lines"
        Clipboard.SetData LoadResPicture("BLANKS", vbResBitmap)
        BlankLine.PasteFace
        Set BlankLinesHandlerPop = vbME.Events.CommandBarEvents(BlankLine)
        Set cbSubMenubar = Nothing
        ISadAddin_Startup = (Err.Number = 0)
    Err.Clear
End Function

Private Sub ProcHeadHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
CodeReview.InsertProcHeaders vbME
End Sub

Private Sub SettingsHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
CodeReview.Settings
End Sub
