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

Private Parent As SliceAndDice.Wizard

Private MySadCommands As SliceAndDice.CSadCommands
Private Externals As SliceAndDice.CAssocArray
Private Sub Class_Initialize()
On Error Resume Next
    Set MySadCommands = New SliceAndDice.CSadCommands
    Set Externals = New SliceAndDice.CAssocArray
    With MySadCommands
         .ParameterDelimiter = ","
         .ParameterTypeDelimiter = ":"
         .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision & " Beta"
    End With

    With Externals
         .Clear
         .Item("&Doom Clock") = "Show frmMain"
    End With
End Sub

Private Sub Class_Terminate()
    Set Externals = Nothing
    Set MySadCommands = Nothing
End Sub

Private Property Get ISadAddin_CommandSet() As SliceAndDice.CSadCommands
On Error Resume Next
    Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
    Dim x As Form

    Select Case UCase(sValue)
           Case "SHOW FRMMAIN"
                If frmMain.Visible Then
                   frmMain.Hide
                Else
                   Set x = Parent.SandyWindow
                       frmMain.Show 0, x
                   Set x = Nothing
                End If
           Case "HIDE ALL WINDOWS", "HIDEALLWINDOWS"
                If frmMain.Visible Then frmMain.Hide
    End Select
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

            Select Case UCase(.SoftCommandName)
                   Case "TESTING"
                         MsgBox "Soft command 'Testing' executed."
                         ISadAddin_ExecuteSoftCommand = True

                   Case "ANOTHERONE"
                         MsgBox "Soft command 'AnotherOne' executed."
                         ISadAddin_ExecuteSoftCommand = True

                  'Case "X"
                  'Case "Y"
                  'Case "Z"
            End Select
         End If
    End With

EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:
    Exit Function

EH_SADAddin_ISadAddin_ExecuteSoftCommand:
    Parent.LogError "SADAddin", "ISadAddin_ExecuteSoftCommand", Err.Number, Err.Description
    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue

    Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
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
       Select Case sInlineSoftCommandName
              Case "INLINE1"
                   sResults = InputBox("What do you want in here ?")
                   ISadAddin_ExecuteSoftCommandInline = True

             'Case "INLINEX"
                   sResults = ""
                   ISadAddin_ExecuteSoftCommandInline = True

             'Case "INLINEY"
                   sResults = ""
                   ISadAddin_ExecuteSoftCommandInline = True

             'Case "INLINEZ"
                   sResults = ""
                   ISadAddin_ExecuteSoftCommandInline = True

        End Select
    End If

EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue:
    Exit Function

EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline:
    Parent.LogError "SADAddin", "ISadAddin_ExecuteSoftCommandInline", Err.Number, Err.Description
    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue

    Resume
End Function

Private Property Get ISadAddin_Externals() As SliceAndDice.CAssocArray
    Set ISadAddin_Externals = Externals
End Property

Private Function ISadAddin_Shutdown() As Boolean
On Error Resume Next
    MySadCommands.Clear
    Set MySadCommands.Parent = Nothing
    Set MySadCommands = Nothing
    Set Parent = Nothing
    frmMain.tmrTimeLeft.Enabled = False
    Unload frmMain

    ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard, vbInst As VBIDE.VBE) As Boolean
On Error Resume Next
    Err.Clear

    'If UCase(oParent.Version) <> UCase(MySadCommands.Attributes("Version")) Then
    '   If MsgBox("Version mismatch:" & Chr(13) & Chr(9) & "SADAddin.NewCommands = " & MySadCommands.Attributes("Version") & Chr(13) & Chr(9) & "Slice and Dice = " & oParent.Version & Chr(13) & Chr(13) & "Continue loading DLL ?", vbYesNo, "*** WARNING - DLL Version mismatch **") = vbNo Then
    '      Exit Function
    '   End If
    'End If

    Set Parent = oParent
    Set MySadCommands.Parent = oParent

    Err.Clear
        'MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath & "sadDoomClock.txt")
        ISadAddin_Startup = (Err.Number = 0)
    Err.Clear
End Function
