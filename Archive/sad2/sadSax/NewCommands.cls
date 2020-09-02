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

Private Patti           As sadSax.Application
Private Parent          As SliceAndDice.Wizard
Private MySadCommands   As SliceAndDice.CSadCommands
Private Externals       As SliceAndDice.CAssocArray

Private Sub Class_Initialize()
On Error Resume Next
    Set Patti = New sadSax.Application
    Set MySadCommands = New SliceAndDice.CSadCommands
    With MySadCommands
         .ParameterDelimiter = ","
         .ParameterTypeDelimiter = ":"
         .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision & " Beta"
    End With
    Set Externals = New SliceAndDice.CAssocArray
    With Externals
        .Remove "EOL"
        .Remove "EOL2"
        .Remove "TAB"
        .Remove "EOLTAB"
        .Remove "EOL2TAB"
        .Item("Sandy Sax") = "SANDYSAX"
    End With
End Sub

Private Sub Class_Terminate()
    With Patti
        Set .Sandy = Nothing
        Set .VbInst = Nothing
    End With
    Set Patti = Nothing
    Set MySadCommands = Nothing
    Set Externals = Nothing
End Sub

Private Property Get ISadAddin_CommandSet() As SliceAndDice.CSadCommands
On Error Resume Next
    Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
34   On Error Resume Next
35       Dim X           As Form
36       Dim sMessage    As String

         With frmSax
37          If .Parent Is Nothing Then
38              Set .Parent = Me
                .saxScript.AddExtension "{}", Patti.Sandy
                .saxScript.AddExtension "{}", Patti.VbInst
                .saxScript.AddExtension "-", Patti
39          End If
    
            Select Case UCase$(sValue)
            Case "SANDYSAX"
                    If .Visible Then
                       .Hide
                    Else
                       Set X = Nothing
                       Set X = Parent.SandyWindow
                           If X Is Nothing Then
                              .Show
                           Else
                              .Show 0, X
                           End If
                           If Not .Visible Then
                              .Show
                           End If
                       Set X = Nothing
                    End If
    
50          Case "HIDE ALL WINDOWS", "HIDEALLWINDOWS"
51              If .Visible Then .Hide
    
52          Case "UNLOAD"
53              Unload frmSax
54          End Select
        End With
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
'On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommand
'    Dim bEOLAtEndOfLine As Boolean
'
'    Dim lWrapLength As Long
'    Dim lThisWrap As Long
'    Dim lLineOffset As Long
'
'    Dim sWordWrapped As String
'    Dim sToWrap As String
'    Dim Token1 As String
'    Dim Token2 As String
'    Dim SoftVar1 As String
'    Dim SoftVar2 As String
'    Dim sOperator As String
'
'    With II
'         If Not MySadCommands(.SoftCommandName & "*C") Is Nothing Then
'            If MySadCommands(.SoftCommandName & "*C").IsInline Then Exit Function
'
'            Select Case UCase(.SoftCommandName)
'                   Case "TESTING"
'                         MsgBox "Soft command 'Testing' executed."
'                         ISadAddin_ExecuteSoftCommand = True
'
'                   Case "ANOTHERONE"
'                         MsgBox "Soft command 'AnotherOne' executed."
'                         ISadAddin_ExecuteSoftCommand = True
'
'                  'Case "X"
'                  'Case "Y"
'                  'Case "Z"
'            End Select
'         End If
'    End With
'
'EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:
'    Exit Function
'
'EH_SADAddin_ISadAddin_ExecuteSoftCommand:
'    Parent.LogError "SADAddin", "ISadAddin_ExecuteSoftCommand", Err.Number, Err.Description
'    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue
'
'    Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
'On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline
'    Dim Template As CTemplate
'
'    Dim Area As Long
'    Dim CurrSet As Long
'    Dim bInlineCommandExecuted As Boolean
'    Dim lParamCount As Long
'
'    Dim sDefault As String
'    Dim sT As String
'    Dim sVar1 As String
'    Dim sVar2 As String
'    Dim sVar3 As String
'
'    If Not MySadCommands(sInlineSoftCommandName & "*I") Is Nothing Then
'       If Not MySadCommands(sInlineSoftCommandName & "*I").IsInline Then Exit Function
'       Select Case sInlineSoftCommandName
'              Case "INLINE1"
'                   sResults = InputBox("What do you want in here ?")
'                   ISadAddin_ExecuteSoftCommandInline = True
'
'             'Case "INLINEX"
'                   sResults = ""
'                   ISadAddin_ExecuteSoftCommandInline = True
'
'             'Case "INLINEY"
'                   sResults = ""
'                   ISadAddin_ExecuteSoftCommandInline = True
'
'             'Case "INLINEZ"
'                   sResults = ""
'                   ISadAddin_ExecuteSoftCommandInline = True
'
'        End Select
'    End If
'
'EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue:
'    Exit Function
'
'EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline:
'    Parent.LogError "SADAddin", "ISadAddin_ExecuteSoftCommandInline", Err.Number, Err.Description
'    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue
'
'    Resume
End Function

Private Property Get ISadAddin_Externals() As SliceAndDice.CAssocArray
74   On Error Resume Next
75       Set ISadAddin_Externals = Externals
End Property
Private Function ISadAddin_Shutdown() As Boolean
On Error Resume Next
    With Patti
        Set .VbInst = Nothing
        Set .Sandy = Nothing
    End With

    MySadCommands.Clear
    Set MySadCommands.Parent = Nothing
    Set MySadCommands = Nothing
    Set Parent = Nothing

    ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard, VbInst As VBIDE.VBE) As Boolean
On Error Resume Next
    Err.Clear

    Set Parent = oParent
    Set MySadCommands.Parent = oParent
    
    With Patti
        Set .Sandy = oParent
        Set .VbInst = VbInst
    End With

    Err.Clear
       'MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath & "sadSax.txt")
        ISadAddin_Startup = (Err.Number = 0)
    Err.Clear
End Function
