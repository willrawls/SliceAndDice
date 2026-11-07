VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements SliceAndDice.ISadAddin

Private Parent As SliceAndDice.Wizard
Private MySadCommands As SliceAndDice.CSadCommands

Private Sub Class_Initialize()
    Set MySadCommands = New SliceAndDice.CSadCommands
    MySadCommands.ParameterDelimiter = ", "
    MySadCommands.ParameterTypeDelimiter = " : "
End Sub

Private Sub Class_Terminate()
    Set MySadCommands = Nothing
End Sub

Private Function ISadAddin_ExecuteInlineSoftCommand(sCommand As String, sParameters As String) As Variant
    '
End Function

Private Function ISadAddin_ExecuteSoftCommand(sCommand As String, sParameters As String) As Variant
On Error GoTo EH_MyFirst_ISadAddin_ExecuteSoftCommand
    Dim CurrFunction As CSadCommand

    If Not MySadCommands(sCommand) Is Nothing Then
       Set CurrFunction = MySadCommands(sCommand)
            CurrFunction.Parameters = sParameters
            If CurrFunction.SyntaxIsValid Then
               Select Case UCase(sCommand)
                      Case "XGETTOKEN"
                      Case "XTOKEN"
               End Select
            End If
    End If
    Set CurrFunction = Nothing

EH_MyFirst_ISadAddin_ExecuteSoftCommand_Continue:
    Exit Function

EH_MyFirst_ISadAddin_ExecuteSoftCommand:
    LogError "MyFirst", "ISadAddin_ExecuteSoftCommand", Err.Number, Err.Description
    Resume EH_MyFirst_ISadAddin_ExecuteSoftCommand_Continue

    Resume
End Function

Private Function ISadAddin_Shutdown() As Boolean
On Error Resume Next
    MySadCommands.Clear
    Set MySadCommands.Parent = Nothing
    Set MySadCommands = Nothing
    Set Parent = Nothing

    ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard) As Boolean
On Error Resume Next
    Err.Clear

    Set Parent = oParent
    Set MySadCommands.Parent = oParent

    SetupSoftCommandCollection

    ISadAddin_Startup = (Err.Number = 0)
End Function

Private Sub SetupSoftCommandCollection()
    With MySadCommands
         .Clear
         .Add "xGetToken", "SoftVarIn : StringOrSoftVar, nToken : Long Opt, sDelim : QuotedString Opt"
         .Add "xToken", "SoftVarIn : StringOrSoftVar, nToken : Long Opt, sDelim : QuotedString Opt"
    End With
End Sub
