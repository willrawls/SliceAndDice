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

Implements SandySupport.ISadAddin

Public Parent As SandySupport.ISandyWizard
Public MySadCommands As SandySupport.CSadCommands
Public Externals As SandySupport.CAssocArray

Private Sub Class_Initialize()
On Error Resume Next
    Set MySadCommands = New SandySupport.CSadCommands
    Set Externals = New SandySupport.CAssocArray
    With MySadCommands
         .Clear
         .ParameterDelimiter = ","
         .ParameterTypeDelimiter = ":"
         .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision
         With .Add("GetURL", False)
              .Aliases = "FileURL, GetWebPage, GetPage, FetchURL"
              .Comments = "Timeout defaults to 120 seconds"
              .OneLineDescription = "Fetches the contents of the specified URL to the indicated SoftVarOut"
              .Syntax = "SoftVarOut:SoftVar = URL:String"
         End With
         With .Add("GetFileURL", False)
              .Aliases = "URL, GetFileFromPage, GetRemoteFile, FetchFile"
              .Comments = "Timeout defaults to 120 seconds"
              .OneLineDescription = "Fetches a binary file from the specified URL to the indicated filename"
              .Syntax = "FilenameOut:String = URL:String"
         End With
    End With

    With Externals
         .Clear
         .Item("&Register Slice and Dice") = "Register"
         .Item("&Check for version updates") = "Update Files"
         .Item("&Submit current Template to CTL") = "Submit Template"
    End With
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Set Externals = Nothing
    Set MySadCommands = Nothing
    Set frmRegister.Parent = Nothing
    Unload frmRegister
End Sub

Private Property Get ISadAddin_CommandSet() As Object
On Error Resume Next
    Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
    Dim X As Form
    Dim sMessage As String

On Error Resume Next
    If frmRegister.Parent Is Nothing Then
       Set frmRegister.Parent = Me
    End If
    
    Select Case UCase(sValue)
           Case "REGISTER"
                If MySadCommands.Attributes("Registered") = "True" Then
                   MsgBox "You have already registered Slice and Dice successfully." & Chr(13) & Chr(9) & "Invoice # " & sadGetLicenseKey("Invoice Number", ""), vbInformation
                Else
                   If frmRegister.Visible Then
                      frmRegister.Hide
                   Else
                      Set X = Nothing
                      Set X = Parent.UI
                          If X Is Nothing Then
                             frmRegister.Show
                          Else
                             frmRegister.Show 0, X
                          End If
                          If Not frmRegister.Visible Then
                             frmRegister.Show
                          End If
                      Set X = Nothing
                   End If
                End If

           Case "UPDATE FILES"
                 If frmRegister.GetCentralUpdateInfo Then
                    MsgBox "Success !" & Chr(13) & Chr(13) & "Please note, no product update has been done. This feature is still beta." & Chr(13) & Chr(9) & "Please check http://www.SandySupport.com for an update."
                    sMessage = sadGetLicenseKey("Message")
                    If Len(sMessage) Then
                       sMessage = Replace(Replace(sMessage, "$%$EOL$%$", Chr(13)), "$%$TAB$%$", Chr(9))
                       MsgBox "A message was received from the Central Server as follows:" & Chr(13) & Chr(13) & sMessage
                    End If
                 End If

           Case "SUBMIT TEMPLATE"
                 frmRegister.SubmitTemplate

           Case "HIDE ALL WINDOWS", "HIDEALLWINDOWS"
                If frmRegister.Visible Then frmRegister.Hide

           Case "UNLOAD"
                Unload frmRegister
    End Select
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommand

    With II
         If Not MySadCommands(.SoftCommandName & "*C") Is Nothing Then
            If MySadCommands(.SoftCommandName & "*C").IsInline Then Exit Function

On Error Resume Next
            Select Case UCase(.SoftCommandName)
                   Case "URL", "GETWEBPAGE", "GETPAGE", "FETCHURL", "GETURL"
                         .SoftVars(.Result) = frmRegister.GetURL(.Expression)
                         ISadAddin_ExecuteSoftCommand = True
                   Case "GETFILEURL", "FILEURL", "GETFILEFROMPAGE", "GETREMOTEFILE", "FETCHFILE"
                         Call frmRegister.GetFile(.Expression, .Result)
                         ISadAddin_ExecuteSoftCommand = True
            End Select
         End If
    End With

EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:
    Exit Function

EH_SADAddin_ISadAddin_ExecuteSoftCommand:
    LogError "SADAddin", "ISadAddin_ExecuteSoftCommand", Err.Number, Err.Description
    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue

    Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
    '
End Function

Private Property Get ISadAddin_Externals() As SandySupport.CAssocArray
On Error Resume Next
    Set ISadAddin_Externals = Externals
End Property

Private Function ISadAddin_Shutdown() As Boolean
On Error Resume Next
    MySadCommands.Clear
    Set MySadCommands.Parent = Nothing
    Set MySadCommands = Nothing
    Set Parent = Nothing

    ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oSliceAndDiceWizard As Object, Application As Object) As Boolean
On Error Resume Next
    Err.Clear

    Set Parent = oSliceAndDiceWizard
    Set MySadCommands.Parent = oSliceAndDiceWizard

    ISadAddin_Startup = (Err.Number = 0)
End Function
