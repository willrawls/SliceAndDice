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

Public Parent As SliceAndDice.Wizard
Public MySadCommands As SliceAndDice.CSadCommands
Public Externals As SliceAndDice.CAssocArray


Public Sub BrowseTo(ByVal sURL As String)
    Dim X As Form
    If frmBrowser.Visible Then
       frmBrowser.Hide
    Else
       Set X = Nothing
       Set X = Parent.SandyWindow
           frmBrowser.StartingAddress = sURL
           frmBrowser.brwWebBrowser.Navigate frmBrowser.StartingAddress
           If X Is Nothing Then
              frmBrowser.Show 0
           Else
              frmBrowser.Show 0, X
           End If
           If Not frmBrowser.Visible Then
              frmBrowser.Show
           End If
       Set X = Nothing
    End If
    DoEvents: DoEvents: DoEvents
End Sub

Private Sub Class_Initialize()
On Error Resume Next
    Set MySadCommands = New SliceAndDice.CSadCommands
    Set Externals = New SliceAndDice.CAssocArray
    With MySadCommands
         .Clear
         .ParameterDelimiter = ","
         .ParameterTypeDelimiter = ":"
         .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision
'         With .Add("GetURL", False)
'              .Aliases = "FileURL, GetWebPage, GetPage, FetchURL"
'              .Comments = "Timeout defaults to 120 seconds"
'              .OneLineDescription = "Fetches the contents of the specified URL to the indicated SoftVarOut"
'              .Syntax = "SoftVarOut:SoftVar = URL:String"
'         End With
'         With .Add("GetFileURL", False)
'              .Aliases = "URL, GetFileFromPage, GetRemoteFile, FetchFile"
'              .Comments = "Timeout defaults to 120 seconds"
'              .OneLineDescription = "Fetches a binary file from the specified URL to the indicated filename"
'              .Syntax = "FilenameOut:String = URL:String"
'         End With
    End With

    With Externals
         .Clear
         .Item("&Import Template(s) from vbcode.com") = "Import Templates"
         .Item("&Submit Current Template to vbcode.com") = "Submit Template"
    End With
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Set Externals = Nothing
    Set MySadCommands = Nothing
    Set frmBrowser.Parent = Nothing
    Unload frmBrowser
End Sub

Private Property Get ISadAddin_CommandSet() As SliceAndDice.CSadCommands
On Error Resume Next
    Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
    Dim sMessage As String
    Dim sAuthorName As String
    Dim sAuthorEmail As String

On Error Resume Next
    If frmBrowser.Parent Is Nothing Then
       Set frmBrowser.Parent = Me
    End If
    
    Select Case UCase(sValue)
           Case "IMPORT TEMPLATES"
                 BrowseTo "http://www.vbcode.com"
                 frmBrowser.ZOrder

           Case "SUBMIT TEMPLATE"
                 BrowseTo "http://www.vbcode.com/submit.htm"
                 Do Until frmBrowser.NavigationComplete
                    DoEvents
                 Loop

                 sAuthorName = GetSetting(App.ProductName, "Last", "AuthorName", "")
                 'If Len(sAuthorName) = 0 Then
                    sAuthorName = InputBox("What is the Template Author's Name ?", "SUBMIT TEMPLATE TO VBCODE.COM", sAuthorName)
                 'End If
                 SaveSetting App.ProductName, "Last", "AuthorName", sAuthorName

                 sAuthorEmail = GetSetting(App.ProductName, "Last", "AuthorEmail", "")
                 'If Len(sAuthorEmail) = 0 Then
                    sAuthorEmail = InputBox("What is the Template Author's Email ?", "SUBMIT TEMPLATE TO VBCODE.COM", sAuthorEmail)
                 'End If
                 SaveSetting App.ProductName, "Last", "AuthorEmail", sAuthorEmail

                 frmBrowser.brwWebBrowser.Document.Forms(0).Item("AuthorName").Value = sAuthorName
                 frmBrowser.brwWebBrowser.Document.Forms(0).Item("AuthorEmail").Value = sAuthorEmail
                 frmBrowser.brwWebBrowser.Document.Forms(0).Item("Task").Value = Parent.CurrentTemplate.Key
                 If Len(Parent.CurrentTemplate.memoCodeAtTop) > 0 Then
                    frmBrowser.brwWebBrowser.Document.Forms(0).Item("Declarations").Value = "~~' Submitted from Slice and Dice" & vbNewLine & Parent.CurrentTemplate.memoCodeAtTop
                    frmBrowser.brwWebBrowser.Document.Forms(0).Item("CodeSnippet").Value = Parent.CurrentTemplate.memoCodeAtBottom
                 ElseIf Len(Parent.CurrentTemplate.memoCodeAtBottom) > 0 Then
                    frmBrowser.brwWebBrowser.Document.Forms(0).Item("CodeSnippet").Value = "~~' Submitted from Slice and Dice" & vbNewLine & Parent.CurrentTemplate.memoCodeAtBottom
                 End If
                 frmBrowser.ZOrder

           Case "HIDE ALL WINDOWS", "HIDEALLWINDOWS"
                If frmBrowser.Visible Then frmBrowser.Hide

           Case "UNLOAD"
                Unload frmBrowser
    End Select
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommand

'    With II
'         If Not MySadCommands(.SoftCommandName & "*C") Is Nothing Then
'            If MySadCommands(.SoftCommandName & "*C").IsInline Then Exit Function
'
'On Error Resume Next
'            Select Case UCase(.SoftCommandName)
'                   Case "URL", "GETWEBPAGE", "GETPAGE", "FETCHURL", "GETURL"
'                         .SoftVars(.Result) = frmBrowser.brwWebBrowser.nav.Expression)
'                         ISadAddin_ExecuteSoftCommand = True
'                   Case "GETFILEURL", "FILEURL", "GETFILEFROMPAGE", "GETREMOTEFILE", "FETCHFILE"
'                         Call frmBrowser.GetFile(.Expression, .Result)
'                         ISadAddin_ExecuteSoftCommand = True
'            End Select
'         End If
'    End With

EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:
    Exit Function

EH_SADAddin_ISadAddin_ExecuteSoftCommand:
    Parent.LogError "SADAddin", "ISadAddin_ExecuteSoftCommand", Err.Number, Err.Description
    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue

    Resume
End Function

Private Function ISadAddin_ExecuteSoftCommandInline(ByRef SoftVars As CAssocArray, ByVal sInlineSoftCommandName As String, ByVal sParameters As String, sResults As String) As Boolean
    '
End Function

Private Property Get ISadAddin_Externals() As SliceAndDice.CAssocArray
On Error Resume Next
    Set ISadAddin_Externals = Externals
End Property

Private Function ISadAddin_Shutdown() As Boolean
On Error Resume Next
    MySadCommands.Clear
    Set MySadCommands.Parent = Nothing
    Set MySadCommands = Nothing
    Set Parent = Nothing

    If Not frmBrowser Is Nothing Then
       frmBrowser.timTimer.Enabled = False
       Unload frmBrowser
    End If

    ISadAddin_Shutdown = True
End Function

Private Function ISadAddin_Startup(oParent As SliceAndDice.Wizard, vbInst As VBIDE.VBE) As Boolean
On Error Resume Next
    Err.Clear

    Set Parent = oParent
    Set MySadCommands.Parent = oParent

    ISadAddin_Startup = (Err.Number = 0)
End Function

