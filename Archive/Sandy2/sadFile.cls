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
Attribute VB_Description = "Sandle - File System Commands"
Option Explicit

Implements SandySupport.ISadAddin

Private Parent As SandySupport.ISandyWizard
Private MySadCommands As SandySupport.CSadCommands

Private fh As Long
Private Sub Class_Initialize()
On Error Resume Next
    Set MySadCommands = New SandySupport.CSadCommands
    With MySadCommands
         .ParameterDelimiter = ","
         .ParameterTypeDelimiter = ":"
         .Attributes("Version") = App.Major & "." & App.Minor & "." & App.Revision
    End With
End Sub

Private Sub Class_Terminate()
    Set MySadCommands = Nothing
End Sub

Private Property Get ISadAddin_CommandSet() As Object
On Error Resume Next
    Set ISadAddin_CommandSet = MySadCommands
End Property

Private Function ISadAddin_ExecuteExternal(ByVal sKey As String, ByVal sValue As String) As Boolean
'    Select Case UCase(sKey)
'           Case "TESTING EXTERNALS"
'                MsgBox sValue
'    End Select
End Function

Private Function ISadAddin_ExecuteSoftCommand(II As CInsertionInfo) As Boolean
On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommand
    Dim sVar1 As String
    Dim sVar2 As String
    Dim sVar3 As String

    With II
         If Not MySadCommands(.SoftCommandName & "*C") Is Nothing Then
            If MySadCommands(.SoftCommandName & "*C").IsInline Then Exit Function
            sVar1 = sGetToken(II.AllParameters, 1, ","): If Len(.SoftVars(sVar1)) Then sVar1 = .SoftVars(sVar1)
            sVar2 = sGetToken(II.AllParameters, 2, ","): If Len(.SoftVars(sVar2)) Then sVar2 = .SoftVars(sVar2)
            sVar3 = sGetToken(II.AllParameters, 3, ","): If Len(.SoftVars(sVar3)) Then sVar3 = .SoftVars(sVar3)

On Error Resume Next
            Select Case UCase(.SoftCommandName)
                   Case "CHANGEDRIVE", "DRIVE", "CHDRIVE"
                         ChDrive sVar1
                         ISadAddin_ExecuteSoftCommand = True

                   Case "CHANGEDIRECTORY", "CHANGEDIR", "CHDIR", "CD"
                         ChDir sVar1
                         ISadAddin_ExecuteSoftCommand = True
                        
                   Case "CREATEFILE", "MAKEFILE", "MKFILE"
                         fh = FreeFile
                         Open sVar1 For Output Access Write As #fh
                              If Len(sVar2) > 0 Then
                                 Print #fh, sVar2
                              End If
                         Close #fh
                         fh = 0
                         ISadAddin_ExecuteSoftCommand = True

                   Case "APPENDTOFILE", "APPENDFILE"
                         fh = FreeFile
                         Open sVar1 For Append Access Write As #fh
                              If Len(sVar2) > 0 Then
                                 Print #fh, sVar2
                              End If
                         Close #fh
                         fh = 0
                         ISadAddin_ExecuteSoftCommand = True

                   Case "DELETEFILE", "DELFILE", "DEL", vbNullString
                         Kill sVar1
                         ISadAddin_ExecuteSoftCommand = True

                   Case "CREATEFOLDER", "CREATEDIRECTORY", "MAKEDIR", "MKDIR", "MD"
                         MkDir sVar1
                         ISadAddin_ExecuteSoftCommand = True

                   Case "DELETEFOLDER", "DELETEDIRECTORY", "DELDIR", "DELETEDIR"
                         RmDir sVar1
                         ISadAddin_ExecuteSoftCommand = True

                   Case "RENAME", "REN", "RENAMEFILE"
                         Name sVar1 As sVar2
                         ISadAddin_ExecuteSoftCommand = True

                   Case "FILECONTENTS", "GETFILE", "FILETOSOFTVAR", "FILETOVAR"
                         II.SoftVars(II.Result) = sFileContents(II.Expression)
                         ISadAddin_ExecuteSoftCommand = True

                   Case "VARTOFILE", "PUTFILE", "SAVETOFILE"
                         fh = FreeFile
                         Open II.Result For Output Access Write As #fh
                              Print #fh, II.SoftVars(II.Expression)
                         Close #fh
                         fh = 0
                         ISadAddin_ExecuteSoftCommand = True

                   Case "APPENDVARTOFILE", "APPENDVAR"
                         fh = FreeFile
                         Open II.Result For Output Access Write As #fh
                              Print #fh, II.SoftVars(II.Expression)
                         Close #fh
                         fh = 0
                         ISadAddin_ExecuteSoftCommand = True

                   Case "RUNPROGRAM", "SHELL", "EXECUTE"
                         II.SoftVars("SHELLID") = Shell(sVar1, vbNormalFocus)
                         ISadAddin_ExecuteSoftCommand = True

                   Case "ACTIVATE", "APPACTIVATE", "ACTIVATEAPPLICATION", "SWITCHTOAPP"
                         AppActivate sVar1
                         ISadAddin_ExecuteSoftCommand = True

                   Case "SENDKEYS"
                         SendKeys sVar1
                         ISadAddin_ExecuteSoftCommand = True

                   Case "SENDKEYSWITHPAUSE", "SENDKEYSANDWAIT", "SENDKEYSANDPAUSE", "SENDKEYSWITHWAIT"
                         SendKeys sVar1, True
                         ISadAddin_ExecuteSoftCommand = True

                   Case "CHECKFORFILE", "PAUSEFORFILE", "WAITFORFILE", "WAITFOR"
                         ISadAddin_ExecuteSoftCommand = True

                   Case "PAUSE", "WAIT"
                         'If Val(II.AllParameters) > 0 Then
                         '   Pause
                         'End If
                         ISadAddin_ExecuteSoftCommand = True

                   Case "OPENFORINPUT", "OPENFILE"
                         fh = FreeFile
                         Open sVar1 For Input Access Read As #fh
                         ISadAddin_ExecuteSoftCommand = True

                   Case "OPENFOROUTPUT"
                         Open sVar1 For Output Access Write As #fh
                         ISadAddin_ExecuteSoftCommand = True

                   Case "NEXTLINE", "LINEINPUTFROMFILE", "READLINEFROMFILE", "READFILELINE", "FILEREADLINE"
                         Line Input #fh, sVar3
                         II.SoftVars(sVar1) = sVar3
                         ISadAddin_ExecuteSoftCommand = True

                   Case "CLOSEFILE"
                         If fh <> 0 Then
                            Close #fh
                            fh = 0
                         End If
                         ISadAddin_ExecuteSoftCommand = True

                   Case "COPYFILE", "FILECOPY"
                         FileCopy II.Expression, II.Result
                         ISadAddin_ExecuteSoftCommand = True

                   Case "SAVESETTING", "SAVEREGISTRYSETTING", "SETREGKEY"
                         SaveSetting "Slice and Dice", sVar1, sVar2, sVar3
                         ISadAddin_ExecuteSoftCommand = True

                   Case "SAVESETTING", "SAVEREGISTRYSETTING", "SETREGKEY"
                         .SoftVars(sVar3) = GetSetting("Slice and Dice", sVar1, sVar2, vbNullString)
                         ISadAddin_ExecuteSoftCommand = True

                   Case "DELETESETTING", "DELETEREGISTRYSETTING", "DELETEREGKEY", "DELREGKEY"
                         DeleteSetting "Slice and Dice", sVar1, sVar2
                         ISadAddin_ExecuteSoftCommand = True

'                   Case "GOTO", "GOSUB", "RETURNFROMGOSUB", "RETURN"
'                         ISadAddin_ExecuteSoftCommand = True
'
'                   Case "IFTHEN", "IF", "OTHERWISE", "ELSEIF", "ENDIF"
'                         ISadAddin_ExecuteSoftCommand = True
'
'                   Case "SELECTCASE", "SWITCHCASE", "SWITCH"
'                         ISadAddin_ExecuteSoftCommand = True
'
'                   Case "CREATETREE", "CREATEDIRECTORYTREE"
'                         ISadAddin_ExecuteSoftCommand = True
'
'                   Case "DELETETREE", "DELETEDIRECTORYTREE"
'                         ISadAddin_ExecuteSoftCommand = True
'
'                   Case "MODIFYTREE", "MODIFYDIRECTORYTREE"
'                         ISadAddin_ExecuteSoftCommand = True
'
'                   Case "FOREACHFILE", "STARTEACHFILE", "ENDEACHFILE"
'                        sResults = vbNullString
'                        ISadAddin_ExecuteSoftCommandInline = True
'
'                   Case "FOREACHDIRECTORY", "STARTEACHDIRECTORY", "ENDEACHDIRECTORY"
'                        sResults = vbNullString
'                        ISadAddin_ExecuteSoftCommandInline = True
'
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
On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline
    Dim Template As CTemplate

    Dim Area As Long
    Dim CurrSet As Long
    Dim bInlineCommandExecuted As Boolean
    Dim lParameterCount As Long

    Dim sDefault As String
    Dim sT As String
    Dim sVar() As String
    Dim asaList As New CAssocArray
    Dim CurrParam As Long
    Dim ParameterCount As Long

    If Not MySadCommands(sInlineSoftCommandName & "*I") Is Nothing Then
       If Not MySadCommands(sInlineSoftCommandName & "*I").IsInline Then Exit Function
       ParameterCount = lTokenCount(sParameters, ",")
       If ParameterCount > 0 Then
          If ParameterCount < 5 Then
             ReDim sVar(1 To 5) As String
          Else
             ReDim sVar(1 To ParameterCount) As String
          End If
          For CurrParam = 1 To ParameterCount
              sVar(CurrParam) = sGetToken(sParameters, CurrParam, ",")
              If Len(SoftVars(sVar(CurrParam))) Then sVar(CurrParam) = SoftVars(sVar(CurrParam))
          Next CurrParam
       Else
          ReDim sVar(1 To 5) As String
       End If

       Select Case sInlineSoftCommandName
              Case "CHOOSEFILE", "PICKFILE", "GETPATHANDFILE"
                   Select Case ParameterCount
                          Case Is < 2: sResults = Parent.sChooseFile(sVar(1))
                          Case 2:    sResults = Parent.sChooseFile(sVar(1), sVar(2))
                          Case Else: sResults = Parent.sChooseFile(sVar(1), sVar(2), sVar(3))
                   End Select
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "CHOOSEDATABASE", "CHOOSEDB", "PICKDATABASE", "PICKDB", "GETDBPATHANDFILE"
                   Select Case ParameterCount
                          Case Is < 2: sResults = Parent.sChooseDatabase(sVar(1))
                          Case Else: sResults = Parent.sChooseDatabase(sVar(1), sVar(2))
                   End Select
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "CHOOSECOLOR", "PICKCOLOR"
                   Select Case ParameterCount
                          Case Is > 0: sResults = Parent.sChooseColor(sVar(1))
                          Case Else: sResults = Parent.sChooseColor
                   End Select
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "READTREE", "READDIRECTORYTREE", "TREETOSOFTVAR"
                   
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "FILEEXISTS", "EXISTENCE", "DOESFILEEXIST"  ' inline
                   
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "MATCHFILE", "DIR", "FINDFIRST", "FIRSTFILE"         'inline
                   sResults = Dir(sVar(1), Val(sVar(2)))
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "MATCHNEXTFILE", "FINDNEXT", "NEXTFILE"             'inline
                   sResults = Dir
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "GETSETTING", "GETREGISTRYSETTING", "GETREGKEY"         'inline
                   sResults = GetSetting("Slice and Dice", sVar(1), sVar(2), sVar(3))
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "FILECONTENTS", "CONTENTS", "INCLUDEFILE"
                   sResults = sFileContents(sParameters)
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "LISTDIRS", "LISTDIRECTORIES", "DIRECTORIES", "DIRS", "ENUMDIRS"
                   sResults = EnumFiles(sVar(1), sVar(2), "DIRECTORY")
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "LISTFILES", "FILES", "ENUMFILES"
                   sResults = EnumFiles(sVar(1), sVar(2), sVar(3))
                   ISadAddin_ExecuteSoftCommandInline = True

              Case "TESTEOF", "EOF", "ENDOFFILE", "ISENDOFFILE", "TESTENDOFFILE"
                   If fh = 0 Then
                      sResults = "0"
                   Else
                      sResults = EOF(fh)
                   End If
                   ISadAddin_ExecuteSoftCommandInline = True
        End Select
    End If

EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue:
    Exit Function

EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline:
    LogError "SADAddin", "ISadAddin_ExecuteSoftCommandInline", Err.Number, Err.Description
    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue

    Resume
End Function

Private Property Get ISadAddin_Externals() As SandySupport.CAssocArray
'    Dim Externals As CAssocArray
'    Set Externals = New CAssocArray
'    Externals.All = "Testing Externals=Testing 123"
'    Set ISadAddin_Externals = Externals
'    Set Externals = Nothing
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

    Err.Clear
        MySadCommands.All = sFileContents(Parent.TemplateDatabasePath & "s2kFile.txt")
        ISadAddin_Startup = (Err.Number = 0)
    Err.Clear
End Function
