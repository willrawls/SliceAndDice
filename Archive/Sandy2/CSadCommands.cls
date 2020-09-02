VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSadCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy - Soft Command Collection"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CSadCommand"
Attribute VB_Ext_KEY = "Member0" ,"CSadCommand"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public ParameterDelimiter As String
Public ParameterTypeDelimiter As String


Public Parent As SandySupport.ISandyWizard
Private MyHelpScreen As SandySupport.ISandyWindowHelp

Private masaParams As CAssocArray
Private masaSyntax As CAssocArray
Private masaAttributes As CAssocArray

Public Enum SadCommandSectionType
     InNone = 0
     InSyntax
     InLong
     InComments
     InSeeAlso
     InExample
End Enum

Public Sub Append(NewItem As Object)
On Error Resume Next
    If TypeOf NewItem Is CSadCommand Then
       mCol.Add NewItem, UCase$(NewItem.SoftCommandName) & IIf(NewItem.IsInline, "*I", "*C")
    End If
End Sub


Public Property Get Attributes() As CAssocArray
    Set Attributes = masaAttributes
End Property

Public Function CheckSyntax(SoftCommandName As String) As Boolean
On Error GoTo EH_CSadCommands_DetermineIfSyntaxIsValid
    Static CurrSyntaxItem As CAssocItem
    Static RequiredParams As Long
   'Dim CurrParamItem As CAssocItem
   'Dim ParamType As String
   
    If Not Item(SoftCommandName) Is Nothing Then
       With Item(SoftCommandName)
            masaSyntax.OrderedMultiAdd .Syntax
            RequiredParams = 0
            For Each CurrSyntaxItem In masaSyntax.mCol
                If InStr(UCase$(sGetToken(CurrSyntaxItem.Value, 2, ParameterTypeDelimiter)), "OPT") = 0 Then
                   RequiredParams = RequiredParams + 1
                End If
            Next CurrSyntaxItem
            masaParams.OrderedMultiAdd .Parameters
            CheckSyntax = (masaParams.Count >= RequiredParams)

           'mbSyntaxIsValid = True
           'For Each CurrSyntaxItem In masaSyntax.mCol
               'Set CurrParamItem = masaParams(CurrSyntaxItem.Key)
               'ParamType = sAfter(CurrSyntaxItem.Value, 1, Parent.ParameterTypeDelimiter)
               'Select Case UCase$(sGetToken(ParamType))
               '       Case "SOFTVAR"
               '       Case "STRINGORSOFTVAR"
               '       Case "STRING"
               '       Case "QUOTEDSTRING"
               '       Case "LONG"
               '       Case vbNullString
               'End Select
           'Next CurrSyntaxItem
       End With
    End If

EH_CSadCommands_DetermineIfSyntaxIsValid_Continue:
    Set CurrSyntaxItem = Nothing
    Exit Function

EH_CSadCommands_DetermineIfSyntaxIsValid:
    LogError "CSadCommands", "DetermineIfSyntaxIsValid", Err.Number, Err.Description
    Resume EH_CSadCommands_DetermineIfSyntaxIsValid_Continue

    Resume
End Function


Public Function Add(SoftCommandName As String, Optional bInLine As Boolean = False) As CSadCommand
On Error GoTo EH_CSadCommands_Add
    Dim sKey As String
    Dim CurrMember As CSadCommand
    Dim objNewMember As CSadCommand

    Set objNewMember = CreateObject("SandySupport.CSadCommand")
    With objNewMember
         .SoftCommandName = SoftCommandName
         .IsInline = bInLine
    End With
    
  ' Fit key into current list in alphabetic order
    sKey = UCase$(SoftCommandName) & IIf(bInLine, "*I", "*C")
    If mCol.Count = 0 Then
       mCol.Add objNewMember, sKey
    Else
       For Each CurrMember In mCol
           If StrComp(sKey, UCase$(CurrMember.SoftCommandName) & IIf(CurrMember.IsInline, "*I", "*C")) < 0 Then
              mCol.Add objNewMember, sKey, UCase$(CurrMember.SoftCommandName) & IIf(CurrMember.IsInline, "*I", "*C")
              Set Add = objNewMember
              GoTo EH_CSadCommands_Add_Continue
           End If
       Next CurrMember

     ' Not found, Insert at the end of the collection
       mCol.Add objNewMember, sKey
    End If
    Set Add = objNewMember

EH_CSadCommands_Add_Continue:
    Set objNewMember = Nothing
    Exit Function

EH_CSadCommands_Add:
    LogError "CSadCommands", "Add", Err.Number, "SoftCommandName=" & SoftCommandName & gsEolTab & Err.Description
    Resume EH_CSadCommands_Add_Continue

    Resume
End Function

Public Property Let All(ByVal NewData As String)
On Error Resume Next
    Dim CurrLine As String
    Dim CurrSection As SadCommandSectionType
    Dim CurrCommand As CSadCommand
    Dim CurrAlias As CSadCommand
    Dim CurrIndex As Long

    Clear
    Do While Len(NewData)
       CurrLine = sExtractToken(NewData, 1, vbNewLine)
       If Left$(CurrLine, 3) = "** " Then
          CurrLine = Mid$(CurrLine, 4)
          Select Case UCase$(Trim$(sGetToken(CurrLine)))
                 Case "'", "REM"   ' Comment
                      CurrSection = InNone

                 Case "COMMAND", "SOFTCOMMAND"
                      Set CurrCommand = Add(Trim$(sAfter(CurrLine)), False)
                      CurrSection = InNone

                 Case "INLINE", "INLINECOMMAND", "INLINESOFTCOMMAND"
                      Set CurrCommand = Add(Trim$(sAfter(CurrLine)), True)
                      CurrSection = InNone

                 Case "ALIAS", "ALIASES"
                      If Not CurrCommand Is Nothing Then CurrCommand.Aliases = sAfter(CurrLine)
                      If Len(CurrCommand.Aliases) Then
                         CurrCommand.Aliases = Replace(Replace(CurrCommand.Aliases, ",", ", "), ",  ", ", ")
                         If Right$(CurrCommand.Aliases, 2) <> ", " Then
                            CurrCommand.Aliases = CurrCommand.Aliases & ", "
                         End If
                         If Left$(CurrCommand.Aliases, 2) <> ", " Then
                            CurrCommand.Aliases = ", " & CurrCommand.Aliases
                         End If
                      End If
                      CurrSection = InNone

                 Case "SHORT", "ONELINE", "SHORTDESCRIPTION", "SHORTDESC", "ONE"
                      If Not CurrCommand Is Nothing Then CurrCommand.OneLineDescription = sAfter(CurrLine)
                      CurrSection = InNone

                 Case "SYNTAX"
                      CurrSection = InSyntax
                      If Not CurrCommand Is Nothing Then CurrCommand.Syntax = sAfter(CurrLine)

                 Case "LONG", "LONGDESCRIPTION", "LONGDESC"
                      CurrSection = InLong
                      If Not CurrCommand Is Nothing Then CurrCommand.LongDescription = sAfter(CurrLine)

                 Case "COMMENTS", "REMARKS", "NOTES"
                      CurrSection = InComments
                      If Not CurrCommand Is Nothing Then CurrCommand.Comments = sAfter(CurrLine)

                 Case "SEEALSO", "SEE"
                      CurrSection = InSeeAlso
                      If Not CurrCommand Is Nothing Then CurrCommand.SeeAlso = sAfter(CurrLine)

                 Case "EXAMPLE", "EXAMPLES"
                      CurrSection = InExample
                      If Not CurrCommand Is Nothing Then CurrCommand.Examples = sAfter(CurrLine)

                 Case "HELPFILE", "HELP"
                      If Not CurrCommand Is Nothing Then CurrCommand.HelpFile = sAfter(CurrLine)
                      CurrSection = InNone

                 Case "HELP TOPIC", "TOPIC"
                      If Not CurrCommand Is Nothing Then CurrCommand.HelpTopic = sAfter(CurrLine)
                      CurrSection = InNone

                 Case "END"
                      NewData = vbNullString
                      CurrSection = InNone

                 Case Else
                      CurrSection = InNone
          End Select
       Else
          Select Case CurrSection
                 Case InNone
                 Case InSyntax:     If Not CurrCommand Is Nothing Then CurrCommand.Syntax = CurrCommand.Syntax & CurrLine & vbNewLine
                 Case InLong:       If Not CurrCommand Is Nothing Then CurrCommand.LongDescription = CurrCommand.LongDescription & CurrLine & vbNewLine
                 Case InComments:   If Not CurrCommand Is Nothing Then CurrCommand.Comments = CurrCommand.Comments & CurrLine & vbNewLine
                 Case InSeeAlso:    If Not CurrCommand Is Nothing Then CurrCommand.SeeAlso = CurrCommand.SeeAlso & CurrLine & vbNewLine
                 Case InExample:    If Not CurrCommand Is Nothing Then CurrCommand.Examples = CurrCommand.Examples & CurrLine & vbNewLine
          End Select
       End If
    Loop
    
    CurrIndex = 1
    For Each CurrCommand In mCol
        CurrCommand.Index = CurrIndex
        CurrIndex = CurrIndex + 1
    Next CurrCommand
End Property

Public Sub Clear(Optional ByVal bRecreateCollection As Boolean = True)
On Error Resume Next
    Set mCol = Nothing
    If bRecreateCollection Then
       Set mCol = New Collection
       ' mCol.CompareMode = TextCompare
    End If
End Sub

Public Sub HideWindow(Optional ByVal bUnloadAsWell As Boolean = False)
    If Not MyHelpScreen Is Nothing Then
       MyHelpScreen.Hide
       If bUnloadAsWell Then
          Unload MyHelpScreen
          Set MyHelpScreen = Nothing
       End If
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As CSadCommand
Attribute Item.VB_UserMemId = 0
On Error Resume Next
    Dim TempItem As CSadCommand
    Dim TempKey As String
    
    If VarType(vntIndexKey) = vbString Then
       Err.Clear
       Set TempItem = Nothing
       Set TempItem = mCol(UCase$(vntIndexKey))
       If TempItem Is Nothing Then
          Set TempItem = mCol(UCase$(vntIndexKey) & "*C")
          If TempItem Is Nothing Then
             Set TempItem = mCol(UCase$(vntIndexKey) & "*I")
             If Not TempItem Is Nothing Then
                Set Item = TempItem
             End If
          Else
             Set Item = TempItem
          End If
       Else
          Set Item = TempItem
       End If
       If TempItem Is Nothing Then
        ' Search Aliases
          If InStr(vntIndexKey, "*I") Then
             TempKey = sGetToken(vntIndexKey, 1, "*I")
          ElseIf InStr(vntIndexKey, "*C") Then
             TempKey = sGetToken(vntIndexKey, 1, "*C")
          Else
             TempKey = vntIndexKey
          End If
          TempKey = ", " & TempKey & ", "
          For Each TempItem In mCol
              If InStr(UCase$(TempItem.Aliases), UCase$(TempKey)) Then
                 Set Item = TempItem
                 Exit For
              End If
          Next TempItem
       End If
    Else
       Set Item = mCol(vntIndexKey)
    End If

    Set TempItem = Nothing

End Property

Public Property Get Count() As Long
On Error Resume Next
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As String)
On Error Resume Next
    mCol.Remove UCase$(vntIndexKey)
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
On Error Resume Next
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Sub ShowHelpScreen(Optional ByVal sSoftCommand As String)
On Error Resume Next
    With MyHelpScreen
         Set .SadCommandSet = Me
         If Len(sSoftCommand) Then
            If Not Item(sSoftCommand) Is Nothing Then
               .CurrCommandKey = Item(sSoftCommand).Index
            End If
         Else
            .CurrCommandKey = 1
         End If
         MsgBox "do fore and back color"
'         .txtAliases.BackColor = "&H" & GetSetting(App.ProductName, "Last", "Background Color", "&H80000018&")
'         .txtSyntax.BackColor = .txtAliases.BackColor
'         .txtOneLineDescription.BackColor = .txtAliases.BackColor
'         .txtHelpFile.BackColor = .txtAliases.BackColor
'         .txtHelpTopic.BackColor = .txtAliases.BackColor
'         .txtLongDescription.BackColor = .txtAliases.BackColor
'         .txtComments.BackColor = .txtAliases.BackColor
'         .txtSeeAlso.BackColor = .txtAliases.BackColor
'         .txtExamples.BackColor = .txtAliases.BackColor
'
'         .txtAliases.ForeColor = "&H" & GetSetting(App.ProductName, "Last", "Foreground Color", "&H80000008&")
'         .txtSyntax.ForeColor = .txtAliases.ForeColor
'         .txtOneLineDescription.ForeColor = .txtAliases.ForeColor
'         .txtHelpFile.ForeColor = .txtAliases.ForeColor
'         .txtHelpTopic.ForeColor = .txtAliases.ForeColor
'         .txtLongDescription.ForeColor = .txtAliases.ForeColor
'         .txtComments.ForeColor = .txtAliases.ForeColor
'         .txtSeeAlso.ForeColor = .txtAliases.ForeColor
'         .txtExamples.ForeColor = .txtAliases.ForeColor
         .Show 0, Parent.UI
    End With
End Sub

Private Sub Class_Initialize()
On Error Resume Next
    Set mCol = New Collection
        ' mCol.CompareMode = TextCompare
        ParameterDelimiter = ","
        ParameterTypeDelimiter = ":"
    
    Set masaParams = CreateObject("SandySupport.CAssocArray")
        masaParams.ItemDelimiter = ParameterDelimiter
        masaParams.FieldDelimiter = ParameterTypeDelimiter
    
    Set masaSyntax = CreateObject("SandySupport.CAssocArray")
        masaSyntax.ItemDelimiter = ParameterDelimiter
        masaSyntax.FieldDelimiter = ParameterTypeDelimiter

    Set masaAttributes = CreateObject("SandySupport.CAssocArray")

    Set MyHelpScreen = CreateObject("SandySupport.frmCommandHelp")
    With MyHelpScreen
        Set .SadCommandSet = Me
        .CurrCommandKey = 1
    End With
    ' LogEvent "CSadCommands: Initialize"
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    With MyHelpScreen
         .FileExit
         Set .SadCommandSet = Nothing
         Set .CurrCommand = Nothing
    End With
    Unload MyHelpScreen
    Set MyHelpScreen = Nothing

    Set masaParams = Nothing
    Set masaSyntax = Nothing
    Set masaAttributes = Nothing
    Set mCol = Nothing
    ' LogEvent "CSadCommands: Terminate"
End Sub

