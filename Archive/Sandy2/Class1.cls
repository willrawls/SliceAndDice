VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private SandyIDE As CSandyIDE

Implements IDTExtensibility

'Private WithEvents mySubClass As CSubClassIDE

'Private WithEvents SubClassIDE As SubClassIDE.CSubClassIDE

Public Function EnumFiles(sPath As String, sMask As String, sAttr As String) As String
    Dim Attr As VbFileAttribute
    
    Select Case Attr 'UCase(sAttr)
           Case "ALIAS":        Attr = vbAlias
           Case "ARCHIVE":      Attr = vbArchive
           Case "DIRECTORY":    Attr = vbDirectory
           Case "HIDDEN":       Attr = vbHidden
           Case "READONLY":     Attr = vbReadOnly
           Case "SYSTEM":       Attr = vbSystem
           Case "VOLUME":       Attr = vbVolume
           Case Else:           Attr = vbNormal
    End Select
    EnumFiles = modGeneral.EnumFiles(sPath, sMask, Attr)
End Function

Public Function Evaluate(ByVal sExpression As String, ByRef asaVar As CAssocArray)
    Evaluate = m_oUI.Evaluate(sExpression, asaVar)
End Function


Public Function FileExists(sFilename As String) As Boolean
    FileExists = modGeneral.FileExists(sFilename)
End Function

Public Sub HandleKeyPress(ByRef KeyCode As Integer, ByRef Shift As Integer)
    If Shift = 3 Then '(Shift And vbShiftMask) And (Shift And vbCtrlMask) Then
       Select Case KeyCode
              Case 69  ' "E"xternals window
                   ExternalsHandler_Click Nothing, False, False
                   KeyCode = 0
                   Shift = 0
              Case 70  ' "F"avorites
                   FavoritesHandler_Click Nothing, False, False
                   KeyCode = 0
                   Shift = 0
              Case 83  ' "S"lice and Dice window
                   MenuHandler_Click Nothing, False, False
                   KeyCode = 0
                   Shift = 0
              Case Else
                   MsgBox "Combination key Shift-Ctrl-" & KeyCode & " pressed"
                   KeyCode = 0
                   Shift = 0
       End Select
    End If
End Sub

Public Function JumpTo(ByVal sTemplateName As String, Optional ByVal bRecordInHistory As Boolean = True, Optional ByVal bSyncCategoryList As Boolean = False) As Boolean
    m_oUI.JumpTo sTemplateName, bRecordInHistory, bSyncCategoryList
End Function

Public Sub NewTemplate(Optional ByVal bAutoCreate As Boolean = False, Optional ByVal sTitle As String, Optional ByVal sDefaultShortName As String, Optional ByVal bJumpToAfterCreate As Boolean = True)
    m_oUI.NewTemplate bAutoCreate, sTitle, sDefaultShortName, bJumpToAfterCreate
End Sub

Public Function sChooseColor(Optional ByVal sInitialColor As String) As String
    sChooseColor = m_oUI.sChooseColor(sInitialColor)
End Function

Public Function sChooseFile(Optional ByVal sPath As String, Optional ByVal sFilename As String, Optional ByVal sFilter As String) As String
    sChooseFile = m_oUI.sChooseFile(sPath, sFilename, sFilter)
End Function

Public Sub ShowSplashScreen()
On Error Resume Next
    frmSplash.DetermineRegistration
    frmSplash.Show
End Sub

Public Property Get TemplateDatabaseName() As String
On Error Resume Next
    TemplateDatabaseName = m_oUI.TemplateDatabaseName
End Property

Public Function GetATemplate(sCategory As String, sTemplate As String) As CTemplate
On Error Resume Next
    Set GetATemplate = Nothing
    Set GetATemplate = m_oUI.SliceAndDice.Categorys(sCategory).Templates(sTemplate)
End Function


Public Function LogError(ByVal sModuleName As String, sProcName As String, lError As Long, sErrorMsg As String) As Boolean
    LogError = modGeneral.LogError(sModuleName, sProcName, lError, sErrorMsg)
End Function

Public Function sFileContents(sPathAndFilename As String) As String
    sFileContents = modGeneral.sFileContents(sPathAndFilename)
End Function

Public Function sGetGUID(ByVal sProgID As String) As String
    sGetGUID = modGeneral.sGetGUID(sProgID)
End Function

Public Function sChoose(sChoices As String, Optional ByVal sDelimiter As String = ";", Optional ByVal sDefault As String)
    sChoose = modGeneral.sChoose(sChoices, sDelimiter, sDefault)
End Function

Public Function sExtractToken(ByRef sOrigStr As String, Optional ByVal nToken As Integer = 1, Optional ByVal strDelim As String = " ")
    sExtractToken = modGeneral.sExtractToken(sOrigStr, nToken, strDelim)
End Function

Public Function BrowseForFolder(ByVal sPrompt As String) As String
    BrowseForFolder = modGeneral.BrowseForFolder(m_oUI.hWnd, sPrompt)
End Function

Public Function bUserSure(Optional ByVal sPrompt As String = "Are you sure this is what you want to do ?") As Boolean
    bUserSure = modGeneral.bUserSure(sPrompt)
End Function

Public Function NextNegativeUnique() As Long
    NextNegativeUnique = modGeneral.NextNegativeUnique
End Function

Public Function lTokenCount(ByVal sAllTokens As String, Optional ByVal sDelim As String = " ") As Long
    lTokenCount = modGeneral.lTokenCount(sAllTokens, sDelim)
End Function

Public Function nZ(ByRef vData As Variant, Optional sDefault As String = "") As String
    nZ = modGeneral.nZ(vData, sDefault)
End Function

Public Function sDenormalize(sLine As String) As String
    sDenormalize = modGeneral.sDenormalize(sLine)
End Function

Public Function sGetToken(ByVal sAllTokens As String, Optional ByVal lToken As Long = 1, Optional ByVal sDelim As String = " ") As String
    sGetToken = modGeneral.sGetToken(sAllTokens, lToken, sDelim)
End Function

Public Function sAfter(ByVal sAllTokens As String, Optional ByVal lToken As Long = 1, Optional ByVal sDelim As String = " ") As String
    sAfter = modGeneral.sAfter(sAllTokens, lToken, sDelim)
End Function

Public Function sBefore(ByVal sAllTokens As String, Optional ByVal lToken As Long = 2, Optional ByVal sDelim As String = " ") As String
    sBefore = modGeneral.sBefore(sAllTokens, lToken, sDelim)
End Function

Public Function sExcept(ByVal sAllTokens As String, Optional ByVal lToken As Long = 1, Optional ByVal sDelim As String = " ") As String
    sExcept = modGeneral.sExcept(sAllTokens, lToken, sDelim)
End Function

Public Function sInsertSpaces(ByVal sToInsertInto As String) As String
    sInsertSpaces = modGeneral.sInsertSpaces(sToInsertInto)
End Function

Public Function sNormalize(sLine As String) As String
    sNormalize = modGeneral.sNormalize(sLine)
End Function

Public Function Replace(ByVal sAll As String, ByVal sFind As String, ByVal ReplaceWith As String) As String
    Replace = modGeneral.Replace(sAll, sFind, ReplaceWith)
End Function

Public Property Get TemplateDatabasePath() As String
On Error Resume Next
    TemplateDatabasePath = sBefore(m_oUI.TemplateDatabaseName, lTokenCount(m_oUI.TemplateDatabaseName, "\"), "\") & "\"
End Property

Public Property Get Version() As String
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Function zn(sData As String) As Variant
    zn = modGeneral.zn(sData)
End Function

Public Function lFindToken(ByVal sAllTokens As String, ByVal sTokenToFind As String, Optional ByVal sDelimiter As String = " ") As Long
    lFindToken = modGeneral.lFindToken(sAllTokens, sTokenToFind, sDelimiter)
End Function

Public Function AddPopup(mcp As CommandBarPopup, sMenu As String, nBefore As Long, sCaption As String, bBeginGroup As Boolean, bNextBeginGroup As Boolean) As String
On Error GoTo EH_Wizard_AddPopup
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Function
    bInHereAlready = True

    If m_oVBInst.CommandBars(sMenu) Is Nothing Then
       MsgBox "Hmm... There doesn't seem to be anywhere I can place the following Slice and Dice button on the (nonexistant) '" & sMenu & "' menu:" & gs2EOLTab & "With Caption: " & sCaption & gs2EOLTab & "At position: " & nBefore
       Exit Function
    End If

    If m_oVBInst.CommandBars(sMenu).Controls.Count < nBefore Then
       If m_oVBInst.CommandBars(sMenu).Controls.Count > 0 Then
          nBefore = m_oVBInst.CommandBars(sMenu).Controls.Count
       Else
          ' No where to put that button
          MsgBox "Hmm... There doesn't seem to be anywhere I can place the following Slice and Dice button on the '" & sMenu & "' menu:" & gs2EOLTab & "With Caption: " & sCaption & gs2EOLTab & "At position: " & nBefore
          Exit Function
       End If
    End If

    Set mcp = m_oVBInst.CommandBars(sMenu).Controls.Add(before:=nBefore, Type:=msoControlPopup)
    If mcp Is Nothing Then
       mcp.BeginGroup = bBeginGroup
       mcp.Caption = sCaption
       If bNextBeginGroup = True Then m_oVBInst.CommandBars(sMenu).Controls(nBefore + 1).BeginGroup = True
       AddPopup = mcp.CommandBar.Name
    End If

EH_Wizard_AddPopup_Continue:
    bInHereAlready = False
    Exit Function

EH_Wizard_AddPopup:
    MsgBox "Error occured in:" & vbCr & vbTab & "Module: Wizard" & vbCr & vbTab & "Procedure: AddPopup" & vbCr & vbCr & Err.Description
    
    Resume EH_Wizard_AddPopup_Continue

    Resume
End Function

Public Function AddButton(mcb As CommandBarButton, sMenu As String, nBefore As Long, sCaption As String, nButtonFace As Long, bBeginGroup As Boolean, bNextBeginGroup As Boolean) As CommandBarEvents
On Error GoTo EH_Wizard_AddButton
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Function
    bInHereAlready = True

    If m_oVBInst.CommandBars(sMenu) Is Nothing Then
       MsgBox "Hmm... There doesn't seem to be anywhere I can place the following Slice and Dice button on the (nonexistant) '" & sMenu & "' menu:" & gs2EOLTab & "With Caption: " & sCaption & gs2EOLTab & "At position: " & nBefore
       Exit Function
    End If

    If m_oVBInst.CommandBars(sMenu).Controls.Count < nBefore Then
       If m_oVBInst.CommandBars(sMenu).Controls.Count > 0 Then
          nBefore = m_oVBInst.CommandBars(sMenu).Controls.Count
       Else
          ' No where to put that button
          MsgBox "Hmm... There doesn't seem to be anywhere I can place the following Slice and Dice button on the '" & sMenu & "' menu:" & gs2EOLTab & "With Caption: " & sCaption & gs2EOLTab & "At position: " & nBefore
          Exit Function
       End If
    End If

    Set mcb = m_oVBInst.CommandBars(sMenu).Controls.Add(before:=nBefore)
    If Not mcb Is Nothing Then
       mcb.BeginGroup = bBeginGroup
       If Len(sCaption) > 0 Then mcb.Caption = sCaption
       If nButtonFace <> 0 Then mcb.FaceId = nButtonFace
On Error Resume Next
       If m_oVBInst.CommandBars(sMenu).Controls.Count >= nBefore + 1 Then
          m_oVBInst.CommandBars(sMenu).Controls(nBefore + 1).BeginGroup = bNextBeginGroup
On Error GoTo EH_Wizard_AddButton
       End If
       Set AddButton = m_oVBInst.Events.CommandBarEvents(mcb)
    End If
    
EH_Wizard_AddButton_Continue:
    bInHereAlready = False
    Exit Function

EH_Wizard_AddButton:
    MsgBox "Error occured in:" & gsEolTab & "Module: Wizard" & gsEolTab & "Procedure: AddButton" & gs2EOL & Err.Description & gs2EOLTab & "Menu: " & sMenu & gs2EOLTab & "Caption: " & sCaption & gs2EOLTab & "At position: " & nBefore
    Resume EH_Wizard_AddButton_Continue

    Resume
End Function

Public Sub DeleteCurrentTextSelection()
    m_oUI.DeleteCurrentTextSelection
End Sub

Public Function DetermineFirstLineInSelection() As Long
    DetermineFirstLineInSelection = m_oUI.DetermineFirstLineInSelection
End Function

Public Function DetermineLastLineInSelection() As Long
    DetermineLastLineInSelection = m_oUI.DetermineLastLineInSelection
End Function


Public Sub DoInsertion(asaV As CAssocArray, sTemplateToInsert As String, Optional ByVal bSkipDeclarations As Boolean = False)
    m_oUI.DoInsertion asaV, sTemplateToInsert, bSkipDeclarations
End Sub


Public Function FillTemplateWithUserInput(ByRef asaX As CAssocArray, ByVal sToParse As String, ByRef sCodeToInsert As String, ByVal sMsgBoxTitle As String) As Boolean
    FillTemplateWithUserInput = m_oUI.FillTemplateWithUserInput(asaX, sToParse, sCodeToInsert, sMsgBoxTitle)
End Function

Public Function FindLastProcLine(sProcName As String, lProcType As Long) As Long
    FindLastProcLine = m_oUI.FindLastProcLine(sProcName, lProcType)
End Function



Public Sub GetCategoryAndName(ByVal sCategoryAndName As String, ByRef sCategory As String, ByRef sShortName As String)
    m_oUI.GetCategoryAndName sCategoryAndName, sCategory, sShortName
End Sub

Public Function GetCurrentTextSelection() As String
    GetCurrentTextSelection = m_oUI.GetCurrentTextSelection
End Function

Public Sub GetProcAtLine(ByVal lCurrentLine As Long, ByRef sProcName As String, ByRef lProcType As Long)
    m_oUI.GetProcAtLine lCurrentLine, sProcName, lProcType
End Sub


Public Function InsertTemplate(ByVal sToParseIn As String, Optional ByVal lLineIn As Long = 1, Optional ByRef SoftVars As CAssocArray, Optional ByVal sTemplateName As String = "Slice and Dice", Optional ByVal sExternalFilename As String = "") As Boolean
    Dim InsertionInfo As CInsertionInfo

    Set InsertionInfo = New CInsertionInfo
        With InsertionInfo
             If SoftVars Is Nothing Then
                Set .SoftVars = New CAssocArray
             Else
                Set .SoftVars = SoftVars
             End If
             .OriginalCodeToInsert = sToParseIn
             .PointOfInsertion = lLineIn
             .ExternalFilename = sExternalFilename
             .TemplateName = sTemplateName
        End With

        InsertTemplate = m_oUI.InternalInsertTemplate(InsertionInfo)
        Set InsertionInfo.SoftVars = Nothing
    Set InsertionInfo = Nothing
End Function

Public Property Get InternalCurrentTemplate() As CTemplate
    Set InternalCurrentTemplate = m_oUI.InternalCurrentTemplate
End Property

Public Property Get CurrentTemplate() As CTemplate
    Set CurrentTemplate = m_oUI.CurrentTemplate
End Property


Public Function sChooseDatabase(Optional ByVal sPath As String, Optional ByVal sFilename As String) As String
    sChooseDatabase = m_oUI.sChooseDatabase(sPath, sFilename)
End Function

Public Property Get SliceAndDice() As CSliceAndDice
    Set SliceAndDice = m_oUI.SliceAndDice
End Property

Public Function sPropertyType(sFieldType As String) As String
    sPropertyType = m_oUI.sPropertyType(sFieldType)
End Function


Friend Property Get UI() As frmMain
    Set UI = m_oUI
End Property

Public Property Get SandyWindow() As Object
    Set SandyWindow = m_oUI
End Property

Friend Property Get Window() As Window
    Set Window = m_oUI
End Property


Friend Property Get VBInst() As VBE
    Set VBInst = m_oVBInst
End Property

Public Property Get CurrentTemplateNameAndCategory() As String
    CurrentTemplateNameAndCategory = m_oUI.txtName.Text
End Property

Private Sub AltChangeToHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ChangeToHandler_Click CommandBarControl, handled, CancelDefault
End Sub

Private Sub BarHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    MenuHandler_Click CommandBarControl, handled, CancelDefault
End Sub

Private Sub ChangeToHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim CurControl  As VBControl
    Dim NewControl  As VBControl
    Dim asaVar      As CAssocArray

    Dim sChoices As String
    Dim sTemplate As String
    Dim sProgID As String
    Dim sLine As String
    Dim sLastClassName As String
    Dim sLastChoice As String
    Dim sCodeToInsert As String
    Dim sCodeToInsert2 As String
    Dim sPropertyName As String
    
    Dim bInTemplate As Boolean

    If PropertyStack Is Nothing Then
       Set PropertyStack = CreateObject("SandySupport.CAssocArray")
    End If
    Set asaVar = CreateObject("SandySupport.CAssocArray")

On Error Resume Next

    

ChangeToHandler_Click_Start_Over:
    With m_oVBInst.SelectedVBComponent.Designer
         gbCancelInsertion = False
         For Each CurControl In .SelectedVBControls
             If gbCancelInsertion Then Exit For
             With CurControl
                  If Not m_oUI.SetInternalCurrentTemplate("Change from - " & .ClassName) Then
                     If Not bUserSure("Please set up a 'Change to' template named" & vbCr & vbTab & "'Change from - " & .ClassName & "'" & vbCr & "before using this function on the '" & .ClassName & "' type control." & vbCr & vbCr & "Select Yes to create the new template." & vbCr & "Select No to abort.") Then
                        Exit Sub
                     End If

                     m_oUI.Visible = True
                     m_oUI.NewTemplate True, "Change from - " & .ClassName
                     Exit Sub
                  End If

                  If .ClassName = sLastClassName Then
                     sProgID = sLastChoice
                  Else
                     sChoices = ""
                     If m_oUI.SliceAndDice.Categorys("Change from").Templates("All Types") Is Nothing Or InStr(1, m_oUI.InternalCurrentTemplate.memoCodeAtTop, "NoAllTypes", vbTextCompare) Then
                        sTemplate = m_oUI.InternalCurrentTemplate.memoCodeAtTop
                     Else
                        With m_oUI.SliceAndDice.Categorys("Change from").Templates("All Types")
                             If Right(.memoCodeAtTop, 2) = gsEOL Then
                                sTemplate = .memoCodeAtTop & m_oUI.InternalCurrentTemplate.memoCodeAtTop
                             Else
                                sTemplate = .memoCodeAtTop & gsEOL & m_oUI.InternalCurrentTemplate.memoCodeAtTop
                             End If
                        End With
                     End If

                     Do Until Len(sTemplate) = 0
                        sLine = sGetToken(sTemplate, 1, gsEOL)
                        sTemplate = sAfter(sTemplate, 1, gsEOL)
                        If Left(sLine, 2) = "**" Then
                           sChoices = sChoices & ";" & Mid(sLine, 4)
                        End If
                     Loop
                     sChoices = Mid(sChoices, 2)
                     sProgID = sChoose(sChoices)
                  End If
                  If Len(sProgID) > 0 Then
                     sLastClassName = .ClassName
                     sLastChoice = sProgID
                     
                     If m_oUI.SliceAndDice.Categorys("Change from").Templates("All Types") Is Nothing Or InStr(1, m_oUI.InternalCurrentTemplate.memoCodeAtTop, "NoAllTypes", vbTextCompare) Then
                        sTemplate = m_oUI.InternalCurrentTemplate.memoCodeAtTop
                     Else
                        With m_oUI.SliceAndDice.Categorys("Change from").Templates("All Types")
                             If Right(.memoCodeAtTop, 2) = gsEOL Then
                                sTemplate = .memoCodeAtTop & m_oUI.InternalCurrentTemplate.memoCodeAtTop
                             Else
                                sTemplate = .memoCodeAtTop & gsEOL & m_oUI.InternalCurrentTemplate.memoCodeAtTop
                             End If
                        End With
                     End If
                     bInTemplate = False
                     Do Until Len(sTemplate) = 0
                        sLine = sGetToken(sTemplate, 1, gsEOL)
                        sTemplate = sAfter(sTemplate, 1, gsEOL)
                        If sLine = "** " & sProgID Then
                           bInTemplate = True
                           sTemplate = sGetToken(sTemplate, 1, gsEOL & "** ")
                           asaVar("ClassName") = CurControl.ClassName
                           asaVar("ProgID") = CurControl.ProgId
                           If Not m_oUI.FillTemplateWithUserInput(asaVar, sTemplate, sCodeToInsert, "GROUP MODIFY") Then
                              Exit Sub
                           End If
                           sTemplate = sCodeToInsert
                        ElseIf Left(sLine, 2) = "**" And bInTemplate = True Then
                           bInTemplate = False
                           sTemplate = ""
                        ElseIf bInTemplate = True Then
                           Select Case UCase(sGetToken(sLine))
                                  Case "ADDCONTROL"
                                        Set NewControl = .Container.ContainedVBControls.Add(sAfter(sLine), CurControl, True)
                                        If NewControl Is Nothing Then
                                           MsgBox "That control has not been referenced yet. Please add a reference first.", vbInformation
                                           sTemplate = ""
                                        End If
                                        
                                  Case "SETNAME"
                                        NewControl.Properties("Name") = sAfter(sLine) & Mid(CurControl.Properties("Name"), Len(sAfter(sLine)))
                                        
                                  Case "COPYPROPERTY"
                                        NewControl.Properties(sAfter(sLine)) = CurControl.Properties(sAfter(sLine))
                                        
                                  Case "SETPROPERTY"
                                        sPropertyName = sGetToken(sAfter(sLine), 1, "=")
                                        If InStr(sPropertyName, ".") > 0 Then
                                           NewControl.Properties(sGetToken(sPropertyName, 1, ".")).Value(sGetToken(sPropertyName, 2, ".")).Value = sAfter(sLine, 1, "=")
                                        Else
                                           NewControl.Properties(sPropertyName) = sAfter(sLine, 1, "=")
                                        End If
                                        If Err.Number <> 0 Then
                                           MsgBox "Warning: Failed to set the property '" & sPropertyName & "' to '" & sAfter(sLine, 1, "=") & vbCr & vbTab & Err.Description
                                        End If
                                       'NewControl.Properties(sGetToken(sAfter(sLine), 1, "=")) = sAfter(sLine, 1, "=")
                                        
                                  Case "SETCURRENTPROPERTY"
                                        Err.Clear
                                        sPropertyName = sGetToken(sAfter(sLine), 1, "=")
                                        If InStr(sPropertyName, ".") > 0 Then
                                           CurControl.Properties(sGetToken(sPropertyName, 1, ".")).Value(sGetToken(sPropertyName, 2, ".")).Value = sAfter(sLine, 1, "=")
                                        Else
                                           CurControl.Properties(sPropertyName) = sAfter(sLine, 1, "=")
                                        End If
                                        If Err.Number <> 0 Then
                                           MsgBox "Warning: Failed to set the property '" & sPropertyName & "' to '" & sAfter(sLine, 1, "=") & vbCr & vbTab & Err.Description
                                        End If
                                        
                                  Case "DELETECURRENTCONTROL"
                                        m_oVBInst.SelectedVBComponent.Designer.VBControls.Remove CurControl
                                        Set CurControl = Nothing

                                  Case "AUTOSPACECURRENT"
                                        CurControl.Properties(sGetToken(sAfter(sLine), 1, "=")) = sInsertSpaces(CurControl.Properties(sGetToken(sAfter(sLine), 1, "=")))
                                        
                                  Case "AUTOSPACE"
                                        NewControl.Properties(sGetToken(sAfter(sLine), 1, "=")) = sInsertSpaces(NewControl.Properties(sGetToken(sAfter(sLine), 1, "=")))
                                        
                                  Case "OFFSETFROMCURRENTLEFT", "OFFSET", "OFFSETLEFT"
                                        NewControl.Properties("Left") = CurControl.Properties("Left") + CurControl.Properties("Width") + IIf(Val(sGetToken(sLine, 2)) = 0, 100, Val(sGetToken(sLine, 2)))
                                        
                                  Case "OFFSETFROMCURRENTTOP", "OFFSETTOP"
                                        NewControl.Properties("Top") = CurControl.Properties("Top") + CurControl.Properties("Height") + IIf(Val(sGetToken(sLine, 2)) = 0, 100, Val(sGetToken(sLine, 2)))
                                        
                                  Case "QUEUE", "QUEUEINSERTION", "ADDTOQUEUE"
                                        asaVar("Queued").Value = asaVar("Queued").Value & sAfter(sLine) & "~"
                                        
                                  Case "COPYLOCATIONCURRENT", "COPYLOCATION"
                                        PropertyStack("Left") = CurControl.Properties("Left")
                                        PropertyStack("Top") = CurControl.Properties("Top")
                                        PropertyStack("Width") = CurControl.Properties("Width")
                                        PropertyStack("Height") = CurControl.Properties("Height")

                                  Case "PASTELOCATIONCURRENT", "PASTELOCATION"
                                        If PropertyStack("Left") < 0 Then PropertyStack("Left") = 0
                                        If PropertyStack("Top") < 0 Then PropertyStack("Top") = 0
                                        If PropertyStack("Width") < 30 Then PropertyStack("Width") = 30
                                        If PropertyStack("Height") < 30 Then PropertyStack("Height") = 30

                                        CurControl.Properties("Left") = PropertyStack("Left")
                                        CurControl.Properties("Top") = PropertyStack("Top")
                                        CurControl.Properties("Width") = PropertyStack("Width")
                                        CurControl.Properties("Height") = PropertyStack("Height")

                                  Case "CLEARPROPERTYSTACK", "CLEARSTACK"
                                        PropertyStack.Clear

                                  Case "PUSHPROPERTY", "PUSH"
                                        PropertyStack(sAfter(sLine)) = CurControl.Properties(sAfter(sLine))

                                  Case "POPPORPERTYCURRENT", "POPPROPERTY", "POP"
                                        CurControl.Properties(sAfter(sLine)) = PropertyStack(sAfter(sLine))

                                  Case "INCREMENT", "INC"
                                  
                                  Case "SETVALUE", "SET"
                                       asaVar.Item(sGetToken(sAfter(sLine), 1, "=")) = sAfter(sLine, 1, "=")

                                  Case "INSERTTEMPLATE", "INSERT"
                                        gbCancelInsertion = False
                                        m_oUI.DoInsertion asaVar, sAfter(sLine)
                                        If gbCancelInsertion Then
                                           sTemplate = ""
                                        End If
                                        
                                  Case "GETPROPERTY", "GET"
                                       asaVar.Item(sGetToken(sAfter(sLine), 1, "=")) = CurControl.Properties(sAfter(sAfter(sLine), 1, "="))
                           End Select
                        End If
                     Loop
                  End If
             End With
         Next CurControl
    End With

    If Not gbCancelInsertion And Len(asaVar("Queued").Value) Then
       m_oUI.QueuedInsertions = asaVar("Queued").Value
    End If

    handled = True

    Set asaVar = Nothing
End Sub

Private Sub Class_Initialize()
    ' LogEvent "Wizard: Initialize"
End Sub

Private Sub Class_Terminate()
    Set PropertyStack = Nothing
    ' LogEvent "Wizard: Terminate"
End Sub


Private Sub ExternalsHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error Resume Next
    m_oUI.ShowExternalsMenu

    handled = True
End Sub

Private Sub FavoritesHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error Resume Next
    m_oUI.FavoriteCalledFromIDE = True
    m_oUI.ShowFavMenu

    handled = True
End Sub

'Private Sub FormChangeToHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'On Error Resume Next
'    MsgBox "Form Change not currently implemented. Sorry."
'End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
    ' Comment to prevent procedure from being deleted on compilation.
End Sub

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
On Error Resume Next
    Dim sTemplateDatabasePath As String
    Dim CurWindow As Window
    Dim bFound As Boolean
    Dim bShown As Boolean
    Dim lFirstButton As Long

    frmSplash.lblDLLsLoaded(1).Caption = "0"
    If GetSetting("SliceAndDice", "Last", "Show Splash", True) Then
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

        If GetSetting("SliceAndDice", "Last", "Show Paitbrush Icon", True) Then
           Set BarHandler = AddButton(mcbEditButton, "Standard", 1, "Slice and Dice", 30, False, True)
        End If

       'Set ImportHandler = AddButton(mcbImportButton, "Code Window", 1, "I&mport selection as new Template", 30, False, False)

        If GetSetting("SliceAndDice", "Last", "Show On Module Right Click", True) Then
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
            .Attach m_oUI.hWnd
            .RegisterKey "Sandy Activate", vbKeyS, MOD_CONTROL + MOD_SHIFT
            .RegisterKey "Sandy Repeat Insertion", vbKeyR, MOD_CONTROL + MOD_SHIFT
            .RegisterKey "Sandy Favorites", vbKeyF, MOD_CONTROL + MOD_SHIFT
            .RegisterKey "Sandy Externals", vbKeyE, MOD_CONTROL + MOD_SHIFT
       End With
    End If
    End If
    If bShown Then frmSplash.Hide
End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)
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
    SaveSetting "SliceAndDice", "Settings", "Exit after insert", IIf(m_oUI.ExitAfterInsert, "True", "False")
    SaveSetting "SliceAndDice", "Settings", "Last Template", m_oUI.CurrentTemplateNameAndCategory
    SaveFormPosition m_oUI
    
   'SaveSetting "SliceAndDice", "Settings", "Bar Arrange", lsbJumpTo.Arrange
   'SaveSetting "SliceAndDice", "Settings", "Bar View", lsbJumpTo.View

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

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
  ' Comment to prevent procedure from being deleted on compilation.
    'Set mySubClass = New CSubClassIDE
    'mySubClass.Key = &H9111973
    'Set mySubClass.VBInstance = vbInst
    'mySubClass.CaptionAdd = " [SubClassIDE Demo On]"
End Sub
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error Resume Next
    m_oUI.Visible = Not m_oUI.Visible
    If m_oUI.Visible Then
       m_oUI.SetFocus
       m_oUI.ZOrder
    Else
       m_oUI.HideAllWindows
    End If

    handled = True
End Sub

'Private Sub mySubClass_KeyDown(KeyCode As Long, ByVal Shift As Integer)
'    If Shift = 3 Then '(Shift And vbShiftMask) And (Shift And vbCtrlMask) Then
'       Select Case KeyCode
'              Case 69  ' "E"xternals window
'                   ExternalsHandler_Click Nothing, False, False
'              Case 70  ' "F"avorites
'                   FavoritesHandler_Click Nothing, False, False
'              Case 83  ' "S"lice and Dice window
'                   MenuHandler_Click Nothing, False, False
'              Case Else
'                   MsgBox "Combination key Shift-Ctrl-" & KeyCode & " pressed"
'       End Select
'    End If
'End Sub
'
Private Sub ShortcutHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    MenuHandler_Click CommandBarControl, handled, CancelDefault
End Sub

'Private Sub SubClassIDE_KeyDown(KeyCode As Long, ByVal Shift As Integer)
'    SubClassIDE_KeyUp KeyCode, Shift
'End Sub

'Private Sub SubClassIDE_KeyPress(KeyAscii As Long)
'    SubClassIDE_KeyUp KeyAscii, 0
'End Sub


'Private Sub SubClassIDE_KeyUp(KeyCode As Long, ByVal Shift As Integer)
'    MsgBox "Bing !"
'    If ((Shift And vbCtrlMask) And vbAltMask) > 0 Then
'       Select Case KeyCode
'              Case vbKeyS
'                   MenuHandler_Click Nothing, False, False
'       End Select
'    End If
'End Sub




