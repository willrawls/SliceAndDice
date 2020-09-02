VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSandy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moIDE  As SandySupport.ISandyIDE
Private moUI   As SandySupport.ISandyWindowMain

Implements SandySupport.ISandyWizard

Private Property Let ISandyWizard_InsertionCancelled(ByVal RHS As Boolean)
    CancelInsertion = RHS
End Property

Private Property Get ISandyWizard_InsertionCancelled() As Boolean
    ISandyWizard_InsertionCancelled = CancelInsertion
End Property

Private Property Get ISandyWizard_SandyIDE() As SandySupport.ISandyIDE
    Set ISandyWizard_SandyIDE = moIDE
End Property

Private Function ISandyWizard_sTableToPropertyName(ByVal sTableName As String) As String
On Error Resume Next
    ISandyWizard_sTableToPropertyName = sTableToPropertyName(sTableName)
End Function

Private Property Get ISandyWizard_Version() As String
    ISandyWizard_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Function ISandyWizard_OnConnection(xIDE As SandySupport.ISandyIDE) As Boolean
On Error Resume Next
    Dim sTemplateDatabasePath   As String
    Dim bFound                  As Boolean
    Dim bShown                  As Boolean

    Set moIDE = xIDE

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

    Set moUI = New frmMain

    Set moUI.Parent = Me
    'Set moUI.DBClassGen = frmDBClassGen
    'Set moUI.DBClassGen.Parent = moUI

    DoEvents
      sTemplateDatabasePath = ISandyWizard_TemplateDatabasePath
    DoEvents: DoEvents: DoEvents

    Call moUI.InitializeAddinDLLs(EnumFiles(App.Path) & EnumFiles(ISandyWizard_TemplateDatabasePath))
   'moUI.InitializeAddinDLLs "sadAddin.NewCommands=LOAD"

    If moUI.mHotKeyOpenWindow Is Nothing Then
       Set moUI.mHotKeyOpenWindow = CreateObject("SandySupport.cRegHotKey")
       If Not moUI.mHotKeyOpenWindow Is Nothing Then
          With moUI.mHotKeyOpenWindow
               .Attach moUI.hwnd
               .RegisterKey "Sandy Activate", vbKeyS, MOD_CONTROL + MOD_SHIFT
               .RegisterKey "Sandy Repeat Insertion", vbKeyR, MOD_CONTROL + MOD_SHIFT
               .RegisterKey "Sandy Favorites", vbKeyF, MOD_CONTROL + MOD_SHIFT
               .RegisterKey "Sandy Externals", vbKeyE, MOD_CONTROL + MOD_SHIFT
          End With
       End If
    End If

    If bShown Then frmSplash.Hide
    
    ISandyWizard_OnConnection = True
End Function

Public Sub ISandyWizard_OnDisconnection()
On Error Resume Next
    Dim Cancel As Integer

  ' Make sure the edited Template (if one) is saved
    moUI.SaveTemplate
    moUI.Visible = False

  ' Save settings for next time
    SaveSetting "SliceAndDice", "Settings", "Exit after insert", IIf(moUI.ExitAfterInsert, "True", "False")
    SaveSetting "SliceAndDice", "Settings", "Last Template", moUI.CurrentTemplateNameAndCategory
    SaveFormPosition moUI
    
    moUI.HideAllWindows True
    moUI.ShutdownDLLs
    
    Set moUI.Parent = Nothing
   'Set moUI.DBClassGen = Nothing
    moUI.FormUnload
    Unload moUI
    Set moUI = Nothing

  ' Disassociate external objects referenced in this object
    Set moIDE = Nothing
End Sub

Private Function ISandyWizard_EnumFiles(sPath As String, sMask As String, sAttr As String) As String
    Dim Attr As VbFileAttribute
    
    Select Case Attr 'UCase$(sAttr)
           Case "ALIAS":        Attr = vbAlias
           Case "ARCHIVE":      Attr = vbArchive
           Case "DIRECTORY":    Attr = vbDirectory
           Case "HIDDEN":       Attr = vbHidden
           Case "READONLY":     Attr = vbReadOnly
           Case "SYSTEM":       Attr = vbSystem
           Case "VOLUME":       Attr = vbVolume
           Case Else:           Attr = vbNormal
    End Select
    ISandyWizard_EnumFiles = EnumFiles(sPath, sMask, Attr)
End Function

Private Function ISandyWizard_Evaluate(ByVal sExpression As String, ByRef asaVar As SandySupport.CAssocArray)
    ISandyWizard_Evaluate = Evaluate(sExpression, asaVar)
End Function

Private Function ISandyWizard_FileExists(sFilename As String) As Boolean
    ISandyWizard_FileExists = FileExists(sFilename)
End Function

Private Function ISandyWizard_JumpTo(ByVal sTemplateName As String, Optional ByVal bRecordInHistory As Boolean = True, Optional ByVal bSyncCategoryList As Boolean = False) As Boolean
    ISandyWizard_JumpTo = moUI.JumpTo(sTemplateName, bRecordInHistory, bSyncCategoryList)
End Function

Private Sub ISandyWizard_NewTemplate(Optional ByVal bAutoCreate As Boolean = False, Optional ByVal sTitle As String, Optional ByVal sDefaultShortName As String, Optional ByVal bJumpToAfterCreate As Boolean = True)
    moUI.NewTemplate bAutoCreate, sTitle, sDefaultShortName, bJumpToAfterCreate
End Sub

Private Function ISandyWizard_sChooseColor(Optional ByVal sInitialColor As String) As String
    ISandyWizard_sChooseColor = moUI.sChooseColor(sInitialColor)
End Function

Private Function ISandyWizard_sChooseFile(Optional ByVal sPath As String, Optional ByVal sFilename As String, Optional ByVal sFilter As String) As String
    ISandyWizard_sChooseFile = moUI.sChooseFile(sPath, sFilename, sFilter)
End Function

Private Sub ISandyWizard_ShowSplashScreen()
On Error Resume Next
    frmSplash.DetermineRegistration
    frmSplash.Show
End Sub

Private Property Get ISandyWizard_TemplateDatabaseName() As String
On Error Resume Next
    ISandyWizard_TemplateDatabaseName = moUI.TemplateDatabaseName
End Property

Private Function ISandyWizard_GetATemplate(sCategory As String, sTemplate As String) As CTemplate
On Error Resume Next
    Set ISandyWizard_GetATemplate = Nothing
    Set ISandyWizard_GetATemplate = moUI.SliceAndDice.Categorys(sCategory).Templates(sTemplate)
End Function

Private Function ISandyWizard_LogError(ByVal sModuleName As String, sProcName As String, lError As Long, sErrorMsg As String) As Boolean
    ISandyWizard_LogError = LogError(sModuleName, sProcName, lError, sErrorMsg)
End Function

Private Function ISandyWizard_sFileContents(sPathAndFilename As String) As String
    ISandyWizard_sFileContents = sFileContents(sPathAndFilename)
End Function

Private Function ISandyWizard_sGetGUID(ByVal sProgID As String) As String
    ISandyWizard_sGetGUID = sGetGUID(sProgID)
End Function

Private Function ISandyWizard_sChoose(sChoices As String, Optional ByVal sDelimiter As String = ";", Optional ByVal sDefault As String)
    ISandyWizard_sChoose = sChoose(sChoices, sDelimiter, sDefault)
End Function

Private Function ISandyWizard_sExtractToken(ByRef sOrigStr As String, Optional ByVal nToken As Integer = 1, Optional ByVal strDelim As String = " ")
    ISandyWizard_sExtractToken = sExtractToken(sOrigStr, nToken, strDelim)
End Function

Private Function ISandyWizard_BrowseForFolder(ByVal sPrompt As String) As String
    ISandyWizard_BrowseForFolder = BrowseForFolder(moUI.hwnd, sPrompt)
End Function

Private Function ISandyWizard_bUserSure(Optional ByVal sPrompt As String = "Are you sure this is what you want to do ?") As Boolean
    ISandyWizard_bUserSure = bUserSure(sPrompt)
End Function

Private Function ISandyWizard_NextNegativeUnique() As Long
    ISandyWizard_NextNegativeUnique = NextNegativeUnique
End Function

Private Function ISandyWizard_lTokenCount(ByVal sAllTokens As String, Optional ByVal sDelim As String = " ") As Long
    ISandyWizard_lTokenCount = lTokenCount(sAllTokens, sDelim)
End Function

Private Function ISandyWizard_nZ(ByRef vData As Variant, Optional sDefault As String = vbNullString) As String
    ISandyWizard_nZ = nZ(vData, sDefault)
End Function

Private Function ISandyWizard_sDenormalize(sLine As String) As String
    ISandyWizard_sDenormalize = sDenormalize(sLine)
End Function

Private Function ISandyWizard_sGetToken(ByVal sAllTokens As String, Optional ByVal lToken As Long = 1, Optional ByVal sDelim As String = " ") As String
    ISandyWizard_sGetToken = sGetToken(sAllTokens, lToken, sDelim)
End Function

Private Function ISandyWizard_sAfter(ByVal sAllTokens As String, Optional ByVal lToken As Long = 1, Optional ByVal sDelim As String = " ") As String
    ISandyWizard_sAfter = sAfter(sAllTokens, lToken, sDelim)
End Function

Private Function ISandyWizard_sBefore(ByVal sAllTokens As String, Optional ByVal lToken As Long = 2, Optional ByVal sDelim As String = " ") As String
    ISandyWizard_sBefore = sBefore(sAllTokens, lToken, sDelim)
End Function

Private Function ISandyWizard_sExcept(ByVal sAllTokens As String, Optional ByVal lToken As Long = 1, Optional ByVal sDelim As String = " ") As String
    ISandyWizard_sExcept = sExcept(sAllTokens, lToken, sDelim)
End Function

Private Function ISandyWizard_sInsertSpaces(ByVal sToInsertInto As String) As String
    ISandyWizard_sInsertSpaces = sInsertSpaces(sToInsertInto)
End Function

Private Function ISandyWizard_sNormalize(sLine As String) As String
    ISandyWizard_sNormalize = sNormalize(sLine)
End Function

Private Property Get ISandyWizard_TemplateDatabasePath() As String
On Error Resume Next
    ISandyWizard_TemplateDatabasePath = sBefore(moUI.TemplateDatabaseName, lTokenCount(moUI.TemplateDatabaseName, "\"), "\") & "\"
End Property

Private Function ISandyWizard_zn(sData As String) As Variant
    ISandyWizard_zn = zn(sData)
End Function

Private Function ISandyWizard_lFindToken(ByVal sAllTokens As String, ByVal sTokenToFind As String, Optional ByVal sDelimiter As String = " ") As Long
    ISandyWizard_lFindToken = lFindToken(sAllTokens, sTokenToFind, sDelimiter)
End Function

Private Sub ISandyWizard_DeleteCurrentTextSelection()
    moUI.DeleteCurrentTextSelection
End Sub

Private Function ISandyWizard_DetermineFirstLineInSelection() As Long
    ISandyWizard_DetermineFirstLineInSelection = moUI.DetermineFirstLineInSelection
End Function

Private Function ISandyWizard_DetermineLastLineInSelection() As Long
    ISandyWizard_DetermineLastLineInSelection = moUI.DetermineLastLineInSelection
End Function

Private Sub ISandyWizard_DoInsertion(asaV As SandySupport.CAssocArray, sTemplateToInsert As String, Optional ByVal bSkipDeclarations As Boolean = False)
    moUI.DoInsertion asaV, sTemplateToInsert, bSkipDeclarations
End Sub

Private Function ISandyWizard_FillTemplateWithUserInput(ByRef asaX As SandySupport.CAssocArray, ByVal sToParse As String, ByRef sCodeToInsert As String, ByVal sMsgBoxTitle As String) As Boolean
    ISandyWizard_FillTemplateWithUserInput = moUI.FillTemplateWithUserInput(asaX, sToParse, sCodeToInsert, sMsgBoxTitle)
End Function

Private Function ISandyWizard_FindLastProcLine(sProcName As String, lProcType As Long) As Long
    ISandyWizard_FindLastProcLine = moUI.FindLastProcLine(sProcName, lProcType)
End Function

Private Sub ISandyWizard_GetCategoryAndName(ByVal sCategoryAndName As String, ByRef sCategory As String, ByRef sShortName As String)
    moUI.GetCategoryAndName sCategoryAndName, sCategory, sShortName
End Sub

Private Function ISandyWizard_GetCurrentTextSelection() As String
    ISandyWizard_GetCurrentTextSelection = moUI.GetCurrentTextSelection
End Function

Private Sub ISandyWizard_GetProcAtLine(ByVal lCurrentLine As Long, ByRef sProcName As String, ByRef lProcType As Long)
    moUI.GetProcAtLine lCurrentLine, sProcName, lProcType
End Sub

Private Function ISandyWizard_InsertTemplate(ByVal sToParseIn As String, Optional ByVal lLineIn As Long = 1, Optional ByRef SoftVars As SandySupport.CAssocArray, Optional ByVal sTemplateName As String = "Slice and Dice", Optional ByVal sExternalFilename As String = vbNullString) As Boolean
    Dim InsertionInfo As CInsertionInfo

    Set InsertionInfo = CreateObject("SandySupport.CInsertionInfo")
        With InsertionInfo
             If SoftVars Is Nothing Then
                Set .SoftVars = CreateObject("SandySupport.CAssocArray")
             Else
                Set .SoftVars = SoftVars
             End If
             .OriginalCodeToInsert = sToParseIn
             .PointOfInsertion = lLineIn
             .ExternalFilename = sExternalFilename
             .TemplateName = sTemplateName
        End With

        ISandyWizard_InsertTemplate = moUI.InternalInsertTemplate(InsertionInfo)
        Set InsertionInfo.SoftVars = Nothing
    Set InsertionInfo = Nothing
End Function

Private Property Get ISandyWizard_InternalCurrentTemplate() As CTemplate
    Set ISandyWizard_InternalCurrentTemplate = moUI.InternalCurrentTemplate
End Property

Private Property Get ISandyWizard_CurrentTemplate() As CTemplate
    Set ISandyWizard_CurrentTemplate = moUI.CurrentTemplate
End Property

Private Function ISandyWizard_sChooseDatabase(Optional ByVal sPath As String, Optional ByVal sFilename As String) As String
    ISandyWizard_sChooseDatabase = moUI.sChooseDatabase(sPath, sFilename)
End Function

Private Property Get ISandyWizard_SliceAndDice() As CSliceAndDice
    Set ISandyWizard_SliceAndDice = moUI.SliceAndDice
End Property

Private Function ISandyWizard_sPropertyType(sFieldType As String) As String
    ISandyWizard_sPropertyType = moUI.sPropertyType(sFieldType)
End Function

Private Property Get ISandyWizard_UI() As ISandyWindowMain
    Set ISandyWizard_UI = moUI
End Property

Private Property Get ISandyWizard_CurrentTemplateNameAndCategory() As String
    ISandyWizard_CurrentTemplateNameAndCategory = moUI.CurrentTemplateNameAndCategory
End Property

