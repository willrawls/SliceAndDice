VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSandyIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private moVBInst                        As VBIDE.VBE
'Private Sandy.UI                    As SandySupport.ISandyWindowMain
Private PropertyStack                   As SandySupport.CAssocArray

Private Sandy                           As SandySupport.ISandyWizard

Private mcbAddinButton                  As Office.CommandBarButton
Private mcbEditButton                   As Office.CommandBarButton
Private mcbShortcutButton               As Office.CommandBarButton
Private mcbChangeToButton               As Office.CommandBarButton
Private mcbAltChangeToButton            As Office.CommandBarButton
Private mcbFavoritesButton              As Office.CommandBarButton
Private mcbExternalsButton              As Office.CommandBarButton

Private WithEvents MenuHandler          As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
Private WithEvents BarHandler           As CommandBarEvents
Attribute BarHandler.VB_VarHelpID = -1
Private WithEvents ShortcutHandler      As CommandBarEvents
Attribute ShortcutHandler.VB_VarHelpID = -1
Private WithEvents ChangeToHandler      As CommandBarEvents
Attribute ChangeToHandler.VB_VarHelpID = -1
Private WithEvents AltChangeToHandler   As CommandBarEvents
Attribute AltChangeToHandler.VB_VarHelpID = -1
Private WithEvents FavoritesHandler     As CommandBarEvents
Attribute FavoritesHandler.VB_VarHelpID = -1
Private WithEvents ExternalsHandler     As CommandBarEvents
Attribute ExternalsHandler.VB_VarHelpID = -1

Implements SandySupport.ISandyIDE

Public Property Get Version() As String
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Function AddPopup(mcp As CommandBarPopup, sMenu As String, nBefore As Long, sCaption As String, bBeginGroup As Boolean, bNextBeginGroup As Boolean) As String
On Error GoTo EH_Wizard_AddPopup
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Function
    bInHereAlready = True

    If moVBInst.CommandBars(sMenu) Is Nothing Then
       MsgBox "Hmm... There doesn't seem to be anywhere I can place the following Slice and Dice button on the (nonexistant) '" & sMenu & "' menu:" & vbNewLine & vbNewLine & vbTab & "With Caption: " & sCaption & vbNewLine & vbNewLine & vbTab & "At position: " & nBefore
       Exit Function
    End If

    If moVBInst.CommandBars(sMenu).Controls.Count < nBefore Then
       If moVBInst.CommandBars(sMenu).Controls.Count > 0 Then
          nBefore = moVBInst.CommandBars(sMenu).Controls.Count
       Else
          ' No where to put that button
          MsgBox "Hmm... There doesn't seem to be anywhere I can place the following Slice and Dice button on the '" & sMenu & "' menu:" & vbNewLine & vbNewLine & vbTab & "With Caption: " & sCaption & vbNewLine & vbNewLine & vbTab & "At position: " & nBefore
          Exit Function
       End If
    End If

    Set mcp = moVBInst.CommandBars(sMenu).Controls.Add(before:=nBefore, Type:=msoControlPopup)
    If mcp Is Nothing Then
       mcp.BeginGroup = bBeginGroup
       mcp.Caption = sCaption
       If bNextBeginGroup = True Then moVBInst.CommandBars(sMenu).Controls(nBefore + 1).BeginGroup = True
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
    Static bInHereAlready As Boolean
    If bInHereAlready Then Exit Function
    bInHereAlready = True

On Error Resume Next
    If moVBInst.CommandBars(sMenu) Is Nothing Then
       MsgBox "Hmm... There doesn't seem to be anywhere I can place the following Slice and Dice button on the (nonexistant) '" & sMenu & "' menu:" & vbNewLine & vbNewLine & vbTab & "With Caption: " & sCaption & vbNewLine & vbNewLine & vbTab & "At position: " & nBefore
       Exit Function
    End If

On Error GoTo EH_Wizard_AddButton
    If moVBInst.CommandBars(sMenu).Controls.Count < nBefore Then
       If moVBInst.CommandBars(sMenu).Controls.Count > 0 Then
          nBefore = moVBInst.CommandBars(sMenu).Controls.Count
       Else
          ' No where to put that button
          MsgBox "Hmm... There doesn't seem to be anywhere I can place the following Slice and Dice button on the '" & sMenu & "' menu:" & vbNewLine & vbNewLine & vbTab & "With Caption: " & sCaption & vbNewLine & vbNewLine & vbTab & "At position: " & nBefore
          Exit Function
       End If
    End If

    Set mcb = moVBInst.CommandBars(sMenu).Controls.Add(before:=nBefore)
    If Not mcb Is Nothing Then
       mcb.BeginGroup = bBeginGroup
       If Len(sCaption) > 0 Then mcb.Caption = sCaption
       If nButtonFace <> 0 Then mcb.FaceId = nButtonFace
On Error Resume Next
       If moVBInst.CommandBars(sMenu).Controls.Count >= nBefore + 1 Then
          moVBInst.CommandBars(sMenu).Controls(nBefore + 1).BeginGroup = bNextBeginGroup
On Error GoTo EH_Wizard_AddButton
       End If
       Set AddButton = moVBInst.Events.CommandBarEvents(mcb)
    End If
    
EH_Wizard_AddButton_Continue:
    bInHereAlready = False
    Exit Function

EH_Wizard_AddButton:
    MsgBox "Error occured in:" & vbNewLine & vbTab & "Module: Wizard" & vbNewLine & vbTab & "Procedure: AddButton" & vbNewLine & vbNewLine & Err.Description & vbNewLine & vbNewLine & vbTab & "Menu: " & sMenu & vbNewLine & vbNewLine & vbTab & "Caption: " & sCaption & vbNewLine & vbNewLine & vbTab & "At position: " & nBefore
    Resume EH_Wizard_AddButton_Continue

    Resume
End Function

Private Sub AltChangeToHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ChangeToHandler_Click CommandBarControl, handled, CancelDefault
End Sub

Private Sub BarHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    MenuHandler_Click CommandBarControl, handled, CancelDefault
End Sub

Private Sub ChangeToHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim CurControl  As Object 'VBControl
    Dim NewControl  As Object 'VBControl
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
    With moVBInst.SelectedVBComponent.Designer
         Sandy.InsertionCancelled = False
         For Each CurControl In .SelectedVBControls
             If Sandy.InsertionCancelled Then Exit For
             With CurControl
                  If Not Sandy.UI.SetInternalCurrentTemplate("Change from - " & .ClassName) Then
                     If MsgBox("Please set up a 'Change to' template named" & vbCr & vbTab & "'Change from - " & .ClassName & "'" & vbCr & "before using this function on the '" & .ClassName & "' type control." & vbCr & vbCr & "Select Yes to create the new template." & vbCr & "Select No to abort.", vbYesNo) = vbNo Then
                        Exit Sub
                     End If

                     Sandy.UI.Visible = True
                     Sandy.UI.NewTemplate True, "Change from - " & .ClassName
                     Exit Sub
                  End If

                  If .ClassName = sLastClassName Then
                     sProgID = sLastChoice
                  Else
                     sChoices = ""
                     If Sandy.UI.SliceAndDice.Categorys("Change from").Templates("All Types") Is Nothing Or InStr(1, Sandy.UI.InternalCurrentTemplate.memoCodeAtTop, "NoAllTypes", vbTextCompare) Then
                        sTemplate = Sandy.UI.InternalCurrentTemplate.memoCodeAtTop
                     Else
                        With Sandy.UI.SliceAndDice.Categorys("Change from").Templates("All Types")
                             If Right(.memoCodeAtTop, 2) = vbNewLine Then
                                sTemplate = .memoCodeAtTop & Sandy.UI.InternalCurrentTemplate.memoCodeAtTop
                             Else
                                sTemplate = .memoCodeAtTop & vbNewLine & Sandy.UI.InternalCurrentTemplate.memoCodeAtTop
                             End If
                        End With
                     End If

                     Do Until Len(sTemplate) = 0
                        sLine = Sandy.sGetToken(sTemplate, 1, vbNewLine)
                        sTemplate = Sandy.sAfter(sTemplate, 1, vbNewLine)
                        If Left(sLine, 2) = "**" Then
                           sChoices = sChoices & ";" & Mid(sLine, 4)
                        End If
                     Loop
                     sChoices = Mid(sChoices, 2)
                     sProgID = Sandy.sChoose(sChoices)
                  End If
                  If Len(sProgID) > 0 Then
                     sLastClassName = .ClassName
                     sLastChoice = sProgID
                     
                     If Sandy.UI.SliceAndDice.Categorys("Change from").Templates("All Types") Is Nothing Or InStr(1, Sandy.UI.InternalCurrentTemplate.memoCodeAtTop, "NoAllTypes", vbTextCompare) Then
                        sTemplate = Sandy.UI.InternalCurrentTemplate.memoCodeAtTop
                     Else
                        With Sandy.UI.SliceAndDice.Categorys("Change from").Templates("All Types")
                             If Right(.memoCodeAtTop, 2) = vbNewLine Then
                                sTemplate = .memoCodeAtTop & Sandy.UI.InternalCurrentTemplate.memoCodeAtTop
                             Else
                                sTemplate = .memoCodeAtTop & vbNewLine & Sandy.UI.InternalCurrentTemplate.memoCodeAtTop
                             End If
                        End With
                     End If
                     bInTemplate = False
                     Do Until Len(sTemplate) = 0
                        sLine = Sandy.sGetToken(sTemplate, 1, vbNewLine)
                        sTemplate = Sandy.sAfter(sTemplate, 1, vbNewLine)
                        If sLine = "** " & sProgID Then
                           bInTemplate = True
                           sTemplate = Sandy.sGetToken(sTemplate, 1, vbNewLine & "** ")
                           asaVar("ClassName") = CurControl.ClassName
                           asaVar("ProgID") = CurControl.ProgId
                           If Not Sandy.UI.FillTemplateWithUserInput(asaVar, sTemplate, sCodeToInsert, "GROUP MODIFY") Then
                              Exit Sub
                           End If
                           sTemplate = sCodeToInsert
                        ElseIf Left(sLine, 2) = "**" And bInTemplate = True Then
                           bInTemplate = False
                           sTemplate = ""
                        ElseIf bInTemplate = True Then
                           Select Case UCase(Sandy.sGetToken(sLine))
                                  Case "ADDCONTROL"
                                        Set NewControl = .Container.ContainedVBControls.Add(Sandy.sAfter(sLine), CurControl, True)
                                        If NewControl Is Nothing Then
                                           MsgBox "That control has not been referenced yet. Please add a reference first.", vbInformation
                                           sTemplate = ""
                                        End If
                                        
                                  Case "SETNAME"
                                        NewControl.Properties("Name") = Sandy.sAfter(sLine) & Mid(CurControl.Properties("Name"), Len(Sandy.sAfter(sLine)))
                                        
                                  Case "COPYPROPERTY"
                                        NewControl.Properties(Sandy.sAfter(sLine)) = CurControl.Properties(Sandy.sAfter(sLine))
                                        
                                  Case "SETPROPERTY"
                                        sPropertyName = Sandy.sGetToken(Sandy.sAfter(sLine), 1, "=")
                                        If InStr(sPropertyName, ".") > 0 Then
                                           NewControl.Properties(Sandy.sGetToken(sPropertyName, 1, ".")).Value(Sandy.sGetToken(sPropertyName, 2, ".")).Value = Sandy.sAfter(sLine, 1, "=")
                                        Else
                                           NewControl.Properties(sPropertyName) = Sandy.sAfter(sLine, 1, "=")
                                        End If
                                        If Err.Number <> 0 Then
                                           MsgBox "Warning: Failed to set the property '" & sPropertyName & "' to '" & Sandy.sAfter(sLine, 1, "=") & vbCr & vbTab & Err.Description
                                        End If
                                       'NewControl.Properties(Sandy.sGetToken(Sandy.sAfter(sLine), 1, "=")) = Sandy.sAfter(sLine, 1, "=")
                                        
                                  Case "SETCURRENTPROPERTY"
                                        Err.Clear
                                        sPropertyName = Sandy.sGetToken(Sandy.sAfter(sLine), 1, "=")
                                        If InStr(sPropertyName, ".") > 0 Then
                                           CurControl.Properties(Sandy.sGetToken(sPropertyName, 1, ".")).Value(Sandy.sGetToken(sPropertyName, 2, ".")).Value = Sandy.sAfter(sLine, 1, "=")
                                        Else
                                           CurControl.Properties(sPropertyName) = Sandy.sAfter(sLine, 1, "=")
                                        End If
                                        If Err.Number <> 0 Then
                                           MsgBox "Warning: Failed to set the property '" & sPropertyName & "' to '" & Sandy.sAfter(sLine, 1, "=") & vbCr & vbTab & Err.Description
                                        End If
                                        
                                  Case "DELETECURRENTCONTROL"
                                        moVBInst.SelectedVBComponent.Designer.VBControls.Remove CurControl
                                        Set CurControl = Nothing

                                  Case "AUTOSPACECURRENT"
                                        CurControl.Properties(Sandy.sGetToken(Sandy.sAfter(sLine), 1, "=")) = Sandy.sInsertSpaces(CurControl.Properties(Sandy.sGetToken(Sandy.sAfter(sLine), 1, "=")))
                                        
                                  Case "AUTOSPACE"
                                        NewControl.Properties(Sandy.sGetToken(Sandy.sAfter(sLine), 1, "=")) = Sandy.sInsertSpaces(NewControl.Properties(Sandy.sGetToken(Sandy.sAfter(sLine), 1, "=")))
                                        
                                  Case "OFFSETFROMCURRENTLEFT", "OFFSET", "OFFSETLEFT"
                                        NewControl.Properties("Left") = CurControl.Properties("Left") + CurControl.Properties("Width") + IIf(Val(Sandy.sGetToken(sLine, 2)) = 0, 100, Val(Sandy.sGetToken(sLine, 2)))
                                        
                                  Case "OFFSETFROMCURRENTTOP", "OFFSETTOP"
                                        NewControl.Properties("Top") = CurControl.Properties("Top") + CurControl.Properties("Height") + IIf(Val(Sandy.sGetToken(sLine, 2)) = 0, 100, Val(Sandy.sGetToken(sLine, 2)))
                                        
                                  Case "QUEUE", "QUEUEINSERTION", "ADDTOQUEUE"
                                        asaVar("Queued").Value = asaVar("Queued").Value & Sandy.sAfter(sLine) & "~"
                                        
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
                                        PropertyStack(Sandy.sAfter(sLine)) = CurControl.Properties(Sandy.sAfter(sLine))

                                  Case "POPPORPERTYCURRENT", "POPPROPERTY", "POP"
                                        CurControl.Properties(Sandy.sAfter(sLine)) = PropertyStack(Sandy.sAfter(sLine))

                                  Case "INCREMENT", "INC"
                                  
                                  Case "SETVALUE", "SET"
                                       asaVar.Item(Sandy.sGetToken(Sandy.sAfter(sLine), 1, "=")) = Sandy.sAfter(sLine, 1, "=")

                                  Case "INSERTTEMPLATE", "INSERT"
                                        Sandy.InsertionCancelled = False
                                        Sandy.UI.DoInsertion asaVar, Sandy.sAfter(sLine)
                                        If Sandy.InsertionCancelled Then
                                           sTemplate = ""
                                        End If
                                        
                                  Case "GETPROPERTY", "GET"
                                       asaVar.Item(Sandy.sGetToken(Sandy.sAfter(sLine), 1, "=")) = CurControl.Properties(Sandy.sAfter(Sandy.sAfter(sLine), 1, "="))
                           End Select
                        End If
                     Loop
                  End If
             End With
         Next CurControl
    End With

    If Not Sandy.InsertionCancelled And Len(asaVar("Queued").Value) Then
       Sandy.UI.QueuedInsertions = asaVar("Queued").Value
    End If

    handled = True

    Set asaVar = Nothing
End Sub

Private Sub Class_Terminate()
    Set PropertyStack = Nothing
End Sub

Private Sub ExternalsHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error Resume Next
    Sandy.UI.ShowExternalsMenu

    handled = True
End Sub

Private Sub FavoritesHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error Resume Next
    Sandy.UI.FavoriteCalledFromIDE = True
    Sandy.UI.ShowFavMenu

    handled = True
End Sub

Private Property Get ISandyIDE_ActiveCodeModule() As SandySupport.ICodeModule
    Dim X As ICodeModule

    Set X = New CCodeModule
        Set X.Link = moVBInst.ActiveCodePane.CodeModule
        Set ISandyIDE_ActiveCodeModule = X
    Set X = Nothing
End Property

Private Property Get ISandyIDE_ActiveCodePane() As SandySupport.ICodePane
    Dim X As SandySupport.ICodePane

    If moVBInst.ActiveCodePane Is Nothing Then Exit Property

    Set X = New CCodePane
        Set X.Link = moVBInst.ActiveCodePane
        Set ISandyIDE_ActiveCodePane = X
    Set X = Nothing
End Property

Private Property Set ISandyIDE_ActiveProject(ByVal RHS As SandySupport.IProject)
    Set moVBInst.ActiveVBProject = RHS.Link
End Property

Private Property Get ISandyIDE_ActiveProject() As SandySupport.IProject
    Dim X As IProject
    Set X = New CProject
    Set X.Link = moVBInst.ActiveVBProject
    Set ISandyIDE_ActiveProject = X
    Set X = Nothing
End Property

Private Property Set ISandyIDE_ActiveWindow(ByVal RHS As Object)
    Set moVBInst.ActiveWindow = RHS
End Property

Private Property Get ISandyIDE_ActiveWindow() As Object
    Set ISandyIDE_ActiveWindow = moVBInst.ActiveWindow
End Property

Private Function ISandyIDE_AddComponent(ComponentType As Long) As IComponent
    Dim X As IComponent
    
    If ComponentType = 5 Then ComponentType = 3
    
    Set X = New CComponent
        Set X.Link = moVBInst.ActiveVBProject.VBComponents.Add(ComponentType)
        Set ISandyIDE_AddComponent = X
    Set X = Nothing
End Function

Private Property Set ISandyIDE_Components(ByVal RHS As Object)
    'Set moVBInst.ActiveVBProject.VBComponents = RHS
End Property

Private Property Get ISandyIDE_Components() As Object
    Set ISandyIDE_Components = moVBInst.ActiveVBProject.VBComponents
End Property

Private Property Set ISandyIDE_SelectedComponent(ByVal RHS As SandySupport.IComponent)
    Set moVBInst.SelectedVBComponent = RHS
End Property

Private Property Get ISandyIDE_SelectedComponent() As SandySupport.IComponent
    Dim X As IComponent
    
    If moVBInst.SelectedVBComponent Is Nothing Then Exit Property
    
    Set X = New CComponent
        Set X.Link = moVBInst.SelectedVBComponent
        Set ISandyIDE_SelectedComponent = X
    Set X = Nothing
End Property


Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error Resume Next
    Sandy.UI.Visible = Not Sandy.UI.Visible
    If Sandy.UI.Visible Then
       Sandy.UI.SetFocus
       Sandy.UI.ZOrder
    Else
       Sandy.UI.HideAllWindows
    End If

    handled = True
End Sub

Private Sub ShortcutHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    MenuHandler_Click CommandBarControl, handled, CancelDefault
End Sub

Public Sub OnConnection(ByVal VBInst As Object)
On Error Resume Next
    Dim lFirstButton As Long

    Set moVBInst = VBInst
    Set Sandy = CreateObject("SandyInstance.CSandy")
    If Sandy Is Nothing Then
       MsgBox "Can't create a 'SandyInstance.CSandy' object. Slice and Dice can't start."
       Exit Sub
    End If

    If Sandy.OnConnection(Me) Then
       Set MenuHandler = AddButton(mcbAddinButton, "Add-Ins", 4, "&Slice and Dice", 30, True, False)
    
       If GetSetting("SliceAndDice", "Last", "Show Paitbrush Icon", True) Then
          Set BarHandler = AddButton(mcbEditButton, "Standard", 1, "Slice and Dice", 30, False, True)
       End If
    
       If GetSetting("SliceAndDice", "Last", "Show On Module Right Click", True) Then
          If Not moVBInst.CommandBars("Code Window") Is Nothing Then
             lFirstButton = IIf(moVBInst.CommandBars("Code Window").Controls.Count > 4, 4, 1)
             Set FavoritesHandler = AddButton(mcbFavoritesButton, "Code Window", lFirstButton, "Insert a &Favorite", 30, True, False)
             Set ExternalsHandler = AddButton(mcbExternalsButton, "Code Window", lFirstButton + 1, "Call an &External", 30, False, False)
             Set ShortcutHandler = AddButton(mcbShortcutButton, "Code Window", lFirstButton + 2, "&Slice and Dice", 30, False, True)
          End If
       End If

       Set ChangeToHandler = AddButton(mcbChangeToButton, "MSForms Control", 1, "Slice and Dice Modify", 30, False, False)
       Set AltChangeToHandler = AddButton(mcbAltChangeToButton, "Format", 8, "&Slice and Dice Modify", 30, False, False)
    Else
       MsgBox "Slice and Dice GUI failed to initialize. No changes have been made to the VB IDE. Try reinstalling."
    End If
End Sub

Public Sub OnDisconnection()
On Error Resume Next
    Dim Cancel As Integer

    Sandy.OnDisconnection

  ' Make sure the edited Template (if one) is saved
    Sandy.UI.SaveTemplate
    Sandy.UI.Visible = False
    
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

    Set moVBInst = Nothing
    Set Sandy = Nothing
End Sub


