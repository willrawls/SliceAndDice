VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy - Template"
Option Explicit

Public MyParent As CTemplates

Public Modified As Boolean
Public Deleted As Boolean

Private m_lID As Long
Private m_sKey As String
Private m_smemoCodeAtCursor As String
Private m_smemoCodeAtTop As String
Private m_smemoCodeAtBottom As String
Private m_smemoCodeToFile As String
Private m_smemoAttributes As String
Private m_bIncludeInMenu As Boolean
Private m_bLocked As Boolean
Private m_bUndeletable As Boolean
Private m_bFavorite As Boolean
Private m_sFilename As String
Private m_sShortTemplateName As String
Private m_lParent As Long

' After thoughts
Public ParentKey As String
Public OriginalShortName As String
Public Selected As Boolean
Public Revision As Long
Public TimerInterval As String
Public FoundInArea As Long
Public FoundAtByte As Long

Public DateCreated As Date
Public DateModified As Date

Public Sub Duplicate(ItemToDuplicate As CTemplate)
On Error GoTo EH_CTemplate_Duplicate
    With ItemToDuplicate
        'Parent = .Parent
         ShortTemplateName = .ShortTemplateName
         Filename = .Filename
         Undeletable = .Undeletable
         Locked = .Locked
         IncludeInMenu = .IncludeInMenu
         Favorite = .Favorite
         memoCodeAtCursor = .memoCodeAtCursor
         memoCodeAtTop = .memoCodeAtTop
         memoCodeAtBottom = .memoCodeAtBottom
         memoCodeToFile = .memoCodeToFile
         memoAttributes = .memoAttributes
    End With

EH_CTemplate_Duplicate_Continue:
    Exit Sub

EH_CTemplate_Duplicate:
    LogError "CTemplate", "Duplicate", Err.Number, Err.Description
    Resume EH_CTemplate_Duplicate_Continue

    Resume
End Sub

Public Property Let Favorite(bData As Boolean)
       If m_bFavorite = bData Then Exit Property

       m_bFavorite = bData
       Modified = True
End Property

Public Property Get Favorite() As Boolean
    Favorite = m_bFavorite
End Property

Public Function Find(sToFind As String, Optional ByVal lAreaToStartSearch As Long, Optional ByVal lByteToStartAt As Long) As Boolean
    Dim lCurArea As Long
    Dim lByteFound As Long
    Dim sStringToSearch As String

    If lByteToStartAt = 0 Then
       For lCurArea = lAreaToStartSearch To 3
           Select Case lCurArea
                  Case 0:  sStringToSearch = memoCodeAtCursor
                  Case 1:  sStringToSearch = memoCodeAtTop
                  Case 2:  sStringToSearch = memoCodeAtBottom
                  Case 3:  sStringToSearch = memoCodeToFile
           End Select
           lByteFound = InStr(sStringToSearch, sToFind)
           If lByteFound > 0 Then
              Find = True
              FoundInArea = lCurArea
              FoundAtByte = lByteFound
              Exit Function
           End If
       Next lCurArea
    Else
       For lCurArea = lAreaToStartSearch To 3
           Select Case lCurArea
                  Case 0:  sStringToSearch = Mid$(memoCodeAtCursor, lByteToStartAt)
                  Case 1:  sStringToSearch = Mid$(memoCodeAtTop, lByteToStartAt)
                  Case 2:  sStringToSearch = Mid$(memoCodeAtBottom, lByteToStartAt)
                  Case 3:  sStringToSearch = Mid$(memoCodeToFile, lByteToStartAt)
           End Select
           lByteFound = InStr(sStringToSearch, sToFind)
           If lByteFound > 0 Then
              lByteFound = lByteFound + lByteToStartAt
              Find = True
              FoundInArea = lCurArea
              FoundAtByte = lByteFound
              Exit Function
           End If
       Next lCurArea
    End If
End Function

Public Property Get ID() As Long
       ID = m_lID
End Property

Public Property Let ID(lData As Long)
       If m_lID = lData Then Exit Property
       m_lID = lData
End Property

Public Property Get Key() As String
Attribute Key.VB_UserMemId = 0
       Key = m_sKey
End Property

Public Property Let Key(sData As String)
       If StrComp(m_sKey, sData) = 0 Then Exit Property
       m_sKey = sData
       Modified = True
End Property

Public Sub LoadProperties(rs As Recordset)
On Error Resume Next
       With rs
            Parent = ![CategoryID]
            ShortTemplateName = nZ(![ShortTemplateName])
            OriginalShortName = ShortTemplateName
            Filename = nZ(![Filename])
            Undeletable = ![Undeletable]
            Favorite = ![Favorite]
            Locked = ![Locked]
            IncludeInMenu = ![IncludeInMenu]
            memoCodeAtCursor = nZ(![memoCodeAtCursor].Value)
            memoCodeAtTop = nZ(![memoCodeAtTop].Value)
            memoCodeAtBottom = nZ(![memoCodeAtBottom].Value)
            memoCodeToFile = nZ(![memoCodeToFile].Value)
            memoAttributes = nZ(![memoAttributes].Value)
            Revision = IIf(![RevisionCount] = 0 Or IsNull(![RevisionCount]), 1, ![RevisionCount])
            TimerInterval = nZ(![TimerInsertion].Value)
            DateCreated = !DateCreated
            DateModified = !DateModified
       End With
End Sub

Public Property Get memoAttributes() As String
    memoAttributes = m_smemoAttributes
End Property

Public Property Let memoAttributes(sData As String)
       If StrComp(m_smemoAttributes, sData) = 0 Then Exit Property

       m_smemoAttributes = sData
       Modified = True
End Property


Public Function Replace(ByVal sToFind As String, ByVal sToReplace As String) As Long
    Dim lTotalCount As Long
    
    If InStr(m_smemoCodeAtCursor, sToFind) > 0 Then
       lTotalCount = lTotalCount + 1
       m_smemoCodeAtCursor = VBA.Replace(m_smemoCodeAtCursor, sToFind, sToReplace)
       Modified = True
    End If
    
    If InStr(m_smemoCodeAtTop, sToFind) > 0 Then
       lTotalCount = lTotalCount + 1
       m_smemoCodeAtTop = VBA.Replace(m_smemoCodeAtTop, sToFind, sToReplace)
       Modified = True
    End If
    
    If InStr(m_smemoCodeAtBottom, sToFind) > 0 Then
       lTotalCount = lTotalCount + 1
       m_smemoCodeAtBottom = VBA.Replace(m_smemoCodeAtBottom, sToFind, sToReplace)
       Modified = True
    End If
    
    If InStr(m_smemoCodeToFile, sToFind) > 0 Then
       lTotalCount = lTotalCount + 1
       m_smemoCodeToFile = VBA.Replace(m_smemoCodeToFile, sToFind, sToReplace)
       Modified = True
    End If
    
    If InStr(m_sFilename, sToFind) > 0 Then
       lTotalCount = lTotalCount + 1
       m_sFilename = VBA.Replace(m_sFilename, sToFind, sToReplace)
       Modified = True
    End If

    Replace = lTotalCount
End Function

Public Sub SaveProperties(rs As Recordset)
On Error Resume Next
       With rs
            ![CategoryID] = Parent
            ![ShortTemplateName] = zn(ShortTemplateName)
            ![Filename] = zn(Filename)
            ![Undeletable] = Undeletable
            ![Locked] = Locked
            ![IncludeInMenu] = IncludeInMenu
            ![Favorite] = Favorite
            ![memoCodeAtCursor] = zn(memoCodeAtCursor)
            ![memoCodeAtTop] = zn(memoCodeAtTop)
            ![memoCodeAtBottom] = zn(memoCodeAtBottom)
            ![memoCodeToFile] = zn(memoCodeToFile)
            ![memoAttributes] = zn(memoAttributes)
            Revision = Revision + 1
            ![RevisionCount] = Revision
       End With
End Sub

Public Property Get MyType() As String
    MyType = "CTemplates"
End Property

Public Property Get Parent() As Long
       Parent = m_lParent
End Property

Public Property Let Parent(lData As Long)
       If m_lParent = lData Then Exit Property

       m_lParent = lData
       Modified = True
End Property

Public Property Get ShortTemplateName() As String
       ShortTemplateName = m_sShortTemplateName
End Property

Public Property Let ShortTemplateName(sData As String)
       If StrComp(m_sShortTemplateName, sData) = 0 Then Exit Property

       If sData = vbNullString Then
          MsgBox "ShortTemplateName set to blank. Aborting"
       Else
          m_sShortTemplateName = sData
       End If
       
       Modified = True
End Property

Public Property Get Filename() As String
       Filename = m_sFilename
End Property

Public Property Let Filename(sData As String)
       If StrComp(m_sFilename, sData) = 0 Then Exit Property

       m_sFilename = sData
       Modified = True
End Property

Public Property Get Undeletable() As Boolean
       Undeletable = m_bUndeletable
End Property

Public Property Let Undeletable(bData As Boolean)
       If m_bUndeletable = bData Then Exit Property

       m_bUndeletable = bData
       Modified = True
End Property

Public Property Get Locked() As Boolean
       Locked = m_bLocked
End Property

Public Property Let Locked(bData As Boolean)
       If m_bLocked = bData Then Exit Property

       m_bLocked = bData
       Modified = True
End Property

Public Property Get IncludeInMenu() As Boolean
       IncludeInMenu = m_bIncludeInMenu
End Property

Public Property Let IncludeInMenu(bData As Boolean)
       If m_bIncludeInMenu = bData Then Exit Property

       m_bIncludeInMenu = bData
       Modified = True
End Property


Public Property Get memoCodeAtCursor() As String
       memoCodeAtCursor = m_smemoCodeAtCursor
End Property

Public Property Let memoCodeAtCursor(sData As String)
       If StrComp(m_smemoCodeAtCursor, sData) = 0 Then Exit Property

       m_smemoCodeAtCursor = sData
       Modified = True
End Property

Public Property Get memoCodeAtTop() As String
       memoCodeAtTop = m_smemoCodeAtTop
End Property

Public Property Let memoCodeAtTop(sData As String)
       If StrComp(m_smemoCodeAtTop, sData) = 0 Then Exit Property

       m_smemoCodeAtTop = sData
       Modified = True
End Property

Public Property Get memoCodeAtBottom() As String
       memoCodeAtBottom = m_smemoCodeAtBottom
End Property

Public Property Let memoCodeAtBottom(sData As String)
       If StrComp(m_smemoCodeAtBottom, sData) = 0 Then Exit Property

       m_smemoCodeAtBottom = sData
       Modified = True
End Property

Public Property Get memoCodeToFile() As String
       memoCodeToFile = m_smemoCodeToFile
End Property

Public Property Let memoCodeToFile(sData As String)
       If StrComp(m_smemoCodeToFile, sData) = 0 Then Exit Property

       m_smemoCodeToFile = sData
       Modified = True
End Property

Private Sub Class_Initialize()

    ' LogEvent "CTemplate: Initialize"
End Sub


Private Sub Class_Terminate()

    ' LogEvent "CTemplate: Terminate"
End Sub


