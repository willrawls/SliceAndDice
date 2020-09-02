VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy - Category"
Option Explicit

Public Templates As CTemplates
Public MyParent As CCategorys

Public Modified As Boolean
Public Deleted As Boolean

Private m_lID As Long
Private m_sKey As String
Private m_iView As Long
Private m_lCategoryType As Long
Private m_iArrange As Long
Private m_smemoAttributes As String

Public LastFound As CTemplate

Public DateCreated As Date
Public DateModified As Date

Public Property Let memoAttributes(sData As String)
    If StrComp(m_smemoAttributes, sData) = 0 Then Exit Property
    
    memoAttributes = sData
    Modified = True
End Property

Public Property Get memoAttributes() As String
    memoAttributes = m_smemoAttributes
End Property

Public Function Replace(ByVal sToFind As String, ByVal sToReplace As String) As Long
    Dim lTotalCount As Long
    Dim CurrTemplate As CTemplate
    
    For Each CurrTemplate In Templates
        lTotalCount = lTotalCount + CurrTemplate.Replace(sToFind, sToReplace)
    Next CurrTemplate

    Replace = lTotalCount
End Function

Public Function Duplicate(Optional ByVal ItemToFill As CCategory, Optional ByVal bCopyInformation As Boolean = True) As CCategory
On Error GoTo EH_CCategory_Duplicate
    Dim CurTemplate As CTemplate
    Dim sNewTemplateName As String

    If ItemToFill Is Nothing Then
       Set ItemToFill = CreateObject("SandySupport.CCategory")
    End If

    With ItemToFill
         .CategoryType = CategoryType
         .View = View
         .Arrange = Arrange
    End With

    For Each CurTemplate In Templates
        sNewTemplateName = ItemToFill.Key & " - " & CurTemplate.ShortTemplateName
        With ItemToFill.Templates.Add(sNewTemplateName)
             If bCopyInformation Then
                .Duplicate CurTemplate
             End If
        End With
    Next CurTemplate

EH_CCategory_Duplicate_Continue:
    Exit Function

EH_CCategory_Duplicate:
    LogError "CCategory", "Duplicate", Err.Number, Err.Description
    Resume EH_CCategory_Duplicate_Continue

    Resume
End Function

Public Function Find(sToFind As String, Optional ByVal sLocationToStart As String, Optional ByVal lAreaToStartSearch As Long, Optional ByVal lByteToStartAt As Long) As Boolean
    Set LastFound = Templates.Find(sToFind, sLocationToStart, lAreaToStartSearch, lByteToStartAt)
    If LastFound Is Nothing Then
       Find = False
    Else
       Find = True
    End If
End Function

Public Property Get ID() As Long
       ID = m_lID
End Property

Public Property Let ID(lData As Long)
       m_lID = lData
End Property

Public Property Get Key() As String
Attribute Key.VB_UserMemId = 0
       Key = m_sKey
End Property

Public Property Let Key(sData As String)
       m_sKey = sData
       Modified = True
End Property

Private Sub Class_Initialize()
    Set Templates = CreateObject("SandySupport.CTemplates")
    Set Templates.MyParent = Me
    ' LogEvent "CCategory: Initialize"
End Sub

Private Sub Class_Terminate()
    Set LastFound = Nothing
    Set Templates.MyParent = Nothing
    Set Templates = Nothing
    ' LogEvent "CCategory: Terminate"
End Sub

Public Sub LoadProperties(rs As Recordset)
On Error Resume Next
       With rs
            CategoryType = ![CategoryType]
            View = !View
            Arrange = !Arrange
            DateModified = !DateModified
            DateCreated = !DateCreated
       End With
End Sub

Public Sub SaveProperties(rs As Recordset)
On Error Resume Next
       With rs
            ![CategoryType] = CategoryType
            ![View] = View
            ![Arrange] = Arrange
            ![memoAttributes] = zn(memoAttributes)
       End With
End Sub

Public Property Get MyType() As String
    MyType = "CCategorys"
End Property

Public Property Get CategoryType() As Long
       CategoryType = m_lCategoryType
End Property

Public Property Let CategoryType(bData As Long)
       If m_lCategoryType = bData Then Exit Property

       m_lCategoryType = bData
       Modified = True
End Property

Public Property Get View() As Long
       View = m_iView
End Property

Public Property Let View(iData As Long)
       If m_iView = iData Then Exit Property

       m_iView = iData
       Modified = True
End Property

Public Property Get Arrange() As Long
       Arrange = m_iArrange
End Property

Public Property Let Arrange(iData As Long)
       If m_iArrange = iData Then Exit Property

       m_iArrange = iData
       Modified = True
End Property

