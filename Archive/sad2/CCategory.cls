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
1        If StrComp(m_smemoAttributes, sData) = 0 Then Exit Property

2        memoAttributes = sData
3        Modified = True
End Property

Public Property Get memoAttributes() As String
4        memoAttributes = m_smemoAttributes
End Property

Public Function Replace(ByVal sToFind As String, ByVal sToReplace As String) As Long
5        Dim lTotalCount As Long
6        Dim CurrTemplate As CTemplate

7        For Each CurrTemplate In Templates
8            lTotalCount = lTotalCount + CurrTemplate.Replace(sToFind, sToReplace)
9        Next CurrTemplate

10       Replace = lTotalCount
End Function

Public Function Duplicate(Optional ByVal ItemToFill As CCategory, Optional ByVal bCopyInformation As Boolean = True) As CCategory
11       On Error GoTo EH_CCategory_Duplicate
12       Dim CurTemplate As CTemplate
13       Dim sNewTemplateName As String

14       If ItemToFill Is Nothing Then
15           Set ItemToFill = New CCategory
16       End If

17       With ItemToFill
18           .CategoryType = CategoryType
19           .View = View
20           .Arrange = Arrange
21       End With

22       For Each CurTemplate In Templates
23           sNewTemplateName = ItemToFill.Key & gsCategoryTemplateDelimiter & CurTemplate.ShortTemplateName
24           With ItemToFill.Templates.Add(sNewTemplateName)
25               If bCopyInformation Then
26                   .Duplicate CurTemplate
27               End If
28           End With
29       Next CurTemplate

30 EH_CCategory_Duplicate_Continue:
31       Exit Function

32 EH_CCategory_Duplicate:
33       LogError "CCategory", "Duplicate", Err.Number, Err.Description, Erl
34       Resume EH_CCategory_Duplicate_Continue

35       Resume
End Function

Public Function Find(sToFind As String, Optional ByVal sLocationToStart As String, Optional ByVal lAreaToStartSearch As Long, Optional ByVal lByteToStartAt As Long) As Boolean
36       Set LastFound = Templates.Find(sToFind, sLocationToStart, lAreaToStartSearch, lByteToStartAt)
37       If LastFound Is Nothing Then
38           Find = False
39       Else
40           Find = True
41       End If
End Function

Public Property Get ID() As Long
42       ID = m_lID
End Property

Public Property Let ID(lData As Long)
43       m_lID = lData
End Property

Public Property Get Key() As String
Attribute Key.VB_UserMemId = 0
44       Key = m_sKey
End Property

Public Property Let Key(sData As String)
45       m_sKey = sData
46       Modified = True
End Property

Private Sub Class_Initialize()
47       Set Templates = New CTemplates
48       Set Templates.MyParent = Me
End Sub

Private Sub Class_Terminate()
49       Set LastFound = Nothing
50       Set Templates.MyParent = Nothing
51       Set Templates = Nothing
End Sub

Public Sub LoadProperties(rs As Recordset)
52       On Error Resume Next
53       With rs
54           CategoryType = ![CategoryType]
55           View = !View
56           Arrange = !Arrange
57           DateModified = !DateModified
58           DateCreated = !DateCreated
59       End With
End Sub

Public Sub SaveProperties(rs As Recordset)
60       On Error Resume Next
61       With rs
62           ![CategoryType] = CategoryType
63           ![View] = View
64           ![Arrange] = Arrange
65           ![memoAttributes] = zn(memoAttributes)
66       End With
End Sub

Public Property Get MyType() As String
67       MyType = "CCategorys"
End Property

Public Property Get CategoryType() As Long
68       CategoryType = m_lCategoryType
End Property

Public Property Let CategoryType(bData As Long)
69       If m_lCategoryType = bData Then Exit Property

70       m_lCategoryType = bData
71       Modified = True
End Property

Public Property Get View() As Long
72       View = m_iView
End Property

Public Property Let View(iData As Long)
73       If m_iView = iData Then Exit Property

74       m_iView = iData
75       Modified = True
End Property

Public Property Get Arrange() As Long
76       Arrange = m_iArrange
End Property

Public Property Let Arrange(iData As Long)
77       If m_iArrange = iData Then Exit Property

78       m_iArrange = iData
79       Modified = True
End Property

