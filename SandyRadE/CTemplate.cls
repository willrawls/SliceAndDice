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
1        On Error GoTo EH_CTemplate_Duplicate
2        With ItemToDuplicate
        'Parent = .Parent
3            ShortTemplateName = .ShortTemplateName
4            FileName = .FileName
5            Undeletable = .Undeletable
6            Locked = .Locked
7            IncludeInMenu = .IncludeInMenu
8            Favorite = .Favorite
9            memoCodeAtCursor = .memoCodeAtCursor
10           memoCodeAtTop = .memoCodeAtTop
11           memoCodeAtBottom = .memoCodeAtBottom
12           memoCodeToFile = .memoCodeToFile
13           memoAttributes = .memoAttributes
14       End With

15 EH_CTemplate_Duplicate_Continue:
16       Exit Sub

17 EH_CTemplate_Duplicate:
18       LogError "CTemplate", "Duplicate", Err.Number, Err.Description, Erl
19       Resume EH_CTemplate_Duplicate_Continue

20       Resume
End Sub

Public Property Let Favorite(bData As Boolean)
21       If m_bFavorite = bData Then Exit Property

22       m_bFavorite = bData
23       Modified = True
End Property

Public Property Get Favorite() As Boolean
24       Favorite = m_bFavorite
End Property

Public Function Find(sToFind As String, Optional ByVal lAreaToStartSearch As Long, Optional ByVal lByteToStartAt As Long) As Boolean
25       Dim lCurArea As Long
26       Dim lByteFound As Long
27       Dim sStringToSearch As String

28       If lByteToStartAt = 0 Then
29           For lCurArea = lAreaToStartSearch To 3
            Select Case lCurArea
                Case 0: sStringToSearch = memoCodeAtCursor
30                   Case 1: sStringToSearch = memoCodeAtTop
31                   Case 2: sStringToSearch = memoCodeAtBottom
32                   Case 3: sStringToSearch = memoCodeToFile
33               End Select
34               lByteFound = InStr(sStringToSearch, sToFind)
35               If lByteFound > 0 Then
36                   Find = True
37                   FoundInArea = lCurArea
38                   FoundAtByte = lByteFound
39                   Exit Function
40               End If
41           Next lCurArea
42       Else
43           For lCurArea = lAreaToStartSearch To 3
            Select Case lCurArea
                Case 0: sStringToSearch = Mid$(memoCodeAtCursor, lByteToStartAt)
44                   Case 1: sStringToSearch = Mid$(memoCodeAtTop, lByteToStartAt)
45                   Case 2: sStringToSearch = Mid$(memoCodeAtBottom, lByteToStartAt)
46                   Case 3: sStringToSearch = Mid$(memoCodeToFile, lByteToStartAt)
47               End Select
48               lByteFound = InStr(sStringToSearch, sToFind)
49               If lByteFound > 0 Then
50                   lByteFound = lByteFound + lByteToStartAt
51                   Find = True
52                   FoundInArea = lCurArea
53                   FoundAtByte = lByteFound
54                   Exit Function
55               End If
56           Next lCurArea
57       End If
End Function

Public Property Get ID() As Long
58       ID = m_lID
End Property

Public Property Let ID(lData As Long)
59       If m_lID = lData Then Exit Property
60       m_lID = lData
End Property

Public Property Get Key() As String
Attribute Key.VB_UserMemId = 0
61       Key = m_sKey
End Property

Public Property Let Key(sData As String)
62       If StrComp(m_sKey, sData) = 0 Then Exit Property
63       m_sKey = sData
64       Modified = True
End Property

Public Sub LoadProperties(rs As Recordset)
65       On Error Resume Next
66       With rs
67           Parent = ![CategoryID]
68           ShortTemplateName = nZ(![ShortTemplateName])
69           OriginalShortName = ShortTemplateName
70           FileName = nZ(![FileName])
71           Undeletable = ![Undeletable]
72           Favorite = ![Favorite]
73           Locked = ![Locked]
74           IncludeInMenu = ![IncludeInMenu]
75           memoCodeAtCursor = nZ(![memoCodeAtCursor].Value)
76           memoCodeAtTop = nZ(![memoCodeAtTop].Value)
77           memoCodeAtBottom = nZ(![memoCodeAtBottom].Value)
78           memoCodeToFile = nZ(![memoCodeToFile].Value)
79           memoAttributes = nZ(![memoAttributes].Value)
80           Revision = IIf(Val(vbNullString & ![RevisionCount]) = 0, 1, ![RevisionCount])
81           TimerInterval = nZ(![TimerInsertion].Value)
82           DateCreated = !DateCreated
83           DateModified = !DateModified
84       End With
End Sub

Public Property Get memoAttributes() As String
85       memoAttributes = m_smemoAttributes
End Property

Public Property Let memoAttributes(sData As String)
86       If StrComp(m_smemoAttributes, sData) = 0 Then Exit Property

87       m_smemoAttributes = sData
88       Modified = True
End Property


Public Function Replace(ByVal sToFind As String, ByVal sToReplace As String) As Long
89       Dim lTotalCount As Long

90       If InStr(m_smemoCodeAtCursor, sToFind) > 0 Then
91           lTotalCount = lTotalCount + 1
92           m_smemoCodeAtCursor = VBA.Replace(m_smemoCodeAtCursor, sToFind, sToReplace)
93           Modified = True
94       End If

95       If InStr(m_smemoCodeAtTop, sToFind) > 0 Then
96           lTotalCount = lTotalCount + 1
97           m_smemoCodeAtTop = VBA.Replace(m_smemoCodeAtTop, sToFind, sToReplace)
98           Modified = True
99       End If

100      If InStr(m_smemoCodeAtBottom, sToFind) > 0 Then
101          lTotalCount = lTotalCount + 1
102          m_smemoCodeAtBottom = VBA.Replace(m_smemoCodeAtBottom, sToFind, sToReplace)
103          Modified = True
104      End If

105      If InStr(m_smemoCodeToFile, sToFind) > 0 Then
106          lTotalCount = lTotalCount + 1
107          m_smemoCodeToFile = VBA.Replace(m_smemoCodeToFile, sToFind, sToReplace)
108          Modified = True
109      End If

110      If InStr(m_sFilename, sToFind) > 0 Then
111          lTotalCount = lTotalCount + 1
112          m_sFilename = VBA.Replace(m_sFilename, sToFind, sToReplace)
113          Modified = True
114      End If

115      Replace = lTotalCount
End Function

Public Sub SaveProperties(rs As Recordset)
116      On Error Resume Next
117      With rs
118          ![CategoryID] = Parent
119          ![ShortTemplateName] = zn(ShortTemplateName)
120          ![FileName] = zn(FileName)
121          ![Undeletable] = Undeletable
122          ![Locked] = Locked
123          ![IncludeInMenu] = IncludeInMenu
124          ![Favorite] = Favorite
125          ![memoCodeAtCursor] = zn(memoCodeAtCursor)
126          ![memoCodeAtTop] = zn(memoCodeAtTop)
127          ![memoCodeAtBottom] = zn(memoCodeAtBottom)
128          ![memoCodeToFile] = zn(memoCodeToFile)
129          ![memoAttributes] = zn(memoAttributes)
130          Revision = Revision + 1
131          ![RevisionCount] = Revision
132      End With
End Sub

Public Property Get MyType() As String
133      MyType = "CTemplates"
End Property

Public Property Get Parent() As Long
134      Parent = m_lParent
End Property

Public Property Let Parent(lData As Long)
135      If m_lParent = lData Then Exit Property

136      m_lParent = lData
137      Modified = True
End Property

Public Property Get ShortTemplateName() As String
138      ShortTemplateName = m_sShortTemplateName
End Property

Public Property Let ShortTemplateName(sData As String)
139      If StrComp(m_sShortTemplateName, sData) = 0 Then Exit Property

140      If sData = vbNullString Then
141          MsgBox "ShortTemplateName set to blank. Aborting"
142      Else
143          m_sShortTemplateName = sData
144      End If

145      Modified = True
End Property

Public Property Get FileName() As String
146      FileName = m_sFilename
End Property

Public Property Let FileName(sData As String)
147      If StrComp(m_sFilename, sData) = 0 Then Exit Property

148      m_sFilename = sData
149      Modified = True
End Property

Public Property Get Undeletable() As Boolean
150      Undeletable = m_bUndeletable
End Property

Public Property Let Undeletable(bData As Boolean)
151      If m_bUndeletable = bData Then Exit Property

152      m_bUndeletable = bData
153      Modified = True
End Property

Public Property Get Locked() As Boolean
154      Locked = m_bLocked
End Property

Public Property Let Locked(bData As Boolean)
155      If m_bLocked = bData Then Exit Property

156      m_bLocked = bData
157      Modified = True
End Property

Public Property Get IncludeInMenu() As Boolean
158      IncludeInMenu = m_bIncludeInMenu
End Property

Public Property Let IncludeInMenu(bData As Boolean)
159      If m_bIncludeInMenu = bData Then Exit Property

160      m_bIncludeInMenu = bData
161      Modified = True
End Property


Public Property Get memoCodeAtCursor() As String
162      memoCodeAtCursor = m_smemoCodeAtCursor
End Property

Public Property Let memoCodeAtCursor(sData As String)
163      If StrComp(m_smemoCodeAtCursor, sData) = 0 Then Exit Property

164      m_smemoCodeAtCursor = sData
165      Modified = True
End Property

Public Property Get memoCodeAtTop() As String
166      memoCodeAtTop = m_smemoCodeAtTop
End Property

Public Property Let memoCodeAtTop(sData As String)
167      If StrComp(m_smemoCodeAtTop, sData) = 0 Then Exit Property

168      m_smemoCodeAtTop = sData
169      Modified = True
End Property

Public Property Get memoCodeAtBottom() As String
170      memoCodeAtBottom = m_smemoCodeAtBottom
End Property

Public Property Let memoCodeAtBottom(sData As String)
171      If StrComp(m_smemoCodeAtBottom, sData) = 0 Then Exit Property

172      m_smemoCodeAtBottom = sData
173      Modified = True
End Property

Public Property Get memoCodeToFile() As String
174      memoCodeToFile = m_smemoCodeToFile
End Property

Public Property Let memoCodeToFile(sData As String)
175      If StrComp(m_smemoCodeToFile, sData) = 0 Then Exit Property

176      m_smemoCodeToFile = sData
177      Modified = True
End Property

