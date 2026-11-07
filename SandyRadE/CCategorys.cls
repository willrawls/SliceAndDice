VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCategorys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy - Category Collection"
Option Explicit

Private mCol As Collection
Public MyParent As CSliceAndDice

Public Property Get ItemByLongTemplateName(sLongTemplateName As String) As CTemplate
1        On Error Resume Next
2        Dim sCategory As String
3        Dim sTemplate As String

4        If InStr(sLongTemplateName, gsCategoryTemplateDelimiter) = 0 Then Exit Property

5        sCategory = sGetToken(sLongTemplateName, 1, gsCategoryTemplateDelimiter)
6        sTemplate = sAfter(sLongTemplateName, 1, gsCategoryTemplateDelimiter)

7        If Len(sCategory) = 0 Or Len(sTemplate) = 0 Then Exit Property

8        Set ItemByLongTemplateName = Item(sCategory).Templates(sTemplate)
End Property

Public Function Find(sToFind As String, Optional ByVal sLocationToStart As String, Optional ByVal lAreaToStartSearchAt As Long, Optional ByVal lByteToStartAt As Long) As CCategory
9        Dim CurItem As CCategory
10       Dim sFoundAt As String

11       For Each CurItem In mCol
12           If CurItem.Find(sToFind, sLocationToStart, lAreaToStartSearchAt, lByteToStartAt) Then
13               Set Find = CurItem
14               Set CurItem = Nothing
15               Exit Function
16           End If
17           lAreaToStartSearchAt = 0
18           lByteToStartAt = 0
19       Next CurItem
End Function

Public Property Get AllKeys(Optional ByVal sDelimiter As String = gsSC, Optional ByVal CategoryType As Long = -1) As String
20       Dim CurItem As CCategory
21       Dim sAllKeys As String

22       If CategoryType > -1 Then
23           For Each CurItem In mCol
24               If CurItem.CategoryType = CategoryType Then
25                   sAllKeys = sAllKeys & CurItem.Key & sDelimiter
26               End If
27           Next CurItem
28       Else
29           For Each CurItem In mCol
30               sAllKeys = sAllKeys & CurItem.Key & sDelimiter
31           Next CurItem
32       End If

33       AllKeys = sAllKeys
End Property

Public Function Choose(Optional ByVal CategoryType As Long = -1) As String
34       Choose = sChoose(AllKeys(, CategoryType))
End Function

Public Sub FillList(lstCtrl As Object, Optional ByVal CategoryType As Long = -1)
35       Dim CurItem As CCategory
36       lstCtrl.Clear
37       If CategoryType > -1 Then
38           For Each CurItem In mCol
39               If CurItem.CategoryType = CategoryType Then
40                   lstCtrl.AddItem CurItem.Key
41               End If
42           Next CurItem
43       Else
44           For Each CurItem In mCol
45               lstCtrl.AddItem CurItem.Key
46           Next CurItem
47       End If
End Sub

Public Function Replace(ByVal sToFind As String, ByVal sToReplace As String) As Long
48       Dim lTotalCount As Long
49       Dim CurrCategory As CCategory

50       For Each CurrCategory In mCol
51           lTotalCount = lTotalCount + CurrCategory.Replace(sToFind, sToReplace)
52       Next CurrCategory

53       Replace = lTotalCount
End Function

Public Sub SaveAll(db As Database)
54       On Error GoTo Categorys_SaveModifiedErr
55       Dim rs As Recordset
56       Dim objCur As CCategory
57       Dim NextID As Long

58 Categorys_SaveModified_Start_Over:
59       For Each objCur In mCol
60           With objCur
61               If .Deleted Then
62                   If .ID > 0 Then
63                       Set rs = db.OpenRecordset(gsSelectFrom & "Category" & gsWhere & "CategoryID=" & .ID, dbOpenDynaset)
64                       rs.Delete
65                       Remove .Key
66                       GoTo Categorys_SaveModified_Start_Over
67                   End If
68               ElseIf .Modified Then
69                   If .ID > 0 Then                       ' Modify the existing record
70                       Set rs = db.OpenRecordset(gsSelectFrom & "Category" & gsWhere & "CategoryID=" & .ID, dbOpenDynaset)
71                       rs.Edit
72                       rs!DateModified = Now()
73                       .DateModified = rs!DateModified
74                   Else                                  ' Create a new record
                    'Set rs = db.OpenRecordset("SELECT MAX(CategoryID)+1 FROM Category")
                    'NextID = IIf(IsNull(rs.Fields(0).Value), 1, rs.Fields(0).Value)
                    'rs.Close
75                       Set rs = db.OpenRecordset(gsSelectFrom & "Category", dbOpenDynaset)
76                       rs.AddNew
77                       rs!DateCreated = Now()
78                       .DateCreated = rs!DateCreated
79                       rs!DateModified = .DateCreated
80                       .DateModified = .DateCreated
81                       .ID = rs!CategoryID
                    'rs!CategoryID = NextID
82                   End If
83                   rs!CategoryName = Left$(.Key, 254)
84                   .SaveProperties rs
85                   rs.Update
86                   rs.Close
87                   .Modified = False
88               End If
89               SaveChildren db, objCur
90           End With
91       Next objCur
92       Exit Sub

93 Categorys_SaveModifiedErr:
94       LogError "CCategorys", "SaveModified", Err.Number, Err.Description, Erl

95       Resume Next
96       Resume
End Sub

Public Sub GetAll(db As Database)
97       On Error GoTo Categorys_GetAllErr
98       Dim rs As Recordset
99       Dim objNew As CCategory
100      Set rs = db.OpenRecordset(gsSelectFrom & "Category ORDER BY CategoryType, CategoryName", dbOpenForwardOnly, dbReadOnly)
101      Clear
102      With rs
103          Do Until .EOF
104              Set objNew = Add(!CategoryName, !CategoryID)
105              With objNew
106                  .LoadProperties rs
107                  .Modified = False
                'LoadChildren db, objNew
108              End With
109              .MoveNext
110          Loop
111          .Close
112      End With
113      Exit Sub

114 Categorys_GetAllErr:
115      LogError "CCategorys", "GetAll", Err.Number, Err.Description, Erl
End Sub

Public Sub Clear()
116      On Error GoTo Categorys_ClearErr
117      Dim CurrCategory As CCategory
118      For Each CurrCategory In mCol
119          Set CurrCategory.MyParent = Nothing
120      Next CurrCategory

121      If mCol.Count > 0 Then
122          Set mCol = Nothing
123          Set mCol = New Collection
124      End If
125      Exit Sub

126 Categorys_ClearErr:
127      LogError "CCategorys", "Clear", Err.Number, Err.Description, Erl
End Sub

Public Function Add(ByVal Key As String, Optional ByVal ID As Long = -1, Optional ByVal sCategoryToCopy As String, Optional ByVal bCopyInformation As Boolean = True) As CCategory
' Create a new Category
128      Dim objNewMember As CCategory
129      Set objNewMember = New CCategory

    ' Set the properties passed into the method
130      With objNewMember
131          If ID = -1 Then ID = NextNegativeUnique
132          .ID = ID
133          .Key = Key
134          .Modified = True
135          Set .MyParent = Me
136      End With

137      If Len(sCategoryToCopy) = 0 Then
138          mCol.Add objNewMember, UCase$(Key)
139      Else
140          If mCol(sCategoryToCopy) Is Nothing Then
141              modGeneral.LogError "CCategorys", "Add", Err.Number, Err.Description & vbNewLine & "   (During Duplication of a " & gsCategory & gsPC, Erl
142          Else
143              If StrComp(Key, sCategoryToCopy, vbTextCompare) <> 0 Then
144                  Item(sCategoryToCopy).Duplicate objNewMember, bCopyInformation
145              End If
146              mCol.Add objNewMember, UCase$(Key)
147          End If
148      End If

    ' Return the object created
149      Set Add = objNewMember
150      Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As CCategory
Attribute Item.VB_UserMemId = 0
151      On Error Resume Next
152      Static CurItem As CCategory

153      Err.Clear
154      If VarType(vntIndexKey) = vbString Then
155          Set Item = mCol(UCase$(vntIndexKey))
156      Else
157          Set Item = mCol(vntIndexKey)
158      End If

159      If Err.Number <> 0 Then
160          Err.Clear
161          For Each CurItem In mCol
162              If CurItem.Key = vntIndexKey Or CurItem.ID = Val(vntIndexKey) Then
163                  Set Item = CurItem
164                  Set CurItem = Nothing
165                  Exit Property
166              End If
167          Next CurItem
168      End If
End Property

Public Property Get Count() As Long
169      Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
170      On Error Resume Next
171      If VarType(vntIndexKey) = vbString Then
172          Set mCol(UCase$(vntIndexKey)).MyParent = Nothing
173          mCol.Remove UCase$(vntIndexKey)
174      Else
175          Set mCol(vntIndexKey).MyParent = Nothing
176          mCol.Remove vntIndexKey
177      End If
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
178      Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
179      Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
180      Set mCol = Nothing
End Sub

Public Property Get MyType() As String
181      MyType = "modGeneral"
End Property

Public Sub LoadChildren(db As Database, objSetToLoad As CCategory)
182      With objSetToLoad
183          .Templates.GetAll db, .ID, .Key
184      End With
End Sub

Public Sub SaveChildren(db As Database, objSetToSave As CCategory)
185      With objSetToSave
186          .Templates.SaveAll db, .ID
187      End With
End Sub

