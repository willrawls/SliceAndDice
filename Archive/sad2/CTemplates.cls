VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy - Template Collection"
Option Explicit

Private mCol As Collection

Public MyParent As CCategory
Private mbNeverAccessed As Boolean
Public Function Find(sToFind As String, Optional ByVal sLocationToStart As String, Optional ByVal lAreaToStartSearchAt As Long, Optional ByVal lByteToStartAt As Long) As CTemplate
1        Dim CurItem As CTemplate
2        Dim sFoundAt As String

3        If Len(sLocationToStart) = 0 Then
4            For Each CurItem In mCol
5                If CurItem.Find(sToFind, lAreaToStartSearchAt, lByteToStartAt) Then
6                    Set Find = CurItem
7                    Set CurItem = Nothing
8                    Exit Function
9                End If
10               lAreaToStartSearchAt = 0
11               lByteToStartAt = 0
12           Next CurItem
13       Else
14           For Each CurItem In mCol
15               If CurItem.Find(sToFind, lAreaToStartSearchAt, lByteToStartAt) Then
16                   Set Find = CurItem
17                   Set CurItem = Nothing
18                   Exit Function
19               End If
20               lAreaToStartSearchAt = 0
21               lByteToStartAt = 0
22           Next CurItem
23       End If
End Function
Public Function Duplicate(Optional ByVal bCopyInformation As Boolean = True) As CTemplates
24       MsgBox "CTemplates.Duplicate : Function disabled"
    '    Dim Dupe As CTemplates
    '    Dim CurItem As CTemplate
    '
    '    Set Dupe = New CTemplates
    '
    '    For Each CurItem In mCol
    '        Dupe.Add CurItem.Key
    '        If bCopyInformation Then
    '           Dupe(CurItem.Key).Duplicate CurItem
    '        Else
    '           ShortTemplateName = CurItem.ShortTemplateName
    '        End If
    '    Next CurItem
    '
    '    Set Duplicate = Dupe
    '    Set Dupe = Nothing
End Function


Public Sub SaveAll(db As Database, CategoryID As Long)
25       On Error GoTo Templates_SaveModifiedErr
26       Dim rs As Recordset
27       Dim objCur As CTemplate
28       Dim NextID As Long

29 Templates_SaveModified_Start_Over:
30       For Each objCur In mCol
31           With objCur
32               If .Deleted Then
33                   If .ID > 0 Then
34                       Set rs = db.OpenRecordset(gsSelectFrom & "Template" & gsWhere & "TemplateID=" & .ID, dbOpenDynaset)
35                       rs.Delete
36                       Remove .Key
37                       GoTo Templates_SaveModified_Start_Over
38                   End If
39               ElseIf .Modified Then
40                   If Len(.Key) = 0 Or Len(.ShortTemplateName) = 0 Then
41                       modGeneral.LogError "CTemplates", "SaveAll", vbObjectError + 1001, "Probable error saving " & gsTemplate & ":" & gsEolTab & "ParentKey=" & .ParentKey & gsEolTab & "ShortTemplateName=" & .ShortTemplateName & gsEolTab & "OriginalShortName=" & .OriginalShortName & vbNewLine & "This " & gsTemplate & " will not be saved.", Erl
42                   Else
43                       If .ID > 0 Then                   ' Modify the existing record
44                           Set rs = db.OpenRecordset(gsSelectFrom & "Template" & gsWhere & "TemplateID=" & .ID, dbOpenDynaset)
45                           rs.Edit
46                           rs!DateModified = Now()
47                           .DateModified = rs!DateModified
48                       Else                              ' Create a new record
                        'Set rs = db.OpenRecordset("SELECT MAX(TemplateID)+1 FROM Template")
                        'NextID = IIf(IsNull(rs.Fields(0).Value), 1, rs.Fields(0).Value)
                        'rs.Close
49                           Set rs = db.OpenRecordset(gsSelectFrom & "Template", dbOpenDynaset)
50                           rs.AddNew
51                           rs!DateCreated = Now()
52                           .DateCreated = rs!DateCreated
53                           rs!DateModified = .DateCreated
54                           .DateModified = .DateCreated
55                           .ID = rs!TemplateID
                        'rs!TemplateID = NextID
56                           .Parent = CategoryID
57                           rs!CategoryID = CategoryID
58                       End If
59                       rs!TemplateName = Left$(.Key, 254)
60                       .SaveProperties rs
61                       rs.Update
62                       rs.Close
63                   End If
64                   .Modified = False
65               End If
66           End With
67       Next objCur
68       Exit Sub

69 Templates_SaveModifiedErr:
70       LogError "CTemplates", "SaveModified", Err.Number, Err.Description, Erl
71       Resume Next

72       Resume
End Sub

Public Sub GetAll(db As Database, CategoryID As Long, CategoryName As String)
73       On Error GoTo Templates_GetAllErr
74       Dim lHotKey As Integer
75       Dim lHotMod As Integer
76       Dim sHotKey As String

77       Dim rs As Recordset
78       Dim objNew As CTemplate

79       Set rs = db.OpenRecordset(gsSelectFrom & "Template" & gsWhere & "CategoryID=" & CategoryID & " ORDER BY ShortTemplateName", dbOpenForwardOnly, dbReadOnly)

80       Clear
81       With rs
82           Do Until .EOF
83               Set objNew = Add(!TemplateName, !TemplateID)
84               With objNew
85                   .LoadProperties rs
86                   .Modified = False
87               End With
88               .MoveNext
89           Loop
90           .Close
91       End With

92 Templates_GetAllErr_Continue:
93       Exit Sub

94 Templates_GetAllErr:
95       LogError "CTemplates", "GetAll", Err.Number, Err.Description, Erl
96       Resume Templates_GetAllErr_Continue

97       Resume
End Sub

Public Sub Clear()
98       On Error GoTo Templates_ClearErr
99       Dim CurrItem As CTemplate
100      For Each CurrItem In mCol
101          Set CurrItem.MyParent = Nothing
102      Next CurrItem

103      If mCol.Count > 0 Then
104          Set mCol = Nothing
105          Set mCol = New Collection
106      End If
107      Exit Sub

108 Templates_ClearErr:
109      LogError "CTemplates", "Clear", Err.Number, Err.Description, Erl
End Sub

Public Function Add(ByVal Key As String, Optional ByVal ID As Long = -1, Optional ByVal sKey As String) As CTemplate
' Create a new Template
110      Dim objNewMember As CTemplate
111      Set objNewMember = New CTemplate

    ' Set the properties passed into the method
112      With objNewMember
113          If ID = -1 Then ID = NextNegativeUnique
114          .ID = ID
115          .Key = Key
116          .ShortTemplateName = sAfter(Key, 1, gsCategoryTemplateDelimiter)
117          .OriginalShortName = .ShortTemplateName
118          .ParentKey = sGetToken(Key, 1, gsCategoryTemplateDelimiter)
119          .Modified = True
120          Set .MyParent = Me
121      End With

122      If Len(sKey) = 0 Then
123          mCol.Add objNewMember, UCase$(Key)
124      Else
125          mCol.Add objNewMember, UCase$(sKey)
126      End If

    ' Return the object created
127      Set Add = objNewMember
128      Set objNewMember = Nothing
End Function

Public Property Get AllKeys(Optional ByVal sDelimiter As String = gsSC) As String
129      Dim CurItem As CTemplate
130      Dim sAllKeys As String

131      For Each CurItem In mCol
132          With CurItem
133              sAllKeys = sAllKeys & .Key & sDelimiter
134          End With
135      Next CurItem

136      AllKeys = sAllKeys
End Property

Public Function Choose() As String
137      Choose = sChoose(AllKeys)
End Function

Public Property Get Item(vntIndexKey As Variant) As CTemplate
Attribute Item.VB_UserMemId = 0
138      On Error Resume Next
139      Static CurItem As CTemplate

140      If mbNeverAccessed Then
141          GetAll MyParent.MyParent.MyParent.db, MyParent.ID, MyParent.Key
142          mbNeverAccessed = False
143      End If

144      Err.Clear
145      If VarType(vntIndexKey) = vbString Then
146          Set Item = mCol(UCase$(vntIndexKey))
147      Else
148          Set Item = mCol(vntIndexKey)
149      End If

150      If Err.Number <> 0 Then
151          Err.Clear
152          For Each CurItem In mCol
153              If CurItem.Key = UCase$(vntIndexKey) Or CurItem.ID = Val(vntIndexKey) Or StrComp(CurItem.ShortTemplateName, vntIndexKey, 1) = 0 Or StrComp(CurItem.ParentKey & gsCategoryTemplateDelimiter & CurItem.OriginalShortName, vntIndexKey, 1) = 0 Then
154                  Set Item = CurItem
155                  Set CurItem = Nothing
156                  Exit Property
157              End If
158          Next CurItem
159      End If
End Property

Public Property Get Count() As Long
160      If mbNeverAccessed Then
161          GetAll MyParent.MyParent.MyParent.db, MyParent.ID, MyParent.Key
162          mbNeverAccessed = False
163      End If

164      Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
165      mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
166      If mbNeverAccessed Then
167          GetAll MyParent.MyParent.MyParent.db, MyParent.ID, MyParent.Key
168          mbNeverAccessed = False
169      End If

170      Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
171      Set mCol = New Collection
172      mbNeverAccessed = True
End Sub

Private Sub Class_Terminate()
173      Set mCol = Nothing
End Sub

Public Property Get MyType() As String
174      MyType = "CCategory"
End Property

