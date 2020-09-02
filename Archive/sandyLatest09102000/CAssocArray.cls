VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAssocArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy - Associative Array"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CAssoc"
Attribute VB_Ext_KEY = "Member0" ,"CAssoc"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Public mCol As Collection

Public FieldDelimiter As String
Public ItemDelimiter As String
Public KeyValueDelimiter As String
Public IconDelimiter As String
Public TagDelimiter As String
Public ChildDelimiter As String
Public EndChildDelimiter As String
Public Section As String

Public AddInOrder As Boolean

Private CurItem As CAssocItem
Public Property Get All() As String
1        Dim sAllKeyValues As String

2        If mCol Is Nothing Then Exit Property

3        For Each CurItem In mCol
4            sAllKeyValues = sAllKeyValues & CurItem.Key & KeyValueDelimiter & CurItem.Value & ItemDelimiter
5        Next CurItem
6        Set CurItem = Nothing

7        All = sAllKeyValues
End Property

Public Sub Clear(Optional ByVal bRecreateCollection As Boolean = True)
8        On Error Resume Next
9        Set mCol = Nothing
10       If bRecreateCollection Then Set mCol = New Collection
End Sub

Public Function Add(ByVal Key As String, Optional ByVal Value As String) As CAssocItem
11       Dim CurrMember As CAssocItem

12       Set CurItem = New CAssocItem
13       With CurItem
14           .Key = Key
15           .Value = Value
16       End With
17       On Error Resume Next
18       Err.Clear

19       If Not AddInOrder Then
20           mCol.Add CurItem, UCase$(Key)
21           If Err.Number = 0 Then
22               Set Add = CurItem
23           Else
24               Set Add = mCol(UCase$(Key))
25               Add.Value = Value
26           End If
27       Else
        ' Fit key into current list in alphabetic order
28           Key = UCase$(Key)
29           If mCol.Count = 0 Then
30               mCol.Add CurItem, Key
31           Else
32               For Each CurrMember In mCol
33                   If StrComp(Key, UCase$(CurrMember.Key)) < 0 Then
34                       mCol.Add CurItem, Key, UCase$(CurrMember.Key)
35                       Set Add = CurItem
36                       Set CurItem = Nothing
37                       Exit Function
38                       Err.Clear
39                   End If
40               Next CurrMember
41           End If

        ' Not found, Insert at the end of the collection
42           mCol.Add CurItem, Key
43           If Err.Number = 0 Then
44               Set Add = CurItem
45           Else
46               Set Add = mCol(Key)
47               Add.Value = Value
48           End If
49       End If
50       Set CurItem = Nothing
51       Err.Clear
End Function

Public Function Column(Optional ByVal lColumn As Long = 0) As String
52       Dim sColumnValues As String

53       If lColumn < 1 Then
54           For Each CurItem In mCol
55               sColumnValues = sColumnValues & CurItem.Value & ItemDelimiter
56           Next CurItem
57       Else
58           For Each CurItem In mCol
59               sColumnValues = sColumnValues & sGetToken(CurItem.Value, lColumn, FieldDelimiter) & ItemDelimiter
60           Next CurItem
61       End If
62       Set CurItem = Nothing

63       Column = sColumnValues
End Function

Public Sub FillList(lstCtrl As Object)
64       lstCtrl.Clear
65       For Each CurItem In mCol
66           lstCtrl.AddItem CurItem.Key
67       Next CurItem
68       Set CurItem = Nothing
End Sub

Friend Sub FillListView(lvwToFill As Object)
69       lvwToFill.ListItems.Clear
70       For Each CurItem In mCol
71           With lvwToFill.ListItems.Add(, CurItem.Key, CurItem.Key, "DocumentAlternate")
72               If Len(CurItem.Value) Then
73                   .SubItems(1) = CurItem.Value
74               End If
75           End With
76       Next CurItem
77       Set CurItem = Nothing
End Sub

'Public Sub FillTreeNode(tvwToFill As Object, nodCur As Object, sImage As String, sTagDelimiter As String, bExpanded As Boolean)
'    Dim nodX As Node
'    Dim tvwX As TreeView
'    Dim sKey As String
'    Dim sTag As String
'
'    Set tvwX = tvwToFill
'        Set nodX = nodCur
'
'            If Not nodX Is Nothing Then
'               For Each CurItem In mCol
'                   'If Len(sTagDelimiter) > 0 Then
'                      sKey = sGetToken(CurItem.Key, 1, sTagDelimiter)
'                      sTag = sGetToken(CurItem.Key, 2, sTagDelimiter)
'                   'Else
'                   '   sKey = CurItem.Key
'                   'End If
'                   With tvwX.Nodes.Add(nodX.Key, tvwChild, nodX.Key & "_" & sKey, sKey, sImage, sImage)
'                        .ExpandedImage = sImage
'                        .Expanded = bExpanded
'                        .Tag = sTag
'                   End With
'               Next CurItem
'            Else
'               For Each CurItem In mCol
'                   sKey = sGetToken(CurItem.Key, 1, sTagDelimiter)
'                   sTag = sGetToken(CurItem.Key, 2, sTagDelimiter)
'                   With tvwX.Nodes.Add(, , sKey, sKey, sImage, sImage)
'                        .ExpandedImage = sImage
'                        .Expanded = bExpanded
'                        .Tag = sTag
'                   End With
'               Next CurItem
'            End If
'            Set CurItem = Nothing
'        Set nodX = Nothing
'    Set tvwX = Nothing
'End Sub
'
Public Function FindKey(sValue As String) As String
78       For Each CurItem In mCol
79           If UCase$(CurItem.Value) = UCase$(sValue) Then
80               FindKey = CurItem.Key
81               Set CurItem = Nothing
82               Exit Function
83           End If
84       Next CurItem
End Function

Public Property Get Item(ByVal sIndexKey As String) As CAssocItem
Attribute Item.VB_UserMemId = 0
85       On Error GoTo EH_CAssocArray_Item
86       Set Item = mCol(UCase$(sIndexKey))

87 EH_CAssocArray_Item_Continue:
88       Exit Property

89 EH_CAssocArray_Item:
90       Set Item = Add(sIndexKey)
91       Err.Clear
92       Resume EH_CAssocArray_Item_Continue
End Property

Public Property Get Count() As Long
93       On Error Resume Next
94       If mCol Is Nothing Then
95           Count = 0
96       Else
97           Count = mCol.Count
98       End If
End Property


Public Function Load(ByVal sFilename As String)
99       If Len(sFilename) = 0 Then Exit Function

100      Dim fh As Long, sT As String
101      fh = FreeFile
102      Open sFilename For Input Access Read As #fh
103      Input #fh, sT
104      All = sT
105      Close #fh
End Function

Public Sub Remove(vntIndexKey As Variant)
106      On Error Resume Next
107      mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
108      Set NewEnum = mCol.[_NewEnum]
End Property


Public Function Save(ByVal sFilename As String)
109      Static fh As Long, sT As String

110      If Len(sFilename) = 0 Then Exit Function

111      fh = FreeFile
112      Open sFilename For Output Access Write As #fh
113      Print #fh, All
114      Close #fh
115      sT = vbNullString
End Function

Public Property Let All(ByVal sAllKeyValues As String)
116      Static sT As String
117      Clear
118      Do While Len(sAllKeyValues)
119          sT = sGetToken(sAllKeyValues, 1, ItemDelimiter)
120          If InStr(sT, KeyValueDelimiter) > 0 Then
121              Item(sGetToken(sT, 1, KeyValueDelimiter)).Value = sGetToken(sT, 2, KeyValueDelimiter)
122          Else
123              Item(sT).Value = vbNullString
124          End If
125          sAllKeyValues = sAfter(sAllKeyValues, 1, ItemDelimiter)
126      Loop
127      sT = vbNullString
End Property

Public Sub OrderedMultiAdd(ByVal sAllKeyValues As String, Optional ByVal bClearFirst As Boolean = True)
128      Static sT As String
129      If bClearFirst Then Clear

130      Do While Len(sAllKeyValues)
131          sT = sGetToken(sAllKeyValues, 1, ItemDelimiter)
132          Add CStr(mCol.Count + 1), sT
133          sAllKeyValues = sAfter(sAllKeyValues, 1, ItemDelimiter)
134      Loop
135      sT = vbNullString
End Sub

Private Sub Class_Initialize()
136      Set mCol = New Collection
137      ItemDelimiter = vbNewLine
138      FieldDelimiter = gsS
139      KeyValueDelimiter = gsE
140      TagDelimiter = "<TAG>"
141      IconDelimiter = "<ICON>"
142      ChildDelimiter = "<CHILD>"
143      EndChildDelimiter = "<ENDCHILD>"

144      Add "EOL", vbNewLine
145      Add "EOLTAB", gsEolTab
146      Add "EOL2", gs2EOL
147      Add "EOL2Tab", gs2EOL & vbTab
End Sub

Private Sub Class_Terminate()
148      Set mCol = Nothing
End Sub

Public Sub TreeToAll(tvwToGet As Object)
149      Dim sAll As String

150      Dim CurNode As Node
151      Dim tvwX As TreeView

152      Set tvwX = tvwToGet
153      Clear
154      For Each CurNode In tvwX.Nodes
155          With CurNode
156              If .Parent Is Nothing Then
157                  If Len(.Tag) = 0 Then
158                      sAll = sAll & ItemDelimiter & .Text & IconDelimiter & .Image
159                  Else
160                      sAll = sAll & ItemDelimiter & .Text & TagDelimiter & .Tag & IconDelimiter & .Image
161                  End If
162                  TreeToAll_AddChildren sAll, CurNode
163              End If
164          End With
165      Next CurNode
166      All = Mid$(sAll, 3)
167      sAll = vbNullString
168      Set CurNode = Nothing
169      Set tvwX = Nothing
End Sub

Private Sub TreeToAll_AddChildren(sAll As String, CurNode As Node)
170      Dim CurChild As Node
171      Set CurChild = CurNode.Child
172      If Not CurChild Is Nothing Then
173          If Len(CurChild.Tag) = 0 Then
174              sAll = sAll & ItemDelimiter & ChildDelimiter & CurChild.Text & IconDelimiter & CurChild.Image
175          Else
176              sAll = sAll & ItemDelimiter & ChildDelimiter & CurChild.Text & TagDelimiter & CurChild.Tag & IconDelimiter & CurChild.Image
177          End If

178          If Not CurChild.Child Is Nothing Then
179              TreeToAll_AddChildren sAll, CurChild
180          End If

181          Set CurChild = CurChild.Next

182          Do Until CurChild Is Nothing
183              If Len(CurChild.Tag) = 0 Then
184                  sAll = sAll & ItemDelimiter & CurChild.Text & IconDelimiter & CurChild.Image
185              Else
186                  sAll = sAll & ItemDelimiter & CurChild.Text & TagDelimiter & CurChild.Tag & IconDelimiter & CurChild.Image
187              End If
188              If Not CurChild.Child Is Nothing Then
189                  TreeToAll_AddChildren sAll, CurChild
190              End If
191              Set CurChild = CurChild.Next
192          Loop
193          sAll = sAll & EndChildDelimiter
194      End If
195      Set CurChild = Nothing
End Sub

Public Sub FillTreeNode(tvwToFill As Object, nodCur As Object, sImage As String, bExpanded As Boolean)
196      Dim nodX As Node
197      Dim PrevNode As Node
198      Dim tvwX As TreeView
199      Dim sKey As String
200      Dim sTag As String
201      Dim sIcon As String
202      Dim nChildDelimiter As Integer
203      Dim nEndChildDelimiter As Integer

204      Dim NodeStack() As Node
205      ReDim NodeStack(0 To 0) As Node

206      nChildDelimiter = Len(ChildDelimiter)
207      nEndChildDelimiter = Len(EndChildDelimiter)

208      Set tvwX = tvwToFill
209      Set nodX = nodCur
210      For Each CurItem In mCol
211          sKey = sGetToken(CurItem.Key, 1, IconDelimiter)
212          sIcon = sGetToken(CurItem.Key, 2, IconDelimiter)
213          sTag = sGetToken(sKey, 2, TagDelimiter)
214          sKey = sGetToken(sKey, 1, TagDelimiter)
215          If Len(sIcon) = 0 Then sIcon = sImage

216          If nChildDelimiter > 0 Then
217              If Left$(sKey, nChildDelimiter) = ChildDelimiter Then
218                  sKey = Mid$(sKey, nChildDelimiter + 1)
219                  If Not PrevNode Is Nothing Then
220                      If UBound(NodeStack) > 0 Then
221                          ReDim Preserve NodeStack(1 To UBound(NodeStack) + 1) As Node
222                      Else
223                          ReDim NodeStack(1 To 1) As Node
224                      End If
225                      Set NodeStack(UBound(NodeStack)) = nodX
226                      Set nodX = PrevNode
227                  End If
228              End If
229          End If
230          On Error Resume Next
231          If nodX Is Nothing Then
232              Set PrevNode = tvwX.Nodes.Add(, , sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter))
233          Else
234              Set PrevNode = tvwX.Nodes.Add(nodX.Key, tvwChild, nodX.Key & "_" & sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter))
235          End If
236          With PrevNode
237              .ExpandedImage = sGetToken(sIcon, 1, EndChildDelimiter)
238              .Expanded = bExpanded
239              .Tag = sTag
240          End With

241          If nEndChildDelimiter > 0 Then
242              Do While Right$(sIcon, nEndChildDelimiter) = EndChildDelimiter
243                  sIcon = Left$(sIcon, Len(sIcon) - nEndChildDelimiter)
244                  Set nodX = NodeStack(UBound(NodeStack))
245                  If UBound(NodeStack) > 1 Then
246                      ReDim Preserve NodeStack(1 To UBound(NodeStack) - 1) As Node
247                  Else
248                      ReDim NodeStack(0 To 0) As Node
249                  End If
250              Loop
251          End If
252      Next CurItem
253      Set CurItem = Nothing
254      Set PrevNode = Nothing
255      Set nodX = Nothing
256      Set tvwX = Nothing
End Sub


Public Property Get XML() As String
1        Dim sAllKeyValues As String

2        If mCol Is Nothing Then Exit Property

3        For Each CurItem In mCol
4            sAllKeyValues = sAllKeyValues & CurItem.Key & KeyValueDelimiter & CurItem.Value & ItemDelimiter
5        Next CurItem
6        Set CurItem = Nothing

7        All = sAllKeyValues
End Property


