VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAssocArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Public Function Add(sKey As String) As CAssocArray
1        Dim objNewMember As CAssocArray

2        If Len(sKey) = 0 Then
3            Err.Raise vbObjectError + 0, "CAssocArrays", "Tryed to add an Assoc Array without a key."
4        End If

5        Set objNewMember = New CAssocArray
6        objNewMember.Section = sKey
7        mCol.Add objNewMember, sKey
8        Set Add = objNewMember
9        Set objNewMember = Nothing
End Function

Public Sub Clear(Optional ByVal bRecreateCollection As Boolean = True)
10       On Error Resume Next
11       Set mCol = Nothing
12       If bRecreateCollection Then Set mCol = New Collection
End Sub


Public Property Get Item(sIndexKey As String) As CAssocArray
Attribute Item.VB_UserMemId = 0
13       On Error GoTo EH_CAssocArrays_Item
14       Set Item = mCol(sIndexKey)

15 EH_CAssocArrays_Item_Continue:
16       Exit Property

17 EH_CAssocArrays_Item:
18       Set Item = Add(sIndexKey)
19       Err.Clear
20       Resume EH_CAssocArrays_Item_Continue
End Property

Public Property Get Count() As Long
21       Count = mCol.Count
End Property

Public Sub Load(sFilename As String, Optional bClearFirst As Boolean = True)
22       If Len(sFilename) = 0 Then
23           Err.Raise vbObjectError + 2, "CAssocArrays_LoadAll", "Tryed to load w/o a filename."
24       End If

25       Dim fh As Integer
26       Dim sLine As String
27       Dim CurAssocArray As CAssocArray

28       fh = FreeFile

29       If bClearFirst = True Then
30           Clear
31       End If

32       Open sFilename For Input Access Read As #fh
33       Do Until EOF(fh)
34           Line Input #fh, sLine
35           If Len(sLine) = 0 Then
            ' Skip it
36           ElseIf Left$(sLine, 1) = "[" Then
37               Set CurAssocArray = Add(Mid$(sLine, 2, Len(sLine) - 2))
38               With CurAssocArray
39                   Line Input #fh, sLine
40                   .ItemDelimiter = sAfter(sDenormalize(sLine), 1, gsE)
41                   Line Input #fh, sLine
42                   .KeyValueDelimiter = sAfter(sDenormalize(sLine), 1, gsE)
43                   Line Input #fh, sLine
44                   .FieldDelimiter = sAfter(sDenormalize(sLine), 1, gsE)
45               End With
46           Else
47               CurAssocArray.Add sGetToken(sLine, 1, gsE), sAfter(sLine, 1, gsE)
48           End If
49       Loop
50       Close #fh
51       Set CurAssocArray = Nothing
End Sub

Public Sub Remove(sIndexKey As String)
52       mCol.Remove sIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
53       Set NewEnum = mCol.[_NewEnum]
End Property

Public Sub Save(sFilename As String)
54       If Len(sFilename) = 0 Then
55           Err.Raise vbObjectError + 1, "CAssocArrays_SaveAll", "Tryed to save w/o a filename."
56       End If

57       Dim fh             As Integer
58       Dim CurAssocArray  As CAssocArray
59       Dim sOldItem       As String
60       Dim sOldKV         As String

61       fh = FreeFile
62       Open sFilename For Output Access Write As #fh
63       For Each CurAssocArray In mCol
64           With CurAssocArray
65               sOldItem = .ItemDelimiter
66               sOldKV = .KeyValueDelimiter
67               Print #fh, "[" & .Section & "]"
68               Print #fh, "Delimiter Item=" & sNormalize(.ItemDelimiter)
69               Print #fh, "Delimiter Key Value=" & sNormalize(.KeyValueDelimiter)
70               Print #fh, "Delimiter Field=" & sNormalize(.FieldDelimiter)
71               .ItemDelimiter = Chr$(13) & Chr$(10)
72               .KeyValueDelimiter = gsE
73               Print #fh, .All
74               .ItemDelimiter = sOldItem
75               .KeyValueDelimiter = sOldKV
76           End With
77       Next CurAssocArray
78       Close #fh
79       Set CurAssocArray = Nothing
End Sub

Private Sub Class_Initialize()
80       Clear
    ' LogEvent "CAssocArrays: Initialize"
End Sub


Private Sub Class_Terminate()
81       Set mCol = Nothing
    ' LogEvent "CAssocArrays: Terminate"
End Sub

Public Sub SaveXML(ByVal sFilename As String)
    Dim fh             As Integer
    Dim CurAssocArray  As CAssocArray
    Dim sOldItem       As String
    Dim sOldKV         As String

    If Len(sFilename) = 0 Then
        Err.Raise vbObjectError + 1, "CAssocArrays_SaveAll", "Tryed to save w/o a filename."
    End If
    
    fh = FreeFile
    Open sFilename For Output Access Write As #fh

    For Each CurAssocArray In mCol
        With CurAssocArray
            sOldItem = .ItemDelimiter
            sOldKV = .KeyValueDelimiter

            Print #fh, "<AssociativeArray>"
            Print #fh, vbTab & "<Section>" & .Section & "</Section>"
            Print #fh, vbTab & vbTab & "<Delimiters>"
            Print #fh, vbTab & vbTab & vbTab & "<Item>" & sNormalize(.ItemDelimiter) & "</Item>"
            Print #fh, vbTab & vbTab & vbTab & "<KeyValue>" & sNormalize(.KeyValueDelimiter) & "</KeyValue>"
            Print #fh, vbTab & vbTab & vbTab & "<Field>" & sNormalize(.FieldDelimiter) & "</Field>"
            Print #fh, vbTab & vbTab & "</Delimiters>"

            .ItemDelimiter = Chr$(13) & Chr$(10)
            .KeyValueDelimiter = gsE

            Print #fh, .All
            .ItemDelimiter = sOldItem
            .KeyValueDelimiter = sOldKV
        End With
    Next CurAssocArray
    Close #fh
    Set CurAssocArray = Nothing
End Sub


