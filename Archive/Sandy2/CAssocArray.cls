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
    Dim sAllKeyValues As String

    If mCol Is Nothing Then Exit Property

    For Each CurItem In mCol
        sAllKeyValues = sAllKeyValues & CurItem.Key & KeyValueDelimiter & CurItem.Value & ItemDelimiter
    Next CurItem
    Set CurItem = Nothing

    All = sAllKeyValues
End Property

Public Sub Clear(Optional ByVal bRecreateCollection As Boolean = True)
On Error Resume Next
    Set mCol = Nothing
    If bRecreateCollection Then
       Set mCol = New Collection
       ' mCol.CompareMode = TextCompare
    End If
End Sub


Public Function Add(ByVal Key As String, Optional ByVal Value As String) As CAssocItem
       Dim CurrMember As CAssocItem

       Set CurItem = CreateObject("SandySupport.CAssocItem")
           With CurItem
                .Key = Key
                .Value = Value
           End With
On Error Resume Next
           Err.Clear
           
           'If Not AddInOrder Then
              mCol.Add CurItem, UCase$(Key)
              If Err.Number = 0 Then
                 Set Add = CurItem
              Else
                 Set Add = mCol(UCase$(Key))
                 Add.Value = Value
              End If
           'Else
           ' ' Fit key into current list in alphabetic order
           '   Key = UCase$(Key)
           '   If mCol.Count = 0 Then
           '      mCol.Add CurItem, Key
           '   Else
           '      For Each CurrMember In mCol
           '          If StrComp(Key, UCase$(CurrMember.Key)) < 0 Then
           '             mCol.Add Key, CurItem
           '             Set Add = CurItem
           '             Set CurItem = Nothing
           '             Exit Function
           '             Err.Clear
           '          End If
           '      Next CurrMember
           '   End If

           ' ' Not found, Insert at the end of the collection
           '   mCol.Add CurItem, Key
           '   If Err.Number = 0 Then
           '      Set Add = CurItem
           '   Else
           '      Set Add = mCol(Key)
           '      Add.Value = Value
           '   End If
           'End If
       Set CurItem = Nothing
       Err.Clear
End Function

Public Function Column(Optional ByVal lColumn As Long = 0) As String
    Dim sColumnValues As String

    If lColumn < 1 Then
        For Each CurItem In mCol
            sColumnValues = sColumnValues & CurItem.Value & ItemDelimiter
        Next CurItem
    Else
        For Each CurItem In mCol
            sColumnValues = sColumnValues & sGetToken(CurItem.Value, lColumn, FieldDelimiter) & ItemDelimiter
        Next CurItem
    End If
    Set CurItem = Nothing

    Column = sColumnValues
End Function

Public Sub FillList(lstCtrl As Object)
    lstCtrl.Clear
    For Each CurItem In mCol
        lstCtrl.AddItem CurItem.Key
    Next CurItem
    Set CurItem = Nothing
End Sub

Public Sub FillListView(ByRef lvwIn As Object)
On Error Resume Next
    Dim lvwToFill As ListView
    Set lvwToFill = lvwIn
    lvwToFill.ListItems.Clear
    For Each CurItem In mCol
        With lvwToFill.ListItems.Add(, CurItem.Key, CurItem.Key, "DocumentAlternate")
             If Len(CurItem.Value) Then
                .SubItems(1) = CurItem.Value
             End If
        End With
    Next CurItem
    Set CurItem = Nothing
    Set lvwToFill = Nothing
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
    For Each CurItem In mCol
        If UCase$(CurItem.Value) = UCase$(sValue) Then
           FindKey = CurItem.Key
           Set CurItem = Nothing
           Exit Function
        End If
    Next CurItem
End Function

Public Property Get Item(sIndexKey As String) As CAssocItem
Attribute Item.VB_UserMemId = 0
On Error GoTo EH_CAssocArray_Item
       Set Item = mCol(UCase$(sIndexKey))

EH_CAssocArray_Item_Continue:
       Exit Property

EH_CAssocArray_Item:
       Set Item = Add(sIndexKey)
       Err.Clear
       Resume EH_CAssocArray_Item_Continue
       
       Resume
End Property

Public Property Get Count() As Long
On Error Resume Next
       If mCol Is Nothing Then
          Count = 0
       Else
          Count = mCol.Count
       End If
End Property


Public Function Load(ByVal sFilename As String)
    If Len(sFilename) = 0 Then Exit Function
    
    Dim fh As Long, sT As String
    fh = FreeFile
    Open sFilename For Input Access Read As #fh
         Input #fh, sT
         All = sT
    Close #fh
End Function

Public Sub Remove(vntIndexKey As Variant)
On Error Resume Next
       mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
       Set NewEnum = mCol.[_NewEnum]
End Property


Public Function Save(ByVal sFilename As String)
    Static fh As Long, sT As String

    If Len(sFilename) = 0 Then Exit Function
    
    fh = FreeFile
    Open sFilename For Output Access Write As #fh
         Print #fh, All
    Close #fh
    sT = vbNullString
End Function

Public Property Let All(ByVal sAllKeyValues As String)
    Static sT As String
    Clear
    Do While Len(sAllKeyValues)
       sT = sGetToken(sAllKeyValues, 1, ItemDelimiter)
       If InStr(sT, KeyValueDelimiter) > 0 Then
          Item(sGetToken(sT, 1, KeyValueDelimiter)).Value = sGetToken(sT, 2, KeyValueDelimiter)
       Else
          Item(sT).Value = vbNullString
       End If
       sAllKeyValues = sAfter(sAllKeyValues, 1, ItemDelimiter)
    Loop
    sT = vbNullString
End Property

Public Sub OrderedMultiAdd(ByVal sAllKeyValues As String, Optional ByVal bClearFirst As Boolean = True)
    Static sT As String
    If bClearFirst Then Clear

    Do While Len(sAllKeyValues)
       sT = sGetToken(sAllKeyValues, 1, ItemDelimiter)
       Add CStr(mCol.Count + 1), sT
       sAllKeyValues = sAfter(sAllKeyValues, 1, ItemDelimiter)
    Loop
    sT = vbNullString
End Sub

Private Sub Class_Initialize()
        Set mCol = New Collection
        ' mCol.CompareMode = TextCompare
        ItemDelimiter = vbNewLine
        FieldDelimiter = " "
        KeyValueDelimiter = "="
        TagDelimiter = "<TAG>"
        IconDelimiter = "<ICON>"
        ChildDelimiter = "<CHILD>"
        EndChildDelimiter = "<ENDCHILD>"

        Add "EOL", vbNewLine
        Add "EOLTAB", gsEolTab
        Add "EOL2", gs2EOL
        Add "EOL2Tab", gs2EOL & vbTab
    ' LogEvent "CAssocArray: Initialize"
End Sub

Private Sub Class_Terminate()
        Set mCol = Nothing
    ' LogEvent "CAssocArray: Terminate"
End Sub

Public Sub TreeToAll(tvwToGet As Object)
    Dim sAll As String

    Dim CurNode As Node
    Dim tvwX As TreeView

    Set tvwX = tvwToGet
        Clear
        For Each CurNode In tvwX.Nodes
            With CurNode
                 If .Parent Is Nothing Then
                    If Len(.Tag) = 0 Then
                       sAll = sAll & ItemDelimiter & .Text & IconDelimiter & .Image
                    Else
                       sAll = sAll & ItemDelimiter & .Text & TagDelimiter & .Tag & IconDelimiter & .Image
                    End If
                    TreeToAll_AddChildren sAll, CurNode
                 End If
            End With
        Next CurNode
        All = Mid$(sAll, 3)
        sAll = vbNullString
    Set CurNode = Nothing
    Set tvwX = Nothing
End Sub

Private Sub TreeToAll_AddChildren(sAll As String, CurNode As Node)
    Dim CurChild As Node
    Set CurChild = CurNode.Child
        If Not CurChild Is Nothing Then
           If Len(CurChild.Tag) = 0 Then
              sAll = sAll & ItemDelimiter & ChildDelimiter & CurChild.Text & IconDelimiter & CurChild.Image
           Else
              sAll = sAll & ItemDelimiter & ChildDelimiter & CurChild.Text & TagDelimiter & CurChild.Tag & IconDelimiter & CurChild.Image
           End If
           
           If Not CurChild.Child Is Nothing Then
              TreeToAll_AddChildren sAll, CurChild
           End If
           
           Set CurChild = CurChild.Next
        
           Do Until CurChild Is Nothing
              If Len(CurChild.Tag) = 0 Then
                 sAll = sAll & ItemDelimiter & CurChild.Text & IconDelimiter & CurChild.Image
              Else
                 sAll = sAll & ItemDelimiter & CurChild.Text & TagDelimiter & CurChild.Tag & IconDelimiter & CurChild.Image
              End If
              If Not CurChild.Child Is Nothing Then
                 TreeToAll_AddChildren sAll, CurChild
              End If
              Set CurChild = CurChild.Next
           Loop
           sAll = sAll & EndChildDelimiter
        End If
    Set CurChild = Nothing
End Sub

Public Sub FillTreeNode(tvwToFill As Object, nodCur As Object, sImage As String, bExpanded As Boolean)
    Dim nodX As Node
    Dim PrevNode As Node
    Dim tvwX As TreeView
    Dim sKey As String
    Dim sTag As String
    Dim sIcon As String
    Dim nChildDelimiter As Integer
    Dim nEndChildDelimiter As Integer

    Dim NodeStack() As Node
    ReDim NodeStack(0 To 0) As Node

    nChildDelimiter = Len(ChildDelimiter)
    nEndChildDelimiter = Len(EndChildDelimiter)

    Set tvwX = tvwToFill
        Set nodX = nodCur
            For Each CurItem In mCol
                sKey = sGetToken(CurItem.Key, 1, IconDelimiter)
                sIcon = sGetToken(CurItem.Key, 2, IconDelimiter)
                sTag = sGetToken(sKey, 2, TagDelimiter)
                sKey = sGetToken(sKey, 1, TagDelimiter)
                If Len(sIcon) = 0 Then sIcon = sImage

                If nChildDelimiter > 0 Then
                   If Left$(sKey, nChildDelimiter) = ChildDelimiter Then
                      sKey = Mid$(sKey, nChildDelimiter + 1)
                      If Not PrevNode Is Nothing Then
                         If UBound(NodeStack) > 0 Then
                            ReDim Preserve NodeStack(1 To UBound(NodeStack) + 1) As Node
                         Else
                           ReDim NodeStack(1 To 1) As Node
                         End If
                         Set NodeStack(UBound(NodeStack)) = nodX
                         Set nodX = PrevNode
                      End If
                   End If
                End If
On Error Resume Next
                If nodX Is Nothing Then
                   Set PrevNode = tvwX.Nodes.Add(, , sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter))
                Else
                   Set PrevNode = tvwX.Nodes.Add(nodX.Key, tvwChild, nodX.Key & "_" & sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter))
                End If
                With PrevNode
                     .ExpandedImage = sGetToken(sIcon, 1, EndChildDelimiter)
                     .Expanded = bExpanded
                     .Tag = sTag
                End With
                
                If nEndChildDelimiter > 0 Then
                   Do While Right$(sIcon, nEndChildDelimiter) = EndChildDelimiter
                      sIcon = Left$(sIcon, Len(sIcon) - nEndChildDelimiter)
                      Set nodX = NodeStack(UBound(NodeStack))
                      If UBound(NodeStack) > 1 Then
                         ReDim Preserve NodeStack(1 To UBound(NodeStack) - 1) As Node
                      Else
                         ReDim NodeStack(0 To 0) As Node
                      End If
                   Loop
                End If
            Next CurItem
        Set CurItem = Nothing
        Set PrevNode = Nothing
        Set nodX = Nothing
    Set tvwX = Nothing
End Sub


