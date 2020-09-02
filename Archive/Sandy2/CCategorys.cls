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
On Error Resume Next
    Dim sCategory As String
    Dim sTemplate As String
    
    If InStr(sLongTemplateName, " - ") = 0 Then Exit Property
    
    sCategory = sGetToken(sLongTemplateName, 1, " - ")
    sTemplate = sAfter(sLongTemplateName, 1, " - ")
    
    If Len(sCategory) = 0 Or Len(sTemplate) = 0 Then Exit Property
    
    Set ItemByLongTemplateName = Item(sCategory).Templates(sTemplate)
End Property

Public Function Find(sToFind As String, Optional ByVal sLocationToStart As String, Optional ByVal lAreaToStartSearchAt As Long, Optional ByVal lByteToStartAt As Long) As CCategory
    Dim CurItem As CCategory
    Dim sFoundAt As String

    For Each CurItem In mCol
        If CurItem.Find(sToFind, sLocationToStart, lAreaToStartSearchAt, lByteToStartAt) Then
           Set Find = CurItem
           Set CurItem = Nothing
           Exit Function
        End If
        lAreaToStartSearchAt = 0
        lByteToStartAt = 0
    Next CurItem
End Function

Public Property Get AllKeys(Optional ByVal sDelimiter As String = ";", Optional ByVal CategoryType As Long = -1) As String
    Dim CurItem As CCategory
    Dim sAllKeys As String

    If CategoryType > -1 Then
        For Each CurItem In mCol
            If CurItem.CategoryType = CategoryType Then
               sAllKeys = sAllKeys & CurItem.Key & sDelimiter
            End If
        Next CurItem
    Else
        For Each CurItem In mCol
            sAllKeys = sAllKeys & CurItem.Key & sDelimiter
        Next CurItem
    End If

    AllKeys = sAllKeys
End Property

Public Function Choose(Optional ByVal CategoryType As Long = -1) As String
    Choose = sChoose(AllKeys(, CategoryType))
End Function

Public Sub FillList(lstCtrl As Object, Optional ByVal CategoryType As Long = -1)
    Dim CurItem As CCategory
    lstCtrl.Clear
    If CategoryType > -1 Then
       For Each CurItem In mCol
           If CurItem.CategoryType = CategoryType Then
              lstCtrl.AddItem CurItem.Key
           End If
       Next CurItem
    Else
       For Each CurItem In mCol
           lstCtrl.AddItem CurItem.Key
       Next CurItem
    End If
End Sub

Public Function Replace(ByVal sToFind As String, ByVal sToReplace As String) As Long
    Dim lTotalCount As Long
    Dim CurrCategory As CCategory
    
    For Each CurrCategory In mCol
        lTotalCount = lTotalCount + CurrCategory.Replace(sToFind, sToReplace)
    Next CurrCategory

    Replace = lTotalCount
End Function

Public Sub SaveAll(db As Database)
On Error GoTo Categorys_SaveModifiedErr
    Dim rs As Recordset
    Dim objCur As CCategory
    Dim NextID As Long

Categorys_SaveModified_Start_Over:
    For Each objCur In mCol
        With objCur
             If .Deleted Then
                If .ID > 0 Then
                   Set rs = db.OpenRecordset("SELECT * FROM Category WHERE CategoryID=" & .ID, dbOpenDynaset)
                   rs.Delete
                   Remove .Key
                   GoTo Categorys_SaveModified_Start_Over
                End If
             ElseIf .Modified Then
                If .ID > 0 Then ' Modify the existing record
                   Set rs = db.OpenRecordset("SELECT * FROM Category WHERE CategoryID=" & .ID, dbOpenDynaset)
                   rs.Edit
                   rs!DateModified = Now()
                   .DateModified = rs!DateModified
                Else            ' Create a new record
                  'Set rs = db.OpenRecordset("SELECT MAX(CategoryID)+1 FROM Category")
                  'NextID = IIf(IsNull(rs.Fields(0).Value), 1, rs.Fields(0).Value)
                  'rs.Close
                   Set rs = db.OpenRecordset("SELECT * FROM Category", dbOpenDynaset)
                   rs.AddNew
                   rs!DateCreated = Now()
                   .DateCreated = rs!DateCreated
                   rs!DateModified = .DateCreated
                   .DateModified = .DateCreated
                   .ID = rs!CategoryID
                  'rs!CategoryID = NextID
                End If
                rs!CategoryName = Left$(.Key, 254)
                .SaveProperties rs
                rs.Update
                rs.Close
                .Modified = False
             End If
             SaveChildren db, objCur
        End With
    Next objCur
Exit Sub

Categorys_SaveModifiedErr:
    LogError "CCategorys", "SaveModified", Err.Number, Err.Description
    
    Resume Next
    Resume
End Sub

Public Sub GetAll(db As Database)
On Error GoTo Categorys_GetAllErr
    Dim rs As Recordset
    Dim objNew As CCategory
    Set rs = db.OpenRecordset("SELECT * FROM Category ORDER BY CategoryType, CategoryName", dbOpenForwardOnly, dbReadOnly)
    Clear
    With rs
         Do Until .EOF
            Set objNew = Add(!CategoryName, !CategoryID)
            With objNew
                 .LoadProperties rs
                 .Modified = False
                 'LoadChildren db, objNew
            End With
            .MoveNext
         Loop
         .Close
    End With
Exit Sub

Categorys_GetAllErr:
    LogError "CCategorys", "GetAll", Err.Number, Err.Description
End Sub

Public Sub Clear()
On Error GoTo Categorys_ClearErr
    Dim CurrCategory As CCategory
    For Each CurrCategory In mCol
        Set CurrCategory.MyParent = Nothing
    Next CurrCategory

    If mCol.Count > 0 Then
       Set mCol = Nothing
       Set mCol = New Collection
       ' mCol.CompareMode = TextCompare
    End If
Exit Sub

Categorys_ClearErr:
    LogError "CCategorys", "Clear", Err.Number, Err.Description
End Sub

Public Function Add(ByVal Key As String, Optional ByVal ID As Long = -1, Optional ByVal sCategoryToCopy As String, Optional ByVal bCopyInformation As Boolean = True) As CCategory
  ' Create a new Category
    Dim objNewMember As CCategory
    Set objNewMember = CreateObject("SandySupport.CCategory")

  ' Set the properties passed into the method
    With objNewMember
         If ID = -1 Then ID = NextNegativeUnique
         .ID = ID
         .Key = Key
         .Modified = True
         Set .MyParent = Me
    End With

    If Len(sCategoryToCopy) = 0 Then
       mCol.Add objNewMember, UCase$(Key)
    Else
       If mCol(sCategoryToCopy) Is Nothing Then
          MsgBox "Error during duplication of a category."
       Else
          If StrComp(Key, sCategoryToCopy, vbTextCompare) <> 0 Then
             Item(sCategoryToCopy).Duplicate objNewMember, bCopyInformation
          End If
          mCol.Add objNewMember, UCase$(Key)
       End If
    End If

  ' Return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As CCategory
Attribute Item.VB_UserMemId = 0
On Error Resume Next
       Static CurItem As CCategory

       Err.Clear
       If VarType(vntIndexKey) = vbString Then
          Set Item = mCol(UCase$(vntIndexKey))
       Else
          Set Item = mCol(vntIndexKey)
       End If
       
       If Err.Number <> 0 Then
          Err.Clear
         For Each CurItem In mCol
              If CurItem.Key = vntIndexKey Or CurItem.ID = Val(vntIndexKey) Then
                 Set Item = CurItem
                 Set CurItem = Nothing
                 Exit Property
              End If
          Next CurItem
       End If
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
On Error Resume Next
    If VarType(vntIndexKey) = vbString Then
       Set mCol(UCase$(vntIndexKey)).MyParent = Nothing
       mCol.Remove UCase$(vntIndexKey)
    Else
       Set mCol(vntIndexKey).MyParent = Nothing
       mCol.Remove vntIndexKey
    End If
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
        ' mCol.CompareMode = TextCompare
    ' LogEvent "CCategorys: Initialize"
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
    ' LogEvent "CCategorys: Terminate"
End Sub

Public Property Get MyType() As String
    MyType = "modGeneral"
End Property

Public Sub LoadChildren(db As Database, objSetToLoad As CCategory)
    With objSetToLoad
         .Templates.GetAll db, .ID, .Key
    End With
End Sub

Public Sub SaveChildren(db As Database, objSetToSave As CCategory)
    With objSetToSave
         .Templates.SaveAll db, .ID
    End With
End Sub

