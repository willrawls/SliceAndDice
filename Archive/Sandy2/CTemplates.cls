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
    Dim CurItem As CTemplate
    Dim sFoundAt As String

    If Len(sLocationToStart) = 0 Then
       For Each CurItem In mCol
           If CurItem.Find(sToFind, lAreaToStartSearchAt, lByteToStartAt) Then
              Set Find = CurItem
              Set CurItem = Nothing
              Exit Function
           End If
           lAreaToStartSearchAt = 0
           lByteToStartAt = 0
       Next CurItem
    Else
       For Each CurItem In mCol
           If CurItem.Find(sToFind, lAreaToStartSearchAt, lByteToStartAt) Then
              Set Find = CurItem
              Set CurItem = Nothing
              Exit Function
           End If
           lAreaToStartSearchAt = 0
           lByteToStartAt = 0
       Next CurItem
    End If
End Function
Public Function Duplicate(Optional ByVal bCopyInformation As Boolean = True) As CTemplates
    MsgBox "CTemplates.Duplicate : Function disabled"
'    Dim Dupe As CTemplates
'    Dim CurItem As CTemplate
'
'    Set Dupe = CreateObject("SandySupport.CTemplates")
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
On Error GoTo Templates_SaveModifiedErr
    Dim rs As Recordset
    Dim objCur As CTemplate
    Dim NextID As Long

Templates_SaveModified_Start_Over:
    For Each objCur In mCol
        With objCur
             If .Deleted Then
                If .ID > 0 Then
                   Set rs = db.OpenRecordset("SELECT * FROM Template WHERE TemplateID=" & .ID, dbOpenDynaset)
                   rs.Delete
                   Remove .Key
                   GoTo Templates_SaveModified_Start_Over
                End If
             ElseIf .Modified Then
                If Len(.Key) = 0 Or Len(.ShortTemplateName) = 0 Then
                   MsgBox "Probable error saving template:" & vbCr & vbTab & "ParentKey=" & .ParentKey & vbCr & vbTab & "ShortTemplateName=" & .ShortTemplateName & vbCr & vbTab & "OriginalShortName=" & .OriginalShortName & vbCr & "This template will not be saved."
                Else
                   If .ID > 0 Then ' Modify the existing record
                      Set rs = db.OpenRecordset("SELECT * FROM Template WHERE TemplateID=" & .ID, dbOpenDynaset)
                      rs.Edit
                      rs!DateModified = Now()
                      .DateModified = rs!DateModified
                   Else            ' Create a new record
                     'Set rs = db.OpenRecordset("SELECT MAX(TemplateID)+1 FROM Template")
                     'NextID = IIf(IsNull(rs.Fields(0).Value), 1, rs.Fields(0).Value)
                     'rs.Close
                      Set rs = db.OpenRecordset("SELECT * FROM Template", dbOpenDynaset)
                      rs.AddNew
                      rs!DateCreated = Now()
                      .DateCreated = rs!DateCreated
                      rs!DateModified = .DateCreated
                      .DateModified = .DateCreated
                      .ID = rs!TemplateID
                     'rs!TemplateID = NextID
                      .Parent = CategoryID
                      rs!CategoryID = CategoryID
                   End If
                   rs!TemplateName = Left$(.Key, 254)
                   .SaveProperties rs
                   rs.Update
                   rs.Close
                End If
                .Modified = False
             End If
        End With
    Next objCur
Exit Sub

Templates_SaveModifiedErr:
    LogError "CTemplates", "SaveModified", Err.Number, Err.Description
    Resume Next

    Resume
End Sub

Public Sub GetAll(db As Database, CategoryID As Long, CategoryName As String)
On Error GoTo Templates_GetAllErr
    Dim lHotKey As Integer
    Dim lHotMod As Integer
    Dim sHotKey As String

    Dim rs As Recordset
    Dim objNew As CTemplate

    Set rs = db.OpenRecordset("SELECT * FROM Template WHERE CategoryID=" & CategoryID & " ORDER BY ShortTemplateName", dbOpenForwardOnly, dbReadOnly)
    
    Clear
    With rs
         Do Until .EOF
            Set objNew = Add(!TemplateName, !TemplateID)
            With objNew
                 .LoadProperties rs
                 .Modified = False
            End With
            .MoveNext
         Loop
         .Close
    End With
    
Templates_GetAllErr_Continue:
    Exit Sub

Templates_GetAllErr:
    LogError "CTemplates", "GetAll", Err.Number, Err.Description
    Resume Templates_GetAllErr_Continue
    
    Resume
End Sub

Public Sub Clear()
On Error GoTo Templates_ClearErr
    Dim CurrItem As CTemplate
    For Each CurrItem In mCol
        Set CurrItem.MyParent = Nothing
    Next CurrItem

    If mCol.Count > 0 Then
       Set mCol = Nothing
       Set mCol = New Collection
       ' mCol.CompareMode = TextCompare
    End If
Exit Sub

Templates_ClearErr:
    LogError "CTemplates", "Clear", Err.Number, Err.Description
End Sub

Public Function Add(ByVal Key As String, Optional ByVal ID As Long = -1, Optional ByVal sKey As String) As CTemplate
  ' Create a new Template
    Dim objNewMember As CTemplate
    Set objNewMember = CreateObject("SandySupport.CTemplate")

  ' Set the properties passed into the method
    With objNewMember
         If ID = -1 Then ID = NextNegativeUnique
         .ID = ID
         .Key = Key
         .ShortTemplateName = sAfter(Key, 1, " - ")
         .OriginalShortName = .ShortTemplateName
         .ParentKey = sGetToken(Key, 1, " - ")
         .Modified = True
         Set .MyParent = Me
    End With

    If Len(sKey) = 0 Then
       mCol.Add objNewMember, UCase$(Key)
    Else
       mCol.Add objNewMember, UCase$(sKey)
    End If

  ' Return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get AllKeys(Optional ByVal sDelimiter As String = ";") As String
    Dim CurItem As CTemplate
    Dim sAllKeys As String

    For Each CurItem In mCol
        With CurItem
             sAllKeys = sAllKeys & .Key & sDelimiter
        End With
    Next CurItem

    AllKeys = sAllKeys
End Property

Public Function Choose() As String
    Choose = sChoose(AllKeys)
End Function

Public Property Get Item(vntIndexKey As Variant) As CTemplate
Attribute Item.VB_UserMemId = 0
On Error Resume Next
    Static CurItem As CTemplate

    If mbNeverAccessed Then
       GetAll MyParent.MyParent.MyParent.db, MyParent.ID, MyParent.Key
       mbNeverAccessed = False
    End If

    Err.Clear
    If VarType(vntIndexKey) = vbString Then
       Set Item = mCol(UCase$(vntIndexKey))
    Else
       Set Item = mCol(vntIndexKey)
    End If
    
    If Err.Number <> 0 Then
       Err.Clear
       For Each CurItem In mCol
           If CurItem.Key = UCase$(vntIndexKey) Or CurItem.ID = Val(vntIndexKey) Or StrComp(CurItem.ShortTemplateName, vntIndexKey, 1) = 0 Or StrComp(CurItem.ParentKey & " - " & CurItem.OriginalShortName, vntIndexKey, 1) = 0 Then
              Set Item = CurItem
              Set CurItem = Nothing
              Exit Property
           End If
       Next CurItem
    End If
End Property

Public Property Get Count() As Long
    If mbNeverAccessed Then
       GetAll MyParent.MyParent.MyParent.db, MyParent.ID, MyParent.Key
       mbNeverAccessed = False
    End If
    
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    If mbNeverAccessed Then
       GetAll MyParent.MyParent.MyParent.db, MyParent.ID, MyParent.Key
       mbNeverAccessed = False
    End If
    
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
    ' mCol.CompareMode = TextCompare
    mbNeverAccessed = True
    ' LogEvent "CTemplates: Initialize"
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
    ' LogEvent "CTemplates: Terminate"
End Sub

Public Property Get MyType() As String
    MyType = "CCategory"
End Property

