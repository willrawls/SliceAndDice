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
    Dim objNewMember As CAssocArray

    If Len(sKey) = 0 Then
       Err.Raise vbObjectError + 0, "CAssocArrays", "Tryed to add an Assoc Array without a key."
    End If

    Set objNewMember = CreateObject("SandySupport.CAssocArray")
        objNewMember.Section = sKey
        mCol.Add objNewMember, sKey
        Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Sub Clear(Optional ByVal bRecreateCollection As Boolean = True)
On Error Resume Next
    Set mCol = Nothing
    If bRecreateCollection Then
       Set mCol = New Collection
       ' mCol.CompareMode = TextCompare
    End If
End Sub


Public Property Get Item(sIndexKey As String) As CAssocArray
Attribute Item.VB_UserMemId = 0
On Error GoTo EH_CAssocArrays_Item
       Set Item = mCol(sIndexKey)

EH_CAssocArrays_Item_Continue:
       Exit Property

EH_CAssocArrays_Item:
       Set Item = Add(sIndexKey)
       Err.Clear
       Resume EH_CAssocArrays_Item_Continue
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Load(sFilename As String, Optional bClearFirst As Boolean = True)
    If Len(sFilename) = 0 Then
       Err.Raise vbObjectError + 2, "CAssocArrays_LoadAll", "Tryed to load w/o a filename."
    End If

    Dim fh As Integer
    Dim sLine As String
    Dim CurAssocArray As CAssocArray

    fh = FreeFile

    If bClearFirst = True Then
       Clear
    End If
    
    Open sFilename For Input Access Read As #fh
         Do Until EOF(fh)
            Line Input #fh, sLine
            If Len(sLine) = 0 Then
             ' Skip it
            ElseIf Left$(sLine, 1) = "[" Then
               Set CurAssocArray = Add(Mid$(sLine, 2, Len(sLine) - 2))
               With CurAssocArray
                    Line Input #fh, sLine
                        .ItemDelimiter = sAfter(sDenormalize(sLine), 1, "=")
                    Line Input #fh, sLine
                        .KeyValueDelimiter = sAfter(sDenormalize(sLine), 1, "=")
                    Line Input #fh, sLine
                        .FieldDelimiter = sAfter(sDenormalize(sLine), 1, "=")
               End With
            Else
               CurAssocArray.Add sGetToken(sLine, 1, "="), sAfter(sLine, 1, "=")
            End If
         Loop
    Close #fh
    Set CurAssocArray = Nothing
End Sub

Public Sub Remove(sIndexKey As String)
    mCol.Remove sIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Sub Save(sFilename As String)
    If Len(sFilename) = 0 Then
       Err.Raise vbObjectError + 1, "CAssocArrays_SaveAll", "Tryed to save w/o a filename."
    End If

    Dim fh As Integer
    Dim CurAssocArray As CAssocArray
    Dim sOldItem As String
    Dim sOldKV As String

    fh = FreeFile
    Open sFilename For Output Access Write As #fh
         For Each CurAssocArray In mCol
             With CurAssocArray
                  sOldItem = .ItemDelimiter
                  sOldKV = .KeyValueDelimiter
                  Print #fh, "[" & .Section & "]"
                  Print #fh, "Delimiter Item=" & sNormalize(.ItemDelimiter)
                  Print #fh, "Delimiter Key Value=" & sNormalize(.KeyValueDelimiter)
                  Print #fh, "Delimiter Field=" & sNormalize(.FieldDelimiter)
                  .ItemDelimiter = Chr$(13) & Chr$(10)
                  .KeyValueDelimiter = "="
                  Print #fh, .All
                  .ItemDelimiter = sOldItem
                  .KeyValueDelimiter = sOldKV
             End With
         Next CurAssocArray
    Close #fh
    Set CurAssocArray = Nothing
End Sub

Private Sub Class_Initialize()
    Clear
    ' LogEvent "CAssocArrays: Initialize"
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
    ' LogEvent "CAssocArrays: Terminate"
End Sub

