VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSliceAndDice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy - MDB Object Hierarchy"
Option Explicit

Public Categorys As CCategorys
Attribute Categorys.VB_VarUserMemId = 0
Public LastFound As CCategory

'Public SystemInfo As CAssocArrays

Public ConnectionOpen As Boolean
Public db As Database

Private m_LastTemplate As CTemplate
Private m_LastCategory As CCategory

Public Function ApplyPatch(sFilename As String) As Boolean
    Dim sFile As String
    Dim sLine As String
    
    If FileExists(sFilename) Then
       sFile = sFileContents(sFilename)
       If bUserSure("This file contains " & lTokenCount(sFile, "!@#$%^&*())(*&^%$#@! Template ") - 1 & " template changes and/or additions." & vbCr & "You will be if one of your templates has a later modified date than an included delta template." & vbCr & vbTab & "Do you wish to continue ?") Then
          Do Until Len(sFile) = 0
             sLine = sGetToken(sFile, 1, vbNewLine)
             sFile = sAfter(sFile, 1, vbNewLine)
             If InStr(sLine, "!@#$%^&*())(*&^%$#@!") Then
                sLine = Trim$(Mid$(sLine, 21))
                Select Case UCase$(sGetToken(sLine))
                       Case "TEMPLATE"
                       Case "ALPHADATE"
                       Case "DELTADATE"
                       Case "CODELOCKED"
                       Case "UNDELETABLE"
                         
                       Case "ATCURSOR"
                       Case "ATTOP"
                       Case "ATBOTTOM"
                       Case "TOFILE"
                         
                       Case Else
               End Select
             Else
             End If
          Loop
       End If
    End If
End Function

Public Function FindFirstTemplateContaining(ByVal sToFind As String) As CTemplate
    Set LastFound = Categorys.Find(sToFind)
    Set FindFirstTemplateContaining = LastFound
End Function

Public Function FindNextTemplate(ByVal sToFind As String) As CTemplate
    
End Function



Public Function GenerateDeltaPatchFile(ByVal DeltaDate As Date, sPatchFilename As String) As Boolean
    Dim CurrCategory As SandySupport.CCategory
    Dim CurrTemplate As SandySupport.CTemplate
    Dim fh As Long

On Error Resume Next
    If Len(Dir$(sPatchFilename)) > 0 Then
       Err.Clear
       Kill sPatchFilename
       If Err.Number <> 0 Then
          Err.Clear
          GoTo SkipMakePathfile
       End If
    End If
    fh = FreeFile
    Open sPatchFilename For Output Access Write As #fh
         Print #fh, "!@#$%^&*())(*&^%$#@! PatchFileVersion " & App.Major & "." & App.Minor & "." & App.Revision
         For Each CurrCategory In Categorys
             For Each CurrTemplate In CurrCategory.Templates
                 If CurrTemplate.DateModified >= DeltaDate Then
                    With CurrTemplate
                         Print #fh, "!@#$%^&*())(*&^%$#@! Template " & .Key
                         Print #fh, "!@#$%^&*())(*&^%$#@! Revision " & .Revision
                         Print #fh, "!@#$%^&*())(*&^%$#@! AlphaDate " & .DateCreated
                         Print #fh, "!@#$%^&*())(*&^%$#@! DeltaDate " & .DateModified
                         Print #fh, "!@#$%^&*())(*&^%$#@! CodeLocked " & IIf(.Locked, "True", "False")
                         Print #fh, "!@#$%^&*())(*&^%$#@! Undeletable " & IIf(.Undeletable, "True", "False")
                         If Len(.memoCodeAtCursor) Then
                            Print #fh, "!@#$%^&*())(*&^%$#@! AtCursor"
                            Print #fh, .memoCodeAtCursor
                         End If
                         If Len(.memoCodeAtTop) Then
                            Print #fh, "!@#$%^&*())(*&^%$#@! AtTop"
                            Print #fh, .memoCodeAtTop
                         End If
                         If Len(.memoCodeAtBottom) Then
                            Print #fh, "!@#$%^&*())(*&^%$#@! AtBottom"
                            Print #fh, .memoCodeAtBottom
                         End If
                         If Len(.memoCodeToFile) Then
                            Print #fh, "!@#$%^&*())(*&^%$#@! ToFile " & .Filename
                            Print #fh, .memoCodeToFile
                         End If
                         If Len(.memoAttributes) Then
                            Print #fh, "!@#$%^&*())(*&^%$#@! Attributes"
                            Print #fh, sadEncrypt(.memoAttributes)
                         End If
                    End With
                 End If
             Next CurrTemplate
         Next CurrCategory
    Close #fh
    GenerateDeltaPatchFile = True
SkipMakePathfile:
End Function

Public Function Load(Optional ByVal DatabaseName As String = "SliceAndDice", Optional ByVal bKeepConnectionOpen As Boolean = True) As Boolean
On Error GoTo EH_CSliceAndDice_Load
    Dim bCategorysFound As Boolean
    Dim bTemplatesFound As Boolean
    Dim bSystemFound As Boolean
    Dim CurTable As TableDef
    Dim rst As Recordset

    Dim tblTemplates     As TableDef
    Dim fldTemplates     As Field
    Dim ndxTemplates     As Index

   'If Not ConnectionOpen Then
   '  rdoEngine.rdoEnvironments(0).CursorDriver = rdUseNone
   'End If

    If ConnectionOpen Then
       db.Close
       ConnectionOpen = False
    End If

    Set db = OpenDatabase(DatabaseName, False, False)
    
  ' Insure this is a Slice and Dice Database
    bCategorysFound = False
    bTemplatesFound = False
On Error Resume Next
    For Each CurTable In db.TableDefs
        If CurTable.Name = "Category" Then
           bCategorysFound = True
           If CurTable.Fields("memoAttributes") Is Nothing Then                             ' Schema change circa 2/18/1999
              CurTable.Fields.Append CurTable.CreateField("memoAttributes", dbMemo)
           End If
        ElseIf CurTable.Name = "Template" Then
           bTemplatesFound = True
           If CurTable.Fields("memoAttributes") Is Nothing Then                             ' Schema change circa 2/18/1999
              CurTable.Fields.Append CurTable.CreateField("memoAttributes", dbMemo)
           End If
           If CurTable.Fields("Favorite") Is Nothing Then                                   ' Schema change circa 2/18/1999
              CurTable.Fields.Append CurTable.CreateField("Favorite", dbBoolean)
           End If
           If CurTable.Fields("RevisionCount") Is Nothing Then                                   ' Schema change circa 3/27/1999
              CurTable.Fields.Append CurTable.CreateField("RevisionCount", dbLong)
           End If
           If CurTable.Fields("TimerInsertion") Is Nothing Then                                   ' Schema change circa 3/27/1999
              CurTable.Fields.Append CurTable.CreateField("TimerInsertion", dbText, 255)
           End If
        ElseIf CurTable.Name = "SystemInfo" Then
           bSystemFound = True
        End If
    Next CurTable

    If Not bSystemFound Then
        Set tblTemplates = db.CreateTableDef("SystemInfo")
        With tblTemplates
             Set fldTemplates = .CreateField("SystemInfoID", dbLong)
             fldTemplates.Attributes = dbAutoIncrField
             .Fields.Append fldTemplates
            .Fields.Append .CreateField("SystemInfoName", dbText, 255)
            .Fields.Append .CreateField("DateCreated", dbDate)
            .Fields.Append .CreateField("DateModified", dbDate)
            .Fields.Append .CreateField("memoAttributes", dbMemo)

            Set ndxTemplates = .CreateIndex("PrimaryKey")
            With ndxTemplates
                 .Fields.Append .CreateField("SystemInfoID")
                 .Primary = True
                 .Unique = True
                 .Required = True
            End With
            .Indexes.Append ndxTemplates

            Set ndxTemplates = .CreateIndex("SystemInfoName")
            With ndxTemplates
                 .Fields.Append .CreateField("SystemInfoName")
                 .Primary = False
                 .Unique = True
                 .Required = True
            End With
            .Indexes.Append ndxTemplates
            
            Set ndxTemplates = Nothing
            
            db.TableDefs.Append tblTemplates
        End With
    End If

On Error GoTo EH_CSliceAndDice_Load
    If Not bCategorysFound Or Not bTemplatesFound Then
       MsgBox "'" & DatabaseName & "' is not a valid Slice and Dice database." & gs2EOLTab & "Please select a valid Slice and Dice database.", vbCritical
       db.Close
       ConnectionOpen = False
       Load = False
       Exit Function
    End If
    
    Categorys.GetAll db
    
'    If Not SystemInfo Is Nothing Then
'       SystemInfo.Clear False
'       Set SystemInfo = Nothing
'    End If
'    Set SystemInfo = CreateObject("SandySupport.CAssocArrays")
'    Set rst = db.OpenRecordset("SELECT * FROM SystemInfo", dbOpenForwardOnly, dbReadOnly)
'    With rst
'         Do Until .EOF
'            SystemInfo(!SystemInfoName).All = nZ(!memoAttributes)
'            .MoveNext
'         Loop
'         .Close
'    End With
'    Set rst = Nothing
    
    If Not bKeepConnectionOpen Then db.Close
    ConnectionOpen = bKeepConnectionOpen
    Load = True

EH_CSliceAndDice_Load_Continue:
    Exit Function

EH_CSliceAndDice_Load:
    LogError "CSliceAndDice", "Load", Err.Number, Err.Description
    Load = False
    Resume EH_CSliceAndDice_Load_Continue
    
    Resume
End Function

Public Function ReplaceInCategory(ByVal sCategory As String, ByVal sToFind As String, ByVal sToReplace As String) As Long
  ' Returns the number of replacements that occured
    If Categorys(sCategory) Is Nothing Then
       MsgBox "The category '" & sCategory & "' does not exist. Aborting category find and replace."
       Exit Function
    End If
    
    ReplaceInCategory = Categorys(sCategory).Replace(sToFind, sToReplace)
End Function

Public Function Save(Optional ByVal DatabaseName As String = "SliceAndDice", Optional ByVal bKeepConnectionOpen As Boolean = True) As Boolean
    Dim rst As Recordset
    Dim asaCurr As CAssocArray
    Dim AtLine As String
    Dim CurrItem As CAssocItem

On Error GoTo EH_CSliceAndDice_Save

    If Categorys Is Nothing Then Exit Function ' Yo. Something's wrong. Don't save shit.
    
    If Not ConnectionOpen Then
       Set db = OpenDatabase(DatabaseName, False, False)
    End If

    If Not Categorys Is Nothing Then
       Categorys.SaveAll db
    End If

'    If Not SystemInfo Is Nothing Then
'       Set rst = db.OpenRecordset("SystemInfo", dbOpenTable)
'       With rst
'            Do Until .EOF
'               .Edit
'                   !memoAttributes = zn(SystemInfo(!SystemInfoName).All)
'                   !DateModified = Now
'                   SystemInfo.Remove !SystemInfoName
'               .Update
'               .MoveNext
'            Loop
'            For Each asaCurr In SystemInfo
'                asaCurr.Remove "EOL"
'                asaCurr.Remove "EOLTAB"
'                asaCurr.Remove "EOL2"
'                asaCurr.Remove "EOL2Tab"
'Save_StartOver:
'                For Each CurrItem In asaCurr
'                    If CurrItem.Value = "0,8" Or CurrItem.Value = "0,0" Or CurrItem.Value = vbNullString Then
'                       asaCurr.Remove CurrItem.Key
'                       Set CurrItem = Nothing
'                       GoTo Save_StartOver
'                    End If
'                Next CurrItem
'                .AddNew
'                   !SystemInfoName = asaCurr.Section
'                   !DateCreated = Now
'                   !DateModified = !DateCreated
'                   !memoAttributes = zn(asaCurr.All)
'                .Update
'            Next asaCurr
'            .Close
'       End With
'       Set rst = Nothing
'    End If
    
    If Not bKeepConnectionOpen Then db.Close
    ConnectionOpen = bKeepConnectionOpen

    Save = True

EH_CSliceAndDice_Save_Continue:
    Exit Function

EH_CSliceAndDice_Save:
    LogError "CSliceAndDice", "Save(" & AtLine & ")", Err.Number, Err.Description
    Save = False
    Resume EH_CSliceAndDice_Save_Continue
End Function


Public Function sChoosePatch(Method As Integer) As String
    Dim CurrCategory As SandySupport.CCategory
    Dim CurrTemplate As SandySupport.CTemplate
    Dim asaDates As SandySupport.CAssocArray

    Set asaDates = CreateObject("SandySupport.CAssocArray")

    asaDates.Clear
    asaDates.AddInOrder = True
       For Each CurrCategory In Categorys
           For Each CurrTemplate In CurrCategory.Templates
               If Method = 0 Then
                  asaDates(vbNullString & Format(CLng(CurrTemplate.DateModified), "00000")) = Format(CurrTemplate.DateModified, "Mmmm D, YYYY H AM/PM")
               Else
                  
               End If
           Next CurrTemplate
       Next CurrCategory
       asaDates.KeyValueDelimiter = "-"
       asaDates.ItemDelimiter = ";"
       sChoosePatch = sChoose(asaDates.Column(0))
End Function

Private Sub Class_Initialize()
    Set Categorys = CreateObject("SandySupport.CCategorys")
    Set Categorys.MyParent = Me
    ' LogEvent "CSliceAndDice: Initialize"
End Sub

Private Sub Class_Terminate()
    Save , False
    Set LastFound = Nothing
    Set Categorys.MyParent = Nothing
    Set Categorys = Nothing
    ' LogEvent "CSliceAndDice: Terminate"
End Sub

