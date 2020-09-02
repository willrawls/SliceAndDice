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
Public db           As Database

Private m_LastTemplate As CTemplate
Private m_LastCategory As CCategory

Public Function ApplyPatch(sFilename As String) As Boolean
'1        Dim sFile As String
'2        Dim sLine As String
'
'3        If FileExists(sFilename) Then
'4            sFile = sFileContents(sFilename)
'5            If bUserSure("This file contains " & lTokenCount(sFile, "!@#$%^&*())(*&^%$#@! Template ") - 1 & gsS & gsTemplate & " changes and/or additions." & vbNewLine & "You will be if one of your templates has a later modified date than an included delta " & gsTemplate & gsP & gsEolTab & "Do you wish to continue ?") Then
'6                Do Until Len(sFile) = 0
'7                    sLine = sGetToken(sFile, 1, vbNewLine)
'8                    sFile = sAfter(sFile, 1, vbNewLine)
'9                    If InStr(sLine, "!@#$%^&*())(*&^%$#@!") Then
'10                       sLine = Trim$(Mid$(sLine, 21))
'                    Select Case UCase$(sGetToken(sLine))
'                        Case "TEMPLATE"
'11                           Case "ALPHADATE"
'12                           Case "DELTADATE"
'13                           Case "CODELOCKED"
'14                           Case "UNDELETABLE"
'
'15                           Case "ATCURSOR"
'16                           Case "ATTOP"
'17                           Case "ATBOTTOM"
'18                           Case "TOFILE"
'
'19                           Case Else
'20                       End Select
'21                   Else
'22                   End If
'23               Loop
'24           End If
'25       End If
End Function

Public Function FindFirstTemplateContaining(ByVal sToFind As String) As CTemplate
26       Set LastFound = Categorys.Find(sToFind)
27       Set FindFirstTemplateContaining = LastFound
End Function

Public Function FindNextTemplate(ByVal sToFind As String) As CTemplate

End Function

Public Function GenerateDeltaPatchFile(ByVal DeltaDate As Date, sPatchFilename As String) As Boolean
'28       Dim CurrCategory   As SliceAndDice.CCategory
'29       Dim CurrTemplate   As SliceAndDice.CTemplate
'30       Dim fh             As Long
'
'31       On Error Resume Next
'32       If Len(Dir$(sPatchFilename)) > 0 Then
'33           Err.Clear
'34           Kill sPatchFilename
'35           If Err.Number <> 0 Then
'36               Err.Clear
'37               GoTo SkipMakePathfile
'38           End If
'39       End If
'40       fh = FreeFile
'41       Open sPatchFilename For Output Access Write As #fh
'
'         Print #fh, "!@#$%^&*())(*&^%$#@! PatchFileVersion " & App.Major & gsP & App.Minor & gsP & App.Revision
'         Print #fh, "!@#$%^&*())(*&^%$#@! UserID " & sFOfX(8)
'
'
'43       For Each CurrCategory In Categorys
'44           For Each CurrTemplate In CurrCategory.Templates
'45               If CurrTemplate.DateModified >= DeltaDate Then
'46                   With CurrTemplate
'47                       Print #fh, "!@#$%^&*())(*&^%$#@! Template " & .Key
'48                       Print #fh, "!@#$%^&*())(*&^%$#@! Revision " & .Revision
'49                       Print #fh, "!@#$%^&*())(*&^%$#@! AlphaDate " & .DateCreated
'50                       Print #fh, "!@#$%^&*())(*&^%$#@! DeltaDate " & .DateModified
'51                       Print #fh, "!@#$%^&*())(*&^%$#@! CodeLocked " & IIf(.Locked, "True", "False")
'52                       Print #fh, "!@#$%^&*())(*&^%$#@! Undeletable " & IIf(.Undeletable, "True", "False")
'53                       If Len(.memoCodeAtCursor) Then
'54                           Print #fh, "!@#$%^&*())(*&^%$#@! AtCursor"
'55                           Print #fh, .memoCodeAtCursor
'56                       End If
'57                       If Len(.memoCodeAtTop) Then
'58                           Print #fh, "!@#$%^&*())(*&^%$#@! AtTop"
'59                           Print #fh, .memoCodeAtTop
'60                       End If
'61                       If Len(.memoCodeAtBottom) Then
'62                           Print #fh, "!@#$%^&*())(*&^%$#@! AtBottom"
'63                           Print #fh, .memoCodeAtBottom
'64                       End If
'65                       If Len(.memoCodeToFile) Then
'66                           Print #fh, "!@#$%^&*())(*&^%$#@! ToFile " & .FileName
'67                           Print #fh, .memoCodeToFile
'68                       End If
'69                       If Len(.memoAttributes) Then
'70                           Print #fh, "!@#$%^&*())(*&^%$#@! Attributes"
'71                           Print #fh, .memoAttributes
'72                       End If
'73                   End With
'74               End If
'75           Next CurrTemplate
'76       Next CurrCategory
'77       Close #fh
'78       GenerateDeltaPatchFile = True
'79 SkipMakePathfile:
End Function

Public Function DumpToDirectory(ByVal sStartPath As String) As Boolean
On Error GoTo ErrorHandler
    Dim sCurrDir       As String
    Dim sCurrFile      As String
    Dim CurrCategory   As SliceAndDice.CCategory
    Dim CurrTemplate   As SliceAndDice.CTemplate
    Dim CurrBag        As PropertyBag
    Dim fh             As Long

    If Len(Dir$(sStartPath, vbDirectory)) = 0 Then
       MsgBox sStartPath & " does not exist or is not accessible."
       Exit Function
    End If
    
    If Right$(sStartPath, 1) <> "\" Then sCurrDir = sCurrDir & "\"

    For Each CurrCategory In Categorys
        sCurrDir = sStartPath & sMassage(CurrCategory.Key)

        If Len(Dir$(sCurrDir, vbDirectory)) = 0 Then
           MkDir sCurrDir
        End If

        For Each CurrTemplate In CurrCategory.Templates
            Set CurrBag = CurrTemplate

            sCurrFile = sCurrDir & "\" & sMassage(CurrTemplate.Key)

            fh = FreeFile
            SaveToFile sCurrFile, CurrBag.Contents
        Next CurrTemplate
    Next CurrCategory
    Close #fh
    DumpToDirectory = True

ErrorHandler_Continue:
    Exit Function

ErrorHandler:
    MsgBox "CSliceAndDice.DumpToDirectory ( " & Err.Number & " ) " & Err.Description
    Resume ErrorHandler_Continue
    
    Resume
End Function


Public Function Load(Optional ByVal DatabaseName As String = "SliceAndDice", Optional ByVal bKeepConnectionOpen As Boolean = True) As Boolean
80       On Error GoTo EH_CSliceAndDice_Load
81       Dim bCategorysFound As Boolean
82       Dim bTemplatesFound As Boolean
83       Dim bSystemFound As Boolean
84       Dim CurTable As TableDef
85       Dim rst As Recordset

86       Dim tblTemplates     As TableDef
87       Dim fldTemplates     As Field
88       Dim ndxTemplates     As Index

    'If Not ConnectionOpen Then
    '  rdoEngine.rdoEnvironments(0).CursorDriver = rdUseNone
    'End If

89       If ConnectionOpen Then
90           db.Close
91           ConnectionOpen = False
92       End If

93       Set db = OpenDatabase(DatabaseName, False, False)

    ' Insure this is a " & gsSliceAndDice & " Database
94       bCategorysFound = False
95       bTemplatesFound = False
96       On Error Resume Next
97       For Each CurTable In db.TableDefs
98           If CurTable.Name = gsCategory Then
99               bCategorysFound = True
100              If CurTable.Fields("memoAttributes") Is Nothing Then    ' Schema change circa 2/18/1999
101                  CurTable.Fields.Append CurTable.CreateField("memoAttributes", dbMemo)
102              End If
103          ElseIf CurTable.Name = "Template" Then
104              bTemplatesFound = True
105              If CurTable.Fields("memoAttributes") Is Nothing Then    ' Schema change circa 2/18/1999
106                  CurTable.Fields.Append CurTable.CreateField("memoAttributes", dbMemo)
107              End If
108              If CurTable.Fields("Favorite") Is Nothing Then    ' Schema change circa 2/18/1999
109                  CurTable.Fields.Append CurTable.CreateField("Favorite", dbBoolean)
110              End If
111              If CurTable.Fields("RevisionCount") Is Nothing Then    ' Schema change circa 3/27/1999
112                  CurTable.Fields.Append CurTable.CreateField("RevisionCount", dbLong)
113              End If
114              If CurTable.Fields("TimerInsertion") Is Nothing Then    ' Schema change circa 3/27/1999
115                  CurTable.Fields.Append CurTable.CreateField("TimerInsertion", dbText, 255)
116              End If
117          ElseIf CurTable.Name = "SystemInfo" Then
118              bSystemFound = True
119          End If
120      Next CurTable

121      If Not bSystemFound Then
122          Set tblTemplates = db.CreateTableDef("SystemInfo")
123          With tblTemplates
124              Set fldTemplates = .CreateField("SystemInfoID", dbLong)
125              fldTemplates.Attributes = dbAutoIncrField
126              .Fields.Append fldTemplates
127              .Fields.Append .CreateField("SystemInfoName", dbText, 255)
128              .Fields.Append .CreateField("DateCreated", dbDate)
129              .Fields.Append .CreateField("DateModified", dbDate)
130              .Fields.Append .CreateField("memoAttributes", dbMemo)

131              Set ndxTemplates = .CreateIndex("PrimaryKey")
132              With ndxTemplates
133                  .Fields.Append .CreateField("SystemInfoID")
134                  .Primary = True
135                  .Unique = True
136                  .Required = True
137              End With
138              .Indexes.Append ndxTemplates

139              Set ndxTemplates = .CreateIndex("SystemInfoName")
140              With ndxTemplates
141                  .Fields.Append .CreateField("SystemInfoName")
142                  .Primary = False
143                  .Unique = True
144                  .Required = True
145              End With
146              .Indexes.Append ndxTemplates

147              Set ndxTemplates = Nothing

148              db.TableDefs.Append tblTemplates
149          End With
150      End If

151      On Error GoTo EH_CSliceAndDice_Load
152      If Not bCategorysFound Or Not bTemplatesFound Then
153          MsgBox gsA & DatabaseName & "' is not a valid " & gsSliceAndDice & " database." & gs2EOLTab & "Please select a valid " & gsSliceAndDice & " database.", vbCritical
154          db.Close
155          ConnectionOpen = False
156          Load = False
157          Exit Function
158      End If

159      Categorys.GetAll db

    '    If Not SystemInfo Is Nothing Then
    '       SystemInfo.Clear False
    '       Set SystemInfo = Nothing
    '    End If
    '    Set SystemInfo = New CAssocArrays
    '    Set rst = db.OpenRecordset( gsSelectFrom & "SystemInfo", dbOpenForwardOnly, dbReadOnly)
    '    With rst
    '         Do Until .EOF
    '            SystemInfo(!SystemInfoName).All = nZ(!memoAttributes)
    '            .MoveNext
    '         Loop
    '         .Close
    '    End With
    '    Set rst = Nothing

160      If Not bKeepConnectionOpen Then db.Close
161      ConnectionOpen = bKeepConnectionOpen
162      Load = True

163 EH_CSliceAndDice_Load_Continue:
164      Exit Function

165 EH_CSliceAndDice_Load:
166      LogError "CSliceAndDice", "Load", Err.Number, Err.Description, Erl
167      Load = False
168      Resume EH_CSliceAndDice_Load_Continue

169      Resume
End Function

Public Function ReplaceInCategory(ByVal sCategory As String, ByVal sToFind As String, ByVal sToReplace As String) As Long
' Returns the number of replacements that occured
170      If Categorys(sCategory) Is Nothing Then
171          MsgBox "The " & gsCategory & " '" & sCategory & "' does not exist. Aborting " & gsCategory & " find and replace."
172          Exit Function
173      End If

174      ReplaceInCategory = Categorys(sCategory).Replace(sToFind, sToReplace)
End Function

Public Function Save(Optional ByVal DatabaseName As String = "SliceAndDice", Optional ByVal bKeepConnectionOpen As Boolean = True) As Boolean
175      Dim rst As Recordset
176      Dim asaCurr As CAssocArray
177      Dim AtLine As String
178      Dim CurrItem As CAssocItem

179      On Error GoTo EH_CSliceAndDice_Save

180      If Categorys Is Nothing Then Exit Function        ' Yo. Something's wrong. Don't save shit.

181      If Not ConnectionOpen Then
182          Set db = OpenDatabase(DatabaseName, False, False)
183      End If

184      If Not Categorys Is Nothing Then
185          Categorys.SaveAll db
186      End If

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

187      If Not bKeepConnectionOpen Then db.Close
188      ConnectionOpen = bKeepConnectionOpen

189      Save = True

190 EH_CSliceAndDice_Save_Continue:
191      Exit Function

192 EH_CSliceAndDice_Save:
193      LogError "CSliceAndDice", "Save(" & AtLine & gsPC, Err.Number, Err.Description, Erl
194      Save = False
195      Resume EH_CSliceAndDice_Save_Continue
End Function


Public Function sChoosePatch(Method As Integer) As String
196      Dim CurrCategory As SliceAndDice.CCategory
197      Dim CurrTemplate As SliceAndDice.CTemplate
198      Dim asaDates As New SliceAndDice.CAssocArray

199      asaDates.Clear
200      asaDates.AddInOrder = True
201      For Each CurrCategory In Categorys
202          For Each CurrTemplate In CurrCategory.Templates
203              If Method = 0 Then
204                  asaDates(vbNullString & Format$(CLng(CurrTemplate.DateModified), "00000")) = Format$(CurrTemplate.DateModified, "Mmmm D, YYYY H AM/PM")
205              Else

206              End If
207          Next CurrTemplate
208      Next CurrCategory
209      asaDates.KeyValueDelimiter = "-"
210      asaDates.ItemDelimiter = gsSC
211      sChoosePatch = sChoose(asaDates.Column(0))
End Function

Private Sub Class_Initialize()
212      Set Categorys = New CCategorys
213      Set Categorys.MyParent = Me
End Sub

Private Sub Class_Terminate()
214      Save , False
215      Set LastFound = Nothing
216      Set Categorys.MyParent = Nothing
217      Set Categorys = Nothing
End Sub

