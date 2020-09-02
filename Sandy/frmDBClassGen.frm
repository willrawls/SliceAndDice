VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E60B3BB8-E409-11D2-BA4F-0080C8C222EC}#15.1#0"; "FirmSolutions.ocx"
Begin VB.Form frmDBClassGen 
   Caption         =   "Database to Code Generator"
   ClientHeight    =   6075
   ClientLeft      =   2205
   ClientTop       =   2475
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8835
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame fraCategory 
      Caption         =   "Data Library Category"
      Height          =   660
      Left            =   4050
      TabIndex        =   2
      Top             =   -30
      Width           =   4755
      Begin VB.CommandButton cmdDeleteCategory 
         Cancel          =   -1  'True
         Caption         =   "Remove"
         Height          =   405
         Left            =   3885
         TabIndex        =   5
         ToolTipText     =   "Remove the current Slice and Dice Category from the 'Code to Generate' list."
         Top             =   180
         Width           =   765
      End
      Begin VB.CommandButton cmdAddCategory 
         Caption         =   "Add"
         Height          =   405
         Left            =   3060
         TabIndex        =   4
         ToolTipText     =   "Add a new Slice and Dice DB to Code Category"
         Top             =   180
         Width           =   765
      End
      Begin VB.ComboBox cboDataLibraryType 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "frmDBClassGen.frx":0000
         Left            =   130
         List            =   "frmDBClassGen.frx":0002
         TabIndex        =   3
         Text            =   "RDO Persisted"
         Top             =   240
         Width           =   2880
      End
   End
   Begin FirmSolutions.DataView dvwTable 
      Height          =   2145
      Left            =   60
      TabIndex        =   1
      Tag             =   "dvwTable"
      Top             =   3960
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   3784
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleWidth      =   8835
      ScaleMode       =   0
      ScaleHeight     =   2145
      HotTracking     =   -1  'True
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSComctlLib.ImageList imlTabs 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBClassGen.frx":0004
            Key             =   "Database"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBClassGen.frx":0464
            Key             =   "FieldString"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBClassGen.frx":08BC
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBClassGen.frx":0D1C
            Key             =   "FieldMemo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBClassGen.frx":117C
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBClassGen.frx":15D0
            Key             =   "FieldDate"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBClassGen.frx":1A30
            Key             =   "TableMarked"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBClassGen.frx":1E84
            Key             =   "FieldNumber"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBClassGen.frx":21A0
            Key             =   "ID"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBClassGen.frx":25F8
            Key             =   "Field"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBClassGen.frx":2A4C
            Key             =   "Marked Table"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBClassGen.frx":2EA0
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwTables 
      DragIcon        =   "frmDBClassGen.frx":31BC
      Height          =   5940
      Left            =   60
      TabIndex        =   0
      Tag             =   "tvwTables"
      Top             =   60
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   10478
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imlTabs"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwFields 
      Height          =   5310
      Left            =   4050
      TabIndex        =   6
      Tag             =   "lvwFields"
      Top             =   675
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   9366
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "imlTabs"
      SmallIcons      =   "imlTabs"
      ColHdrIcons     =   "imlTabs"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuX 
      Caption         =   "X"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open an Access97 database"
      End
      Begin VB.Menu mnuFileOpenODBC 
         Caption         =   "Open an O&DBC database"
      End
      Begin VB.Menu mnuFileOpenVBIDE 
         Caption         =   """Open"" &VB IDE"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpenPrevious 
         Caption         =   "Open a &Previously used database"
         Begin VB.Menu mnuFavorite 
            Caption         =   "-Empty-"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuFileSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFavRemoveAll 
            Caption         =   "Remove all Favorites"
         End
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "Create a &New Access97 database"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRelateOnLoad 
         Caption         =   "Relate tabes on load (S&&D mentality)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFreeAssociateTables 
         Caption         =   "Free Associate tables (No limitations)"
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuGenerate 
      Caption         =   "&Generate"
      Begin VB.Menu mnuGenerateEntireDatabase 
         Caption         =   "Entire &Database                 (everything and a wrapper class)"
      End
      Begin VB.Menu mnuGenerateClass 
         Caption         =   "Selected &Collection Class  (and Collection Member Class)"
      End
      Begin VB.Menu mnuGenerateEnterBranch 
         Caption         =   "Entire &Branch                    (selected and all children)"
      End
      Begin VB.Menu mnuGenerateCustom 
         Caption         =   "Custom Generation               (requires a custom gen template)"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuRules 
      Caption         =   "&Rules"
      Begin VB.Menu mnuRulesAutoAdd 
         Caption         =   "Automatically add"
         Begin VB.Menu mnuRulesAutoAddDateCreated 
            Caption         =   "DateCreated"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuRulesAutoAddDateModified 
            Caption         =   "DateModified"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuRulesAutoAddKey 
            Caption         =   "Name / Key (highly recommended)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuRulesAutoAddSep0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRulesAutoAddCustom 
            Caption         =   "Custom"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuRelationship 
         Caption         =   "Parent/Child Relationship"
         Begin VB.Menu mnuRulesEnforce 
            Caption         =   "Enforce referencial integrity"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuRulesCascadeUpdates 
            Caption         =   "Cascade Updates"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuRulesCascadeDeletes 
            Caption         =   "Cascade Deletes"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuRulesUseAutoNumber 
         Caption         =   "Use AutoNumber for PrimaryID"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Slice and Dice"
      End
   End
   Begin VB.Menu mnuShortcut 
      Caption         =   "Shortcut Menus"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuTable 
      Caption         =   "Table"
      Begin VB.Menu mnuTableNew 
         Caption         =   "&New Table"
      End
      Begin VB.Menu mnuTableRename 
         Caption         =   "&Rename"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToggleTableMark 
         Caption         =   "&Mark/Unmark Table"
      End
      Begin VB.Menu mnuRemoveUnmarked 
         Caption         =   "Remove Unmarked Tables from list"
      End
      Begin VB.Menu mnuRemoveMarked 
         Caption         =   "Remove Marked Tables from list"
      End
      Begin VB.Menu mnuUnhideTable 
         Caption         =   "Unhide a Table by Name"
      End
      Begin VB.Menu mnuShowAllTables 
         Caption         =   "Show all tables"
      End
      Begin VB.Menu mnuViewTableData 
         Caption         =   "&View Selected Table's Data"
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTableDelete 
         Caption         =   "Delete Table"
      End
   End
   Begin VB.Menu mnuField 
      Caption         =   "Field"
      Begin VB.Menu mnuFieldNew 
         Caption         =   "New Field"
      End
      Begin VB.Menu mnuFieldRename 
         Caption         =   "Modify"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFieldDelete 
         Caption         =   "Delete Field"
      End
   End
End
Attribute VB_Name = "frmDBClassGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbCanceled As Boolean
Private mbGenerateDatabase As Boolean
Private mbGenerateBranch As Boolean
Private mbLoadingCategories As Boolean
Private mbOpenVBIDE As Boolean

Private msClassDatabaseName As String
Private msClassDatabaseOptions As String

Private KeepDatabaseOpen As Boolean
Private IsInODBCDatabaseMode As Boolean
Private ODBCTableNamePrefix As String
Private ODBCPassword As String
Private TableList As String

Private db As Database

Private Favorites As CAssocArray
Private FavoriteCount As Long
Private RetrievingAFavoriteNow As Boolean

Public Parent As frmMain

Private NodeDragged As Node

Public Sub GenerateChildren(ByRef asaPass As CAssocArray, sDataLibraryType As String, nodChild As Node)
1        Dim CurChild As Node
2        Set CurChild = nodChild
3        Do Until CurChild Is Nothing
4            CurChild.Selected = True
5            tvwTables_NodeClick CurChild
6            If Not CurChild.Parent.Parent Is Nothing Then
7                asaPass("Parent Table Name") = CurChild.Parent.Text
8            Else
9                asaPass("Parent Table Name") = "Root"
10           End If
11           GenerateClass asaPass, sDataLibraryType, tvwTables, lvwFields
12           If gbCancelInsertion Then Exit Sub
13           If Not CurChild.Child Is Nothing Then
14               GenerateChildren asaPass, sDataLibraryType, CurChild.Child
15               If gbCancelInsertion Then Exit Sub
16           End If
17           Set CurChild = CurChild.Next
18       Loop
End Sub


Public Sub GenerateClass(ByRef asaPass As CAssocArray, sDataLibraryType As String, tvwTables As TreeView, lvwFields As ListView)
19       Dim asaV As CAssocArray
20       Dim CurListItem As ListItem
21       Dim CurChild As Node

22       Dim sTableName As String
23       Dim sFieldType As String
24       Dim sDBPCType As String
25       Dim sParentName As String
26       Dim sClassToCollect As String
27       Dim sChildTableName As String
28       Dim bSingularCollects As Boolean
29       Dim sCategoryName As String

30       On Error Resume Next

31       sCategoryName = sGetToken(sDataLibraryType, 1, gsCategoryTemplateDelimiter)

    ' Determine if the singular member of the collection will be collecting anything
32       bSingularCollects = (Not tvwTables.SelectedItem.Child Is Nothing)
33       If bSingularCollects Then
34           sClassToCollect = tvwTables.SelectedItem.Child.Text & "s"
35           sChildTableName = tvwTables.SelectedItem.Child.Text
36       Else
37           sClassToCollect = vbNullString
38           sChildTableName = tvwTables.SelectedItem.Child.Text
39       End If

    ' Generate the Collection Class
40       sTableName = sTableToPropertyName(tvwTables.SelectedItem.Text)
    'sTableName = Replace(Replace(Replace(tvwTables.SelectedItem.Text, "_", vbNullString), gsS, vbNullString), gsP, "__")
41       Err.Clear

42       If Not asaPass Is Nothing Then
43           Set asaV = asaPass
44       Else
45           Set asaV = New CAssocArray
46       End If

47       If Right$(lvwFields.ListItems(2).Key, 2) = "ID" Then
        ' Collection has a parent object
48           sParentName = lvwFields.ListItems(2).Key
49           asaV("Parent AutoNumber Field Name") = sParentName
50           asaV("Parent AutoNumber Property Name") = sTableToPropertyName(sParentName)
        'asaV("Parent AutoNumber Property Name") = Replace(Replace(Replace(sParentName, gsS, "_"), "*", "_"), "-", "_")
51           sDBPCType = vbNullString
52       ElseIf Not tvwTables.SelectedItem.Parent Is Nothing Then
53           If StrComp(tvwTables.SelectedItem.Parent.Key, "ODBC") = 0 Or StrComp(tvwTables.SelectedItem.Parent.Key, "Root") = 0 Then
            ' Collection DOESN'T have a parent object
54               asaV("Parent AutoNumber Field Name") = vbNullString
55               asaV("Parent AutoNumber Property Name") = vbNullString
56               sParentName = vbNullString
57               sDBPCType = ", No Parent"
58           Else
            ' Collection has a parent object
59               sParentName = tvwTables.SelectedItem.Text
60               asaV("Parent AutoNumber Field Name") = sParentName
61               asaV("Parent AutoNumber Property Name") = sTableToPropertyName(sParentName)
            'asaV("Parent AutoNumber Property Name") = Replace(Replace(Replace(Replace(sParentName, gsS, "_"), "*", "_"), "-", "_"), gsP, "__")
62               sDBPCType = vbNullString
63           End If
64       Else
        ' Collection DOESN'T have a parent object
65           asaV("Parent AutoNumber Field Name") = vbNullString
66           asaV("Parent AutoNumber Property Name") = vbNullString
67           sParentName = vbNullString
68           sDBPCType = ", No Parent"
69       End If

    'asaV("Collection Member Subcollection Property Name") = sClassToCollect
70       asaV("Property Name") = sClassToCollect
71       asaV("Child Table Name") = sChildTableName
72       asaV("Singular Property Name") = sChildTableName

    'asaV("Primary AutoNumber Field for Collection Member") = lvwFields.ListItems(1).Key
73       asaV("AutoNumber Field Name") = lvwFields.ListItems(1).Key
74       asaV("AutoNumber Property Name") = sTableToPropertyName(lvwFields.ListItems(1).Key)
    'asaV("AutoNumber Property Name") = Replace(Replace(Replace(lvwFields.ListItems(1).Key, gsS, "_"), "*", "_"), "-", "_")

    'asaV("Table that stores this collection") = sTableName
75       asaV("Pure Table Name") = tvwTables.SelectedItem.Text
76       asaV("Table Name") = sTableName
    'asaV("Object Name of Collection Member") = sTableName
77       asaV("Object Name") = sTableName

78       asaV("Spaced Table Name") = sInsertSpaces(sTableName)
79       asaV("Spaced Object Name") = sInsertSpaces(sTableName)
    'asaV("Label Name of Collection Member") = sInsertSpaces(sTableName)
80       asaV("Label Name") = sInsertSpaces(sTableName)

81       If Len(sDBPCType) = 0 Then
        ' Collection has a parent object
        'asaV("Field to use as Key") = lvwFields.ListItems(3).Key
82           asaV("Key Field Name") = lvwFields.ListItems(3).Key
83           asaV("Key Property Name") = sTableToPropertyName(lvwFields.ListItems(3).Key)
        'asaV("Key Property Name") = Replace(Replace(Replace(lvwFields.ListItems(3).Key, gsS, "_"), "*", "_"), "-", "_")
84       Else
        ' Collection DOESN'T have a parent object
        'asaV("Field to use as Key") = lvwFields.ListItems(2).Key
85           asaV("Key Field Name") = lvwFields.ListItems(2).Key
86           asaV("Key Property Name") = sTableToPropertyName(lvwFields.ListItems(2).Key)
        'asaV("Key Property Name") = Replace(Replace(Replace(lvwFields.ListItems(2).Key, gsS, "_"), "*", "_"), "-", "_")
87       End If

88       If bSingularCollects = False Then
89           sDBPCType = sDBPCType & ", No Child"
90       End If

91       If Not Parent.SliceAndDice.Categorys(sCategoryName).Templates("Table - " & tvwTables.SelectedItem.Text) Is Nothing Then
92           Parent.DoInsertion asaV, sDataLibraryType & "Table - " & tvwTables.SelectedItem.Text
93       End If
94       If gbCancelInsertion Then Exit Sub

95       Parent.DoInsertion asaV, sDataLibraryType & "Collection" & sDBPCType
96       If gbCancelInsertion Then Exit Sub

    ' Generate the Collection MEMBER Class
    'asaV.Clear
    'asaV("Object Name of Collection Member")=sTableName
    'asaV("Object Name")=sTableName
    'asaV("Table Name")=sTableName

    'asaV("Label Name of Collection Member")=sInsertSpaces(sTableName)
    'asaV("Label Name")=sInsertSpaces(sTableName)
    'asaV("Spaced Table Name")=sInsertSpaces(sTableName)

97       If bSingularCollects Then
98           If Len(sClassToCollect) = 0 Then sClassToCollect = "SubClass"
        'asaV("Property name of Class to collect") = sClassToCollect
99           asaV("Class to collect") = sClassToCollect
100          asaV("Property Name") = sClassToCollect
        'asaV("Collection Member Subcollection Property Name") = sClassToCollect
101          Parent.DoInsertion asaV, sDataLibraryType & "Collection Member"
102          If gbCancelInsertion Then Exit Sub
103      Else
104          Parent.DoInsertion asaV, sDataLibraryType & "Collection Member, Terminal"
105          If gbCancelInsertion Then Exit Sub
106      End If

107      For Each CurListItem In lvwFields.ListItems
        'asaV("Field Name of Property") = CurListItem.Key
108          asaV("Property Name") = sTableToPropertyName(CurListItem.Key)
        'asaV("Property Name") = Replace(Replace(Replace(CurListItem.Key, gsS, "_"), "*", "_"), "-", "_")
        'asaV("Pure Field Name") = CurListItem.Key
109          asaV("Spaced Field Name") = sInsertSpaces(CurListItem.Key)
110          asaV("Field Name") = CurListItem.Key
111          asaV("Table Name") = sTableName
112          asaV("Spaced Table Name") = sInsertSpaces(sTableName)
113          asaV("Property Type") = CurListItem.SubItems(1)
114          asaV("Property Size") = CurListItem.SubItems(2)
115          asaV("Property Length") = CurListItem.SubItems(2)
116          asaV("Field Type") = CurListItem.SubItems(1)
117          asaV("Field Length") = CurListItem.SubItems(2)

118          If Left$(CurListItem.Key, 2) = "s_" Then
            ' A field that shouldn't be messed with
119          ElseIf Right$(CurListItem.Key, 2) = "ID" Then
120              Parent.DoInsertion asaV, sDataLibraryType & "Property - " & Parent.sPropertyType(CurListItem.SubItems(1))
121              If gbCancelInsertion Then Exit Sub

122              If StrComp(CurListItem.Key, sTableName & "ID") <> 0 Then
123                  Parent.DoInsertion asaV, sDataLibraryType & "Property - 3D Link"
124                  If gbCancelInsertion Then Exit Sub
125              End If
126          Else
127              Parent.DoInsertion asaV, sDataLibraryType & "Property - " & Parent.sPropertyType(CurListItem.SubItems(1))
128              If gbCancelInsertion Then Exit Sub
129          End If
130      Next CurListItem

131      If Not Parent.SliceAndDice.Categorys(sCategoryName).Templates("Table - " & tvwTables.SelectedItem.Text & " - Finalize") Is Nothing Then
132          Parent.DoInsertion asaV, sDataLibraryType & "Table - " & tvwTables.SelectedItem.Text & " - Finalize"
133      ElseIf Not Parent.SliceAndDice.Categorys(sCategoryName).Templates("Table - " & tvwTables.SelectedItem.Text & ", Finalize") Is Nothing Then
134          Parent.DoInsertion asaV, sDataLibraryType & "Table - " & tvwTables.SelectedItem.Text & ", Finalize"
135      End If
136      If gbCancelInsertion Then Exit Sub

137      If bSingularCollects Then
        ' Generate any additional Collection MEMBER sub-collection linkages
138          sClassToCollect = tvwTables.SelectedItem.Child.Key
139          Set CurChild = tvwTables.SelectedItem.Child
140          Do Until CurChild Is Nothing
141              If sClassToCollect <> CurChild.Key Then
142                  asaV("Singular Property Name") = sTableToPropertyName(CurChild.Key)
                'asaV("Singular Property Name") = Replace(Replace(CurChild.Key, "_", vbNullString), gsS, vbNullString)
143                  asaV("Property Name") = asaV("Singular Property name") & "s"
144                  asaV("Pure Child Table Name") = CurChild.Key
145                  asaV("Child Table Name") = asaV("Singular Property Name")
146                  asaV("Pure Table Name") = tvwTables.SelectedItem.Text
147                  asaV("Table Name") = sTableName
148                  asaV("Spaced Table Name") = sInsertSpaces(sTableName)

149                  Parent.DoInsertion asaV, sDataLibraryType & "Collection Member - New Subcollection"
150                  If gbCancelInsertion Then Exit Sub
151              End If
152              Set CurChild = CurChild.Next
153          Loop
154      End If
End Sub


Public Sub TriggerClassGeneration()
155      If gbProcessing Then Exit Sub

156      Dim asaV As CAssocArray
157      Dim CurChild As Node

158      Set asaV = New CAssocArray

159      If Canceled = False Then
        'masaMisc.Clear                                                                 ' Clean out the assoc used for 'session' long inserts
160          If GenerateDatabase = False Then
            ' Generate a class for the currently selected class
161              GenerateClass asaV, cboDataLibraryType.Text & gsCategoryTemplateDelimiter, tvwTables, lvwFields
162              If GenerateBranch = True Then
                ' Generate a class for each child table and each of its children tables
163                  If Not tvwTables.SelectedItem.Child Is Nothing Then
164                      GenerateChildren asaV, cboDataLibraryType.Text & gsCategoryTemplateDelimiter, tvwTables.SelectedItem.Child
165                  End If
166              End If
167          Else
            ' Create a wrapper class
            'frmLog.tvwLog.Nodes.Clear
168              If Len(DBName) Then
169                  asaV("DSN") = DBName
170                  asaV("Database Name") = DBName
171                  asaV("Database Path") = DBPathAndFilename
172                  asaV("Spaced Database Name") = sInsertSpaces(DBName)
173              Else
174                  asaV("DSN") = sGetToken(sGetToken(ConnectString, 2, "DSN="), 1, gsSC)
175                  asaV("DSNConnect") = ConnectString
176                  asaV("ConnectString") = ConnectString
177                  asaV("Connect") = ConnectString
178                  asaV("Database Name") = ODBCDatabaseName
179                  asaV("Database Path") = "xxx No database path available. ODBC generation xxx"
180                  asaV("Spaced Database Name") = sInsertSpaces(ODBCDatabaseName)
181              End If
182              Parent.DoInsertion asaV, cboDataLibraryType.Text & gsCategoryTemplateDelimiter & "Settings"
183              If gbCancelInsertion Then Exit Sub
184              Parent.DoInsertion asaV, cboDataLibraryType.Text & gsCategoryTemplateDelimiter & "Routines"
185              If gbCancelInsertion Then Exit Sub
186              Parent.DoInsertion asaV, cboDataLibraryType.Text & gsCategoryTemplateDelimiter & "Wrapper class"
187              If gbCancelInsertion Then Exit Sub

            ' In the wrapper class, add collections for each top level DB table
188              Set CurChild = tvwTables.SelectedItem.Child
189              Do Until CurChild Is Nothing
                'asaV.Clear
190                  asaV("Pure Table Name") = CurChild.Text
191                  asaV("Table Name") = CurChild.Text
192                  asaV("Property Name") = sTableToPropertyName(CurChild.Text) & "s"
                'asaV("Plural Table Name") = asaV("Property Name")
193                  asaV("Spaced Table Name") = sInsertSpaces(sTableToPropertyName(CurChild.Text))
194                  Parent.DoInsertion asaV, cboDataLibraryType.Text & gsCategoryTemplateDelimiter & "Wrapper class - Add collection"
195                  If gbCancelInsertion Then Exit Sub
196                  Set CurChild = CurChild.Next
197              Loop

            ' Insert extra functions
            'asaV.Clear
198              If Len(DBName) Then
199                  asaV("DSN") = DBName
200                  asaV("Database Name") = DBName
201                  asaV("Database Path") = DBPathAndFilename
202                  asaV("Spaced Database Name") = sInsertSpaces(DBName)
203              Else
204                  asaV("DSN") = ConnectString
205                  asaV("Database Name") = ODBCDatabaseName
206                  asaV("Database Path") = "xxx No database path available. ODBC generation xxx"
207                  asaV("Spaced Database Name") = sInsertSpaces(ODBCDatabaseName)
208              End If

            ' Walk down the tree creating everything
209              GenerateChildren asaV, cboDataLibraryType.Text & gsCategoryTemplateDelimiter, tvwTables.SelectedItem.Child

            ' Finalize the insertions
210              On Error Resume Next
211              If Not Parent.SliceAndDice.Categorys(cboDataLibraryType.Text).Templates("Finalize") Is Nothing Then
212                  Parent.DoInsertion asaV, cboDataLibraryType.Text & gsCategoryTemplateDelimiter & "Finalize"
213              End If

            'frmLog.Show vbModal
214              If gbCancelInsertion Then
215                  MsgBox "Processing canceled at user's request."
216              Else
217                  MsgBox "Done generating database.", vbOKOnly, "CODE GENERATION COMPLETE"
218              End If
219          End If
220      End If
End Sub


Private Sub AddFavorite()
221      Dim NewKey As String
222      If Len(msClassDatabaseName) Then
223          Favorites(msClassDatabaseName) = "|||" & sNormalize(TableList)
224      Else
225          If Len(ODBCTableNamePrefix) Then
226              Favorites("ODBC: " & sGetToken(sGetToken(msClassDatabaseOptions, 2, "DSN="), 1, gsSC) & " (" & ODBCTableNamePrefix & gsPC & IIf(Len(TableList), " Limited (" & tvwTables.Nodes.Count - 1 & gsPC, vbNullString)) = msClassDatabaseOptions & "|" & ODBCTableNamePrefix & "|" & ODBCPassword & "|" & sNormalize(TableList)
227          Else
228              Favorites("ODBC: " & sGetToken(sGetToken(msClassDatabaseOptions, 2, "DSN="), 1, gsSC) & IIf(Len(TableList), " Limited (" & tvwTables.Nodes.Count - 1 & gsPC, vbNullString)) = msClassDatabaseOptions & "|" & ODBCTableNamePrefix & "|" & ODBCPassword & "|" & sNormalize(TableList)
229          End If
230      End If
231      SaveSetting App.ProductName, "DB Class Gen", "Favorites", Favorites.All
232      UpdateFavorites
End Sub


Public Sub AddTable(sTableName As String, sParentTable As String)
233      On Error GoTo EH_frmDBClassGen_AddTable
234      Dim tdfNew As TableDef
235      Dim relNew As Relation
236      Dim idxNew As Index

237      If IsInODBCDatabaseMode Then
238          If Not KeepDatabaseOpen Then Set db = OpenDatabase(msClassDatabaseName, dbDriverComplete, False, msClassDatabaseOptions)
239      Else
240          Set db = OpenDatabase(msClassDatabaseName, False, False)
241      End If
242      Set tdfNew = db.CreateTableDef(sTableName)
243      With tdfNew
        ' Create the primary unique index field
244          .Fields.Append .CreateField(sTableName & "ID", dbLong)
245          If mnuRulesUseAutoNumber.Checked Then
246              .Fields(sTableName & "ID").Attributes = (.Fields(sTableName & "ID").Attributes And dbAutoIncrField)
247          End If
248          Set idxNew = .CreateIndex(sTableName & "IDIndex")
249          With idxNew
250              .Fields.Append .CreateField(sTableName & "ID")
251              .Unique = True
252              .Primary = True
253              .Required = True

254          End With
255          .Indexes.Append idxNew

256          If Len(sParentTable) > 0 Then
            ' Attach to any indicated parent
257              .Fields.Append .CreateField(sParentTable & "ID", dbLong)
258              Set idxNew = .CreateIndex(sParentTable & "IDIndex")
259              With idxNew
260                  .Fields.Append .CreateField(sParentTable & "ID")
261                  .Required = True
262              End With
263              .Indexes.Append idxNew
264          End If
265          Set idxNew = Nothing

266          If mnuRulesAutoAddKey.Checked Then
267              .Fields.Append .CreateField(sTableName & "Name", dbText)
268          End If
269          If mnuRulesAutoAddDateCreated.Checked Then
270              .Fields.Append .CreateField("DateCreated", dbDate)
271          End If
272          If mnuRulesAutoAddDateModified.Checked Then
273              .Fields.Append .CreateField("DateModified", dbDate)
274          End If
275      End With
276      db.TableDefs.Append tdfNew

277      If Len(sParentTable) > 0 Then
        ' Create the needed cascading update/delete relationship between the parent table and child table
278          If mnuRulesEnforce.Checked Then
279              If mnuRulesCascadeUpdates.Checked Then
280                  If mnuRulesCascadeDeletes.Checked Then
281                      Set relNew = db.CreateRelation(sParentTable & "_" & sTableName, sParentTable, sTableName, dbRelationUpdateCascade + dbRelationDeleteCascade)
282                  Else
283                      Set relNew = db.CreateRelation(sParentTable & "_" & sTableName, sParentTable, sTableName, dbRelationUpdateCascade)
284                  End If
285              Else
286                  If mnuRulesCascadeDeletes.Checked Then
287                      Set relNew = db.CreateRelation(sParentTable & "_" & sTableName, sParentTable, sTableName, dbRelationDeleteCascade)
288                  Else
289                      Set relNew = db.CreateRelation(sParentTable & "_" & sTableName, sParentTable, sTableName)
290                  End If
291              End If
292              relNew.Fields.Append relNew.CreateField(sParentTable & "ID")
293              relNew.Fields(sParentTable & "ID").ForeignName = sParentTable & "ID"
294              db.Relations.Append relNew
295              Set relNew = Nothing
296          End If
297      End If

298      Set tdfNew = Nothing
299      If Not KeepDatabaseOpen Then db.Close

300      PopulateTree

301 EH_frmDBClassGen_AddTable_Continue:
302      Exit Sub

303 EH_frmDBClassGen_AddTable:
304      MsgBox "Error in SliceAndDice.frmDBClassGen_AddTable" & gs2EOLTab & Err.Description
305      Resume EH_frmDBClassGen_AddTable_Continue

306      Resume
End Sub

Public Property Get Canceled() As Boolean
307      Canceled = mbCanceled
End Property

Public Property Get ConnectString() As String
308      ConnectString = msClassDatabaseOptions
End Property

Public Property Get DBName() As String
309      DBName = sGetToken(sGetToken(msClassDatabaseName, lTokenCount(msClassDatabaseName, gsBS), gsBS), 1, ".mdb")
End Property

Public Property Get DBPathAndFilename() As String
310      DBPathAndFilename = msClassDatabaseName
End Property


Public Property Get ODBCDatabaseName() As String
311      ODBCDatabaseName = sGetToken(sGetToken(msClassDatabaseOptions, 2, "DSN="), 1, gsSC)
End Property

Private Sub RefreshTableList()
312      Dim TreeReading As CAssocArray
313      Set TreeReading = New CAssocArray
314      TreeReading.TreeToAll tvwTables
315      TableList = TreeReading.All
316      Set TreeReading = Nothing
End Sub

Private Function RemoveNodes(tvwX As TreeView, sNodeImageNameToRemove) As String
317      Dim CurrNode As Node
318      Dim sNodesLeft As String
319 RemoveNodes_Restart:
320      For Each CurrNode In tvwX.Nodes
321          If StrComp(UCase$(CurrNode.Image), UCase$(sNodeImageNameToRemove)) = 0 Then
322              tvwX.Nodes.Remove CurrNode.Key
323              GoTo RemoveNodes_Restart
324          ElseIf Not CurrNode.Parent Is Nothing Then
325              sNodesLeft = sNodesLeft & CurrNode.Key & gsSC
326          End If
327      Next CurrNode

328      For Each CurrNode In tvwX.Nodes
329          sNodesLeft = sNodesLeft & CurrNode.Key & gsSC
330      Next CurrNode

331      RemoveNodes = sNodesLeft
End Function

Public Sub UpdateFavorites()
332      On Error Resume Next
333      Dim CurrFav As Long
334      Dim CurrAssoc As CAssocItem

335      If FavoriteCount > 0 Then                         ' Clear out previous entries
336          For CurrFav = FavoriteCount To 1 Step -1
337              Unload mnuFavorite(CurrFav)
338          Next CurrFav
339          mnuFavorite(0).Caption = "-Empty-"
340          mnuFavorite(0).Enabled = False
341          FavoriteCount = 0
342      End If

343      For Each CurrAssoc In Favorites
344          If FavoriteCount > 0 Then
345              Load mnuFavorite(FavoriteCount)
346          End If
347          mnuFavorite(FavoriteCount).Caption = CurrAssoc.Key
348          mnuFavorite(FavoriteCount).Enabled = True
349          FavoriteCount = FavoriteCount + 1
350      Next CurrAssoc
End Sub


Public Property Get GenerateBranch() As Boolean
351      GenerateBranch = mbGenerateBranch
End Property

Public Property Get GenerateDatabase() As Boolean
352      GenerateDatabase = mbGenerateDatabase
End Property

Public Sub RefreshCategories()
353      On Error Resume Next
354      mbLoadingCategories = True
355      Parent.SliceAndDice.Categorys.FillList cboDataLibraryType, 1
356      cboDataLibraryType.ListIndex = FindListIndex(cboDataLibraryType, GetSetting$(App.ProductName, "DB Class Gen", "Last " & gsCategory, "RDO Persisted"))
357      mbLoadingCategories = False
End Sub

Public Sub PopulateTree()
358      On Error GoTo EH_PopulateTree
359      Dim CurTable As TableDef
360      Dim nodX     As Node
361      Dim TableListCount As Long
362      Dim asaX As CAssocArray
363      Dim CurrItem As CAssocItem

    'On Error Resume Next
    'msClassDatabaseName = sGetToken(sGetToken(msClassDatabaseName, lTokenCount(msClassDatabaseName, gsBS), gsBS), 1, ".mdb")

364      Screen.MousePointer = vbHourglass
365      If Len(msClassDatabaseName) = 0 Then
366          If Not KeepDatabaseOpen Then
367              If Len(msClassDatabaseOptions) = 0 Then
368                  msClassDatabaseOptions = "ODBC;"
369              End If
370              Set db = OpenDatabase(msClassDatabaseName, dbDriverPrompt, False, msClassDatabaseOptions)
371              If Not db Is Nothing Then
372                  IsInODBCDatabaseMode = True
373                  KeepDatabaseOpen = True

374                  If Len(ODBCPassword) = 0 Then
375                      msClassDatabaseOptions = db.Connect
376                      ODBCPassword = InputBox("What is the database password you just entered ?" & gsEolTab & "NOTE: Entering it again here prevents you having to reenter it throughout this session.", "ENTER ODBC DATABASE PASSWORD")
377                      If Len(ODBCPassword) Then
378                          If Right$(msClassDatabaseOptions, 1) <> gsSC Then
379                              msClassDatabaseOptions = msClassDatabaseOptions & ";PWD=" & ODBCPassword & gsSC
380                          Else
381                              msClassDatabaseOptions = msClassDatabaseOptions & "PWD=" & ODBCPassword & gsSC
382                          End If
383                      End If
384                  End If
385                  If Len(ODBCTableNamePrefix) = 0 Then
386                      If db.TableDefs.Count > 100 Then
387                          ODBCTableNamePrefix = InputBox("There are " & db.TableDefs.Count & " tables in this database." & gs2EOL & "Would you like to limit the tables included by a table name prefix ?" & gsEolTab & "(Leave blank to include all tables)", "LIMIT TABLES BY TABLE NAME PREFIX")
388                      End If
389                  End If
390                  If Not RetrievingAFavoriteNow Then
391                      AddFavorite
392                  End If
393              End If
394          End If
395      Else
396          msClassDatabaseOptions = vbNullString
397          Set db = OpenDatabase(msClassDatabaseName, False, False)
398          IsInODBCDatabaseMode = False
399          KeepDatabaseOpen = False
400          If Not RetrievingAFavoriteNow Then
401              AddFavorite
402          End If
403      End If

404      lvwFields.ListItems.Clear
405      With tvwTables.Nodes
406          .Clear
407          If IsInODBCDatabaseMode Then
408              Set nodX = .Add(, , "Root", "ODBC", "Database", "Database")
409          Else
410              Set nodX = .Add(, , "Root", sGetToken(sGetToken(msClassDatabaseName, lTokenCount(msClassDatabaseName, gsBS), gsBS), 1, ".mdb"), "Database", "Database")
411          End If
412          nodX.ExpandedImage = "Database"
413          nodX.Expanded = True

414          If Len(TableList) = 0 Then
415              Screen.MousePointer = vbHourglass
416              TableListCount = 0
417              For Each CurTable In db.TableDefs
418                  If Left$(CurTable.Name, 4) <> "MSys" And Left$(CurTable.Name, 4) <> "SYS." And Left$(CurTable.Name, 4) <> "ALL_" Then
419                      If Len(ODBCTableNamePrefix) = 0 Or UCase$(Left$(CurTable.Name, Len(ODBCTableNamePrefix))) = UCase$(ODBCTableNamePrefix) Then
420                          Set nodX = .Add("Root", tvwChild, CurTable.Name, CurTable.Name, "Table", "Table")
421                          nodX.ExpandedImage = "Table"
422                          nodX.Expanded = True
423                          TableListCount = TableListCount + 1
424                      End If
425                  End If
426              Next CurTable
427          ElseIf InStr(TableList, "<ICON>") = 0 And InStr(TableList, "<CHILD>") = 0 And InStr(TableList, "<ENDCHILD>") = 0 Then
428              Set asaX = New CAssocArray
429              asaX.ItemDelimiter = gsSC
430              asaX.All = TableList
431              For Each CurrItem In asaX
432                  Set nodX = .Add("Root", tvwChild, CurrItem.Key, CurrItem.Key, "Table", "Table")
433                  nodX.ExpandedImage = "Table"
434                  nodX.Expanded = True
435              Next CurrItem
436              TableListCount = asaX.Count
437              asaX.Clear
438              Set asaX = Nothing
439          Else
440              Set asaX = New CAssocArray
441              asaX.All = TableList
442              tvwTables.Nodes.Clear
443              asaX.FillTreeNode tvwTables, Nothing, "Table", True
444              Set asaX = Nothing
445          End If

446          If Not mnuRelateOnLoad.Checked Then
447              GoTo SKIP_RELATION_STEP
448          End If

449          If db.TableDefs.Count > 40 Then
450              If bUserSure("There are a lot of table to attempt to relate. If there are no " & gsSliceAndDice & " relations, this step can be skipped." & gsEolTab & "Would you like to skip this step ?") Then
451                  GoTo SKIP_RELATION_STEP
452              End If
453          End If
454          Screen.MousePointer = vbHourglass
455          On Error Resume Next
456          For Each CurTable In db.TableDefs
457              If Left$(CurTable.Name, 4) <> "MSys" And Left$(CurTable.Name, 4) <> "SYS." And Left$(CurTable.Name, 4) <> "ALL_" Then
458                  If Len(ODBCTableNamePrefix) = 0 Or UCase$(Left$(CurTable.Name, Len(ODBCTableNamePrefix))) = UCase$(ODBCTableNamePrefix) Then
459                      If CurTable.Fields.Count > 1 Then
460                          Screen.MousePointer = vbHourglass
461                          If Right$(CurTable.Fields(1).Name, 2) = "ID" Then
462                              Set nodX = Nothing
463                              Set nodX = .Item(Left$(CurTable.Fields(1).Name, Len(CurTable.Fields(1).Name) - 2))
464                              If Not nodX Is Nothing Then
465                                  Set .Item(CurTable.Name).Parent = nodX
466                              End If
467                          End If
468                      End If
469                  End If
470              End If
471          Next CurTable

472 SKIP_RELATION_STEP:
473      End With
474      If Not KeepDatabaseOpen Then db.Close

475 EH_PopulateTree_Continue:
476      Set nodX = Nothing
477      Screen.MousePointer = vbDefault
478      Exit Sub

479 EH_PopulateTree:
480      LogError "frmDBClassGen", "PopulateTree", Err.Number, Err.Description, Erl
481      Resume EH_PopulateTree_Continue

482      Resume
End Sub

Public Function sFieldType(iFieldType As Long) As String
    Select Case iFieldType
        Case dbBigInt: sFieldType = "Big Integer"
483          Case dbBinary: sFieldType = "Binary"
484          Case dbBoolean: sFieldType = "Boolean"
485          Case dbByte: sFieldType = "Byte"
486          Case dbChar: sFieldType = "Char"
487          Case dbCurrency: sFieldType = "Currency"
488          Case dbDate: sFieldType = "Date / Time"
489          Case dbDecimal: sFieldType = "Decimal"
490          Case dbDouble: sFieldType = "Double"
491          Case dbFloat: sFieldType = "Float"
492          Case dbGUID: sFieldType = "Guid"
493          Case dbInteger: sFieldType = "Integer"
494          Case dbLong: sFieldType = "Long"
495          Case dbLongBinary: sFieldType = "Long Binary (OLE Object)"
496          Case dbMemo: sFieldType = "Memo"
497          Case dbNumeric: sFieldType = "Numeric"
498          Case dbSingle: sFieldType = "Single"
499          Case dbText: sFieldType = "Text"
500          Case dbTime: sFieldType = "Time"
501          Case dbTimeStamp: sFieldType = "Time Stamp"
502          Case dbVarBinary: sFieldType = "VarBinary"
503      End Select
End Function

Private Sub cboDataLibraryType_Click()
504      If mbLoadingCategories Then Exit Sub
505      SaveSetting App.ProductName, "DB Class Gen", "Last " & gsCategory, cboDataLibraryType.Text
End Sub

Private Sub cmdAddCategory_Click()
506      Dim sCategoryChosen As String
507      With Parent.SliceAndDice.Categorys
508          sCategoryChosen = .Choose(0)
509          If Len(sCategoryChosen) > 0 Then
510              .Item(sCategoryChosen).CategoryType = 1
511              Parent.SliceAndDice.Save
512          End If
513      End With
514      RefreshCategories
End Sub

Private Sub cmdDeleteCategory_Click()
515      Dim sCategoryChosen As String
516      With Parent.SliceAndDice.Categorys
517          sCategoryChosen = .Choose(1)
518          If Len(sCategoryChosen) > 0 Then
519              .Item(sCategoryChosen).CategoryType = 0
520              Parent.SliceAndDice.Save
521          End If
522      End With
523      RefreshCategories
End Sub


Private Sub Form_Initialize()

' LogEvent "frmDBClassGen: Initialize"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
524      If UnloadMode = vbFormControlMenu Then
525          Cancel = True
526      End If
527      mnuFileExit_Click
End Sub

Private Sub Form_Resize()
528      On Error Resume Next
529      If Not dvwTable.Visible Then
530          lvwFields.Move ScaleWidth - fraCategory.Width, lvwFields.Top, fraCategory.Width, ScaleHeight - lvwFields.Top
531          tvwTables.Move 0, 60, ScaleWidth - fraCategory.Width, ScaleHeight
532          fraCategory.Move lvwFields.Left, 0
533      Else
534          lvwFields.Move ScaleWidth - fraCategory.Width, lvwFields.Top, fraCategory.Width, ScaleHeight - fraCategory.Height - dvwTable.Height
535          tvwTables.Move 0, 60, ScaleWidth - fraCategory.Width, ScaleHeight - dvwTable.Height
536          fraCategory.Move lvwFields.Left, 0
537          dvwTable.Move 0, lvwFields.Top + lvwFields.Height, ScaleWidth
538      End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
539      SaveFormPosition Me
End Sub

Private Sub lvwFields_DblClick()
540      mnuFieldNew_Click
End Sub

Private Sub lvwFields_KeyPress(KeyAscii As Integer)
541      If KeyAscii = 13 Then
542          mnuFieldNew_Click
543      End If
End Sub

Private Sub mnuFavorite_Click(Index As Integer)
544      Screen.MousePointer = vbDefault
545      With Favorites(mnuFavorite(Index).Caption)
546          If Len(sGetToken(.Value, 1, "|")) = 0 Then
547              msClassDatabaseName = sGetToken(.Key, 1, "|")
548              msClassDatabaseOptions = vbNullString
549              ODBCTableNamePrefix = vbNullString
550              ODBCPassword = vbNullString
551              TableList = sDenormalize(sAfter(.Value, 3, "|"))
552          Else
553              msClassDatabaseName = vbNullString
554              msClassDatabaseOptions = sGetToken(.Value, 1, "|")
555              ODBCTableNamePrefix = sGetToken(.Value, 2, "|")
556              ODBCPassword = sGetToken(.Value, 3, "|")
557              TableList = sDenormalize(sAfter(.Value, 3, "|"))
558          End If
559          IsInODBCDatabaseMode = False
560          KeepDatabaseOpen = False
561          RetrievingAFavoriteNow = True
562          PopulateTree
563          RetrievingAFavoriteNow = False
564      End With
End Sub

Private Sub mnuFavRemoveAll_Click()
565      If bUserSure() Then
566          Favorites.All = vbNullString
567          SaveSetting App.ProductName, "DB Class Gen", "Favorites", vbNullString
568          UpdateFavorites
569      End If
End Sub


Private Sub mnuFileExit_Click()
570      mbCanceled = True
571      lvwFields.ListItems.Clear
572      SaveFormPosition Me
573      Hide
End Sub

Private Sub mnuFileNew_Click()
574      Dim sDatabasePath    As String
575      Dim sNewDatabaseName As String

576      sDatabasePath = Trim$(BrowseForFolder(hwnd, "Where should database go ?"))
577      If Len(sDatabasePath) = 0 Then Exit Sub

578      sNewDatabaseName = Trim$(InputBox("What should the name of the new database be ?", "CREATE BLANK DATABASE"))
579      If Len(sNewDatabaseName) = 0 Then Exit Sub

580      If Right$(sDatabasePath, 1) <> gsBS Then sDatabasePath = sDatabasePath & gsBS
581      If Right$(LCase$(sNewDatabaseName), 4) <> ".mdb" Then sNewDatabaseName = sDatabasePath & sNewDatabaseName & ".mdb"

582      On Error Resume Next
583      Set db = CreateDatabase(sNewDatabaseName, dbLangGeneral, dbVersion30)
584      db.Close

585      msClassDatabaseName = sNewDatabaseName
586      msClassDatabaseOptions = vbNullString
587      ODBCTableNamePrefix = vbNullString
588      ODBCPassword = vbNullString
589      IsInODBCDatabaseMode = False
590      KeepDatabaseOpen = False
591      PopulateTree
End Sub

Private Sub mnuFileOpenODBC_Click()
592      mbOpenVBIDE = False
593      msClassDatabaseName = vbNullString
594      msClassDatabaseOptions = vbNullString
595      ODBCTableNamePrefix = vbNullString
596      ODBCPassword = vbNullString
597      IsInODBCDatabaseMode = False
598      KeepDatabaseOpen = False
599      PopulateTree
End Sub

Private Sub mnuFileOpenVBIDE_Click()
600      mbOpenVBIDE = True
601      msClassDatabaseName = vbNullString
602      msClassDatabaseOptions = vbNullString
603      ODBCTableNamePrefix = vbNullString
604      ODBCPassword = vbNullString
605      IsInODBCDatabaseMode = False
606      KeepDatabaseOpen = False
607      PopulateTree
End Sub


Private Sub mnuFreeAssociateTables_Click()
608      mnuFreeAssociateTables.Checked = Not mnuFreeAssociateTables.Checked
609      SaveSetting App.ProductName, "DB Class Gen", "Free Associate Tables", mnuFreeAssociateTables.Checked
End Sub

Private Sub mnuGenerateClass_Click()
610      If lvwFields.ListItems.Count = 0 Then
611          MsgBox "Please select a table first."
612          Exit Sub
613      End If

614      mbCanceled = False
615      mbGenerateBranch = False
616      mbGenerateDatabase = False
617      Hide
618      TriggerClassGeneration
End Sub

Private Sub mnuGenerateEnterBranch_Click()
619      If lvwFields.ListItems.Count = 0 Then
620          MsgBox "Please select a branch first."
621          Exit Sub
622      End If

623      mbCanceled = False
624      mbGenerateBranch = True
625      mbGenerateDatabase = False
626      Hide
627      TriggerClassGeneration
End Sub


Private Sub mnuGenerateEntireDatabase_Click()
628      If tvwTables.Nodes.Count = 0 Then
629          MsgBox "Please select a database first."
630          Exit Sub
631      End If

632      tvwTables.Nodes(1).Selected = True
633      lvwFields.ListItems.Clear

634      mbCanceled = False
635      mbGenerateBranch = False
636      mbGenerateDatabase = True
637      Hide
638      TriggerClassGeneration
End Sub

Public Sub mnuFileOpen_Click()
639      mbOpenVBIDE = False
640      msClassDatabaseName = Parent.sChooseDatabase()
641      msClassDatabaseOptions = vbNullString
642      ODBCTableNamePrefix = vbNullString
643      ODBCPassword = vbNullString
644      IsInODBCDatabaseMode = False
645      KeepDatabaseOpen = False
646      If Len(msClassDatabaseName) > 0 Then
647          PopulateTree
648      End If
End Sub

Private Sub Form_Load()
649      On Error Resume Next

650      RefreshCategories

651      mnuRulesAutoAddKey.Checked = GetSetting(App.ProductName, "DB Class Gen", "AutoAddKey", True)
652      mnuRulesAutoAddDateModified.Checked = GetSetting(App.ProductName, "DB Class Gen", "AutoAddDateModified", True)
653      mnuRulesAutoAddDateCreated.Checked = GetSetting(App.ProductName, "DB Class Gen", "AutoAddCreated", True)

654      mnuRulesEnforce.Checked = GetSetting(App.ProductName, "DB Class Gen", "Enforce", True)
655      mnuRulesCascadeUpdates.Checked = GetSetting(App.ProductName, "DB Class Gen", "CascadeUpdates", True)
656      mnuRulesCascadeDeletes.Checked = GetSetting(App.ProductName, "DB Class Gen", "CascadeDeletes", True)

657      mnuRulesUseAutoNumber.Checked = GetSetting(App.ProductName, "DB Class Gen", "UseAutoNumber", True)
658      mnuViewTableData.Checked = GetSetting(App.ProductName, "DB Class Gen", "ViewTableData", False)
659      mnuRelateOnLoad.Checked = GetSetting(App.ProductName, "DB Class Gen", "Relate on load", True)
660      mnuFreeAssociateTables.Checked = GetSetting(App.ProductName, "DB Class Gen", "Free Asoociate Tables", True)

661      If mnuRulesEnforce.Checked Then
662          mnuRulesCascadeUpdates.Enabled = True
663          mnuRulesCascadeDeletes.Enabled = True
664      Else
665          mnuRulesCascadeUpdates.Enabled = False
666          mnuRulesCascadeDeletes.Enabled = False
667      End If

668      Set Favorites = New CAssocArray
669      Favorites.KeyValueDelimiter = "=="
670      Favorites.FieldDelimiter = "|"
671      Favorites.ItemDelimiter = "<EOE>"
672      Favorites.All = GetSetting$(App.ProductName, "DB Class Gen", "Favorites", vbNullString)

673      UpdateFavorites

674      LoadFormPosition Me

675      dvwTable.Visible = mnuViewTableData.Checked
End Sub


Private Sub Form_Terminate()
676      SaveFormPosition Me
677      Set Favorites = Nothing
678      Set Parent = Nothing
    ' LogEvent "frmDBClassGen: Terminate"
End Sub


Private Sub lvwFields_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
679      On Error Resume Next
680      Dim ItemClicked As ListItem

681      If Button = vbRightButton Then
682          Set ItemClicked = lvwFields.HitTest(X, Y)
683          If Not ItemClicked Is Nothing Then
684              ItemClicked.Selected = True
685              PopupMenu mnuField
686          End If
687      End If
End Sub


Private Sub mnuFieldDelete_Click()
688      Dim sParentTable As String
689      Dim sTable       As String
690      Dim sField       As String

691      On Error Resume Next
692      If tvwTables.SelectedItem.Parent.Key = msClassDatabaseName Then
693          sParentTable = vbNullString
694      Else
695          sParentTable = tvwTables.SelectedItem.Parent.Key
696      End If
697      sTable = tvwTables.SelectedItem.Key
698      sField = lvwFields.SelectedItem.Key

    Select Case sField
        Case sTable & "ID", sTable & "Name", "DateCreated", "DateModified", sParentTable & "ID"
699              MsgBox "Can't delete that field (required for correct object/database operation), sorry."
700              Exit Sub
701      End Select

702      If IsInODBCDatabaseMode Then
703          If Not KeepDatabaseOpen Then Set db = OpenDatabase(msClassDatabaseName, dbDriverComplete, False, msClassDatabaseOptions)
704      Else
705          Set db = OpenDatabase(msClassDatabaseName, False, False)
706      End If
707      db.TableDefs(sTable).Fields.Delete sField
708      If Not KeepDatabaseOpen Then db.Close

709      tvwTables_NodeClick tvwTables.SelectedItem
End Sub

Private Sub mnuFieldNew_Click()
710      Dim CurTable   As TableDef
711      Dim CurItem    As ListItem
712      Dim NewField   As frmFieldType

713      Set NewField = New frmFieldType

714 mnuFieldNew_Click_TryAgain:
715      With NewField
716          On Error Resume Next
717          Err.Clear
718          .Show vbModal, Me
719          If Err.Number <> 0 Then Exit Sub
720          If .Canceled = True Then Exit Sub

721          For Each CurItem In lvwFields.ListItems
722              If UCase$(CurItem.Text) = UCase$(.FieldName) Then
723                  MsgBox "That field already exists in this table... Try again."
724                  GoTo mnuFieldNew_Click_TryAgain
725              End If
726          Next CurItem

727          On Error Resume Next
728          If IsInODBCDatabaseMode Then
729              If Not KeepDatabaseOpen Then Set db = OpenDatabase(msClassDatabaseName, dbDriverComplete, False, msClassDatabaseOptions)
730          Else
731              Set db = OpenDatabase(msClassDatabaseName, False, False)
732          End If
733          Set CurTable = db.TableDefs(tvwTables.SelectedItem.Key)
734          Err.Clear
735          If .dbFieldType = dbText Then
736              CurTable.Fields.Append CurTable.CreateField(.FieldName, .dbFieldType, .Length)
737          Else
738              CurTable.Fields.Append CurTable.CreateField(.FieldName, .dbFieldType)
739          End If
740          If Err.Number Then
741              MsgBox "Error occured adding field '" & .FieldName & "' to table '" & CurTable.Name & gsA & gsEolTab & "Err #" & Err.Number & gsEolTab & "Desc:" & Err.Description
742              Err.Clear
743          End If
744          If Not KeepDatabaseOpen Then db.Close

745          tvwTables_NodeClick tvwTables.SelectedItem
746      End With
End Sub

Private Sub mnuRelateOnLoad_Click()
747      mnuRelateOnLoad.Checked = Not mnuRelateOnLoad.Checked
748      SaveSetting App.ProductName, "DB Class Gen", "Relate on load", mnuRelateOnLoad.Checked
End Sub

Private Sub mnuRemoveMarked_Click()
749      RemoveNodes tvwTables, "TableMarked"
750      RefreshTableList
751      AddFavorite
End Sub

Private Sub mnuRemoveUnmarked_Click()
752      RemoveNodes tvwTables, "Table"
753      RefreshTableList
754      AddFavorite
End Sub

Private Sub mnuRulesAutoAddDateCreated_Click()
755      mnuRulesAutoAddDateCreated.Checked = Not mnuRulesAutoAddDateCreated.Checked
756      SaveSetting App.ProductName, "DB Class Gen", "AutoAddDateCreated", mnuRulesAutoAddDateCreated.Checked
End Sub

Private Sub mnuRulesAutoAddDateModified_Click()
757      mnuRulesAutoAddDateModified.Checked = Not mnuRulesAutoAddDateModified.Checked
758      SaveSetting App.ProductName, "DB Class Gen", "AutoAddDateModified", mnuRulesAutoAddDateModified.Checked
End Sub


Private Sub mnuRulesAutoAddKey_Click()
759      mnuRulesAutoAddKey.Checked = Not mnuRulesAutoAddKey.Checked
760      SaveSetting App.ProductName, "DB Class Gen", "AutoAddKey", mnuRulesAutoAddKey.Checked
End Sub

Private Sub mnuRulesCascadeDeletes_Click()
761      mnuRulesCascadeDeletes.Checked = Not mnuRulesCascadeDeletes.Checked
762      SaveSetting App.ProductName, "DB Class Gen", "CascadeDeletes", mnuRulesCascadeDeletes.Checked
End Sub


Private Sub mnuRulesCascadeUpdates_Click()
763      mnuRulesCascadeUpdates.Checked = Not mnuRulesCascadeUpdates.Checked
764      SaveSetting App.ProductName, "DB Class Gen", "CascadeUpdates", mnuRulesCascadeUpdates.Checked
End Sub

Private Sub mnuRulesEnforce_Click()
765      mnuRulesEnforce.Checked = Not mnuRulesEnforce.Checked
766      SaveSetting App.ProductName, "DB Class Gen", "Enforce", mnuRulesEnforce.Checked
767      If mnuRulesEnforce.Checked Then
768          mnuRulesCascadeUpdates.Enabled = True
769          mnuRulesCascadeDeletes.Enabled = True
770      Else
771          mnuRulesCascadeUpdates.Enabled = False
772          mnuRulesCascadeDeletes.Enabled = False
773      End If
End Sub

Private Sub mnuRulesUseAutoNumber_Click()
774      mnuRulesUseAutoNumber.Checked = Not mnuRulesUseAutoNumber.Checked
775      SaveSetting App.ProductName, "DB Class Gen", "UseAutoNumber", mnuRulesUseAutoNumber.Checked
End Sub

Private Sub mnuShowAllTables_Click()
776      TableList = vbNullString
777      AddFavorite
778      PopulateTree
End Sub

Private Sub mnuTableDelete_Click()
779      On Error Resume Next
780      If bUserSure("This will PERMANENTLY remove the table selected." & gsEolTab & "ARE YOU ABSOLUTELY SURE ?") Then
781          If IsInODBCDatabaseMode Then
782              If Not KeepDatabaseOpen Then Set db = OpenDatabase(msClassDatabaseName, dbDriverComplete, False, msClassDatabaseOptions)
783          Else
784              Set db = OpenDatabase(msClassDatabaseName, False, False)
785          End If
786          db.TableDefs.Delete tvwTables.SelectedItem.Key
787          If Not KeepDatabaseOpen Then db.Close
788          PopulateTree
789      End If
End Sub

Private Sub mnuTableNew_Click()
790      Dim sTable As String

791 mnuTableNew_Click_TryAgain:
792      sTable = Replace(Trim$(InputBox("What should the name of the new table be ?" & gsEolTab & "Note: Table names should be singular, such as:" & gsEolTab & "Book, Publisher, etc.")), gsS, vbNullString)
793      If Len(sTable) = 0 Then Exit Sub

794      If Right$(sTable, 1) = "s" Then
795          MsgBox "Table names MUST be singular."
796          GoTo mnuTableNew_Click_TryAgain
797      End If

798      If tvwTables.SelectedItem.Text <> DBName Then
799          AddTable sTable, tvwTables.SelectedItem.Text
800      Else
801          AddTable sTable, vbNullString
802      End If

803      PopulateTree
End Sub

Private Sub mnuToggleTableMark_Click()
804      If tvwTables.SelectedItem.Image = "Table" Then
805          tvwTables.SelectedItem.Image = "TableMarked"
806          tvwTables.SelectedItem.ExpandedImage = "TableMarked"
807          tvwTables.SelectedItem.SelectedImage = "TableMarked"
808      Else
809          tvwTables.SelectedItem.Image = "Table"
810          tvwTables.SelectedItem.ExpandedImage = "Table"
811          tvwTables.SelectedItem.SelectedImage = "Table"
812      End If

End Sub

Private Sub mnuUnhideTable_Click()
813      On Error Resume Next
814      Dim sTableName As String
815      sTableName = InputBox("What is the fully qualified name of the table to unhide ?", "UNHIDE A TABLE")
816      If Len(sTableName) Then
817          With tvwTables.Nodes.Add(tvwTables.SelectedItem.Key, tvwChild, sTableName, sTableName, "Table", "Table")
818              .ExpandedImage = "Table"
819              .Expanded = True
820          End With
821      End If
822      RefreshTableList
823      AddFavorite
End Sub

Private Sub mnuViewTableData_Click()
824      On Error Resume Next
825      mnuViewTableData.Checked = Not mnuViewTableData.Checked
826      SaveSetting App.ProductName, "DB Class Gen", "ViewTableData", mnuViewTableData.Checked
827      dvwTable.Visible = mnuViewTableData.Checked
828      Form_Resize
End Sub

Private Sub mnuX_Click()
829      mnuFileExit_Click
End Sub

Private Sub tvwTables_DblClick()
830      mnuTableNew_Click
End Sub

Private Sub tvwTables_DragDrop(Source As Control, X As Single, Y As Single)
831      On Error Resume Next
832      Dim NodeDroppedOnto As Node

833      If Source.Tag = "tvwTables" Then
834          If Not NodeDragged Is Nothing Then
835              Set NodeDroppedOnto = Nothing
836              Set NodeDroppedOnto = tvwTables.HitTest(X, Y)
837              If Not NodeDroppedOnto Is Nothing Then
838                  Set NodeDragged.Parent = NodeDroppedOnto
839                  RefreshTableList
840                  AddFavorite
841              End If
842          End If
843      End If
End Sub

Private Sub tvwTables_KeyPress(KeyAscii As Integer)
844      If KeyAscii = 13 Then
845          mnuTableNew_Click
846      End If
End Sub

Private Sub tvwTables_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
847      If Button = vbLeftButton And (Shift And vbShiftMask) <> 0 Then
848          Set NodeDragged = Nothing
849          Set NodeDragged = tvwTables.HitTest(X, Y)
850          If Not NodeDragged Is Nothing Then
851              tvwTables.Drag vbBeginDrag
852          End If
853      End If
End Sub

Private Sub tvwTables_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
854      Dim NodeClicked As Node

855      If Button = vbRightButton Then
856          Set NodeClicked = tvwTables.HitTest(X, Y)
857          If Not NodeClicked Is Nothing Then
858              NodeClicked.Selected = True
859              PopupMenu mnuTable
860          End If
861      End If
End Sub

Public Sub tvwTables_NodeClick(ByVal Node As MSComctlLib.Node)
862      Dim CurTable As TableDef
863      Dim CurField As Field
864      Dim nodX     As Node
865      Dim litX     As ListItem
866      Dim sIcon    As String

867      On Error Resume Next
868      Screen.MousePointer = vbHourglass
869      If IsInODBCDatabaseMode Then
870          If Not KeepDatabaseOpen Then Set db = OpenDatabase(msClassDatabaseName, dbDriverComplete, False, msClassDatabaseOptions)
871      Else
872          Set db = OpenDatabase(msClassDatabaseName, False, False)
873      End If
874      With lvwFields.ListItems
875          .Clear
876          With lvwFields.ColumnHeaders
877              .Clear
878              .Add , "Field Name", "Field Name", 2600
879              .Add , "Field Type", "Type", 1000
880              .Add , "Field Length", "Length", 500
881          End With
882          lvwFields.View = lvwReport
883          Set CurTable = db.TableDefs(tvwTables.SelectedItem.Text)

884          For Each CurField In CurTable.Fields
885              If Right$(CurField.Name, 2) = "ID" Then
886                  If Left$(CurField.Name, Len(CurField.Name) - 2) = CurTable.Name Then
887                      sIcon = "Key"
888                  Else
889                      sIcon = "ID"
890                  End If
891              ElseIf CurField.Type = dbDate Then
892                  sIcon = "FieldDate"
893              ElseIf CurField.Type = dbMemo Then
894                  sIcon = "FieldMemo"
895              ElseIf CurField.Type = dbText Then
896                  sIcon = "FieldString"
897              Else
898                  sIcon = "FieldNumber"
899              End If
900              Set litX = .Add(, CurField.Name, CurField.Name, sIcon, sIcon)
901              litX.SubItems(1) = sFieldType(CurField.Type)
902              litX.SubItems(2) = CurField.Size
903          Next CurField
904      End With

905      If dvwTable.Visible And Not tvwTables.SelectedItem.Parent Is Nothing Then
906          On Error Resume Next
907          With dvwTable
908              .DatabaseName = "x"
909              .RecordSource = gsSelectFrom & "[" & tvwTables.SelectedItem.Text & "]"
910              .View = lvwReport
911              .GridLines = True
912              .FullRowSelect = True
913              .Requery db
914          End With
915      End If

916      If Not KeepDatabaseOpen Then db.Close
917      Set nodX = Nothing
918      Screen.MousePointer = vbDefault
End Sub

