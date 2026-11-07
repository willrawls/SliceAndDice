VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmODBCClassGen 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "ODBC Database to Code Generator"
   ClientHeight    =   6165
   ClientLeft      =   1710
   ClientTop       =   2910
   ClientWidth     =   8835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Caption         =   " Code to Generate "
      Height          =   660
      Left            =   4050
      TabIndex        =   2
      Top             =   -15
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
         ItemData        =   "frmODBCClassGen.frx":0000
         Left            =   135
         List            =   "frmODBCClassGen.frx":0002
         TabIndex        =   3
         Text            =   "RDO Persisted"
         Top             =   240
         Width           =   2880
      End
   End
   Begin ComctlLib.ListView lvwFields 
      Height          =   5310
      Left            =   4050
      TabIndex        =   0
      Top             =   705
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   9366
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "imlDB"
      SmallIcons      =   "imlDB"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Field Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Length"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.TreeView tvwTables 
      DragIcon        =   "frmODBCClassGen.frx":0004
      Height          =   5940
      Left            =   60
      TabIndex        =   1
      Top             =   75
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   10478
      _Version        =   327682
      Indentation     =   265
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imlDB"
      Appearance      =   1
   End
   Begin ComctlLib.ImageList imlDB 
      Left            =   390
      Top             =   -225
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmODBCClassGen.frx":0446
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmODBCClassGen.frx":0760
            Key             =   "Field"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmODBCClassGen.frx":0872
            Key             =   "Database"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmODBCClassGen.frx":0B8C
            Key             =   "Date"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmODBCClassGen.frx":0EA6
            Key             =   "ID"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmODBCClassGen.frx":11C0
            Key             =   "Key"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuX 
      Caption         =   "X"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open an Access97 database"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "Create a &New Access97 database"
      End
      Begin VB.Menu mnuFileOpen2 
         Caption         =   "Open an ODBC data source"
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
         Caption         =   "Enter &Database                 (everything and a wrapper class)"
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
      Begin VB.Menu mnuTable 
         Caption         =   "Table Right Click"
         Begin VB.Menu mnuTableNew 
            Caption         =   "New Table"
         End
         Begin VB.Menu mnuTableRename 
            Caption         =   "Rename"
         End
         Begin VB.Menu mnuSep0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTableDelete 
            Caption         =   "Delete Table"
         End
      End
      Begin VB.Menu mnuField 
         Caption         =   "Field Right Click"
         Begin VB.Menu mnuFieldNew 
            Caption         =   "New Field"
         End
         Begin VB.Menu mnuFieldRename 
            Caption         =   "Modify"
         End
         Begin VB.Menu mnuSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFieldDelete 
            Caption         =   "Delete Field"
         End
      End
   End
End
Attribute VB_Name = "frmODBCClassGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bCanceled As Boolean
Private m_bGenerateDatabase As Boolean
Private m_bGenerateBranch As Boolean

Private m_sClassDatabaseName As String

Public Parent As frmMain
Public Sub AddTable(sTableName As String, sParentTable As String)
On Error GoTo EH_frmODBCClassGen_AddTable
    Dim db     As Database
    Dim tdfNew As TableDef
    Dim relNew As Relation
    Dim idxNew As Index

    Set db = OpenDatabase(m_sClassDatabaseName, "ODBC;")
        Set tdfNew = db.CreateTableDef(sTableName)
        With tdfNew
           ' Create the primary unique index field
             .Fields.Append .CreateField(sTableName & "ID", dbLong)
             If mnuRulesUseAutoNumber.Checked Then
                .Fields(sTableName & "ID").Attributes = (.Fields(sTableName & "ID").Attributes And dbAutoIncrField)
             End If
             Set idxNew = .CreateIndex(sTableName & "IDIndex")
             With idxNew
                  .Fields.Append .CreateField(sTableName & "ID")
                  .Unique = True
                  .Primary = True
                  .Required = True
                  
             End With
             .Indexes.Append idxNew

             If Len(sParentTable) > 0 Then
              ' Attach to any indicated parent
                .Fields.Append .CreateField(sParentTable & "ID", dbLong)
                Set idxNew = .CreateIndex(sParentTable & "IDIndex")
                With idxNew
                     .Fields.Append .CreateField(sParentTable & "ID")
                     .Required = True
                End With
                .Indexes.Append idxNew
             End If
             Set idxNew = Nothing

             If mnuRulesAutoAddKey.Checked Then
                .Fields.Append .CreateField(sTableName & "Name", dbText)
             End If
             If mnuRulesAutoAddDateCreated.Checked Then
                .Fields.Append .CreateField("DateCreated", dbDate)
             End If
             If mnuRulesAutoAddDateModified.Checked Then
                .Fields.Append .CreateField("DateModified", dbDate)
             End If
        End With
        db.TableDefs.Append tdfNew

        If Len(sParentTable) > 0 Then
         ' Create the needed cascading update/delete relationship between the parent table and child table
           If mnuRulesEnforce.Checked Then
              If mnuRulesCascadeUpdates.Checked Then
                 If mnuRulesCascadeDeletes.Checked Then
                    Set relNew = db.CreateRelation(sParentTable & "_" & sTableName, sParentTable, sTableName, dbRelationUpdateCascade + dbRelationDeleteCascade)
                 Else
                    Set relNew = db.CreateRelation(sParentTable & "_" & sTableName, sParentTable, sTableName, dbRelationUpdateCascade)
                 End If
              Else
                 If mnuRulesCascadeDeletes.Checked Then
                    Set relNew = db.CreateRelation(sParentTable & "_" & sTableName, sParentTable, sTableName, dbRelationDeleteCascade)
                 Else
                    Set relNew = db.CreateRelation(sParentTable & "_" & sTableName, sParentTable, sTableName)
                 End If
              End If
              relNew.Fields.Append relNew.CreateField(sParentTable & "ID")
              relNew.Fields(sParentTable & "ID").ForeignName = sParentTable & "ID"
              db.Relations.Append relNew
              Set relNew = Nothing
           End If
        End If
                
        Set tdfNew = Nothing
    db.Close
    
    PopulateTree

EH_frmODBCClassGen_AddTable_Continue:
    Exit Sub

EH_frmODBCClassGen_AddTable:
    MsgBox "Error in SliceAndDice.frmODBCClassGen_AddTable" & Chr(13) & Chr(13) & Chr(9) & Err.Description
    Resume EH_frmODBCClassGen_AddTable_Continue
    
    Resume
End Sub

Public Property Get Canceled() As Boolean
    Canceled = m_bCanceled
End Property

Public Property Get DBName() As String
    DBName = sGetToken(sGetToken(m_sClassDatabaseName, lTokenCount(m_sClassDatabaseName, "\"), "\"), 1, ".mdb")
End Property

Public Property Get DBPathAndFilename() As String
    DBPathAndFilename = m_sClassDatabaseName
End Property


Public Property Get GenerateBranch() As Boolean
    GenerateBranch = m_bGenerateBranch
End Property

Public Property Get GenerateDatabase() As Boolean
    GenerateDatabase = m_bGenerateDatabase
End Property

Private Sub LoadCategories()
On Error Resume Next
    Parent.SliceAndDice.Categorys.FillList cboDataLibraryType, 1
    cboDataLibraryType.ListIndex = FindListIndex(cboDataLibraryType, GetSetting(App.ProductName, "Last", "DPCCG Code to generate", "RDO Persisted"))
End Sub

Public Sub PopulateTree()
    Dim db       As Database
    Dim CurTable As TableDef
    Dim nodX     As Node

On Error Resume Next
   'm_sClassDatabaseName = sGetToken(sGetToken(m_sClassDatabaseName, lTokenCount(m_sClassDatabaseName, "\"), "\"), 1, ".mdb")

    Set db = OpenDatabase(m_sClassDatabaseName, m_sClassDatabaseOptions)
        lvwFields.ListItems.Clear
        With tvwTables.Nodes
             .Clear
             Set nodX = .Add(, , "Root", sGetToken(sGetToken(m_sClassDatabaseName, lTokenCount(m_sClassDatabaseName, "\"), "\"), 1, ".mdb"), "Database", "Database")
             nodX.ExpandedImage = "Database"
             nodX.Expanded = True

             For Each CurTable In db.TableDefs
                 If Left(CurTable.Name, 4) <> "MSys" Then
                    Set nodX = .Add("Root", tvwChild, CurTable.Name, CurTable.Name, "Table", "Table")
                    nodX.ExpandedImage = "Table"
                    nodX.Expanded = True
                 End If
             Next CurTable

             For Each CurTable In db.TableDefs
                 If Left(CurTable.Name, 4) <> "MSys" Then
                    If Right(CurTable.Fields(1).Name, 2) = "ID" Then
                       Set .Item(CurTable.Name).Parent = .Item(Left(CurTable.Fields(1).Name, Len(CurTable.Fields(1).Name) - 2))
                    End If
                 End If
             Next CurTable

        End With
    db.Close
    Set nodX = Nothing
End Sub

Public Function sFieldType(iFieldType As Long) As String
       Select Case iFieldType
              Case dbBigInt:        sFieldType = "Big Integer"
              Case dbBinary:        sFieldType = "Binary"
              Case dbBoolean:       sFieldType = "Boolean"
              Case dbByte:          sFieldType = "Byte"
              Case dbChar:          sFieldType = "Char"
              Case dbCurrency:      sFieldType = "Currency"
              Case dbDate:          sFieldType = "Date / Time"
              Case dbDecimal:       sFieldType = "Decimal"
              Case dbDouble:        sFieldType = "Double"
              Case dbFloat:         sFieldType = "Float"
              Case dbGUID:          sFieldType = "Guid"
              Case dbInteger:       sFieldType = "Integer"
              Case dbLong:          sFieldType = "Long"
              Case dbLongBinary:    sFieldType = "Long Binary (OLE Object)"
              Case dbMemo:          sFieldType = "Memo"
              Case dbNumeric:       sFieldType = "Numeric"
              Case dbSingle:        sFieldType = "Single"
              Case dbText:          sFieldType = "Text"
              Case dbTime:          sFieldType = "Time"
              Case dbTimeStamp:     sFieldType = "Time Stamp"
              Case dbVarBinary:     sFieldType = "VarBinary"
        End Select
End Function

Private Sub cmdAddCategory_Click()
    Dim sCategoryChosen As String
    With Parent.SliceAndDice.Categorys
         sCategoryChosen = .Choose(0)
         If Len(sCategoryChosen) > 0 Then
            .Item(sCategoryChosen).CategoryType = 1
            Parent.SliceAndDice.Save
         End If
    End With
    LoadCategories
End Sub

Private Sub cmdDeleteCategory_Click()
    Dim sCategoryChosen As String
    With Parent.SliceAndDice.Categorys
         sCategoryChosen = .Choose(1)
         If Len(sCategoryChosen) > 0 Then
            .Item(sCategoryChosen).CategoryType = 0
            Parent.SliceAndDice.Save
         End If
    End With
    LoadCategories
End Sub


Private Sub Form_Resize()
On Error Resume Next
    lvwFields.Width = Me.ScaleWidth - lvwFields.Left - 100
    lvwFields.Height = Me.ScaleHeight - lvwFields.Top - 100
    tvwTables.Height = Me.ScaleHeight - tvwTables.Top - 100
End Sub

Private Sub lvwFields_DblClick()
    mnuFieldNew_Click
End Sub

Private Sub lvwFields_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       mnuFieldNew_Click
    End If
End Sub

Private Sub mnuFileExit_Click()
    m_bCanceled = True
    lvwFields.ListItems.Clear
    Hide
End Sub

Private Sub mnuFileNew_Click()
    Dim sDatabasePath    As String
    Dim sNewDatabaseName As String
    Dim db               As Database

    sDatabasePath = Trim(BrowseForFolder(hWnd, "Where should database go ?"))
    If Len(sDatabasePath) = 0 Then Exit Sub

    sNewDatabaseName = Trim(InputBox("What should the name of the new database be ?", "CREATE BLANK DATABASE"))
    If Len(sNewDatabaseName) = 0 Then Exit Sub

    If Right(sDatabasePath, 1) <> "\" Then sDatabasePath = sDatabasePath & "\"
    If Right(LCase(sNewDatabaseName), 4) <> ".mdb" Then sNewDatabaseName = sDatabasePath & sNewDatabaseName & ".mdb"

On Error Resume Next
    Set db = CreateDatabase(sNewDatabaseName, dbLangGeneral, dbVersion30)
    db.Close

    m_sClassDatabaseName = sNewDatabaseName
    PopulateTree
End Sub

Private Sub mnuFileOpen2_Click()
    m_sClassDatabaseName = ""
    m_sClassDatabaseOptions = "ODBC;"
    PopulateTree
End Sub

Private Sub mnuGenerateClass_Click()
    If lvwFields.ListItems.Count = 0 Then
       MsgBox "Please select a table first."
       Exit Sub
    End If

    m_bCanceled = False
    m_bGenerateBranch = False
    m_bGenerateDatabase = False
    Hide
End Sub

Private Sub mnuGenerateEnterBranch_Click()
    If lvwFields.ListItems.Count = 0 Then
       MsgBox "Please select a branch first."
       Exit Sub
    End If

    m_bCanceled = False
    m_bGenerateBranch = True
    m_bGenerateDatabase = False
    Hide
End Sub


Private Sub mnuGenerateEntireDatabase_Click()
    If tvwTables.Nodes.Count = 0 Then
       MsgBox "Please select a database first."
       Exit Sub
    End If

    tvwTables.Nodes("Root").Selected = True
    lvwFields.ListItems.Clear

    m_bCanceled = False
    m_bGenerateBranch = False
    m_bGenerateDatabase = True
    Hide
End Sub

Public Sub mnuFileOpen_Click()
    m_sClassDatabaseName = Parent.sChooseDatabase()
    If Len(m_sClassDatabaseName) > 0 Then
       PopulateTree
    End If
End Sub

Private Sub Form_Load()
   'mnuFileOpen_Click
   
   LoadCategories
   
   mnuRulesAutoAddKey.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "AutoAddKey", True)
   mnuRulesAutoAddDateModified.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "AutoAddDateModified", True)
   mnuRulesAutoAddDateCreated.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "AutoAddCreated", True)

   mnuRulesEnforce.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "Enforce", True)
   mnuRulesCascadeUpdates.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "CascadeUpdates", True)
   mnuRulesCascadeDeletes.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "CascadeDeletes", True)

   mnuRulesUseAutoNumber.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "UseAutoNumber", True)

   If mnuRulesEnforce.Checked Then
      mnuRulesCascadeUpdates.Enabled = True
      mnuRulesCascadeDeletes.Enabled = True
   Else
      mnuRulesCascadeUpdates.Enabled = False
      mnuRulesCascadeDeletes.Enabled = False
   End If

End Sub


Private Sub Form_Terminate()
    Set Parent = Nothing
End Sub


Private Sub Label2_Click()

End Sub


Private Sub lvwFields_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    Dim ItemClicked As ListItem

    If Button = vbRightButton Then
       Set ItemClicked = lvwFields.HitTest(x, y)
       If Not ItemClicked Is Nothing Then
          ItemClicked.Selected = True
          PopupMenu mnuField
       End If
    End If
End Sub


Private Sub mnuFieldDelete_Click()
    Dim db           As Database
    Dim sParentTable As String
    Dim sTable       As String
    Dim sField       As String

On Error Resume Next
    If tvwTables.SelectedItem.Parent.Key = m_sClassDatabaseName Then
       sParentTable = ""
    Else
       sParentTable = tvwTables.SelectedItem.Parent.Key
    End If
    sTable = tvwTables.SelectedItem.Key
    sField = lvwFields.SelectedItem.Key

    Select Case sField
           Case sTable & "ID", sTable & "Name", "DateCreated", "DateModified", sParentTable & "ID"
                MsgBox "Can't delete that field (required for correct object/database operation), sorry."
                Exit Sub
    End Select

    Set db = OpenDatabase(m_sClassDatabaseName, "ODBC;")
        db.TableDefs(sTable).Fields.Delete sField
    db.Close

    tvwTables_NodeClick tvwTables.SelectedItem
End Sub

Private Sub mnuFieldNew_Click()
    Dim db         As Database
    Dim CurTable   As TableDef
    Dim CurItem    As ListItem
    Dim NewField   As New frmODBCFieldType

mnuFieldNew_Click_TryAgain:
    With NewField
On Error Resume Next
         Err.Clear
         .Show vbModal, Me
         If Err.Number <> 0 Then Exit Sub
         If .Canceled = True Then Exit Sub

         For Each CurItem In lvwFields.ListItems
             If UCase(CurItem.Text) = UCase(.FieldName) Then
                MsgBox "That field already exists in this table... Try again."
                GoTo mnuFieldNew_Click_TryAgain
             End If
         Next CurItem

On Error Resume Next
         Set db = OpenDatabase(m_sClassDatabaseName, "ODBC;")
             Set CurTable = db.TableDefs(tvwTables.SelectedItem.Key)
             If .dbFieldType = dbText Then
                CurTable.Fields.Append CurTable.CreateField(.FieldName, .dbFieldType, .Length)
             Else
                CurTable.Fields.Append CurTable.CreateField(.FieldName, .dbFieldType)
             End If
         db.Close

         tvwTables_NodeClick tvwTables.SelectedItem
    End With
End Sub

Private Sub mnuRulesAutoAddDateCreated_Click()
    mnuRulesAutoAddDateCreated.Checked = Not mnuRulesAutoAddDateCreated.Checked
    SaveSetting App.ProductName, "ODBC Class Gen", "AutoAddDateCreated", mnuRulesAutoAddDateCreated.Checked
End Sub

Private Sub mnuRulesAutoAddDateModified_Click()
    mnuRulesAutoAddDateModified.Checked = Not mnuRulesAutoAddDateModified.Checked
    SaveSetting App.ProductName, "ODBC Class Gen", "AutoAddDateModified", mnuRulesAutoAddDateModified.Checked
End Sub


Private Sub mnuRulesAutoAddKey_Click()
    mnuRulesAutoAddKey.Checked = Not mnuRulesAutoAddKey.Checked
    SaveSetting App.ProductName, "ODBC Class Gen", "AutoAddKey", mnuRulesAutoAddKey.Checked
End Sub

Private Sub mnuRulesCascadeDeletes_Click()
    mnuRulesCascadeDeletes.Checked = Not mnuRulesCascadeDeletes.Checked
    SaveSetting App.ProductName, "ODBC Class Gen", "CascadeDeletes", mnuRulesCascadeDeletes.Checked
End Sub


Private Sub mnuRulesCascadeUpdates_Click()
    mnuRulesCascadeUpdates.Checked = Not mnuRulesCascadeUpdates.Checked
    SaveSetting App.ProductName, "ODBC Class Gen", "CascadeUpdates", mnuRulesCascadeUpdates.Checked
End Sub

Private Sub mnuRulesEnforce_Click()
    mnuRulesEnforce.Checked = Not mnuRulesEnforce.Checked
    SaveSetting App.ProductName, "ODBC Class Gen", "Enforce", mnuRulesEnforce.Checked
    If mnuRulesEnforce.Checked Then
       mnuRulesCascadeUpdates.Enabled = True
       mnuRulesCascadeDeletes.Enabled = True
    Else
       mnuRulesCascadeUpdates.Enabled = False
       mnuRulesCascadeDeletes.Enabled = False
    End If
End Sub

Private Sub mnuRulesUseAutoNumber_Click()
    mnuRulesUseAutoNumber.Checked = Not mnuRulesUseAutoNumber.Checked
    SaveSetting App.ProductName, "ODBC Class Gen", "UseAutoNumber", mnuRulesUseAutoNumber.Checked
End Sub

Private Sub mnuTableDelete_Click()
    Dim db As Database

On Error Resume Next
    If bUserSure("This will PERMANENTLY remove the table selected." & gsEolTab & "ARE YOU ABSOLUTELY SURE ?") Then
       Set db = OpenDatabase(m_sClassDatabaseName, "ODBC;")
           db.TableDefs.Delete tvwTables.SelectedItem.Key
       db.Close
       PopulateTree
    End If
End Sub

Private Sub mnuTableNew_Click()
    Dim sTable As String

mnuTableNew_Click_TryAgain:
    sTable = sReplace(Trim(InputBox("What should the name of the new table be ?" & gsEolTab & "Note: Table names should be singular, such as:" & gsEolTab & "Book, Publisher, etc.")), " ", "")
    If Len(sTable) = 0 Then Exit Sub
    
    If Right(sTable, 1) = "s" Then
       MsgBox "Table names MUST be singular."
       GoTo mnuTableNew_Click_TryAgain
    End If

    If tvwTables.SelectedItem.Text <> DBName Then
       AddTable sTable, tvwTables.SelectedItem.Text
    Else
       AddTable sTable, ""
    End If
    
    PopulateTree
End Sub

Private Sub mnuX_Click()
    mnuFileExit_Click
End Sub

Private Sub tvwTables_DblClick()
    mnuTableNew_Click
End Sub

Private Sub tvwTables_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       mnuTableNew_Click
    End If
End Sub

Private Sub tvwTables_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim NodeClicked As Node

    If Button = vbRightButton Then
       Set NodeClicked = tvwTables.HitTest(x, y)
       If Not NodeClicked Is Nothing Then
          NodeClicked.Selected = True
          PopupMenu mnuTable
       End If
    End If
End Sub

Public Sub tvwTables_NodeClick(ByVal Node As ComctlLib.Node)
    Dim db       As Database
    Dim CurTable As TableDef
    Dim CurField As Field
    Dim nodX     As Node
    Dim litX     As ListItem
    Dim sIcon    As String

On Error Resume Next
    Set db = OpenDatabase(m_sClassDatabaseName, "ODBC;")
        With lvwFields.ListItems
             .Clear
             With lvwFields.ColumnHeaders
                  .Clear
                  .Add , "Field Name", "Field Name", 2000
                  .Add , "Field Type", "Type", 1000
                  .Add , "Field Length", "Length", 500
             End With
             lvwFields.View = lvwReport
             Set CurTable = db.TableDefs(tvwTables.SelectedItem.Key)
             For Each CurField In CurTable.Fields
                 If Right(CurField.Name, 2) = "ID" Then
                    If Left(CurField.Name, Len(CurField.Name) - 2) = CurTable.Name Then
                       sIcon = "Key"
                    Else
                       sIcon = "ID"
                    End If
                 ElseIf CurField.Name = "DateCreated" Or CurField.Name = "DateModified" Then
                    sIcon = "Date"
                 Else
                    sIcon = "Field"
                 End If
                 Set litX = .Add(, CurField.Name, CurField.Name, sIcon, sIcon)
                 litX.SubItems(1) = sFieldType(CurField.Type)
                 litX.SubItems(2) = CurField.Size
             Next CurField
        End With
    db.Close
    Set nodX = Nothing
End Sub

