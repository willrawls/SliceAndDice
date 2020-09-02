VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
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
   Tag             =   "ForeVB DB=S:\Projects - Firm Solutions\SliceAndDice\SliceAndDice.dba"
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

Private Favorites As SandySupport.CAssocArray
Private FavoriteCount As Long
Private RetrievingAFavoriteNow As Boolean

Public Parent As SandySupport.ISandyWindowMain

Private NodeDragged As Node

Implements SandySupport.ISandyWindowGen
Public Sub GenerateChildren(ByRef asaPass As SandySupport.CAssocArray, sDataLibraryType As String, nodChild As Node)
    Dim CurChild As Node
    Set CurChild = nodChild
    Do Until CurChild Is Nothing
       CurChild.Selected = True
       tvwTables_NodeClick CurChild

       If Not CurChild.Parent.Parent Is Nothing Then
          asaPass("Parent Table Name") = CurChild.Parent.Text
       Else
          asaPass("Parent Table Name") = "Root"
       End If

       GenerateClass asaPass, sDataLibraryType, tvwTables, lvwFields
       If Parent.Parent.InsertionCancelled Then Exit Sub

       If Not CurChild.Child Is Nothing Then
          GenerateChildren asaPass, sDataLibraryType, CurChild.Child
          If Parent.Parent.InsertionCancelled Then Exit Sub
       End If
       Set CurChild = CurChild.Next
    Loop
End Sub


Public Sub GenerateClass(ByRef asaPass As SandySupport.CAssocArray, sDataLibraryType As String, tvwTables As TreeView, lvwFields As ListView)
    Dim asaV As SandySupport.CAssocArray
    Dim CurListItem As ListItem
    Dim CurChild As Node

    Dim sTableName As String
    Dim sFieldType As String
   'Dim bParentAdded As Boolean
    Dim sDBPCType As String
    Dim sParentName As String
    Dim sClassToCollect As String
    Dim sChildTableName As String
    Dim bSingularCollects As Boolean
    Dim sCategoryName As String

On Error Resume Next

    sCategoryName = sGetToken(sDataLibraryType, 1, " - ")

  ' Determine if the singular member of the collection will be collecting anything
    bSingularCollects = (Not tvwTables.SelectedItem.Child Is Nothing)
    If bSingularCollects Then
       sClassToCollect = tvwTables.SelectedItem.Child.Text & "s"
       sChildTableName = tvwTables.SelectedItem.Child.Text
    Else
       sClassToCollect = vbNullString
       sChildTableName = tvwTables.SelectedItem.Child.Text
    End If

  ' Generate the Collection Class
    sTableName = Parent.Parent.sTableToPropertyName(tvwTables.SelectedItem.Text)
   'sTableName = Replace(Replace(Replace(tvwTables.SelectedItem.Text, "_", vbNullString), " ", vbNullString), ".", "__")
    Err.Clear

    If Not asaPass Is Nothing Then
       Set asaV = asaPass
    Else
       Set asaV = CreateObject("SandySupport.CAssocArray")
    End If

    If Right$(lvwFields.ListItems(2).Key, 2) = "ID" Then
     ' Collection has a parent object
       sParentName = lvwFields.ListItems(2).Key
       asaV("Parent AutoNumber Field Name") = sParentName
       asaV("Parent AutoNumber Property Name") = Parent.Parent.sTableToPropertyName(sParentName)
      'asaV("Parent AutoNumber Property Name") = Replace(Replace(Replace(sParentName, " ", "_"), "*", "_"), "-", "_")
       sDBPCType = vbNullString
    ElseIf Not tvwTables.SelectedItem.Parent Is Nothing Then
       If StrComp(tvwTables.SelectedItem.Parent.Key, "ODBC") = 0 Or StrComp(tvwTables.SelectedItem.Parent.Key, "Root") = 0 Then
        ' Collection DOESN'T have a parent object
          asaV("Parent AutoNumber Field Name") = vbNullString
          asaV("Parent AutoNumber Property Name") = vbNullString
          sParentName = vbNullString
          sDBPCType = ", No Parent"
       Else
        ' Collection has a parent object
          sParentName = tvwTables.SelectedItem.Text
          asaV("Parent AutoNumber Field Name") = sParentName
          asaV("Parent AutoNumber Property Name") = Parent.Parent.sTableToPropertyName(sParentName)
         'asaV("Parent AutoNumber Property Name") = Replace(Replace(Replace(Replace(sParentName, " ", "_"), "*", "_"), "-", "_"), ".", "__")
          sDBPCType = vbNullString
       End If
    Else
     ' Collection DOESN'T have a parent object
       asaV("Parent AutoNumber Field Name") = vbNullString
       asaV("Parent AutoNumber Property Name") = vbNullString
       sParentName = vbNullString
       sDBPCType = ", No Parent"
    End If

   'asaV("Collection Member Subcollection Property Name") = sClassToCollect
    asaV("Property Name") = sClassToCollect
    asaV("Child Table Name") = sChildTableName
    asaV("Singular Property Name") = sChildTableName

   'asaV("Primary AutoNumber Field for Collection Member") = lvwFields.ListItems(1).Key
    asaV("AutoNumber Field Name") = lvwFields.ListItems(1).Key
    asaV("AutoNumber Property Name") = Parent.Parent.sTableToPropertyName(lvwFields.ListItems(1).Key)
   'asaV("AutoNumber Property Name") = Replace(Replace(Replace(lvwFields.ListItems(1).Key, " ", "_"), "*", "_"), "-", "_")
    
   'asaV("Table that stores this collection") = sTableName
    asaV("Pure Table Name") = tvwTables.SelectedItem.Text
    asaV("Table Name") = sTableName
   'asaV("Object Name of Collection Member") = sTableName
    asaV("Object Name") = sTableName
    
    asaV("Spaced Table Name") = sInsertSpaces(sTableName)
    asaV("Spaced Object Name") = sInsertSpaces(sTableName)
   'asaV("Label Name of Collection Member") = sInsertSpaces(sTableName)
    asaV("Label Name") = sInsertSpaces(sTableName)

    If Len(sDBPCType) = 0 Then
     ' Collection has a parent object
      'asaV("Field to use as Key") = lvwFields.ListItems(3).Key
       asaV("Key Field Name") = lvwFields.ListItems(3).Key
       asaV("Key Property Name") = Parent.Parent.sTableToPropertyName(lvwFields.ListItems(3).Key)
      'asaV("Key Property Name") = Replace(Replace(Replace(lvwFields.ListItems(3).Key, " ", "_"), "*", "_"), "-", "_")
    Else
     ' Collection DOESN'T have a parent object
      'asaV("Field to use as Key") = lvwFields.ListItems(2).Key
       asaV("Key Field Name") = lvwFields.ListItems(2).Key
       asaV("Key Property Name") = Parent.Parent.sTableToPropertyName(lvwFields.ListItems(2).Key)
      'asaV("Key Property Name") = Replace(Replace(Replace(lvwFields.ListItems(2).Key, " ", "_"), "*", "_"), "-", "_")
    End If

    If bSingularCollects = False Then
       sDBPCType = sDBPCType & ", No Child"
    End If

    If Not Parent.SliceAndDice.Categorys(sCategoryName).Templates("Table - " & tvwTables.SelectedItem.Text) Is Nothing Then
       Parent.DoInsertion asaV, sDataLibraryType & "Table - " & tvwTables.SelectedItem.Text
    End If
    If Parent.Parent.InsertionCancelled Then Exit Sub

    Parent.DoInsertion asaV, sDataLibraryType & "Collection" & sDBPCType
    If Parent.Parent.InsertionCancelled Then Exit Sub
    
  ' Generate the Collection MEMBER Class
   'asaV.Clear
   'asaV("Object Name of Collection Member")=sTableName
   'asaV("Object Name")=sTableName
   'asaV("Table Name")=sTableName

   'asaV("Label Name of Collection Member")=sInsertSpaces(sTableName)
   'asaV("Label Name")=sInsertSpaces(sTableName)
   'asaV("Spaced Table Name")=sInsertSpaces(sTableName)

    If bSingularCollects Then
       If Len(sClassToCollect) = 0 Then sClassToCollect = "SubClass"
      'asaV("Property name of Class to collect") = sClassToCollect
       asaV("Class to collect") = sClassToCollect
       asaV("Property Name") = sClassToCollect
      'asaV("Collection Member Subcollection Property Name") = sClassToCollect
       Parent.DoInsertion asaV, sDataLibraryType & "Collection Member"
       If Parent.Parent.InsertionCancelled Then Exit Sub
    Else
       Parent.DoInsertion asaV, sDataLibraryType & "Collection Member, Terminal"
       If Parent.Parent.InsertionCancelled Then Exit Sub
    End If

    For Each CurListItem In lvwFields.ListItems
       'asaV("Field Name of Property") = CurListItem.Key
        asaV("Property Name") = Parent.Parent.sTableToPropertyName(CurListItem.Key)
       'asaV("Property Name") = Replace(Replace(Replace(CurListItem.Key, " ", "_"), "*", "_"), "-", "_")
       'asaV("Pure Field Name") = CurListItem.Key
        asaV("Spaced Field Name") = sInsertSpaces(CurListItem.Key)
        asaV("Field Name") = CurListItem.Key
        asaV("Table Name") = sTableName
        asaV("Spaced Table Name") = sInsertSpaces(sTableName)
        asaV("Property Type") = CurListItem.SubItems(1)
        asaV("Property Size") = CurListItem.SubItems(2)
        asaV("Property Length") = CurListItem.SubItems(2)
        asaV("Field Type") = CurListItem.SubItems(1)
        asaV("Field Length") = CurListItem.SubItems(2)

        If Left$(CurListItem.Key, 2) = "s_" Then
         ' A field that shouldn't be messed with
        ElseIf Right$(CurListItem.Key, 2) = "ID" Then
           Parent.DoInsertion asaV, sDataLibraryType & "Property - " & Parent.sPropertyType(CurListItem.SubItems(1))
           If Parent.Parent.InsertionCancelled Then Exit Sub

           If StrComp(CurListItem.Key, sTableName & "ID") <> 0 Then
              Parent.DoInsertion asaV, sDataLibraryType & "Property - 3D Link"
              If Parent.Parent.InsertionCancelled Then Exit Sub
           End If
        Else
           Parent.DoInsertion asaV, sDataLibraryType & "Property - " & Parent.sPropertyType(CurListItem.SubItems(1))
           If Parent.Parent.InsertionCancelled Then Exit Sub
        End If
    Next CurListItem

    If Not Parent.SliceAndDice.Categorys(sCategoryName).Templates("Table - " & tvwTables.SelectedItem.Text & " - Finalize") Is Nothing Then
       Parent.DoInsertion asaV, sDataLibraryType & "Table - " & tvwTables.SelectedItem.Text & " - Finalize"
    ElseIf Not Parent.SliceAndDice.Categorys(sCategoryName).Templates("Table - " & tvwTables.SelectedItem.Text & ", Finalize") Is Nothing Then
       Parent.DoInsertion asaV, sDataLibraryType & "Table - " & tvwTables.SelectedItem.Text & ", Finalize"
    End If
    If Parent.Parent.InsertionCancelled Then Exit Sub

    If bSingularCollects Then
     ' Generate any additional Collection MEMBER sub-collection linkages
       sClassToCollect = tvwTables.SelectedItem.Child.Key
       Set CurChild = tvwTables.SelectedItem.Child
       Do Until CurChild Is Nothing
          If sClassToCollect <> CurChild.Key Then
             asaV("Singular Property Name") = Parent.Parent.sTableToPropertyName(CurChild.Key)
            'asaV("Singular Property Name") = Replace(Replace(CurChild.Key, "_", vbNullString), " ", vbNullString)
             asaV("Property Name") = asaV("Singular Property name") & "s"
             asaV("Pure Child Table Name") = CurChild.Key
             asaV("Child Table Name") = asaV("Singular Property Name")
             asaV("Pure Table Name") = tvwTables.SelectedItem.Text
             asaV("Table Name") = sTableName
             asaV("Spaced Table Name") = sInsertSpaces(sTableName)

             Parent.DoInsertion asaV, sDataLibraryType & "Collection Member - New Subcollection"
             If Parent.Parent.InsertionCancelled Then Exit Sub
          End If
          Set CurChild = CurChild.Next
       Loop
    End If
End Sub


Public Sub TriggerClassGeneration()
   'If gbProcessing Then Exit Sub

    Dim asaV As New CAssocArray
    Dim CurChild As Node

    If Canceled = False Then
      'masaMisc.Clear                                                                 ' Clean out the assoc used for 'session' long inserts
       If GenerateDatabase = False Then
        ' Generate a class for the currently selected class
          GenerateClass asaV, cboDataLibraryType.Text & " - ", tvwTables, lvwFields
          If GenerateBranch = True Then
           ' Generate a class for each child table and each of its children tables
             If Not tvwTables.SelectedItem.Child Is Nothing Then
                GenerateChildren asaV, cboDataLibraryType.Text & " - ", tvwTables.SelectedItem.Child
             End If
          End If
       Else
        ' Create a wrapper class
         'frmLog.tvwLog.Nodes.Clear
          If Len(DBName) Then
             asaV("DSN") = DBName
             asaV("Database Name") = DBName
             asaV("Database Path") = DBPathAndFilename
             asaV("Spaced Database Name") = sInsertSpaces(DBName)
          Else
             asaV("DSN") = sGetToken(sGetToken(ConnectString, 2, "DSN="), 1, ";")
             asaV("DSNConnect") = ConnectString
             asaV("ConnectString") = ConnectString
             asaV("Connect") = ConnectString
             asaV("Database Name") = ODBCDatabaseName
             asaV("Database Path") = "xxx No database path available. ODBC generation xxx"
             asaV("Spaced Database Name") = sInsertSpaces(ODBCDatabaseName)
          End If
          Parent.DoInsertion asaV, cboDataLibraryType.Text & " - " & "Settings"
          If Parent.Parent.InsertionCancelled Then Exit Sub
          Parent.DoInsertion asaV, cboDataLibraryType.Text & " - " & "Routines"
          If Parent.Parent.InsertionCancelled Then Exit Sub
          Parent.DoInsertion asaV, cboDataLibraryType.Text & " - " & "Wrapper class"
          If Parent.Parent.InsertionCancelled Then Exit Sub

        ' In the wrapper class, add collections for each top level DB table
          Set CurChild = tvwTables.SelectedItem.Child
          Do Until CurChild Is Nothing
            'asaV.Clear
             asaV("Pure Table Name") = CurChild.Text
             asaV("Table Name") = CurChild.Text
             asaV("Property Name") = Parent.Parent.sTableToPropertyName(CurChild.Text) & "s"
            'asaV("Plural Table Name") = asaV("Property Name")
             asaV("Spaced Table Name") = sInsertSpaces(Parent.Parent.sTableToPropertyName(CurChild.Text))
             Parent.DoInsertion asaV, cboDataLibraryType.Text & " - " & "Wrapper class - Add collection"
             If Parent.Parent.InsertionCancelled Then Exit Sub
             Set CurChild = CurChild.Next
          Loop

        ' Insert extra functions
         'asaV.Clear
          If Len(DBName) Then
             asaV("DSN") = DBName
             asaV("Database Name") = DBName
             asaV("Database Path") = DBPathAndFilename
             asaV("Spaced Database Name") = sInsertSpaces(DBName)
          Else
             asaV("DSN") = ConnectString
             asaV("Database Name") = ODBCDatabaseName
             asaV("Database Path") = "xxx No database path available. ODBC generation xxx"
             asaV("Spaced Database Name") = sInsertSpaces(ODBCDatabaseName)
          End If

        ' Walk down the tree creating everything
          GenerateChildren asaV, cboDataLibraryType.Text & " - ", tvwTables.SelectedItem.Child
          
        ' Finalize the insertions
On Error Resume Next
          If Not Parent.SliceAndDice.Categorys(cboDataLibraryType.Text).Templates("Finalize") Is Nothing Then
             Parent.DoInsertion asaV, cboDataLibraryType.Text & " - " & "Finalize"
          End If
          
         'frmLog.Show vbModal
          If Parent.Parent.InsertionCancelled Then
             MsgBox "Processing canceled at user's request."
          Else
             MsgBox "Done generating database.", vbOKOnly, "CODE GENERATION COMPLETE"
          End If
       End If
    End If
End Sub


Private Sub AddFavorite()
    Dim NewKey As String
    If Len(msClassDatabaseName) Then
       Favorites(msClassDatabaseName) = "|||" & sNormalize(TableList)
    Else
       If Len(ODBCTableNamePrefix) Then
          Favorites("ODBC: " & sGetToken(sGetToken(msClassDatabaseOptions, 2, "DSN="), 1, ";") & " (" & ODBCTableNamePrefix & ")" & IIf(Len(TableList), " Limited (" & tvwTables.Nodes.Count - 1 & ")", vbNullString)) = msClassDatabaseOptions & "|" & ODBCTableNamePrefix & "|" & ODBCPassword & "|" & sNormalize(TableList)
       Else
          Favorites("ODBC: " & sGetToken(sGetToken(msClassDatabaseOptions, 2, "DSN="), 1, ";") & IIf(Len(TableList), " Limited (" & tvwTables.Nodes.Count - 1 & ")", vbNullString)) = msClassDatabaseOptions & "|" & ODBCTableNamePrefix & "|" & ODBCPassword & "|" & sNormalize(TableList)
       End If
    End If
    SaveSetting "SliceAndDice", "DB Class Gen", "Favorites", Favorites.All
    UpdateFavorites
End Sub


Public Sub AddTable(sTableName As String, sParentTable As String)
On Error GoTo EH_frmDBClassGen_AddTable
    Dim tdfNew As TableDef
    Dim relNew As Relation
    Dim idxNew As Index

    If IsInODBCDatabaseMode Then
       If Not KeepDatabaseOpen Then Set db = OpenDatabase(msClassDatabaseName, dbDriverComplete, False, msClassDatabaseOptions)
    Else
       Set db = OpenDatabase(msClassDatabaseName, False, False)
    End If
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
    If Not KeepDatabaseOpen Then db.Close
    
    PopulateTree

EH_frmDBClassGen_AddTable_Continue:
    Exit Sub

EH_frmDBClassGen_AddTable:
    MsgBox "Error in SliceAndDice.frmDBClassGen_AddTable" & vbCr & vbCr & vbTab & Err.Description
    Resume EH_frmDBClassGen_AddTable_Continue
    
    Resume
End Sub

Public Property Get Canceled() As Boolean
    Canceled = mbCanceled
End Property

Public Property Get ConnectString() As String
    ConnectString = msClassDatabaseOptions
End Property

Public Property Get DBName() As String
    DBName = sGetToken(sGetToken(msClassDatabaseName, lTokenCount(msClassDatabaseName, "\"), "\"), 1, ".mdb")
End Property

Public Property Get DBPathAndFilename() As String
    DBPathAndFilename = msClassDatabaseName
End Property


Public Property Get ODBCDatabaseName() As String
    ODBCDatabaseName = sGetToken(sGetToken(msClassDatabaseOptions, 2, "DSN="), 1, ";")
End Property

Private Sub RefreshTableList()
    Dim TreeReading As SandySupport.CAssocArray
    Set TreeReading = CreateObject("SandySupport.CAssocArray")
        TreeReading.TreeToAll tvwTables
        TableList = TreeReading.All
    Set TreeReading = Nothing
End Sub

Private Function RemoveNodes(tvwX As TreeView, sNodeImageNameToRemove) As String
    Dim CurrNode As Node
    Dim sNodesLeft As String
RemoveNodes_Restart:
    For Each CurrNode In tvwX.Nodes
        If StrComp(UCase$(CurrNode.Image), UCase$(sNodeImageNameToRemove)) = 0 Then
           tvwX.Nodes.Remove CurrNode.Key
           GoTo RemoveNodes_Restart
        ElseIf Not CurrNode.Parent Is Nothing Then
           sNodesLeft = sNodesLeft & CurrNode.Key & ";"
        End If
    Next CurrNode
    
    For Each CurrNode In tvwX.Nodes
        sNodesLeft = sNodesLeft & CurrNode.Key & ";"
    Next CurrNode
    
    RemoveNodes = sNodesLeft
End Function

Public Sub UpdateFavorites()
On Error Resume Next
    Dim CurrFav As Long
    Dim CurrAssoc As SandySupport.CAssocItem
    
    If FavoriteCount > 0 Then ' Clear out previous entries
       For CurrFav = FavoriteCount To 1 Step -1
           Unload mnuFavorite(CurrFav)
       Next CurrFav
       mnuFavorite(0).Caption = "-Empty-"
       mnuFavorite(0).Enabled = False
       FavoriteCount = 0
    End If
    
    For Each CurrAssoc In Favorites
        If FavoriteCount > 0 Then
           Load mnuFavorite(FavoriteCount)
        End If
        mnuFavorite(FavoriteCount).Caption = CurrAssoc.Key
        mnuFavorite(FavoriteCount).Enabled = True
        FavoriteCount = FavoriteCount + 1
    Next CurrAssoc
End Sub


Public Property Get GenerateBranch() As Boolean
    GenerateBranch = mbGenerateBranch
End Property

Public Property Get GenerateDatabase() As Boolean
    GenerateDatabase = mbGenerateDatabase
End Property

Public Sub RefreshCategories()
On Error Resume Next
    mbLoadingCategories = True
        Parent.SliceAndDice.Categorys.FillList cboDataLibraryType, 1
        cboDataLibraryType.ListIndex = FindListIndex(cboDataLibraryType, GetSetting("SliceAndDice", "DB Class Gen", "Last Category", "RDO Persisted"))
    mbLoadingCategories = False
End Sub

Public Sub PopulateTree()
On Error GoTo EH_PopulateTree
    Dim CurTable As TableDef
    Dim nodX     As Node
    Dim TableListCount As Long
    Dim asaX As SandySupport.CAssocArray
    Dim CurrItem As SandySupport.CAssocItem

'On Error Resume Next
   'msClassDatabaseName = sGetToken(sGetToken(msClassDatabaseName, lTokenCount(msClassDatabaseName, "\"), "\"), 1, ".mdb")

    Screen.MousePointer = vbHourglass
    If Len(msClassDatabaseName) = 0 Then
       If Not KeepDatabaseOpen Then
          If Len(msClassDatabaseOptions) = 0 Then
             msClassDatabaseOptions = "ODBC;"
          End If
          Set db = OpenDatabase(msClassDatabaseName, dbDriverPrompt, False, msClassDatabaseOptions)
          If Not db Is Nothing Then
             IsInODBCDatabaseMode = True
             KeepDatabaseOpen = True

             If Len(ODBCPassword) = 0 Then
                msClassDatabaseOptions = db.Connect
                ODBCPassword = InputBox("What is the database password you just entered ?" & vbCr & vbTab & "NOTE: Entering it again here prevents you having to reenter it throughout this session.", "ENTER ODBC DATABASE PASSWORD")
                If Len(ODBCPassword) Then
                   If Right$(msClassDatabaseOptions, 1) <> ";" Then
                      msClassDatabaseOptions = msClassDatabaseOptions & ";PWD=" & ODBCPassword & ";"
                   Else
                      msClassDatabaseOptions = msClassDatabaseOptions & "PWD=" & ODBCPassword & ";"
                   End If
                End If
             End If
             If Len(ODBCTableNamePrefix) = 0 Then
                If db.TableDefs.Count > 100 Then
                   ODBCTableNamePrefix = InputBox("There are " & db.TableDefs.Count & " tables in this database." & vbCr & "Would you like to limit the tables included by a table name prefix ?" & vbCr & vbTab & "(Leave blank to include all tables)", "LIMIT TABLES BY TABLE NAME PREFIX")
                End If
             End If
             If Not RetrievingAFavoriteNow Then
                AddFavorite
             End If
          End If
       End If
    Else
       msClassDatabaseOptions = vbNullString
       Set db = OpenDatabase(msClassDatabaseName, False, False)
       IsInODBCDatabaseMode = False
       KeepDatabaseOpen = False
       If Not RetrievingAFavoriteNow Then
          AddFavorite
       End If
    End If

        lvwFields.ListItems.Clear
        With tvwTables.Nodes
             .Clear
             If IsInODBCDatabaseMode Then
                Set nodX = .Add(, , "Root", "ODBC", "Database", "Database")
             Else
                Set nodX = .Add(, , "Root", sGetToken(sGetToken(msClassDatabaseName, lTokenCount(msClassDatabaseName, "\"), "\"), 1, ".mdb"), "Database", "Database")
             End If
             nodX.ExpandedImage = "Database"
             nodX.Expanded = True

             If Len(TableList) = 0 Then
                Screen.MousePointer = vbHourglass
                TableListCount = 0
                For Each CurTable In db.TableDefs
                    If Left$(CurTable.Name, 4) <> "MSys" And Left$(CurTable.Name, 4) <> "SYS." And Left$(CurTable.Name, 4) <> "ALL_" Then
                       If Len(ODBCTableNamePrefix) = 0 Or UCase$(Left$(CurTable.Name, Len(ODBCTableNamePrefix))) = UCase$(ODBCTableNamePrefix) Then
                          Set nodX = .Add("Root", tvwChild, CurTable.Name, CurTable.Name, "Table", "Table")
                          nodX.ExpandedImage = "Table"
                          nodX.Expanded = True
                          TableListCount = TableListCount + 1
                       End If
                    End If
                Next CurTable
             ElseIf InStr(TableList, "<ICON>") = 0 And InStr(TableList, "<CHILD>") = 0 And InStr(TableList, "<ENDCHILD>") = 0 Then
                Set asaX = CreateObject("SandySupport.CAssocArray")
                asaX.ItemDelimiter = ";"
                asaX.All = TableList
                For Each CurrItem In asaX
                    Set nodX = .Add("Root", tvwChild, CurrItem.Key, CurrItem.Key, "Table", "Table")
                    nodX.ExpandedImage = "Table"
                    nodX.Expanded = True
                Next CurrItem
                TableListCount = asaX.Count
                asaX.Clear
                Set asaX = Nothing
             Else
                Set asaX = CreateObject("SandySupport.CAssocArray")
                    asaX.All = TableList
                    tvwTables.Nodes.Clear
                    asaX.FillTreeNode tvwTables, Nothing, "Table", True
                Set asaX = Nothing
             End If
             
             If Not mnuRelateOnLoad.Checked Then
                GoTo SKIP_RELATION_STEP
             End If

             If db.TableDefs.Count > 40 Then
                If bUserSure("There are a lot of table to attempt to relate. If there are no Slice and Dice relations, this step can be skipped." & vbCr & vbTab & "Would you like to skip this step ?") Then
                   GoTo SKIP_RELATION_STEP
                End If
             End If
             Screen.MousePointer = vbHourglass
On Error Resume Next
             For Each CurTable In db.TableDefs
                 If Left$(CurTable.Name, 4) <> "MSys" And Left$(CurTable.Name, 4) <> "SYS." And Left$(CurTable.Name, 4) <> "ALL_" Then
                    If Len(ODBCTableNamePrefix) = 0 Or UCase$(Left$(CurTable.Name, Len(ODBCTableNamePrefix))) = UCase$(ODBCTableNamePrefix) Then
                       If CurTable.Fields.Count > 1 Then
                          Screen.MousePointer = vbHourglass
                          If Right$(CurTable.Fields(1).Name, 2) = "ID" Then
                             Set nodX = Nothing
                             Set nodX = .Item(Left$(CurTable.Fields(1).Name, Len(CurTable.Fields(1).Name) - 2))
                             If Not nodX Is Nothing Then
                                Set .Item(CurTable.Name).Parent = nodX
                             End If
                          End If
                       End If
                    End If
                 End If
             Next CurTable
             
SKIP_RELATION_STEP:
        End With
    If Not KeepDatabaseOpen Then db.Close

EH_PopulateTree_Continue:
    Set nodX = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

EH_PopulateTree:
    LogError "frmDBClassGen", "PopulateTree", Err.Number, Err.Description
    Resume EH_PopulateTree_Continue
    
    Resume
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

Private Sub cboDataLibraryType_Click()
    If mbLoadingCategories Then Exit Sub
    SaveSetting "SliceAndDice", "DB Class Gen", "Last Category", cboDataLibraryType.Text
End Sub

Private Sub cmdAddCategory_Click()
    Dim sCategoryChosen As String
    With Parent.SliceAndDice.Categorys
         sCategoryChosen = .Choose(0)
         If Len(sCategoryChosen) > 0 Then
            .Item(sCategoryChosen).CategoryType = 1
            Parent.SliceAndDice.Save
         End If
    End With
    RefreshCategories
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
    RefreshCategories
End Sub


Private Sub Form_Initialize()

    ' LogEvent "frmDBClassGen: Initialize"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
       Cancel = True
    End If
    mnuFileExit_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Not dvwTable.Visible Then
       lvwFields.Move ScaleWidth - fraCategory.Width, lvwFields.Top, fraCategory.Width, ScaleHeight - lvwFields.Top
       tvwTables.Move 0, 60, ScaleWidth - fraCategory.Width, ScaleHeight
       fraCategory.Move lvwFields.Left, 0
    Else
       lvwFields.Move ScaleWidth - fraCategory.Width, lvwFields.Top, fraCategory.Width, ScaleHeight - fraCategory.Height - dvwTable.Height
       tvwTables.Move 0, 60, ScaleWidth - fraCategory.Width, ScaleHeight - dvwTable.Height
       fraCategory.Move lvwFields.Left, 0
       dvwTable.Move 0, lvwFields.Top + lvwFields.Height, ScaleWidth
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPosition Me
End Sub

Private Sub ISandyWindowGen_AddTable(sTableName As String, sParentTable As String)
    AddTable sTableName, sParentTable
End Sub

Private Property Get ISandyWindowGen_Canceled() As Boolean
    ISandyWindowGen_Canceled = mbCanceled
End Property


Private Property Get ISandyWindowGen_ConnectString() As String
    ISandyWindowGen_ConnectString = msClassDatabaseOptions
End Property


Private Property Get ISandyWindowGen_DBName() As String
    ISandyWindowGen_DBName = DBName
End Property


Private Property Get ISandyWindowGen_DBPathAndFilename() As String
    ISandyWindowGen_DBPathAndFilename = DBPathAndFilename
End Property


Private Property Get ISandyWindowGen_GenerateBranch() As Boolean
    ISandyWindowGen_GenerateBranch = mbGenerateBranch
End Property


Private Sub ISandyWindowGen_GenerateChildren(asaPass As SandySupport.CAssocArray, sDataLibraryType As String, nodChild As Object)
    GenerateChildren asaPass, sDataLibraryType, nodChild
End Sub


Private Sub ISandyWindowGen_GenerateClass(asaPass As SandySupport.CAssocArray, sDataLibraryType As String, tvwTables As Object, lvwFields As Object)
    GenerateClass asaPass, sDataLibraryType, tvwTables, lvwFields
End Sub


Private Property Get ISandyWindowGen_GenerateDatabase() As Boolean
    ISandyWindowGen_GenerateDatabase = GenerateDatabase
End Property


Private Sub ISandyWindowGen_Hide()
    Hide
End Sub


Private Sub ISandyWindowGen_NodeClick(ByRef Node As MSComctlLib.INode)
    tvwTables_NodeClick Node
End Sub

Private Property Get ISandyWindowGen_ODBCDatabaseName() As String
    ISandyWindowGen_ODBCDatabaseName = ODBCDatabaseName
End Property

Private Sub ISandyWindowGen_OpenFile()
    mnuFileOpen_Click
End Sub


Private Property Set ISandyWindowGen_Parent(ByVal RHS As SandySupport.ISandyWindowMain)
    Set Parent = RHS
End Property


Private Property Get ISandyWindowGen_Parent() As SandySupport.ISandyWindowMain
    Set ISandyWindowGen_Parent = Parent
End Property

Private Sub ISandyWindowGen_PopulateTree()
    PopulateTree
End Sub

Private Sub ISandyWindowGen_RefreshCategories()
    RefreshCategories
End Sub


Private Sub ISandyWindowGen_SetColors(ByVal ForeColor As Long, ByVal BackColor As Long)
    lvwFields.BackColor = BackColor
    lvwFields.ForeColor = ForeColor
    dvwTable.BackColor = BackColor
    dvwTable.ForeColor = ForeColor
End Sub

Private Function ISandyWindowGen_sFieldType(iFieldType As Long) As String
    ISandyWindowGen_sFieldType = sFieldType(iFieldType)
End Function


Private Sub ISandyWindowGen_Show(Optional ByVal ModalSetting As Integer, Optional ParentWindow As Object)
    If IsMissing(ParentWindow) Then
       Show ModalSetting
    Else
       Show ModalSetting, ParentWindow
    End If
End Sub


Private Sub ISandyWindowGen_TriggerClassGeneration()
    TriggerClassGeneration
End Sub


Private Sub ISandyWindowGen_UpdateFavorites()
    UpdateFavorites
End Sub

Private Sub ISandyWindowGen_ZOrder()
    ZOrder
End Sub


Private Sub lvwFields_DblClick()
    mnuFieldNew_Click
End Sub

Private Sub lvwFields_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       mnuFieldNew_Click
    End If
End Sub

Private Sub mnuFavorite_Click(Index As Integer)
    Screen.MousePointer = vbDefault
    With Favorites(mnuFavorite(Index).Caption)
         If Len(sGetToken(.Value, 1, "|")) = 0 Then
            msClassDatabaseName = sGetToken(.Key, 1, "|")
            msClassDatabaseOptions = vbNullString
            ODBCTableNamePrefix = vbNullString
            ODBCPassword = vbNullString
            TableList = sDenormalize(sAfter(.Value, 3, "|"))
         Else
            msClassDatabaseName = vbNullString
            msClassDatabaseOptions = sGetToken(.Value, 1, "|")
            ODBCTableNamePrefix = sGetToken(.Value, 2, "|")
            ODBCPassword = sGetToken(.Value, 3, "|")
            TableList = sDenormalize(sAfter(.Value, 3, "|"))
         End If
         IsInODBCDatabaseMode = False
         KeepDatabaseOpen = False
         RetrievingAFavoriteNow = True
         PopulateTree
         RetrievingAFavoriteNow = False
    End With
End Sub

Private Sub mnuFavRemoveAll_Click()
    If bUserSure() Then
       Favorites.All = vbNullString
       SaveSetting "SliceAndDice", "DB Class Gen", "Favorites", vbNullString
       UpdateFavorites
    End If
End Sub


Private Sub mnuFileExit_Click()
    mbCanceled = True
    lvwFields.ListItems.Clear
    SaveFormPosition Me
    Hide
End Sub

Private Sub mnuFileNew_Click()
    Dim sDatabasePath    As String
    Dim sNewDatabaseName As String

    sDatabasePath = Trim$(BrowseForFolder(Me.hWnd, "Where should database go ?"))
    If Len(sDatabasePath) = 0 Then Exit Sub

    sNewDatabaseName = Trim$(InputBox("What should the name of the new database be ?", "CREATE BLANK DATABASE"))
    If Len(sNewDatabaseName) = 0 Then Exit Sub

    If Right$(sDatabasePath, 1) <> "\" Then sDatabasePath = sDatabasePath & "\"
    If Right$(LCase$(sNewDatabaseName), 4) <> ".mdb" Then sNewDatabaseName = sDatabasePath & sNewDatabaseName & ".mdb"

On Error Resume Next
    Set db = CreateDatabase(sNewDatabaseName, dbLangGeneral, dbVersion30)
    db.Close

    msClassDatabaseName = sNewDatabaseName
    msClassDatabaseOptions = vbNullString
    ODBCTableNamePrefix = vbNullString
    ODBCPassword = vbNullString
    IsInODBCDatabaseMode = False
    KeepDatabaseOpen = False
    PopulateTree
End Sub

Private Sub mnuFileOpenODBC_Click()
    mbOpenVBIDE = False
    msClassDatabaseName = vbNullString
    msClassDatabaseOptions = vbNullString
    ODBCTableNamePrefix = vbNullString
    ODBCPassword = vbNullString
    IsInODBCDatabaseMode = False
    KeepDatabaseOpen = False
    PopulateTree
End Sub

Private Sub mnuFileOpenVBIDE_Click()
    mbOpenVBIDE = True
    msClassDatabaseName = vbNullString
    msClassDatabaseOptions = vbNullString
    ODBCTableNamePrefix = vbNullString
    ODBCPassword = vbNullString
    IsInODBCDatabaseMode = False
    KeepDatabaseOpen = False
    PopulateTree
End Sub


Private Sub mnuFreeAssociateTables_Click()
    mnuFreeAssociateTables.Checked = Not mnuFreeAssociateTables.Checked
    SaveSetting "SliceAndDice", "DB Class Gen", "Free Associate Tables", mnuFreeAssociateTables.Checked
End Sub

Private Sub mnuGenerateClass_Click()
    If lvwFields.ListItems.Count = 0 Then
       MsgBox "Please select a table first."
       Exit Sub
    End If

    mbCanceled = False
    mbGenerateBranch = False
    mbGenerateDatabase = False
    Hide
    TriggerClassGeneration
End Sub

Private Sub mnuGenerateEnterBranch_Click()
    If lvwFields.ListItems.Count = 0 Then
       MsgBox "Please select a branch first."
       Exit Sub
    End If

    mbCanceled = False
    mbGenerateBranch = True
    mbGenerateDatabase = False
    Hide
    TriggerClassGeneration
End Sub


Private Sub mnuGenerateEntireDatabase_Click()
    If tvwTables.Nodes.Count = 0 Then
       MsgBox "Please select a database first."
       Exit Sub
    End If

    tvwTables.Nodes(1).Selected = True
    lvwFields.ListItems.Clear

    mbCanceled = False
    mbGenerateBranch = False
    mbGenerateDatabase = True
    Hide
    TriggerClassGeneration
End Sub

Public Sub mnuFileOpen_Click()
    mbOpenVBIDE = False
    msClassDatabaseName = Parent.sChooseDatabase()
    msClassDatabaseOptions = vbNullString
    ODBCTableNamePrefix = vbNullString
    ODBCPassword = vbNullString
    IsInODBCDatabaseMode = False
    KeepDatabaseOpen = False
    If Len(msClassDatabaseName) > 0 Then
       PopulateTree
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
       
   RefreshCategories
   
   mnuRulesAutoAddKey.Checked = GetSetting("SliceAndDice", "DB Class Gen", "AutoAddKey", True)
   mnuRulesAutoAddDateModified.Checked = GetSetting("SliceAndDice", "DB Class Gen", "AutoAddDateModified", True)
   mnuRulesAutoAddDateCreated.Checked = GetSetting("SliceAndDice", "DB Class Gen", "AutoAddCreated", True)

   mnuRulesEnforce.Checked = GetSetting("SliceAndDice", "DB Class Gen", "Enforce", True)
   mnuRulesCascadeUpdates.Checked = GetSetting("SliceAndDice", "DB Class Gen", "CascadeUpdates", True)
   mnuRulesCascadeDeletes.Checked = GetSetting("SliceAndDice", "DB Class Gen", "CascadeDeletes", True)

   mnuRulesUseAutoNumber.Checked = GetSetting("SliceAndDice", "DB Class Gen", "UseAutoNumber", True)
   mnuViewTableData.Checked = GetSetting("SliceAndDice", "DB Class Gen", "ViewTableData", False)
   mnuRelateOnLoad.Checked = GetSetting("SliceAndDice", "DB Class Gen", "Relate on load", True)
   mnuFreeAssociateTables.Checked = GetSetting("SliceAndDice", "DB Class Gen", "Free Asoociate Tables", True)

   If mnuRulesEnforce.Checked Then
      mnuRulesCascadeUpdates.Enabled = True
      mnuRulesCascadeDeletes.Enabled = True
   Else
      mnuRulesCascadeUpdates.Enabled = False
      mnuRulesCascadeDeletes.Enabled = False
   End If

   Set Favorites = CreateObject("SandySupport.CAssocArray")
       Favorites.KeyValueDelimiter = "=="
       Favorites.FieldDelimiter = "|"
       Favorites.ItemDelimiter = "<EOE>"
       Favorites.All = GetSetting("SliceAndDice", "DB Class Gen", "Favorites", vbNullString)
   
   UpdateFavorites
   
   LoadFormPosition Me

   dvwTable.Visible = mnuViewTableData.Checked
End Sub

    
Private Sub Form_Terminate()
    SaveFormPosition Me
    Set Favorites = Nothing
    Set Parent = Nothing
    ' LogEvent "frmDBClassGen: Terminate"
End Sub


Private Sub lvwFields_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim ItemClicked As ListItem

    If Button = vbRightButton Then
       Set ItemClicked = lvwFields.HitTest(X, Y)
       If Not ItemClicked Is Nothing Then
          ItemClicked.Selected = True
          PopupMenu mnuField
       End If
    End If
End Sub


Private Sub mnuFieldDelete_Click()
    Dim sParentTable As String
    Dim sTable       As String
    Dim sField       As String

On Error Resume Next
    If tvwTables.SelectedItem.Parent.Key = msClassDatabaseName Then
       sParentTable = vbNullString
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

    If IsInODBCDatabaseMode Then
       If Not KeepDatabaseOpen Then Set db = OpenDatabase(msClassDatabaseName, dbDriverComplete, False, msClassDatabaseOptions)
    Else
       Set db = OpenDatabase(msClassDatabaseName, False, False)
    End If
        db.TableDefs(sTable).Fields.Delete sField
    If Not KeepDatabaseOpen Then db.Close

    tvwTables_NodeClick tvwTables.SelectedItem
End Sub

Private Sub mnuFieldNew_Click()
    Dim CurTable   As TableDef
    Dim CurItem    As ListItem
    Dim NewField   As New frmFieldType

mnuFieldNew_Click_TryAgain:
    With NewField
On Error Resume Next
         Err.Clear
         .Show vbModal, Me
         If Err.Number <> 0 Then Exit Sub
         If .Canceled = True Then Exit Sub

         For Each CurItem In lvwFields.ListItems
             If UCase$(CurItem.Text) = UCase$(.FieldName) Then
                MsgBox "That field already exists in this table... Try again."
                GoTo mnuFieldNew_Click_TryAgain
             End If
         Next CurItem

On Error Resume Next
         If IsInODBCDatabaseMode Then
            If Not KeepDatabaseOpen Then Set db = OpenDatabase(msClassDatabaseName, dbDriverComplete, False, msClassDatabaseOptions)
         Else
            Set db = OpenDatabase(msClassDatabaseName, False, False)
         End If
             Set CurTable = db.TableDefs(tvwTables.SelectedItem.Key)
             Err.Clear
             If .dbFieldType = dbText Then
                CurTable.Fields.Append CurTable.CreateField(.FieldName, .dbFieldType, .Length)
             Else
                CurTable.Fields.Append CurTable.CreateField(.FieldName, .dbFieldType)
             End If
             If Err.Number Then
                MsgBox "Error occured adding field '" & .FieldName & "' to table '" & CurTable.Name & "'" & vbCr & vbTab & "Err #" & Err.Number & vbCr & vbTab & "Desc:" & Err.Description
                Err.Clear
             End If
         If Not KeepDatabaseOpen Then db.Close

         tvwTables_NodeClick tvwTables.SelectedItem
    End With
End Sub

Private Sub mnuRelateOnLoad_Click()
    mnuRelateOnLoad.Checked = Not mnuRelateOnLoad.Checked
    SaveSetting "SliceAndDice", "DB Class Gen", "Relate on load", mnuRelateOnLoad.Checked
End Sub

Private Sub mnuRemoveMarked_Click()
    RemoveNodes tvwTables, "TableMarked"
    RefreshTableList
    AddFavorite
End Sub

Private Sub mnuRemoveUnmarked_Click()
    RemoveNodes tvwTables, "Table"
    RefreshTableList
    AddFavorite
End Sub

Private Sub mnuRulesAutoAddDateCreated_Click()
    mnuRulesAutoAddDateCreated.Checked = Not mnuRulesAutoAddDateCreated.Checked
    SaveSetting "SliceAndDice", "DB Class Gen", "AutoAddDateCreated", mnuRulesAutoAddDateCreated.Checked
End Sub

Private Sub mnuRulesAutoAddDateModified_Click()
    mnuRulesAutoAddDateModified.Checked = Not mnuRulesAutoAddDateModified.Checked
    SaveSetting "SliceAndDice", "DB Class Gen", "AutoAddDateModified", mnuRulesAutoAddDateModified.Checked
End Sub


Private Sub mnuRulesAutoAddKey_Click()
    mnuRulesAutoAddKey.Checked = Not mnuRulesAutoAddKey.Checked
    SaveSetting "SliceAndDice", "DB Class Gen", "AutoAddKey", mnuRulesAutoAddKey.Checked
End Sub

Private Sub mnuRulesCascadeDeletes_Click()
    mnuRulesCascadeDeletes.Checked = Not mnuRulesCascadeDeletes.Checked
    SaveSetting "SliceAndDice", "DB Class Gen", "CascadeDeletes", mnuRulesCascadeDeletes.Checked
End Sub


Private Sub mnuRulesCascadeUpdates_Click()
    mnuRulesCascadeUpdates.Checked = Not mnuRulesCascadeUpdates.Checked
    SaveSetting "SliceAndDice", "DB Class Gen", "CascadeUpdates", mnuRulesCascadeUpdates.Checked
End Sub

Private Sub mnuRulesEnforce_Click()
    mnuRulesEnforce.Checked = Not mnuRulesEnforce.Checked
    SaveSetting "SliceAndDice", "DB Class Gen", "Enforce", mnuRulesEnforce.Checked
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
    SaveSetting "SliceAndDice", "DB Class Gen", "UseAutoNumber", mnuRulesUseAutoNumber.Checked
End Sub

Private Sub mnuShowAllTables_Click()
    TableList = vbNullString
    AddFavorite
    PopulateTree
End Sub

Private Sub mnuTableDelete_Click()
On Error Resume Next
    If bUserSure("This will PERMANENTLY remove the table selected." & vbNewLine & vbTab & "ARE YOU ABSOLUTELY SURE ?") Then
       If IsInODBCDatabaseMode Then
          If Not KeepDatabaseOpen Then Set db = OpenDatabase(msClassDatabaseName, dbDriverComplete, False, msClassDatabaseOptions)
       Else
          Set db = OpenDatabase(msClassDatabaseName, False, False)
       End If
           db.TableDefs.Delete tvwTables.SelectedItem.Key
       If Not KeepDatabaseOpen Then db.Close
       PopulateTree
    End If
End Sub

Private Sub mnuTableNew_Click()
    Dim sTable As String

mnuTableNew_Click_TryAgain:
    sTable = Replace(Trim$(InputBox("What should the name of the new table be ?" & vbNewLine & vbTab & "Note: Table names should be singular, such as:" & vbNewLine & vbTab & "Book, Publisher, etc.")), " ", vbNullString)
    If Len(sTable) = 0 Then Exit Sub
    
    If Right$(sTable, 1) = "s" Then
       MsgBox "Table names MUST be singular."
       GoTo mnuTableNew_Click_TryAgain
    End If

    If tvwTables.SelectedItem.Text <> DBName Then
       AddTable sTable, tvwTables.SelectedItem.Text
    Else
       AddTable sTable, vbNullString
    End If
    
    PopulateTree
End Sub

Private Sub mnuToggleTableMark_Click()
    If tvwTables.SelectedItem.Image = "Table" Then
       tvwTables.SelectedItem.Image = "TableMarked"
       tvwTables.SelectedItem.ExpandedImage = "TableMarked"
       tvwTables.SelectedItem.SelectedImage = "TableMarked"
    Else
       tvwTables.SelectedItem.Image = "Table"
       tvwTables.SelectedItem.ExpandedImage = "Table"
       tvwTables.SelectedItem.SelectedImage = "Table"
    End If
    
End Sub

Private Sub mnuUnhideTable_Click()
On Error Resume Next
    Dim sTableName As String
    sTableName = InputBox("What is the fully qualified name of the table to unhide ?", "UNHIDE A TABLE")
    If Len(sTableName) Then
       With tvwTables.Nodes.Add(tvwTables.SelectedItem.Key, tvwChild, sTableName, sTableName, "Table", "Table")
            .ExpandedImage = "Table"
            .Expanded = True
       End With
    End If
    RefreshTableList
    AddFavorite
End Sub

Private Sub mnuViewTableData_Click()
On Error Resume Next
    mnuViewTableData.Checked = Not mnuViewTableData.Checked
    SaveSetting "SliceAndDice", "DB Class Gen", "ViewTableData", mnuViewTableData.Checked
    dvwTable.Visible = mnuViewTableData.Checked
    Form_Resize
End Sub

Private Sub mnuX_Click()
    mnuFileExit_Click
End Sub

Private Sub tvwTables_DblClick()
    mnuTableNew_Click
End Sub

Private Sub tvwTables_DragDrop(Source As Control, X As Single, Y As Single)
On Error Resume Next
    Dim NodeDroppedOnto As Node
    
    If Source.Tag = "tvwTables" Then
       If Not NodeDragged Is Nothing Then
          Set NodeDroppedOnto = Nothing
          Set NodeDroppedOnto = tvwTables.HitTest(X, Y)
          If Not NodeDroppedOnto Is Nothing Then
             Set NodeDragged.Parent = NodeDroppedOnto
             RefreshTableList
             AddFavorite
          End If
       End If
    End If
End Sub

Private Sub tvwTables_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       mnuTableNew_Click
    End If
End Sub

Private Sub tvwTables_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And (Shift And vbShiftMask) <> 0 Then
       Set NodeDragged = Nothing
       Set NodeDragged = tvwTables.HitTest(X, Y)
       If Not NodeDragged Is Nothing Then
          tvwTables.Drag vbBeginDrag
       End If
    End If
End Sub

Private Sub tvwTables_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NodeClicked As Node

    If Button = vbRightButton Then
       Set NodeClicked = tvwTables.HitTest(X, Y)
       If Not NodeClicked Is Nothing Then
          NodeClicked.Selected = True
          PopupMenu mnuTable
       End If
    End If
End Sub

Public Sub tvwTables_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim CurTable As TableDef
    Dim CurField As Field
    Dim nodX     As Node
    Dim litX     As ListItem
    Dim sIcon    As String

On Error Resume Next
    Screen.MousePointer = vbHourglass
    If IsInODBCDatabaseMode Then
       If Not KeepDatabaseOpen Then Set db = OpenDatabase(msClassDatabaseName, dbDriverComplete, False, msClassDatabaseOptions)
    Else
       Set db = OpenDatabase(msClassDatabaseName, False, False)
    End If
        With lvwFields.ListItems
             .Clear
             With lvwFields.ColumnHeaders
                  .Clear
                  .Add , "Field Name", "Field Name", 2600
                  .Add , "Field Type", "Type", 1000
                  .Add , "Field Length", "Length", 500
             End With
             lvwFields.View = lvwReport
             Set CurTable = db.TableDefs(tvwTables.SelectedItem.Text)

             For Each CurField In CurTable.Fields
                 If Right$(CurField.Name, 2) = "ID" Then
                    If Left$(CurField.Name, Len(CurField.Name) - 2) = CurTable.Name Then
                       sIcon = "Key"
                    Else
                       sIcon = "ID"
                    End If
                 ElseIf CurField.Type = dbDate Then
                    sIcon = "FieldDate"
                 ElseIf CurField.Type = dbMemo Then
                    sIcon = "FieldMemo"
                 ElseIf CurField.Type = dbText Then
                    sIcon = "FieldString"
                 Else
                    sIcon = "FieldNumber"
                 End If
                 Set litX = .Add(, CurField.Name, CurField.Name, sIcon, sIcon)
                 litX.SubItems(1) = sFieldType(CurField.Type)
                 litX.SubItems(2) = CurField.Size
             Next CurField
        End With
        
        If dvwTable.Visible And Not tvwTables.SelectedItem.Parent Is Nothing Then
On Error Resume Next
           With dvwTable
                .DatabaseName = "x"
                .RecordSource = "SELECT * FROM [" & tvwTables.SelectedItem.Text & "]"
                .View = lvwReport
                .GridLines = True
                .FullRowSelect = True
                .Requery db
           End With
        End If

    If Not KeepDatabaseOpen Then db.Close
    Set nodX = Nothing
    Screen.MousePointer = vbDefault
End Sub

