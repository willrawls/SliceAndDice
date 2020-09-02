using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class frmDBClassGen : System.Windows.Forms.Form
        {
         public System.Windows.Forms.VB.Frame fraCategory;
         public System.Windows.Forms.VB.CommandButton cmdDeleteCategory;
         public System.Windows.Forms.VB.CommandButton cmdAddCategory;
         public System.Windows.Forms.VB.ComboBox cboDataLibraryType;
         public System.Windows.Forms.FirmSolutions.DataView dvwTable;
         public System.Windows.Forms.MSComctlLib.ImageList imlTabs;
         public System.Windows.Forms.MSComctlLib.TreeView tvwTables;
         public System.Windows.Forms.MSComctlLib.ListView lvwFields;
         public System.Windows.Forms.VB.Menu mnuX;
         public System.Windows.Forms.VB.Menu mnuFile;
         public System.Windows.Forms.VB.Menu mnuFileOpen;
         public System.Windows.Forms.VB.Menu mnuFileOpenODBC;
         public System.Windows.Forms.VB.Menu mnuFileOpenVBIDE;
         public System.Windows.Forms.VB.Menu mnuFileSep1;
         public System.Windows.Forms.VB.Menu mnuFileOpenPrevious;
         public System.Windows.Forms.VB.Menu mnuFavorite0;
         public System.Windows.Forms.VB.Menu mnuFileSep3;
         public System.Windows.Forms.VB.Menu mnuFavRemoveAll;
         public System.Windows.Forms.VB.Menu mnuFileSep2;
         public System.Windows.Forms.VB.Menu mnuFileNew;
         public System.Windows.Forms.VB.Menu mnuSep5;
         public System.Windows.Forms.VB.Menu mnuRelateOnLoad;
         public System.Windows.Forms.VB.Menu mnuFreeAssociateTables;
         public System.Windows.Forms.VB.Menu mnuFileSep0;
         public System.Windows.Forms.VB.Menu mnuFileExit;
         public System.Windows.Forms.VB.Menu mnuGenerate;
         public System.Windows.Forms.VB.Menu mnuGenerateEntireDatabase;
         public System.Windows.Forms.VB.Menu mnuGenerateClass;
         public System.Windows.Forms.VB.Menu mnuGenerateEnterBranch;
         public System.Windows.Forms.VB.Menu mnuGenerateCustom;
         public System.Windows.Forms.VB.Menu mnuRules;
         public System.Windows.Forms.VB.Menu mnuRulesAutoAdd;
         public System.Windows.Forms.VB.Menu mnuRulesAutoAddDateCreated;
         public System.Windows.Forms.VB.Menu mnuRulesAutoAddDateModified;
         public System.Windows.Forms.VB.Menu mnuRulesAutoAddKey;
         public System.Windows.Forms.VB.Menu mnuRulesAutoAddSep0;
         public System.Windows.Forms.VB.Menu mnuRulesAutoAddCustom;
         public System.Windows.Forms.VB.Menu mnuRelationship;
         public System.Windows.Forms.VB.Menu mnuRulesEnforce;
         public System.Windows.Forms.VB.Menu mnuRulesCascadeUpdates;
         public System.Windows.Forms.VB.Menu mnuRulesCascadeDeletes;
         public System.Windows.Forms.VB.Menu mnuRulesUseAutoNumber;
         public System.Windows.Forms.VB.Menu mnuHelp;
         public System.Windows.Forms.VB.Menu mnuHelpContents;
         public System.Windows.Forms.VB.Menu mnuHelpAbout;
         public System.Windows.Forms.VB.Menu mnuShortcut;
         public System.Windows.Forms.VB.Menu mnuTable;
         public System.Windows.Forms.VB.Menu mnuTableNew;
         public System.Windows.Forms.VB.Menu mnuTableRename;
         public System.Windows.Forms.VB.Menu mnuSep3;
         public System.Windows.Forms.VB.Menu mnuToggleTableMark;
         public System.Windows.Forms.VB.Menu mnuRemoveUnmarked;
         public System.Windows.Forms.VB.Menu mnuRemoveMarked;
         public System.Windows.Forms.VB.Menu mnuUnhideTable;
         public System.Windows.Forms.VB.Menu mnuShowAllTables;
         public System.Windows.Forms.VB.Menu mnuViewTableData;
         public System.Windows.Forms.VB.Menu mnuSep0;
         public System.Windows.Forms.VB.Menu mnuTableDelete;
         public System.Windows.Forms.VB.Menu mnuField;
         public System.Windows.Forms.VB.Menu mnuFieldNew;
         public System.Windows.Forms.VB.Menu mnuFieldRename;
         public System.Windows.Forms.VB.Menu mnuSep1;
         public System.Windows.Forms.VB.Menu mnuFieldDelete;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.Container components = null;

        public frmDBClassGen()
        {
            // Required for Windows Form Designer support
            InitializeComponent();

            // TODO: Add any constructor code after InitializeComponent call
        }
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose( bool disposing )
        {
            if( disposing )
            {
              if (components != null)
              {
                components.Dispose();
              }
            }
            base.Dispose( disposing );
        }
        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        public void InitializeComponent()
        {
            this.fraCategory = new System.Windows.Forms.VB.Frame();
            this.cmdDeleteCategory = new System.Windows.Forms.VB.CommandButton();
            this.cmdAddCategory = new System.Windows.Forms.VB.CommandButton();
            this.cboDataLibraryType = new System.Windows.Forms.VB.ComboBox();
            this.dvwTable = new System.Windows.Forms.FirmSolutions.DataView();
            this.imlTabs = new System.Windows.Forms.MSComctlLib.ImageList();
            this.tvwTables = new System.Windows.Forms.MSComctlLib.TreeView();
            this.lvwFields = new System.Windows.Forms.MSComctlLib.ListView();
            this.mnuX = new System.Windows.Forms.VB.Menu();
            this.mnuFile = new System.Windows.Forms.VB.Menu();
            this.mnuFileOpen = new System.Windows.Forms.VB.Menu();
            this.mnuFileOpenODBC = new System.Windows.Forms.VB.Menu();
            this.mnuFileOpenVBIDE = new System.Windows.Forms.VB.Menu();
            this.mnuFileSep1 = new System.Windows.Forms.VB.Menu();
            this.mnuFileOpenPrevious = new System.Windows.Forms.VB.Menu();
            this.mnuFavorite0 = new System.Windows.Forms.VB.Menu();
            this.mnuFileSep3 = new System.Windows.Forms.VB.Menu();
            this.mnuFavRemoveAll = new System.Windows.Forms.VB.Menu();
            this.mnuFileSep2 = new System.Windows.Forms.VB.Menu();
            this.mnuFileNew = new System.Windows.Forms.VB.Menu();
            this.mnuSep5 = new System.Windows.Forms.VB.Menu();
            this.mnuRelateOnLoad = new System.Windows.Forms.VB.Menu();
            this.mnuFreeAssociateTables = new System.Windows.Forms.VB.Menu();
            this.mnuFileSep0 = new System.Windows.Forms.VB.Menu();
            this.mnuFileExit = new System.Windows.Forms.VB.Menu();
            this.mnuGenerate = new System.Windows.Forms.VB.Menu();
            this.mnuGenerateEntireDatabase = new System.Windows.Forms.VB.Menu();
            this.mnuGenerateClass = new System.Windows.Forms.VB.Menu();
            this.mnuGenerateEnterBranch = new System.Windows.Forms.VB.Menu();
            this.mnuGenerateCustom = new System.Windows.Forms.VB.Menu();
            this.mnuRules = new System.Windows.Forms.VB.Menu();
            this.mnuRulesAutoAdd = new System.Windows.Forms.VB.Menu();
            this.mnuRulesAutoAddDateCreated = new System.Windows.Forms.VB.Menu();
            this.mnuRulesAutoAddDateModified = new System.Windows.Forms.VB.Menu();
            this.mnuRulesAutoAddKey = new System.Windows.Forms.VB.Menu();
            this.mnuRulesAutoAddSep0 = new System.Windows.Forms.VB.Menu();
            this.mnuRulesAutoAddCustom = new System.Windows.Forms.VB.Menu();
            this.mnuRelationship = new System.Windows.Forms.VB.Menu();
            this.mnuRulesEnforce = new System.Windows.Forms.VB.Menu();
            this.mnuRulesCascadeUpdates = new System.Windows.Forms.VB.Menu();
            this.mnuRulesCascadeDeletes = new System.Windows.Forms.VB.Menu();
            this.mnuRulesUseAutoNumber = new System.Windows.Forms.VB.Menu();
            this.mnuHelp = new System.Windows.Forms.VB.Menu();
            this.mnuHelpContents = new System.Windows.Forms.VB.Menu();
            this.mnuHelpAbout = new System.Windows.Forms.VB.Menu();
            this.mnuShortcut = new System.Windows.Forms.VB.Menu();
            this.mnuTable = new System.Windows.Forms.VB.Menu();
            this.mnuTableNew = new System.Windows.Forms.VB.Menu();
            this.mnuTableRename = new System.Windows.Forms.VB.Menu();
            this.mnuSep3 = new System.Windows.Forms.VB.Menu();
            this.mnuToggleTableMark = new System.Windows.Forms.VB.Menu();
            this.mnuRemoveUnmarked = new System.Windows.Forms.VB.Menu();
            this.mnuRemoveMarked = new System.Windows.Forms.VB.Menu();
            this.mnuUnhideTable = new System.Windows.Forms.VB.Menu();
            this.mnuShowAllTables = new System.Windows.Forms.VB.Menu();
            this.mnuViewTableData = new System.Windows.Forms.VB.Menu();
            this.mnuSep0 = new System.Windows.Forms.VB.Menu();
            this.mnuTableDelete = new System.Windows.Forms.VB.Menu();
            this.mnuField = new System.Windows.Forms.VB.Menu();
            this.mnuFieldNew = new System.Windows.Forms.VB.Menu();
            this.mnuFieldRename = new System.Windows.Forms.VB.Menu();
            this.mnuSep1 = new System.Windows.Forms.VB.Menu();
            this.mnuFieldDelete = new System.Windows.Forms.VB.Menu();
            this.SuspendLayout();
            this.fraCategory.SuspendLayout();
            this.mnuFile.SuspendLayout();
            this.mnuFileOpenPrevious.SuspendLayout();
            this.mnuGenerate.SuspendLayout();
            this.mnuRules.SuspendLayout();
            this.mnuRulesAutoAdd.SuspendLayout();
            this.mnuRelationship.SuspendLayout();
            this.mnuHelp.SuspendLayout();
            this.mnuTable.SuspendLayout();
            this.mnuField.SuspendLayout();
            //
            // fraCategory
            //
            this.fraCategory.Name = "fraCategory";
            this.fraCategory.Text = "Data Library Category";
            this.fraCategory.Size = new System.Drawing.Size(317, 44);
            this.fraCategory.Location = new System.Drawing.Point(270, -2);
            this.fraCategory.TabIndex = 2;
            this.fraCategory.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.cmdDeleteCategory,
                        this.cmdAddCategory,
                        this.cboDataLibraryType
            });
            //
            // cmdDeleteCategory
            //
            this.cmdDeleteCategory.Name = "cmdDeleteCategory";
            this.cmdDeleteCategory.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cmdDeleteCategory.Text = "Remove";
            this.cmdDeleteCategory.Size = new System.Drawing.Size(51, 27);
            this.cmdDeleteCategory.Location = new System.Drawing.Point(259, 12);
            this.cmdDeleteCategory.TabIndex = 5;
//            this.cmdDeleteCategory.ToolTipText = "Remove the current Slice and Dice Category from the;
            //
            // cmdAddCategory
            //
            this.cmdAddCategory.Name = "cmdAddCategory";
            this.cmdAddCategory.Text = "Add";
            this.cmdAddCategory.Size = new System.Drawing.Size(51, 27);
            this.cmdAddCategory.Location = new System.Drawing.Point(204, 12);
            this.cmdAddCategory.TabIndex = 4;
//            this.cmdAddCategory.ToolTipText = "Add a new Slice and Dice DB to Code Category";
            //
            // cboDataLibraryType
            //
            this.cboDataLibraryType.Name = "cboDataLibraryType";
            this.cboDataLibraryType.Size = new System.Drawing.Size(192, 21);
//            this.cboDataLibraryType.IntegralHeight = 0;
//            this.cboDataLibraryType.ItemData = "frmDBClassGen.frx":0000;
            this.cboDataLibraryType.Location = new System.Drawing.Point(8, 16);
//            this.cboDataLibraryType.List = "frmDBClassGen.frx":0002;
            this.cboDataLibraryType.TabIndex = 3;
            this.cboDataLibraryType.Text = "RDO Persisted";
            //
            // dvwTable
            //
            this.dvwTable.Name = "dvwTable";
            this.dvwTable.Size = new System.Drawing.Size(589, 143);
            this.dvwTable.Location = new System.Drawing.Point(4, 264);
            this.dvwTable.TabIndex = 1;
            this.dvwTable.Tag = "dvwTable";
            this.dvwTable.BackColor = System.Drawing.Color.FromArgb(-2147483633);
            this.dvwTable.Font = new System.Drawing.Font("MS Sans Serif",8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
//            this.dvwTable.ScaleMode = 0;
//            this.dvwTable.HotTracking = -1;
//            this.dvwTable.FontSize = 8.25;
//            this.dvwTable.FontName = "MS Sans Serif";
            //
            // imlTabs
            //
            this.imlTabs.Name = "imlTabs";
            this.imlTabs.Location = new System.Drawing.Point(0, 0);
            this.imlTabs.BackColor = System.Drawing.Color.FromArgb(-2147483643);
//            this.imlTabs.ImageWidth = 16;
//            this.imlTabs.ImageHeight = 16;
//            this.imlTabs.MaskColor = 12632256;
//            this.imlTabs.ListImage1 = ;
//            this.imlTabs.ListImage2 = ;
//            this.imlTabs.ListImage3 = ;
//            this.imlTabs.ListImage4 = ;
//            this.imlTabs.ListImage5 = ;
//            this.imlTabs.ListImage6 = ;
//            this.imlTabs.ListImage7 = ;
//            this.imlTabs.ListImage8 = ;
//            this.imlTabs.ListImage9 = ;
//            this.imlTabs.ListImage10 = ;
//            this.imlTabs.ListImage11 = ;
//            this.imlTabs.ListImage12 = ;
//            this.imlTabs.ListImage12 = ;
            //
            // tvwTables
            //
            this.tvwTables.Name = "tvwTables";
//            this.tvwTables.DragIcon = "frmDBClassGen.frx":31BC;
            this.tvwTables.Size = new System.Drawing.Size(262, 396);
            this.tvwTables.Location = new System.Drawing.Point(4, 4);
            this.tvwTables.TabIndex = 0;
            this.tvwTables.Tag = "tvwTables";
            this.tvwTables.HideSelection = false;
//            this.tvwTables.Indentation = 353;
            this.tvwTables.LabelEdit = true;
//            this.tvwTables.ImageList = "imlTabs";
            this.tvwTables.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            //
            // lvwFields
            //
            this.lvwFields.Name = "lvwFields";
            this.lvwFields.Size = new System.Drawing.Size(319, 354);
            this.lvwFields.Location = new System.Drawing.Point(270, 45);
            this.lvwFields.TabIndex = 6;
            this.lvwFields.Tag = "lvwFields";
            this.lvwFields.LabelWrap = false;
            this.lvwFields.HideSelection = false;
//            this.lvwFields.FullRowSelect = -1;
//            this.lvwFields.GridLines = -1;
//            this.lvwFields.HotTracking = -1;
//            this.lvwFields.Icons = "imlTabs";
//            this.lvwFields.SmallIcons = "imlTabs";
//            this.lvwFields.ColHdrIcons = "imlTabs";
            this.lvwFields.ForeColor = System.Drawing.Color.FromArgb(-2147483640);
            this.lvwFields.BackColor = System.Drawing.Color.FromArgb(-2147483643);
            this.lvwFields.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
//            this.lvwFields.NumItems = 0;
            //
            // mnuX
            //
            this.mnuX.Name = "mnuX";
            this.mnuX.Text = "X";
            this.mnuX.Enabled = false;
            this.mnuX.Visible = false;
            //
            // mnuFile
            //
            this.mnuFile.Name = "mnuFile";
            this.mnuFile.Text = "&File";
            this.mnuFile.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuFileOpen,
                        this.mnuFileOpenODBC,
                        this.mnuFileOpenVBIDE,
                        this.mnuFileSep1,
                        this.mnuFileOpenPrevious,
                        this.mnuFileSep2,
                        this.mnuFileNew,
                        this.mnuSep5,
                        this.mnuRelateOnLoad,
                        this.mnuFreeAssociateTables,
                        this.mnuFileSep0,
                        this.mnuFileExit
            });
            //
            // mnuFileOpen
            //
            this.mnuFileOpen.Name = "mnuFileOpen";
            this.mnuFileOpen.Text = "&Open an Access97 database";
            //
            // mnuFileOpenODBC
            //
            this.mnuFileOpenODBC.Name = "mnuFileOpenODBC";
            this.mnuFileOpenODBC.Text = "Open an O&DBC database";
            //
            // mnuFileOpenVBIDE
            //
            this.mnuFileOpenVBIDE.Name = "mnuFileOpenVBIDE";
            this.mnuFileOpenVBIDE.Text = """Open"" &VB IDE";
            this.mnuFileOpenVBIDE.Enabled = false;
            //
            // mnuFileSep1
            //
            this.mnuFileSep1.Name = "mnuFileSep1";
            this.mnuFileSep1.Text = "-";
            //
            // mnuFileOpenPrevious
            //
            this.mnuFileOpenPrevious.Name = "mnuFileOpenPrevious";
            this.mnuFileOpenPrevious.Text = "Open a &Previously used database";
            this.mnuFileOpenPrevious.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuFavorite0,
                        this.mnuFileSep3,
                        this.mnuFavRemoveAll
            });
            //
            // mnuFavorite0
            //
            this.mnuFavorite0.Name = "mnuFavorite0";
            this.mnuFavorite0.Text = "-Empty-";
            this.mnuFavorite0.Enabled = false;
            //
            // mnuFileSep3
            //
            this.mnuFileSep3.Name = "mnuFileSep3";
            this.mnuFileSep3.Text = "-";
            //
            // mnuFavRemoveAll
            //
            this.mnuFavRemoveAll.Name = "mnuFavRemoveAll";
            this.mnuFavRemoveAll.Text = "Remove all Favorites";
            //
            // mnuFileSep2
            //
            this.mnuFileSep2.Name = "mnuFileSep2";
            this.mnuFileSep2.Text = "-";
            //
            // mnuFileNew
            //
            this.mnuFileNew.Name = "mnuFileNew";
            this.mnuFileNew.Text = "Create a &New Access97 database";
            //
            // mnuSep5
            //
            this.mnuSep5.Name = "mnuSep5";
            this.mnuSep5.Text = "-";
            //
            // mnuRelateOnLoad
            //
            this.mnuRelateOnLoad.Name = "mnuRelateOnLoad";
            this.mnuRelateOnLoad.Text = "Relate tabes on load (S&&D mentality)";
//            this.mnuRelateOnLoad.Checked = -1;
            //
            // mnuFreeAssociateTables
            //
            this.mnuFreeAssociateTables.Name = "mnuFreeAssociateTables";
            this.mnuFreeAssociateTables.Text = "Free Associate tables (No limitations)";
            //
            // mnuFileSep0
            //
            this.mnuFileSep0.Name = "mnuFileSep0";
            this.mnuFileSep0.Text = "-";
            //
            // mnuFileExit
            //
            this.mnuFileExit.Name = "mnuFileExit";
            this.mnuFileExit.Text = "E&xit";
            //
            // mnuGenerate
            //
            this.mnuGenerate.Name = "mnuGenerate";
            this.mnuGenerate.Text = "&Generate";
            this.mnuGenerate.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuGenerateEntireDatabase,
                        this.mnuGenerateClass,
                        this.mnuGenerateEnterBranch,
                        this.mnuGenerateCustom
            });
            //
            // mnuGenerateEntireDatabase
            //
            this.mnuGenerateEntireDatabase.Name = "mnuGenerateEntireDatabase";
            this.mnuGenerateEntireDatabase.Text = "Entire &Database                 (everything and a wrapper class)";
            //
            // mnuGenerateClass
            //
            this.mnuGenerateClass.Name = "mnuGenerateClass";
            this.mnuGenerateClass.Text = "Selected &Dictionary<string,string>() Class  (and Dictionary<string,string>() Member Class)";
            //
            // mnuGenerateEnterBranch
            //
            this.mnuGenerateEnterBranch.Name = "mnuGenerateEnterBranch";
            this.mnuGenerateEnterBranch.Text = "Entire &Branch                    (selected and all children)";
            //
            // mnuGenerateCustom
            //
            this.mnuGenerateCustom.Name = "mnuGenerateCustom";
            this.mnuGenerateCustom.Text = "Custom Generation               (requires a custom gen template)";
            this.mnuGenerateCustom.Enabled = false;
            //
            // mnuRules
            //
            this.mnuRules.Name = "mnuRules";
            this.mnuRules.Text = "&Rules";
            this.mnuRules.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuRulesAutoAdd,
                        this.mnuRelationship,
                        this.mnuRulesUseAutoNumber
            });
            //
            // mnuRulesAutoAdd
            //
            this.mnuRulesAutoAdd.Name = "mnuRulesAutoAdd";
            this.mnuRulesAutoAdd.Text = "Automatically add";
            this.mnuRulesAutoAdd.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuRulesAutoAddDateCreated,
                        this.mnuRulesAutoAddDateModified,
                        this.mnuRulesAutoAddKey,
                        this.mnuRulesAutoAddSep0,
                        this.mnuRulesAutoAddCustom
            });
            //
            // mnuRulesAutoAddDateCreated
            //
            this.mnuRulesAutoAddDateCreated.Name = "mnuRulesAutoAddDateCreated";
            this.mnuRulesAutoAddDateCreated.Text = "DateCreated";
//            this.mnuRulesAutoAddDateCreated.Checked = -1;
            //
            // mnuRulesAutoAddDateModified
            //
            this.mnuRulesAutoAddDateModified.Name = "mnuRulesAutoAddDateModified";
            this.mnuRulesAutoAddDateModified.Text = "DateModified";
//            this.mnuRulesAutoAddDateModified.Checked = -1;
            //
            // mnuRulesAutoAddKey
            //
            this.mnuRulesAutoAddKey.Name = "mnuRulesAutoAddKey";
            this.mnuRulesAutoAddKey.Text = "Name / Key (highly recommended)";
//            this.mnuRulesAutoAddKey.Checked = -1;
            //
            // mnuRulesAutoAddSep0
            //
            this.mnuRulesAutoAddSep0.Name = "mnuRulesAutoAddSep0";
            this.mnuRulesAutoAddSep0.Text = "-";
            //
            // mnuRulesAutoAddCustom
            //
            this.mnuRulesAutoAddCustom.Name = "mnuRulesAutoAddCustom";
            this.mnuRulesAutoAddCustom.Text = "Custom";
            this.mnuRulesAutoAddCustom.Enabled = false;
            //
            // mnuRelationship
            //
            this.mnuRelationship.Name = "mnuRelationship";
            this.mnuRelationship.Text = "Parent/Child Relationship";
            this.mnuRelationship.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuRulesEnforce,
                        this.mnuRulesCascadeUpdates,
                        this.mnuRulesCascadeDeletes
            });
            //
            // mnuRulesEnforce
            //
            this.mnuRulesEnforce.Name = "mnuRulesEnforce";
            this.mnuRulesEnforce.Text = "Enforce referencial integrity";
//            this.mnuRulesEnforce.Checked = -1;
            //
            // mnuRulesCascadeUpdates
            //
            this.mnuRulesCascadeUpdates.Name = "mnuRulesCascadeUpdates";
            this.mnuRulesCascadeUpdates.Text = "Cascade Updates";
//            this.mnuRulesCascadeUpdates.Checked = -1;
            //
            // mnuRulesCascadeDeletes
            //
            this.mnuRulesCascadeDeletes.Name = "mnuRulesCascadeDeletes";
            this.mnuRulesCascadeDeletes.Text = "Cascade Deletes";
//            this.mnuRulesCascadeDeletes.Checked = -1;
            //
            // mnuRulesUseAutoNumber
            //
            this.mnuRulesUseAutoNumber.Name = "mnuRulesUseAutoNumber";
            this.mnuRulesUseAutoNumber.Text = "Use AutoNumber for PrimaryID";
            //
            // mnuHelp
            //
            this.mnuHelp.Name = "mnuHelp";
            this.mnuHelp.Text = "&Help";
            this.mnuHelp.Visible = false;
            this.mnuHelp.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuHelpContents,
                        this.mnuHelpAbout
            });
            //
            // mnuHelpContents
            //
            this.mnuHelpContents.Name = "mnuHelpContents";
            this.mnuHelpContents.Text = "&Contents";
            //
            // mnuHelpAbout
            //
            this.mnuHelpAbout.Name = "mnuHelpAbout";
            this.mnuHelpAbout.Text = "&About Slice and Dice";
            //
            // mnuShortcut
            //
            this.mnuShortcut.Name = "mnuShortcut";
            this.mnuShortcut.Text = "Shortcut Menus";
            this.mnuShortcut.Visible = false;
            //
            // mnuTable
            //
            this.mnuTable.Name = "mnuTable";
            this.mnuTable.Text = "Table";
            this.mnuTable.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuTableNew,
                        this.mnuTableRename,
                        this.mnuSep3,
                        this.mnuToggleTableMark,
                        this.mnuRemoveUnmarked,
                        this.mnuRemoveMarked,
                        this.mnuUnhideTable,
                        this.mnuShowAllTables,
                        this.mnuViewTableData,
                        this.mnuSep0,
                        this.mnuTableDelete
            });
            //
            // mnuTableNew
            //
            this.mnuTableNew.Name = "mnuTableNew";
            this.mnuTableNew.Text = "&New Table";
            //
            // mnuTableRename
            //
            this.mnuTableRename.Name = "mnuTableRename";
            this.mnuTableRename.Text = "&Rename";
            this.mnuTableRename.Enabled = false;
            //
            // mnuSep3
            //
            this.mnuSep3.Name = "mnuSep3";
            this.mnuSep3.Text = "-";
            //
            // mnuToggleTableMark
            //
            this.mnuToggleTableMark.Name = "mnuToggleTableMark";
            this.mnuToggleTableMark.Text = "&Mark/Unmark Table";
            //
            // mnuRemoveUnmarked
            //
            this.mnuRemoveUnmarked.Name = "mnuRemoveUnmarked";
            this.mnuRemoveUnmarked.Text = "Remove Unmarked Tables from list";
            //
            // mnuRemoveMarked
            //
            this.mnuRemoveMarked.Name = "mnuRemoveMarked";
            this.mnuRemoveMarked.Text = "Remove Marked Tables from list";
            //
            // mnuUnhideTable
            //
            this.mnuUnhideTable.Name = "mnuUnhideTable";
            this.mnuUnhideTable.Text = "Unhide a Table by Name";
            //
            // mnuShowAllTables
            //
            this.mnuShowAllTables.Name = "mnuShowAllTables";
            this.mnuShowAllTables.Text = "Show all tables";
            //
            // mnuViewTableData
            //
            this.mnuViewTableData.Name = "mnuViewTableData";
            this.mnuViewTableData.Text = "&View Selected Table;
            //
            // mnuSep0
            //
            this.mnuSep0.Name = "mnuSep0";
            this.mnuSep0.Text = "-";
            //
            // mnuTableDelete
            //
            this.mnuTableDelete.Name = "mnuTableDelete";
            this.mnuTableDelete.Text = "Delete Table";
            //
            // mnuField
            //
            this.mnuField.Name = "mnuField";
            this.mnuField.Text = "Field";
            this.mnuField.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuFieldNew,
                        this.mnuFieldRename,
                        this.mnuSep1,
                        this.mnuFieldDelete
            });
            //
            // mnuFieldNew
            //
            this.mnuFieldNew.Name = "mnuFieldNew";
            this.mnuFieldNew.Text = "New Field";
            //
            // mnuFieldRename
            //
            this.mnuFieldRename.Name = "mnuFieldRename";
            this.mnuFieldRename.Text = "Modify";
            this.mnuFieldRename.Enabled = false;
            //
            // mnuSep1
            //
            this.mnuSep1.Name = "mnuSep1";
            this.mnuSep1.Text = "-";
            //
            // mnuFieldDelete
            //
            this.mnuFieldDelete.Name = "mnuFieldDelete";
            this.mnuFieldDelete.Text = "Delete Field";
            //
            // frmDBClassGen
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.fraCategory,
                  this.dvwTable,
                  this.imlTabs,
                  this.tvwTables,
                  this.lvwFields,
                  this.mnuX,
                  this.mnuFile,
                  this.mnuGenerate,
                  this.mnuRules,
                  this.mnuHelp,
                  this.mnuShortcut,
                  this.mnuTable,
                  this.mnuField
            });
            this.Name = "frmDBClassGen";
            this.fraCategory.ResumeLayout(false);
            this.mnuFile.ResumeLayout(false);
            this.mnuFileOpenPrevious.ResumeLayout(false);
            this.mnuGenerate.ResumeLayout(false);
            this.mnuRules.ResumeLayout(false);
            this.mnuRulesAutoAdd.ResumeLayout(false);
            this.mnuRelationship.ResumeLayout(false);
            this.mnuHelp.ResumeLayout(false);
            this.mnuTable.ResumeLayout(false);
            this.mnuField.ResumeLayout(false);
            this.ResumeLayout(false);
        }
        #endregion

        public bool mbCanceled;
        public bool mbGenerateDatabase;
        public bool mbGenerateBranch;
        public bool mbLoadingCategories;
        public bool mbOpenVBIDE;
        public string msClassDatabaseName;
        public string msClassDatabaseOptions;
        public bool KeepDatabaseOpen;
        public bool IsInODBCDatabaseMode;
        public string ODBCTableNamePrefix;
        public string ODBCPassword;
        public string TableList;
        public Database db;
        public CAssocArray Favorites;
        public int FavoriteCount;
        public bool RetrievingAFavoriteNow;
        public frmMain Parent;
        public Node NodeDragged;
        public Node CurChild;
        public CAssocArray asaV;
        public ListItem CurListItem;
        public Node CurChild;
        public string sTableName;
        public string sFieldType;
        public string sDBPCType;
        public string sParentName;
        public string sClassToCollect;
        public string sChildTableName;
        public bool bSingularCollects;
        public string sCategoryName;
        public CAssocArray asaV;
        public Node CurChild;
        public string NewKey;
        public TableDef tdfNew;
        public Relation relNew;
        public Index idxNew;
        public CAssocArray TreeReading;
        public Node CurrNode;
        public string sNodesLeft;
        public int CurrFav;
        public CAssocItem CurrAssoc;
        public TableDef CurTable;
        public object nodX;
        public int TableListCount;
        public CAssocArray asaX;
        public CAssocItem CurrItem;
        public string sCategoryChosen;
        public string sCategoryChosen;
        public object sDatabasePath;
        public string sNewDatabaseName;
        public ListItem ItemClicked;
        public string sParentTable;
        public object sTable;
        public object sField;
        public object CurTable;
        public object CurItem;
        public object NewField;
        public string sTable;
        public string sTableName;
        public Node NodeDroppedOnto;
        public Node NodeClicked;
        public TableDef CurTable;
        public Field CurField;
        public object nodX;
        public object litX;
        public object sIcon;


                public bool Canceled
    {
        get
        {
        Canceled = mbCanceled;
        }

    }


                public string Connectstring
    {
        get
        {
        Connectstring = msClassDatabaseOptions;
        }

    }


                public string DBName
    {
        get
        {
        DBName = sGetToken(sGetToken(msClassDatabaseName, lTokenCount(msClassDatabaseName, gsBS), gsBS), 1, ".mdb");
        }

    }


                public string DBPathAndFilename
    {
        get
        {
        DBPathAndFilename = msClassDatabaseName;
        }

    }


                public string ODBCDatabaseName
    {
        get
        {
        ODBCDatabaseName = sGetToken(sGetToken(msClassDatabaseOptions, 2, "DSN="), 1, gsSC);
        }

    }


                public bool GenerateBranch
    {
        get
        {
        GenerateBranch = mbGenerateBranch;
        }

    }


                public bool GenerateDatabase
    {
        get
        {
        GenerateDatabase = mbGenerateDatabase;
        }

    }



            public void GenerateChildren            {
                CurChild = nodChild;
                Do Until CurChild Is null;
                CurChild.Selected = true;
                tvwTables_NodeClick CurChild;
                if ( ! CurChild.Parent.Parent Is null )
            {;
                asaPass("Parent Table Name") = CurChild.Parent.Text;
                }
            else
            {;
                asaPass("Parent Table Name") = "Root";
            }

            public void GenerateClass            {


                try
{;

                sCategoryName = sGetToken(sDataLibraryType, 1, gsCategoryTemplateDelimiter);

                //    ' Determine if the singular member of the collection will be collecting anything;
                bSingularCollects = (! tvwTables.SelectedItem.Child Is null);
                if ( bSingularCollects )
            {;
                sClassToCollect = tvwTables.SelectedItem.Child.Text + "s";
                sChildTableName = tvwTables.SelectedItem.Child.Text;
                }
            else
            {;
                sClassToCollect = string.Empty;
                sChildTableName = tvwTables.SelectedItem.Child.Text;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void TriggerClassGeneration()
            {
                if ( gbProcessing )
            {
 return;


                asaV = new CAssocArray();

                if ( Canceled = false )
            {;
                //        'masaMisc.this.Clear()                                                                ' Clean out the assoc used for 'session' long inserts;
                if ( GenerateDatabase = false )
            {;
                //            ' Generate a class for the currently selected class;
                GenerateClass asaV, cboDataLibraryType.Text + gsCategoryTemplateDelimiter, tvwTables, lvwFields;
                if ( GenerateBranch = true )
            {;
                //                ' Generate a class for each child table and each of its children tables;
                if ( ! tvwTables.SelectedItem.Child Is null )
            {;
                GenerateChildren asaV, cboDataLibraryType.Text + gsCategoryTemplateDelimiter, tvwTables.SelectedItem.Child;
            }

            public void AddFavorite()
            {
                if ( Len(msClassDatabaseName) )
            {;
                Favorites(msClassDatabaseName) = "|||" + sNormalize(TableList);
                }
            else
            {;
                if ( Len(ODBCTableNamePrefix) )
            {;
                Favorites("ODBC: " + sGetToken(sGetToken(msClassDatabaseOptions, 2, "DSN="), 1, gsSC) + " (" + ODBCTableNamePrefix + gsPC + IIf(Len(TableList), " Limited (" + tvwTables.Nodes.Count - 1 + gsPC, string.Empty)) = msClassDatabaseOptions + "|" + ODBCTableNamePrefix + "|" + ODBCPassword + "|" + sNormalize(TableList);
                }
            else
            {;
                Favorites("ODBC: " + sGetToken(sGetToken(msClassDatabaseOptions, 2, "DSN="), 1, gsSC) + IIf(Len(TableList), " Limited (" + tvwTables.Nodes.Count - 1 + gsPC, string.Empty)) = msClassDatabaseOptions + "|" + ODBCTableNamePrefix + "|" + ODBCPassword + "|" + sNormalize(TableList);
            }

            public void AddTable            {
                // TODO: Rewrite try/catch and/or goto. EH_frmDBClassGen_AddTable;

                if ( IsInODBCDatabaseMode )
            {;
                if ( ! KeepDatabaseOpen )
            {
  db == OpenDatabase(msClassDatabaseName, dbDriverComplete, false, msClassDatabaseOptions);
                }
            else
            {;
                db = OpenDatabase(msClassDatabaseName, false, false);
            }

            public void RefreshTableList()
            {
                TreeReading = new CAssocArray();
                TreeReading.TreeToAll tvwTables;
                TableList = TreeReading.All;
                TreeReading = null;
            }

            public As RemoveNodes            {
                RemoveNodes_Restart:;
                foreach( var CurrNode in tvwX.Nodes );
                if ( StrComp(CurrNode.Image), UCase$(sNodeImageNameToRemove)) = 0 .ToUpper()
            {;
                tvwX.Nodes.Remove CurrNode.Key;
                GoTo RemoveNodes_Restart;
                }
            else
            {if ( ! CurrNode.Parent Is null )
            {;
                sNodesLeft +=  CurrNode.Key + gsSC;
            }

            public void UpdateFavorites()
            {
                try
{;

                if ( FavoriteCount > 0 )
            {
                         ' this.Clear()out previous entries;
                for(var CurrFav = FavoriteCount; CurrFav < 1 Step -1; CurrFav++)  {;
                Unload mnuFavorite(CurrFav);
                } // CurrFav;
                mnuFavorite(0).Text = "-Empty-";
                mnuFavorite(0).Enabled = false;
                FavoriteCount = 0;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void RefreshCategories()
            {
                try
{;
                mbLoadingCategories = true;
                Parent.SliceAndDice.Categorys.FillList cboDataLibraryType, 1;
                cboDataLibraryType.ListIndex = FindListIndex(cboDataLibraryType, GetSetting$(App.ProductName, "DB Class Gen", "Last " + gsCategory, "RDO Persisted"));
                mbLoadingCategories = false;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void PopulateTree()
            {
                // TODO: Rewrite try/catch and/or goto. EH_PopulateTree;

                //    'On Error goto Next;
                //    'msClassDatabaseName = sGetToken(sGetToken(msClassDatabaseName, lTokenCount(msClassDatabaseName, gsBS), gsBS), 1, ".mdb");

                Screen.MousePointer = vbHourglass;
                if ( Len(msClassDatabaseName) = 0 )
            {;
                if ( ! KeepDatabaseOpen )
            {;
                if ( Len(msClassDatabaseOptions) = 0 )
            {;
                msClassDatabaseOptions = "ODBC;";
            }

            public As sFieldType            {
                switch iFieldType;
                Case dbBigInt: sFieldType = "Big Integer";
                Case dbBinary: sFieldType = "Binary";
                Case dbBoolean: sFieldType = "Boolean";
                Case dbByte: sFieldType = "Byte";
                Case dbChar: sFieldType = "Char";
                Case dbCurrency: sFieldType = "Currency";
                Case dbDate: sFieldType = "Date / Time";
                Case dbDecimal: sFieldType = "Decimal";
                Case dbDouble: sFieldType = "Double";
                Case dbFloat: sFieldType = "Float";
                Case dbGUID: sFieldType = "Guid";
                Case dbInteger: sFieldType = "Integer";
                Case dbLong: sFieldType = "Long";
                Case dbLongBinary: sFieldType = "long Binary (OLE Object)";
                Case dbMemo: sFieldType = "Memo";
                Case dbNumeric: sFieldType = "Numeric";
                Case dbSingle: sFieldType = "Single";
                Case dbText: sFieldType = "Text";
                Case dbTime: sFieldType = "Time";
                Case dbTimeStamp: sFieldType = "Time Stamp";
                Case dbVarBinary: sFieldType = "VarBinary";
            }

            public void cboDataLibraryType_Click()
            {
                if ( mbLoadingCategories )
            {
 return;
                SaveSetting App.ProductName, "DB Class Gen", "Last " + gsCategory, cboDataLibraryType.Text;
            }

            public void cmdAddCategory_Click()
            {

                sCategoryChosen = Parent.SliceAndDice.Categorys.Choose(0);
                if ( Len(sCategoryChosen) > 0 )
            {;
                Parent.SliceAndDice.Categorys.Item(sCategoryChosen).CategoryType = 1;
                Parent.SliceAndDice.Save;
            }

            public void cmdDeleteCategory_Click()
            {

                sCategoryChosen = Parent.SliceAndDice.Categorys.Choose(1);
                if ( Len(sCategoryChosen) > 0 )
            {;
                Parent.SliceAndDice.Categorys.Item(sCategoryChosen).CategoryType = 0;
                Parent.SliceAndDice.Save;
            }

            public void Form_Initialize()
            {

            }

            public void Form_QueryUnload            {
                if ( UnloadMode = vbFormControlMenu )
            {;
                Cancel = true;
            }

            public void Form_Resize()
            {
                try
{;
                if ( ! dvwTable.Visible )
            {;
                lvwFields.Move ScaleWidth - fraCategory.Width, lvwFields.Top, fraCategory.Width, ScaleHeight - lvwFields.Top;
                tvwTables.Move 0, 60, ScaleWidth - fraCategory.Width, ScaleHeight;
                fraCategory.Move lvwFields.Left, 0;
                }
            else
            {;
                lvwFields.Move ScaleWidth - fraCategory.Width, lvwFields.Top, fraCategory.Width, ScaleHeight - fraCategory.Height - dvwTable.Height;
                tvwTables.Move 0, 60, ScaleWidth - fraCategory.Width, ScaleHeight - dvwTable.Height;
                fraCategory.Move lvwFields.Left, 0;
                dvwTable.Move 0, lvwFields.Top + lvwFields.Height, ScaleWidth;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Form_Unload            {
                SaveFormPosition Me;
            }

            public void lvwFields_DblClick()
            {
                mnuFieldNew_Click;
            }

            public void lvwFields_KeyPress            {
                if ( KeyAscii = 13 )
            {;
                mnuFieldNew_Click;
            }

            public void mnuFavorite_Click            {
                Screen.MousePointer = vbDefault;

                if ( Len(sGetToken(.Value, 1, "|")) = 0 )
            {;
                msClassDatabaseName = sGetToken(.Key, 1, "|");
                msClassDatabaseOptions = string.Empty;
                ODBCTableNamePrefix = string.Empty;
                ODBCPassword = string.Empty;
                TableList = sDenormalize(sAfter(.Value, 3, "|"));
                }
            else
            {;
                msClassDatabaseName = string.Empty;
                msClassDatabaseOptions = sGetToken(.Value, 1, "|");
                ODBCTableNamePrefix = sGetToken(.Value, 2, "|");
                ODBCPassword = sGetToken(.Value, 3, "|");
                TableList = sDenormalize(sAfter(.Value, 3, "|"));
            }

            public void mnuFavRemoveAll_Click()
            {
                if ( bUserSure() )
            {;
                Favorites.All = string.Empty;
                SaveSetting App.ProductName, "DB Class Gen", "Favorites", string.Empty;
                UpdateFavorites;
            }

            public void mnuFileExit_Click()
            {
                mbCanceled = true;
                lvwFields.ListItems.Clear;
                SaveFormPosition Me;
                Hide;
            }

            public void mnuFileNew_Click()
            {

                sDatabasePath = Trim$(BrowseForFolder(hwnd, "Where should database go ?"));
                if ( Len(sDatabasePath) == 0 )
            {
 return;

                sNewDatabaseName = Trim$(InputBox("What should the name of the new database be ?", "CREATE BLANK DATABASE"));
                if ( Len(snewDatabaseName) == 0 )
            {
 return;

                sDatabasePath.Substring(sDatabasePath.Length - 1) <> gsBS )
            {
 sDatabasePath == sDatabasePath + gsBS;
                LCase$(snewDatabaseName).Substring(LCase$(snewDatabaseName).Length - 4) <> ".mdb" )
            {
 snewDatabaseName == sDatabasePath + snewDatabaseName + ".mdb";

                try
{;
                db = CreateDatabase(sNewDatabaseName, dbLangGeneral, dbVersion30);
                db.Close;

                msClassDatabaseName = sNewDatabaseName;
                msClassDatabaseOptions = string.Empty;
                ODBCTableNamePrefix = string.Empty;
                ODBCPassword = string.Empty;
                IsInODBCDatabaseMode = false;
                KeepDatabaseOpen = false;
                PopulateTree;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuFileOpenODBC_Click()
            {
                mbOpenVBIDE = false;
                msClassDatabaseName = string.Empty;
                msClassDatabaseOptions = string.Empty;
                ODBCTableNamePrefix = string.Empty;
                ODBCPassword = string.Empty;
                IsInODBCDatabaseMode = false;
                KeepDatabaseOpen = false;
                PopulateTree;
            }

            public void mnuFileOpenVBIDE_Click()
            {
                mbOpenVBIDE = true;
                msClassDatabaseName = string.Empty;
                msClassDatabaseOptions = string.Empty;
                ODBCTableNamePrefix = string.Empty;
                ODBCPassword = string.Empty;
                IsInODBCDatabaseMode = false;
                KeepDatabaseOpen = false;
                PopulateTree;
            }

            public void mnuFreeAssociateTables_Click()
            {
                mnuFreeAssociateTables.Checked = ! mnuFreeAssociateTables.Checked;
                SaveSetting App.ProductName, "DB Class Gen", "Free Associate Tables", mnuFreeAssociateTables.Checked;
            }

            public void mnuGenerateClass_Click()
            {
                if ( lvwFields.ListItems.Count = 0 )
            {;
                MsgBox "Please select a table first.";
                return;
            }

            public void mnuGenerateEnterBranch_Click()
            {
                if ( lvwFields.ListItems.Count = 0 )
            {;
                MsgBox "Please select a branch first.";
                return;
            }

            public void mnuGenerateEntireDatabase_Click()
            {
                if ( tvwTables.Nodes.Count = 0 )
            {;
                MsgBox "Please select a database first.";
                return;
            }

            public void mnuFileOpen_Click()
            {
                mbOpenVBIDE = false;
                msClassDatabaseName = Parent.sChooseDatabase();
                msClassDatabaseOptions = string.Empty;
                ODBCTableNamePrefix = string.Empty;
                ODBCPassword = string.Empty;
                IsInODBCDatabaseMode = false;
                KeepDatabaseOpen = false;
                if ( Len(msClassDatabaseName) > 0 )
            {;
                PopulateTree;
            }

            public void Form_Load()
            {
                try
{;

                RefreshCategories;

                mnuRulesAutoAddKey.Checked = GetSetting(App.ProductName, "DB Class Gen", "AutoAddKey", true);
                mnuRulesAutoAddDateModified.Checked = GetSetting(App.ProductName, "DB Class Gen", "AutoAddDateModified", true);
                mnuRulesAutoAddDateCreated.Checked = GetSetting(App.ProductName, "DB Class Gen", "AutoAddCreated", true);

                mnuRulesEnforce.Checked = GetSetting(App.ProductName, "DB Class Gen", "Enforce", true);
                mnuRulesCascadeUpdates.Checked = GetSetting(App.ProductName, "DB Class Gen", "CascadeUpdates", true);
                mnuRulesCascadeDeletes.Checked = GetSetting(App.ProductName, "DB Class Gen", "CascadeDeletes", true);

                mnuRulesUseAutoNumber.Checked = GetSetting(App.ProductName, "DB Class Gen", "UseAutoNumber", true);
                mnuViewTableData.Checked = GetSetting(App.ProductName, "DB Class Gen", "ViewTableData", false);
                mnuRelateOnLoad.Checked = GetSetting(App.ProductName, "DB Class Gen", "Relate on load", true);
                mnuFreeAssociateTables.Checked = GetSetting(App.ProductName, "DB Class Gen", "Free Asoociate Tables", true);

                if ( mnuRulesEnforce.Checked )
            {;
                mnuRulesCascadeUpdates.Enabled = true;
                mnuRulesCascadeDeletes.Enabled = true;
                }
            else
            {;
                mnuRulesCascadeUpdates.Enabled = false;
                mnuRulesCascadeDeletes.Enabled = false;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Form_Terminate()
            {
                SaveFormPosition Me;
                Favorites = null;
                Parent = null;
                //    ' LogEvent "frmDBClassGen: Terminate";
            }

            public void lvwFields_MouseUp            {
                try
{;

                if ( Button = vbRightButton )
            {;
                ItemClicked = lvwFields.HitTest(X, Y);
                if ( ! ItemClicked Is null )
            {;
                ItemClicked.Selected = true;
                PopupMenu mnuField;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuFieldDelete_Click()
            {

                try
{;
                if ( tvwTables.SelectedItem.Parent.Key = msClassDatabaseName )
            {;
                sParentTable = string.Empty;
                }
            else
            {;
                sParentTable = tvwTables.SelectedItem.Parent.Key;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuFieldNew_Click()
            {

                NewField = new frmFieldType();

                mnuFieldNew_Click_TryAgain:;

                try
{;

                NewField.Show vbModal, Me;
                if ( ex <> 0 )
            {
 return;
                if ( newField.Canceled == true )
            {
 return;

                foreach( var CurItem in lvwFields.ListItems );
                if ( CurItem.Text) = UCase$(.FieldName) .ToUpper()
            {;
                MsgBox "That field already exists in this table... Try again.";
                GoTo mnuFieldNew_Click_TryAgain;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuRelateOnLoad_Click()
            {
                mnuRelateOnLoad.Checked = ! mnuRelateOnLoad.Checked;
                SaveSetting App.ProductName, "DB Class Gen", "Relate on load", mnuRelateOnLoad.Checked;
            }

            public void mnuRemoveMarked_Click()
            {
                RemoveNodes tvwTables, "TableMarked";
                RefreshTableList;
                AddFavorite;
            }

            public void mnuRemoveUnmarked_Click()
            {
                RemoveNodes tvwTables, "Table";
                RefreshTableList;
                AddFavorite;
            }

            public void mnuRulesAutoAddDateCreated_Click()
            {
                mnuRulesAutoAddDateCreated.Checked = ! mnuRulesAutoAddDateCreated.Checked;
                SaveSetting App.ProductName, "DB Class Gen", "AutoAddDateCreated", mnuRulesAutoAddDateCreated.Checked;
            }

            public void mnuRulesAutoAddDateModified_Click()
            {
                mnuRulesAutoAddDateModified.Checked = ! mnuRulesAutoAddDateModified.Checked;
                SaveSetting App.ProductName, "DB Class Gen", "AutoAddDateModified", mnuRulesAutoAddDateModified.Checked;
            }

            public void mnuRulesAutoAddKey_Click()
            {
                mnuRulesAutoAddKey.Checked = ! mnuRulesAutoAddKey.Checked;
                SaveSetting App.ProductName, "DB Class Gen", "AutoAddKey", mnuRulesAutoAddKey.Checked;
            }

            public void mnuRulesCascadeDeletes_Click()
            {
                mnuRulesCascadeDeletes.Checked = ! mnuRulesCascadeDeletes.Checked;
                SaveSetting App.ProductName, "DB Class Gen", "CascadeDeletes", mnuRulesCascadeDeletes.Checked;
            }

            public void mnuRulesCascadeUpdates_Click()
            {
                mnuRulesCascadeUpdates.Checked = ! mnuRulesCascadeUpdates.Checked;
                SaveSetting App.ProductName, "DB Class Gen", "CascadeUpdates", mnuRulesCascadeUpdates.Checked;
            }

            public void mnuRulesEnforce_Click()
            {
                mnuRulesEnforce.Checked = ! mnuRulesEnforce.Checked;
                SaveSetting App.ProductName, "DB Class Gen", "Enforce", mnuRulesEnforce.Checked;
                if ( mnuRulesEnforce.Checked )
            {;
                mnuRulesCascadeUpdates.Enabled = true;
                mnuRulesCascadeDeletes.Enabled = true;
                }
            else
            {;
                mnuRulesCascadeUpdates.Enabled = false;
                mnuRulesCascadeDeletes.Enabled = false;
            }

            public void mnuRulesUseAutoNumber_Click()
            {
                mnuRulesUseAutoNumber.Checked = ! mnuRulesUseAutoNumber.Checked;
                SaveSetting App.ProductName, "DB Class Gen", "UseAutoNumber", mnuRulesUseAutoNumber.Checked;
            }

            public void mnuShowAllTables_Click()
            {
                TableList = string.Empty;
                AddFavorite;
                PopulateTree;
            }

            public void mnuTableDelete_Click()
            {
                try
{;
                if ( bUserSure("This will PERMANENTLY remove the table selected." + gsEolTab + "ARE YOU ABSOLUTELY SURE ?") )
            {;
                if ( IsInODBCDatabaseMode )
            {;
                if ( ! KeepDatabaseOpen )
            {
  db == OpenDatabase(msClassDatabaseName, dbDriverComplete, false, msClassDatabaseOptions);
                }
            else
            {;
                db = OpenDatabase(msClassDatabaseName, false, false);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuTableNew_Click()
            {

                mnuTableNew_Click_TryAgain:;
                sTable = Replace(Trim$(InputBox("What should the name of the new table be ?" + gsEolTab + "Note: Table names should be singular, such as:" + gsEolTab + "Book, Publisher, etc.")), gsS, string.Empty);
                if ( Len(sTable) == 0 )
            {
 return;

                sTable.Substring(sTable.Length - 1) = "s" )
            {;
                MsgBox "Table names MUST be singular.";
                GoTo mnuTableNew_Click_TryAgain;
            }

            public void mnuToggleTableMark_Click()
            {
                if ( tvwTables.SelectedItem.Image = "Table" )
            {;
                tvwTables.SelectedItem.Image = "TableMarked";
                tvwTables.SelectedItem.ExpandedImage = "TableMarked";
                tvwTables.SelectedItem.SelectedImage = "TableMarked";
                }
            else
            {;
                tvwTables.SelectedItem.Image = "Table";
                tvwTables.SelectedItem.ExpandedImage = "Table";
                tvwTables.SelectedItem.SelectedImage = "Table";
            }

            public void mnuUnhideTable_Click()
            {
                try
{;
                sTableName = InputBox("What is the fully qualified name of the table to unhide ?", "UNHIDE A TABLE");
                if ( Len(sTableName) )
            {;

                tvwTables.Nodes.Add(                                                                                                                                                                                                      tvwTables.SelectedItem.Key, tvwChild, sTableName, sTableName, "Table", "Table").ExpandedImage = "Table";
                tvwTables.Nodes.Add(                                                                                                                                                                                                      tvwTables.SelectedItem.Key, tvwChild, sTableName, sTableName, "Table", "Table").Expanded = true;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuViewTableData_Click()
            {
                try
{;
                mnuViewTableData.Checked = ! mnuViewTableData.Checked;
                SaveSetting App.ProductName, "DB Class Gen", "ViewTableData", mnuViewTableData.Checked;
                dvwTable.Visible = mnuViewTableData.Checked;
                Form_Resize;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuX_Click()
            {
                mnuFileExit_Click;
            }

            public void tvwTables_DblClick()
            {
                mnuTableNew_Click;
            }

            public void tvwTables_DragDrop            {
                try
{;

                if ( Source.Tag = "tvwTables" )
            {;
                if ( ! NodeDragged Is null )
            {;
                NodeDroppedOnto = null;
                NodeDroppedOnto = tvwTables.HitTest(X, Y);
                if ( ! NodeDroppedOnto Is null )
            {;
                NodeDragged.Parent = NodeDroppedOnto;
                RefreshTableList;
                AddFavorite;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void tvwTables_KeyPress            {
                if ( KeyAscii = 13 )
            {;
                mnuTableNew_Click;
            }

            public void tvwTables_MouseMove            {
                if ( Button = vbLeftButton And (Shift And vbShiftMask) <> 0 )
            {;
                NodeDragged = null;
                NodeDragged = tvwTables.HitTest(X, Y);
                if ( ! NodeDragged Is null )
            {;
                tvwTables.Drag vbBeginDrag;
            }

            public void tvwTables_MouseUp            {

                if ( Button = vbRightButton )
            {;
                NodeClicked = tvwTables.HitTest(X, Y);
                if ( ! NodeClicked Is null )
            {;
                NodeClicked.Selected = true;
                PopupMenu mnuTable;
            }

            public void tvwTables_NodeClick            {

                try
{;
                Screen.MousePointer = vbHourglass;
                if ( IsInODBCDatabaseMode )
            {;
                if ( ! KeepDatabaseOpen )
            {
  db == OpenDatabase(msClassDatabaseName, dbDriverComplete, false, msClassDatabaseOptions);
                }
            else
            {;
                db = OpenDatabase(msClassDatabaseName, false, false);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

        }
    }
