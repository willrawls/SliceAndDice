using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class frmODBCClassGen : System.Windows.Forms.Form
        {
         public System.Windows.Forms.VB.Frame Frame1;
         public System.Windows.Forms.VB.CommandButton cmdDeleteCategory;
         public System.Windows.Forms.VB.CommandButton cmdAddCategory;
         public System.Windows.Forms.VB.ComboBox cboDataLibraryType;
         public System.Windows.Forms.ComctlLib.ListView lvwFields;
         public System.Windows.Forms.ComctlLib.TreeView tvwTables;
         public System.Windows.Forms.ComctlLib.ImageList imlDB;
         public System.Windows.Forms.VB.Menu mnuX;
         public System.Windows.Forms.VB.Menu mnuFile;
         public System.Windows.Forms.VB.Menu mnuFileOpen;
         public System.Windows.Forms.VB.Menu mnuFileNew;
         public System.Windows.Forms.VB.Menu mnuFileOpen2;
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

        public frmODBCClassGen()
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
            this.Frame1 = new System.Windows.Forms.VB.Frame();
            this.cmdDeleteCategory = new System.Windows.Forms.VB.CommandButton();
            this.cmdAddCategory = new System.Windows.Forms.VB.CommandButton();
            this.cboDataLibraryType = new System.Windows.Forms.VB.ComboBox();
            this.lvwFields = new System.Windows.Forms.ComctlLib.ListView();
            this.tvwTables = new System.Windows.Forms.ComctlLib.TreeView();
            this.imlDB = new System.Windows.Forms.ComctlLib.ImageList();
            this.mnuX = new System.Windows.Forms.VB.Menu();
            this.mnuFile = new System.Windows.Forms.VB.Menu();
            this.mnuFileOpen = new System.Windows.Forms.VB.Menu();
            this.mnuFileNew = new System.Windows.Forms.VB.Menu();
            this.mnuFileOpen2 = new System.Windows.Forms.VB.Menu();
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
            this.mnuSep0 = new System.Windows.Forms.VB.Menu();
            this.mnuTableDelete = new System.Windows.Forms.VB.Menu();
            this.mnuField = new System.Windows.Forms.VB.Menu();
            this.mnuFieldNew = new System.Windows.Forms.VB.Menu();
            this.mnuFieldRename = new System.Windows.Forms.VB.Menu();
            this.mnuSep1 = new System.Windows.Forms.VB.Menu();
            this.mnuFieldDelete = new System.Windows.Forms.VB.Menu();
            this.SuspendLayout();
            this.Frame1.SuspendLayout();
            this.mnuFile.SuspendLayout();
            this.mnuGenerate.SuspendLayout();
            this.mnuRules.SuspendLayout();
            this.mnuRulesAutoAdd.SuspendLayout();
            this.mnuRelationship.SuspendLayout();
            this.mnuHelp.SuspendLayout();
            this.mnuShortcut.SuspendLayout();
            this.mnuTable.SuspendLayout();
            this.mnuField.SuspendLayout();
            //
            // Frame1
            //
            this.Frame1.Name = "Frame1";
            this.Frame1.Text = " Code to Generate ";
            this.Frame1.Size = new System.Drawing.Size(317, 44);
            this.Frame1.Location = new System.Drawing.Point(270, -1);
            this.Frame1.TabIndex = 2;
            this.Frame1.Controls.AddRange(new System.Windows.Forms.Control[]
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
//            this.cboDataLibraryType.ItemData = "frmODBCClassGen.frx":0000;
            this.cboDataLibraryType.Location = new System.Drawing.Point(9, 16);
//            this.cboDataLibraryType.List = "frmODBCClassGen.frx":0002;
            this.cboDataLibraryType.TabIndex = 3;
            this.cboDataLibraryType.Text = "RDO Persisted";
            //
            // lvwFields
            //
            this.lvwFields.Name = "lvwFields";
            this.lvwFields.Size = new System.Drawing.Size(319, 354);
            this.lvwFields.Location = new System.Drawing.Point(270, 47);
            this.lvwFields.TabIndex = 0;
            this.lvwFields.LabelEdit = true;
            this.lvwFields.LabelWrap = true;
            this.lvwFields.HideSelection = true;
//            this.lvwFields.Icons = "imlDB";
//            this.lvwFields.SmallIcons = "imlDB";
            this.lvwFields.ForeColor = System.Drawing.Color.FromArgb(-2147483640);
            this.lvwFields.BackColor = System.Drawing.Color.FromArgb(-2147483643);
            this.lvwFields.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
//            this.lvwFields.NumItems = 4;
//            this.lvwFields.ColumnHeader(1) = ;
//            this.lvwFields.ColumnHeader(2) = ;
//            this.lvwFields.ColumnHeader(3) = ;
//            this.lvwFields.ColumnHeader(4) = ;
            //
            // tvwTables
            //
            this.tvwTables.Name = "tvwTables";
//            this.tvwTables.DragIcon = "frmODBCClassGen.frx":0004;
            this.tvwTables.Size = new System.Drawing.Size(262, 396);
            this.tvwTables.Location = new System.Drawing.Point(4, 5);
            this.tvwTables.TabIndex = 1;
//            this.tvwTables.Indentation = 265;
            this.tvwTables.LabelEdit = true;
//            this.tvwTables.ImageList = "imlDB";
            //
            // imlDB
            //
            this.imlDB.Name = "imlDB";
            this.imlDB.Location = new System.Drawing.Point(26, -15);
            this.imlDB.BackColor = System.Drawing.Color.FromArgb(-2147483643);
//            this.imlDB.ImageWidth = 16;
//            this.imlDB.ImageHeight = 16;
//            this.imlDB.MaskColor = 12632256;
//            this.imlDB.ListImage1 = ;
//            this.imlDB.ListImage2 = ;
//            this.imlDB.ListImage3 = ;
//            this.imlDB.ListImage4 = ;
//            this.imlDB.ListImage5 = ;
//            this.imlDB.ListImage6 = ;
//            this.imlDB.ListImage6 = ;
            //
            // mnuX
            //
            this.mnuX.Name = "mnuX";
            this.mnuX.Text = "X";
            //
            // mnuFile
            //
            this.mnuFile.Name = "mnuFile";
            this.mnuFile.Text = "&File";
            this.mnuFile.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuFileOpen,
                        this.mnuFileNew,
                        this.mnuFileOpen2,
                        this.mnuFileSep0,
                        this.mnuFileExit
            });
            //
            // mnuFileOpen
            //
            this.mnuFileOpen.Name = "mnuFileOpen";
            this.mnuFileOpen.Text = "&Open an Access97 database";
            //
            // mnuFileNew
            //
            this.mnuFileNew.Name = "mnuFileNew";
            this.mnuFileNew.Text = "Create a &New Access97 database";
            //
            // mnuFileOpen2
            //
            this.mnuFileOpen2.Name = "mnuFileOpen2";
            this.mnuFileOpen2.Text = "Open an ODBC data source";
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
            this.mnuGenerateEntireDatabase.Text = "Enter &Database                 (everything and a wrapper class)";
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
            this.mnuShortcut.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuTable,
                        this.mnuField
            });
            //
            // mnuTable
            //
            this.mnuTable.Name = "mnuTable";
            this.mnuTable.Text = "Table Right Click";
            this.mnuTable.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuTableNew,
                        this.mnuTableRename,
                        this.mnuSep0,
                        this.mnuTableDelete
            });
            //
            // mnuTableNew
            //
            this.mnuTableNew.Name = "mnuTableNew";
            this.mnuTableNew.Text = "New Table";
            //
            // mnuTableRename
            //
            this.mnuTableRename.Name = "mnuTableRename";
            this.mnuTableRename.Text = "Rename";
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
            this.mnuField.Text = "Field Right Click";
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
            // frmODBCClassGen
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.Frame1,
                  this.lvwFields,
                  this.tvwTables,
                  this.imlDB,
                  this.mnuX,
                  this.mnuFile,
                  this.mnuGenerate,
                  this.mnuRules,
                  this.mnuHelp,
                  this.mnuShortcut
            });
            this.Name = "frmODBCClassGen";
            this.Frame1.ResumeLayout(false);
            this.mnuFile.ResumeLayout(false);
            this.mnuGenerate.ResumeLayout(false);
            this.mnuRules.ResumeLayout(false);
            this.mnuRulesAutoAdd.ResumeLayout(false);
            this.mnuRelationship.ResumeLayout(false);
            this.mnuHelp.ResumeLayout(false);
            this.mnuShortcut.ResumeLayout(false);
            this.mnuTable.ResumeLayout(false);
            this.mnuField.ResumeLayout(false);
            this.ResumeLayout(false);
        }
        #endregion

        public bool m_bCanceled;
        public bool m_bGenerateDatabase;
        public bool m_bGenerateBranch;
        public string m_sClassDatabaseName;
        public frmMain Parent;


                public bool Canceled
    {
        get
        {
        Canceled = m_bCanceled;
        }

    }


                public string DBName
    {
        get
        {
        DBName = sGetToken(sGetToken(m_sClassDatabaseName, lTokenCount(m_sClassDatabaseName, "\"), "\"), 1, ".mdb");
        }

    }


                public string DBPathAndFilename
    {
        get
        {
        DBPathAndFilename = m_sClassDatabaseName;
        }

    }


                public bool GenerateBranch
    {
        get
        {
        GenerateBranch = m_bGenerateBranch;
        }

    }


                public bool GenerateDatabase
    {
        get
        {
        GenerateDatabase = m_bGenerateDatabase;
        }

    }



            public void AddTable            {
                // TODO: Rewrite try/catch and/or goto. EH_frmODBCClassGen_AddTable;
                ;
                ;
                ;
                ;

                db = OpenDatabase(m_sClassDatabaseName, "ODBC;");
                tdfnew = db.CreateTableDef(sTableName);

                //           ' Create the primary unique index field;
                tdfNew.Fields.Append tdfNew.CreateField(sTableName + "ID", dbLong);
                if ( mnuRulesUseAutoNumber.Checked )
            {;
                tdfNew.Fields(sTableName + "ID").Attributes = (.Fields(sTableName + "ID").Attributes And dbAutoIncrField);
                };
                idxnew = tdfNew.CreateIndex(sTableName + "IDIndex");

                idxNew.Fields.Append idxNew.CreateField(sTableName + "ID");
                idxNew.Unique = true;
                idxNew.Primary = true;
                idxNew.Required = true;
                ;

                tdfNew.Indexes.Append idxNew;

                if ( Len(sParentTable) > 0 )
            {;
                //              ' Attach to any indicated parent;
                tdfNew.Fields.Append tdfNew.CreateField(sParentTable + "ID", dbLong);
                idxnew = tdfNew.CreateIndex(sParentTable + "IDIndex");

                idxNew.Fields.Append idxNew.CreateField(sParentTable + "ID");
                idxNew.Required = true;

                tdfNew.Indexes.Append idxNew;
                };
                idxnew = null();

                if ( mnuRulesAutoAddKey.Checked )
            {;
                tdfNew.Fields.Append tdfNew.CreateField(sTableName + "Name", dbText);
                };
                if ( mnuRulesAutoAddDateCreated.Checked )
            {;
                tdfNew.Fields.Append tdfNew.CreateField("DateCreated", dbDate);
                };
                if ( mnuRulesAutoAddDateModified.Checked )
            {;
                tdfNew.Fields.Append tdfNew.CreateField("DateModified", dbDate);
                };

                db.TableDefs.Append tdfNew;

                if ( Len(sParentTable) > 0 )
            {;
                //         ' Create the needed cascading update/delete relationship between the parent table and child table;
                if ( mnuRulesEnforce.Checked )
            {;
                if ( mnuRulesCascadeUpdates.Checked )
            {;
                if ( mnuRulesCascadeDeletes.Checked )
            {;
                relnew = db.CreateRelation(sParentTable + "_" + sTableName, sParentTable, sTableName, dbRelationUpdateCascade + dbRelationDeleteCascade);
                }
            else
            {;
                relnew = db.CreateRelation(sParentTable + "_" + sTableName, sParentTable, sTableName, dbRelationUpdateCascade);
                };
                }
            else
            {;
                if ( mnuRulesCascadeDeletes.Checked )
            {;
                relnew = db.CreateRelation(sParentTable + "_" + sTableName, sParentTable, sTableName, dbRelationDeleteCascade);
                }
            else
            {;
                relnew = db.CreateRelation(sParentTable + "_" + sTableName, sParentTable, sTableName);
                };
                };
                relNew.Fields.Append relNew.CreateField(sParentTable + "ID");
                relNew.Fields(sParentTable + "ID").ForeignName = sParentTable + "ID";
                db.Relations.Append relNew;
                relnew = null();
                };
                };
                ;
                tdfnew = null();
                db.Close;
                ;
                PopulateTree;

                EH_frmODBCClassGen_AddTable_Continue:;
                return;

                EH_frmODBCClassGen_AddTable:;
                MsgBox "Error in SliceAndDice.frmODBCClassGen_AddTable" + Chr(13) + Chr(13) + Chr(9) + Err.Description;
                goto EH_frmODBCClassGen_AddTable_Continue;
                ;
                Resume;
            }

            public void LoadCategories()
            {
                try
{;
                Parent.SliceAndDice.Categorys.FillList cboDataLibraryType, 1;
                cboDataLibraryType.ListIndex = FindListIndex(cboDataLibraryType, GetSetting(App.ProductName, "Last", "DPCCG Code to generate", "RDO Persisted"));
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void PopulateTree()
            {
                ;
                ;
                ;

                try
{;
                //   'm_sClassDatabaseName = sGetToken(sGetToken(m_sClassDatabaseName, lTokenCount(m_sClassDatabaseName, "\"), "\"), 1, ".mdb");

                db = OpenDatabase(m_sClassDatabaseName, m_sClassDatabaseOptions);
                lvwFields.ListItems.Clear;

                tvwTables.Nodes.Clear;
                nodX = tvwTables.Nodes.Add(                                                                                                                                                                                                      , , "Root", sGetToken(sGetToken(m_sClassDatabaseName, lTokenCount(m_sClassDatabaseName, "\"), "\"), 1, ".mdb"), "Database", "Database");
                nodX.ExpandedImage = "Database";
                nodX.Expanded = true;

                foreach( var CurTable in db.TableDefs;
                CurTable.Name.Substring(0, 4) <> "MSys" )
            {;
                nodX = tvwTables.Nodes.Add(                                                                                                                                                                                                      "Root", tvwChild, CurTable.Name, CurTable.Name, "Table", "Table");
                nodX.ExpandedImage = "Table";
                nodX.Expanded = true;
                };
                } // CurTable;

                foreach( var CurTable in db.TableDefs;
                CurTable.Name.Substring(0, 4) <> "MSys" )
            {;
                CurTable.Fields(1).Name.Substring(CurTable.Fields(1).Name.Length - 2) = "ID" )
            {;
                CurTable.Fields(1).Name.Substring(0, Len(CurTable.Fields(1).Name) - 2));
                };
                };
                } // CurTable;


                db.Close;
                nodX = null;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As sFieldType            {
                switch iFieldType;
                Case dbBigInt:        sFieldType = "Big Integer";
                Case dbBinary:        sFieldType = "Binary";
                Case dbBoolean:       sFieldType = "Boolean";
                Case dbByte:          sFieldType = "Byte";
                Case dbChar:          sFieldType = "Char";
                Case dbCurrency:      sFieldType = "Currency";
                Case dbDate:          sFieldType = "Date / Time";
                Case dbDecimal:       sFieldType = "Decimal";
                Case dbDouble:        sFieldType = "Double";
                Case dbFloat:         sFieldType = "Float";
                Case dbGUID:          sFieldType = "Guid";
                Case dbInteger:       sFieldType = "Integer";
                Case dbLong:          sFieldType = "Long";
                Case dbLongBinary:    sFieldType = "long Binary (OLE Object)";
                Case dbMemo:          sFieldType = "Memo";
                Case dbNumeric:       sFieldType = "Numeric";
                Case dbSingle:        sFieldType = "Single";
                Case dbText:          sFieldType = "Text";
                Case dbTime:          sFieldType = "Time";
                Case dbTimeStamp:     sFieldType = "Time Stamp";
                Case dbVarBinary:     sFieldType = "VarBinary";
                };
            }

            public void cmdAddCategory_Click()
            {
                ;

                sCategoryChosen = Parent.SliceAndDice.Categorys.Choose(0);
                if ( Len(sCategoryChosen) > 0 )
            {;
                Parent.SliceAndDice.Categorys.Item(sCategoryChosen).CategoryType = 1;
                Parent.SliceAndDice.Save;
                };

                LoadCategories;
            }

            public void cmdDeleteCategory_Click()
            {
                ;

                sCategoryChosen = Parent.SliceAndDice.Categorys.Choose(1);
                if ( Len(sCategoryChosen) > 0 )
            {;
                Parent.SliceAndDice.Categorys.Item(sCategoryChosen).CategoryType = 0;
                Parent.SliceAndDice.Save;
                };

                LoadCategories;
            }

            public void Form_Resize()
            {
                try
{;
                lvwFields.Width = this.ScaleWidth - lvwFields.Left - 100;
                lvwFields.Height = this.ScaleHeight - lvwFields.Top - 100;
                tvwTables.Height = this.ScaleHeight - tvwTables.Top - 100;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void lvwFields_DblClick()
            {
                mnuFieldNew_Click;
            }

            public void lvwFields_KeyPress            {
                if ( KeyAscii = 13 )
            {;
                mnuFieldNew_Click;
                };
            }

            public void mnuFileExit_Click()
            {
                m_bCanceled = true;
                lvwFields.ListItems.Clear;
                Hide;
            }

            public void mnuFileNew_Click()
            {
                ;
                ;
                ;

                sDatabasePath = Trim(BrowseForFolder(hWnd, "Where should database go ?"));
                if ( Len(sDatabasePath) == 0 )
            {
 return;

                sNewDatabaseName = Trim(InputBox("What should the name of the new database be ?", "CREATE BLANK DATABASE"));
                if ( Len(snewDatabaseName) == 0 )
            {
 return;

                sDatabasePath.Substring(sDatabasePath.Length - 1) <> "\" )
            {
 sDatabasePath == sDatabasePath + "\";
                LCase(snewDatabaseName).Substring(LCase(snewDatabaseName).Length - 4) <> ".mdb" )
            {
 snewDatabaseName == sDatabasePath + snewDatabaseName + ".mdb";

                try
{;
                db = CreateDatabase(sNewDatabaseName, dbLangGeneral, dbVersion30);
                db.Close;

                m_sClassDatabaseName = sNewDatabaseName;
                PopulateTree;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuFileOpen2_Click()
            {
                m_sClassDatabaseName = "";
                m_sClassDatabaseOptions = "ODBC;";
                PopulateTree;
            }

            public void mnuGenerateClass_Click()
            {
                if ( lvwFields.ListItems.Count = 0 )
            {;
                MsgBox "Please select a table first.";
                return;
                };

                m_bCanceled = false;
                m_bGenerateBranch = false;
                m_bGenerateDatabase = false;
                Hide;
            }

            public void mnuGenerateEnterBranch_Click()
            {
                if ( lvwFields.ListItems.Count = 0 )
            {;
                MsgBox "Please select a branch first.";
                return;
                };

                m_bCanceled = false;
                m_bGenerateBranch = true;
                m_bGenerateDatabase = false;
                Hide;
            }

            public void mnuGenerateEntireDatabase_Click()
            {
                if ( tvwTables.Nodes.Count = 0 )
            {;
                MsgBox "Please select a database first.";
                return;
                };

                tvwTables.Nodes("Root").Selected = true;
                lvwFields.ListItems.Clear;

                m_bCanceled = false;
                m_bGenerateBranch = false;
                m_bGenerateDatabase = true;
                Hide;
            }

            public void mnuFileOpen_Click()
            {
                m_sClassDatabaseName = Parent.sChooseDatabase();
                if ( Len(m_sClassDatabaseName) > 0 )
            {;
                PopulateTree;
                };
            }

            public void Form_Load()
            {
                //   'mnuFileOpen_Click;
                ;
                LoadCategories;
                ;
                mnuRulesAutoAddKey.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "AutoAddKey", true);
                mnuRulesAutoAddDateModified.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "AutoAddDateModified", true);
                mnuRulesAutoAddDateCreated.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "AutoAddCreated", true);

                mnuRulesEnforce.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "Enforce", true);
                mnuRulesCascadeUpdates.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "CascadeUpdates", true);
                mnuRulesCascadeDeletes.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "CascadeDeletes", true);

                mnuRulesUseAutoNumber.Checked = GetSetting(App.ProductName, "ODBC Class Gen", "UseAutoNumber", true);

                if ( mnuRulesEnforce.Checked )
            {;
                mnuRulesCascadeUpdates.Enabled = true;
                mnuRulesCascadeDeletes.Enabled = true;
                }
            else
            {;
                mnuRulesCascadeUpdates.Enabled = false;
                mnuRulesCascadeDeletes.Enabled = false;
                };

            }

            public void Form_Terminate()
            {
                Parent = null;
            }

            public void Label2_Click()
            {

            }

            public void lvwFields_MouseUp            {
                try
{;
                ;

                if ( Button = vbRightButton )
            {;
                ItemClicked = lvwFields.HitTest(x, y);
                if ( ! ItemClicked Is null )
            {;
                ItemClicked.Selected = true;
                PopupMenu mnuField;
                };
                };
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuFieldDelete_Click()
            {
                ;
                ;
                ;
                ;

                try
{;
                if ( tvwTables.SelectedItem.Parent.Key = m_sClassDatabaseName )
            {;
                sParentTable = "";
                }
            else
            {;
                sParentTable = tvwTables.SelectedItem.Parent.Key;
                };
                sTable = tvwTables.SelectedItem.Key;
                sField = lvwFields.SelectedItem.Key;

                switch sField;
                Case sTable + "ID", sTable + "Name", "DateCreated", "DateModified", sParentTable + "ID";
                MsgBox "Can't delete that field (required for correct object/database operation), sorry.";
                return;
                };

                db = OpenDatabase(m_sClassDatabaseName, "ODBC;");
                db.TableDefs(sTable).Fields.Delete sField;
                db.Close;

                tvwTables_NodeClick tvwTables.SelectedItem;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuFieldNew_Click()
            {
                ;
                ;
                ;
                ;

                mnuFieldNew_Click_TryAgain:;

                try
{;
                ;
                NewField.Show vbModal, Me;
                if ( ex <> 0 )
            {
 return;
                if ( newField.Canceled == true )
            {
 return;

                foreach( var CurItem in lvwFields.ListItems;
                if ( CurItem.Text) = UCase(.FieldName) .ToUpper()
            {;
                MsgBox "That field already exists in this table... Try again.";
                GoTo mnuFieldNew_Click_TryAgain;
                };
                } // CurItem;

                try
{;
                db = OpenDatabase(m_sClassDatabaseName, "ODBC;");
                CurTable = db.TableDefs(tvwTables.SelectedItem.Key);
                if ( NewField.dbFieldType = dbText )
            {;
                CurTable.Fields.Append CurTable.CreateField(.FieldName, NewField.dbFieldType, NewField.Length);
                }
            else
            {;
                CurTable.Fields.Append CurTable.CreateField(.FieldName, NewField.dbFieldType);
                };
                db.Close;

                tvwTables_NodeClick tvwTables.SelectedItem;

                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuRulesAutoAddDateCreated_Click()
            {
                mnuRulesAutoAddDateCreated.Checked = ! mnuRulesAutoAddDateCreated.Checked;
                SaveSetting App.ProductName, "ODBC Class Gen", "AutoAddDateCreated", mnuRulesAutoAddDateCreated.Checked;
            }

            public void mnuRulesAutoAddDateModified_Click()
            {
                mnuRulesAutoAddDateModified.Checked = ! mnuRulesAutoAddDateModified.Checked;
                SaveSetting App.ProductName, "ODBC Class Gen", "AutoAddDateModified", mnuRulesAutoAddDateModified.Checked;
            }

            public void mnuRulesAutoAddKey_Click()
            {
                mnuRulesAutoAddKey.Checked = ! mnuRulesAutoAddKey.Checked;
                SaveSetting App.ProductName, "ODBC Class Gen", "AutoAddKey", mnuRulesAutoAddKey.Checked;
            }

            public void mnuRulesCascadeDeletes_Click()
            {
                mnuRulesCascadeDeletes.Checked = ! mnuRulesCascadeDeletes.Checked;
                SaveSetting App.ProductName, "ODBC Class Gen", "CascadeDeletes", mnuRulesCascadeDeletes.Checked;
            }

            public void mnuRulesCascadeUpdates_Click()
            {
                mnuRulesCascadeUpdates.Checked = ! mnuRulesCascadeUpdates.Checked;
                SaveSetting App.ProductName, "ODBC Class Gen", "CascadeUpdates", mnuRulesCascadeUpdates.Checked;
            }

            public void mnuRulesEnforce_Click()
            {
                mnuRulesEnforce.Checked = ! mnuRulesEnforce.Checked;
                SaveSetting App.ProductName, "ODBC Class Gen", "Enforce", mnuRulesEnforce.Checked;
                if ( mnuRulesEnforce.Checked )
            {;
                mnuRulesCascadeUpdates.Enabled = true;
                mnuRulesCascadeDeletes.Enabled = true;
                }
            else
            {;
                mnuRulesCascadeUpdates.Enabled = false;
                mnuRulesCascadeDeletes.Enabled = false;
                };
            }

            public void mnuRulesUseAutoNumber_Click()
            {
                mnuRulesUseAutoNumber.Checked = ! mnuRulesUseAutoNumber.Checked;
                SaveSetting App.ProductName, "ODBC Class Gen", "UseAutoNumber", mnuRulesUseAutoNumber.Checked;
            }

            public void mnuTableDelete_Click()
            {
                ;

                try
{;
                if ( bUserSure("This will PERMANENTLY remove the table selected." + gsEolTab + "ARE YOU ABSOLUTELY SURE ?") )
            {;
                db = OpenDatabase(m_sClassDatabaseName, "ODBC;");
                db.TableDefs.Delete tvwTables.SelectedItem.Key;
                db.Close;
                PopulateTree;
                };
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuTableNew_Click()
            {
                ;

                mnuTableNew_Click_TryAgain:;
                sTable = sReplace(Trim(InputBox("What should the name of the new table be ?" + gsEolTab + "Note: Table names should be singular, such as:" + gsEolTab + "Book, Publisher, etc.")), " ", "");
                if ( Len(sTable) == 0 )
            {
 return;
                ;
                sTable.Substring(sTable.Length - 1) = "s" )
            {;
                MsgBox "Table names MUST be singular.";
                GoTo mnuTableNew_Click_TryAgain;
                };

                if ( tvwTables.SelectedItem.Text <> DBName )
            {;
                AddTable sTable, tvwTables.SelectedItem.Text;
                }
            else
            {;
                AddTable sTable, "";
                };
                ;
                PopulateTree;
            }

            public void mnuX_Click()
            {
                mnuFileExit_Click;
            }

            public void tvwTables_DblClick()
            {
                mnuTableNew_Click;
            }

            public void tvwTables_KeyPress            {
                if ( KeyAscii = 13 )
            {;
                mnuTableNew_Click;
                };
            }

            public void tvwTables_MouseUp            {
                ;

                if ( Button = vbRightButton )
            {;
                NodeClicked = tvwTables.HitTest(x, y);
                if ( ! NodeClicked Is null )
            {;
                NodeClicked.Selected = true;
                PopupMenu mnuTable;
                };
                };
            }

            public void tvwTables_NodeClick            {
                ;
                ;
                ;
                ;
                ;
                ;

                try
{;
                db = OpenDatabase(m_sClassDatabaseName, "ODBC;");

                lvwFields.ListItems.Clear;

                lvwFields.ColumnHeaders.Clear;
                lvwFields.ColumnHeaders.Add(                                                                                                   , "Field Name", "Field Name", 2000);
                lvwFields.ColumnHeaders.Add(                                                                                                   , "Field Type", "Type", 1000);
                lvwFields.ColumnHeaders.Add(                                                                                                   , "Field Length", "Length", 500);

                lvwFields.View = lvwReport;
                CurTable = db.TableDefs(tvwTables.SelectedItem.Key);
                foreach( var CurField in CurTable.Fields;
                CurField.Name.Substring(CurField.Name.Length - 2) = "ID" )
            {;
                CurField.Name.Substring(0, Len(CurField.Name) - 2) = CurTable.Name )
            {;
                sIcon = "Key";
                }
            else
            {;
                sIcon = "ID";
                };
                }
            else
            {if ( CurField.Name = "DateCreated" Or CurField.Name = "DateModified" )
            {;
                sIcon = "Date";
                }
            else
            {;
                sIcon = "Field";
                };
                litX = lvwFields.ListItems.Add(                                                                                                                                                                                                      , CurField.Name, CurField.Name, sIcon, sIcon);
                litX.SubItems(1) = sFieldType(CurField.Type);
                litX.SubItems(2) = CurField.Size;
                } // CurField;

                db.Close;
                nodX = null;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

        }
    }
