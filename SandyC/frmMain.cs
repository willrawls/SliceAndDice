using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class frmMain : System.Windows.Forms.Form
        {
         public System.Windows.Forms.VB.Frame frmTemplateInfo;
         public System.Windows.Forms.VB.CommandButton cmdRecalc;
         public System.Windows.Forms.VB.CheckBox chkAutoRecalc;
         public System.Windows.Forms.VB.ListBox lstSoftVariables;
         public System.Windows.Forms.VB.ListBox lstSoftCommands;
         public System.Windows.Forms.VB.Label lblTemplateInfo0;
         public System.Windows.Forms.VB.Label lblTemplateInfo1;
         public System.Windows.Forms.VB.Frame frmOptions;
         public System.Windows.Forms.VB.Frame Frame2;
         public System.Windows.Forms.VB.Label lblDelta;
         public System.Windows.Forms.VB.Label lblAlpha;
         public System.Windows.Forms.VB.Label lblRevision;
         public System.Windows.Forms.VB.Frame Frame1;
         public System.Windows.Forms.VB.CheckBox chkSelected;
         public System.Windows.Forms.VB.CheckBox chkFavorite;
         public System.Windows.Forms.VB.CheckBox chkUndeletable;
         public System.Windows.Forms.VB.CheckBox chkLocked;
         public System.Windows.Forms.FirmSolutions.FSListBar lsbJumpTo;
         public System.Windows.Forms.MSComctlLib.ImageList imlTabs;
         public System.Windows.Forms.VB.Timer tmrDoAction;
         public System.Windows.Forms.VB.TextBox txtName;
         public System.Windows.Forms.VB.TextBox txtCode2;
         public System.Windows.Forms.VB.TextBox txtCode1;
         public System.Windows.Forms.VB.TextBox txtCode0;
         public System.Windows.Forms.VB.Frame frmFile;
         public System.Windows.Forms.VB.TextBox txtCodeToFile;
         public System.Windows.Forms.VB.TextBox txtFilename;
         public System.Windows.Forms.VB.Label Label2;
         public System.Windows.Forms.VB.Timer tmrActivateDBClassGen;
         public System.Windows.Forms.VB.TextBox txtShortName;
         public System.Windows.Forms.MSComDlg.CommonDialog cdgSelect;
         public System.Windows.Forms.MSComctlLib.TabStrip tabCode;
         public System.Windows.Forms.VB.Label lblCode3;
         public System.Windows.Forms.VB.Label Label1;
         public System.Windows.Forms.VB.Menu mnuX;
         public System.Windows.Forms.VB.Menu mnuFile;
         public System.Windows.Forms.VB.Menu mnuSpecialOpenDatabase;
         public System.Windows.Forms.VB.Menu mnuSpecialNewDatabase;
         public System.Windows.Forms.VB.Menu mnuSpecialExportSnippet;
         public System.Windows.Forms.VB.Menu mnuSep13;
         public System.Windows.Forms.VB.Menu mnuFileExit;
         public System.Windows.Forms.VB.Menu mnuEdit;
         public System.Windows.Forms.VB.Menu mnuEditCut;
         public System.Windows.Forms.VB.Menu mnuEditCopy;
         public System.Windows.Forms.VB.Menu mnuEditPaste;
         public System.Windows.Forms.VB.Menu mnuEditSep0;
         public System.Windows.Forms.VB.Menu mnuEditFind;
         public System.Windows.Forms.VB.Menu mnuEditReplace;
         public System.Windows.Forms.VB.Menu mnuTemplate;
         public System.Windows.Forms.VB.Menu mnuFileNew;
         public System.Windows.Forms.VB.Menu mnuFileCopy;
         public System.Windows.Forms.VB.Menu mnuSep5;
         public System.Windows.Forms.VB.Menu mnuIsFavorite;
         public System.Windows.Forms.VB.Menu mnuSep25;
         public System.Windows.Forms.VB.Menu mnuInsertTemplate;
         public System.Windows.Forms.VB.Menu mnuFileImport;
         public System.Windows.Forms.VB.Menu mnuSep2;
         public System.Windows.Forms.VB.Menu mnuFileDelete;
         public System.Windows.Forms.VB.Menu mnuCategories;
         public System.Windows.Forms.VB.Menu mnuCategoriesNewMethod0;
         public System.Windows.Forms.VB.Menu mnuCategoriesNewMethod1;
         public System.Windows.Forms.VB.Menu mnuCategoriesNewMethod2;
         public System.Windows.Forms.VB.Menu mnuSep1;
         public System.Windows.Forms.VB.Menu mnuFileRefresh;
         public System.Windows.Forms.VB.Menu mnuSep10;
         public System.Windows.Forms.VB.Menu mnuCategoriesDeleteCurrent;
         public System.Windows.Forms.VB.Menu mnuSpecial;
         public System.Windows.Forms.VB.Menu mnuSpecialViewLog;
         public System.Windows.Forms.VB.Menu mnuProjectProcessor;
         public System.Windows.Forms.VB.Menu mnuSep7;
         public System.Windows.Forms.VB.Menu mnuExitAfterInsert;
         public System.Windows.Forms.VB.Menu mnuSep11;
         public System.Windows.Forms.VB.Menu mnuShowSplash;
         public System.Windows.Forms.VB.Menu mnuShowPaintbrushIcon;
         public System.Windows.Forms.VB.Menu mnuShowOnModuleRightClick;
         public System.Windows.Forms.VB.Menu mnuSwitchTabsAutomatically;
         public System.Windows.Forms.VB.Menu mnuPasswordProtection;
         public System.Windows.Forms.VB.Menu mnuSep15;
         public System.Windows.Forms.VB.Menu mnuOLEDragDrop;
         public System.Windows.Forms.VB.Menu mnuTakeOverKeys;
         public System.Windows.Forms.VB.Menu mnuSep12;
         public System.Windows.Forms.VB.Menu mnuChangeBackgroundColors;
         public System.Windows.Forms.VB.Menu mnuChangeForegroundColor;
         public System.Windows.Forms.VB.Menu mnuHistory;
         public System.Windows.Forms.VB.Menu mnuBack;
         public System.Windows.Forms.VB.Menu mnuForward;
         public System.Windows.Forms.VB.Menu mnuHistorySep0;
         public System.Windows.Forms.VB.Menu mnuHistoryList;
         public System.Windows.Forms.VB.Menu mnuFav;
         public System.Windows.Forms.VB.Menu mnuFavorite0;
         public System.Windows.Forms.VB.Menu mnuExternalFunctions;
         public System.Windows.Forms.VB.Menu mnuDBClassGen;
         public System.Windows.Forms.VB.Menu mnuFileSep1;
         public System.Windows.Forms.VB.Menu mnuExternals0;
         public System.Windows.Forms.VB.Menu mnuHelp;
         public System.Windows.Forms.VB.Menu mnuHelpSoftCommandReference;
         public System.Windows.Forms.VB.Menu mnuHelpCodeGenSoftVarRef;
         public System.Windows.Forms.VB.Menu mnuHelpSep2;
         public System.Windows.Forms.VB.Menu mnuHelpOnlineDocumentation;
         public System.Windows.Forms.VB.Menu mnuHelpReportIssue;
         public System.Windows.Forms.VB.Menu mnuHelpEmailWilliamRawls;
         public System.Windows.Forms.VB.Menu mnuHelpVisitHomePage;
         public System.Windows.Forms.VB.Menu mnuHelpSep1;
         public System.Windows.Forms.VB.Menu mnuHelpContents;
         public System.Windows.Forms.VB.Menu mnuHelpIndex;
         public System.Windows.Forms.VB.Menu mnuHelpSep0;
         public System.Windows.Forms.VB.Menu mnuHelpAbout;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.Container components = null;

        public frmMain()
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
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmMain));
            this.frmTemplateInfo = new System.Windows.Forms.VB.Frame();
            this.cmdRecalc = new System.Windows.Forms.VB.CommandButton();
            this.chkAutoRecalc = new System.Windows.Forms.VB.CheckBox();
            this.lstSoftVariables = new System.Windows.Forms.VB.ListBox();
            this.lstSoftCommands = new System.Windows.Forms.VB.ListBox();
            this.lblTemplateInfo0 = new System.Windows.Forms.VB.Label();
            this.lblTemplateInfo1 = new System.Windows.Forms.VB.Label();
            this.frmOptions = new System.Windows.Forms.VB.Frame();
            this.Frame2 = new System.Windows.Forms.VB.Frame();
            this.lblDelta = new System.Windows.Forms.VB.Label();
            this.lblAlpha = new System.Windows.Forms.VB.Label();
            this.lblRevision = new System.Windows.Forms.VB.Label();
            this.Frame1 = new System.Windows.Forms.VB.Frame();
            this.chkSelected = new System.Windows.Forms.VB.CheckBox();
            this.chkFavorite = new System.Windows.Forms.VB.CheckBox();
            this.chkUndeletable = new System.Windows.Forms.VB.CheckBox();
            this.chkLocked = new System.Windows.Forms.VB.CheckBox();
            this.lsbJumpTo = new System.Windows.Forms.FirmSolutions.FSListBar();
            this.imlTabs = new System.Windows.Forms.MSComctlLib.ImageList();
            this.tmrDoAction = new System.Windows.Forms.VB.Timer();
            this.txtName = new System.Windows.Forms.VB.TextBox();
            this.txtCode2 = new System.Windows.Forms.VB.TextBox();
            this.txtCode1 = new System.Windows.Forms.VB.TextBox();
            this.txtCode0 = new System.Windows.Forms.VB.TextBox();
            this.frmFile = new System.Windows.Forms.VB.Frame();
            this.txtCodeToFile = new System.Windows.Forms.VB.TextBox();
            this.txtFilename = new System.Windows.Forms.VB.TextBox();
            this.Label2 = new System.Windows.Forms.VB.Label();
            this.tmrActivateDBClassGen = new System.Windows.Forms.VB.Timer();
            this.txtShortName = new System.Windows.Forms.VB.TextBox();
            this.cdgSelect = new System.Windows.Forms.MSComDlg.CommonDialog();
            this.tabCode = new System.Windows.Forms.MSComctlLib.TabStrip();
            this.lblCode3 = new System.Windows.Forms.VB.Label();
            this.Label1 = new System.Windows.Forms.VB.Label();
            this.mnuX = new System.Windows.Forms.VB.Menu();
            this.mnuFile = new System.Windows.Forms.VB.Menu();
            this.mnuSpecialOpenDatabase = new System.Windows.Forms.VB.Menu();
            this.mnuSpecialNewDatabase = new System.Windows.Forms.VB.Menu();
            this.mnuSpecialExportSnippet = new System.Windows.Forms.VB.Menu();
            this.mnuSep13 = new System.Windows.Forms.VB.Menu();
            this.mnuFileExit = new System.Windows.Forms.VB.Menu();
            this.mnuEdit = new System.Windows.Forms.VB.Menu();
            this.mnuEditCut = new System.Windows.Forms.VB.Menu();
            this.mnuEditCopy = new System.Windows.Forms.VB.Menu();
            this.mnuEditPaste = new System.Windows.Forms.VB.Menu();
            this.mnuEditSep0 = new System.Windows.Forms.VB.Menu();
            this.mnuEditFind = new System.Windows.Forms.VB.Menu();
            this.mnuEditReplace = new System.Windows.Forms.VB.Menu();
            this.mnuTemplate = new System.Windows.Forms.VB.Menu();
            this.mnuFileNew = new System.Windows.Forms.VB.Menu();
            this.mnuFileCopy = new System.Windows.Forms.VB.Menu();
            this.mnuSep5 = new System.Windows.Forms.VB.Menu();
            this.mnuIsFavorite = new System.Windows.Forms.VB.Menu();
            this.mnuSep25 = new System.Windows.Forms.VB.Menu();
            this.mnuInsertTemplate = new System.Windows.Forms.VB.Menu();
            this.mnuFileImport = new System.Windows.Forms.VB.Menu();
            this.mnuSep2 = new System.Windows.Forms.VB.Menu();
            this.mnuFileDelete = new System.Windows.Forms.VB.Menu();
            this.mnuCategories = new System.Windows.Forms.VB.Menu();
            this.mnuCategoriesNewMethod0 = new System.Windows.Forms.VB.Menu();
            this.mnuCategoriesNewMethod1 = new System.Windows.Forms.VB.Menu();
            this.mnuCategoriesNewMethod2 = new System.Windows.Forms.VB.Menu();
            this.mnuSep1 = new System.Windows.Forms.VB.Menu();
            this.mnuFileRefresh = new System.Windows.Forms.VB.Menu();
            this.mnuSep10 = new System.Windows.Forms.VB.Menu();
            this.mnuCategoriesDeleteCurrent = new System.Windows.Forms.VB.Menu();
            this.mnuSpecial = new System.Windows.Forms.VB.Menu();
            this.mnuSpecialViewLog = new System.Windows.Forms.VB.Menu();
            this.mnuProjectProcessor = new System.Windows.Forms.VB.Menu();
            this.mnuSep7 = new System.Windows.Forms.VB.Menu();
            this.mnuExitAfterInsert = new System.Windows.Forms.VB.Menu();
            this.mnuSep11 = new System.Windows.Forms.VB.Menu();
            this.mnuShowSplash = new System.Windows.Forms.VB.Menu();
            this.mnuShowPaintbrushIcon = new System.Windows.Forms.VB.Menu();
            this.mnuShowOnModuleRightClick = new System.Windows.Forms.VB.Menu();
            this.mnuSwitchTabsAutomatically = new System.Windows.Forms.VB.Menu();
            this.mnuPasswordProtection = new System.Windows.Forms.VB.Menu();
            this.mnuSep15 = new System.Windows.Forms.VB.Menu();
            this.mnuOLEDragDrop = new System.Windows.Forms.VB.Menu();
            this.mnuTakeOverKeys = new System.Windows.Forms.VB.Menu();
            this.mnuSep12 = new System.Windows.Forms.VB.Menu();
            this.mnuChangeBackgroundColors = new System.Windows.Forms.VB.Menu();
            this.mnuChangeForegroundColor = new System.Windows.Forms.VB.Menu();
            this.mnuHistory = new System.Windows.Forms.VB.Menu();
            this.mnuBack = new System.Windows.Forms.VB.Menu();
            this.mnuForward = new System.Windows.Forms.VB.Menu();
            this.mnuHistorySep0 = new System.Windows.Forms.VB.Menu();
            this.mnuHistoryList = new System.Windows.Forms.VB.Menu();
            this.mnuFav = new System.Windows.Forms.VB.Menu();
            this.mnuFavorite0 = new System.Windows.Forms.VB.Menu();
            this.mnuExternalFunctions = new System.Windows.Forms.VB.Menu();
            this.mnuDBClassGen = new System.Windows.Forms.VB.Menu();
            this.mnuFileSep1 = new System.Windows.Forms.VB.Menu();
            this.mnuExternals0 = new System.Windows.Forms.VB.Menu();
            this.mnuHelp = new System.Windows.Forms.VB.Menu();
            this.mnuHelpSoftCommandReference = new System.Windows.Forms.VB.Menu();
            this.mnuHelpCodeGenSoftVarRef = new System.Windows.Forms.VB.Menu();
            this.mnuHelpSep2 = new System.Windows.Forms.VB.Menu();
            this.mnuHelpOnlineDocumentation = new System.Windows.Forms.VB.Menu();
            this.mnuHelpReportIssue = new System.Windows.Forms.VB.Menu();
            this.mnuHelpEmailWilliamRawls = new System.Windows.Forms.VB.Menu();
            this.mnuHelpVisitHomePage = new System.Windows.Forms.VB.Menu();
            this.mnuHelpSep1 = new System.Windows.Forms.VB.Menu();
            this.mnuHelpContents = new System.Windows.Forms.VB.Menu();
            this.mnuHelpIndex = new System.Windows.Forms.VB.Menu();
            this.mnuHelpSep0 = new System.Windows.Forms.VB.Menu();
            this.mnuHelpAbout = new System.Windows.Forms.VB.Menu();
            this.SuspendLayout();
            this.frmTemplateInfo.SuspendLayout();
            this.frmOptions.SuspendLayout();
            this.Frame2.SuspendLayout();
            this.Frame1.SuspendLayout();
            this.frmFile.SuspendLayout();
            this.mnuFile.SuspendLayout();
            this.mnuEdit.SuspendLayout();
            this.mnuTemplate.SuspendLayout();
            this.mnuCategories.SuspendLayout();
            this.mnuSpecial.SuspendLayout();
            this.mnuSep7.SuspendLayout();
            this.mnuHistory.SuspendLayout();
            this.mnuFav.SuspendLayout();
            this.mnuExternalFunctions.SuspendLayout();
            this.mnuHelp.SuspendLayout();
            //
            // frmTemplateInfo
            //
            this.frmTemplateInfo.Name = "frmTemplateInfo";
            this.frmTemplateInfo.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.frmTemplateInfo.Text = "Frame1";
            this.frmTemplateInfo.Size = new System.Drawing.Size(269, 129);
            this.frmTemplateInfo.Location = new System.Drawing.Point(350, 85);
            this.frmTemplateInfo.TabIndex = 1;
            this.frmTemplateInfo.Visible = false;
            this.frmTemplateInfo.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.cmdRecalc,
                        this.chkAutoRecalc,
                        this.lstSoftVariables,
                        this.lstSoftCommands,
                        this.lblTemplateInfo0,
                        this.lblTemplateInfo1
            });
            //
            // cmdRecalc
            //
            this.cmdRecalc.Name = "cmdRecalc";
            this.cmdRecalc.Text = "Recalc";
            this.cmdRecalc.Size = new System.Drawing.Size(54, 29);
            this.cmdRecalc.Location = new System.Drawing.Point(8, 5);
            this.cmdRecalc.TabIndex = 5;
            //
            // chkAutoRecalc
            //
            this.chkAutoRecalc.Name = "chkAutoRecalc";
            this.chkAutoRecalc.Text = "Auto Recalc when tab selected.";
            this.chkAutoRecalc.Size = new System.Drawing.Size(174, 20);
            this.chkAutoRecalc.Location = new System.Drawing.Point(66, 11);
            this.chkAutoRecalc.TabIndex = 4;
            //
            // lstSoftVariables
            //
            this.lstSoftVariables.Name = "lstSoftVariables";
            this.lstSoftVariables.BackColor = System.Drawing.Color.FromArgb(-2147483624);
            this.lstSoftVariables.Size = new System.Drawing.Size(122, 67);
//            this.lstSoftVariables.IntegralHeight = 0;
            this.lstSoftVariables.Location = new System.Drawing.Point(13, 56);
            this.lstSoftVariables.TabIndex = 3;
            //
            // lstSoftCommands
            //
            this.lstSoftCommands.Name = "lstSoftCommands";
            this.lstSoftCommands.BackColor = System.Drawing.Color.FromArgb(-2147483624);
            this.lstSoftCommands.Size = new System.Drawing.Size(124, 67);
//            this.lstSoftCommands.IntegralHeight = 0;
            this.lstSoftCommands.Location = new System.Drawing.Point(142, 57);
            this.lstSoftCommands.TabIndex = 2;
            //
            // lblTemplateInfo0
            //
            this.lblTemplateInfo0.Name = "lblTemplateInfo0";
            this.lblTemplateInfo0.Text = "Soft Variables in Use";
            this.lblTemplateInfo0.Size = new System.Drawing.Size(99, 13);
            this.lblTemplateInfo0.Location = new System.Drawing.Point(13, 38);
            this.lblTemplateInfo0.TabIndex = 7;
            //
            // lblTemplateInfo1
            //
            this.lblTemplateInfo1.Name = "lblTemplateInfo1";
            this.lblTemplateInfo1.Text = "Soft Commands in Use";
            this.lblTemplateInfo1.Size = new System.Drawing.Size(108, 13);
            this.lblTemplateInfo1.Location = new System.Drawing.Point(143, 39);
            this.lblTemplateInfo1.TabIndex = 6;
            //
            // frmOptions
            //
            this.frmOptions.Name = "frmOptions";
            this.frmOptions.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.frmOptions.Text = "Frame1";
            this.frmOptions.Size = new System.Drawing.Size(337, 247);
            this.frmOptions.Location = new System.Drawing.Point(292, 104);
            this.frmOptions.TabIndex = 8;
            this.frmOptions.Visible = false;
            this.frmOptions.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.Frame2,
                        this.Frame1
            });
            //
            // Frame2
            //
            this.Frame2.Name = "Frame2";
            this.Frame2.Text = " Statistics ";
            this.Frame2.Size = new System.Drawing.Size(329, 91);
            this.Frame2.Location = new System.Drawing.Point(4, 141);
            this.Frame2.TabIndex = 26;
            this.Frame2.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.lblDelta,
                        this.lblAlpha,
                        this.lblRevision
            });
            //
            // lblDelta
            //
            this.lblDelta.Name = "lblDelta";
            this.lblDelta.Text = "Delta Date";
            this.lblDelta.Font = new System.Drawing.Font("Times New Roman",12F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblDelta.Size = new System.Drawing.Size(74, 19);
            this.lblDelta.Location = new System.Drawing.Point(10, 62);
            this.lblDelta.TabIndex = 29;
            //
            // lblAlpha
            //
            this.lblAlpha.Name = "lblAlpha";
            this.lblAlpha.Text = "Alpha Date: September 15, 2000 12:00:00 PM";
            this.lblAlpha.Font = new System.Drawing.Font("Times New Roman",12F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblAlpha.Size = new System.Drawing.Size(308, 19);
            this.lblAlpha.Location = new System.Drawing.Point(10, 38);
            this.lblAlpha.TabIndex = 28;
            //
            // lblRevision
            //
            this.lblRevision.Name = "lblRevision";
            this.lblRevision.Text = "Revision # ";
            this.lblRevision.Font = new System.Drawing.Font("Times New Roman",12F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblRevision.Size = new System.Drawing.Size(75, 19);
            this.lblRevision.Location = new System.Drawing.Point(10, 14);
            this.lblRevision.TabIndex = 27;
            //
            // Frame1
            //
            this.Frame1.Name = "Frame1";
            this.Frame1.Text = " Basic ";
            this.Frame1.Size = new System.Drawing.Size(115, 131);
            this.Frame1.Location = new System.Drawing.Point(4, 0);
            this.Frame1.TabIndex = 21;
            this.Frame1.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.chkSelected,
                        this.chkFavorite,
                        this.chkUndeletable,
                        this.chkLocked
            });
            //
            // chkSelected
            //
            this.chkSelected.Name = "chkSelected";
            this.chkSelected.Text = "Selected";
            this.chkSelected.Size = new System.Drawing.Size(100, 25);
            this.chkSelected.Location = new System.Drawing.Point(10, 101);
            this.chkSelected.TabIndex = 25;
//            this.chkSelected.ToolTipText = "When checked, the Template will be available for direct insertion on the code window;
            //
            // chkFavorite
            //
            this.chkFavorite.Name = "chkFavorite";
            this.chkFavorite.Text = "Favorite";
            this.chkFavorite.Size = new System.Drawing.Size(100, 25);
            this.chkFavorite.Location = new System.Drawing.Point(10, 74);
            this.chkFavorite.TabIndex = 24;
//            this.chkFavorite.ToolTipText = "When checked, the Template will be available for direct insertion on the code window;
            //
            // chkUndeletable
            //
            this.chkUndeletable.Name = "chkUndeletable";
            this.chkUndeletable.Text = "Undeletable";
            this.chkUndeletable.Size = new System.Drawing.Size(100, 25);
            this.chkUndeletable.Location = new System.Drawing.Point(10, 20);
            this.chkUndeletable.TabIndex = 23;
//            this.chkUndeletable.ToolTipText = "When checked, this Template will not allow users to delete  it.";
            //
            // chkLocked
            //
            this.chkLocked.Name = "chkLocked";
            this.chkLocked.Text = "Code Locked";
            this.chkLocked.Size = new System.Drawing.Size(100, 25);
            this.chkLocked.Location = new System.Drawing.Point(10, 47);
            this.chkLocked.TabIndex = 22;
//            this.chkLocked.ToolTipText = "When checked, this Template will not allow users to modify its code contents.";
            //
            // lsbJumpTo
            //
            this.lsbJumpTo.Name = "lsbJumpTo";
//            this.lsbJumpTo.Align = 3;
            this.lsbJumpTo.Size = new System.Drawing.Size(236, 528);
            this.lsbJumpTo.Location = new System.Drawing.Point(0, 0);
            this.lsbJumpTo.TabIndex = 20;
            this.lsbJumpTo.Font = new System.Drawing.Font("MS Sans Serif",8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lsbJumpTo.BackColor = System.Drawing.Color.FromArgb(-2147483624);
//            this.lsbJumpTo.Arrange = 1;
            this.lsbJumpTo.LabelEdit = true;
            this.lsbJumpTo.View = System.Windows.Forms.View.List;
            //
            // imlTabs
            //
            this.imlTabs.Name = "imlTabs";
            this.imlTabs.Location = new System.Drawing.Point(244, 396);
            this.imlTabs.BackColor = System.Drawing.Color.FromArgb(-2147483638);
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
//            this.imlTabs.ListImage7 = ;
            //
            // tmrDoAction
            //
            this.tmrDoAction.Name = "tmrDoAction";
            this.tmrDoAction.Enabled = false;
            this.tmrDoAction.Interval = 500;
            this.tmrDoAction.Location = new System.Drawing.Point(184, 68);
            //
            // txtName
            //
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(276, 20);
            this.txtName.Location = new System.Drawing.Point(306, 2);
//            this.txtName.OLEDragMode = 1;
            this.txtName.TabIndex = 16;
//            this.txtName.ToolTipText = "Enter the name of the Template here.";
            this.txtName.Visible = false;
            //
            // txtCode2
            //
            this.txtCode2.Name = "txtCode2";
            this.txtCode2.BackColor = System.Drawing.Color.FromArgb(-2147483624);
            this.txtCode2.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtCode2.CausesValidatio = 0;
            this.txtCode2.Font = new System.Drawing.Font("Fixedsys",9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.txtCode2.Size = new System.Drawing.Size(217, 65);
            this.txtCode2.Location = new System.Drawing.Point(268, 68);
//            this.txtCode2.MultiLine = -1;
            this.txtCode2.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtCode2.TabIndex = 15;
            this.txtCode2.Tag = "Code Area 2";
            this.txtCode2.Visible = false;
            //
            // txtCode1
            //
            this.txtCode1.Name = "txtCode1";
            this.txtCode1.BackColor = System.Drawing.Color.FromArgb(-2147483624);
            this.txtCode1.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtCode1.CausesValidatio = 0;
            this.txtCode1.Font = new System.Drawing.Font("Fixedsys",9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.txtCode1.Size = new System.Drawing.Size(213, 55);
            this.txtCode1.Location = new System.Drawing.Point(268, 144);
//            this.txtCode1.MultiLine = -1;
            this.txtCode1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtCode1.TabIndex = 14;
            this.txtCode1.Tag = "Code Area 1";
            this.txtCode1.Visible = false;
            //
            // txtCode0
            //
            this.txtCode0.Name = "txtCode0";
            this.txtCode0.BackColor = System.Drawing.Color.FromArgb(-2147483624);
            this.txtCode0.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtCode0.CausesValidatio = 0;
            this.txtCode0.Font = new System.Drawing.Font("Fixedsys",9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.txtCode0.Size = new System.Drawing.Size(219, 59);
            this.txtCode0.Location = new System.Drawing.Point(268, 204);
//            this.txtCode0.MultiLine = -1;
            this.txtCode0.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtCode0.TabIndex = 13;
            this.txtCode0.Tag = "Code Area 0";
            //
            // frmFile
            //
            this.frmFile.Name = "frmFile";
            this.frmFile.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.frmFile.Text = "Frame1";
            this.frmFile.Size = new System.Drawing.Size(269, 129);
            this.frmFile.Location = new System.Drawing.Point(291, 66);
            this.frmFile.TabIndex = 9;
            this.frmFile.Visible = false;
            this.frmFile.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.txtCodeToFile,
                        this.txtFilename,
                        this.Label2
            });
            //
            // txtCodeToFile
            //
            this.txtCodeToFile.Name = "txtCodeToFile";
            this.txtCodeToFile.BackColor = System.Drawing.Color.FromArgb(-2147483624);
            this.txtCodeToFile.Font = new System.Drawing.Font("Fixedsys",9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.txtCodeToFile.Size = new System.Drawing.Size(217, 65);
            this.txtCodeToFile.Location = new System.Drawing.Point(6, 50);
//            this.txtCodeToFile.MultiLine = -1;
            this.txtCodeToFile.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtCodeToFile.TabIndex = 11;
            this.txtCodeToFile.Tag = "Code Area File";
            //
            // txtFilename
            //
            this.txtFilename.Name = "txtFilename";
            this.txtFilename.Size = new System.Drawing.Size(223, 19);
            this.txtFilename.Location = new System.Drawing.Point(6, 16);
            this.txtFilename.MaxLength = 255;
            this.txtFilename.TabIndex = 10;
//            this.txtFilename.ToolTipText = "This can include template variables.";
            //
            // Label2
            //
            this.Label2.Name = "Label2";
            this.Label2.Text = "Filename to send output to:";
            this.Label2.Size = new System.Drawing.Size(191, 23);
            this.Label2.Location = new System.Drawing.Point(6, 2);
            this.Label2.TabIndex = 12;
            //
            // tmrActivateDBClassGen
            //
            this.tmrActivateDBClassGen.Name = "tmrActivateDBClassGen";
            this.tmrActivateDBClassGen.Enabled = false;
            this.tmrActivateDBClassGen.Interval = 500;
            this.tmrActivateDBClassGen.Location = new System.Drawing.Point(184, 36);
            //
            // txtShortName
            //
            this.txtShortName.Name = "txtShortName";
            this.txtShortName.Size = new System.Drawing.Size(276, 20);
            this.txtShortName.Location = new System.Drawing.Point(306, 2);
            this.txtShortName.TabIndex = 0;
//            this.txtShortName.ToolTipText = "Enter the name of the Snippet here.";
            //
            // cdgSelect
            //
            this.cdgSelect.Name = "cdgSelect";
            this.cdgSelect.Location = new System.Drawing.Point(292, 414);
//            this.cdgSelect.CancelError = -1;
//            this.cdgSelect.DefaultExt = ".mdb";
//            this.cdgSelect.DialogTitle = "Select Access97 DB to work on";
//            this.cdgSelect.Filter = "*.mdb";
            //
            // tabCode
            //
            this.tabCode.Name = "tabCode";
            this.tabCode.Size = new System.Drawing.Size(500, 363);
            this.tabCode.Location = new System.Drawing.Point(235, 28);
            this.tabCode.TabIndex = 19;
//            this.tabCode.HotTracking = -1;
//            this.tabCode.ImageList = "imlTabs";
//            this.tabCode.Tab1 = ;
//            this.tabCode.Tab2 = ;
//            this.tabCode.Tab3 = ;
//            this.tabCode.Tab4 = ;
//            this.tabCode.Tab5 = ;
//            this.tabCode.Tab6 = ;
//            this.tabCode.Tab6 = ;
            //
            // lblCode3
            //
            this.lblCode3.Name = "lblCode3";
            this.lblCode3.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.lblCode3.Text = "Name:";
            this.lblCode3.Font = new System.Drawing.Font("Times New Roman",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblCode3.Size = new System.Drawing.Size(34, 13);
            this.lblCode3.Location = new System.Drawing.Point(267, 5);
            this.lblCode3.TabIndex = 18;
            //
            // Label1
            //
            this.Label1.Name = "Label1";
            this.Label1.Text = "Label1";
            this.Label1.Size = new System.Drawing.Size(77, 23);
            this.Label1.Location = new System.Drawing.Point(300, 64);
            this.Label1.TabIndex = 17;
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
                        this.mnuSpecialOpenDatabase,
                        this.mnuSpecialNewDatabase,
                        this.mnuSpecialExportSnippet,
                        this.mnuSep13,
                        this.mnuFileExit
            });
            //
            // mnuSpecialOpenDatabase
            //
            this.mnuSpecialOpenDatabase.Name = "mnuSpecialOpenDatabase";
            this.mnuSpecialOpenDatabase.Text = "&Open Slice and Dice database";
            //
            // mnuSpecialNewDatabase
            //
            this.mnuSpecialNewDatabase.Name = "mnuSpecialNewDatabase";
            this.mnuSpecialNewDatabase.Text = "&New Slice and Dice database";
            //
            // mnuSpecialExportSnippet
            //
            this.mnuSpecialExportSnippet.Name = "mnuSpecialExportSnippet";
            this.mnuSpecialExportSnippet.Text = "Export current Template";
            this.mnuSpecialExportSnippet.Enabled = false;
            this.mnuSpecialExportSnippet.Visible = false;
            //
            // mnuSep13
            //
            this.mnuSep13.Name = "mnuSep13";
            this.mnuSep13.Text = "-";
            //
            // mnuFileExit
            //
            this.mnuFileExit.Name = "mnuFileExit";
            this.mnuFileExit.Text = "E&xit";
            //
            // mnuEdit
            //
            this.mnuEdit.Name = "mnuEdit";
            this.mnuEdit.Text = "&Edit";
            this.mnuEdit.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuEditCut,
                        this.mnuEditCopy,
                        this.mnuEditPaste,
                        this.mnuEditSep0,
                        this.mnuEditFind,
                        this.mnuEditReplace
            });
            //
            // mnuEditCut
            //
            this.mnuEditCut.Name = "mnuEditCut";
            this.mnuEditCut.Text = "Cut";
            //
            // mnuEditCopy
            //
            this.mnuEditCopy.Name = "mnuEditCopy";
            this.mnuEditCopy.Text = "Copy";
            //
            // mnuEditPaste
            //
            this.mnuEditPaste.Name = "mnuEditPaste";
            this.mnuEditPaste.Text = "Paste";
            //
            // mnuEditSep0
            //
            this.mnuEditSep0.Name = "mnuEditSep0";
            this.mnuEditSep0.Text = "-";
            //
            // mnuEditFind
            //
            this.mnuEditFind.Name = "mnuEditFind";
            this.mnuEditFind.Text = "&Find";
            //
            // mnuEditReplace
            //
            this.mnuEditReplace.Name = "mnuEditReplace";
            this.mnuEditReplace.Text = "&Replace";
            //
            // mnuTemplate
            //
            this.mnuTemplate.Name = "mnuTemplate";
            this.mnuTemplate.Text = "&Template";
//            this.mnuTemplate.NegotiatePositio = 3;
            this.mnuTemplate.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuFileNew,
                        this.mnuFileCopy,
                        this.mnuSep5,
                        this.mnuIsFavorite,
                        this.mnuSep25,
                        this.mnuInsertTemplate,
                        this.mnuFileImport,
                        this.mnuSep2,
                        this.mnuFileDelete
            });
            //
            // mnuFileNew
            //
            this.mnuFileNew.Name = "mnuFileNew";
            this.mnuFileNew.Text = "&New";
            //
            // mnuFileCopy
            //
            this.mnuFileCopy.Name = "mnuFileCopy";
            this.mnuFileCopy.Text = "&Copy current";
            //
            // mnuSep5
            //
            this.mnuSep5.Name = "mnuSep5";
            this.mnuSep5.Text = "-";
            //
            // mnuIsFavorite
            //
            this.mnuIsFavorite.Name = "mnuIsFavorite";
            this.mnuIsFavorite.Text = "On &Favorites Menu";
            //
            // mnuSep25
            //
            this.mnuSep25.Name = "mnuSep25";
            this.mnuSep25.Text = "-";
            //
            // mnuInsertTemplate
            //
            this.mnuInsertTemplate.Name = "mnuInsertTemplate";
            this.mnuInsertTemplate.Text = "&Insert Template into VB";
            //
            // mnuFileImport
            //
            this.mnuFileImport.Name = "mnuFileImport";
            this.mnuFileImport.Text = "I&mport selected code from VB";
            //
            // mnuSep2
            //
            this.mnuSep2.Name = "mnuSep2";
            this.mnuSep2.Text = "-";
            //
            // mnuFileDelete
            //
            this.mnuFileDelete.Name = "mnuFileDelete";
            this.mnuFileDelete.Text = "Delete";
            //
            // mnuCategories
            //
            this.mnuCategories.Name = "mnuCategories";
            this.mnuCategories.Text = "&Categories";
            this.mnuCategories.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuCategoriesNewMethod0,
                        this.mnuCategoriesNewMethod1,
                        this.mnuCategoriesNewMethod2,
                        this.mnuSep1,
                        this.mnuFileRefresh,
                        this.mnuSep10,
                        this.mnuCategoriesDeleteCurrent
            });
            //
            // mnuCategoriesNewMethod0
            //
            this.mnuCategoriesNewMethod0.Name = "mnuCategoriesNewMethod0";
            this.mnuCategoriesNewMethod0.Text = "&New category";
            //
            // mnuCategoriesNewMethod1
            //
            this.mnuCategoriesNewMethod1.Name = "mnuCategoriesNewMethod1";
            this.mnuCategoriesNewMethod1.Text = "&Duplicate a category. Template names and code";
            //
            // mnuCategoriesNewMethod2
            //
            this.mnuCategoriesNewMethod2.Name = "mnuCategoriesNewMethod2";
            this.mnuCategoriesNewMethod2.Text = "Duplicate a category. Template names only";
            //
            // mnuSep1
            //
            this.mnuSep1.Name = "mnuSep1";
            this.mnuSep1.Text = "-";
            //
            // mnuFileRefresh
            //
            this.mnuFileRefresh.Name = "mnuFileRefresh";
            this.mnuFileRefresh.Text = "Refresh Category and Template List";
            //
            // mnuSep10
            //
            this.mnuSep10.Name = "mnuSep10";
            this.mnuSep10.Text = "-";
            //
            // mnuCategoriesDeleteCurrent
            //
            this.mnuCategoriesDeleteCurrent.Name = "mnuCategoriesDeleteCurrent";
            this.mnuCategoriesDeleteCurrent.Text = "Delete current category";
            //
            // mnuSpecial
            //
            this.mnuSpecial.Name = "mnuSpecial";
            this.mnuSpecial.Text = "Tools";
            this.mnuSpecial.Visible = false;
            this.mnuSpecial.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuSpecialViewLog,
                        this.mnuProjectProcessor
            });
            //
            // mnuSpecialViewLog
            //
            this.mnuSpecialViewLog.Name = "mnuSpecialViewLog";
            this.mnuSpecialViewLog.Text = "&View Insertion Log";
            this.mnuSpecialViewLog.Enabled = false;
            //
            // mnuProjectProcessor
            //
            this.mnuProjectProcessor.Name = "mnuProjectProcessor";
            this.mnuProjectProcessor.Text = "&Project Processor (Future)";
            this.mnuProjectProcessor.Enabled = false;
            //
            // mnuSep7
            //
            this.mnuSep7.Name = "mnuSep7";
            this.mnuSep7.Text = "&Options";
            this.mnuSep7.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuExitAfterInsert,
                        this.mnuSep11,
                        this.mnuShowSplash,
                        this.mnuShowPaintbrushIcon,
                        this.mnuShowOnModuleRightClick,
                        this.mnuSwitchTabsAutomatically,
                        this.mnuPasswordProtection,
                        this.mnuSep15,
                        this.mnuOLEDragDrop,
                        this.mnuTakeOverKeys,
                        this.mnuSep12,
                        this.mnuChangeBackgroundColors,
                        this.mnuChangeForegroundColor
            });
            //
            // mnuExitAfterInsert
            //
            this.mnuExitAfterInsert.Name = "mnuExitAfterInsert";
            this.mnuExitAfterInsert.Text = "Exit after insert ?";
            //
            // mnuSep11
            //
            this.mnuSep11.Name = "mnuSep11";
            this.mnuSep11.Text = "-";
            //
            // mnuShowSplash
            //
            this.mnuShowSplash.Name = "mnuShowSplash";
            this.mnuShowSplash.Text = "Show splash screen at startup";
//            this.mnuShowSplash.Checked = -1;
            //
            // mnuShowPaintbrushIcon
            //
            this.mnuShowPaintbrushIcon.Name = "mnuShowPaintbrushIcon";
            this.mnuShowPaintbrushIcon.Text = "Show Paintbrush icon on ""Standard"" menu";
//            this.mnuShowPaintbrushIcon.Checked = -1;
            //
            // mnuShowOnModuleRightClick
            //
            this.mnuShowOnModuleRightClick.Name = "mnuShowOnModuleRightClick";
            this.mnuShowOnModuleRightClick.Text = "Show ""Slice and Dice"" on Module right click";
//            this.mnuShowOnModuleRightClick.Checked = -1;
            //
            // mnuSwitchTabsAutomatically
            //
            this.mnuSwitchTabsAutomatically.Name = "mnuSwitchTabsAutomatically";
            this.mnuSwitchTabsAutomatically.Text = "Switch to first tab with code when switching templates";
//            this.mnuSwitchTabsAutomatically.Checked = -1;
            //
            // mnuPasswordProtection
            //
            this.mnuPasswordProtection.Name = "mnuPasswordProtection";
            this.mnuPasswordProtection.Text = "Password Protection";
            this.mnuPasswordProtection.Enabled = false;
            this.mnuPasswordProtection.Visible = false;
            //
            // mnuSep15
            //
            this.mnuSep15.Name = "mnuSep15";
            this.mnuSep15.Text = "-";
            //
            // mnuOLEDragDrop
            //
            this.mnuOLEDragDrop.Name = "mnuOLEDragDrop";
            this.mnuOLEDragDrop.Text = "Use OLE Text Editing - Drag && Drop";
            //
            // mnuTakeOverKeys
            //
            this.mnuTakeOverKeys.Name = "mnuTakeOverKeys";
            this.mnuTakeOverKeys.Text = "Take over CTRL-SHIFT-1234567890";
            //
            // mnuSep12
            //
            this.mnuSep12.Name = "mnuSep12";
            this.mnuSep12.Text = "-";
            //
            // mnuChangeBackgroundColors
            //
            this.mnuChangeBackgroundColors.Name = "mnuChangeBackgroundColors";
            this.mnuChangeBackgroundColors.Text = "Change Background Colors";
            //
            // mnuChangeForegroundColor
            //
            this.mnuChangeForegroundColor.Name = "mnuChangeForegroundColor";
            this.mnuChangeForegroundColor.Text = "Change Foreground Color";
            //
            // mnuHistory
            //
            this.mnuHistory.Name = "mnuHistory";
            this.mnuHistory.Text = "&History";
            this.mnuHistory.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuBack,
                        this.mnuForward,
                        this.mnuHistorySep0,
                        this.mnuHistoryList
            });
            //
            // mnuBack
            //
            this.mnuBack.Name = "mnuBack";
            this.mnuBack.Text = "&Back";
            //
            // mnuForward
            //
            this.mnuForward.Name = "mnuForward";
            this.mnuForward.Text = "&Forward";
            //
            // mnuHistorySep0
            //
            this.mnuHistorySep0.Name = "mnuHistorySep0";
            this.mnuHistorySep0.Text = "-";
            //
            // mnuHistoryList
            //
            this.mnuHistoryList.Name = "mnuHistoryList";
            this.mnuHistoryList.Text = "&List";
            //
            // mnuFav
            //
            this.mnuFav.Name = "mnuFav";
            this.mnuFav.Text = "Fa&vorites";
            this.mnuFav.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuFavorite0
            });
            //
            // mnuFavorite0
            //
            this.mnuFavorite0.Name = "mnuFavorite0";
            this.mnuFavorite0.Text = "-Empty-";
            this.mnuFavorite0.Enabled = false;
            //
            // mnuExternalFunctions
            //
            this.mnuExternalFunctions.Name = "mnuExternalFunctions";
            this.mnuExternalFunctions.Text = "Exte&rnals";
            this.mnuExternalFunctions.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuDBClassGen,
                        this.mnuFileSep1,
                        this.mnuExternals0
            });
            //
            // mnuDBClassGen
            //
            this.mnuDBClassGen.Name = "mnuDBClassGen";
            this.mnuDBClassGen.Text = "&Database to Code Generator";
            //
            // mnuFileSep1
            //
            this.mnuFileSep1.Name = "mnuFileSep1";
            this.mnuFileSep1.Text = "-";
            //
            // mnuExternals0
            //
            this.mnuExternals0.Name = "mnuExternals0";
            this.mnuExternals0.Text = "-Empty-";
            this.mnuExternals0.Enabled = false;
            //
            // mnuHelp
            //
            this.mnuHelp.Name = "mnuHelp";
            this.mnuHelp.Text = "Hel&p";
            this.mnuHelp.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.mnuHelpSoftCommandReference,
                        this.mnuHelpCodeGenSoftVarRef,
                        this.mnuHelpSep2,
                        this.mnuHelpOnlineDocumentation,
                        this.mnuHelpReportIssue,
                        this.mnuHelpEmailWilliamRawls,
                        this.mnuHelpVisitHomePage,
                        this.mnuHelpSep1,
                        this.mnuHelpContents,
                        this.mnuHelpIndex,
                        this.mnuHelpSep0,
                        this.mnuHelpAbout
            });
            //
            // mnuHelpSoftCommandReference
            //
            this.mnuHelpSoftCommandReference.Name = "mnuHelpSoftCommandReference";
            this.mnuHelpSoftCommandReference.Text = "Soft &Command Reference";
            //
            // mnuHelpCodeGenSoftVarRef
            //
            this.mnuHelpCodeGenSoftVarRef.Name = "mnuHelpCodeGenSoftVarRef";
            this.mnuHelpCodeGenSoftVarRef.Text = "Code &Gen Soft Variable Reference";
            this.mnuHelpCodeGenSoftVarRef.Enabled = false;
            //
            // mnuHelpSep2
            //
            this.mnuHelpSep2.Name = "mnuHelpSep2";
            this.mnuHelpSep2.Text = "-";
            //
            // mnuHelpOnlineDocumentation
            //
            this.mnuHelpOnlineDocumentation.Name = "mnuHelpOnlineDocumentation";
            this.mnuHelpOnlineDocumentation.Text = "Online &Documentation";
            //
            // mnuHelpReportIssue
            //
            this.mnuHelpReportIssue.Name = "mnuHelpReportIssue";
            this.mnuHelpReportIssue.Text = "Report an &Issue";
            //
            // mnuHelpEmailWilliamRawls
            //
            this.mnuHelpEmailWilliamRawls.Name = "mnuHelpEmailWilliamRawls";
            this.mnuHelpEmailWilliamRawls.Text = "&Email William Rawls";
            //
            // mnuHelpVisitHomePage
            //
            this.mnuHelpVisitHomePage.Name = "mnuHelpVisitHomePage";
            this.mnuHelpVisitHomePage.Text = "Visit the &Home Page";
            //
            // mnuHelpSep1
            //
            this.mnuHelpSep1.Name = "mnuHelpSep1";
            this.mnuHelpSep1.Text = "-";
            //
            // mnuHelpContents
            //
            this.mnuHelpContents.Name = "mnuHelpContents";
            this.mnuHelpContents.Text = "Co&ntents";
            this.mnuHelpContents.Enabled = false;
            this.mnuHelpContents.Visible = false;
            //
            // mnuHelpIndex
            //
            this.mnuHelpIndex.Name = "mnuHelpIndex";
            this.mnuHelpIndex.Text = "&Index";
            this.mnuHelpIndex.Enabled = false;
            this.mnuHelpIndex.Visible = false;
            //
            // mnuHelpSep0
            //
            this.mnuHelpSep0.Name = "mnuHelpSep0";
            this.mnuHelpSep0.Text = "-";
            this.mnuHelpSep0.Visible = false;
            //
            // mnuHelpAbout
            //
            this.mnuHelpAbout.Name = "mnuHelpAbout";
            this.mnuHelpAbout.Text = "&About";
            //
            // frmMain
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.frmTemplateInfo,
                  this.frmOptions,
                  this.lsbJumpTo,
                  this.imlTabs,
                  this.tmrDoAction,
                  this.txtName,
                  this.txtCode2,
                  this.txtCode1,
                  this.txtCode0,
                  this.frmFile,
                  this.tmrActivateDBClassGen,
                  this.txtShortName,
                  this.cdgSelect,
                  this.tabCode,
                  this.lblCode3,
                  this.Label1,
                  this.mnuX,
                  this.mnuFile,
                  this.mnuEdit,
                  this.mnuTemplate,
                  this.mnuCategories,
                  this.mnuSpecial,
                  this.mnuSep7,
                  this.mnuHistory,
                  this.mnuFav,
                  this.mnuExternalFunctions,
                  this.mnuHelp
            });
            this.Name = "frmMain";
            this.frmTemplateInfo.ResumeLayout(false);
            this.frmOptions.ResumeLayout(false);
            this.Frame2.ResumeLayout(false);
            this.Frame1.ResumeLayout(false);
            this.frmFile.ResumeLayout(false);
            this.mnuFile.ResumeLayout(false);
            this.mnuEdit.ResumeLayout(false);
            this.mnuTemplate.ResumeLayout(false);
            this.mnuCategories.ResumeLayout(false);
            this.mnuSpecial.ResumeLayout(false);
            this.mnuSep7.ResumeLayout(false);
            this.mnuHistory.ResumeLayout(false);
            this.mnuFav.ResumeLayout(false);
            this.mnuExternalFunctions.ResumeLayout(false);
            this.mnuHelp.ResumeLayout(false);
            this.ResumeLayout(false);
        }
        #endregion

        public object m_sTemplateDatabaseName;
        public string m_sCurrentEventResponseCategory;
        public object CurrentCodeArea;
        public object Parent;
        public object m_oDBClassGen;
        public object m_asaHistory;
        public object m_asaAttributes;
        public object SliceAndDice;
        public object CurrentTemplate;
        public object InternalCurrentTemplate;
        public object Complete;
        public object SadCommands();
        public object SadCommandSetCount;
        public object FavoriteCount;
        public object ExternalCount;
        public object CurrentHistoryEntry;
        public object ActionToDo;
        public object ActionParam;
        public object mbScramFormKey;
        public object mbFillingAddInScreen;
        public object mbIgnoreBlanks;
        public object mbIgnoreReadOnly;
        public object OkayToDoAction;
        public object FavoriteCalledFromIDE;
        public object OkayToUnload;
        public object WithEvents;
        public object Externals;
        public object CurrExternal;
        public string sResult;
        public string sOperator;
        public string sLeftOfOp;
        public string sRightOfOp;
        public Variant vLeftOfOp;
        public Variant vRightOfOp;
        public int CurrSet;
        public CAssocArray asaTemp;
        public CAssocItem CurrAssocItem;
        public ISadAddin CurrDLL;
        public string sCategory;
        public string sShortName;
        public New asaVar;
        public New asaV;
        public CAssocItem CurItem;
        public int lLine;
        public int lTemp;
        public string sCodeToInsert;
        public string sProcName;
        public int lProcType;
        public string sProcTypeLong;
        public CAssocArray asaVar;
        public CAssocItem CurItem;
        public object CurFrame;
        public object CurForm;
        public object CurReference;
        public object CurControl;
        public object CurModule;
        public object ControlVars;
        public object tTemplate;
        public object tProject;
        public object tModule;
        public object tPane;
        public object tWindow;
        public object tWindows;
        public object tComponent;
        public int fh;
        public int CurrSet;
        public int CurrParam;
        public int lParamCount;
        public int lStartLine;
        public int lEndLine;
        public int lStartColFound;
        public int lEndColFound;
        public int lProcType;
        public MousePointerConstants lMouseState;
        public int lIfLoops;
        public int CodaIterations;
        public int NextElse;
        public int NextElseIf;
        public int NextEndIf;
        public int CmdIterations;
        public bool bFunction;
        public bool bFoundReference;
        public bool bDoCoda;
        public bool bT;
        public string sT;
        public string sHold1;
        public string sProcName;
        public string sProcType;
        public string sHold2;
        public string sHold3;
        public string sHold4;
        public string sCurParam;
        public string sCurType;
        public string CommandReference;
        public string sDelim1;
        public string sDelim2;
        public string sT2;
        public string scT2;
        public string scT3;
        public string scT1;
        public string scT4;
        public vbext_ProcKind ProcType;
        public Object lvwX;
        public Object tvwX;
        public Object tvwY;
        public CCategory CurrCategory;
        public CTemplate CurrTemplate;
        public string sOpened;
        public string sClosed;
        public int lCount;
        public int CurrSet;
        public int CurrFav;
        public CCategory CurrCategory;
        public CTemplate CurrTemplate;
        public string sPasswordCheck;
        public string sPasswordCheck;
        public 4) sCodeToCheck(0;
        public int CurCodeWindow;
        public int lTokens;
        public int CurToken;
        public int CurListItem;
        public string sCurToken;
        public CTemplate TemplateFound;
        public object sSelectedText;
        public object sOrigText;
        public object CurrSet;
        public string sCurrentCategory;
        public string sNewCategoryName;
        public string sCategoryToDuplicate;
        public string ColorSelected;
        public string ColorSelected;
        public string sFilename;
        public string sDate;
        public string PatchFilename;
        public int CurrSet;
        public CSadCommand CurrCommand;
        public string sChoices;
        public string sChoice;
        public string sPasswordCheck;
        public object sDatabasePath;
        public string sNewDatabaseName;
        public object db;
        public object tblTemplates;
        public object fldTemplates;
        public object ndxTemplates;
        public object rstCategory;
        public string sTemplateDatabaseName;
        public string sOldDatabaseName;
        public string sLastTemplate;
        public string sCategory;
        public string sShortName;
        public short Red;
        public short Green;
        public short Blue;
        public object sSelectedText;
        public object sOrigText;
        public object CurrSet;
        public int CurCodeArea;
        public int lLastFound;
        public bool bSomethingFound;
        public string sCategory;
        public string sShortName;
        public string sName;
        public 4) sCode(0;
        public object lLine;
        public As lLastLine;
        public object lTemp;
        public As lFirstCol;
        public object lLastCol;
        public object sCode;
        public As sProcName;
        public As lProcType;
        public object lLine;
        public As lLastLine;
        public As lFirstCol;
        public object lLastCol;
        public object sCode;
        public int lLine;
        public int lLastLine;
        public int lFirstCol;
        public int lLastCol;
        public string sCode;
        public int lLine;
        public int lLastLine;
        public int lFirstCol;
        public int lLastCol;
        public string sCode;
        public int lLine;
        public int lLastLine;
        public int lFirstCol;
        public int lLastCol;
        public string sCode;
        public int lLine;
        public int lLastLine;
        public int lFirstCol;
        public int lLastCol;
        public string sCode;
        public object lLine;
        public As lLastLine;
        public As lFirstCol;
        public object lLastCol;
        public object sCode;
        public string sTitle;
        public short Cancel;
        public object sKey;
        public string sRegValue;


                public object CurrentEventResponseCategory
    {
        get
        {
        if ( Len(m_sCurrentEventResponseCategory) == 0 )
            {
 m_sCurrentEventResponseCategory == "Event Response";
        CurrentEventResponseCategory = m_sCurrentEventResponseCategory;
        }

        set
        {
        m_sCurrentEventResponseCategory = value;
        }

    }


                public string CurrentTemplateNameAndCategory
    {
        get
        {
        CurrentTemplateNameAndCategory = txtName.Text;
        }

    }


                public object DBClassGen
    {
        get
        {
         DBClassGen = m_oDBClassGen;
        }

        set
        {
        try
{

         m_oDBClassGen = value;
        SetColors GetSetting$(App.ProductName, "Last", "Background Color", "&H80000018&"), GetSetting$(App.ProductName, "Last", "Foreground Color", "&H80000008&");
        ;
        }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        }

        }

    }


                public bool ExitAfterInsert
    {
        get
        {
        ExitAfterInsert = mnuExitAfterInsert.Checked;
        }

    }


                public object QueuedInsertions
    {
        set
        {
        // TODO: Rewrite try/catch and/or goto. EH_frmMain_QueuedInsertions
        Boolean static bInHereAlready;
        if ( bInHereAlready )
            {
 return; // ???
        bInHereAlready = true;
        asaVar.ItemDelimiter = "~";
        asaVar.All = value;
        foreach( var CurItem in asaVar.mCol )
        DoInsertion asaV, CurItem.Key;
        if ( gbCancelInsertion )
            {
 return; // ???
        } // CurItem
        EH_frmMain_QueuedInsertions_Continue:
        bInHereAlready = false;
        return; // ???
        EH_frmMain_QueuedInsertions:
        MsgBox "Error occured in:" + gsEolTab + "Module: frmMain" + gsEolTab + "Procedure: QueuedInsertions" + gs2EOL + Err.Description;
        goto EH_frmMain_QueuedInsertions_Continue;
        Resume;
        }

    }


                public string TemplateDatabaseName
    {
        get
        {
        TemplateDatabaseName = m_sTemplateDatabaseName;
        }

    }



            public As AddSadCommandSet            {
                try
{;
                ;

                SadCommandSetCount +=  1;
                ReDim Preserve SadCommands(1 To SadCommandSetCount);
                SadCommands(SadCommandSetCount) = oCommands;
                if ( SadCommands(SadCommandSetCount).Startup(Parent, Parent.vbInst) )
            {;
                AddSadCommand = true;
                frmSplash.lblDLLsLoaded(1).Text = string.Empty + SadCommandSetCount;
                frmSplash.lblDLLsLoaded(1).Refresh;
                Externals = oCommands.Externals;
                if ( ! Externals Is null )
            {;
                if ( ExternalCount > 0 )
            {;
                Load mnuExternals(ExternalCount);
                mnuExternals(ExternalCount).Text = "-";
                mnuExternals(ExternalCount).Tag = string.Empty;
                mnuExternals(ExternalCount).Enabled = true;
                mnuExternals(ExternalCount).Visible = true;
                ExternalCount +=  1;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void DeleteTemplate            {
                // TODO: Rewrite try/catch and/or goto. EH_frmMain_DeleteTemplate;
                Boolean static bInHereAlready;
                if ( bInHereAlready )
            {
 return;
                bInHereAlready = true;

                if ( CurrentTemplate Is null )
            {;
                if ( bAutoDelete )
            {;
                MsgBox "DeleteTemplate failed because nothing is selected.";
                }
            else
            {;
                MsgBox "Please select a " + gsTemplate + " to delete first.";
            }

            public As Evaluate            {
                try
{;


                if ( InStr(sExpression), " AND ") .ToUpper()
            {;
                vLeftOfOp = InStr(sExpression), " AND ".ToUpper();
                sExpression.Substring(0, vLeftOfOp - 1);
                sRightOfOp = sExpression.Substring( vLeftOfOp + 5);
                sLeftOfOp = Evaluate(Trim$(sLeftOfOp), asaVar);
                sRightOfOp = Evaluate(Trim$(sRightOfOp), asaVar);
                Evaluate = IIf((Val(sLeftOfOp) <> 0) And (Val(sRightOfOp) <> 0), "1", "0");
                return; // ???;

                }
            else
            {if ( InStr(sExpression), " OR ") .ToUpper()
            {;
                vLeftOfOp = InStr(sExpression), " OR ".ToUpper();
                sExpression.Substring(0, vLeftOfOp - 1);
                sRightOfOp = sExpression.Substring( vLeftOfOp + 5);
                sLeftOfOp = Evaluate(Trim$(sLeftOfOp), asaVar);
                sRightOfOp = Evaluate(Trim$(sRightOfOp), asaVar);
                Evaluate = IIf((Val(sLeftOfOp) <> 0) Or (Val(sRightOfOp) <> 0), "1", "0");
                return; // ???;

                }
            else
            {if ( InStr(sExpression), " XOR ") .ToUpper()
            {;
                vLeftOfOp = InStr(sExpression), " XOR ".ToUpper();
                sExpression.Substring(0, vLeftOfOp - 1);
                sRightOfOp = sExpression.Substring( vLeftOfOp + 5);
                sLeftOfOp = Evaluate(Trim$(sLeftOfOp), asaVar);
                sRightOfOp = Evaluate(Trim$(sRightOfOp), asaVar);
                Evaluate = IIf((Val(sLeftOfOp) <> 0) Xor (Val(sRightOfOp) <> 0), "1", "0");
                return; // ???;

                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void FillAddInScreen()
            {
                // TODO: Rewrite try/catch and/or goto. EH_frmMain_FillAddInScreen;
                Boolean static bInHereAlready;
                if ( bInHereAlready )
            {
 return;
                bInHereAlready = true;
                mbFillingAddInScreen = true;
                if ( CurrentTemplate Is null )
            {
 return;

                txtName = CurrentTemplate.Key;
                txtShortName = CurrentTemplate.ShortTemplateName;

                txtCode(0) = CurrentTemplate.memoCodeAtTop;
                txtCode(1) = CurrentTemplate.memoCodeAtCursor;
                txtCode(2) = CurrentTemplate.memoCodeAtBottom;

                txtFilename = CurrentTemplate.FileName;
                txtCodeToFile = CurrentTemplate.memoCodeToFile;

                chkUndeletable = Abs(.Undeletable);
                chkLocked = Abs(.Locked);
                chkFavorite = Abs(.Favorite);
                mnuIsFavorite.Checked = CurrentTemplate.Favorite;
                chkSelected = Abs(.Selected);
                lblRevision.Text = "Revision #: " + CurrentTemplate.Revision;
                lblAlpha.Text = "Alpha Date: " + Format$(.DateCreated, "Mmmm D, YYYY H:NN:SS AM/PM");
                lblDelta.Text = "Delta Date: " + Format$(.DateModified, "Mmmm D, YYYY H:NN:SS AM/PM");
                //        '         With SliceAndDice.SystemInfo("Hotkey Templates").Item(.Key);
                //        '              If Len(.Value) Then;
                //        '                 hkyInstantInsert.HotKeyModifier = Val(sGetToken(.Value, 2, gsC));
                //        '                 hkyInstantInsert.HotKey = Val(sGetToken(.Value, 1, gsC));
                //        '              Else;
                //        '                 hkyInstantInsert.HotKeyModifier = HOTKEYF_EXT;
                //        '                 hkyInstantInsert.HotKey = 0;
                //        '              End If;
                //        '         End With;
            }

            public void GetCategoryAndName            {
                if ( lTokenCount(sCategoryAndName, gsCategoryTemplateDelimiter) < 2 )
            {;
                sCategory = "Unknown";
                sShortName = sCategoryAndName;
                }
            else
            {;
                sCategory = sGetToken(sCategoryAndName, 1, gsCategoryTemplateDelimiter);
                sShortName = sAfter(sCategoryAndName, 1, gsCategoryTemplateDelimiter);
                if ( Len(sShortName) = 0 )
            {;
                sCategory = "Unknown";
                sShortName = sCategoryAndName;
            }

            public void HandleIDEEvents            {
                // On Error GoTo EH_frmMain_HandleIDEEvents;
                // On Error goto Next;
                // EH_frmMain_HandleIDEEvents_Continue:;
                // EH_frmMain_HandleIDEEvents:;
            }

            public void HideAllWindows            {
                try
{;

                if ( ! m_oDBClassGen Is null )
            {;
                m_oDBClassGen.Hide;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As InitializeAddinDLLs            {

                ShutdownDLLs;

                if ( Len(sAddinList) = 0 )
            {;
                InitializeAddinDLLs = true;
                return; // ???;
            }

            public void NewTemplate            {
                // TODO: Rewrite try/catch and/or goto. EH_frmMain_NewTemplate;

                if ( Len(sTitle) = 0 )
            {;
                sCategory = lsbJumpTo.BarKey;
                if ( Len(sDefaultShortName) = 0 )
            {;
                sDefaultShortName = Abs(NextNegativeUnique());
            }

            public void QueueAction            {
                OkayToDoAction = false;
                ActionToDo = sAction;
                ActionParam = sParam;
                tmrDoAction.Interval = IIf(Interval > 65535, 65535, IIf(Interval < 100, 100, Interval));
                tmrDoAction.Enabled = true;
            }

            public As RefreshDatabaseConnection()
            {
                // TODO: Rewrite try/catch and/or goto. EH_frmMain_RefreshDatabaseConnection;

                Call NextNegativeUnique;

                CurrentTemplate = null;
                InternalCurrentTemplate = null;
                SliceAndDice = null;

                SliceAndDice = new CSliceAndDice();
                if ( ! SliceAndDice.Load(m_sTemplateDatabaseName) )
            {;
                RefreshDatabaseConnection = false;
                return; // ???;
            }

            public void DoInsertion            {
                // TODO: Rewrite try/catch and/or goto. EH_frmMain_DoInsertion;
                Boolean static bInHereAlready;
                if ( bInHereAlready )
            {
 return;
                bInHereAlready = true;

                lsbJumpTo.Enabled = false;

                gbCancelInsertion = false;
                mbIgnoreBlanks = false;
                mbIgnoreReadOnly = false;



                //    'If txtName <> sTemplateToInsert Then;
                if ( ! SetInternalCurrentTemplate(sTemplateToInsert) )
            {;
                LogError "frmMain", "DoInsertion", vbObjectError + 100, "Can't find the " + gsTemplate + gsS + gsA + sTemplateToInsert + "' to insert." + gsEolTab + "Aborting this insertion.", Erl;
                GoTo EH_frmMain_DoInsertion_Continue;
            }

            public As FillTemplateWithUserInput            {
                string static sVarName;
                string static sVarPhrase;
                string static sDefault;
                string static sT;
                string static sVar1;
                string static sVar2;
                string static sVar3;
                string static sNow;
                long static lParamCount;
                long static Curr;
                Boolean static bInlineCommandExecuted;

                sGetToken(sToParse, 1, vbNewLine).Contains(gsSoftVarDelimiter) > 0    ' For each soft variable found) {;
                sVarPhrase = sGetToken(sToParse, 2, gsSoftVarDelimiter)    ' Get the Variable name and default if provided;
                sVarName = sGetToken(sVarPhrase, 1, gsInlineCmdDelimiter)    ' Extract just the variable name;
                sNow = string.Empty;
                bInlineCommandExecuted = false;
                if ( SadCommandSetCount > 0 )
            {;
                sVar1 = sAfter(sVarPhrase, 1, gsInlineCmdDelimiter);
                for(var Curr = 1; Curr < SadCommandSetCount; Curr++)  {;
                if ( SadCommands(CurrSet).ExecuteSoftCommandInline(asaX, sVarName), sVar1, sNow) .ToUpper()
            {;
                bInlineCommandExecuted = true;
                Exit For;
            }

            public As InternalInsertTemplate            {







                //    ' For use by internal commands only (not outside select case statment);


                // TODO: Rewrite try/catch and/or goto. EH_InsertTemplate;

                if ( II Is null )
            {;
                InternalInsertTemplate = true;
                GoTo EH_InsertTemplate_Continue;
            }

            public void GetProcAtLine            {

                if ( ! Parent.HostedByVB )
            {
 return;   ' Shell override;

                if ( lCurrentLine < 1 Or lCurrentLine > Parent.vbInst.ActiveCodePane.CodeModule.CountOfLines )
            {;
                sProcName = string.Empty;
                lProcType = 0;
                return;
            }

            public As FindLastProcLine            {
                long static lLine          ;
                long static lCurLine       ;
                long static lLastLine      ;
                string static sFindstring    ;
                string static sFunctionHeader;

                if ( ! Parent.HostedByVB )
            {
 return; // ???   ' Shell override;

                try
{;


                lLine = Parent.vbInst.ActiveCodePane.CodeModule.ProcStartLine(sProcName, lProcType)  ' Get the first line number of the procedure;
                if ( ex <> 0 )
            {;
                if ( InStr(sProcName, "_") )
            {;

                lLine = Parent.vbInst.ActiveCodePane.CodeModule.CreateEventProc(sGetToken(sProcName, 2, "_"), sGetToken(sProcName, 1, "_"));
                if ( ex <> 0 )
            {;
                MsgBox "FindLastProcLine" + gsEolTab + "Can't find:" + gsEolTab + vbTab + sProcName + gsEolTab + "In module:" + gsEolTab + vbTab + Parent.vbInst.ActiveCodePane.CodeModule.Parent.Name;
                gbCancelInsertion = bUserSure("Cancel processing ?");
                FindLastProcLine = 0;

                return; // ???;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As JumpTo            {
                // TODO: Rewrite try/catch and/or goto. frmMain_EH_JumpTo;
                string static sCategoryName;
                string static sShortTemplateName;
                long static CurrHE;


                SaveTemplate;

                sCategoryName = sGetToken(sTemplateName, 1, gsCategoryTemplateDelimiter);
                sShortTemplateName = sAfter(sTemplateName, 1, gsCategoryTemplateDelimiter);
                try
{;
                if ( SliceAndDice.Categorys(sCategoryName).Templates(sTemplateName) Is null )
            {;
                JumpTo = false;
                return; // ???;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void RefillList()
            {
                // TODO: Rewrite try/catch and/or goto. EH_frmMain_RefillList;
                Boolean static bInHereAlready;
                if ( bInHereAlready )
            {
 return;
                bInHereAlready = true;


                lsbJumpTo.Visible = false;
                lsbJumpTo.Clear;
                foreach( var CurrCategory in SliceAndDice.Categorys );

                if ( CurrCategory.Deleted )
            {;
                //                ' Ignore this one;
                //                'Else;
                }
            else
            {if ( CurrCategory.CategoryType = 0 )
            {;
                if ( lsbJumpTo.BarKey = "Bar 1" )
            {;
                lsbJumpTo.CurBar = 0;
                lsbJumpTo.BarName = CurrCategory.Key + " (" + Format$(CurrCategory.Templates.Count, "00") + gsPC;
                lsbJumpTo.BarKey = CurrCategory.Key;
                lsbJumpTo.View = 3;
                lsbJumpTo.Arrange = CurrCategory.Arrange;
                lsbJumpTo.BarType = "List";
                try
{;
                lsbJumpTo.Bars(0).ColumnHeaders(1).Width = 3400;
                }
            else
            {;
                lsbJumpTo.AddBar(.Key + " (" + Format$(CurrCategory.Templates.Count, "00") + gsPC, CurrCategory.Key).ColumnHeaders(1).Width = 3400;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void SetColors            {
                try
{;
                BackColor.Substring(BackColor.Length - 1) == "&" )
            {
BackColor.Substring(0, Len(BackColor) - 1);
                ForeColor.Substring(ForeColor.Length - 1) == "&" )
            {
ForeColor.Substring(0, Len(ForeColor) - 1);

                BackColor.Substring(0, 2) <> "&H" )
            {
 BackColor == "&H" + BackColor;
                ForeColor.Substring(0, 2) <> "&H" )
            {
 ForeColor == "&H" + ForeColor;

                lsbJumpTo.BackColor = BackColor;
                txtCode(0).BackColor = BackColor;
                txtCode(1).BackColor = BackColor;
                txtCode(2).BackColor = BackColor;
                txtCodeToFile.BackColor = BackColor;
                lstSoftCommands.BackColor = BackColor;
                lstSoftVariables.BackColor = BackColor;

                lsbJumpTo.ForeColor = ForeColor;
                txtCode(0).ForeColor = ForeColor;
                txtCode(1).ForeColor = ForeColor;
                txtCode(2).ForeColor = ForeColor;
                txtCodeToFile.ForeColor = ForeColor;
                lstSoftCommands.ForeColor = ForeColor;
                lstSoftVariables.ForeColor = ForeColor;

                if ( ! m_oDBClassGen Is null )
            {;

                m_oDBClassGen.lvwFields.BackColor = BackColor;
                m_oDBClassGen.lvwFields.ForeColor = ForeColor;
                m_oDBClassGen.dvwTable.BackColor = BackColor;
                m_oDBClassGen.dvwTable.ForeColor = ForeColor;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As SetInternalCurrentTemplate            {
                try
{;
                string static sCategoryName;
                string static sShortTemplateName;


                SaveTemplate;

                sCategoryName = sGetToken(sTemplateName, 1, gsCategoryTemplateDelimiter);
                sShortTemplateName = sAfter(sTemplateName, 1, gsCategoryTemplateDelimiter);

                if ( SliceAndDice(sCategoryName).Templates(sTemplateName) Is null )
            {;
                SetInternalCurrentTemplate = false;
                }
            else
            {;
                InternalCurrentTemplate = SliceAndDice(sCategoryName).Templates(sTemplateName);
                SetInternalCurrentTemplate = true;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As sGetCurrentLineAtCharacter            {
                sTextToSearch, lCharToStart).Substring(0, vbNewLine);
                if ( lCount > 0 )
            {;
                sGetCurrentLineAtCharacter = sGetToken(sTextToSearch, lCount, vbNewLine);
                }
            else
            {;
                sGetCurrentLineAtCharacter = sTextToSearch;
            }

            public void ShowExternalsMenu()
            {
                try
{;
                PopupMenu mnuExternalFunctions;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void ShowFavMenu()
            {
                try
{;
                PopupMenu mnuFav;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As ShutdownDLLs()
            {
                try
{;

                for(var Curr = 1; Curr < SadCommandSetCount; Curr++)  {;
                Call SadCommands(CurrSet).Shutdown;
                SadCommands(CurrSet) = null;
                } // CurrSet;
                ReDim SadCommands(1 To 1);
                SadCommandSetCount = 0;

                ShutdownDLLs = true;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void SaveTemplate()
            {
                // TODO: Rewrite try/catch and/or goto. EH_frmMain_SaveTemplate;
                Boolean static bInHereAlready;
                if ( bInHereAlready )
            {
 return;
                bInHereAlready = true;

                MSComctlLib.ListView static CurrBar;

                if ( ! CurrentTemplate Is null )
            {;

                if ( ! CurrentTemplate.Deleted )
            {;
                if ( CurrentTemplate.ParentKey = string.Empty )
            {;
                MsgBox "frmMain.SaveTemplate : Error found. ParentKey blank";
                GoTo EH_frmMain_SaveTemplate_Continue;
            }

            public void UpdateFavorites()
            {
                try
{;

                DoEvents: DoEvents: DoEvents;

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

            public As sTemplateIcon            {
                if ( CurrTemplate Is null )
            {;
                sTemplateIcon = "!";
                }
            else
            {if ( Len(CurrTemplate.memoCodeAtBottom + CurrTemplate.memoCodeAtCursor + CurrTemplate.memoCodeAtTop + CurrTemplate.memoCodeToFile) > 0 )
            {;
                sTemplateIcon = gsCategory;
                }
            else
            {if ( CurrTemplate.Selected )
            {;
                sTemplateIcon = "Check";
                }
            else
            {if ( CurrTemplate.Undeletable Or CurrTemplate.Locked )
            {;
                sTemplateIcon = "Key";
                }
            else
            {;
                sTemplateIcon = "Document";
            }

            public void UpdateHotKeys()
            {
                // On Error GoTo EH_UpdateHotKeys;
                // EH_UpdateHotKeys_Continue:;
                // EH_UpdateHotKeys:;
            }

            public void chkAutoRecalc_Click()
            {
                SaveSetting App.ProductName, "Last", "Auto Recalc", chkAutoRecalc.Value;
            }

            public void chkFavorite_Click()
            {
                if ( mbFillingAddInScreen )
            {
 return;
                CurrentTemplate.Favorite = (chkFavorite.Value <> 0);
                mnuIsFavorite.Checked = CurrentTemplate.Favorite;
                UpdateFavorites;
            }

            public void chkUndeletable_Click()
            {
                Boolean static bInHereAlready;

                if ( mbFillingAddInScreen )
            {
 return;
                if ( bInHereAlready )
            {
 return;
                if ( ! mnuPasswordProtection.Checked )
            {;
                lsbJumpTo.BarItemIcon = IIf(chkLocked.Value = 0 And chkUndeletable.Value = 0, "Document" + IIf(lsbJumpTo.BarItemIcon = "DocumentAlternate", "Alternate", string.Empty), "Key");
                return;
            }

            public void chkUndeletable_Validate            {
            }

            public void cmdRecalc_Click()
            {
                // TODO: Rewrite try/catch and/or goto. EH_cmdRecalc_Click;


                Screen.MousePointer = vbHourglass;

                lstSoftVariables.Clear;
                lstSoftCommands.Clear;
                //    'MsgBox "Recalc to occur here.";

                sCodeToCheck(0) = txtCode(1);
                sCodeToCheck(1) = txtCode(0);
                sCodeToCheck(2) = txtCode(2);


                switch txtShortName.ToUpper();
                Case "COLLECTION", "COLLECTION, NO CHILD", "COLLECTION, NO PARENT", "COLLECTION, NO PARENT, NO CHILD";
                if ( InStr(txtShortName), "NO PARENT") = 0 .ToUpper()
            {;
                lstSoftVariables.AddItem("* Parent AutoNumber Field Name");
                lstSoftVariables.AddItem("* Parent AutoNumber Property Name");
            }

            public void Form_LostFocus()
            {
                try
{;
                lsbJumpTo.HideCategories;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Form_QueryUnload            {
                if ( UnloadMode = vbFormControlMenu )
            {;
                Cancel = true;
                mnuFileExit_Click;
                }
            else
            {;
                Form_Unload Cancel;
            }

            public void lsbJumpTo_AfterBarClick()
            {
                try
{;
                JumpTo lsbJumpTo.BarKey + gsCategoryTemplateDelimiter + lsbJumpTo.BarItemName;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void lsbJumpTo_BarItemClick            {
                try
{;
                if ( lsbJumpTo.BarType = "List" )
            {;
                JumpTo BarItemKey;
                }
            else
            {;
                TemplateFound = SliceAndDice.Categorys.ItemByLongTemplateName(BarItemKey);
                if ( TemplateFound Is null )
            {;
                Beep;
                if ( MsgBox("That " + gsTemplate + " does not exist (yet)." + gsEolTab + "Create " + gsTemplate + " now ?", vbYesNo, "NO " + gsTemplate + ": " + BarItemKey) = vbYes )
            {;
                QueueAction "NewTemplate", BarItemKey;
                OkayToDoAction = true;
                }
            else
            {;
                if ( Val(CurrentHistoryEntry) > 0 )
            {;
                JumpTo m_asaHistory(CurrentHistoryEntry), false, true;
                }
            else
            {if ( SliceAndDice.Categorys(sGetToken(BarItemKey, 1, gsCategoryTemplateDelimiter)).Templates.Count > 1 )
            {;
                JumpTo SliceAndDice.Categorys(sGetToken(BarItemKey, 1, gsCategoryTemplateDelimiter)).Templates(1);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void lsbJumpTo_BarItemDblClick            {
                if ( Len(BarItemKey) == 0 )
            {
 return;

                mnuInsertTemplate_Click;
            }

            public void lsbJumpTo_KeyDown            {
                // TODO: Rewrite try/catch and/or goto. EH_frmMain_lsbJumpTo_KeyDown;


                if ( (Shift And vbShiftMask) > 0 )
            {
               ' Shift Key         *******************;
                switch KeyCode;
                ;
                Case vbKeyInsert: KeyCode = 0: Shift = 0  ' Paste;
                if ( TypeOf ActiveControl Is TextBox )
            {;
                ActiveControl.SelText = Clipboard.GetText;
            }

            public void lsbJumpTo_LostFocus()
            {
                try
{;
                lsbJumpTo.HideCategories;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void lsbJumpTo_MouseDown            {
                if ( Button == vbRightButton And Shift == 0 )
            {
      ' Right click, pop-up menu;
                PopupMenu mnuTemplate;
            }

            public void lsbJumpTo_MouseDownOnCategory            {
                if ( Button == vbRightButton And Shift == 0 )
            {
      ' Right click, pop-up menu;
                PopupMenu mnuCategories;
            }

            public void lstSoftCommands_DblClick()
            {
                try
{;
                lstSoftVariables.ListIndex = -1;
                txtCode(0).SelStart = 0: txtCode(0).SelLength = 0;
                txtCode(1).SelStart = 0: txtCode(1).SelLength = 0;
                txtCode(2).SelStart = 0: txtCode(2).SelLength = 0;
                FindInCurrent false, false, true;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void lstSoftVariables_DblClick()
            {
                try
{;
                lstSoftCommands.ListIndex = -1;
                txtCode(0).SelStart = 0: txtCode(0).SelLength = 0;
                txtCode(1).SelStart = 0: txtCode(1).SelLength = 0;
                txtCode(2).SelStart = 0: txtCode(2).SelLength = 0;
                FindInCurrent false, false, true;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuBack_Click()
            {
                try
{;
                if ( Val(CurrentHistoryEntry) < 2 )
            {;
                Beep;
                mnuBack.Enabled = false;
                return;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuCategoriesDeleteCurrent_Click()
            {
                // TODO: Rewrite try/catch and/or goto. EH_mnuCategoriesDeleteCurrent_Click;
                Boolean static bInHereAlready;

                if ( bInHereAlready )
            {
 return;
                bInHereAlready = true;

                SaveTemplate;

                sCurrentCategory = lsbJumpTo.BarKey;

                if ( sCurrentCategory) = "CHANGE FROM" .ToUpper()
            {;
                MsgBox "The 'Change From' " + gsCategory + " is not removable.", vbExclamation;
                GoTo EH_mnuCategoriesDeleteCurrent_Click_Continue;
                return;
                }
            else
            {if ( SliceAndDice.Categorys(sCurrentCategory).CategoryType <> 0 )
            {;
                if ( ! bUserSure("This " + gsCategory + " is used by the code generators. Deleting it is unadvisable." + gs2EOLTab + "Are you sure you want to permanently erase this " + gsCategory + " ?") )
            {;
                GoTo EH_mnuCategoriesDeleteCurrent_Click_Continue;
                return;
            }

            public void mnuCategoriesNewMethod_Click            {

                switch Index;
                Case 0                                        ' New, Blank Category;
                sNewCategoryName = InputBox("What should the name of the new, blank " + gsCategory + " be ?", "NEW " + gsCategory, string.Empty);
                if ( Len(snewCategoryName) == 0 )
            {
 return;
                if ( SliceAndDice(sNewCategoryName) Is null )
            {;
                SliceAndDice.Categorys.Add(                                                                                                   sNewCategoryName);
                SliceAndDice.Save;
                RefillList;
                }
            else
            {;
                MsgBox "There is already a " + gsCategory + " by that name. Aborting.", vbInformation;
            }

            public void mnuChangeBackgroundColors_Click()
            {
                ColorSelected = sChooseColor(lsbJumpTo.BackColor);
                if ( Len(ColorSelected) == 0 )
            {
 return;

                SaveSetting App.ProductName, "Last", "Background Color", ColorSelected;
                SetColors ColorSelected, GetSetting$(App.ProductName, "Last", "Foreground Color", "&H80000008&");
            }

            public void mnuChangeForegroundColor_Click()
            {
                ColorSelected = sChooseColor(lsbJumpTo.ForeColor);
                if ( Len(ColorSelected) == 0 )
            {
 return;
                SaveSetting App.ProductName, "Last", "Foreground Color", ColorSelected;

                SetColors GetSetting$(App.ProductName, "Last", "Background Color", "&H80000018&"), ColorSelected;
            }

            public void mnuEditCopy_Click()
            {
                if ( ! chkLocked )
            {;
                switch tabCode.SelectedItem.Index;
                Case 1: stringToClipboard txtCode(0).SelText;
                Case 2: stringToClipboard txtCode(1).SelText;
                Case 3: stringToClipboard txtCode(2).SelText;
                Case 4: stringToClipboard txtCode(3).SelText;
            }

            public void mnuEditCut_Click()
            {
                if ( ! chkLocked )
            {;
                switch tabCode.SelectedItem.Index;
                Case 1: if ( stringToClipboard(txtCode(0).SelText) )
            {
 txtCode(0).SelText == string.Empty;
                Case 2: if ( stringToClipboard(txtCode(1).SelText) )
            {
 txtCode(1).SelText == string.Empty;
                Case 3: if ( stringToClipboard(txtCode(2).SelText) )
            {
 txtCode(2).SelText == string.Empty;
                Case 4: if ( stringToClipboard(txtCode(3).SelText) )
            {
 txtCode(3).SelText == string.Empty;
            }

            public void mnuEditFind_Click()
            {
                FindInCurrent;
            }

            public void mnuEditPaste_Click()
            {
                if ( ! chkLocked )
            {;
                switch tabCode.SelectedItem.Index;
                Case 1: txtCode(0).SelText = Clipboard.GetText;
                Case 2: txtCode(1).SelText = Clipboard.GetText;
                Case 3: txtCode(2).SelText = Clipboard.GetText;
                Case 4: txtCode(3).SelText = Clipboard.GetText;
            }

            public void mnuEditReplace_Click()
            {
                FindInCurrent false, true;
            }

            public void mnuExternals_Click            {
                try
{;
                SadCommands(Val(sGetToken(mnuExternals(Index).Tag, 1, "|"))).ExecuteExternal mnuExternals(Index).Text, sAfter(mnuExternals(Index).Tag, 1, "|");
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuFavorite_Click            {
                if ( FavoriteCalledFromIDE )
            {;
                FavoriteCalledFromIDE = false;
                DoInsertion null, mnuFavorite(Index).Text;
                }
            else
            {;
                JumpTo mnuFavorite(Index).Text, , true;
                lsbJumpTo.HideCategories;
            }

            public void mnuFileApplyDeltaPatch_Click()
            {
                sFilename = sChooseFile(, , "Sandy Delta Patch (*.sad)|*.sad|All Files (*.*)|*.*");
                if ( Len(sFilename) )
            {;
                SliceAndDice.ApplyPatch sFilename;
            }

            public void mnuFileGenerateDeltaPatch_Click            {

                sDate = SliceAndDice.sChoosePatch(Index);
                if ( Len(sDate) )
            {;
                App.Path, 1) <> gsBS, gsBS.Substring(App.Path, 1) <> gsBS, gsBS.Length - string.Empty) + "MDBPatch" + Replace(Format$(sDate, "00000.00"), gsP, "-") + ".sad";
                SliceAndDice.GenerateDeltaPatchFile CVDate(sDate), PatchFilename;
                if ( Len(Dir$(PatchFilename)) )
            {;
                if ( bUserSure("File created successfully." + gsEolTab + "Filename:" + PatchFilename + gs2EOL + "Would you like to view it now ?") )
            {;
                try
{;
                Shell WindowsDirectory + "NOTEPAD.EXE " + gsQ + PatchFilename + gsQ;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuForward_Click()
            {
                if ( Val(CurrentHistoryEntry) >= m_asaHistory.Count )
            {;
                Beep;
                mnuForward.Enabled = false;
                return;
            }

            public void mnuHelpAbout_Click()
            {
                try
{;

                frmSplash.lblDLLsLoaded(1).Text = string.Empty + SadCommandSetCount;
                frmSplash.Show;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuHelpEmailWilliamRawls_Click()
            {
                BrowseTo "mailto:wrawls@firmsolutions.com";
            }

            public void mnuHelpReportIssue_Click()
            {
                BrowseTo "http://www.sliceanddice.com/sadissue.html";
            }

            public void mnuHelpSoftCommandReference_Click()
            {
                if ( SadCommandSetCount > 0 )
            {;
                Complete.ShowHelpScreen;
                // End If;
                }
            else
            {;
                the frmSplash.MDB you have loaded." MsgBox "No command set DLLs loaded." + gsEolTab + "No Soft Command Reference available." + gsEolTab + "Make sure S&D DLLs are in the same directory;
            }

            public void UpdateCompleteListOfSoftCommands()
            {
                try
{;

                if ( ! Complete Is null )
            {;
                Complete.this.Clear()false;
                Complete.Parent = null;
                Complete = null;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuHelpVisitHomePage_Click()
            {
                BrowseTo "http://www.sliceanddice.com";
            }

            public void mnuHistoryList_Click()
            {
                try
{;

                if ( m_asaHistory.Count > 0 )
            {;
                m_asaHistory.ItemDelimiter = gsSC;
                sChoices = m_asaHistory.Column;
                sChoice = sChoose(sChoices, , m_asaHistory(CurrentHistoryEntry).Value);
                if ( Len(sChoice) )
            {;
                CurrentHistoryEntry = string.Empty + m_asaHistory.FindKey(sChoice);
                JumpTo m_asaHistory(CurrentHistoryEntry), false, true;
                mnuForward.Enabled = Val(CurrentHistoryEntry) < m_asaHistory.Count;
                mnuBack.Enabled = Val(CurrentHistoryEntry) > 1;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuHelpOnlineDocumentation_Click()
            {
                BrowseTo "http://www.sliceanddice.com/saddoc.html";
            }

            public void mnuIsFavorite_Click()
            {
                if ( mbFillingAddInScreen )
            {
 return;
                chkFavorite = Abs(! -chkFavorite);
                chkFavorite_Click;
            }

            public void mnuPasswordProtection_Click()
            {
                mnuPasswordProtection.Checked = ! mnuPasswordProtection.Checked;
                SaveSetting App.ProductName, "Last", "Password Protection", mnuPasswordProtection.Checked;
            }

            public void mnuShowOnModuleRightClick_Click()
            {
                mnuShowOnModuleRightClick.Checked = ! mnuShowOnModuleRightClick.Checked;
                SaveSetting App.ProductName, "Last", "Show On Module Right Click", mnuShowOnModuleRightClick.Checked;
                MsgBox "This will take effect the next time Visual Basic or " + gsSliceAndDice + " is restarted.", vbInformation;
            }

            public void mnuShowPaintbrushIcon_Click()
            {
                mnuShowPaintbrushIcon.Checked = ! mnuShowPaintbrushIcon.Checked;
                SaveSetting App.ProductName, "Last", "Show Paitbrush Icon", mnuShowPaintbrushIcon.Checked;
                MsgBox "This will take effect the next time Visual Basic is restarted.", vbInformation;
            }

            public void mnuOLEDragDrop_Click()
            {
                try
{;
                mnuOLEDragDrop.Checked = ! mnuOLEDragDrop.Checked;
                SaveSetting App.ProductName, gsLast, "OLEDragDrop", mnuOLEDragDrop.Checked;

                txtCode(0).OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0);
                txtCode(0).OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0);
                txtCode(1).OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0);
                txtCode(1).OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0);
                txtCode(2).OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0);
                txtCode(2).OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0);
                txtCodeToFile.OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0);
                txtCodeToFile.OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuShowSplash_Click()
            {
                mnuShowSplash.Checked = ! mnuShowSplash.Checked;
                SaveSetting App.ProductName, "Last", "Show Splash", mnuShowSplash.Checked;
            }

            public void mnuSwitchTabsAutomatically_Click()
            {
                mnuSwitchTabsAutomatically.Checked = ! mnuSwitchTabsAutomatically.Checked;
                SaveSetting App.ProductName, "Last", "Switch tabs automatically", mnuSwitchTabsAutomatically.Checked;
            }

            public void mnuX_Click()
            {
                mnuFileExit_Click;
            }

            public void tmrActivateDBClassGen_Timer()
            {
                if ( gbProcessing )
            {
 return;
                tmrActivateDBClassGen.Enabled = false;

                DBClassGen.RefreshCategories;
                DBClassGen.Show , Me;
            }

            public void chkLocked_Click()
            {
                Boolean static bInHereAlready;

                if ( mbFillingAddInScreen )
            {
 return;
                if ( bInHereAlready )
            {
 return;
                if ( ! mnuPasswordProtection.Checked )
            {;
                lsbJumpTo.BarItemIcon = IIf(chkLocked.Value = 0 And chkUndeletable.Value = 0, "Document" + IIf(lsbJumpTo.BarItemIcon = "DocumentAlternate", "Alternate", string.Empty), "Key");
                return;
            }

            public void mnuSpecialNewDatabase_Click()
            {

                sDatabasePath = Trim$(BrowseForFolder(DBClassGen.hwnd, "Where should database go ?"));
                if ( Len(sDatabasePath) == 0 )
            {
 return;

                sNewDatabaseName = Trim$(InputBox("What should the name of the new " + gsTemplate + " database be ?", "CREATE " + gsTemplate + " DATABASE"));
                if ( Len(snewDatabaseName) == 0 )
            {
 return;

                sDatabasePath.Substring(sDatabasePath.Length - 1) <> gsBS )
            {
 sDatabasePath == sDatabasePath + gsBS;
                LCase$(snewDatabaseName).Substring(LCase$(snewDatabaseName).Length - 4) <> ".mdb" )
            {
 snewDatabaseName == sDatabasePath + snewDatabaseName + ".mdb";

                // TODO: Rewrite try/catch and/or goto. mnuSpecialNewDatabase_Click;

                db = CreateDatabase(sNewDatabaseName, dbLangGeneral, dbVersion30);
                if ( ex <> 0 )
            {;
                MsgBox "Error creating " + gsTemplate + " database. Aborting.";
                return;
            }

            public void mnuSpecialOpenDatabase_Click()
            {

                sTemplateDatabaseName = sChooseDatabase(App.Path);
                if ( Len(sTemplateDatabaseName) )
            {;
                SaveTemplate;
                sOldDatabaseName = m_sTemplateDatabaseName;
                m_sTemplateDatabaseName = sTemplateDatabaseName;
                if ( RefreshDatabaseConnection )
            {;
                SaveSetting App.ProductName, "Settings", "Current Database", sTemplateDatabaseName;
                if ( ! SliceAndDice(1) Is null )
            {;
                if ( ! SliceAndDice(1).Templates(1) Is null )
            {;
                JumpTo SliceAndDice(1).Templates(1).Key;
                lsbJumpTo.BarAndItem SliceAndDice(1).Key, SliceAndDice(1).Templates(1).ShortTemplateName;
                }
            else
            {if ( ! SliceAndDice(2) Is null )
            {;
                if ( ! SliceAndDice(2).Templates(1) Is null )
            {;
                JumpTo SliceAndDice(2).Templates(1).Key;
                lsbJumpTo.BarAndItem SliceAndDice(2).Key, SliceAndDice(2).Templates(1).ShortTemplateName;
            }

            public void tabCode_MouseUp            {
                // TODO: Rewrite try/catch and/or goto. EH_frmMain_tabCode_MouseUp;
                Boolean static bInHereAlready;

                switch tabCode.SelectedItem.Index;
                Case 1;
                txtCode(0).Visible = true;
                txtCode(1).Visible = false;
                txtCode(2).Visible = false;
                frmFile.Visible = false;
                frmOptions.Visible = false;
                frmTemplateInfo.Visible = false;
                try
{;
                txtCode(0).SetFocus;

                Case 2;
                txtCode(0).Visible = false;
                txtCode(1).Visible = true;
                txtCode(2).Visible = false;
                frmFile.Visible = false;
                frmOptions.Visible = false;
                frmTemplateInfo.Visible = false;
                try
{;
                txtCode(1).SetFocus;

                Case 3;
                txtCode(0).Visible = false;
                txtCode(1).Visible = false;
                txtCode(2).Visible = true;
                frmFile.Visible = false;
                frmOptions.Visible = false;
                frmTemplateInfo.Visible = false;
                try
{;
                txtCode(2).SetFocus;

                Case 4;
                txtCode(0).Visible = false;
                txtCode(1).Visible = false;
                txtCode(2).Visible = false;
                frmFile.Visible = true;
                frmOptions.Visible = false;
                frmTemplateInfo.Visible = false;
                try
{;
                txtCodeToFile.SetFocus;

                Case 5;
                txtCode(0).Visible = false;
                txtCode(1).Visible = false;
                txtCode(2).Visible = false;
                frmFile.Visible = false;
                frmOptions.Visible = true;
                frmTemplateInfo.Visible = false;
                try
{;
                chkFavorite.SetFocus;

                Case 6;
                txtCode(0).Visible = false;
                txtCode(1).Visible = false;
                txtCode(2).Visible = false;
                frmFile.Visible = false;
                frmOptions.Visible = false;
                frmTemplateInfo.Visible = true;
                try
{;
                cmdRecalc.SetFocus;

                if ( chkAutoRecalc.Value <> 0 )
            {;
                cmdRecalc_Click;
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

            public void tmrDoAction_Timer()
            {
                try
{;
                if ( ! OkayToDoAction )
            {
 return;

                tmrDoAction.Enabled = false;

                switch ActionToDo.ToUpper();
                Case "NEWTEMPLATE";
                NewTemplate true, ActionParam;

                //            'Case "DELTACHECK", "DELTA CHECK";
                //            '     If Len(Dir$(Parent.TemplateDatabasePath + "MDBPatch*.sad", vbNormal)) Then;
                //            '        If bUserSure("A Delta Patch file has been found. Would you like to apply it now ?") Then;
                //            '           SliceAndDice.ApplyPatch Dir$(Parent.TemplateDatabasePath + "MDBPatch*.sad", vbNormal);
                //            '        End If;
                //            '     End If;
                //            '     QueueAction "DeltaCheck", vbNullstring, 65535;
                //            '     OkayToDoAction = true;

                Case "DOINSERTION";
                DoInsertion null, ActionParam;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void txtCode_GotFocus            {
                CurrentCodeArea = Index;
            }

            public void txtCode_KeyDown            {
                Form_KeyDown KeyCode, Shift;
                if ( mbScramFormKey )
            {
 KeyCode == 0: Shift == 0;
            }

            public void Form_GotFocus()
            {
                try
{;
                lsbJumpTo.SetFocus                                ' More than likely the user is going to want to insert a pre-existing Template.;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Form_Initialize()
            {

                InitPublic;

                mnuExitAfterInsert.Checked = GetSetting(App.ProductName, "Settings", "Exit after insert", true);

                mnuShowPaintbrushIcon.Checked = GetSetting(App.ProductName, "Last", "Show Paitbrush Icon", true);
                mnuShowOnModuleRightClick.Checked = GetSetting(App.ProductName, "Last", "Show On Module Right Click", true);

                mnuSwitchTabsAutomatically.Checked = GetSetting(App.ProductName, "Last", "Switch tabs automatically", true);
                mnuPasswordProtection.Checked = GetSetting(App.ProductName, "Last", "Password Protection", false);
                mnuShowSplash.Checked = GetSetting(App.ProductName, "Last", "Show Splash", true);
                mnuOLEDragDrop.Checked = GetSetting(App.ProductName, "Last", "OLEDragDrop", false);

                txtCode(0).OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0):     txtCode(0).OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0);
                txtCode(1).OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0):     txtCode(1).OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0);
                txtCode(2).OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0):     txtCode(2).OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0);
                txtCodeToFile.OLEDragMode = IIf(mnuOLEDragDrop.Checked, 1, 0):  txtCodeToFile.OLEDropMode = IIf(mnuOLEDragDrop.Checked, 2, 0);
                ;
                chkAutoRecalc.Value = GetSetting(App.ProductName, "Last", "Auto Recalc", 0);

                lsbJumpTo.Arrange = GetSetting(App.ProductName, "Settings", "Bar Arrange", "1");
                lsbJumpTo.View = GetSetting(App.ProductName, "Settings", "Bar View", "1");

                if ( Len(Dir$(App.Path + gsBS + "SliceAndDice.mdb")) = 0 And Len(Dir$(App.Path + gsBS + "SliceAndDiceNew.mdb")) <> 0 )
            {;
                App.Path + gsBS + "SliceAndDice.mdb" Name App.Path + gsBS + "SliceAndDiceNew.mdb";
            }

            public As sChooseDatabase            {
                try
{;


                cdgSelect.Filter = "Access Database (*.mdb)|*.mdb|All Files (*.*)|*.*";
                cdgSelect.FilterIndex = 0;
                if ( Len(sPath) > 0 )
            {
 cdgSelect.InitDir == sPath;
                if ( Len(sFilename) > 0 )
            {
 cdgSelect.FileName == sFilename;
                cdgSelect.ShowOpen;
                if ( Err <> 0 )
            {;

                return; // ???;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As sChooseFile            {
                try
{;


                cdgSelect.Filter = IIf(Len(sFilter) And InStr(sFilter, "|"), sFilter, "All Files (*.*)|*.*");
                cdgSelect.FilterIndex = 0;
                if ( Len(sPath) > 0 )
            {
 cdgSelect.InitDir == sPath;
                if ( Len(sFilename) > 0 )
            {
 cdgSelect.FileName == sFilename;
                cdgSelect.ShowOpen;
                if ( Err <> 0 )
            {;

                return; // ???;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As sChooseColor            {
                try
{;



                cdgSelect.CancelError = true;
                if ( lTokenCount(sInitialColor, gsSC) = 3 )
            {;
                Red = sGetToken(sInitialColor, 1, gsSC);
                Green = sGetToken(sInitialColor, 1, gsSC);
                Blue = sGetToken(sInitialColor, 1, gsSC);
                if ( Red > 255 )
            {
 Red == 255;
                if ( Red < 0 )
            {
 Red == 0;
                if ( Green > 255 )
            {
 Green == 255;
                if ( Green < 0 )
            {
 Green == 0;
                if ( Blue > 255 )
            {
 Blue == 255;
                if ( Blue < 0 )
            {
 Blue == 0;
                cdgSelect.Color = RGB(Red, Green, Blue);
                }
            else
            {if ( Val(sInitialColor) > 0 )
            {;
                cdgSelect.Color = Val(sInitialColor);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Form_KeyDown            {
                // TODO: Rewrite try/catch and/or goto. EH_frmMain_Form_KeyDown;

                mbScramFormKey = false;

                if ( (Shift And vbShiftMask) > 0 )
            {
               ' Shift Key         *******************;
                switch KeyCode;
                ;
                Case vbKeyInsert: KeyCode = 0: Shift = 0: mbScramFormKey = true   ' Paste;
                if ( TypeOf ActiveControl Is TextBox )
            {;
                ActiveControl.SelText = Clipboard.GetText;
            }

            public void FindInCurrent            {
                // TODO: Rewrite try/catch and/or goto. EH_frmMain_FindInCurrent;
                Boolean static bInHereAlready;
                if ( bInHereAlready )
            {
 return;
                bInHereAlready = true;


                if ( CurrentTemplate Is null )
            {;
                MsgBox "Please select a " + gsTemplate + " before selecting to search.";
                return;
            }

            public void Form_Resize()
            {
                // TODO: Rewrite try/catch and/or goto. EH_Form_Resize;

                tabCode                                      ' Position the code entry areas.Height = ScaleHeight - tabCode                                      ' Position the code entry areas.Top;
                //        'lsbJumpTo.Height = ScaleHeight - 415;
                if ( ScaleWidth - tabCode                                      ' Position the code entry areas.Left < 0 )
            {
 return;       ' if ( there isn't enough display area to show the code entry areas, don't attempt to redraw it;
                tabCode                                      ' Position the code entry areas.Width = ScaleWidth - tabCode                                      ' Position the code entry areas.Left;
                txtName.Move lblCode(3).Left + lblCode(3).Width + 40, txtName.Top, tabCode                                      ' Position the code entry areas.Width - (lblCode(3).Left + lblCode(3).Width + 40 - tabCode                                      ' Position the code entry areas.Left), txtName.Height;
                txtShortName.Move lblCode(3).Left + lblCode(3).Width + 40, txtName.Top, tabCode                                      ' Position the code entry areas.Width - (lblCode(3).Left + lblCode(3).Width + 40 - tabCode                                      ' Position the code entry areas.Left), txtName.Height;

                txtCode(0).Move tabCode                                      ' Position the code entry areas.Left + 100, tabCode                                      ' Position the code entry areas.Top + 500, tabCode                                      ' Position the code entry areas.Width - 200, tabCode                                      ' Position the code entry areas.Height - 600;
                txtCode(1).Move tabCode                                      ' Position the code entry areas.Left + 100, tabCode                                      ' Position the code entry areas.Top + 500, tabCode                                      ' Position the code entry areas.Width - 200, tabCode                                      ' Position the code entry areas.Height - 600;
                txtCode(2).Move tabCode                                      ' Position the code entry areas.Left + 100, tabCode                                      ' Position the code entry areas.Top + 500, tabCode                                      ' Position the code entry areas.Width - 200, tabCode                                      ' Position the code entry areas.Height - 600;
                frmOptions.Move tabCode                                      ' Position the code entry areas.Left + 100, tabCode                                      ' Position the code entry areas.Top + 500, tabCode                                      ' Position the code entry areas.Width - 200, tabCode                                      ' Position the code entry areas.Height - 600;
                frmFile.Move tabCode                                      ' Position the code entry areas.Left + 100, tabCode                                      ' Position the code entry areas.Top + 500, tabCode                                      ' Position the code entry areas.Width - 200, tabCode                                      ' Position the code entry areas.Height - 600;
                txtFilename.Width = frmFile.Width - txtFilename.Left * 2;
                txtCodeToFile.Width = txtFilename.Width;
                txtCodeToFile.Height = frmFile.Height - txtCodeToFile.Top - 100;
                frmTemplateInfo.Move tabCode                                      ' Position the code entry areas.Left + 100, tabCode                                      ' Position the code entry areas.Top + 500, tabCode                                      ' Position the code entry areas.Width - 200, tabCode                                      ' Position the code entry areas.Height - 600;
                lstSoftVariables.Width = (frmTemplateInfo.Width - lstSoftVariables.Left * 3) \ 2;
                lstSoftCommands.Left = lstSoftVariables.Left * 2 + lstSoftVariables.Width;
                lstSoftCommands.Width = lstSoftVariables.Width;
                lblTemplateInfo(0).Left = lstSoftVariables.Left;
                lblTemplateInfo(1).Left = lstSoftCommands.Left;
                lstSoftVariables.Height = frmTemplateInfo.Height - lstSoftVariables.Top - 100;
                lstSoftCommands.Height = frmTemplateInfo.Height - lstSoftVariables.Top - 100;
            }

            public void mnuDBClassGen_Click()
            {
                tmrActivateDBClassGen.Enabled = true;
            }

            public As sPropertyType            {
                switch sFieldType;
                Case "Big Integer": sPropertyType = "Long";
                Case "Binary": sPropertyType = "Variant";
                Case "Boolean": sPropertyType = "Boolean";
                Case "Byte": sPropertyType = "Byte";
                Case "Char": sPropertyType = "string";
                Case "Currency": sPropertyType = "Currency";
                Case "Date / Time": sPropertyType = "Date";
                Case "Decimal": sPropertyType = "Variant";
                Case "Double": sPropertyType = "Double";
                Case "Float": sPropertyType = "Double";
                Case "Guid": sPropertyType = "string";
                Case "Integer": sPropertyType = "Integer";
                Case "Long": sPropertyType = "Long";
                Case "long Binary (OLE Object)": sPropertyType = "Variant";
                Case "Memo": sPropertyType = "Memo";
                Case "Numeric": sPropertyType = "Variant";
                Case "Single": sPropertyType = "Single";
                Case "Text": sPropertyType = "string";
                Case "Time": sPropertyType = "Date";
                Case "Time Stamp": sPropertyType = "Date";
                Case "VarBinary": sPropertyType = "Variant";
                Case }
            else
            {: sPropertyType = "Variant";
            }

            public void mnuExitAfterInsert_Click()
            {
                mnuExitAfterInsert.Checked = ! mnuExitAfterInsert.Checked;
            }

            public void mnuFileCopy_Click()
            {

                if ( CurrentTemplate Is null )
            {;
                MsgBox "Please select a " + gsTemplate + " to copy before selecting this option.";
                return;
            }

            public void mnuFileDelete_Click()
            {
                DeleteTemplate;
            }

            public void mnuFileExit_Click()
            {
                SaveTemplate;

                Hide;
                HideAllWindows;
                //    'VBIDEWindow.Visible = false      '   So hiding it will return control to VB;
            }

            public void mnuFileImport_Click()
            {

                if ( ! Parent.HostedByVB )
            {
 return;

                try
{;

                Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane.GetSelection lLine, lFirstCol, lLastLine, lLastCol    ' Determine where the cursor is;
                lLastLine +=  lLine + Abs(lLastCol > 1)    ' Determine what the last line selected is (discard last line if at beginning);
                sCode = Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane.CodeModule.Lines(lLine, lLastLine)   ' Grab the code selected from the active pane;
                GetProcAtLine lLine, sProcName, lProcType;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As GetCurrentTextSelection()
            {

                if ( ! Parent.HostedByVB )
            {
 return; // ???;

                try
{;

                Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane.GetSelection lLine, lFirstCol, lLastLine, lLastCol    ' Determine where the cursor is;
                lLastLine +=  lLine + Abs(lLastCol > 1)    ' Determine what the last line selected is (discard last line if at beginning);
                GetCurrentTextSelection = Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane.CodeModule.Lines(lLine, lLastLine)    ' Grab the code selected from the active pane;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void DeleteCurrentTextSelection()
            {

                if ( ! Parent.HostedByVB )
            {
 return;

                try
{;

                Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane.GetSelection lLine, lFirstCol, lLastLine, lLastCol;
                lLastLine +=  lLine + Abs(lLastCol > 1);
                Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane.CodeModule.DeleteLines lLine, lLastLine;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As DetermineLastLineInSelection()
            {

                if ( ! Parent.HostedByVB )
            {
 return; // ???;

                try
{;

                Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane.GetSelection lLine, lFirstCol, lLastLine, lLastCol;
                DetermineLastLineInSelection = lLastLine;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As DetermineFirstLineInSelection()
            {

                if ( ! Parent.HostedByVB )
            {
 return; // ???;

                try
{;

                Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane.GetSelection lLine, lFirstCol, lLastLine, lLastCol;
                DetermineFirstLineInSelection = lLine;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As DetermineFirstColumnInSelection()
            {

                if ( ! Parent.HostedByVB )
            {
 return; // ???;

                try
{;

                Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane.GetSelection lLine, lFirstCol, lLastLine, lLastCol;
                DetermineFirstColumnInSelection = lFirstCol;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As DetermineLastColumnInSelection()
            {

                if ( ! Parent.HostedByVB )
            {
 return; // ???;

                try
{;

                Parent.vbInst.ActiveCodePane                 ' Send all output to the active code pane.GetSelection lLine, lFirstCol, lLastLine, lLastCol;
                DetermineLastColumnInSelection = lLastCol;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuFileNew_Click()
            {
                NewTemplate;
            }

            public void mnuFileRefresh_Click()
            {
                try
{;

                sTitle = lsbJumpTo.BarKey + gsCategoryTemplateDelimiter + lsbJumpTo.BarItemName;
                RefillList;
                JumpTo sTitle, false, true;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuInsertTemplate_Click()
            {
                DoInsertion null, txtName;
            }

            public void Form_Terminate()
            {
                Form_Unload Cancel;
                //    ' LogEvent "frmMain: Terminate";
            }

            public void Form_Load()
            {
                m_asaHistory.Clear;
                LoadFormPosition Me;
                SetColors GetSetting$(App.ProductName, "Last", "Background Color", "&H80000018&"), GetSetting$(App.ProductName, "Last", "Foreground Color", "&H80000008&");
            }

            public void Form_Unload            {
                if ( ! mHotKeyOpenWindow Is null )
            {;
                mHotKeyOpenWindow.Clear;
                mHotKeyOpenWindow = null;
            }

            public void mHotKeyOpenWindow_HotKeyPress            {

                if ( sName = "Sandy Cancel Insertion" )
            {;
                gbCancelInsertion = true;
                }
            else
            {if ( sName = "Sandy Activate" )
            {;
                mHotKeyOpenWindow.RestoreAndActivate this.hwnd;
                }
            else
            {if ( sName = "Sandy Repeat Insertion" )
            {;
                if ( ! InternalCurrentTemplate Is null )
            {;
                sKey = InternalCurrentTemplate.Key;
            }

            public void txtCodeToFile_KeyDown            {
                Form_KeyDown KeyCode, Shift;
                if ( mbScramFormKey )
            {
 KeyCode == 0: Shift == 0;
            }

            public void txtFilename_KeyDown            {
                Form_KeyDown KeyCode, Shift;
                if ( mbScramFormKey )
            {
 KeyCode == 0: Shift == 0;
            }

            public void txtShortName_KeyDown            {
                Form_KeyDown KeyCode, Shift;
                if ( mbScramFormKey )
            {
 KeyCode == 0: Shift == 0;
            }

        }
    }
