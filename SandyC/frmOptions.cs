using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class frmOptions : System.Windows.Forms.Form
        {
         public System.Windows.Forms.VB.PictureBox picBottom;
         public System.Windows.Forms.VB.CheckBox chkProcessIncludes;
         public System.Windows.Forms.VB.CommandButton cmdOK;
         public System.Windows.Forms.VB.CommandButton cmdCancel;
         public System.Windows.Forms.VB.CommandButton cmdGenerate;
         public System.Windows.Forms.VB.CommandButton cmdPickFile;
         public System.Windows.Forms.VB.PictureBox picRight;
         public System.Windows.Forms.MSComctlLib.ListView lvwContents;
         public System.Windows.Forms.VB.PictureBox picLeft;
         public System.Windows.Forms.MSComctlLib.TreeView tvwHierarchy;
         public System.Windows.Forms.MSComctlLib.ImageList imlIcons;
         public System.Windows.Forms.VB.PictureBox picOptions3;
         public System.Windows.Forms.MSComctlLib.ListView lvwInfo3;
         public System.Windows.Forms.VB.PictureBox picOptions2;
         public System.Windows.Forms.MSComctlLib.ListView lvwInfo2;
         public System.Windows.Forms.VB.PictureBox picOptions1;
         public System.Windows.Forms.MSComctlLib.ListView lvwInfo1;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.Container components = null;

        public frmOptions()
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
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmOptions));
            this.picBottom = new System.Windows.Forms.VB.PictureBox();
            this.chkProcessIncludes = new System.Windows.Forms.VB.CheckBox();
            this.cmdOK = new System.Windows.Forms.VB.CommandButton();
            this.cmdCancel = new System.Windows.Forms.VB.CommandButton();
            this.cmdGenerate = new System.Windows.Forms.VB.CommandButton();
            this.cmdPickFile = new System.Windows.Forms.VB.CommandButton();
            this.picRight = new System.Windows.Forms.VB.PictureBox();
            this.lvwContents = new System.Windows.Forms.MSComctlLib.ListView();
            this.picLeft = new System.Windows.Forms.VB.PictureBox();
            this.tvwHierarchy = new System.Windows.Forms.MSComctlLib.TreeView();
            this.imlIcons = new System.Windows.Forms.MSComctlLib.ImageList();
            this.picOptions3 = new System.Windows.Forms.VB.PictureBox();
            this.lvwInfo3 = new System.Windows.Forms.MSComctlLib.ListView();
            this.picOptions2 = new System.Windows.Forms.VB.PictureBox();
            this.lvwInfo2 = new System.Windows.Forms.MSComctlLib.ListView();
            this.picOptions1 = new System.Windows.Forms.VB.PictureBox();
            this.lvwInfo1 = new System.Windows.Forms.MSComctlLib.ListView();
            this.SuspendLayout();
            this.picBottom.SuspendLayout();
            this.picRight.SuspendLayout();
            this.picLeft.SuspendLayout();
            this.picOptions3.SuspendLayout();
            this.picOptions2.SuspendLayout();
            this.picOptions1.SuspendLayout();
            //
            // picBottom
            //
            this.picBottom.Name = "picBottom";
//            this.picBottom.Align = 2;
            this.picBottom.BackColor = System.Drawing.Color.FromArgb(-2147483648);
            this.picBottom.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.picBottom.ForeColor = System.Drawing.Color.FromArgb(-2147483640);
            this.picBottom.Size = new System.Drawing.Size(574, 33);
            this.picBottom.Location = new System.Drawing.Point(0, 390);
            this.picBottom.TabIndex = 6;
            this.picBottom.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.chkProcessIncludes,
                        this.cmdOK,
                        this.cmdCancel,
                        this.cmdGenerate,
                        this.cmdPickFile
            });
            //
            // chkProcessIncludes
            //
            this.chkProcessIncludes.Name = "chkProcessIncludes";
            this.chkProcessIncludes.Text = "Process #include <x.h> files";
            this.chkProcessIncludes.Size = new System.Drawing.Size(171, 35);
            this.chkProcessIncludes.Location = new System.Drawing.Point(332, 0);
            this.chkProcessIncludes.TabIndex = 15;
            //
            // cmdOK
            //
            this.cmdOK.Name = "cmdOK";
            this.cmdOK.Text = "OK";
            this.cmdOK.Size = new System.Drawing.Size(73, 25);
            this.cmdOK.Location = new System.Drawing.Point(9, 4);
            this.cmdOK.TabIndex = 10;
            //
            // cmdCancel
            //
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.Size = new System.Drawing.Size(73, 25);
            this.cmdCancel.Location = new System.Drawing.Point(90, 4);
            this.cmdCancel.TabIndex = 9;
            //
            // cmdGenerate
            //
            this.cmdGenerate.Name = "cmdGenerate";
            this.cmdGenerate.Text = "Generate";
            this.cmdGenerate.Size = new System.Drawing.Size(73, 25);
            this.cmdGenerate.Location = new System.Drawing.Point(252, 4);
            this.cmdGenerate.TabIndex = 8;
            //
            // cmdPickFile
            //
            this.cmdPickFile.Name = "cmdPickFile";
            this.cmdPickFile.Text = "&Browse";
            this.cmdPickFile.Size = new System.Drawing.Size(73, 25);
            this.cmdPickFile.Location = new System.Drawing.Point(171, 4);
            this.cmdPickFile.TabIndex = 7;
            //
            // picRight
            //
            this.picRight.Name = "picRight";
//            this.picRight.Align = 4;
            this.picRight.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.picRight.Size = new System.Drawing.Size(388, 390);
            this.picRight.Location = new System.Drawing.Point(186, 0);
            this.picRight.TabIndex = 13;
            this.picRight.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.lvwContents
            });
            //
            // lvwContents
            //
            this.lvwContents.Name = "lvwContents";
            this.lvwContents.Size = new System.Drawing.Size(385, 376);
            this.lvwContents.Location = new System.Drawing.Point(0, 1);
            this.lvwContents.TabIndex = 14;
            this.lvwContents.View = System.Windows.Forms.View.List;
//            this.lvwContents.Arrange = 1;
            this.lvwContents.LabelEdit = true;
            this.lvwContents.LabelWrap = true;
            this.lvwContents.HideSelection = false;
//            this.lvwContents.FullRowSelect = -1;
//            this.lvwContents.GridLines = -1;
//            this.lvwContents.HotTracking = -1;
//            this.lvwContents.Icons = "imlIcons";
//            this.lvwContents.SmallIcons = "imlIcons";
//            this.lvwContents.ColHdrIcons = "imlIcons";
            this.lvwContents.ForeColor = System.Drawing.Color.FromArgb(-2147483640);
            this.lvwContents.BackColor = System.Drawing.Color.FromArgb(12648447);
            this.lvwContents.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
//            this.lvwContents.NumItems = 2;
//            this.lvwContents.ColumnHeader(1) = ;
//            this.lvwContents.ColumnHeader(2) = ;
            //
            // picLeft
            //
            this.picLeft.Name = "picLeft";
//            this.picLeft.Align = 3;
            this.picLeft.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.picLeft.Size = new System.Drawing.Size(176, 390);
            this.picLeft.Location = new System.Drawing.Point(0, 0);
            this.picLeft.TabIndex = 11;
            this.picLeft.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.tvwHierarchy,
                        this.imlIcons
            });
            //
            // tvwHierarchy
            //
            this.tvwHierarchy.Name = "tvwHierarchy";
            this.tvwHierarchy.Size = new System.Drawing.Size(172, 354);
            this.tvwHierarchy.Location = new System.Drawing.Point(1, 1);
            this.tvwHierarchy.TabIndex = 12;
            this.tvwHierarchy.HideSelection = false;
//            this.tvwHierarchy.Indentation = 265;
            this.tvwHierarchy.LabelEdit = true;
//            this.tvwHierarchy.FullRowSelect = -1;
//            this.tvwHierarchy.ImageList = "imlIcons";
            this.tvwHierarchy.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tvwHierarchy.Font = new System.Drawing.Font("MS Sans Serif",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            //
            // imlIcons
            //
            this.imlIcons.Name = "imlIcons";
            this.imlIcons.Location = new System.Drawing.Point(3, 346);
            this.imlIcons.BackColor = System.Drawing.Color.FromArgb(-2147483648);
//            this.imlIcons.ImageWidth = 16;
//            this.imlIcons.ImageHeight = 16;
//            this.imlIcons.MaskColor = 16777215;
//            this.imlIcons.ListImage1 = ;
//            this.imlIcons.ListImage2 = ;
//            this.imlIcons.ListImage3 = ;
//            this.imlIcons.ListImage4 = ;
//            this.imlIcons.ListImage5 = ;
//            this.imlIcons.ListImage6 = ;
//            this.imlIcons.ListImage7 = ;
//            this.imlIcons.ListImage8 = ;
//            this.imlIcons.ListImage9 = ;
//            this.imlIcons.ListImage10 = ;
//            this.imlIcons.ListImage11 = ;
//            this.imlIcons.ListImage12 = ;
//            this.imlIcons.ListImage13 = ;
//            this.imlIcons.ListImage14 = ;
//            this.imlIcons.ListImage15 = ;
//            this.imlIcons.ListImage15 = ;
            //
            // picOptions3
            //
            this.picOptions3.Name = "picOptions3";
            this.picOptions3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.picOptions3.Size = new System.Drawing.Size(379, 252);
            this.picOptions3.Location = new System.Drawing.Point(3666, 32);
            this.picOptions3.TabIndex = 2;
            this.picOptions3.TabStop = false;
            //
            // lvwInfo3
            //
            this.lvwInfo3.Name = "lvwInfo3";
            this.lvwInfo3.Size = new System.Drawing.Size(376, 249);
            this.lvwInfo3.Location = new System.Drawing.Point(0, 0);
            this.lvwInfo3.TabIndex = 5;
            this.lvwInfo3.View = System.Windows.Forms.View.List;
//            this.lvwInfo3.Arrange = 1;
            this.lvwInfo3.LabelEdit = true;
            this.lvwInfo3.LabelWrap = true;
            this.lvwInfo3.HideSelection = false;
//            this.lvwInfo3.Checkboxes = -1;
//            this.lvwInfo3.FullRowSelect = -1;
//            this.lvwInfo3.GridLines = -1;
//            this.lvwInfo3.HotTracking = -1;
//            this.lvwInfo3.Icons = "imlIcons";
//            this.lvwInfo3.SmallIcons = "imlIcons";
//            this.lvwInfo3.ColHdrIcons = "imlIcons";
            this.lvwInfo3.ForeColor = System.Drawing.Color.FromArgb(-2147483640);
            this.lvwInfo3.BackColor = System.Drawing.Color.FromArgb(-2147483643);
            this.lvwInfo3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
//            this.lvwInfo3.NumItems = 0;
            //
            // picOptions2
            //
            this.picOptions2.Name = "picOptions2";
            this.picOptions2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.picOptions2.Size = new System.Drawing.Size(379, 252);
            this.picOptions2.Location = new System.Drawing.Point(3666, 32);
            this.picOptions2.TabIndex = 1;
            this.picOptions2.TabStop = false;
            //
            // lvwInfo2
            //
            this.lvwInfo2.Name = "lvwInfo2";
            this.lvwInfo2.Size = new System.Drawing.Size(376, 249);
            this.lvwInfo2.Location = new System.Drawing.Point(0, 0);
            this.lvwInfo2.TabIndex = 4;
            this.lvwInfo2.View = System.Windows.Forms.View.List;
//            this.lvwInfo2.Arrange = 1;
            this.lvwInfo2.LabelEdit = true;
            this.lvwInfo2.LabelWrap = true;
            this.lvwInfo2.HideSelection = false;
//            this.lvwInfo2.Checkboxes = -1;
//            this.lvwInfo2.FullRowSelect = -1;
//            this.lvwInfo2.GridLines = -1;
//            this.lvwInfo2.HotTracking = -1;
//            this.lvwInfo2.Icons = "imlIcons";
//            this.lvwInfo2.SmallIcons = "imlIcons";
//            this.lvwInfo2.ColHdrIcons = "imlIcons";
            this.lvwInfo2.ForeColor = System.Drawing.Color.FromArgb(-2147483640);
            this.lvwInfo2.BackColor = System.Drawing.Color.FromArgb(-2147483643);
            this.lvwInfo2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
//            this.lvwInfo2.NumItems = 0;
            //
            // picOptions1
            //
            this.picOptions1.Name = "picOptions1";
            this.picOptions1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.picOptions1.Size = new System.Drawing.Size(379, 252);
            this.picOptions1.Location = new System.Drawing.Point(3666, 32);
            this.picOptions1.TabIndex = 0;
            this.picOptions1.TabStop = false;
            //
            // lvwInfo1
            //
            this.lvwInfo1.Name = "lvwInfo1";
            this.lvwInfo1.Size = new System.Drawing.Size(376, 249);
            this.lvwInfo1.Location = new System.Drawing.Point(0, 0);
            this.lvwInfo1.TabIndex = 3;
            this.lvwInfo1.View = System.Windows.Forms.View.List;
//            this.lvwInfo1.Arrange = 1;
            this.lvwInfo1.LabelEdit = true;
            this.lvwInfo1.LabelWrap = true;
            this.lvwInfo1.HideSelection = false;
//            this.lvwInfo1.Checkboxes = -1;
//            this.lvwInfo1.FullRowSelect = -1;
//            this.lvwInfo1.GridLines = -1;
//            this.lvwInfo1.HotTracking = -1;
//            this.lvwInfo1.Icons = "imlIcons";
//            this.lvwInfo1.SmallIcons = "imlIcons";
//            this.lvwInfo1.ColHdrIcons = "imlIcons";
            this.lvwInfo1.ForeColor = System.Drawing.Color.FromArgb(-2147483640);
            this.lvwInfo1.BackColor = System.Drawing.Color.FromArgb(-2147483643);
            this.lvwInfo1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
//            this.lvwInfo1.NumItems = 0;
            //
            // frmOptions
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.picBottom,
                  this.picRight,
                  this.picLeft,
                  this.picOptions3,
                  this.picOptions2,
                  this.picOptions1
            });
            this.Name = "frmOptions";
            this.picBottom.ResumeLayout(false);
            this.picRight.ResumeLayout(false);
            this.picLeft.ResumeLayout(false);
            this.picOptions3.ResumeLayout(false);
            this.picOptions2.ResumeLayout(false);
            this.picOptions1.ResumeLayout(false);
            this.ResumeLayout(false);
        }
        #endregion

        public NewCommands Parent;
        public New asaDefines;
        public New asaTemp;
        public string sContents;
        public string sLine;
        public int LineNumber;
        public string sRootName;
        public bool bInAComment;
        public string sEOL;
        public string sCurrNode;
        public string sObject;
        public string sScope;
        public string sLastMethod;
        public string sEnum;
        public string sStruct;
        public string sUnion;
        public int lBraceCount;
        public object CurrItem;
        public object sOut;
        public string sFilename;
        public SliceAndDice.CAssocItem CurrItem;
        public Node CurrNode;


            public void AddNode            {
                // TODO: Rewrite try/catch and/or goto. 0;
                if ( Len(sText) == 0 Or Len(sIcon) == 0 )
            {
 return;
                if ( Len(sKey) == 0 )
            {
 sKey == sText;

                if ( Len(sParent) )
            {;

                tvwHierarchy.Nodes.Add(                                                                                                                                                                                                      sParent, tvwChild, sKey, sText, sIcon, sIcon).ExpandedImage = sIcon;
                tvwHierarchy.Nodes.Add(                                                                                                                                                                                                      sParent, tvwChild, sKey, sText, sIcon, sIcon).Expanded = bExpanded;
                tvwHierarchy.Nodes.Add(                                                                                                                                                                                                      sParent, tvwChild, sKey, sText, sIcon, sIcon).BackColor = lvwContents.BackColor;
                tvwHierarchy.Nodes.Add(                                                                                                                                                                                                      sParent, tvwChild, sKey, sText, sIcon, sIcon).ForeColor = lvwContents.ForeColor;
                tvwHierarchy.Nodes.Add(                                                                                                                                                                                                      sParent, tvwChild, sKey, sText, sIcon, sIcon).Tag = sTag;
            }

            public void ProcessFile            {
                ;
                ;

                Form_Load;

                sContents = Parent.Parent.sFileContents(sFilename);
                Screen.MousePointer = vbHourglass;
                if ( Len(sContents) )
            {;
                sEOL = Chr(13);
                sRootName = Parent.Parent.sGetToken(sFilename, Parent.Parent.lTokenCount(sFilename, "\"), "\");
                if ( bClearFirst )
            {;
                lvwContents.ListItems.Clear;
                tvwHierarchy.Nodes.Clear;
            }

            public As MassReplace            {

                sOut = sLine;
                foreach( var CurrItem in asaDefines );
                if ( Len(CurrItem.Key) > 0 )
            {;
                sOut = Replace(sOut, CurrItem.Key, CurrItem.Value);
            }

            public void chkProcessIncludes_Click()
            {
                SaveSetting App.ProductName, "Last", "Process Includes", chkProcessIncludes.Value;
            }

            public void cmdCancel_Click()
            {
                Form_Unload 0;
                Hide;
            }

            public void cmdGenerate_Click()
            {
                Form_Unload 0;
                MsgBox "Generation would occur here";
                Hide;
            }

            public void cmdOK_Click()
            {
                Form_Unload 0;
                Hide;
            }

            public void cmdPickFile_Click()
            {
                sFilename = Parent.Parent.sChooseFile(, , "C Header|*.h|C++ Header|*.hpp|All Files|*.*");
                if ( Len(sFilename) )
            {
 ProcessFile sFilename, true;
            }

            public void Form_Load()
            {
                LoadFormPosition Me;
                Form_Resize;

                asaDefines.Clear;
                asaDefines.Item("&") = "!!!";
                asaDefines.Item("virtual ") = "!!!";
                asaDefines.Item("char *") = "string ";
                asaDefines.Item("LPCTSTR ") = "string ";
                asaDefines.Item("BSTR ") = "string ";
                asaDefines.Item("void *") = "long ";
                asaDefines.Item("int ") = "long ";
                asaDefines.Item("short") = "Integer";
                //       '.Item("long ") = "long ";
                asaDefines.Item("char ") = "Byte ";
                asaDefines.Item("") = " ";
                asaDefines.Item("") = " ";
                asaDefines.Item("") = " ";
                asaDefines.Item("operator==") = "CompareForEquality ";
                asaDefines.Item("operator!=") = "CompareForInequality ";
                asaDefines.Item("_exports") = "!!!";
                asaDefines.Item(": public") = " - Implements ";
                asaDefines.Item("const ") = "!!!";
                asaDefines.Item("afx_msg ") = "!!!";
                asaDefines.Item(" const") = "!!!";
                asaDefines.Item("(void)") = "()";
                asaDefines.Item("void ") = "Sub ";
                asaDefines.Item(" void") = "!!!";
                asaDefines.Item("(string ") = "(ByVal string ";
                asaDefines.Item(", string ") = ", ByVal string ";
            }

            public void Form_Resize()
            {
                try
{;
                picLeft.Width = ScaleWidth * 0.55;
                picRight.Width = ScaleWidth - picLeft.Width - 100;
                tvwHierarchy.Move 30, 30, picLeft.Width - 30, picLeft.Height - 30;
                lvwContents.Move 30, 30, picRight.Width - 30, picRight.Height - 30;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Form_Unload            {
                SaveFormPosition Me;
            }

            public void tvwHierarchy_MouseUp            {
                ;
                if ( Button = 1 And Shift = 0 )
            {;
                CurrNode = null;
                CurrNode = tvwHierarchy.HitTest(x, y);
                if ( ! CurrNode Is null )
            {;
                if ( lvwContents.Tag <> CurrNode.Tag) .ToUpper()
            {;
                lvwContents.ListItems.Clear;
                switch CurrNode.Tag.ToUpper();
                Case "ASADEFINES";
                foreach( var CurrItem in asaDefines );

                lvwContents.ListItems.Add(                                                                                                                                                                                                      , , CurrItem.Key, "Constant", "Constant").SubItems(1) = CurrItem.Value;
            }

        }
    }
