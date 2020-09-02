using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class frmCommandHelp : System.Windows.Forms.Form
        {
         public System.Windows.Forms.MSComDlg.CommonDialog cdgHelp;
         public System.Windows.Forms.VB.TextBox txtSoftCommandName;
         public System.Windows.Forms.VB.TextBox txtAliases;
         public System.Windows.Forms.VB.TextBox txtSyntax;
         public System.Windows.Forms.VB.TextBox txtOneLineDescription;
         public System.Windows.Forms.VB.TextBox txtHelpFile;
         public System.Windows.Forms.VB.TextBox txtHelpTopic;
         public System.Windows.Forms.VB.TextBox txtLongDescription;
         public System.Windows.Forms.VB.TextBox txtComments;
         public System.Windows.Forms.VB.TextBox txtSeeAlso;
         public System.Windows.Forms.VB.TextBox txtExamples;
         public System.Windows.Forms.VB.CheckBox chkIsInline;
         public System.Windows.Forms.VB.PictureBox picTH;
         public System.Windows.Forms.VB.Label lblSoftCommandName;
         public System.Windows.Forms.VB.Label lblAliases;
         public System.Windows.Forms.VB.Label lblSyntax;
         public System.Windows.Forms.VB.Label lblOneLineDescription;
         public System.Windows.Forms.VB.Label lblHelpFile;
         public System.Windows.Forms.VB.Label lblHelpTopic;
         public System.Windows.Forms.VB.Label lblLongDescription;
         public System.Windows.Forms.VB.Label lblComments;
         public System.Windows.Forms.VB.Label lblSeeAlso;
         public System.Windows.Forms.VB.Label lblExamples;
         public System.Windows.Forms.VB.Label lblIsInline;
         public System.Windows.Forms.VB.Menu mnuFileExit;
         public System.Windows.Forms.VB.Menu mnuFirst;
         public System.Windows.Forms.VB.Menu mnuPrevious;
         public System.Windows.Forms.VB.Menu mnuNext;
         public System.Windows.Forms.VB.Menu mnuLast;
         public System.Windows.Forms.VB.Menu mnuFileFind;
         public System.Windows.Forms.VB.Menu mnuChangeCommandSets;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.Container components = null;

        public frmCommandHelp()
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
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCommandHelp));
            this.cdgHelp = new System.Windows.Forms.MSComDlg.CommonDialog();
            this.txtSoftCommandName = new System.Windows.Forms.VB.TextBox();
            this.txtAliases = new System.Windows.Forms.VB.TextBox();
            this.txtSyntax = new System.Windows.Forms.VB.TextBox();
            this.txtOneLineDescription = new System.Windows.Forms.VB.TextBox();
            this.txtHelpFile = new System.Windows.Forms.VB.TextBox();
            this.txtHelpTopic = new System.Windows.Forms.VB.TextBox();
            this.txtLongDescription = new System.Windows.Forms.VB.TextBox();
            this.txtComments = new System.Windows.Forms.VB.TextBox();
            this.txtSeeAlso = new System.Windows.Forms.VB.TextBox();
            this.txtExamples = new System.Windows.Forms.VB.TextBox();
            this.chkIsInline = new System.Windows.Forms.VB.CheckBox();
            this.picTH = new System.Windows.Forms.VB.PictureBox();
            this.lblSoftCommandName = new System.Windows.Forms.VB.Label();
            this.lblAliases = new System.Windows.Forms.VB.Label();
            this.lblSyntax = new System.Windows.Forms.VB.Label();
            this.lblOneLineDescription = new System.Windows.Forms.VB.Label();
            this.lblHelpFile = new System.Windows.Forms.VB.Label();
            this.lblHelpTopic = new System.Windows.Forms.VB.Label();
            this.lblLongDescription = new System.Windows.Forms.VB.Label();
            this.lblComments = new System.Windows.Forms.VB.Label();
            this.lblSeeAlso = new System.Windows.Forms.VB.Label();
            this.lblExamples = new System.Windows.Forms.VB.Label();
            this.lblIsInline = new System.Windows.Forms.VB.Label();
            this.mnuFileExit = new System.Windows.Forms.VB.Menu();
            this.mnuFirst = new System.Windows.Forms.VB.Menu();
            this.mnuPrevious = new System.Windows.Forms.VB.Menu();
            this.mnu} // = new System.Windows.Forms.VB.Menu();
            this.mnuLast = new System.Windows.Forms.VB.Menu();
            this.mnuFileFind = new System.Windows.Forms.VB.Menu();
            this.mnuChangeCommandSets = new System.Windows.Forms.VB.Menu();
            this.SuspendLayout();
            //
            // cdgHelp
            //
            this.cdgHelp.Name = "cdgHelp";
            this.cdgHelp.Location = new System.Drawing.Point(12, 252);
            //
            // txtSoftCommandName
            //
            this.txtSoftCommandName.Name = "txtSoftCommandName";
            this.txtSoftCommandName.BackColor = System.Drawing.Color.Transparent;
            this.txtSoftCommandName.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtSoftCommandName.CausesValidatio = 0;
            this.txtSoftCommandName.Size = new System.Drawing.Size(440, 16);
            this.txtSoftCommandName.Location = new System.Drawing.Point(70, 20);
            this.txtSoftCommandName.Locked = true;
            this.txtSoftCommandName.TabIndex = 10;
            //
            // txtAliases
            //
            this.txtAliases.Name = "txtAliases";
            this.txtAliases.BackColor = System.Drawing.Color.Transparent;
            this.txtAliases.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtAliases.CausesValidatio = 0;
            this.txtAliases.Size = new System.Drawing.Size(440, 20);
            this.txtAliases.Location = new System.Drawing.Point(70, 40);
            this.txtAliases.Locked = true;
            this.txtAliases.TabIndex = 9;
            //
            // txtSyntax
            //
            this.txtSyntax.Name = "txtSyntax";
            this.txtSyntax.BackColor = System.Drawing.Color.Transparent;
            this.txtSyntax.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtSyntax.CausesValidatio = 0;
            this.txtSyntax.Size = new System.Drawing.Size(440, 20);
            this.txtSyntax.Location = new System.Drawing.Point(70, 60);
            this.txtSyntax.Locked = true;
            this.txtSyntax.TabIndex = 8;
            //
            // txtOneLineDescription
            //
            this.txtOneLineDescription.Name = "txtOneLineDescription";
            this.txtOneLineDescription.BackColor = System.Drawing.Color.Transparent;
            this.txtOneLineDescription.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtOneLineDescription.CausesValidatio = 0;
            this.txtOneLineDescription.Size = new System.Drawing.Size(440, 20);
            this.txtOneLineDescription.Location = new System.Drawing.Point(70, 80);
            this.txtOneLineDescription.Locked = true;
            this.txtOneLineDescription.TabIndex = 7;
            //
            // txtHelpFile
            //
            this.txtHelpFile.Name = "txtHelpFile";
            this.txtHelpFile.BackColor = System.Drawing.Color.Transparent;
            this.txtHelpFile.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtHelpFile.CausesValidatio = 0;
            this.txtHelpFile.Size = new System.Drawing.Size(440, 20);
            this.txtHelpFile.Location = new System.Drawing.Point(70, 100);
            this.txtHelpFile.Locked = true;
            this.txtHelpFile.TabIndex = 6;
            //
            // txtHelpTopic
            //
            this.txtHelpTopic.Name = "txtHelpTopic";
            this.txtHelpTopic.BackColor = System.Drawing.Color.Transparent;
            this.txtHelpTopic.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtHelpTopic.CausesValidatio = 0;
            this.txtHelpTopic.Font = new System.Drawing.Font("MS Sans Serif",8F, ( System.Drawing.FontStyle.Underline ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.txtHelpTopic.Size = new System.Drawing.Size(440, 20);
            this.txtHelpTopic.Location = new System.Drawing.Point(70, 120);
            this.txtHelpTopic.Locked = true;
            this.txtHelpTopic.TabIndex = 5;
            //
            // txtLongDescription
            //
            this.txtLongDescription.Name = "txtLongDescription";
            this.txtLongDescription.BackColor = System.Drawing.Color.Transparent;
            this.txtLongDescription.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtLongDescription.CausesValidatio = 0;
            this.txtLongDescription.Size = new System.Drawing.Size(440, 80);
            this.txtLongDescription.Location = new System.Drawing.Point(70, 140);
            this.txtLongDescription.Locked = true;
//            this.txtLongDescription.MultiLine = -1;
            this.txtLongDescription.TabIndex = 4;
            //
            // txtComments
            //
            this.txtComments.Name = "txtComments";
            this.txtComments.BackColor = System.Drawing.Color.Transparent;
            this.txtComments.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtComments.CausesValidatio = 0;
            this.txtComments.Size = new System.Drawing.Size(440, 80);
            this.txtComments.Location = new System.Drawing.Point(70, 220);
            this.txtComments.Locked = true;
//            this.txtComments.MultiLine = -1;
            this.txtComments.TabIndex = 3;
            //
            // txtSeeAlso
            //
            this.txtSeeAlso.Name = "txtSeeAlso";
            this.txtSeeAlso.BackColor = System.Drawing.Color.Transparent;
            this.txtSeeAlso.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtSeeAlso.CausesValidatio = 0;
            this.txtSeeAlso.Size = new System.Drawing.Size(440, 80);
            this.txtSeeAlso.Location = new System.Drawing.Point(70, 300);
            this.txtSeeAlso.Locked = true;
//            this.txtSeeAlso.MultiLine = -1;
            this.txtSeeAlso.TabIndex = 2;
            //
            // txtExamples
            //
            this.txtExamples.Name = "txtExamples";
            this.txtExamples.BackColor = System.Drawing.Color.Transparent;
            this.txtExamples.BorderStyle = System.Windows.Forms.BorderStyle.None;
//            this.txtExamples.CausesValidatio = 0;
            this.txtExamples.Size = new System.Drawing.Size(440, 80);
            this.txtExamples.Location = new System.Drawing.Point(70, 380);
            this.txtExamples.Locked = true;
//            this.txtExamples.MultiLine = -1;
            this.txtExamples.TabIndex = 1;
            //
            // chkIsInline
            //
            this.chkIsInline.Name = "chkIsInline";
            this.chkIsInline.Text = "This is an Inline command if checked ( ie. %%Command::Parameters%% )";
//            this.chkIsInline.CausesValidatio = 0;
            this.chkIsInline.Enabled = false;
            this.chkIsInline.Font = new System.Drawing.Font("MS Sans Serif",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.chkIsInline.ForeColor = System.Drawing.Color.FromArgb(-2147483640);
            this.chkIsInline.Size = new System.Drawing.Size(458, 20);
            this.chkIsInline.Location = new System.Drawing.Point(70, 0);
            this.chkIsInline.TabIndex = 0;
            //
            // picTH
            //
            this.picTH.Name = "picTH";
            this.picTH.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.picTH.ForeColor = System.Drawing.Color.FromArgb(-2147483640);
            this.picTH.Size = new System.Drawing.Size(51, 43);
            this.picTH.Location = new System.Drawing.Point(4, 164);
            this.picTH.TabIndex = 22;
            //
            // lblSoftCommandName
            //
            this.lblSoftCommandName.Name = "lblSoftCommandName";
            this.lblSoftCommandName.Text = "Name";
            this.lblSoftCommandName.Font = new System.Drawing.Font("Times New Roman",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblSoftCommandName.Size = new System.Drawing.Size(30, 13);
            this.lblSoftCommandName.Location = new System.Drawing.Point(2, 22);
            this.lblSoftCommandName.TabIndex = 21;
            //
            // lblAliases
            //
            this.lblAliases.Name = "lblAliases";
            this.lblAliases.Text = "Aliases";
            this.lblAliases.Font = new System.Drawing.Font("Times New Roman",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblAliases.Size = new System.Drawing.Size(38, 13);
            this.lblAliases.Location = new System.Drawing.Point(2, 42);
            this.lblAliases.TabIndex = 20;
            //
            // lblSyntax
            //
            this.lblSyntax.Name = "lblSyntax";
            this.lblSyntax.Text = "Syntax";
            this.lblSyntax.Font = new System.Drawing.Font("Times New Roman",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblSyntax.Size = new System.Drawing.Size(36, 13);
            this.lblSyntax.Location = new System.Drawing.Point(2, 62);
            this.lblSyntax.TabIndex = 19;
            //
            // lblOneLineDescription
            //
            this.lblOneLineDescription.Name = "lblOneLineDescription";
            this.lblOneLineDescription.Text = "Summary";
            this.lblOneLineDescription.Font = new System.Drawing.Font("Times New Roman",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblOneLineDescription.Size = new System.Drawing.Size(51, 13);
            this.lblOneLineDescription.Location = new System.Drawing.Point(2, 82);
            this.lblOneLineDescription.TabIndex = 18;
            //
            // lblHelpFile
            //
            this.lblHelpFile.Name = "lblHelpFile";
            this.lblHelpFile.Text = "Help File";
            this.lblHelpFile.Font = new System.Drawing.Font("Times New Roman",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblHelpFile.Size = new System.Drawing.Size(48, 13);
            this.lblHelpFile.Location = new System.Drawing.Point(2, 102);
            this.lblHelpFile.TabIndex = 17;
            //
            // lblHelpTopic
            //
            this.lblHelpTopic.Name = "lblHelpTopic";
            this.lblHelpTopic.Text = "Help Topic";
            this.lblHelpTopic.Font = new System.Drawing.Font("Times New Roman",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblHelpTopic.Size = new System.Drawing.Size(55, 13);
            this.lblHelpTopic.Location = new System.Drawing.Point(2, 122);
            this.lblHelpTopic.TabIndex = 16;
            //
            // lblLongDescription
            //
            this.lblLongDescription.Name = "lblLongDescription";
            this.lblLongDescription.Text = "Description";
            this.lblLongDescription.Font = new System.Drawing.Font("Times New Roman",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblLongDescription.Size = new System.Drawing.Size(60, 13);
            this.lblLongDescription.Location = new System.Drawing.Point(2, 142);
            this.lblLongDescription.TabIndex = 15;
            //
            // lblComments
            //
            this.lblComments.Name = "lblComments";
            this.lblComments.Text = "Comments";
            this.lblComments.Font = new System.Drawing.Font("Times New Roman",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblComments.Size = new System.Drawing.Size(57, 13);
            this.lblComments.Location = new System.Drawing.Point(2, 222);
            this.lblComments.TabIndex = 14;
            //
            // lblSeeAlso
            //
            this.lblSeeAlso.Name = "lblSeeAlso";
            this.lblSeeAlso.Text = "See Also";
            this.lblSeeAlso.Font = new System.Drawing.Font("Times New Roman",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblSeeAlso.Size = new System.Drawing.Size(45, 13);
            this.lblSeeAlso.Location = new System.Drawing.Point(2, 302);
            this.lblSeeAlso.TabIndex = 13;
            //
            // lblExamples
            //
            this.lblExamples.Name = "lblExamples";
            this.lblExamples.Text = "Examples";
            this.lblExamples.Font = new System.Drawing.Font("Times New Roman",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblExamples.Size = new System.Drawing.Size(49, 13);
            this.lblExamples.Location = new System.Drawing.Point(2, 382);
            this.lblExamples.TabIndex = 12;
            //
            // lblIsInline
            //
            this.lblIsInline.Name = "lblIsInline";
            this.lblIsInline.Text = "Is Inline";
            this.lblIsInline.Font = new System.Drawing.Font("Times New Roman",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblIsInline.Size = new System.Drawing.Size(44, 13);
            this.lblIsInline.Location = new System.Drawing.Point(2, 4);
            this.lblIsInline.TabIndex = 11;
            //
            // mnuFileExit
            //
            this.mnuFileExit.Name = "mnuFileExit";
            this.mnuFileExit.Text = "&X";
            this.mnuFileExit.Enabled = false;
            this.mnuFileExit.Visible = false;
            //
            // mnuFirst
            //
            this.mnuFirst.Name = "mnuFirst";
            this.mnuFirst.Text = "Fi&rst";
            //
            // mnuPrevious
            //
            this.mnuPrevious.Name = "mnuPrevious";
            this.mnuPrevious.Text = "&<<";
            //
            // mnuNext
            //
            this.mnuNext.Name = "mnuNext";
            this.mnuNext.Text = "&>>";
            //
            // mnuLast
            //
            this.mnuLast.Name = "mnuLast";
            this.mnuLast.Text = "&Last";
            //
            // mnuFileFind
            //
            this.mnuFileFind.Name = "mnuFileFind";
            this.mnuFileFind.Text = "&Find";
            //
            // mnuChangeCommandSets
            //
            this.mnuChangeCommandSets.Name = "mnuChangeCommandSets";
            this.mnuChangeCommandSets.Text = "&Command Set";
            this.mnuChangeCommandSets.Enabled = false;
            this.mnuChangeCommandSets.Visible = false;
            //
            // frmCommandHelp
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.cdgHelp,
                  this.txtSoftCommandName,
                  this.txtAliases,
                  this.txtSyntax,
                  this.txtOneLineDescription,
                  this.txtHelpFile,
                  this.txtHelpTopic,
                  this.txtLongDescription,
                  this.txtComments,
                  this.txtSeeAlso,
                  this.txtExamples,
                  this.chkIsInline,
                  this.picTH,
                  this.lblSoftCommandName,
                  this.lblAliases,
                  this.lblSyntax,
                  this.lblOneLineDescription,
                  this.lblHelpFile,
                  this.lblHelpTopic,
                  this.lblLongDescription,
                  this.lblComments,
                  this.lblSeeAlso,
                  this.lblExamples,
                  this.lblIsInline,
                  this.mnuFileExit,
                  this.mnuFirst,
                  this.mnuPrevious,
                  this.mnuNext,
                  this.mnuLast,
                  this.mnuFileFind,
                  this.mnuChangeCommandSets
            });
            this.Name = "frmCommandHelp";
            this.ResumeLayout(false);
        }
        #endregion

        public CSadCommands SadCommandSet;
        public CSadCommand CurrCommand;
        public Variant vCurrCommandKey;
        public int NextTop;
        public CSadCommand CurrMember;
        public string sChoices;
        public string sChoice;
        public CAssocArray asaOrdered;


                public object CurrCommandKey
    {
        set
        {
        try
{

        if ( ! SadCommandSet.Item(vKey) Is null )
            {

         CurrCommand = SadCommandSet.Item(value);
        vCurrCommandKey = CurrCommand.Index;
        Populate;
        ;
        }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        }

        }

    }



            public void Populate()
            {
                try
{;

                if ( SadCommandIs null )
            {
 return;
                if ( CurrCommand Is null )
            {
 return;


                Caption = "SAD Soft Command Reference - " + SadCommandSet.Attributes("Name") + " ( " + SadCommandSet.Count + " commands)";
                chkIsInline.Enabled = true;
                chkIsInline.Value = Abs(.IsInline);
                if ( CurrCommand.IsInline )
            {;
                chkIsInline.Text = "This is an INLINE Soft Command (ie. " + gsSoftVarDelimiter + CurrCommand.SoftCommandName + gsInlineCmdDelimiter + "Parameters" + gsSoftVarDelimiter + gsS + gsPC;
                }
            else
            {;
                chkIsInline.Text = "This is a REGULAR Soft Command (ie. " + gsSoftCmdDelimiter + CurrCommand.SoftCommandName + " )";
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void chkIsInline_Click()
            {
                try
{;
                chkIsInline.Value = Abs(CurrCommand.IsInline);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Form_Initialize()
            {

            }

            public void Form_KeyDown            {
                try
{;
                if ( (Shift And vbCtrlMask) > 0 )
            {;
                switch KeyCode;
                Case vbKeyPageUp: mnuPrevious_Click;
                Case vbKeyPageDown: mnuNext_Click;
                Case vbKeyF: mnuFileFind_Click;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Form_Load()
            {
                try
{;
                LoadFormPosition Me, , false;
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
            }

            public void Form_Terminate()
            {

            }

            public void Form_Unload            {
                try
{;
                SaveFormPosition Me;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuFileExit_Click()
            {
                try
{;
                SaveFormPosition Me;
                CurrCommand = null;
                SadCommand= null;
                Hide;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuFileFind_Click()
            {
                try
{;

                sChoices = string.Empty;
                foreach( var CurrMember in SadCommandSet );
                sChoices +=  CurrMember.SoftCommandName + IIf(CurrMember.IsInline, " (Inline)" + gsE + CurrMember.OneLineDescription + gsSC, gsE + CurrMember.OneLineDescription + gsSC);
                if ( Len(CurrMember.Aliases) )
            {;
                sChoices +=  Replace(CurrMember.Aliases, ", ", IIf(CurrMember.IsInline, " (Inline)" + "=See " + CurrMember.SoftCommandName + gsSC, "=See " + CurrMember.SoftCommandName + gsSC));
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuFirst_Click()
            {
                try
{;
                CurrCommandKey = 1;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuLast_Click()
            {
                try
{;
                CurrCommandKey = SadCommandSet.Count;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuNext_Click()
            {
                try
{;
                if ( vCurrCommandKey + 1 < SadCommandSet.Count )
            {;
                CurrCommandKey = vCurrCommandKey + 1;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void mnuPrevious_Click()
            {
                try
{;
                if ( vCurrCommandKey - 1 > 0 )
            {;
                CurrCommandKey = vCurrCommandKey - 1;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void txtHelpFile_Click()
            {
                txtHelpTopic_Click;
            }

            public void txtHelpFile_DblClick()
            {
                txtHelpTopic_Click;
            }

            public void txtHelpTopic_Click()
            {
                if ( Len(txtHelpFile) > 0 And Len(txtHelpTopic) = 0 )
            {;

                cdgHelp.HelpFile = txtHelpFile;
                //            ' Go to the Click Event topic in the Help file.;
                //            ' The number is determined in the [MAP] section;
                //            ' of the .HPJ file for the .chm file. You can;
                //            ' edit this number only if you are using the;
                //            ' Microsoft Help Workshop to build your;
                //            ' own Help file.;
                cdgHelp.HelpContext = txtHelpTopic;
                cdgHelp.HelpCommand = cdlHelpContext;
                cdgHelp.ShowHelp;
            }

            public void txtHelpTopic_DblClick()
            {
                txtHelpTopic_Click;
            }

        }
    }
