using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class frmRegisterCodeSnippet : System.Windows.Forms.Form
        {
         public System.Windows.Forms.VB.TextBox txtObjectName;
         public System.Windows.Forms.VB.CommandButton cmdRemoveAddin;
         public System.Windows.Forms.VB.TextBox txtLocation;
         public System.Windows.Forms.VB.CommandButton cmdAddAddin;
         public System.Windows.Forms.VB.Label Label11;
         public System.Windows.Forms.VB.Label Label10;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.Container components = null;

        public frmRegisterCodeSnippet()
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
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRegisterCodeSnippet));
            this.txtObjectName = new System.Windows.Forms.VB.TextBox();
            this.cmdRemoveAddin = new System.Windows.Forms.VB.CommandButton();
            this.txtLocation = new System.Windows.Forms.VB.TextBox();
            this.cmdAddAddin = new System.Windows.Forms.VB.CommandButton();
            this.Label11 = new System.Windows.Forms.VB.Label();
            this.Label10 = new System.Windows.Forms.VB.Label();
            this.SuspendLayout();
            //
            // txtObjectName
            //
            this.txtObjectName.Name = "txtObjectName";
            this.txtObjectName.Size = new System.Drawing.Size(187, 21);
            this.txtObjectName.Location = new System.Drawing.Point(144, 32);
            this.txtObjectName.TabIndex = 1;
            this.txtObjectName.Text = "SliceAndDice.Wizard";
            //
            // cmdRemoveAddin
            //
            this.cmdRemoveAddin.Name = "cmdRemoveAddin";
            this.cmdRemoveAddin.Text = "Remove from VB5 Add-in list";
            this.cmdRemoveAddin.Size = new System.Drawing.Size(187, 29);
            this.cmdRemoveAddin.Location = new System.Drawing.Point(146, 102);
            this.cmdRemoveAddin.TabIndex = 3;
            //
            // txtLocation
            //
            this.txtLocation.Name = "txtLocation";
            this.txtLocation.Size = new System.Drawing.Size(187, 21);
            this.txtLocation.Location = new System.Drawing.Point(144, 8);
            this.txtLocation.TabIndex = 0;
            this.txtLocation.Text = "c:\windows\vbaddin.ini";
            //
            // cmdAddAddin
            //
            this.cmdAddAddin.Name = "cmdAddAddin";
            this.cmdAddAddin.Text = "Add to VB5 Add-in list";
            this.cmdAddAddin.Size = new System.Drawing.Size(187, 29);
            this.cmdAddAddin.Location = new System.Drawing.Point(146, 62);
            this.cmdAddAddin.TabIndex = 2;
            //
            // Label11
            //
            this.Label11.Name = "Label11";
            this.Label11.Text = "Addin Object Name";
            this.Label11.Size = new System.Drawing.Size(92, 13);
            this.Label11.Location = new System.Drawing.Point(44, 36);
            this.Label11.TabIndex = 5;
            //
            // Label10
            //
            this.Label10.Name = "Label10";
            this.Label10.Text = "Addin list Location";
            this.Label10.Size = new System.Drawing.Size(86, 13);
            this.Label10.Location = new System.Drawing.Point(52, 12);
            this.Label10.TabIndex = 4;
            //
            // frmRegisterCodeSnippet
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.txtObjectName,
                  this.cmdRemoveAddin,
                  this.txtLocation,
                  this.cmdAddAddin,
                  this.Label11,
                  this.Label10
            });
            this.Name = "frmRegisterCodeSnippet";
            this.ResumeLayout(false);
        }
        #endregion

            public void RegisterIt()
            {
                ;
                ;

                fh = FreeFile;
                ;
                //  ' Make sure it's not already there;
                #fh     Open txtLocation For Input Access Read;
                Do Until EOF(fh);
                Input #fh, sLine;
                sLine).Contains(UCase(txtObjectName)) > 0 .ToUpper()
            {;
                MsgBox "'" + txtObjectName + "' has already been registered. Aborting action.";
                Close #fh;
                return;
                };
                Loop;
                Close #fh;
                ;
                #fh     Open txtLocation For Append Access Write;
                Print #fh, txtObjectName + "=0";
                Close #fh;
            }

            public void cmdAddAddin_Click()
            {
                RegisterIt;
                MsgBox "Code Snippet Add-in added successfully. Use the Add-in Manager to activate it.";
            }

            public void cmdRemoveAddin_Click()
            {
                ;
                ;
                ;

                fh = FreeFile;
                #fh     Open txtLocation For Input Access Read;
                Do Until EOF(fh);
                Input #fh, sLine;
                sLine.Contains(txtObjectName) = 0 )
            {;
                sBackOut +=  sLine + Chr$(13) + Chr$(10);
                };
                Loop;
                Close #fh;
                ;
                #fh     Open txtLocation For Output Access Write;
                Print #fh, sBackOut;
                Close #fh;

                MsgBox "Code Snippet Add-in removed successfully.";
            }

            public void Form_Load()
            {
                txtLocation.Text = sGetWindowsDir() + "vbaddin.ini";
                ;
                if ( Command) = "REGISTER SLICE AND DICE" .ToUpper()
            {;
                RegisterIt;
                End;
                };
            }

        }
    }
