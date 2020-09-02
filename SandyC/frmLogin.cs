using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class frmLogin : System.Windows.Forms.Form
        {
         public System.Windows.Forms.VB.TextBox txtUserName5;
         public System.Windows.Forms.VB.TextBox txtPassword5;
         public System.Windows.Forms.VB.TextBox txtUserName4;
         public System.Windows.Forms.VB.TextBox txtPassword4;
         public System.Windows.Forms.VB.TextBox txtUserName0;
         public System.Windows.Forms.VB.TextBox txtPassword1;
         public System.Windows.Forms.VB.TextBox txtUserName1;
         public System.Windows.Forms.VB.TextBox txtPassword0;
         public System.Windows.Forms.VB.TextBox txtPassword3;
         public System.Windows.Forms.VB.TextBox txtUserName3;
         public System.Windows.Forms.VB.TextBox txtPassword2;
         public System.Windows.Forms.VB.TextBox txtUserName2;
         public System.Windows.Forms.VB.CommandButton cmdOK;
         public System.Windows.Forms.VB.CommandButton cmdCancel;
         public System.Windows.Forms.VB.Label lblLabels11;
         public System.Windows.Forms.VB.Label lblLabels10;
         public System.Windows.Forms.VB.Label lblLabels9;
         public System.Windows.Forms.VB.Label lblLabels8;
         public System.Windows.Forms.VB.Label lblLabels7;
         public System.Windows.Forms.VB.Label lblLabels6;
         public System.Windows.Forms.VB.Label lblLabels5;
         public System.Windows.Forms.VB.Label lblLabels4;
         public System.Windows.Forms.VB.Label lblLabels0;
         public System.Windows.Forms.VB.Label lblLabels3;
         public System.Windows.Forms.VB.Label lblLabels2;
         public System.Windows.Forms.VB.Label lblLabels1;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.Container components = null;

        public frmLogin()
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
            this.txtUserName5 = new System.Windows.Forms.VB.TextBox();
            this.txtPassword5 = new System.Windows.Forms.VB.TextBox();
            this.txtUserName4 = new System.Windows.Forms.VB.TextBox();
            this.txtPassword4 = new System.Windows.Forms.VB.TextBox();
            this.txtUserName0 = new System.Windows.Forms.VB.TextBox();
            this.txtPassword1 = new System.Windows.Forms.VB.TextBox();
            this.txtUserName1 = new System.Windows.Forms.VB.TextBox();
            this.txtPassword0 = new System.Windows.Forms.VB.TextBox();
            this.txtPassword3 = new System.Windows.Forms.VB.TextBox();
            this.txtUserName3 = new System.Windows.Forms.VB.TextBox();
            this.txtPassword2 = new System.Windows.Forms.VB.TextBox();
            this.txtUserName2 = new System.Windows.Forms.VB.TextBox();
            this.cmdOK = new System.Windows.Forms.VB.CommandButton();
            this.cmdCancel = new System.Windows.Forms.VB.CommandButton();
            this.lblLabels11 = new System.Windows.Forms.VB.Label();
            this.lblLabels10 = new System.Windows.Forms.VB.Label();
            this.lblLabels9 = new System.Windows.Forms.VB.Label();
            this.lblLabels8 = new System.Windows.Forms.VB.Label();
            this.lblLabels7 = new System.Windows.Forms.VB.Label();
            this.lblLabels6 = new System.Windows.Forms.VB.Label();
            this.lblLabels5 = new System.Windows.Forms.VB.Label();
            this.lblLabels4 = new System.Windows.Forms.VB.Label();
            this.lblLabels0 = new System.Windows.Forms.VB.Label();
            this.lblLabels3 = new System.Windows.Forms.VB.Label();
            this.lblLabels2 = new System.Windows.Forms.VB.Label();
            this.lblLabels1 = new System.Windows.Forms.VB.Label();
            this.SuspendLayout();
            //
            // txtUserName5
            //
            this.txtUserName5.Name = "txtUserName5";
            this.txtUserName5.Size = new System.Drawing.Size(155, 23);
            this.txtUserName5.Location = new System.Drawing.Point(10, 54);
            this.txtUserName5.TabIndex = 24;
            //
            // txtPassword5
            //
            this.txtPassword5.Name = "txtPassword5";
            this.txtPassword5.Size = new System.Drawing.Size(155, 23);
            this.txtPassword5.Location = new System.Drawing.Point(476, 54);
            this.txtPassword5.TabIndex = 21;
            //
            // txtUserName4
            //
            this.txtUserName4.Name = "txtUserName4";
            this.txtUserName4.Size = new System.Drawing.Size(155, 23);
            this.txtUserName4.Location = new System.Drawing.Point(320, 54);
            this.txtUserName4.TabIndex = 20;
            //
            // txtPassword4
            //
            this.txtPassword4.Name = "txtPassword4";
            this.txtPassword4.Size = new System.Drawing.Size(155, 23);
            this.txtPassword4.Location = new System.Drawing.Point(165, 54);
            this.txtPassword4.TabIndex = 18;
            //
            // txtUserName0
            //
            this.txtUserName0.Name = "txtUserName0";
            this.txtUserName0.Size = new System.Drawing.Size(155, 23);
            this.txtUserName0.Location = new System.Drawing.Point(10, 10);
            this.txtUserName0.TabIndex = 12;
            this.txtUserName0.Text = "Sequence from Server";
            //
            // txtPassword1
            //
            this.txtPassword1.Name = "txtPassword1";
            this.txtPassword1.Size = new System.Drawing.Size(155, 23);
            this.txtPassword1.Location = new System.Drawing.Point(476, 10);
            this.txtPassword1.TabIndex = 9;
            this.txtPassword1.Text = "Fx(Local Generated)";
            //
            // txtUserName1
            //
            this.txtUserName1.Name = "txtUserName1";
            this.txtUserName1.Size = new System.Drawing.Size(155, 23);
            this.txtUserName1.Location = new System.Drawing.Point(320, 10);
            this.txtUserName1.TabIndex = 8;
            this.txtUserName1.Text = "Checksum";
            //
            // txtPassword0
            //
            this.txtPassword0.Name = "txtPassword0";
            this.txtPassword0.Size = new System.Drawing.Size(155, 23);
            this.txtPassword0.Location = new System.Drawing.Point(165, 10);
            this.txtPassword0.TabIndex = 6;
            this.txtPassword0.Text = "Random";
            //
            // txtPassword3
            //
            this.txtPassword3.Name = "txtPassword3";
            this.txtPassword3.Size = new System.Drawing.Size(155, 23);
            this.txtPassword3.Location = new System.Drawing.Point(476, 98);
            this.txtPassword3.TabIndex = 5;
            //
            // txtUserName3
            //
            this.txtUserName3.Name = "txtUserName3";
            this.txtUserName3.Size = new System.Drawing.Size(155, 23);
            this.txtUserName3.Location = new System.Drawing.Point(320, 98);
            this.txtUserName3.TabIndex = 4;
            //
            // txtPassword2
            //
            this.txtPassword2.Name = "txtPassword2";
            this.txtPassword2.Size = new System.Drawing.Size(155, 23);
            this.txtPassword2.Location = new System.Drawing.Point(165, 98);
            this.txtPassword2.TabIndex = 3;
            //
            // txtUserName2
            //
            this.txtUserName2.Name = "txtUserName2";
            this.txtUserName2.Size = new System.Drawing.Size(155, 23);
            this.txtUserName2.Location = new System.Drawing.Point(10, 98);
            this.txtUserName2.TabIndex = 2;
            //
            // cmdOK
            //
            this.cmdOK.Name = "cmdOK";
            this.cmdOK.Text = "OK";
            this.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.cmdOK.Size = new System.Drawing.Size(76, 26);
            this.cmdOK.Location = new System.Drawing.Point(8, 150);
            this.cmdOK.TabIndex = 0;
            //
            // cmdCancel
            //
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.Size = new System.Drawing.Size(76, 26);
            this.cmdCancel.Location = new System.Drawing.Point(86, 150);
            this.cmdCancel.TabIndex = 1;
            //
            // lblLabels11
            //
            this.lblLabels11.Name = "lblLabels11";
            this.lblLabels11.Text = "Rep Public ID";
            this.lblLabels11.Size = new System.Drawing.Size(66, 13);
            this.lblLabels11.Location = new System.Drawing.Point(10, 80);
            this.lblLabels11.TabIndex = 25;
            //
            // lblLabels10
            //
            this.lblLabels10.Name = "lblLabels10";
            this.lblLabels10.Text = "Rep Remote Confirmation";
            this.lblLabels10.Size = new System.Drawing.Size(121, 13);
            this.lblLabels10.Location = new System.Drawing.Point(476, 80);
            this.lblLabels10.TabIndex = 23;
            //
            // lblLabels9
            //
            this.lblLabels9.Name = "lblLabels9";
            this.lblLabels9.Text = "Rep Remote Generated";
            this.lblLabels9.Size = new System.Drawing.Size(113, 13);
            this.lblLabels9.Location = new System.Drawing.Point(320, 80);
            this.lblLabels9.TabIndex = 22;
            //
            // lblLabels8
            //
            this.lblLabels8.Name = "lblLabels8";
            this.lblLabels8.Text = "Rep Local Generated";
            this.lblLabels8.Size = new System.Drawing.Size(102, 13);
            this.lblLabels8.Location = new System.Drawing.Point(165, 80);
            this.lblLabels8.TabIndex = 19;
            //
            // lblLabels7
            //
            this.lblLabels7.Name = "lblLabels7";
            this.lblLabels7.Text = "Team Public ID";
            this.lblLabels7.Size = new System.Drawing.Size(73, 13);
            this.lblLabels7.Location = new System.Drawing.Point(10, 124);
            this.lblLabels7.TabIndex = 17;
            //
            // lblLabels6
            //
            this.lblLabels6.Name = "lblLabels6";
            this.lblLabels6.Text = "Team Remote Confirmation";
            this.lblLabels6.Size = new System.Drawing.Size(128, 13);
            this.lblLabels6.Location = new System.Drawing.Point(476, 124);
            this.lblLabels6.TabIndex = 16;
            //
            // lblLabels5
            //
            this.lblLabels5.Name = "lblLabels5";
            this.lblLabels5.Text = "Team Remote Generated";
            this.lblLabels5.Size = new System.Drawing.Size(120, 13);
            this.lblLabels5.Location = new System.Drawing.Point(320, 124);
            this.lblLabels5.TabIndex = 15;
            //
            // lblLabels4
            //
            this.lblLabels4.Name = "lblLabels4";
            this.lblLabels4.Text = "Team Local Generated";
            this.lblLabels4.Size = new System.Drawing.Size(109, 13);
            this.lblLabels4.Location = new System.Drawing.Point(165, 124);
            this.lblLabels4.TabIndex = 14;
            //
            // lblLabels0
            //
            this.lblLabels0.Name = "lblLabels0";
            this.lblLabels0.Text = "SHPC Public ID";
            this.lblLabels0.Size = new System.Drawing.Size(75, 13);
            this.lblLabels0.Location = new System.Drawing.Point(10, 36);
            this.lblLabels0.TabIndex = 13;
            //
            // lblLabels3
            //
            this.lblLabels3.Name = "lblLabels3";
            this.lblLabels3.Text = "SHPC Remote Confirmation";
            this.lblLabels3.Size = new System.Drawing.Size(130, 13);
            this.lblLabels3.Location = new System.Drawing.Point(476, 36);
            this.lblLabels3.TabIndex = 11;
            //
            // lblLabels2
            //
            this.lblLabels2.Name = "lblLabels2";
            this.lblLabels2.Text = "SHPC Remote Generated";
            this.lblLabels2.Size = new System.Drawing.Size(122, 13);
            this.lblLabels2.Location = new System.Drawing.Point(320, 36);
            this.lblLabels2.TabIndex = 10;
            //
            // lblLabels1
            //
            this.lblLabels1.Name = "lblLabels1";
            this.lblLabels1.Text = "SHPC Local Generated";
            this.lblLabels1.Size = new System.Drawing.Size(111, 13);
            this.lblLabels1.Location = new System.Drawing.Point(165, 36);
            this.lblLabels1.TabIndex = 7;
            //
            // frmLogin
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.txtUserName5,
                  this.txtPassword5,
                  this.txtUserName4,
                  this.txtPassword4,
                  this.txtUserName0,
                  this.txtPassword1,
                  this.txtUserName1,
                  this.txtPassword0,
                  this.txtPassword3,
                  this.txtUserName3,
                  this.txtPassword2,
                  this.txtUserName2,
                  this.cmdOK,
                  this.cmdCancel,
                  this.lblLabels11,
                  this.lblLabels10,
                  this.lblLabels9,
                  this.lblLabels8,
                  this.lblLabels7,
                  this.lblLabels6,
                  this.lblLabels5,
                  this.lblLabels4,
                  this.lblLabels0,
                  this.lblLabels3,
                  this.lblLabels2,
                  this.lblLabels1
            });
            this.Name = "frmLogin";
            this.ResumeLayout(false);
        }
        #endregion

        public bool LoginSucceeded;


            public void cmdCancel_Click()
            {
                //    'set the global var to false;
                //    'to denote a failed login;
                LoginSucceeded = false;
                this.Hide;
            }

            public void cmdOK_Click()
            {
                //    'check for correct password;
                if ( txtPassword = "password" )
            {;
                //        'place code to here to pass the;
                //        'success to the calling sub;
                //        'setting a global var is the easiest;
                LoginSucceeded = true;
                this.Hide;
                }
            else
            {;
                MsgBox "Invalid Password, try again!", , "Login";
                txtPassword.SetFocus;
                SendKeys "{Home}+{End}";
                };
            }

        }
    }
