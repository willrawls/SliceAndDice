using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class Form1 : System.Windows.Forms.Form
        {
         public System.Windows.Forms.VB.CommandButton cmdDecrypt;
         public System.Windows.Forms.VB.TextBox txtDec;
         public System.Windows.Forms.VB.CommandButton cmdEncrypt;
         public System.Windows.Forms.VB.TextBox txtOut;
         public System.Windows.Forms.VB.TextBox txtIn;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.Container components = null;

        public Form1()
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
            this.cmdDecrypt = new System.Windows.Forms.VB.CommandButton();
            this.txtDec = new System.Windows.Forms.VB.TextBox();
            this.cmdEncrypt = new System.Windows.Forms.VB.CommandButton();
            this.txtOut = new System.Windows.Forms.VB.TextBox();
            this.txtIn = new System.Windows.Forms.VB.TextBox();
            this.SuspendLayout();
            //
            // cmdDecrypt
            //
            this.cmdDecrypt.Name = "cmdDecrypt";
            this.cmdDecrypt.Text = "&Decrypt";
            this.cmdDecrypt.Size = new System.Drawing.Size(83, 35);
            this.cmdDecrypt.Location = new System.Drawing.Point(286, 208);
            this.cmdDecrypt.TabIndex = 4;
            //
            // txtDec
            //
            this.txtDec.Name = "txtDec";
            this.txtDec.Size = new System.Drawing.Size(279, 177);
            this.txtDec.Location = new System.Drawing.Point(2, 230);
//            this.txtDec.MultiLine = -1;
            this.txtDec.TabIndex = 3;
            //
            // cmdEncrypt
            //
            this.cmdEncrypt.Name = "cmdEncrypt";
            this.cmdEncrypt.Text = "&Encrypt";
            this.cmdEncrypt.Size = new System.Drawing.Size(83, 35);
            this.cmdEncrypt.Location = new System.Drawing.Point(286, 34);
            this.cmdEncrypt.TabIndex = 2;
            //
            // txtOut
            //
            this.txtOut.Name = "txtOut";
            this.txtOut.Size = new System.Drawing.Size(279, 177);
            this.txtOut.Location = new System.Drawing.Point(2, 53);
//            this.txtOut.MultiLine = -1;
            this.txtOut.TabIndex = 1;
            //
            // txtIn
            //
            this.txtIn.Name = "txtIn";
            this.txtIn.Size = new System.Drawing.Size(279, 51);
            this.txtIn.Location = new System.Drawing.Point(2, 2);
//            this.txtIn.MultiLine = -1;
            this.txtIn.TabIndex = 0;
            this.txtIn.Text = "TestCrypt.frx":0000;
            //
            // Form1
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.cmdDecrypt,
                  this.txtDec,
                  this.cmdEncrypt,
                  this.txtOut,
                  this.txtIn
            });
            this.Name = "Form1";
            this.ResumeLayout(false);
        }
        #endregion

            public void cmdDecrypt_Click()
            {
                txtDec.Text = sadDecrypt(txtOut.Text);
            }

            public void cmdEncrypt_Click()
            {
                txtOut.Text = sadEncrypt(txtIn.Text);
            }

        }
    }
