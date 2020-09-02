using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class frmTest : System.Windows.Forms.Form
        {
         public System.Windows.Forms.VB.Label Label2;
         public System.Windows.Forms.VB.Image Image1;
         public System.Windows.Forms.VB.Label Label1;
         public System.Windows.Forms.VB.Label lblInfo;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.Container components = null;

        public frmTest()
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
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTest));
            this.Label2 = new System.Windows.Forms.VB.Label();
            this.Image1 = new System.Windows.Forms.VB.Image();
            this.Label1 = new System.Windows.Forms.VB.Label();
            this.lblInfo = new System.Windows.Forms.VB.Label();
            this.SuspendLayout();
            //
            // Label2
            //
            this.Label2.Name = "Label2";
            this.Label2.Text = "The HotKey for this demo has been set to CTRL-ALT-UP";
            this.Label2.Font = new System.Drawing.Font("Tahoma",8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.Label2.Size = new System.Drawing.Size(269, 21);
            this.Label2.Location = new System.Drawing.Point(4, 144);
            this.Label2.TabIndex = 2;
            //
            // Image1
            //
            this.Image1.Name = "Image1";
            this.Image1.Size = new System.Drawing.Size(128, 24);
            this.Image1.Location = new System.Drawing.Point(8, 8);
            this.Image1.Image = ((System.Drawing.Bitmap)(resources.GetObject("Image1.Image")));
            //
            // Label1
            //
            this.Label1.Name = "Label1";
            this.Label1.BackColor = System.Drawing.Color.FromArgb(0);
            this.Label1.Text = "Label1";
            this.Label1.Size = new System.Drawing.Size(273, 33);
            this.Label1.Location = new System.Drawing.Point(4, 4);
            this.Label1.TabIndex = 1;
            //
            // lblInfo
            //
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Text = $"frmTest.frx":197E;
            this.lblInfo.Font = new System.Drawing.Font("Tahoma",8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblInfo.Size = new System.Drawing.Size(269, 77);
            this.lblInfo.Location = new System.Drawing.Point(8, 52);
            this.lblInfo.TabIndex = 0;
            //
            // frmTest
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.Label2,
                  this.Image1,
                  this.Label1,
                  this.lblInfo
            });
            this.Name = "frmTest";
            this.ResumeLayout(false);
        }
        #endregion

        public object WithEvents;


            public void Form_Load()
            {
                m_cHotKey = new cRegHotKey();
                m_cHotKey.Attach this.hwnd;
                m_cHotKey.RegisterKey "Activate", vbKeyUp, MOD_ALT + MOD_CONTROL;
            }

            public void m_cHotKey_HotKeyPress            {
                m_cHotKey.RestoreAndActivate this.hwnd;
                MsgBox "Got HotKey: " + sName;
            }

        }
    }
