using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class frmSplash : System.Windows.Forms.Form
        {
         public System.Windows.Forms.VB.Frame Frame1;
         public System.Windows.Forms.VB.Label Label14;
         public System.Windows.Forms.VB.Label Label411;
         public System.Windows.Forms.VB.Label Label12;
         public System.Windows.Forms.VB.Label Label42;
         public System.Windows.Forms.VB.Label lblDaysLeft;
         public System.Windows.Forms.VB.Label lblDayLeftCaption;
         public System.Windows.Forms.VB.Label lblDLLsLoaded1;
         public System.Windows.Forms.VB.Shape Shape21;
         public System.Windows.Forms.VB.Label TeamIP;
         public System.Windows.Forms.VB.Label CentralIP;
         public System.Windows.Forms.VB.Label Label410;
         public System.Windows.Forms.VB.Label Label49;
         public System.Windows.Forms.VB.Label ParentRepID;
         public System.Windows.Forms.VB.Label UserRepID;
         public System.Windows.Forms.VB.Label TeamID;
         public System.Windows.Forms.VB.Label UserID;
         public System.Windows.Forms.VB.Label Label48;
         public System.Windows.Forms.VB.Label Label47;
         public System.Windows.Forms.VB.Label Label46;
         public System.Windows.Forms.VB.Label Label45;
         public System.Windows.Forms.VB.Label Label31;
         public System.Windows.Forms.VB.Label lblDLLsLoaded0;
         public System.Windows.Forms.VB.Shape Shape20;
         public System.Windows.Forms.VB.Label Label44;
         public System.Windows.Forms.VB.Label Label43;
         public System.Windows.Forms.VB.Label Label13;
         public System.Windows.Forms.VB.Label Label41;
         public System.Windows.Forms.VB.Label Label11;
         public System.Windows.Forms.VB.Label Label40;
         public System.Windows.Forms.VB.Label Label30;
         public System.Windows.Forms.VB.Label lblPlatform;
         public System.Windows.Forms.VB.Label Label2;
         public System.Windows.Forms.VB.Label Label10;
         public System.Windows.Forms.VB.Label lblCopyright;
         public System.Windows.Forms.VB.Label lblWarning;
         public System.Windows.Forms.VB.Label lblVersion;
         public System.Windows.Forms.VB.Label lblLicenseTo;
         public System.Windows.Forms.VB.Label lblCompanyProduct;
         public System.Windows.Forms.VB.Label lblProductName;
         public System.Windows.Forms.VB.Shape Shape1;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.Container components = null;

        public frmSplash()
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
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSplash));
            this.Frame1 = new System.Windows.Forms.VB.Frame();
            this.Label14 = new System.Windows.Forms.VB.Label();
            this.Label411 = new System.Windows.Forms.VB.Label();
            this.Label12 = new System.Windows.Forms.VB.Label();
            this.Label42 = new System.Windows.Forms.VB.Label();
            this.lblDaysLeft = new System.Windows.Forms.VB.Label();
            this.lblDayLeftCaption = new System.Windows.Forms.VB.Label();
            this.lblDLLsLoaded1 = new System.Windows.Forms.VB.Label();
            this.Shape21 = new System.Windows.Forms.VB.Shape();
            this.TeamIP = new System.Windows.Forms.VB.Label();
            this.CentralIP = new System.Windows.Forms.VB.Label();
            this.Label410 = new System.Windows.Forms.VB.Label();
            this.Label49 = new System.Windows.Forms.VB.Label();
            this.ParentRepID = new System.Windows.Forms.VB.Label();
            this.UserRepID = new System.Windows.Forms.VB.Label();
            this.TeamID = new System.Windows.Forms.VB.Label();
            this.UserID = new System.Windows.Forms.VB.Label();
            this.Label48 = new System.Windows.Forms.VB.Label();
            this.Label47 = new System.Windows.Forms.VB.Label();
            this.Label46 = new System.Windows.Forms.VB.Label();
            this.Label45 = new System.Windows.Forms.VB.Label();
            this.Label31 = new System.Windows.Forms.VB.Label();
            this.lblDLLsLoaded0 = new System.Windows.Forms.VB.Label();
            this.Shape20 = new System.Windows.Forms.VB.Shape();
            this.Label44 = new System.Windows.Forms.VB.Label();
            this.Label43 = new System.Windows.Forms.VB.Label();
            this.Label13 = new System.Windows.Forms.VB.Label();
            this.Label41 = new System.Windows.Forms.VB.Label();
            this.Label11 = new System.Windows.Forms.VB.Label();
            this.Label40 = new System.Windows.Forms.VB.Label();
            this.Label30 = new System.Windows.Forms.VB.Label();
            this.lblPlatform = new System.Windows.Forms.VB.Label();
            this.Label2 = new System.Windows.Forms.VB.Label();
            this.Label10 = new System.Windows.Forms.VB.Label();
            this.lblCopyright = new System.Windows.Forms.VB.Label();
            this.lblWarning = new System.Windows.Forms.VB.Label();
            this.lblVersion = new System.Windows.Forms.VB.Label();
            this.lblLicenseTo = new System.Windows.Forms.VB.Label();
            this.lblCompanyProduct = new System.Windows.Forms.VB.Label();
            this.lblProductName = new System.Windows.Forms.VB.Label();
            this.Shape1 = new System.Windows.Forms.VB.Shape();
            this.SuspendLayout();
            this.Frame1.SuspendLayout();
            //
            // Frame1
            //
            this.Frame1.Name = "Frame1";
            this.Frame1.BackColor = System.Drawing.Color.Transparent;
            this.Frame1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Frame1.ForeColor = System.Drawing.Color.Transparent;
            this.Frame1.Size = new System.Drawing.Size(568, 308);
            this.Frame1.Location = new System.Drawing.Point(4, 4);
            this.Frame1.TabIndex = 0;
//            this.Frame1.ToolTipText = "Click on any non-web link to close this form. Thank you for using Slice and Dice !";
            this.Frame1.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.Label14,
                        this.Label411,
                        this.Label12,
                        this.Label42,
                        this.lblDaysLeft,
                        this.lblDayLeftCaption,
                        this.lblDLLsLoaded1,
                        this.Shape21,
                        this.TeamIP,
                        this.CentralIP,
                        this.Label410,
                        this.Label49,
                        this.ParentRepID,
                        this.UserRepID,
                        this.TeamID,
                        this.UserID,
                        this.Label48,
                        this.Label47,
                        this.Label46,
                        this.Label45,
                        this.Label31,
                        this.lblDLLsLoaded0,
                        this.Shape20,
                        this.Label44,
                        this.Label43,
                        this.Label13,
                        this.Label41,
                        this.Label11,
                        this.Label40,
                        this.Label30,
                        this.lblPlatform,
                        this.Label2,
                        this.Label10,
                        this.lblCopyright,
                        this.lblWarning,
                        this.lblVersion,
                        this.lblLicenseTo,
                        this.lblCompanyProduct,
                        this.lblProductName,
                        this.Shape1
            });
            //
            // Label14
            //
            this.Label14.Name = "Label14";
            this.Label14.Text = "http://willrawls.ewebcity.com/coreTeam";
            this.Label14.Font = new System.Drawing.Font("Arial",8F, ( System.Drawing.FontStyle.Underline ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.Label14.ForeColor = System.Drawing.Color.Transparent;
            this.Label14.Size = new System.Drawing.Size(193, 12);
            this.Label14.Location = new System.Drawing.Point(92, 223);
//            this.Label14.MouseIcon = "frmSplash.frx":000C;
//            this.Label14.MousePointer = 99;
            this.Label14.TabIndex = 37;
//            this.Label14.ToolTipText = "Manually submit a template or category for inclusion in the general S&D product.";
            //
            // Label411
            //
            this.Label411.Name = "Label411";
            this.Label411.Text = "Opensource Dev:";
            this.Label411.ForeColor = System.Drawing.Color.Transparent;
            this.Label411.Size = new System.Drawing.Size(84, 12);
            this.Label411.Location = new System.Drawing.Point(2, 223);
            this.Label411.TabIndex = 36;
            //
            // Label12
            //
            this.Label12.Name = "Label12";
            this.Label12.Text = "http://willrawls.ewebcity.com/sndforum";
            this.Label12.Font = new System.Drawing.Font("Arial",8F, ( System.Drawing.FontStyle.Underline ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.Label12.ForeColor = System.Drawing.Color.Transparent;
            this.Label12.Size = new System.Drawing.Size(189, 12);
            this.Label12.Location = new System.Drawing.Point(92, 206);
//            this.Label12.MouseIcon = "frmSplash.frx":044E;
//            this.Label12.MousePointer = 99;
            this.Label12.TabIndex = 35;
//            this.Label12.ToolTipText = "Manually submit a template or category for inclusion in the general S&D product.";
            //
            // Label42
            //
            this.Label42.Name = "Label42";
            this.Label42.Text = "User forums:";
            this.Label42.ForeColor = System.Drawing.Color.Transparent;
            this.Label42.Size = new System.Drawing.Size(60, 12);
            this.Label42.Location = new System.Drawing.Point(2, 206);
            this.Label42.TabIndex = 34;
            //
            // lblDaysLeft
            //
            this.lblDaysLeft.Name = "lblDaysLeft";
            this.lblDaysLeft.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.lblDaysLeft.Text = "??????????";
            this.lblDaysLeft.Font = new System.Drawing.Font("MS Sans Serif",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblDaysLeft.ForeColor = System.Drawing.Color.Transparent;
            this.lblDaysLeft.Size = new System.Drawing.Size(113, 13);
            this.lblDaysLeft.Location = new System.Drawing.Point(448, 26);
            this.lblDaysLeft.TabIndex = 33;
            //
            // lblDayLeftCaption
            //
            this.lblDayLeftCaption.Name = "lblDayLeftCaption";
            this.lblDayLeftCaption.Text = "Days Left in your Free Evaluation Period:";
            this.lblDayLeftCaption.Font = new System.Drawing.Font("MS Sans Serif",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblDayLeftCaption.ForeColor = System.Drawing.Color.Transparent;
            this.lblDayLeftCaption.Size = new System.Drawing.Size(108, 39);
            this.lblDayLeftCaption.Location = new System.Drawing.Point(448, 0);
            this.lblDayLeftCaption.TabIndex = 32;
            //
            // lblDLLsLoaded1
            //
            this.lblDLLsLoaded1.Name = "lblDLLsLoaded1";
            this.lblDLLsLoaded1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.lblDLLsLoaded1.Text = "0";
            this.lblDLLsLoaded1.ForeColor = System.Drawing.Color.Transparent;
            this.lblDLLsLoaded1.Size = new System.Drawing.Size(6, 12);
            this.lblDLLsLoaded1.Location = new System.Drawing.Point(496, 159);
            this.lblDLLsLoaded1.TabIndex = 31;
//            this.lblDLLsLoaded1.ToolTipText = "This is the number of S&D add-in DLLs currently loaded in memory.";
            //
            // Shape21
            //
            this.Shape21.Name = "Shape21";
            this.Shape21.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Shape21.Size = new System.Drawing.Size(127, 45);
            this.Shape21.Location = new System.Drawing.Point(444, -2);
            //
            // TeamIP
            //
            this.TeamIP.Name = "TeamIP";
            this.TeamIP.Text = "209.196.104.22";
            this.TeamIP.ForeColor = System.Drawing.Color.FromArgb(0);
            this.TeamIP.Size = new System.Drawing.Size(75, 13);
            this.TeamIP.Location = new System.Drawing.Point(480, 233);
            this.TeamIP.TabIndex = 30;
//            this.TeamIP.ToolTipText = "This is the TCP/IP address of your Team;
            //
            // CentralIP
            //
            this.CentralIP.Name = "CentralIP";
            this.CentralIP.Text = "209.196.104.22";
            this.CentralIP.ForeColor = System.Drawing.Color.FromArgb(0);
            this.CentralIP.Size = new System.Drawing.Size(75, 13);
            this.CentralIP.Location = new System.Drawing.Point(480, 220);
            this.CentralIP.TabIndex = 29;
//            this.CentralIP.ToolTipText = "This is the TCP/IP address of the S&D Internet Server.";
            //
            // Label410
            //
            this.Label410.Name = "Label410";
            this.Label410.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.Label410.Text = "Team IP:";
            this.Label410.ForeColor = System.Drawing.Color.FromArgb(0);
            this.Label410.Size = new System.Drawing.Size(43, 13);
            this.Label410.Location = new System.Drawing.Point(432, 233);
            this.Label410.TabIndex = 28;
            //
            // Label49
            //
            this.Label49.Name = "Label49";
            this.Label49.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.Label49.Text = "Central IP:";
            this.Label49.ForeColor = System.Drawing.Color.FromArgb(0);
            this.Label49.Size = new System.Drawing.Size(49, 13);
            this.Label49.Location = new System.Drawing.Point(426, 220);
            this.Label49.TabIndex = 27;
            //
            // ParentRepID
            //
            this.ParentRepID.Name = "ParentRepID";
            this.ParentRepID.Text = "999-999999";
            this.ParentRepID.ForeColor = System.Drawing.Color.FromArgb(0);
            this.ParentRepID.Size = new System.Drawing.Size(57, 13);
            this.ParentRepID.Location = new System.Drawing.Point(480, 260);
            this.ParentRepID.TabIndex = 26;
//            this.ParentRepID.ToolTipText = "This is the ID of the Representative who made you a S&D representative.";
            //
            // UserRepID
            //
            this.UserRepID.Name = "UserRepID";
            this.UserRepID.Text = "999-999999";
            this.UserRepID.ForeColor = System.Drawing.Color.FromArgb(0);
            this.UserRepID.Size = new System.Drawing.Size(57, 13);
            this.UserRepID.Location = new System.Drawing.Point(480, 247);
            this.UserRepID.TabIndex = 25;
//            this.UserRepID.ToolTipText = "This is your S&D Representative ID (Only if you;
            //
            // TeamID
            //
            this.TeamID.Name = "TeamID";
            this.TeamID.Text = "999-999999";
            this.TeamID.ForeColor = System.Drawing.Color.FromArgb(0);
            this.TeamID.Size = new System.Drawing.Size(57, 13);
            this.TeamID.Location = new System.Drawing.Point(480, 206);
            this.TeamID.TabIndex = 24;
//            this.TeamID.ToolTipText = "This is your Team;
            //
            // UserID
            //
            this.UserID.Name = "UserID";
            this.UserID.Text = "999-999999";
            this.UserID.ForeColor = System.Drawing.Color.Transparent;
            this.UserID.Size = new System.Drawing.Size(57, 13);
            this.UserID.Location = new System.Drawing.Point(480, 193);
            this.UserID.TabIndex = 23;
//            this.UserID.ToolTipText = "This is your S&D serial number.";
            //
            // Label48
            //
            this.Label48.Name = "Label48";
            this.Label48.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.Label48.Text = "Parent Rep ID:";
            this.Label48.ForeColor = System.Drawing.Color.FromArgb(0);
            this.Label48.Size = new System.Drawing.Size(71, 13);
            this.Label48.Location = new System.Drawing.Point(404, 260);
            this.Label48.TabIndex = 22;
            //
            // Label47
            //
            this.Label47.Name = "Label47";
            this.Label47.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.Label47.Text = "Your Rep ID:";
            this.Label47.ForeColor = System.Drawing.Color.FromArgb(0);
            this.Label47.Size = new System.Drawing.Size(62, 13);
            this.Label47.Location = new System.Drawing.Point(413, 247);
            this.Label47.TabIndex = 21;
            //
            // Label46
            //
            this.Label46.Name = "Label46";
            this.Label46.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.Label46.Text = "Team ID:";
            this.Label46.ForeColor = System.Drawing.Color.FromArgb(0);
            this.Label46.Size = new System.Drawing.Size(44, 13);
            this.Label46.Location = new System.Drawing.Point(431, 206);
            this.Label46.TabIndex = 20;
            //
            // Label45
            //
            this.Label45.Name = "Label45";
            this.Label45.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.Label45.Text = "Your ID:";
            this.Label45.ForeColor = System.Drawing.Color.Transparent;
            this.Label45.Size = new System.Drawing.Size(39, 13);
            this.Label45.Location = new System.Drawing.Point(436, 193);
            this.Label45.TabIndex = 19;
            //
            // Label31
            //
            this.Label31.Name = "Label31";
            this.Label31.Text = "Miscellaneous Info:";
            this.Label31.Font = new System.Drawing.Font("MS Sans Serif",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.Label31.ForeColor = System.Drawing.Color.Transparent;
            this.Label31.Size = new System.Drawing.Size(111, 13);
            this.Label31.Location = new System.Drawing.Point(400, 138);
            this.Label31.TabIndex = 18;
            //
            // lblDLLsLoaded0
            //
            this.lblDLLsLoaded0.Name = "lblDLLsLoaded0";
            this.lblDLLsLoaded0.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.lblDLLsLoaded0.Text = "Sandals Loaded:";
            this.lblDLLsLoaded0.ForeColor = System.Drawing.Color.Transparent;
            this.lblDLLsLoaded0.Size = new System.Drawing.Size(82, 12);
            this.lblDLLsLoaded0.Location = new System.Drawing.Point(409, 159);
            this.lblDLLsLoaded0.TabIndex = 17;
//            this.lblDLLsLoaded0.ToolTipText = "This is the number of S&D add-in DLLs currently loaded in memory.";
            //
            // Shape20
            //
            this.Shape20.Name = "Shape20";
            this.Shape20.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Shape20.Size = new System.Drawing.Size(179, 143);
            this.Shape20.Location = new System.Drawing.Point(396, 134);
            //
            // Label44
            //
            this.Label44.Name = "Label44";
            this.Label44.Text = "User Agreement:";
            this.Label44.ForeColor = System.Drawing.Color.Transparent;
            this.Label44.Size = new System.Drawing.Size(79, 13);
            this.Label44.Location = new System.Drawing.Point(2, 244);
            this.Label44.TabIndex = 16;
            //
            // Label43
            //
            this.Label43.Name = "Label43";
            this.Label43.Text = "Submit Template:";
            this.Label43.ForeColor = System.Drawing.Color.Transparent;
            this.Label43.Size = new System.Drawing.Size(82, 12);
            this.Label43.Location = new System.Drawing.Point(2, 188);
            this.Label43.TabIndex = 15;
            //
            // Label13
            //
            this.Label13.Name = "Label13";
            this.Label13.Text = "http://www.sliceanddice.com/submit.html";
            this.Label13.Font = new System.Drawing.Font("Arial",8F, ( System.Drawing.FontStyle.Underline ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.Label13.ForeColor = System.Drawing.Color.Transparent;
            this.Label13.Size = new System.Drawing.Size(199, 14);
            this.Label13.Location = new System.Drawing.Point(92, 188);
//            this.Label13.MouseIcon = "frmSplash.frx":0890;
//            this.Label13.MousePointer = 99;
            this.Label13.TabIndex = 14;
//            this.Label13.ToolTipText = "Manually submit a template or category for inclusion in the general S&D product.";
            //
            // Label41
            //
            this.Label41.Name = "Label41";
            this.Label41.Text = "Report an Issue:";
            this.Label41.ForeColor = System.Drawing.Color.Transparent;
            this.Label41.Size = new System.Drawing.Size(78, 13);
            this.Label41.Location = new System.Drawing.Point(2, 172);
            this.Label41.TabIndex = 13;
            //
            // Label11
            //
            this.Label11.Name = "Label11";
            this.Label11.Text = "http://www.sliceanddice.com/sadissue.html";
            this.Label11.Font = new System.Drawing.Font("Arial",8F, ( System.Drawing.FontStyle.Underline ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.Label11.ForeColor = System.Drawing.Color.Transparent;
            this.Label11.Size = new System.Drawing.Size(212, 14);
            this.Label11.Location = new System.Drawing.Point(92, 172);
//            this.Label11.MouseIcon = "frmSplash.frx":0CD2;
//            this.Label11.MousePointer = 99;
            this.Label11.TabIndex = 12;
//            this.Label11.ToolTipText = "Submit an issue / bug / feature directly to the S&D developer.";
            //
            // Label40
            //
            this.Label40.Name = "Label40";
            this.Label40.Text = "Latest updates:";
            this.Label40.ForeColor = System.Drawing.Color.Transparent;
            this.Label40.Size = new System.Drawing.Size(73, 13);
            this.Label40.Location = new System.Drawing.Point(2, 154);
            this.Label40.TabIndex = 11;
            //
            // Label30
            //
            this.Label30.Name = "Label30";
            this.Label30.Text = "Some places on the web to go for more information:";
            this.Label30.ForeColor = System.Drawing.Color.Transparent;
            this.Label30.Size = new System.Drawing.Size(242, 13);
            this.Label30.Location = new System.Drawing.Point(2, 138);
            this.Label30.TabIndex = 10;
            //
            // lblPlatform
            //
            this.lblPlatform.Name = "lblPlatform";
            this.lblPlatform.Text = "Visual Basic (SP3) - Win95";
            this.lblPlatform.Font = new System.Drawing.Font("Arial",15F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblPlatform.Size = new System.Drawing.Size(264, 24);
            this.lblPlatform.Location = new System.Drawing.Point(36, 62);
            this.lblPlatform.TabIndex = 5;
            //
            // Label2
            //
            this.Label2.Name = "Label2";
            this.Label2.Text = "Use of Slice and Dice is governed by the Slice and Dice end-user agreement. Click here to view it.";
            this.Label2.Font = new System.Drawing.Font("MS Sans Serif",8F, ( System.Drawing.FontStyle.Underline ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.Label2.ForeColor = System.Drawing.Color.Transparent;
            this.Label2.Size = new System.Drawing.Size(274, 31);
            this.Label2.Location = new System.Drawing.Point(90, 244);
//            this.Label2.MouseIcon = "frmSplash.frx":0FDC;
//            this.Label2.MousePointer = 99;
            this.Label2.TabIndex = 9;
//            this.Label2.ToolTipText = "View end-user agreement.";
            //
            // Label10
            //
            this.Label10.Name = "Label10";
            this.Label10.Text = "http://www.sliceanddice.com/dl.html";
            this.Label10.Font = new System.Drawing.Font("Arial",8F, ( System.Drawing.FontStyle.Underline ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.Label10.ForeColor = System.Drawing.Color.Transparent;
            this.Label10.Size = new System.Drawing.Size(176, 14);
            this.Label10.Location = new System.Drawing.Point(92, 154);
//            this.Label10.MouseIcon = "frmSplash.frx":141E;
//            this.Label10.MousePointer = 99;
            this.Label10.TabIndex = 8;
//            this.Label10.ToolTipText = "Visit the main S&D Web Site";
            //
            // lblCopyright
            //
            this.lblCopyright.Name = "lblCopyright";
            this.lblCopyright.Text = "William Rawls retains rights to everything before 4/1/2000. All public domain thereafter.";
            this.lblCopyright.Font = new System.Drawing.Font("Arial",8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblCopyright.Size = new System.Drawing.Size(566, 12);
            this.lblCopyright.Location = new System.Drawing.Point(4, 100);
            this.lblCopyright.TabIndex = 3;
//            this.lblCopyright.ToolTipText = "Slice and Dice is copyright 1999 by Firm Solutions and William M. Rawls. Slice and Dice is a trademark of Firm Solutions.";
            //
            // lblWarning
            //
            this.lblWarning.Name = "lblWarning";
            this.lblWarning.Text = "Warning: You are responsible for any damages caused by this program. Use with care.";
            this.lblWarning.Font = new System.Drawing.Font("Arial",8F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblWarning.Size = new System.Drawing.Size(482, 12);
            this.lblWarning.Location = new System.Drawing.Point(22, 278);
            this.lblWarning.TabIndex = 2;
            //
            // lblVersion
            //
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Text = "Version";
            this.lblVersion.Font = new System.Drawing.Font("Arial",12F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblVersion.Size = new System.Drawing.Size(59, 19);
            this.lblVersion.Location = new System.Drawing.Point(50, 84);
            this.lblVersion.TabIndex = 4;
            //
            // lblLicenseTo
            //
            this.lblLicenseTo.Name = "lblLicenseTo";
            this.lblLicenseTo.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.lblLicenseTo.Text = "Licensed to: William M. Rawls, Super Human Programmer";
            this.lblLicenseTo.Font = new System.Drawing.Font("Arial",9F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblLicenseTo.Size = new System.Drawing.Size(389, 17);
            this.lblLicenseTo.Location = new System.Drawing.Point(168, 114);
            this.lblLicenseTo.TabIndex = 1;
//            this.lblLicenseTo.ToolTipText = "This is the name of the person who owns this copy of S&D";
            //
            // lblCompanyProduct
            //
            this.lblCompanyProduct.Name = "lblCompanyProduct";
            this.lblCompanyProduct.Text = "all freeware, all opensource";
            this.lblCompanyProduct.Font = new System.Drawing.Font("Times New Roman",18F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblCompanyProduct.Size = new System.Drawing.Size(283, 26);
            this.lblCompanyProduct.Location = new System.Drawing.Point(6, 1);
            this.lblCompanyProduct.TabIndex = 6;
            //
            // lblProductName
            //
            this.lblProductName.Name = "lblProductName";
            this.lblProductName.Text = "Slice and Dice";
            this.lblProductName.Font = new System.Drawing.Font("Arial",32F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblProductName.Size = new System.Drawing.Size(424, 50);
            this.lblProductName.Location = new System.Drawing.Point(13, 22);
            this.lblProductName.TabIndex = 7;
            //
            // Shape1
            //
            this.Shape1.Name = "Shape1";
            this.Shape1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Shape1.Size = new System.Drawing.Size(395, 143);
            this.Shape1.Location = new System.Drawing.Point(0, 134);
            //
            // frmSplash
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.Frame1
            });
            this.Name = "frmSplash";
            this.Frame1.ResumeLayout(false);
            this.ResumeLayout(false);
        }
        #endregion

        public string sLicenseInfo;
        public Date RightNow;


            public void Form_KeyPress            {
                Close();
            }

            public void Form_Load()
            {
                try
{;
                DetermineRegistration;
                Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2;

                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Frame1_Click()
            {
                Close();
            }

            public void imgLogo_Click()
            {
                BrowseTo "http://www.sliceanddice.com";
            }

            public void Label1_Click            {
                BrowseTo string.Empty + Label1(Index).Text;
            }

            public void Label2_Click()
            {
                BrowseTo "http://www.sliceanddice.com/agreement.html";
            }

            public void lblCompanyProduct_Click()
            {
                Frame1_Click;
            }

            public void lblCopyright_Click()
            {
                Frame1_Click;
            }

            public void lblLicenseTo_Click()
            {
                Frame1_Click;
            }

            public void lblPlatform_Click()
            {
                Frame1_Click;
            }

            public void lblProductName_Click()
            {
                Frame1_Click;
            }

            public void lblVersion_Click()
            {
                Frame1_Click;
            }

            public void lblWarning_Click()
            {
                Frame1_Click;
            }

        }
    }
