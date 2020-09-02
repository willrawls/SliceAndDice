using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class frmBrowser : System.Windows.Forms.Form
        {
         public System.Windows.Forms.MSComctlLib.Toolbar tbToolBar;
         public System.Windows.Forms.SHDocVwCtl.WebBrowser brwWebBrowser;
         public System.Windows.Forms.VB.Timer timTimer;
         public System.Windows.Forms.VB.PictureBox picAddress;
         public System.Windows.Forms.VB.ComboBox cboAddress;
         public System.Windows.Forms.VB.Label lblAddress;
         public System.Windows.Forms.MSComctlLib.ImageList imlIcons;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.Container components = null;

        public frmBrowser()
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
            this.tbToolBar = new System.Windows.Forms.MSComctlLib.Toolbar();
            this.brwWebBrowser = new System.Windows.Forms.SHDocVwCtl.WebBrowser();
            this.timTimer = new System.Windows.Forms.VB.Timer();
            this.picAddress = new System.Windows.Forms.VB.PictureBox();
            this.cboAddress = new System.Windows.Forms.VB.ComboBox();
            this.lblAddress = new System.Windows.Forms.VB.Label();
            this.imlIcons = new System.Windows.Forms.MSComctlLib.ImageList();
            this.SuspendLayout();
            this.picAddress.SuspendLayout();
            //
            // tbToolBar
            //
            this.tbToolBar.Name = "tbToolBar";
//            this.tbToolBar.Align = 1;
            this.tbToolBar.Size = new System.Drawing.Size(436, 36);
            this.tbToolBar.Location = new System.Drawing.Point(0, 0);
            this.tbToolBar.TabIndex = 3;
//            this.tbToolBar.ButtonWidth = 820;
//            this.tbToolBar.ButtonHeight = 794;
//            this.tbToolBar.ImageList = "imlIcons";
//            this.tbToolBar.Button1 = ;
//            this.tbToolBar.Button2 = ;
//            this.tbToolBar.Button3 = ;
//            this.tbToolBar.Button4 = ;
//            this.tbToolBar.Button5 = ;
//            this.tbToolBar.Button6 = ;
//            this.tbToolBar.Button6 = ;
            //
            // brwWebBrowser
            //
            this.brwWebBrowser.Name = "brwWebBrowser";
            this.brwWebBrowser.Size = new System.Drawing.Size(359, 248);
            this.brwWebBrowser.Location = new System.Drawing.Point(3, 81);
            this.brwWebBrowser.TabIndex = 0;
//            this.brwWebBrowser.ExtentX = 9525;
//            this.brwWebBrowser.ExtentY = 6588;
//            this.brwWebBrowser.ViewMode = 1;
//            this.brwWebBrowser.Offline = 0;
//            this.brwWebBrowser.Silent = 0;
//            this.brwWebBrowser.RegisterAsBrowse = 0;
//            this.brwWebBrowser.RegisterAsDropTarge = 0;
//            this.brwWebBrowser.AutoArrange = -1;
//            this.brwWebBrowser.NoClientEdge = -1;
//            this.brwWebBrowser.AlignLeft = 0;
//            this.brwWebBrowser.ViewID = "{0057D0E0-3573-11CF-AE69-08002B2E1262}";
//            this.brwWebBrowser.Location = "res://E:\WINNT\system32\shdoclc.dll/dnserror.htm#http:///";
            //
            // timTimer
            //
            this.timTimer.Name = "timTimer";
            this.timTimer.Enabled = false;
            this.timTimer.Interval = 5;
            this.timTimer.Location = new System.Drawing.Point(412, 100);
            //
            // picAddress
            //
            this.picAddress.Name = "picAddress";
//            this.picAddress.Align = 1;
            this.picAddress.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.picAddress.Size = new System.Drawing.Size(436, 45);
            this.picAddress.Location = new System.Drawing.Point(0, 36);
            this.picAddress.TabIndex = 4;
            this.picAddress.TabStop = false;
            this.picAddress.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                        this.cboAddress,
                        this.lblAddress
            });
            //
            // cboAddress
            //
            this.cboAddress.Name = "cboAddress";
            this.cboAddress.Size = new System.Drawing.Size(253, 21);
            this.cboAddress.Location = new System.Drawing.Point(3, 20);
            this.cboAddress.TabIndex = 2;
            this.cboAddress.Text = "��END!";
            //
            // lblAddress
            //
            this.lblAddress.Name = "lblAddress";
            this.lblAddress.Text = "&Address:";
            this.lblAddress.Size = new System.Drawing.Size(205, 17);
            this.lblAddress.Location = new System.Drawing.Point(3, 4);
            this.lblAddress.TabIndex = 1;
            this.lblAddress.Tag = "&Address:";
            //
            // imlIcons
            //
            this.imlIcons.Name = "imlIcons";
            this.imlIcons.Location = new System.Drawing.Point(178, 155);
            this.imlIcons.BackColor = System.Drawing.Color.FromArgb(-2147483643);
//            this.imlIcons.ImageWidth = 24;
//            this.imlIcons.ImageHeight = 24;
//            this.imlIcons.MaskColor = 12632256;
//            this.imlIcons.ListImage1 = ;
//            this.imlIcons.ListImage2 = ;
//            this.imlIcons.ListImage3 = ;
//            this.imlIcons.ListImage4 = ;
//            this.imlIcons.ListImage5 = ;
//            this.imlIcons.ListImage6 = ;
//            this.imlIcons.ListImage6 = ;
            //
            // frmBrowser
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.tbToolBar,
                  this.brwWebBrowser,
                  this.timTimer,
                  this.picAddress,
                  this.imlIcons
            });
            this.Name = "frmBrowser";
            this.picAddress.ResumeLayout(false);
            this.ResumeLayout(false);
        }
        #endregion

        public As StartingAddress;
        public As DontNavigateNow;
        public object Parent;
        public object Sites;


            public void Form_Initialize()
            {
                LoadSites;
            }

            public void Form_Load()
            {
                try
{;
                Show;
                tbToolBar.Refresh;
                Form_Resize;

                cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15;

                GotoSite 1;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void brwWebBrowser_DownloadComplete()
            {
                try
{;
                Caption = brwWebBrowser.LocationName;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void brwWebBrowser_NavigateComplete            {
                ;
                ;

                Caption = brwWebBrowser.LocationName;
                for(var CurrAddress = 0; CurrAddress < cboAddress.ListCount - 1; CurrAddress++)  {;
                if ( cboAddress.List(CurrAddress) = brwWebBrowser.LocationURL )
            {;
                bFound = true;
                Exit For;
                };
                } // CurrAddress;

                DontNavigateNow = true;
                if ( bFound )
            {
 cboAddress.RemoveItem CurrAddress;
                cboAddress.AddItem(brwWebBrowser.LocationURL, 0);
                cboAddress.ListIndex = 0;
                DontNavigateNow = false;
            }

            public void cboAddress_Click()
            {
                if ( DontNavigateNow )
            {
 return;
                timTimer.Enabled = true;
                brwWebBrowser.Navigate cboAddress.Text;
            }

            public void cboAddress_KeyPress            {
                try
{;
                if ( KeyAscii == vbKeyReturn )
            {
 cboAddress_Click;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Form_Resize()
            {
                cboAddress.Width = ScaleWidth - 100;
                brwWebBrowser.Width = ScaleWidth - 100;
                brwWebBrowser.Height = ScaleHeight - (picAddress.Top + picAddress.Height) - 100;
            }

            public void Form_Terminate()
            {
                SaveSites;
                Sites.this.Clear()false;
                Sites = null;
            }

            public void timTimer_Timer()
            {
                if ( brwWebBrowser.Busy = false )
            {;
                timTimer.Enabled = false;
                Caption = brwWebBrowser.LocationName;
                }
            else
            {;
                Caption = "Working...";
                };
            }

            public void tbToolBar_ButtonClick            {
                try
{;
                timTimer.Enabled = true;
                switch Button.Key;
                Case "Back":            brwWebBrowser.GoBack;
                Case "Forward":         brwWebBrowser.GoForward;
                Case "Refresh":         brwWebBrowser.Refresh;
                Case "Home":            brwWebBrowser.GoHome;
                Case "Search":          brwWebBrowser.GoSearch;
                Case "Stop":            timTimer.Enabled = false;
                brwWebBrowser.Stop;
                Caption = brwWebBrowser.LocationName;
                };
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void GotoSite            {
                StartingAddress = sGetToken(Sites("1"), "|||");

                if ( Len(StartingAddress) > 0 )
            {;
                cboAddress.Text = StartingAddress;
                cboAddress.AddItem(cboAddress.Text);

                //      ' Try to navigate to the starting address;
                timTimer.Enabled = true;
                brwWebBrowser.Navigate StartingAddress;
                };
            }

            public void LoadSites()
            {
                if ( ! Sites Is null )
            {;
                Sites.Clear;
                Sites = null;
                };

                Sites = new CAssocArray();


                Sites.FieldDelimiter = "|||";
                Sites.KeyValueDelimiter = "=";
                Sites.ItemDelimiter = vbNewLine;

                Sites.All = GetSetting("Slice And Dice", "Sandal", "AutoWeb Sites", sDefaults);

                if ( Sites.Count = 0 )
            {;
                Sites.All = "0=Add(                                                                                                   a new site|||***ADD SITE***" + vbNewLine +)
                   "1=http://www.vbcode.com|||VBCode.com||Desc of VBCode.com|||" + vbNewLine +
                   "2=http://www.planet-source-code.com|||Planet Source Code|||Desc of Planet Source Code|||" + vbNewLine;
                };

            }

            public void SaveSites()
            {
                if ( ! Sites Is null )
            {;
                SaveSetting "Slice And Dice", "Sandal", "AutoWeb Sites", Sites.All;
                };
            }

        }
    }
