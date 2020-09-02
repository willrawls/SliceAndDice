using System;
using System.Drawing;
using System.Dictionary<string,string>()s;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
		public class frmFindReplace : System.Windows.Forms.Form
		{
		 public System.Windows.Forms.VB.CommandButton cmdReplaceAll;
		 public System.Windows.Forms.VB.CommandButton cmdReplace;
		 public System.Windows.Forms.VB.CheckBox chkUsePatternMatching;
		 public System.Windows.Forms.VB.CheckBox chkMatchCase;
		 public System.Windows.Forms.VB.CheckBox chkFindWholeWordOnly;
		 public System.Windows.Forms.VB.ComboBox cboDirection;
		 public System.Windows.Forms.VB.TextBox txtFind;
		 public System.Windows.Forms.VB.CommandButton cmdOK;
		 public System.Windows.Forms.VB.CommandButton cmdCancel;
		 public System.Windows.Forms.VB.TextBox txtReplace;
		 public System.Windows.Forms.VB.Frame Frame1;
		 public System.Windows.Forms.VB.OptionButton optSearchArea3;
		 public System.Windows.Forms.VB.OptionButton optSearchArea2;
		 public System.Windows.Forms.VB.OptionButton optSearchArea1;
		 public System.Windows.Forms.VB.OptionButton optSearchArea0;
		 public System.Windows.Forms.VB.Label lblLabels2;
		 public System.Windows.Forms.VB.Label lblLabels0;
		 public System.Windows.Forms.VB.Label lblLabels1;
			/// <summary>
			/// Required designer variable.
			/// </summary>
			public System.ComponentModel.Container components = null;

			public frmFindReplace()
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
				System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFindReplace));
				this.cmdReplaceAll = new System.Windows.Forms.VB.CommandButton();
				this.cmdReplace = new System.Windows.Forms.VB.CommandButton();
				this.chkUsePatternMatching = new System.Windows.Forms.VB.CheckBox();
				this.chkMatchCase = new System.Windows.Forms.VB.CheckBox();
				this.chkFindWholeWordOnly = new System.Windows.Forms.VB.CheckBox();
				this.cboDirection = new System.Windows.Forms.VB.ComboBox();
				this.txtFind = new System.Windows.Forms.VB.TextBox();
				this.cmdOK = new System.Windows.Forms.VB.CommandButton();
				this.cmdCancel = new System.Windows.Forms.VB.CommandButton();
				this.txtReplace = new System.Windows.Forms.VB.TextBox();
				this.Frame1 = new System.Windows.Forms.VB.Frame();
				this.optSearchArea3 = new System.Windows.Forms.VB.OptionButton();
				this.optSearchArea2 = new System.Windows.Forms.VB.OptionButton();
				this.optSearchArea1 = new System.Windows.Forms.VB.OptionButton();
				this.optSearchArea0 = new System.Windows.Forms.VB.OptionButton();
				this.lblLabels2 = new System.Windows.Forms.VB.Label();
				this.lblLabels0 = new System.Windows.Forms.VB.Label();
				this.lblLabels1 = new System.Windows.Forms.VB.Label();
				this.SuspendLayout();
				this.Frame1.SuspendLayout();
				//
				// cmdReplaceAll
				//
				this.cmdReplaceAll.Name = "cmdReplaceAll";
				this.cmdReplaceAll.Text = "Replace &All";
				this.cmdReplaceAll.Size = new System.Drawing.Size(76, 26);
				this.cmdReplaceAll.Location = new System.Drawing.Point(361, 98);
				this.cmdReplaceAll.TabIndex = 13;
				//
				// cmdReplace
				//
				this.cmdReplace.Name = "cmdReplace";
				this.cmdReplace.Text = "&Replace";
				this.cmdReplace.Size = new System.Drawing.Size(76, 26);
				this.cmdReplace.Location = new System.Drawing.Point(361, 70);
				this.cmdReplace.TabIndex = 12;
				//
				// chkUsePatternMatching
				//
				this.chkUsePatternMatching.Name = "chkUsePatternMatching";
				this.chkUsePatternMatching.Text = "&Use Pattern Matching";
				this.chkUsePatternMatching.Size = new System.Drawing.Size(138, 19);
				this.chkUsePatternMatching.Location = new System.Drawing.Point(121, 128);
				this.chkUsePatternMatching.TabIndex = 9;
				this.chkUsePatternMatching.Visible = false;
				//
				// chkMatchCase
				//
				this.chkMatchCase.Name = "chkMatchCase";
				this.chkMatchCase.Text = "Match Ca&se";
				this.chkMatchCase.Size = new System.Drawing.Size(138, 19);
				this.chkMatchCase.Location = new System.Drawing.Point(121, 110);
				this.chkMatchCase.TabIndex = 8;
				this.chkMatchCase. = Both;
				//
				// chkFindWholeWordOnly
				//
				this.chkFindWholeWordOnly.Name = "chkFindWholeWordOnly";
				this.chkFindWholeWordOnly.Text = "Find Whole World &Only";
				this.chkFindWholeWordOnly.Size = new System.Drawing.Size(138, 19);
				this.chkFindWholeWordOnly.Location = new System.Drawing.Point(121, 94);
				this.chkFindWholeWordOnly.TabIndex = 7;
				this.chkFindWholeWordOnly.Visible = false;
				//
				// cboDirection
				//
				this.cboDirection.Name = "cboDirection";
				this.cboDirection.Size = new System.Drawing.Size(57, 21);
//				this.cboDirection.ItemData = "frmFindReplace.frx":0442;
				this.cboDirection.Location = new System.Drawing.Point(182, 68);
//				this.cboDirection.List = "frmFindReplace.frx":044F;
				this.cboDirection.TabIndex = 6;
				this.cboDirection.Visible = false;
				//
				// txtFind
				//
				this.txtFind.Name = "txtFind";
				this.txtFind.Size = new System.Drawing.Size(269, 23);
				this.txtFind.Location = new System.Drawing.Point(83, 5);
				this.txtFind.TabIndex = 1;
				//
				// cmdOK
				//
				this.cmdOK.Name = "cmdOK";
				this.cmdOK.Text = "Find &Next";
				this.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK;
				this.cmdOK.Size = new System.Drawing.Size(76, 26);
				this.cmdOK.Location = new System.Drawing.Point(361, 5);
				this.cmdOK.TabIndex = 10;
				//
				// cmdCancel
				//
				this.cmdCancel.Name = "cmdCancel";
				this.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
				this.cmdCancel.Text = "Cancel";
				this.cmdCancel.Size = new System.Drawing.Size(76, 26);
				this.cmdCancel.Location = new System.Drawing.Point(361, 33);
				this.cmdCancel.TabIndex = 11;
				//
				// txtReplace
				//
				this.txtReplace.Name = "txtReplace";
				this.txtReplace.Size = new System.Drawing.Size(269, 23);
				this.txtReplace.Location = new System.Drawing.Point(83, 33);
				this.txtReplace.TabIndex = 2;
				//
				// Frame1
				//
				this.Frame1.Name = "Frame1";
				this.Frame1.Text = " Search ";
				this.Frame1.Size = new System.Drawing.Size(114, 91);
				this.Frame1.Location = new System.Drawing.Point(2, 55);
				this.Frame1.TabIndex = 16;
				this.Frame1.Controls.AddRange(new System.Windows.Forms.Control[]
				{
								this.optSearchArea3,
								this.optSearchArea2,
								this.optSearchArea1,
								this.optSearchArea0
				});
				//
				// optSearchArea3
				//
				this.optSearchArea3.Name = "optSearchArea3";
				this.optSearchArea3.Text = "Current &Database";
				this.optSearchArea3.Size = new System.Drawing.Size(106, 13);
				this.optSearchArea3.Location = new System.Drawing.Point(5, 71);
				this.optSearchArea3.TabIndex = 17;
				//
				// optSearchArea2
				//
				this.optSearchArea2.Name = "optSearchArea2";
				this.optSearchArea2.Text = "Current &Category";
				this.optSearchArea2.Size = new System.Drawing.Size(106, 13);
				this.optSearchArea2.Location = new System.Drawing.Point(5, 53);
				this.optSearchArea2.TabIndex = 5;
				//
				// optSearchArea1
				//
				this.optSearchArea1.Name = "optSearchArea1";
				this.optSearchArea1.Text = "Current &Template";
				this.optSearchArea1.Size = new System.Drawing.Size(106, 13);
				this.optSearchArea1.Location = new System.Drawing.Point(5, 35);
				this.optSearchArea1.TabIndex = 4;
				this.optSearchArea1. = Both;
				//
				// optSearchArea0
				//
				this.optSearchArea0.Name = "optSearchArea0";
				this.optSearchArea0.Text = "Current &Pane";
				this.optSearchArea0.Size = new System.Drawing.Size(85, 13);
				this.optSearchArea0.Location = new System.Drawing.Point(5, 18);
				this.optSearchArea0.TabIndex = 3;
				//
				// lblLabels2
				//
				this.lblLabels2.Name = "lblLabels2";
				this.lblLabels2.Text = "&Direction:";
				this.lblLabels2.Size = new System.Drawing.Size(45, 13);
				this.lblLabels2.Location = new System.Drawing.Point(124, 70);
				this.lblLabels2.TabIndex = 15;
				this.lblLabels2.Visible = false;
				//
				// lblLabels0
				//
				this.lblLabels0.Name = "lblLabels0";
				this.lblLabels0.Text = "&Find What:";
				this.lblLabels0.Size = new System.Drawing.Size(72, 18);
				this.lblLabels0.Location = new System.Drawing.Point(10, 9);
				this.lblLabels0.TabIndex = 0;
				//
				// lblLabels1
				//
				this.lblLabels1.Name = "lblLabels1";
				this.lblLabels1.Text = "Replace &With:";
				this.lblLabels1.Size = new System.Drawing.Size(72, 18);
				this.lblLabels1.Location = new System.Drawing.Point(10, 37);
				this.lblLabels1.TabIndex = 14;
				//
				// frmFindReplace
				//
				this.Controls.AddRange(new System.Windows.Forms.Control[]
				{
				      this.cmdReplaceAll,
				      this.cmdReplace,
				      this.chkUsePatternMatching,
				      this.chkMatchCase,
				      this.chkFindWholeWordOnly,
				      this.cboDirection,
				      this.txtFind,
				      this.cmdOK,
				      this.cmdCancel,
				      this.txtReplace,
				      this.Frame1,
				      this.lblLabels2,
				      this.lblLabels0,
				      this.lblLabels1
				});
				this.Name = "frmFindReplace";
				this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
				this.Text = "Find/Replace";
				this.ClientSize = new System.Drawing.Size(444, 149);
				this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
				this.MaximizeBox = false;
				this.MinimizeBox = false;
////				this.ScaleMode = 0;
				this.ShowInTaskbar = false;
				this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
				this.Frame1.ResumeLayout(false);
				this.ResumeLayout(false);
			}
			#endregion

			public enum FindReplaceSearchArea
			{
				SearchAreaCurrentPane = 0,
				SearchAreaCurrentTemplate = 1,
				SearchAreaCurrentCategory = 2,
				SearchAreaCurrentDatabase = 3
			};
			public enum FindReplaceDirection
			{
				DirectionAll = 0,
				DirectionDown = 1,
				DirectionUp = 2
			};

			public bool DoFindNext;
			public bool DoReplace;
			public bool DoReplaceAll;
			public bool Canceled;
			public FindReplaceSearchArea SearchArea;
			public FindReplaceDirection Direction;
			public bool FindWholeWordOnly;
			public bool MatchCase;
			public bool UsePatternMatching;

			public object IsReplace
			{
				get
				{
					IsReplace = txtReplace.Enabled;
				}

				set
				{
					lblLabels(1).Enabled = value;
					txtReplace.Visible = value;
					cmdReplaceAll.Visible = value;
					cmdReplace.Enabled = true                         'value;
					// txtReplace.Enabled = New_IsReplace;
					// cmdReplace.Enabled = New_IsReplace;
					// cmdReplaceAll.Enabled = New_IsReplace;
					Me.Show vbModal;
				}

			}


			public void cmdCancel_Click()
			{
				DoFindNext = false;
				DoReplace = false;
				DoReplaceAll = false;
				Canceled = true;
				Me.Hide;
			}
			public void cmdOK_Click()
			{
				DoFindNext = true;
				DoReplace = false;
				DoReplaceAll = false;
				Canceled = false;
				Me.Hide;
			}
			public void cmdReplace_Click()
			{
				if ( cmdReplace.Visible )
				{;
				DoFindNext = false;
				DoReplace = true;
				DoReplaceAll = false;
				Canceled = false;
				Me.Hide;
				}
				else
				{;
				DoFindNext = false;
				DoReplace = true;
				DoReplaceAll = false;
				Canceled = false;
				Me.Hide;
			}
			public void cmdReplaceAll_Click()
			{
				DoFindNext = false;
				DoReplace = false;
				DoReplaceAll = true;
				Canceled = false;
				Me.Hide;
			}
			public void Form_Activate()
			{
				if ( Len(txtFind) <> 0 And Len(txtReplace) = 0 And txtReplace.Visible )
				{;
				txtReplace.SetFocus;
				}
				else
				{;
				txtFind.SetFocus;
			}
			public void Form_Load()
			{
				txtFind = GetSetting$(App.ProductName, "Last", "Find Text", string.Empty);
				txtReplace = GetSetting$(App.ProductName, "Last", "Replace Text", string.Empty);
				optSearchArea(GetSetting(App.ProductName, "Last", "Search Area", 0)).Value = true;
				cboDirection.ListIndex = GetSetting(App.ProductName, "Last", "Search Direction", 0);
				chkFindWholeWordOnly.Value = GetSetting(App.ProductName, "Last", "Find Whole Word Only", 0);
				chkMatchCase.Value = GetSetting(App.ProductName, "Last", "Match Case", 0);
				chkUsePatternMatching.Value = GetSetting(App.ProductName, "Last", "Use Pattern Matching", 0);
			}
			public void Form_Unload			{
				SaveSetting App.ProductName, "Last", "Find Text", txtFind.Text;
				SaveSetting App.ProductName, "Last", "Replace Text", txtReplace;
				SaveSetting App.ProductName, "Last", "Search Direction", cboDirection.ListIndex;
				SaveSetting App.ProductName, "Last", "Find Whole Word Only", chkFindWholeWordOnly.Value;
				SaveSetting App.ProductName, "Last", "Match Case", chkMatchCase.Value;
				SaveSetting App.ProductName, "Last", "Use Pattern Matching", chkUsePatternMatching.Value;

				if ( optSearchArea(0).Value )
				{;
				SaveSetting App.ProductName, "Last", "Search Area", 0;
				}
				else
				{if ( optSearchArea(1).Value )
				{;
				SaveSetting App.ProductName, "Last", "Search Area", 1;
				}
				else
				{if ( optSearchArea(2).Value )
				{;
				SaveSetting App.ProductName, "Last", "Search Area", 2;
				}
				else
				{;
				SaveSetting App.ProductName, "Last", "Search Area", 3;
			}
			public void optSearchArea_Click			{
				if ( optSearchArea(0).Value )
				{;
				SearchArea = SearchAreaCurrentPane;
				}
				else
				{if ( optSearchArea(1).Value )
				{;
				SearchArea = SearchAreaCurrentTemplate;
				}
				else
				{if ( optSearchArea(2).Value )
				{;
				SearchArea = SearchAreaCurrentCategory;
				}
				else
				{;
				SearchArea = SearchAreaCurrentDatabase;
			}
		}
}
