using System;
using System.Drawing;
using System.Dictionary<string,string>()s;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
		public class Search2Form : System.Windows.Forms.Form
		{
		 public System.Windows.Forms.VB.CheckBox chkIncludeSubDirs;
		 public System.Windows.Forms.VB.TextBox txtFilePattern;
		 public System.Windows.Forms.VB.TextBox txtStartDir;
		 public System.Windows.Forms.VB.CommandButton cmdSearch;
		 public System.Windows.Forms.VB.Label Label11;
		 public System.Windows.Forms.VB.Label Label10;
			/// <summary>
			/// Required designer variable.
			/// </summary>
			public System.ComponentModel.Container components = null;

			public Search2Form()
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
				this.chkIncludeSubDirs = new System.Windows.Forms.VB.CheckBox();
				this.txtFilePattern = new System.Windows.Forms.VB.TextBox();
				this.txtStartDir = new System.Windows.Forms.VB.TextBox();
				this.cmdSearch = new System.Windows.Forms.VB.CommandButton();
				this.Label11 = new System.Windows.Forms.VB.Label();
				this.Label10 = new System.Windows.Forms.VB.Label();
				this.SuspendLayout();
				//
				// chkIncludeSubDirs
				//
				this.chkIncludeSubDirs.Name = "chkIncludeSubDirs";
				this.chkIncludeSubDirs.Text = "Include sub-directories";
				this.chkIncludeSubDirs.Enabled = false;
				this.chkIncludeSubDirs.Size = new System.Drawing.Size(201, 21);
				this.chkIncludeSubDirs.Location = new System.Drawing.Point(86, 60);
				this.chkIncludeSubDirs.TabIndex = 5;
				this.chkIncludeSubDirs. = Both;
				//
				// txtFilePattern
				//
				this.txtFilePattern.Name = "txtFilePattern";
				this.txtFilePattern.Size = new System.Drawing.Size(85, 19);
				this.txtFilePattern.Location = new System.Drawing.Point(88, 32);
				this.txtFilePattern.TabIndex = 1;
				this.txtFilePattern.Text = "*.*";
//				this.txtFilePattern.ToolTipText = "Enter a pattern to search for (*.* finds everything, etc.)";
				//
				// txtStartDir
				//
				this.txtStartDir.Name = "txtStartDir";
				this.txtStartDir.Size = new System.Drawing.Size(277, 19);
				this.txtStartDir.Location = new System.Drawing.Point(88, 8);
				this.txtStartDir.TabIndex = 0;
//				this.txtStartDir.ToolTipText = "Enter the drive and path to start the search at";
				//
				// cmdSearch
				//
				this.cmdSearch.Name = "cmdSearch";
				this.cmdSearch.DialogResult = System.Windows.Forms.DialogResult.Cancel;
				this.cmdSearch.Text = "&Get File List";
				this.cmdSearch.DialogResult = System.Windows.Forms.DialogResult.OK;
				this.cmdSearch.Size = new System.Drawing.Size(81, 33);
				this.cmdSearch.Location = new System.Drawing.Point(284, 32);
				this.cmdSearch.TabIndex = 2;
				//
				// Label11
				//
				this.Label11.Name = "Label11";
				this.Label11.Text = "File Pattern";
				this.Label11.Size = new System.Drawing.Size(73, 17);
				this.Label11.Location = new System.Drawing.Point(8, 32);
				this.Label11.TabIndex = 4;
				//
				// Label10
				//
				this.Label10.Name = "Label10";
				this.Label10.Text = "Start Directory";
				this.Label10.Size = new System.Drawing.Size(73, 17);
				this.Label10.Location = new System.Drawing.Point(8, 8);
				this.Label10.TabIndex = 3;
				//
				// Search2Form
				//
				this.Controls.AddRange(new System.Windows.Forms.Control[]
				{
				      this.chkIncludeSubDirs,
				      this.txtFilePattern,
				      this.txtStartDir,
				      this.cmdSearch,
				      this.Label11,
				      this.Label10
				});
				this.Name = "Search2Form";
				this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
				this.Text = "Get file list";
				this.ClientSize = new System.Drawing.Size(374, 87);
				this.MaximizeBox = false;
				this.MinimizeBox = false;
				this.ShowInTaskbar = false;
				this.ResumeLayout(false);
			}
			#endregion

			public string FileList;

			public void Form_Load()
			{
				String sFilePath;
				sFilePath = App.Path;
				if ( Right$(sFilePath, 1) <> "\" )
				{
 sFilePath = sFilePath + "\";
				txtStartDir.Text = sFilePath;
			}
			public void cmdSearch_Click()
			{
				String sFileList;
				FileList = GetFileList(txtStartDir.Text, txtFilePattern.Text);
				Hide;
			}
		}
}
