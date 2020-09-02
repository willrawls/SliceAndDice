using System;
using System.Drawing;
using System.Dictionary<string,string>()s;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
		public class frmProgress : System.Windows.Forms.Form
		{
		 public System.Windows.Forms.MSComctlLib.ProgressBar pbrProgress0;
		 public System.Windows.Forms.MSComctlLib.ProgressBar pbrProgress1;
		 public System.Windows.Forms.MSComctlLib.ProgressBar pbrProgress2;
			/// <summary>
			/// Required designer variable.
			/// </summary>
			public System.ComponentModel.Container components = null;

			public frmProgress()
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
				System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmProgress));
				this.pbrProgress0 = new System.Windows.Forms.MSComctlLib.ProgressBar();
				this.pbrProgress1 = new System.Windows.Forms.MSComctlLib.ProgressBar();
				this.pbrProgress2 = new System.Windows.Forms.MSComctlLib.ProgressBar();
				this.SuspendLayout();
				//
				// pbrProgress0
				//
				this.pbrProgress0.Name = "pbrProgress0";
//				this.pbrProgress0.Align = 1;
				this.pbrProgress0.Size = new System.Drawing.Size(402, 35);
				this.pbrProgress0.Location = new System.Drawing.Point(0, 70);
				this.pbrProgress0.TabIndex = 0;
				this.pbrProgress0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
				//
				// pbrProgress1
				//
				this.pbrProgress1.Name = "pbrProgress1";
//				this.pbrProgress1.Align = 1;
				this.pbrProgress1.Size = new System.Drawing.Size(402, 35);
				this.pbrProgress1.Location = new System.Drawing.Point(0, 35);
				this.pbrProgress1.TabIndex = 1;
				this.pbrProgress1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
				//
				// pbrProgress2
				//
				this.pbrProgress2.Name = "pbrProgress2";
//				this.pbrProgress2.Align = 1;
				this.pbrProgress2.Size = new System.Drawing.Size(402, 35);
				this.pbrProgress2.Location = new System.Drawing.Point(0, 0);
				this.pbrProgress2.TabIndex = 2;
				this.pbrProgress2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
				//
				// frmProgress
				//
				this.Controls.AddRange(new System.Windows.Forms.Control[]
				{
				      this.pbrProgress0,
				      this.pbrProgress1,
				      this.pbrProgress2
				});
				this.Name = "frmProgress";
				this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
				this.Text = "Progress Indicator";
				this.ClientSize = new System.Drawing.Size(402, 107);
				this.ControlBox = false;
				this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
				this.MaximizeBox = false;
				this.MinimizeBox = false;
				this.ShowInTaskbar = false;
				this.ResumeLayout(false);
			}
			#endregion
		}
}
