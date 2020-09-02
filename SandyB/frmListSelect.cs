using System;
using System.Drawing;
using System.Dictionary<string,string>()s;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
		public class frmListSelect : System.Windows.Forms.Form
		{
		 public System.Windows.Forms.MSComctlLib.ListView lstChoose;
		 public System.Windows.Forms.VB.CommandButton cmdCancel;
		 public System.Windows.Forms.VB.CommandButton cmdOkay;
		 public System.Windows.Forms.MSComctlLib.ImageList imlSmallIcons;
			/// <summary>
			/// Required designer variable.
			/// </summary>
			public System.ComponentModel.Container components = null;

			public frmListSelect()
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
				this.lstChoose = new System.Windows.Forms.MSComctlLib.ListView();
				this.cmdCancel = new System.Windows.Forms.VB.CommandButton();
				this.cmdOkay = new System.Windows.Forms.VB.CommandButton();
				this.imlSmallIcons = new System.Windows.Forms.MSComctlLib.ImageList();
				this.SuspendLayout();
				//
				// lstChoose
				//
				this.lstChoose.Name = "lstChoose";
				this.lstChoose.Size = new System.Drawing.Size(475, 341);
				this.lstChoose.Location = new System.Drawing.Point(0, 0);
				this.lstChoose.TabIndex = 0;
				this.lstChoose.View = System.Windows.Forms.View.List;
//				this.lstChoose.Arrange = 1;
				this.lstChoose.LabelEdit = true;
				this.lstChoose.LabelWrap = false;
				this.lstChoose.HideSelection = true;
//				this.lstChoose.HideColumnHeader = -1;
//				this.lstChoose.FullRowSelect = -1;
//				this.lstChoose.GridLines = -1;
//				this.lstChoose.Icons = "imlSmallIcons";
//				this.lstChoose.SmallIcons = "imlSmallIcons";
//				this.lstChoose.ColHdrIcons = "imlSmallIcons";
				this.lstChoose.ForeColor = System.Drawing.Color.Black;
				this.lstChoose.BackColor = System.Drawing.Color.FromArgb(-2147483624);
//				this.lstChoose.NumItems = 0;
				//
				// cmdCancel
				//
				this.cmdCancel.Name = "cmdCancel";
				this.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
				this.cmdCancel.Text = "&Cancel";
				this.cmdCancel.Size = new System.Drawing.Size(73, 33);
				this.cmdCancel.Location = new System.Drawing.Point(96, 342);
				this.cmdCancel.TabIndex = 2;
				//
				// cmdOkay
				//
				this.cmdOkay.Name = "cmdOkay";
				this.cmdOkay.Text = "&OK";
				this.cmdOkay.DialogResult = System.Windows.Forms.DialogResult.OK;
				this.cmdOkay.Size = new System.Drawing.Size(73, 33);
				this.cmdOkay.Location = new System.Drawing.Point(0, 342);
				this.cmdOkay.TabIndex = 1;
				//
				// imlSmallIcons
				//
				this.imlSmallIcons.Name = "imlSmallIcons";
				this.imlSmallIcons.Location = new System.Drawing.Point(0, 0);
				this.imlSmallIcons.BackColor = System.Drawing.Color.FromArgb(-2147483643);
//				this.imlSmallIcons.ImageWidth = 16;
//				this.imlSmallIcons.ImageHeight = 16;
//				this.imlSmallIcons.MaskColor = 12632256;
//				this.imlSmallIcons.ListImage1 = ;
//				this.imlSmallIcons.ListImage2 = ;
//				this.imlSmallIcons.ListImage3 = ;
//				this.imlSmallIcons.ListImage4 = ;
//				this.imlSmallIcons.ListImage5 = ;
//				this.imlSmallIcons.ListImage6 = ;
//				this.imlSmallIcons.ListImage7 = ;
//				this.imlSmallIcons.ListImage8 = ;
//				this.imlSmallIcons.ListImage9 = ;
//				this.imlSmallIcons.ListImage10 = ;
//				this.imlSmallIcons.ListImage11 = ;
//				this.imlSmallIcons.ListImage12 = ;
//				this.imlSmallIcons.ListImage13 = ;
//				this.imlSmallIcons.ListImage14 = ;
//				this.imlSmallIcons.ListImage15 = ;
//				this.imlSmallIcons.ListImage16 = ;
//				this.imlSmallIcons.ListImage17 = ;
//				this.imlSmallIcons.ListImage18 = ;
//				this.imlSmallIcons.ListImage18 = ;
				//
				// frmListSelect
				//
				this.Controls.AddRange(new System.Windows.Forms.Control[]
				{
				      this.lstChoose,
				      this.cmdCancel,
				      this.cmdOkay,
				      this.imlSmallIcons
				});
				this.Name = "frmListSelect";
				this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
				this.Text = "Select one";
				this.ClientSize = new System.Drawing.Size(474, 374);
				this.ControlBox = false;
				this.MaximizeBox = false;
				this.MinimizeBox = false;
				this.ShowInTaskbar = false;
				this.ResumeLayout(false);
			}
			#endregion

			public New Choices;
			public string m_sChoice;
			public string Key;

			public string Choice
			{
				get
				{
					Choice = m_sChoice;
				}

			}


			public void cmdCancel_Click()
			{
				m_sChoice = string.Empty;
				Hide;
			}
			public void cmdOkay_Click()
			{
				if ( lstChoose.SelectedItem Is null )
				{
;
				// If lstChoose.ListIndex < 0 Then;
				MsgBox "Please choose one before pressing OK.", vbInformation;
				return;
			}
			public void Initialize			{
				try
{;

				Choices.ItemDelimiter = sDelimiter;
				Choices.All = sChoices;
				Key = Left$(sChoices, 25);
				if ( Len(Key) > 0 )
				{;
				LoadFormPosition Me, , , Key;
				}
				else
				{;
				LoadFormPosition Me;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void Form_Load()
			{

				lstChoose.ColumnHeaders.Clear;
				lstChoose.ColumnHeaders.Add , , "Title", 2800;
				lstChoose.ColumnHeaders.Add , , "Description", 9000;
			}
			public void Form_Resize()
			{
				cmdOkay.Move 0, ScaleHeight - cmdOkay.Height;
				cmdCancel.Move ScaleWidth - cmdCancel.Width, cmdOkay.Top;
				lstChoose.Move 0, 0, ScaleWidth, ScaleHeight - cmdOkay.Height;
			}
			public void Form_Unload			{
				if ( Len(Key) )
				{;
				SaveFormPosition Me, Key;
				}
				else
				{;
				SaveFormPosition Me;
			}
			public void lstChoose_DblClick()
			{
				cmdOkay_Click;
			}
		}
}
