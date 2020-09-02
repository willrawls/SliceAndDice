using System;
using System.Drawing;
using System.Dictionary<string,string>()s;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
		public class frmODBCFieldType : System.Windows.Forms.Form
		{
		 public System.Windows.Forms.VB.TextBox FieldName;
		 public System.Windows.Forms.VB.CommandButton cmdCancel;
		 public System.Windows.Forms.VB.CommandButton cmdOkay;
		 public System.Windows.Forms.VB.TextBox txtLength;
		 public System.Windows.Forms.VB.ListBox lstType;
		 public System.Windows.Forms.VB.Label Label3;
		 public System.Windows.Forms.VB.Label Label2;
		 public System.Windows.Forms.VB.Label Label1;
			/// <summary>
			/// Required designer variable.
			/// </summary>
			public System.ComponentModel.Container components = null;

			public frmODBCFieldType()
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
				this.FieldName = new System.Windows.Forms.VB.TextBox();
				this.cmdCancel = new System.Windows.Forms.VB.CommandButton();
				this.cmdOkay = new System.Windows.Forms.VB.CommandButton();
				this.txtLength = new System.Windows.Forms.VB.TextBox();
				this.lstType = new System.Windows.Forms.VB.ListBox();
				this.Label3 = new System.Windows.Forms.VB.Label();
				this.Label2 = new System.Windows.Forms.VB.Label();
				this.Label1 = new System.Windows.Forms.VB.Label();
				this.SuspendLayout();
				//
				// FieldName
				//
				this.FieldName.Name = "FieldName";
				this.FieldName.Size = new System.Drawing.Size(171, 20);
				this.FieldName.Location = new System.Drawing.Point(4, 18);
				this.FieldName.TabIndex = 0;
				//
				// cmdCancel
				//
				this.cmdCancel.Name = "cmdCancel";
				this.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
				this.cmdCancel.Text = "&Cancel";
				this.cmdCancel.Size = new System.Drawing.Size(57, 29);
				this.cmdCancel.Location = new System.Drawing.Point(118, 96);
				this.cmdCancel.TabIndex = 4;
				//
				// cmdOkay
				//
				this.cmdOkay.Name = "cmdOkay";
				this.cmdOkay.Text = "&OK";
				this.cmdOkay.DialogResult = System.Windows.Forms.DialogResult.OK;
				this.cmdOkay.Size = new System.Drawing.Size(57, 29);
				this.cmdOkay.Location = new System.Drawing.Point(118, 60);
				this.cmdOkay.TabIndex = 3;
				//
				// txtLength
				//
				this.txtLength.Name = "txtLength";
				this.txtLength.TextAlign = System.Drawing.ContentAlignment.TopRight;
				this.txtLength.Size = new System.Drawing.Size(107, 20);
				this.txtLength.Location = new System.Drawing.Point(4, 222);
//				this.txtLength.MultiLine = -1;
				this.txtLength.TabIndex = 2;
				this.txtLength.Text = "frmODBCFieldType.frx":0000;
				//
				// lstType
				//
				this.lstType.Name = "lstType";
				this.lstType.Size = new System.Drawing.Size(107, 138);
//				this.lstType.IntegralHeight = 0;
//				this.lstType.ItemData = "frmODBCFieldType.frx":0004;
				this.lstType.Location = new System.Drawing.Point(4, 60);
//				this.lstType.List = "frmODBCFieldType.frx":0026;
				this.lstType.TabIndex = 1;
				//
				// Label3
				//
				this.Label3.Name = "Label3";
				this.Label3.Text = "Field Name:";
				this.Label3.Size = new System.Drawing.Size(56, 13);
				this.Label3.Location = new System.Drawing.Point(4, 0);
				this.Label3.TabIndex = 7;
				//
				// Label2
				//
				this.Label2.Name = "Label2";
				this.Label2.Text = "Field Length:";
				this.Label2.Size = new System.Drawing.Size(61, 13);
				this.Label2.Location = new System.Drawing.Point(4, 204);
				this.Label2.TabIndex = 6;
				//
				// Label1
				//
				this.Label1.Name = "Label1";
				this.Label1.Text = "Field Type:";
				this.Label1.Size = new System.Drawing.Size(52, 13);
				this.Label1.Location = new System.Drawing.Point(4, 44);
				this.Label1.TabIndex = 5;
				//
				// frmODBCFieldType
				//
				this.Controls.AddRange(new System.Windows.Forms.Control[]
				{
				      this.FieldName,
				      this.cmdCancel,
				      this.cmdOkay,
				      this.txtLength,
				      this.lstType,
				      this.Label3,
				      this.Label2,
				      this.Label1
				});
				this.Name = "frmODBCFieldType";
				this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
				this.Text = "Set Field Name, type, size";
				this.ClientSize = new System.Drawing.Size(178, 247);
				this.ControlBox = false;
				this.MaximizeBox = false;
				this.MinimizeBox = false;
				this.ShowInTaskbar = false;
				this.ResumeLayout(false);
			}
			#endregion

			public bool m_bCanceled;

			public bool Canceled
			{
				get
				{
					Canceled = m_bCanceled;
				}

			}

			public DataTypeEnum dbFieldType
			{
				get
				{
					switch lstType;
					Case "Text":          dbFieldType = dbText;
					Case "Long Integer":  dbFieldType = dbLong;
					Case "Boolean":       dbFieldType = dbBoolean;
					Case "Currency":      dbFieldType = dbCurrency;
					Case "Date/Time":     dbFieldType = dbDate;
					Case "Double":        dbFieldType = dbDouble;
					Case "Integer":       dbFieldType = dbInteger;
					Case "Binary":        dbFieldType = dbLongBinary;
					Case "Memo":          dbFieldType = dbMemo;
					Case "Single":        dbFieldType = dbSingle;
					}
				}

			}

			public int Length
			{
				get
				{
					if ( lstType = "Text" )
				{
					Length = Val(txtLength);
					}
				else
				{
					Length = 0;
					}
				}

			}


			public void cmdCancel_Click()
			{
				m_bCanceled = true;
				Hide;
			}
			public void cmdOkay_Click()
			{
				FieldName = Trim(FieldName);

				if ( Len(FieldName) = 0 )
				{;
				MsgBox "Please enter a name before pressing 'OK'", vbInformation;
				return;
				}
				else
				{if ( lstType.ListIndex < 0 )
				{;
				MsgBox "Please select a field type before pressing 'OK'", vbInformation;
				return;
				};

				m_bCanceled = false;
				Hide;
			}
			public void Form_Activate()
			{
				try
{;
				FieldName.SetFocus;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void Form_Load()
			{
				try
{;
				if ( lstType.ListIndex = -1 )
				{;
				lstType.ListIndex = 0;
				};
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void lstType_Click()
			{
				txtLength.Enabled = (lstType = "Text");
			}
			public void lstType_DblClick()
			{
				cmdOkay_Click;
			}
			public void txtLength_Change()
			{
				if ( CStr(Val(txtLength)) <> txtLength )
				{;
				if ( Val(txtLength) < 1 )
				{;
				txtLength = "1";
				}
				else
				{if ( Val(txtLength) > 255 )
				{;
				txtLength = "255";
				}
				else
				{;
				txtLength = CStr(Val(txtLength));
				};
				};
			}
		}
}
