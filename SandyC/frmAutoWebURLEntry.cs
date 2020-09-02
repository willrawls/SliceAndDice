using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class frmURLEntry : System.Windows.Forms.Form
        {
         public System.Windows.Forms.VB.TextBox txtDataPaste;
         public System.Windows.Forms.VB.CommandButton cmdRemove;
         public System.Windows.Forms.VB.ListBox lstDataToPost;
         public System.Windows.Forms.VB.CommandButton cmdAdd;
         public System.Windows.Forms.VB.CommandButton cmdCancel;
         public System.Windows.Forms.VB.CommandButton cmdOkay;
         public System.Windows.Forms.VB.TextBox txtObjectToActivate;
         public System.Windows.Forms.VB.TextBox txtAltURL;
         public System.Windows.Forms.VB.TextBox txtURL;
         public System.Windows.Forms.VB.Label lblDataToPost;
         public System.Windows.Forms.VB.Label lblObjectToActivate;
         public System.Windows.Forms.VB.Label lblAltURL;
         public System.Windows.Forms.VB.Label lblURL;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.Container components = null;

        public frmURLEntry()
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
            this.txtDataPaste = new System.Windows.Forms.VB.TextBox();
            this.cmdRemove = new System.Windows.Forms.VB.CommandButton();
            this.lstDataToPost = new System.Windows.Forms.VB.ListBox();
            this.cmdAdd = new System.Windows.Forms.VB.CommandButton();
            this.cmdCancel = new System.Windows.Forms.VB.CommandButton();
            this.cmdOkay = new System.Windows.Forms.VB.CommandButton();
            this.txtObjectToActivate = new System.Windows.Forms.VB.TextBox();
            this.txtAltURL = new System.Windows.Forms.VB.TextBox();
            this.txtURL = new System.Windows.Forms.VB.TextBox();
            this.lblDataToPost = new System.Windows.Forms.VB.Label();
            this.lblObjectToActivate = new System.Windows.Forms.VB.Label();
            this.lblAltURL = new System.Windows.Forms.VB.Label();
            this.lblURL = new System.Windows.Forms.VB.Label();
            this.SuspendLayout();
            //
            // txtDataPaste
            //
            this.txtDataPaste.Name = "txtDataPaste";
            this.txtDataPaste.Size = new System.Drawing.Size(366, 20);
            this.txtDataPaste.Location = new System.Drawing.Point(139, 227);
            this.txtDataPaste.TabIndex = 12;
            //
            // cmdRemove
            //
            this.cmdRemove.Name = "cmdRemove";
            this.cmdRemove.Text = "-";
            this.cmdRemove.Font = new System.Drawing.Font("MS Sans Serif",12F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.cmdRemove.Size = new System.Drawing.Size(19, 60);
            this.cmdRemove.Location = new System.Drawing.Point(121, 168);
            this.cmdRemove.TabIndex = 11;
//            this.cmdRemove.ToolTipText = "Remove the currently selected entry from the list";
            //
            // lstDataToPost
            //
            this.lstDataToPost.Name = "lstDataToPost";
            this.lstDataToPost.Size = new System.Drawing.Size(366, 121);
//            this.lstDataToPost.IntegralHeight = 0;
//            this.lstDataToPost.ItemData = "frmAutoWebURLEntry.frx":0000;
            this.lstDataToPost.Location = new System.Drawing.Point(139, 106);
//            this.lstDataToPost.List = "frmAutoWebURLEntry.frx":0002;
            this.lstDataToPost.TabIndex = 10;
            //
            // cmdAdd
            //
            this.cmdAdd.Name = "cmdAdd";
            this.cmdAdd.Text = "+";
            this.cmdAdd.Font = new System.Drawing.Font("MS Sans Serif",12F, ( System.Drawing.FontStyle.Bold ), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.cmdAdd.Size = new System.Drawing.Size(19, 60);
            this.cmdAdd.Location = new System.Drawing.Point(121, 107);
            this.cmdAdd.TabIndex = 9;
//            this.cmdAdd.ToolTipText = "Add a new entry to the list";
            //
            // cmdCancel
            //
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cmdCancel.Text = "&Cancel";
            this.cmdCancel.Size = new System.Drawing.Size(81, 33);
            this.cmdCancel.Location = new System.Drawing.Point(1, 194);
            this.cmdCancel.TabIndex = 8;
            //
            // cmdOkay
            //
            this.cmdOkay.Name = "cmdOkay";
            this.cmdOkay.Text = "&Okay";
            this.cmdOkay.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.cmdOkay.Size = new System.Drawing.Size(81, 33);
            this.cmdOkay.Location = new System.Drawing.Point(1, 158);
            this.cmdOkay.TabIndex = 7;
            //
            // txtObjectToActivate
            //
            this.txtObjectToActivate.Name = "txtObjectToActivate";
            this.txtObjectToActivate.Size = new System.Drawing.Size(366, 20);
            this.txtObjectToActivate.Location = new System.Drawing.Point(139, 73);
            this.txtObjectToActivate.TabIndex = 5;
            //
            // txtAltURL
            //
            this.txtAltURL.Name = "txtAltURL";
            this.txtAltURL.Size = new System.Drawing.Size(366, 20);
            this.txtAltURL.Location = new System.Drawing.Point(139, 40);
            this.txtAltURL.TabIndex = 3;
            //
            // txtURL
            //
            this.txtURL.Name = "txtURL";
            this.txtURL.Size = new System.Drawing.Size(366, 20);
            this.txtURL.Location = new System.Drawing.Point(139, 6);
            this.txtURL.TabIndex = 1;
            //
            // lblDataToPost
            //
            this.lblDataToPost.Name = "lblDataToPost";
            this.lblDataToPost.Text = "Data To Post";
            this.lblDataToPost.Size = new System.Drawing.Size(80, 32);
            this.lblDataToPost.Location = new System.Drawing.Point(5, 107);
            this.lblDataToPost.TabIndex = 6;
            //
            // lblObjectToActivate
            //
            this.lblObjectToActivate.Name = "lblObjectToActivate";
            this.lblObjectToActivate.Text = "Object To Activate";
            this.lblObjectToActivate.Size = new System.Drawing.Size(124, 32);
            this.lblObjectToActivate.Location = new System.Drawing.Point(5, 73);
            this.lblObjectToActivate.TabIndex = 4;
            //
            // lblAltURL
            //
            this.lblAltURL.Name = "lblAltURL";
            this.lblAltURL.Text = "Alt URL";
            this.lblAltURL.Size = new System.Drawing.Size(80, 32);
            this.lblAltURL.Location = new System.Drawing.Point(4, 39);
            this.lblAltURL.TabIndex = 2;
            //
            // lblURL
            //
            this.lblURL.Name = "lblURL";
            this.lblURL.Text = "URL";
            this.lblURL.Size = new System.Drawing.Size(80, 32);
            this.lblURL.Location = new System.Drawing.Point(5, 7);
            this.lblURL.TabIndex = 0;
            //
            // frmURLEntry
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.txtDataPaste,
                  this.cmdRemove,
                  this.lstDataToPost,
                  this.cmdAdd,
                  this.cmdCancel,
                  this.cmdOkay,
                  this.txtObjectToActivate,
                  this.txtAltURL,
                  this.txtURL,
                  this.lblDataToPost,
                  this.lblObjectToActivate,
                  this.lblAltURL,
                  this.lblURL
            });
            this.Name = "frmURLEntry";
            this.ResumeLayout(false);
        }
        #endregion

        public string msURLEntry;
        public bool Canceled;


                public string URLEntry
    {
        get
        {
        msURLEntry = txtURL + "~~~";
        msURLEntry +=  txtAltURL + "~~~";
        msURLEntry +=  ListTostring(lstDataToPost, false, "$$$") + "~~~";
        msURLEntry +=  txtObjectToActivate + "~~~";
        URLEntry = msURLEntry;
        }

        set
        {
        msURLEntry = value;
        txtURL = sGetToken(value, 1, "~~~");
        txtAltURL = sGetToken(value, 2, "~~~");
        stringToList sGetToken(value, 3, "~~~"), lstDataToPost, true, "$$$";
        txtObjectToActivate = sGetToken(value, 4, "~~~");
        }

    }



            public void cmdAdd_Click()
            {
                ;
                sNewEntry = InputBox("What should the new entry be ?" + vbCrLf + vbTab + "Name=Value", "Add(                                                                                                   DATA TO POST"));
                if ( Len(sNewEntry) )
            {;
                lstDataToPost.AddItem(sNewEntry);
                };
            }

            public void cmdCancel_Click()
            {
                Canceled = true;
                Hide;
            }

            public void cmdOkay_Click()
            {
                Canceled = false;
                Hide;
            }

            public void cmdRemove_Click()
            {
                if ( lstDataToPost.ListIndex > -1 )
            {;
                //       'If MsgBox("Are you sure you want to remove that item ?", vbYesNo, "REMOVE DATA ITEM TO POST WITH THIS URL") = vbYes Then;
                lstDataToPost.RemoveItem lstDataToPost.ListIndex;
                //       'End If;
                };
            }

            public void Form_Load()
            {

                LoadFormPosition Me;
            }

            public void Form_Resize()
            {
                try
{;
                txtURL.Width = ScaleWidth - txtURL.Left - 50;
                txtAltURL.Width = ScaleWidth - txtAltURL.Left - 50;
                txtObjectToActivate.Width = ScaleWidth - txtObjectToActivate.Left - 50;
                lstDataToPost.Width = ScaleWidth - lstDataToPost.Left - 50;
                txtDataPaste.Width = ScaleWidth - txtDataPaste.Left - 50;
                txtDataPaste.Top = ScaleHeight - txtDataPaste.Height - 50;
                lstDataToPost.Height = ScaleHeight - lstDataToPost.Top - txtDataPaste.Height - 100;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Form_Unload            {

                SaveFormPosition Me;
            }

            public void lstDataToPost_DblClick()
            {
                ;
                if ( lstDataToPost.ListIndex > -1 )
            {;
                sNewEntry = InputBox("Edit the value to post", "EDIT POST VALUE", lstDataToPost.List(lstDataToPost.ListIndex));
                if ( Len(sNewEntry) > 0 )
            {;
                lstDataToPost.List(lstDataToPost.ListIndex) = sNewEntry;
                };
                };
            }

            public void lstDataToPost_MouseUp            {
                if ( Button = vbRightButton And Shift = 0 )
            {;
                lstDataToPost_DblClick;
                };
            }

            public void txtDataPaste_Change()
            {
                if ( Len(txtDataPaste) == 0 )
            {
 return;
                if ( lTokenCount(txtDataPaste, "$$$") > 1 )
            {;
                stringToList txtDataPaste, lstDataToPost, false, "$$$";
                }
            else
            {if ( lTokenCount(txtDataPaste, "&") > 1 )
            {;
                stringToList txtDataPaste, lstDataToPost, false, "&";
                };
                txtDataPaste = string.Empty;
            }

        }
    }
