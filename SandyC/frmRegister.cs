using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
        public class frmRegister : System.Windows.Forms.Form
        {
         public System.Windows.Forms.VB.CommandButton cmdDone;
         public System.Windows.Forms.VB.TextBox txtInvoiceNumber;
         public System.Windows.Forms.VB.CommandButton cmdStepTwo;
         public System.Windows.Forms.VB.CommandButton cmdStepOne;
         public System.Windows.Forms.InetCtlsObjects.Inet inetRegister0;
         public System.Windows.Forms.VB.Label lblInvoiceNumber;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.Container components = null;

        public frmRegister()
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
            this.cmdDone = new System.Windows.Forms.VB.CommandButton();
            this.txtInvoiceNumber = new System.Windows.Forms.VB.TextBox();
            this.cmdStepTwo = new System.Windows.Forms.VB.CommandButton();
            this.cmdStepOne = new System.Windows.Forms.VB.CommandButton();
            this.inetRegister0 = new System.Windows.Forms.InetCtlsObjects.Inet();
            this.lblInvoiceNumber = new System.Windows.Forms.VB.Label();
            this.SuspendLayout();
            //
            // cmdDone
            //
            this.cmdDone.Name = "cmdDone";
            this.cmdDone.Text = "Done";
            this.cmdDone.Size = new System.Drawing.Size(257, 35);
            this.cmdDone.Location = new System.Drawing.Point(2, 124);
            this.cmdDone.TabIndex = 3;
            //
            // txtInvoiceNumber
            //
            this.txtInvoiceNumber.Name = "txtInvoiceNumber";
            this.txtInvoiceNumber.Size = new System.Drawing.Size(147, 20);
            this.txtInvoiceNumber.Location = new System.Drawing.Point(108, 52);
            this.txtInvoiceNumber.TabIndex = 1;
//            this.txtInvoiceNumber.ToolTipText = "Enter the Invoice number given to you during step 1 here.";
            //
            // cmdStepTwo
            //
            this.cmdStepTwo.Name = "cmdStepTwo";
            this.cmdStepTwo.Text = "Step 2: Inform Central Server of Invoice Number";
            this.cmdStepTwo.Enabled = false;
            this.cmdStepTwo.Size = new System.Drawing.Size(257, 35);
            this.cmdStepTwo.Location = new System.Drawing.Point(2, 84);
            this.cmdStepTwo.TabIndex = 2;
            //
            // cmdStepOne
            //
            this.cmdStepOne.Name = "cmdStepOne";
            this.cmdStepOne.Text = "Step 1: Secure Ordering / Payment Online";
            this.cmdStepOne.Size = new System.Drawing.Size(257, 35);
            this.cmdStepOne.Location = new System.Drawing.Point(2, 2);
            this.cmdStepOne.TabIndex = 0;
            //
            // inetRegister0
            //
            this.inetRegister0.Name = "inetRegister0";
            this.inetRegister0.Location = new System.Drawing.Point(214, 10);
//            this.inetRegister0.Protocol = 4;
//            this.inetRegister0.RequestTimeout = 100;
            //
            // lblInvoiceNumber
            //
            this.lblInvoiceNumber.Name = "lblInvoiceNumber";
            this.lblInvoiceNumber.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.lblInvoiceNumber.Text = "Invoice Number from Step 1:";
            this.lblInvoiceNumber.Size = new System.Drawing.Size(83, 26);
            this.lblInvoiceNumber.Location = new System.Drawing.Point(14, 48);
            this.lblInvoiceNumber.TabIndex = 4;
            //
            // frmRegister
            //
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                  this.cmdDone,
                  this.txtInvoiceNumber,
                  this.cmdStepTwo,
                  this.cmdStepOne,
                  this.inetRegister0,
                  this.lblInvoiceNumber
            });
            this.Name = "frmRegister";
            this.ResumeLayout(false);
        }
        #endregion

        public int CurrentStage;
        public NewCommands Parent;
        public string sCategory;
        public string sTemplate;
        public string sResponse;
        public CAssocArray asaX;
        public CAssocItem CurrItem;
        public int fh;
        public byte b();
        public string strURL;
        public object InvoiceNumber;
        public As DecryptedInvoiceNumber;
        public object sEncryptedRegKey;
        public object sRegKey;
        public object sResponse;
        public object CurrValue;
        public object bOkaySoFar;
        public object asaX;
        public object fh;
        public object Value08;
        public object ProductID;
        public object NumberOfLicenses;
        public object LicensesRemaining;
        public object strOut;
        public As bytArray();
        public object CurrByte;
        public object ValueLen;
        public object OffsetLen;
        public object CharLoc;
        public object StartAt;
        public As CurrOffset;
        public object CheckSum;
        public As CheckValue;


            public void SubmitTemplate()
            {
                try
{;


                if ( Len(.CurrentTemplateNameAndCategory) > 0 )
            {;
                Parent.Parent.GetCategoryAndName Parent.Parent.CurrentTemplateNameAndCategory, sCategory, sTemplate;
                if ( ! Parent.Parent.SliceAndDice(sCategory).Templates(sTemplate) Is null )
            {;

                ;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As GetCentralUpdateInfo            {
                try
{;
                Screen.MousePointer = vbHourglass;
                sResponse = GetURL("http://www.sliceanddice.com/central.update");
                Screen.MousePointer = vbDefault;
                if ( Len(sResponse) = 0 )
            {;
                if ( bUserSure("The Central Server Update Information cannot be acceessed right now." + vbCr + vbTab + "Continue with current settings ?") )
            {;
                GetCentralUpdateInfo = true;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As GetFile            {
                try
{;

                if ( Len(sURL) == 0 Or Len(sFilename) == 0 )
            {
 return; // ???;

                if ( InStr(sFilename, "\") == 0 )
            {
 sFilename == App.Path + "\" + sFilename;

                if ( Len(Dir(sFilename)) > 0 )
            {;
                Kill sFilename;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As GetURL            {
                try
{;
                //   'Load inetRegister(0);
                inetRegister(0).RequestTimeout = 60;
                GetURL = inetRegister(0).OpenURL(sURL);
                //   'Unload inetRegister(0);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void PostURL            {
                try
{;
                //   'Load inetRegister(0);
                inetRegister(0).RequestTimeout = 60;
                inetRegister(0).Execute sURL, "POST", sData;
                //   'Unload inetRegister(0);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void cmdDone_Click()
            {
                Hide;
            }

            public void cmdStepOne_Click()
            {
                try
{;
                if ( Len(txtInvoiceNumber) > 0 )
            {;
                if ( bUserSure("It appears you have already ordered Slice and Dice because there is an Invoice number." + vbCr + vbTab + "Would you like to continue to the online ordering system ?") )
            {;
                CurrentStage = 2;
                sadSaveLicenseKey "Current Stage", 2;
                Shell "start http://www.sliceanddice.com/register.html", vbNormalFocus;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void cmdStepTwo_Click()
            {
                ;
                ;
                try
{;
                ;
                InvoiceNumber = txtInvoiceNumber.Text;
                ;
                if ( Now - CVDate(sadGetLicenseKey("Last Updated", CDbl(Now))) > 7 )
            {;
                bOkaySoFar = GetCentralUpdateInfo;
                }
            else
            {;
                bOkaySoFar = true;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As sadInvoiceDecrypt            {
                try
{;
                string = "615243516784259045218002180248620684102579462315787815795168911248961534896127811596154329617581123589402160548" Const Offsets  ;
                string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890" Const Values   ;

                if ( Len(sInvoiceNumber) == 0 )
            {
 return; // ???;

                ValueLen = Len(Values);
                OffsetLen = Len(Offsets);
                sInvoiceNumber = Replace(Replace(Replace(sInvoiceNumber), "O", "0"), ".", "O"), "-", "".ToUpper();
                sInvoiceNumber.Substring(0, 1));
                sInvoiceNumber = sInvoiceNumber, 2.Substring( Len(sInvoiceNumber) - 2);
                Values, Right$(sInvoiceNumber.Contains(1));
                sInvoiceNumber.Substring(0, Len(sInvoiceNumber) - 1);
                CurrOffset = StartAt;
                bytArray = StrConv(sInvoiceNumber, vbFromUnicode);
                strOut = "";

                for(var CurrByte = 0; CurrByte < UBound(bytArray); CurrByte++)  {;
                Values.Contains(Chr$(bytArray(CurrByte)));
                if ( CharLoc < 1 Or CharLoc > ValueLen )
            {
 return; // ???;
                CharLoc +=  Val(Offsets, CurrOffset.Substring( 1));
                if ( CharLoc < 1 )
            {
 CharLoc == ValueLen + CharLoc;
                CheckSum = (CheckSum + Asc(Values, CharLoc.Substring( 1))) Mod ValueLen;
                strOut +=  Values, CharLoc.Substring( 1);
                CurrOffset +=  1;
                if ( CurrOffset > OffsetLen )
            {
 CurrOffset == 1;
                } // CurrByte;

                strOut.Substring(0.Substring(strOut.Substring(0.Length - 1) = "S" )
            {;
                if ( CheckSum < 1 )
            {
 CheckSum == 1;
                strOut = strOut, 2.Substring( Len(strOut) - 2);
                if ( CheckSum = CheckValue )
            {;
                sadInvoiceDecrypt = strOut;
                }
            else
            {;
                sadInvoiceDecrypt = "";
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Form_Load()
            {
                try
{;
                Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2;
                CurrentStage = sadGetLicenseKey("Current Stage", 1);
                txtInvoiceNumber = sadGetLicenseKey("Invoice Number", "");
                if ( Len(txtInvoiceNumber) > 10 )
            {;
                cmdStepTwo.Enabled = true;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void txtInvoiceNumber_Change()
            {
                try
{;
                if ( Len(txtInvoiceNumber) > 15 )
            {;
                cmdStepTwo.Enabled = true;
                }
            else
            {;
                cmdStepTwo.Enabled = false;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

        }
    }
