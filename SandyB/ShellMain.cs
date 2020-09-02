using System;
using System.Drawing;
using System.Dictionary<string,string>()s;
using System.ComponentModel;
using System.Windows.Forms;

namespace MetX.SliceAndDice
{
		public class frmMain : System.Windows.Forms.Form
		{
		 public System.Windows.Forms.SysTrayCtl.cSysTray trayMain;
		 public System.Windows.Forms.VB.Menu mnuTray;
		 public System.Windows.Forms.VB.Menu mnuExit;
		 public System.Windows.Forms.VB.Menu mnuTraySep1;
		 public System.Windows.Forms.VB.Menu mnuAbout;
		 public System.Windows.Forms.VB.Menu mnuTraySep9;
		 public System.Windows.Forms.VB.Menu mnuShowExternals;
		 public System.Windows.Forms.VB.Menu mnuFavorites;
		 public System.Windows.Forms.VB.Menu mnuTraySep0;
		 public System.Windows.Forms.VB.Menu mnuMainWindow;
			/// <summary>
			/// Required designer variable.
			/// </summary>
			public System.ComponentModel.Container components = null;

			public frmMain()
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
				System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmMain));
				this.trayMain = new System.Windows.Forms.SysTrayCtl.cSysTray();
				this.mnuTray = new System.Windows.Forms.VB.Menu();
				this.mnuExit = new System.Windows.Forms.VB.Menu();
				this.mnuTraySep1 = new System.Windows.Forms.VB.Menu();
				this.mnuAbout = new System.Windows.Forms.VB.Menu();
				this.mnuTraySep9 = new System.Windows.Forms.VB.Menu();
				this.mnuShowExternals = new System.Windows.Forms.VB.Menu();
				this.mnuFavorites = new System.Windows.Forms.VB.Menu();
				this.mnuTraySep0 = new System.Windows.Forms.VB.Menu();
				this.mnuMainWindow = new System.Windows.Forms.VB.Menu();
				this.SuspendLayout();
				this.mnuTray.SuspendLayout();
				//
				// trayMain
				//
				this.trayMain.Name = "trayMain";
				this.trayMain.Location = new System.Drawing.Point(0, 0);
//				this.trayMain.InTray = -1;
//				this.trayMain.TrayIcon = "ShellMain.frx":014A;
//				this.trayMain.TrayTip = "Slice and Dice Shell";
				//
				// mnuTray
				//
				this.mnuTray.Name = "mnuTray";
				this.mnuTray.Text = "Tray";
				this.mnuTray.Visible = false;
				this.mnuTray.Controls.AddRange(new System.Windows.Forms.Control[]
				{
								this.mnuExit,
								this.mnuTraySep1,
								this.mnuAbout,
								this.mnuTraySep9,
								this.mnuShowExternals,
								this.mnuFavorites,
								this.mnuTraySep0,
								this.mnuMainWindow
				});
				//
				// mnuExit
				//
				this.mnuExit.Name = "mnuExit";
				this.mnuExit.Text = "E&xit";
				//
				// mnuTraySep1
				//
				this.mnuTraySep1.Name = "mnuTraySep1";
				this.mnuTraySep1.Text = "-";
				//
				// mnuAbout
				//
				this.mnuAbout.Name = "mnuAbout";
				this.mnuAbout.Text = "&About...";
				//
				// mnuTraySep9
				//
				this.mnuTraySep9.Name = "mnuTraySep9";
				this.mnuTraySep9.Text = "-";
				//
				// mnuShowExternals
				//
				this.mnuShowExternals.Name = "mnuShowExternals";
				this.mnuShowExternals.Text = "&Externals...";
				//
				// mnuFavorites
				//
				this.mnuFavorites.Name = "mnuFavorites";
				this.mnuFavorites.Text = "&Favorites...";
				//
				// mnuTraySep0
				//
				this.mnuTraySep0.Name = "mnuTraySep0";
				this.mnuTraySep0.Text = "-";
				//
				// mnuMainWindow
				//
				this.mnuMainWindow.Name = "mnuMainWindow";
				this.mnuMainWindow.Text = "&Main Window";
				//
				// frmMain
				//
				this.Controls.AddRange(new System.Windows.Forms.Control[]
				{
				      this.trayMain,
				      this.mnuTray
				});
				this.Name = "frmMain";
				this.Text = "Slice and Dice Shell";
				this.ClientSize = new System.Drawing.Size(197, 35);
////				this.HasDC = 0;
				this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
				this.MaximizeBox = false;
				this.MinimizeBox = false;
				this.ShowInTaskbar = false;
				this.Visible = false;
				this.mnuTray.ResumeLayout(false);
				this.ResumeLayout(false);
			}
			#endregion

			public  myIDTExtender;
			public  QueuedPopupMenu;
			public 1) CustomVars(0;
			public  sQueuedPopupMenu;
			public  IsMenuDisplayed;

			public void ShowQueuedPopupMenu()
			{
				try
{;
				MenuToShow;
				mySandyWizard;

				TryAgain:;
				if ( ! QueuedPopupMenu Is null )
				{
;
				MenuToShow = QueuedPopupMenu;
				QueuedPopupMenu = null;
				IsMenuDisplayed = true;
				PopupMenu MenuToShow, , , , mnuMainWindow;
				IsMenuDisplayed = false;
				GoTo TryAgain;
				}
				else
				{if ( Len(sQueuedPopupMenu) > 0 )
				{;
				mySandyWizard = myIDTExtender;
				switch UCase$(sQueuedPopupMenu);
				Case "FAVORITES": mySandyWizard.FavoriteCalledFromIDE = true: mySandyWizard.ShowFavoritesMenu;
				Case "EXTERNALS": mySandyWizard.ShowExternalsMenu;
				Case "ABOUT":     mySandyWizard.ShowSplashScreen;
				};
				mySandyWizard = null;
				sQueuedPopupMenu = string.Empty;
				};
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
				myIDTExtender = new SliceAndDice.Wizard();

				CustomVars(0) = "sadAddin|sadFile|sadRegister|sadSoftCodeWmr";
				CustomVars(1) = "";

				myIDTExtender.OnConnection null, vbext_cm_External, null, CustomVars;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void Form_Unload			{
				try
{;
				myIDTExtender.OnDisconnection vbext_dm_HostShutdown, CustomVars;
				myIDTExtender = null;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void mnuAbout_Click()
			{
				sQueuedPopupMenu = "About";
			}
			public void mnuExit_Click()
			{
				try
{;
				SliceAndDice.Wizard mySandyWizard;
				mySandyWizard = myIDTExtender;
				mySandyWizard.HideWindows;
				mySandyWizard = null;

				Close();
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void mnuFavorites_Click()
			{
				sQueuedPopupMenu = "Favorites";
			}
			public void mnuMainWindow_Click()
			{
				try
{;
				SliceAndDice.Wizard mySandyWizard;
				mySandyWizard = myIDTExtender;
				mySandyWizard.ShowMainWindow;
				mySandyWizard = null;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void mnuShowExternals_Click()
			{
				sQueuedPopupMenu = "Externals";
			}
			public void trayMain_MouseDblClick			{
				try
{;
				mnuMainWindow_Click;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void trayMain_MouseUp			{
				try
{;
				QueuedPopupMenu = mnuTray;
				ShowQueuedPopupMenu;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
		}
}
