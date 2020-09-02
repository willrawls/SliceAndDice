using System;

namespace MetX.SliceAndDice
{
		public class ISandyEnv
		{

			public VBIDE.VBE moVBE;
			public VBIDE.CodePane moRHS;

			public object sadEnvironment
			{
				get
				{
					sadEnvironment = moVBE;
				}

				set
				{
					moVBE = value;
				}

			}

			public object VBE_ActiveCodePane
			{
				get
				{
					VBE_ActiveCodePane = moVBE.ActiveCodePane;
				}

				set
				{
					movalue = value;
				}

			}

			public object VBE_ActiveVBProject
				;

			public VBIDE.Window VBE_ActiveWindow
			{
				get; set; // Was get only
			}

			public VBIDE.Addins VBE_Addins
			{
				get; set; // Was get only
			}

			public VBIDE.CodePanes VBE_CodePanes
			{
				get; set; // Was get only
			}

			public Office.CommandBars VBE_CommandBars
			{
				get; set; // Was get only
			}

			public object VBE_DisplayModel
				;

			public VBIDE.Events VBE_Events
			{
				get; set; // Was get only
			}

			public string VBE_FullName
			{
				get; set; // Was get only
			}

			public object VBE_LastUsedPath
				;

			public VBIDE.Window VBE_MainWindow
			{
				get; set; // Was get only
			}

			public string VBE_Name
			{
				get; set; // Was get only
			}

			public object VBE_ReadOnlyMode
				;

			public VBIDE.VBComponent VBE_SelectedVBComponent
			{
				get; set; // Was get only
			}

			public string VBE_TemplatePath
			{
				get; set; // Was get only
			}

			public VBIDE.VBProjects VBE_VBProjects
			{
				get; set; // Was get only
			}

			public string VBE_Version
			{
				get; set; // Was get only
			}

			public VBIDE.Windows VBE_Windows
			{
				get; set; // Was get only
			}


			public void VBE_Quit()
			{
			}
		}
}
