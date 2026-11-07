using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
    public class ISandyEnv
    {

        public VBIDE.VBE moVBE;
        public VBIDE.CodePane moRHS;


        public unknown sadEnvironment
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


        public unknown VBE_ActiveCodePane
        {
            get
            {
                 VBE_ActiveCodePane = moVBE.ActiveCodePane;
            }

            set
            {
            }

        }


        public unknown VBE_ActiveVBProject
        {
            get
            {
            }

            set
            {
            }

        }



            // 
            

        public VBIDE.Window VBE_ActiveWindow
        {
            get
            {
            }

        }



            // 
            

        public VBIDE.Addins VBE_Addins
        {
            get
            {
            }

        }



            // 
            

        public VBIDE.CodePanes VBE_CodePanes
        {
            get
            {
            }

        }



            // 
            

        public Office.CommandBars VBE_CommandBars
        {
            get
            {
            }

        }



            // 
            

        public unknown VBE_DisplayModel
        {
            get
            {
            }

            set
            {
            }

        }



            // 
            

        public VBIDE.Events VBE_Events
        {
            get
            {
            }

        }



            // 
            

        public string VBE_FullName
        {
            get
            {
            }

        }



            // 
            

        public unknown VBE_LastUsedPath
        {
            get
            {
            }

            set
            {
            }

        }



            // 
            

        public VBIDE.Window VBE_MainWindow
        {
            get
            {
            }

        }



            // 
            

        public string VBE_Name
        {
            get
            {
            }

        }



            // 
            

        public unknown VBE_ReadOnlyMode
        {
            get
            {
            }

            set
            {
            }

        }



            // 
            

        public VBIDE.VBComponent VBE_SelectedVBComponent
        {
            get
            {
            }

        }



            // 
            

        public string VBE_TemplatePath
        {
            get
            {
            }

        }



            // 
            

        public VBIDE.VBProjects VBE_VBProjects
        {
            get
            {
            }

        }



            // 
            

        public string VBE_Version
        {
            get
            {
            }

        }



            // 
            

        public VBIDE.Windows VBE_Windows
        {
            get
            {
            }

        }



    public void VBE_Quit()
    {
    }

    }
    }
