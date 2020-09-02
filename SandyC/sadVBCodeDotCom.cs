using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
        public class NewCommands
        {

        public SliceAndDice.Wizard Parent;
        public SliceAndDice.CSadCommands MySadCommands;
        public SliceAndDice.CAssocArray Externals;


                public SliceAndDice.CSadCommands ISadAddin_CommandSet
    {
        get
        {
        try
{

         ISadAddin_Command = MySadCommands;
        ;
        }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        }

        }

    }


                public SliceAndDice.CAssocArray ISadAddin_Externals
    {
        get
        {
        try
{

         ISadAddin_Externals = Externals;
        ;
        }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        }

        }

    }



            public void Class_Initialize()
            {
                try
{;
                MySadCommands = new SliceAndDice.CSadCommands();
                Externals = new SliceAndDice.CAssocArray();

                MySadCommands.Clear;
                MySadCommands.ParameterDelimiter = ",";
                MySadCommands.ParameterTypeDelimiter = ":";
                MySadCommands.Attributes("Version") = App.Major + "." + App.Minor + "." + App.Revision;



                Externals.Clear;
                Externals.Item("&Import Template(s) from vbcode.com") = "Import Templates";
                Externals.Item("&Submit Current Template to vbcode.com") = "Submit Template";

                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Class_Terminate()
            {
                try
{;
                Externals = null;
                MySadCommands = null;
                frmBrowser.Parent = null;
                Unload frmBrowser;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As ISadAddin_ExecuteExternal            {
                ;
                ;

                try
{;
                if ( frmBrowser.Parent Is null )
            {;
                frmBrowser.Parent = Me;
                };
                ;
                switch sValue.ToUpper();
                Case "IMPORT TEMPLATES";
                if ( frmBrowser.Visible )
            {;
                frmBrowser.Hide;
                }
            else
            {;
                X = null;
                X = Parent.SandyWindow;
                if ( X Is null )
            {;
                frmBrowser.Show;
                }
            else
            {;
                frmBrowser.Show 0, X;
                };
                if ( ! frmBrowser.Visible )
            {;
                frmBrowser.Show;
                };
                X = null;
                };

                Case "SUBMIT TEMPLATE";
                //                 'frmBrowser.SubmitTemplate;

                Case "HIDE ALL WINDOWS", "HIDEALLWINDOWS";
                if ( frmBrowser.Visible )
            {
 frmBrowser.Hide;

                Case "UNLOAD";
                Unload frmBrowser;
                };
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As ISadAddin_ExecuteSoftCommand            {
                // TODO: Rewrite try/catch and/or goto. EH_SADAddin_ISadAddin_ExecuteSoftCommand;

                // On Error goto Next;

                EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:;
                return; // ???;

                EH_SADAddin_ISadAddin_ExecuteSoftCommand:;
                Parent.LogError "SADAddin", "ISadAddin_ExecuteSoftCommand", ex, Err.Description;
                goto EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue;

                Resume;
            }

            public As ISadAddin_ExecuteSoftCommandInline            {
                //    ';
            }

            public As ISadAddin_Shutdown()
            {
                try
{;
                MySadCommands.Clear;
                MySadCommands.Parent = null;
                MySadCommands = null;
                Parent = null;

                ISadAddin_Shutdown = true;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As ISadAddin_Startup            {
                try
{;
                ;

                Parent = oParent;
                MySadCommands.Parent = oParent;

                ISadAddin_Startup = (ex = 0);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

        }
    }
