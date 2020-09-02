using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
        public class NewCommands
        {

        public SliceAndDice.Wizard Parent;
        public SliceAndDice.CSadCommands MySadCommands;


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
        CAssocArray Externals ;
         Externals = new CAssocArray();
        Externals.All = "Testing Externals=Testing 123";
         ISadAddin_Externals = Externals;
        Externals = null;
        }

    }



            public void Class_Initialize()
            {
                try
{;
                MySadCommands = new SliceAndDice.CSadCommands();

                MySadCommands.ParameterDelimiter = ",";
                MySadCommands.ParameterTypeDelimiter = ":";
                MySadCommands.Attributes("Version") = App.Major + "." + App.Minor + "." + App.Revision + " Beta";

                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void Class_Terminate()
            {
                MySadCommands = null;
            }

            public As ISadAddin_ExecuteExternal            {
                switch sKey.ToUpper();
                Case "TESTING EXTERNALS";
                MsgBox sValue;
                };
            }

            public As ISadAddin_ExecuteSoftCommand            {
                // TODO: Rewrite try/catch and/or goto. EH_SADAddin_ISadAddin_ExecuteSoftCommand;
                ;

                ;
                ;
                ;

                ;
                ;
                ;
                ;
                ;
                ;
                ;


                if ( ! MySadCommands(.SoftCommandName + "*C") Is null )
            {;
                if ( MySadCommands(.SoftCommandName + "*C").IsInline )
            {
 return; // ???;

                switch .SoftCommandName.ToUpper();
                Case "TESTING";
                MsgBox "Soft command 'Testing' executed.";
                ISadAddin_ExecuteSoftCommand = true;

                Case "ANOTHERONE";
                MsgBox "Soft command 'AnotherOne' executed.";
                ISadAddin_ExecuteSoftCommand = true;

                //                  'Case "X";
                //                  'Case "Y";
                //                  'Case "Z";
                };
                };


                EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue:;
                return; // ???;

                EH_SADAddin_ISadAddin_ExecuteSoftCommand:;
                Parent.LogError "SADAddin", "ISadAddin_ExecuteSoftCommand", ex, Err.Description;
                goto EH_SADAddin_ISadAddin_ExecuteSoftCommand_Continue;

                Resume;
            }

            public As ISadAddin_ExecuteSoftCommandInline            {
                // TODO: Rewrite try/catch and/or goto. EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline;
                ;

                ;
                ;
                ;
                ;

                ;
                ;
                ;
                ;
                ;

                if ( ! MySadCommands(sInlineSoftCommandName + "*I") Is null )
            {;
                if ( ! MySadCommands(sInlineSoftCommandName + "*I").IsInline )
            {
 return; // ???;
                switch sInlineSoftCommandName;
                Case "INLINE1";
                sResults = InputBox("What do you want in here ?");
                ISadAddin_ExecuteSoftCommandInline = true;

                //             'Case "INLINEX";
                sResults = "";
                ISadAddin_ExecuteSoftCommandInline = true;

                //             'Case "INLINEY";
                sResults = "";
                ISadAddin_ExecuteSoftCommandInline = true;

                //             'Case "INLINEZ";
                sResults = "";
                ISadAddin_ExecuteSoftCommandInline = true;

                };
                };

                EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue:;
                return; // ???;

                EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline:;
                Parent.LogError "SADAddin", "ISadAddin_ExecuteSoftCommandInline", ex, Err.Description;
                goto EH_SADAddin_ISadAddin_ExecuteSoftCommandInlineInline_Continue;

                Resume;
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

                if ( oParent.Version) <> UCase(MySadCommands.Attributes("Version")) .ToUpper()
            {;
                if ( MsgBox("Version mismatch:" + Chr(13) + Chr(9) + "SADAddin.NewCommands = " + MySadCommands.Attributes("Version") + Chr(13) + Chr(9) + "Slice and Dice = " + oParent.Version + Chr(13) + Chr(13) + "Continue loading DLL ?", vbYesNo, "*** WARNING - DLL Version mismatch **") = vbNo )
            {;
                return; // ???;
                };
                };

                Parent = oParent;
                MySadCommands.Parent = oParent;

                ;
                MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath + "sadMyFirstAddin.txt");
                ISadAddin_Startup = (ex = 0);
                ;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

        }
    }
