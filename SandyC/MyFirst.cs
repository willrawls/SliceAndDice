using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
        public class MyFirst
        {

        public SliceAndDice.Wizard Parent;
        public SliceAndDice.CSadCommands MySadCommands;


            public void Class_Initialize()
            {
                MySadCommands = new SliceAndDice.CSadCommands();
                MySadCommands.ParameterDelimiter = ", ";
                MySadCommands.ParameterTypeDelimiter = " : ";
            }

            public void Class_Terminate()
            {
                MySadCommands = null;
            }

            public As ISadAddin_ExecuteInlineSoftCommand            {
                //    ';
            }

            public As ISadAddin_ExecuteSoftCommand            {
                // TODO: Rewrite try/catch and/or goto. EH_MyFirst_ISadAddin_ExecuteSoftCommand;
                ;

                if ( ! MySadCommands(sCommand) Is null )
            {;
                CurrFunction = MySadCommands(sCommand);
                CurrFunction.Parameters = sParameters;
                if ( CurrFunction.SyntaxIsValid )
            {;
                switch sCommand.ToUpper();
                Case "XGETTOKEN";
                Case "XTOKEN";
                };
                };
                };
                CurrFunction = null;

                EH_MyFirst_ISadAddin_ExecuteSoftCommand_Continue:;
                return; // ???;

                EH_MyFirst_ISadAddin_ExecuteSoftCommand:;
                LogError "MyFirst", "ISadAddin_ExecuteSoftCommand", ex, Err.Description;
                goto EH_MyFirst_ISadAddin_ExecuteSoftCommand_Continue;

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

                Parent = oParent;
                MySadCommands.Parent = oParent;

                SetupSoftCommandDictionary<string,string>();

                ISadAddin_Startup = (ex = 0);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void SetupSoftCommandDictionary<string,string>()
            {

                MySadCommands.Clear;
                MySadCommands.Add(                                                                                                   "xGetToken", "SoftVarIn : stringOrSoftVar, nToken : long Opt, sDelim : Quotedstring Opt");
                MySadCommands.Add(                                                                                                   "xToken", "SoftVarIn : stringOrSoftVar, nToken : long Opt, sDelim : Quotedstring Opt");

            }

        }
    }
