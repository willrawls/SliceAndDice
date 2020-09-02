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



            public void BrowseTo            {
                ;
                if ( frmBrowser.Visible )
            {;
                frmBrowser.Hide;
                }
            else
            {;
                X = null;
                X = Parent.SandyWindow;
                frmBrowser.StartingAddress = sURL;
                frmBrowser.brwWebBrowser.Navigate frmBrowser.StartingAddress;
                if ( X Is null )
            {;
                frmBrowser.Show 0;
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
                DoEvents: DoEvents: DoEvents;
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
                ;

                try
{;
                if ( frmBrowser.Parent Is null )
            {;
                frmBrowser.Parent = Me;
                };

                switch sValue.ToUpper();
                Case "IMPORT TEMPLATES";
                BrowseTo "http://www.vbcode.com";
                frmBrowser.ZOrder;

                Case "SUBMIT TEMPLATE";
                BrowseTo "http://www.vbcode.com/submit.htm";
                Do Until frmBrowser.NavigationComplete;
                DoEvents;
                Loop;

                sAuthorName = GetSetting(App.ProductName, "Last", "AuthorName", string.Empty);
                //                 'If Len(sAuthorName) = 0 Then;
                sAuthorName = InputBox("What is the Template Author's Name ?", "SUBMIT TEMPLATE TO VBCODE.COM", sAuthorName);
                //                 'End If;
                SaveSetting App.ProductName, "Last", "AuthorName", sAuthorName;

                sAuthorEmail = GetSetting(App.ProductName, "Last", "AuthorEmail", string.Empty);
                //                 'If Len(sAuthorEmail) = 0 Then;
                sAuthorEmail = InputBox("What is the Template Author's Email ?", "SUBMIT TEMPLATE TO VBCODE.COM", sAuthorEmail);
                //                 'End If;
                SaveSetting App.ProductName, "Last", "AuthorEmail", sAuthorEmail;

                frmBrowser.brwWebBrowser.Document.Forms(0).Item("AuthorName").Value = sAuthorName;
                frmBrowser.brwWebBrowser.Document.Forms(0).Item("AuthorEmail").Value = sAuthorEmail;
                frmBrowser.brwWebBrowser.Document.Forms(0).Item("Task").Value = Parent.CurrentTemplate.Key;
                if ( Len(Parent.CurrentTemplate.memoCodeAtTop) > 0 )
            {;
                frmBrowser.brwWebBrowser.Document.Forms(0).Item("Declarations").Value = "~~' Submitted from Slice and Dice" + vbNewLine + Parent.CurrentTemplate.memoCodeAtTop;
                frmBrowser.brwWebBrowser.Document.Forms(0).Item("CodeSnippet").Value = Parent.CurrentTemplate.memoCodeAtBottom;
                }
            else
            {if ( Len(Parent.CurrentTemplate.memoCodeAtBottom) > 0 )
            {;
                frmBrowser.brwWebBrowser.Document.Forms(0).Item("CodeSnippet").Value = "~~' Submitted from Slice and Dice" + vbNewLine + Parent.CurrentTemplate.memoCodeAtBottom;
                };
                frmBrowser.ZOrder;

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

                if ( ! frmBrowser Is null )
            {;
                frmBrowser.timTimer.Enabled = false;
                Unload frmBrowser;
                };

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
