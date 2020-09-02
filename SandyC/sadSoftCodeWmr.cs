using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
        public class NewCommands
        {

        public object Parent;
        public object CommandsSupported;
        public object ItemLists;
        public object ItemsToProcess;
        public 3) vParams(0;
        public object CurrItem;
        public object CurrSection;
        public object Template;
        public object sDefault;
        public object sT;
        public As bInlineCommandExecuted;
        public object Area;
        public object CurrSet;
        public object lParameterCount;
        public object CurrParam;
        public object ParameterCount;
        public object CurrItem;
        public object CurrSection;
        public 3) vParams(0;
        public CInsertionInfo X;
        public object fh;
        public object sMessage;


                public SliceAndDice.CSadCommands ISadAddin_CommandSet
    {
        get
        {
        try
{

         ISadAddin_Command = CommandsSupported;
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
        Externals.this.Clear()true;
        Externals.Item("SoftCommand HTML Reference") = "HTMLREFERENCE";
        Externals.Item("Revision History") = "REVISIONHISTORY";
         ISadAddin_Externals = Externals;
        Externals = null;
        }

    }



            public As ShowMessage            {
                ;
                ;
                ;
                ;
                X = new frmMessage();

                X.Text = sTitle;

                .txtMessage.Text = sMessageToShow;
                .txtMessage.ToolTipText = sToolTip;
                .txtMessage.SelStart = 0;
                .txtMessage.SelLength = Len(.Text) + 1;


                xWidth = X.TextWidth(sMessageToShow) + 500;
                xHeight = X.TextHeight(sMessageToShow) + 500;
                ;
                xWidth +=  X.TextWidth("Q") * 3;
                xHeight +=  X.TextHeight("QWERTY") * 2;

                if ( xWidth > Screen.Width - 1000 )
            {
 xWidth == Screen.Width - 1000;
                if ( xWidth < 1000 )
            {
 xWidth == 1000;

                if ( xHeight > Screen.Height - 1000 )
            {
 xHeight == Screen.Height - 1000;
                if ( xHeight < 1000 )
            {
 xHeight == 1000;

                X.Width = xWidth;
                X.Height = xHeight;

                if ( ! Parent.SandyWindow Is null )
            {;
                X.Show vbModal, Parent.SandyWindow;
                }
            else
            {;
                X.Show vbModal;
                };

                if ( StrComp(NewMessageText, sMessageToShow, vbTextCompare) <> 0 )
            {;
                ShowMessage = NewMessageText;
                };

                X = null;
            }

            public void Class_Terminate()
            {
                try
{;
                CommandsSupported = null;
                //        'Set asaList = Nothing;
                ItemLists = null;
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

                switch sValue.ToUpper();
                Case "REVISIONHISTORY";
                if ( Len(Dir$(App.Path + "\SandyRevisions.txt")) = 0 )
            {;
                MsgBox "Revision history file ('SandyRevisions.txt') not in the application path." + vbNewLine + vbTab + "Unable to view at this time.";
                }
            else
            {;
                ShowMessage Parent.sFileContents(App.Path + "\SandyRevisions.txt"), "Slice and Dice Revision History", "Abridged history of changes made between revisions of Slice and Dice.";
                };

                Case "HTMLREFERENCE";
                Screen.MousePointer = vbHourglass;
                Commands = Parent.SoftCommands;
                sHTML = "<HTML><BODY><TABLE>";
                foreach( var CurrCommand in Commands;

                sHTML +=  vbTab + "<TR>" + vbNewLine;
                sHTML +=  vbTab + vbTab + "<TD><B><H2>" + CurrCommand.SoftCommandName + "</h2></B><UL><BR>" + vbNewLine;
                if ( Len(.Aliases) > 0 )
            {
.Aliases.Substring(0, 3).Substring( Len(.Aliases) - 4) + "</H3></BLOCKQUOTE><BR>" + vbnewLine;
                if ( Len(.OneLineDescription) > 0 )
            {
 sHTML == sHTML + vbTab + vbTab + "<LI><B><H3>Summary</H3></B> " + "<BLOCKQUOTE>" + CurrCommand.OneLineDescription + "</BLOCKQUOTE><BR>" + vbnewLine;
                if ( Len(.SeeAlso) > 0 )
            {
 sHTML == sHTML + vbTab + vbTab + "<LI><B><H3>See Also</H3></B> " + "<BLOCKQUOTE>" + Replace(.SeeAlso, vbnewLine, "<BR>" + vbnewLine) + "</BLOCKQUOTE><BR>" + vbnewLine;
                if ( Len(.LongDescription) > 0 )
            {
 sHTML == sHTML + vbTab + vbTab + "<LI><B><H3>long Description</H3></B> " + "<BLOCKQUOTE>" + Replace(.LongDescription, vbnewLine, "<BR>" + vbnewLine) + "</BLOCKQUOTE><BR>" + vbnewLine;
                if ( Len(.Comments) > 0 )
            {
 sHTML == sHTML + vbTab + vbTab + "<LI><B><H3>Comments</H3></B> " + "<BLOCKQUOTE>" + Replace(.Comments, vbnewLine, "<BR>" + vbnewLine) + "</BLOCKQUOTE><BR>" + vbnewLine;
                if ( Len(.Examples) > 0 )
            {
 sHTML == sHTML + vbTab + vbTab + "<LI><B><H3>Examples</H3></B> " + "<BLOCKQUOTE>" + Replace(.Examples, vbnewLine, "<BR>" + vbnewLine) + "</BLOCKQUOTE><BR>" + vbnewLine;
                sHTML +=  vbTab + "</UL></TD></TR>" + vbNewLine;

                } // CurrCommand;
                sHTML +=  vbNewLine + "</TABLE></BODY></HTML>" + vbNewLine;
                Parent.SaveToFile Parent.TemplateDatabasePath + "sadReference.html", sHTML;
                sHTML = string.Empty;
                Parent.BrowseTo Parent.TemplateDatabasePath + "sadReference.html";
                ISadAddin_ExecuteExternal = true;
                Screen.MousePointer = vbDefault;
            }

            public As ISadAddin_ExecuteSoftCommand            {
                // TODO: Rewrite try/catch and/or goto. SandalError_ExecuteSoftCommand;
                ;


                if ( ! CommandsSupported(.SoftCommandName + "*C") Is null )
            {;

                if ( CommandsSupported(.SoftCommandName + "*C").IsInline )
            {
 return; // ???;

                vParams(0) = Split(II.AllParameters, ",");
                vParams(1) = Array(Parent.sGetToken(II.AllParameters, 1, " - "), Parent.sAfter(II.AllParameters, 1, " - "));
                vParams(2) = vParams(0);
                vParams(3) = vParams(1);

                for(var CurrSection = 0; CurrSection < 1; CurrSection++)  {;
                for(var CurrItem = 0; CurrItem < UBound(vParams(CurrSection)); CurrItem++)  {;
                if ( Len(.SoftVars(vParams(CurrSection)(CurrItem) + string.Empty)) )
            {;
                vParams(2)(CurrItem) = vParams(CurrSection)(CurrItem);
            }

            public As ISadAddin_ExecuteSoftCommandInline            {
                // TODO: Rewrite try/catch and/or goto. SandalError_ExecuteSoftCommandInline;


                if ( ! CommandsSupported(sInlineSoftCommandName + "*I") Is null )
            {;
                if ( ! CommandsSupported(sInlineSoftCommandName + "*I").IsInline )
            {
 return; // ???;

                ParameterCount = Parent.lTokenCount(sParameters, ",");

                if ( ParameterCount > 0 )
            {;
                vParams(0) = Split(sParameters, ",");
                vParams(1) = Array(Parent.sGetToken(sParameters, 1, " - "), Parent.sAfter(sParameters, 1, " - "));
                vParams(2) = vParams(0);
                vParams(3) = vParams(1);

                for(var CurrSection = 0; CurrSection < 1; CurrSection++)  {;
                for(var CurrItem = 0; CurrItem < UBound(vParams(CurrSection)); CurrItem++)  {;
                if ( Len(SoftVars(vParams(CurrSection)(CurrItem) + string.Empty)) )
            {;
                vParams(2)(CurrItem) = vParams(CurrSection)(CurrItem);
            }

            public As ISadAddin_Shutdown()
            {
                try
{;
                CommandsSupported.Clear;
                CommandsSupported.Parent = null;
                CommandsSupported = null;
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


                Parent = oParent;
                ;
                try
{;
                //        'Set asaList = New CAssocArray;

                CommandsSupported = new SliceAndDice.CSadCommands();

                CommandsSupported.Parent = oParent;
                CommandsSupported.ParameterDelimiter = ",";
                CommandsSupported.ParameterTypeDelimiter = ":";
                CommandsSupported.Attributes("Version") = App.Major + "." + App.Minor + "." + App.Revision;


                .Add(                                                                                                                                                                                                      "ShowMessage", false).Aliases = ", ShowMsg, ";
                .Add(                                                                                                                                                                                                      "ShowMessage", false).Examples = "~~ShowMessage Testing" + vbNewLine + "~~ X=This is a test" + vbNewLine + "~~ShowMessage X";
                .Add(                                                                                                                                                                                                      "ShowMessage", false).Comments = "Shows a multi-line message form capable of showing a read-only message to the user.";
                .Add(                                                                                                                                                                                                      "ShowMessage", false).SeeAlso = "EditMessage";
                stringOrSoftVar]" .Add(                                                                                                                                                                                                      "ShowMessage", false).Syntax = "[Message;
                .Add(                                                                                                                                                                                                      "ShowMessage", false).OneLineDescription = "Show a multi-line soft variable / string.";



                .Add(                                                                                                                                                                                                      "EditMessage", false).Aliases = ", EditMsg, ";
                .Add(                                                                                                                                                                                                      "EditMessage", false).Examples = "~~EditMessage Y=Some test to edit here." + vbNewLine + "~~ X=This is another test" + vbNewLine + "~~EditMessage Z=X";
                .Add(                                                                                                                                                                                                      "EditMessage", false).Comments = "Shows a multi-line message form with a simple text box editor" + vbNewLine + "that can be used to edit a soft variable and return/overwrite that variable or another.";
                .Add(                                                                                                                                                                                                      "EditMessage", false).SeeAlso = "EditMessage";
                stringOrSoftVar]" .Add(                                                                                                                                                                                                      "EditMessage", false).Syntax = "[EditedOut As SoftVar] = [Message;
                .Add(                                                                                                                                                                                                      "EditMessage", false).OneLineDescription = "Edit a multi-line soft variable.";



                .Add(                                                                                                                                                                                                      "ExecuteSoftCode", true).Aliases = ", RunSoftCode, ";
                .Add(                                                                                                                                                                                                      "ExecuteSoftCode", true).Examples = "~~GetClipboardText cbText" + vbNewLine + "~~ExecuteSoftCode cbText";
                SoftCode]" .Add(                                                                                                                                                                                                      "ExecuteSoftCode", true).Syntax = "[CodeToExecute;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As ErrorsOcurred            {
                try
{;

                if ( ex <> 0 )
            {;
                sMessage = "Error executing SoftCode:" + vbNewLine;
                sMessage +=  vbTab + "Occured:      " + Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM");
                sMessage +=  vbTab + "Sandal:       sadSoftCodeWmr.NewCommands" + vbNewLine;
                if ( Erl <> 0 )
            {
 sMessage == sMessage + vbTab + "Sandal Line:  " + Erl + vbnewLine;
                sMessage +=  vbTab + "Error Number: " + ex + vbNewLine;
                sMessage +=  vbTab + "Description:  " + Err.Description + vbNewLine;

                sMessage +=  vbNewLine + vbTab + "SoftCode Parameters (Resolved):" + vbNewLine;
                sMessage +=  vbTab + vbTab + "sVar(1) = " + sVar1 + vbNewLine;
                sMessage +=  vbTab + vbTab + "sVar(2) = " + sVar2 + vbNewLine;
                sMessage +=  vbTab + vbTab + "sVar(3) = " + sVar3 + vbNewLine;
                sMessage +=  vbTab + vbTab + "sVar(4) = " + sVar4 + vbNewLine;
                sMessage +=  vbTab + vbTab + "sVar(5) = " + sVar5 + vbNewLine;

                if ( ! II Is null )
            {;
                sMessage +=  vbNewLine + vbTab + "(Unresolved) CInsertionInfo contents:" + vbNewLine;
                sMessage +=  vbTab + vbTab + "TemplateName = " + II.TemplateName + vbNewLine;
                sMessage +=  vbTab + vbTab + "CurrentLineToProcess = " + II.CurrentLineToProcess + vbNewLine;
                sMessage +=  vbTab + vbTab + "SoftCommandName = " + II.SoftCommandName + vbNewLine;
                sMessage +=  vbTab + vbTab + "AllParameters = " + II.AllParameters + vbNewLine;
                sMessage +=  vbTab + vbTab + "Result     (LHS) = " + II.Result + vbNewLine;
                sMessage +=  vbTab + vbTab + "Expression (RHS) = " + II.Expression + vbNewLine;
                sMessage +=  vbTab + vbTab + "ExternalFilename = " + II.ExternalFilename + vbNewLine;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

        }
    }
