using System;

namespace MetX.SliceAndDice
{
		public class NewCommands
		{

			public  Parent;
			public  CommandsSupported;
			public  ItemLists;
			public  ItemsToProcess;
			public 3) vParams(0;
			public  CurrItem;
			public  CurrSection;
			public  Template;
			public  sDefault;
			public  sT;
			public As bInlineCommandExecuted;
			public  Area;
			public  CurrSet;
			public  lParameterCount;
			public  CurrParam;
			public  ParameterCount;
			public  CurrItem;
			public  CurrSection;
			public 3) vParams(0;
			public CInsertionInfo X;
			public  fh;
			public  sMessage;

			public SliceAndDice.CSadCommands ISadAddin_CommandSet
			{
				get
				{
					try
{
					ISadAddin_Command = CommandsSupported;
					
        }
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
				}

			}

			public SliceAndDice.CAssocArray ISadAddin_Externals
			{
				get
				{
					CAssocArray Externals;
					Externals = new CAssocArray();
					Externals.Clear true;
					Externals.Item("SoftCommand HTML Reference") = "HTMLREFERENCE";
					Externals.Item("Revision History") = "REVISIONHISTORY";
					End With;
					ISadAddin_Externals = Externals;
					Externals = null;
				}

			}


			public As ShowMessage			{
				X;
				As xWidth;
				Single xHeight;

				X = new frmMessage();

				X.Text = sTitle;

				X.txtMessage.Text = sMessageToShow;
				X.txtMessage.ToolTipText = sToolTip;
				X.txtMessage.SelStart = 0;
				X.txtMessage.SelLength = Len(.Text) + 1;
				End With;

				xWidth = X.txtMessage.TextWidth(sMessageToShow) + 500;
				xHeight = X.txtMessage.TextHeight(sMessageToShow) + 500;

				xWidth = xWidth + X.txtMessage.TextWidth("Q") * 3;
				xHeight = xHeight + X.txtMessage.TextHeight("QWERTY") * 2;

				if ( xWidth > Screen.Width - 1000 )
				{
 xWidth = Screen.Width - 1000;
				if ( xWidth < 1000 )
				{
 xWidth = 1000;

				if ( xHeight > Screen.Height - 1000 )
				{
 xHeight = Screen.Height - 1000;
				if ( xHeight < 1000 )
				{
 xHeight = 1000;

				X.txtMessage.Width = xWidth;
				X.txtMessage.Height = xHeight;

				if ( ! Parent.SandyWindow Is null )
				{
;
				X.txtMessage.Show vbModal, Parent.SandyWindow;
				}
				else
				{;
				X.txtMessage.Show vbModal;
				};

				if ( StrComp(NewMessageText, sMessageToShow, vbTextCompare) <> 0 )
				{;
				ShowMessage = NewMessageText;
				};
				End With;
				X = null;
			}
			public void Class_Terminate()
			{
				try
{;
				CommandsSupported = null;
				// Set asaList = Nothing;
				ItemLists = null;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As ISadAddin_ExecuteExternal			{
				CSadCommands Commands;
				CSadCommand CurrCommand;
				String sHTML;

				switch UCase$(sValue);
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
				foreach( var CurrCommand in Commands );

				sHTML = sHTML + vbTab + "<TR>" + vbNewLine;
				sHTML = sHTML + vbTab + vbTab + "<TD><B><H2>" + CurrCommand.SoftCommandName + "</h2></B><UL><BR>" + vbNewLine;
				if ( Len(.Aliases) > 0 )
				{
 sHTML = sHTML + vbTab + vbTab + "<LI><B><H3>Aliases</H3></B> " + "<BLOCKQUOTE><H3>" + Left$(Mid$(.Aliases, 3), Len(.Aliases) - 4) + "</H3></BLOCKQUOTE><BR>" + vbnewLine;
				if ( Len(.OneLineDescription) > 0 )
				{
 sHTML = sHTML + vbTab + vbTab + "<LI><B><H3>Summary</H3></B> " + "<BLOCKQUOTE>" + CurrCommand.OneLineDescription + "</BLOCKQUOTE><BR>" + vbnewLine;
				if ( Len(.SeeAlso) > 0 )
				{
 sHTML = sHTML + vbTab + vbTab + "<LI><B><H3>See Also</H3></B> " + "<BLOCKQUOTE>" + Replace(.SeeAlso, vbnewLine, "<BR>" + vbnewLine) + "</BLOCKQUOTE><BR>" + vbnewLine;
				if ( Len(.LongDescription) > 0 )
				{
 sHTML = sHTML + vbTab + vbTab + "<LI><B><H3>Long Description</H3></B> " + "<BLOCKQUOTE>" + Replace(.LongDescription, vbnewLine, "<BR>" + vbnewLine) + "</BLOCKQUOTE><BR>" + vbnewLine;
				if ( Len(.Comments) > 0 )
				{
 sHTML = sHTML + vbTab + vbTab + "<LI><B><H3>Comments</H3></B> " + "<BLOCKQUOTE>" + Replace(.Comments, vbnewLine, "<BR>" + vbnewLine) + "</BLOCKQUOTE><BR>" + vbnewLine;
				if ( Len(.Examples) > 0 )
				{
 sHTML = sHTML + vbTab + vbTab + "<LI><B><H3>Examples</H3></B> " + "<BLOCKQUOTE>" + Replace(.Examples, vbnewLine, "<BR>" + vbnewLine) + "</BLOCKQUOTE><BR>" + vbnewLine;
				sHTML = sHTML + vbTab + "</UL></TD></TR>" + vbNewLine;
				End With;
				};
				sHTML = sHTML + vbNewLine + "</TABLE></BODY></HTML>" + vbNewLine;
				Parent.SaveToFile Parent.TemplateDatabasePath + "sadReference.html", sHTML;
				sHTML = string.Empty;
				Parent.BrowseTo Parent.TemplateDatabasePath + "sadReference.html";
				ISadAddin_ExecuteExternal = true;
				Screen.MousePointer = vbDefault;
			}
			public As ISadAddin_ExecuteSoftCommand			{
				// TODO: Rewrite try/catch and/or goto. SandalError_ExecuteSoftCommand;
				EditedMessage;


				if ( ! CommandsSupported(.SoftCommandName + "*C") Is null )
				{
;

				if ( CommandsSupported(.SoftCommandName + "*C").IsInline )
				{
 return; // ???;

				vParams(0) = Split(II.AllParameters, ",");
				vParams(1) = Array(Parent.sGetToken(II.AllParameters, 1, " - "), Parent.sAfter(II.AllParameters, 1, " - "));
				vParams(2) = vParams(0);
				vParams(3) = vParams(1);

				For CurrSection = 0 To 1;
				For CurrItem = 0 To UBound(vParams(CurrSection));
				if ( Len(.SoftVars(vParams(CurrSection)(CurrItem) + string.Empty)) )
				{;
				vParams(2)(CurrItem) = vParams(CurrSection)(CurrItem);
			}
			public As ISadAddin_ExecuteSoftCommandInline			{
				// TODO: Rewrite try/catch and/or goto. SandalError_ExecuteSoftCommandInline;


				if ( ! CommandsSupported(sInlineSoftCommandName + "*I") Is null )
				{
;
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

				For CurrSection = 0 To 1;
				For CurrItem = 0 To UBound(vParams(CurrSection));
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
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As ISadAddin_Startup			{
				try
{;


				Parent = oParent;

				try
{;
				// Set asaList = New CAssocArray;

				CommandsSupported = new SliceAndDice.CSadCommands();

				CommandsSupported.Parent = oParent;
				CommandsSupported.ParameterDelimiter = ",";
				CommandsSupported.ParameterTypeDelimiter = ":";
				CommandsSupported.Attributes("Version") = App.Major + "." + App.Minor + "." + App.Revision;


				CommandsSupported.Add("ShowMessage",.Aliases = ", ShowMsg, ";
				CommandsSupported.Add("ShowMessage",.Examples = "~~ShowMessage Testing" + vbNewLine + "~~ X=This is a test" + vbNewLine + "~~ShowMessage X";
				CommandsSupported.Add("ShowMessage",.Comments = "Shows a multi-line message form capable of showing a read-only message to the user.";
				CommandsSupported.Add("ShowMessage",.SeeAlso = "EditMessage";
				CommandsSupported.Add("ShowMessage",.Syntax = "[Message As StringOrSoftVar]";
				CommandsSupported.Add("ShowMessage",.OneLineDescription = "Show a multi-line soft variable / string.";
				End With;


				CommandsSupported.Add("ShowMessage",.Add("EditMessage",.Aliases = ", EditMsg, ";
				CommandsSupported.Add("ShowMessage",.Add("EditMessage",.Examples = "~~EditMessage Y=Some test to edit here." + vbNewLine + "~~ X=This is another test" + vbNewLine + "~~EditMessage Z=X";
				CommandsSupported.Add("ShowMessage",.Add("EditMessage",.Comments = "Shows a multi-line message form with a simple text box editor" + vbNewLine + "that can be used to edit a soft variable and return/overwrite that variable or another.";
				CommandsSupported.Add("ShowMessage",.Add("EditMessage",.SeeAlso = "EditMessage";
				CommandsSupported.Add("ShowMessage",.Add("EditMessage",.Syntax = "[EditedOut As SoftVar] = [Message As StringOrSoftVar]";
				CommandsSupported.Add("ShowMessage",.Add("EditMessage",.OneLineDescription = "Edit a multi-line soft variable.";
				End With;


				CommandsSupported.Add("ShowMessage",.Add("EditMessage",.Add("ExecuteSoftCode",.Aliases = ", RunSoftCode, ";
				CommandsSupported.Add("ShowMessage",.Add("EditMessage",.Add("ExecuteSoftCode",.Examples = "~~GetClipboardText cbText" + vbNewLine + "~~ExecuteSoftCode cbText";
				CommandsSupported.Add("ShowMessage",.Add("EditMessage",.Add("ExecuteSoftCode",.Syntax = "[CodeToExecute As SoftCode]";
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As ErrorsOcurred			{
				try
{;

				if ( ex <> 0 )
				{;
				sMessage = "Error executing SoftCode:" + vbNewLine;
				sMessage = sMessage + vbTab + "Occured:      " + Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM");
				sMessage = sMessage + vbTab + "Sandal:       sadSoftCodeWmr.NewCommands" + vbNewLine;
				if ( Erl <> 0 )
				{
 sMessage = sMessage + vbTab + "Sandal Line:  " + Erl + vbnewLine;
				sMessage = sMessage + vbTab + "Error Number: " + ex + vbNewLine;
				sMessage = sMessage + vbTab + "Description:  " + Err.Description + vbNewLine;

				sMessage = sMessage + vbNewLine + vbTab + "SoftCode Parameters (Resolved):" + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "sVar(1) = " + sVar1 + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "sVar(2) = " + sVar2 + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "sVar(3) = " + sVar3 + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "sVar(4) = " + sVar4 + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "sVar(5) = " + sVar5 + vbNewLine;

				if ( ! II Is null )
				{
;
				sMessage = sMessage + vbNewLine + vbTab + "(Unresolved) CInsertionInfo contents:" + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "TemplateName = " + II.TemplateName + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "CurrentLineToProcess = " + II.CurrentLineToProcess + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "SoftCommandName = " + II.SoftCommandName + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "AllParameters = " + II.AllParameters + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "Result     (LHS) = " + II.Result + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "Expression (RHS) = " + II.Expression + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "ExternalFilename = " + II.ExternalFilename + vbNewLine;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
		}
}
