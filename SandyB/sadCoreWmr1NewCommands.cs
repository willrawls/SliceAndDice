using System;

namespace MetX.SliceAndDice
{
		public class NewCommands
		{

			public  Parent;
			public  CommandsSupported;

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
				get; set; // Was get only
			}


			public void Class_Initialize()
			{
				try
{;
				CommandsSupported = new SliceAndDice.CSadCommands();

				CommandsSupported.ParameterDelimiter = ",";
				CommandsSupported.ParameterTypeDelimiter = ":";
				CommandsSupported.Attributes("Version") = App.Major + "." + App.Minor + "." + App.Revision;
				End With;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void Class_Terminate()
			{
				try
{;
				CommandsSupported = null;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As ISadAddin_ExecuteExternal			{
			}
			public As ISadAddin_ExecuteSoftCommand			{
				// TODO: Rewrite try/catch and/or goto. SandalError_ExecuteSoftCommand;
				String sVar1;
				String sVar2;
				String sVar3;
				String sVar4;


				if ( ! CommandsSupported(.SoftCommandName + "*C") Is null )
				{
;

				if ( CommandsSupported(.SoftCommandName + "*C").IsInline )
				{
 return; // ???;

				sVar1 = Parent.sGetToken(II.AllParameters, 1, ","): if ( Len(.SoftVars(sVar1)) )
				{
 sVar1 = II.SoftVars(sVar1);
				sVar2 = Parent.sGetToken(II.AllParameters, 2, ","): if ( Len(.SoftVars(sVar2)) )
				{
 sVar2 = II.SoftVars(sVar2);
				sVar3 = Parent.sGetToken(II.AllParameters, 3, ","): if ( Len(.SoftVars(sVar3)) )
				{
 sVar3 = II.SoftVars(sVar3);
				sVar4 = Parent.sGetToken(II.AllParameters, 4, ","): if ( Len(.SoftVars(sVar4)) )
				{
 sVar4 = II.SoftVars(sVar4);

				try
{;

				switch UCase$(.SoftCommandName);
				// Case "SOMETHING";
				//       ISadAddin_ExecuteSoftCommand = true;
				};
				};
				End With;

				SandalError_ExecuteSoftCommand_Continue:;
				return; // ???;

				SandalError_ExecuteSoftCommand:;

				ErrorsOcurred II, sVar1, sVar2, sVar3, sVar4, string.Empty;
				Resume SandalError_ExecuteSoftCommand_Continue;

				Resume;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As ISadAddin_ExecuteSoftCommandInline			{
				// TODO: Rewrite try/catch and/or goto. SandalError_ExecuteSoftCommandInline;
				Template;
				Area;
				Curr;
				As bInlineCommandExecuted;
				lParameterCount;
				sDefault;
				sT;
				sVar();

				CurrParam;
				ParameterCount;

				5) sVar(1;

				if ( ! CommandsSupported(sInlineSoftCommandName + "*I") Is null )
				{
;
				if ( ! CommandsSupported(sInlineSoftCommandName + "*I").IsInline )
				{
 return; // ???;

				ParameterCount = Parent.lTokenCount(sParameters, ",");

				if ( ParameterCount > 0 )
				{;
				if ( ParameterCount < 5 )
				{;
				5) sVar(1;
				}
				else
				{;
				ParameterCount) sVar(1;
				};
				For CurrParam = 1 To ParameterCount;
				sVar(CurrParam) = Parent.sGetToken(sParameters, CurrParam, ",");
				if ( Len(SoftVars(sVar(CurrParam))) )
				{
 sVar(CurrParam) = SoftVars(sVar(CurrParam));
				};
				};

				switch sInlineSoftCommandName;
				Case "FINDPATTERNINFILE", "FINDINFILE", "FINDREGEXPINFILE";
				sResults = FindPatternInFile(Parent.sGetToken(sParameters, 1, "="), Parent.sAfter(sParameters, 1, "="));
				ISadAddin_ExecuteSoftCommandInline = true;
				};
				};

				SandalError_ExecuteSoftCommandInline_Continue:;
				return; // ???;

				SandalError_ExecuteSoftCommandInline:;
				CInsertionInfo X;

				X = new CInsertionInfo();

				X.AllParameters = sParameters;
				X.Result = sResults;
				X.SoftVars = SoftVars;
				X.SoftCommandName = sInlineSoftCommandName;
				X.CurrentLineToProcess = "-Inline Substition-";
				End With;

				ErrorsOcurred X, sVar(1), sVar(2), sVar(3), sVar(4), sVar(5);

				X.SoftVars = null;
				X = null;

				Resume SandalError_ExecuteSoftCommandInline_Continue;
				Resume;
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
				CommandsSupported.Parent = oParent;



				CommandsSupported.Add("FindInFile",.Aliases = "FindPatternInFile";
				CommandsSupported.Add("FindInFile",.Examples = "~~FindInFile c:\sadResults.txt=occ*";
				CommandsSupported.Add("FindInFile",.OneLineDescription = "[FN As PathAndFile] = [Expression As WSHRegExpString]";
				CommandsSupported.Add("FindInFile",.Comments = "Expression is as the WSH ";
				End With;

				CommandsSupported.All = Parent.sFileContents(Parent.TemplateDatabasePath + "sadCoreWmr.txt");
				ISadAddin_Startup = (ex = 0);

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
				fh;
				sMessage;

				if ( ex <> 0 )
				{;
				sMessage = "Error executing SoftCode:" + vbNewLine;
				sMessage = sMessage + vbTab + "Occured:      " + Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM");
				sMessage = sMessage + vbTab + "Sandal:       sadCoreWmr.NewCommands" + vbNewLine;
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
				};

				sMessage = sMessage + vbNewLine + vbNewLine + "Would you like to cancel processing ?" + vbNewLine;
				sMessage = sMessage + vbNewLine + vbTab + "IF YOU SELECT:" + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "YES: This template should be cancelled.";
				sMessage = sMessage + vbTab + vbTab + vbTab + "NOTE: Due to the nature of this addin, processing is not guarenteed to stop.";
				sMessage = sMessage + vbTab + vbTab + "NO : Processing will continue with the next line of SoftCode.";
				sMessage = sMessage + vbTab + "NOTE: This information will be stored in: """ + App.Path + IIf(Right$(App.Path, 1) <> "\", "\", string.Empty) + "\sadCoreWmrError.Log""";

				if ( MsgBox(sMessage, vbYesNo, "CANCEL PROCESSING ?") = vbYes )
				{;
				if ( ! II Is null )
				{
;
				II.LinesLeftToProcess = string.Empty;
				};
				sMessage = sMessage + vbNewLine + "  *** User decided to CANCEL processing.";
				ErrorsOcurred = true ' Tell parent function processing has been cancelled.;
				}
				else
				{;
				sMessage = sMessage + vbNewLine + "  *** User choose to CONTINUE after error.";

				};

				fh = FreeFile;
				Open App.Path + IIf(Right$(App.Path, 1) <> "\", "\", string.Empty) + "\sadCoreWmrError.Log" For Append As #fh;
				Print #fh, sMessage;
				Close #fh;
				};
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As FindPatternInFile			{
				FileSystemObject fs;
				TextStream ts;
				RegExp re;
				String text;

				fs = new FileSystemObject();
				ts = fs.OpenTextFile(sFilename, ForReading);
				text = ts.ReadAll;
				re = new RegExp();
				re.Pattern = sRegularExpression;
				FindPatternInFile = re.Test(text);
			}
		}
}
