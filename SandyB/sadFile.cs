using System;

namespace MetX.SliceAndDice
{
		public class NewCommands
		{

			public  Parent;
			public  MySadCommands;
			public int fh;
			public string sVar1;
			public string sVar2;
			public string sVar3;
			public string sVar4;
			public  Template;
			public  Area;
			public  CurrSet;
			public As bInlineCommandExecuted;
			public  lParameterCount;
			public  sDefault;
			public  sT;
			public  sVar();
			public  CurrParam;
			public  ParameterCount;
			public CInsertionInfo X;
			public  fh;
			public  sMessage;

			public SliceAndDice.CSadCommands ISadAddin_CommandSet
			{
				get
				{
					try
{
					ISadAddin_Command = MySadCommands;
					
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
				MySadCommands = new SliceAndDice.CSadCommands();

				MySadCommands.ParameterDelimiter = ",";
				MySadCommands.ParameterTypeDelimiter = ":";
				MySadCommands.Attributes("Version") = App.Major + "." + App.Minor + "." + App.Revision;
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
				if ( fh <> 0 )
				{;
				Close #fh;
				fh = 0;
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
				// TODO: Rewrite try/catch and/or goto. EH_SADAddin_ISadAddin_ExecuteSoftCommand;


				if ( ! MySadCommands(.SoftCommandName + "*C") Is null )
				{
;

				if ( MySadCommands(.SoftCommandName + "*C").IsInline )
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
				Case "CHANGEDRIVE", "DRIVE", "CHDRIVE";
				ChDrive sVar1;
				if ( ErrorsOcurred(II, sVar1, sVar2, sVar3, sVar4, string.Empty) )
				{;
				if ( fh <> 0 )
				{
 Close #fh: fh = 0;
				ISadAddin_ExecuteSoftCommand = true;
				return; // ???;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As ISadAddin_ExecuteSoftCommandInline			{
				// TODO: Rewrite try/catch and/or goto. EH_SADAddin_ISadAddin_ExecuteSoftCommandInline;


				5) sVar(1;

				if ( ! MySadCommands(sInlineSoftCommandName + "*I") Is null )
				{
;
				if ( ! MySadCommands(sInlineSoftCommandName + "*I").IsInline )
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
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As ISadAddin_Startup			{
				try
{;



				Parent = oParent;
				MySadCommands.Parent = oParent;


				MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath + "sadFile.txt");
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

				if ( ex <> 0 )
				{;
				sMessage = "Error executing SoftCode:" + vbNewLine;
				sMessage = sMessage + vbTab + "Occured:      " + Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM");
				sMessage = sMessage + vbTab + "Sandal:       sadFile.NewCommands" + vbNewLine;
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
