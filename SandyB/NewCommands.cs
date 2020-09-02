using System;

namespace MetX.SliceAndDice
{
		public class NewCommands
		{

			public  mvbInst;
			public  Parent;
			public  asaSaved;
			public  MySadCommands;
			public As Const;
			public  bEOLAtEndOfLine;
			public  lWrapLength;
			public  lThisWrap;
			public  lLineOffset;
			public  Decision;
			public  sWordWrapped;
			public  sToWrap;
			public  Token1;
			public  Token2;
			public  SoftVar1;
			public  SoftVar2;
			public  sOperator;
			public  bPreviousWasAnUnderscore;
			public  sOrig;
			public  CurrChar;
			public  LenOrig;
			public  sChar;
			public  sOut;
			public  Template;
			public  Area;
			public  lParamCount;
			public  CurrProject;
			public  CurrModule;
			public  CurrMember;
			public  sT;
			public  sVar1;
			public  sVar2;
			public  sVar3;
			public  ProcType;
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
				get
				{
					// ;
				}

			}


			public void Class_Initialize()
			{
				try
{;
				MySadCommands = new SliceAndDice.CSadCommands();

				MySadCommands.ParameterDelimiter = gsC;
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
				MySadCommands = null;
			}
			public As ISadAddin_ExecuteExternal			{
				// ;
			}
			public As ISadAddin_ExecuteSoftCommand			{
				// TODO: Rewrite try/catch and/or goto. EH_SADAddin_ISadAddin_ExecuteSoftCommand;




				if ( ! MySadCommands(.SoftCommandName + "*C") Is null )
				{
;
				if ( MySadCommands(.SoftCommandName + "*C").IsInline )
				{
 return; // ???;

				switch UCase(.SoftCommandName);
				Case "SAVEVARS";
				asaSaved = new SliceAndDice.CAssocArray();
				asaSaved.All = II.SoftVars.All;

				Case "RESTOREVARS";
				if ( ! asaSaved Is null )
				{
;
				//    MsgBox "You must execute the ~~SaveVars command first (See F1 help) before using the ~~RestoreVars command.";
				// Else;
				asaSaved.KeyValueDelimiter = " saved=";
				II.SoftVars.All = Replace(asaSaved.All + vbNewLine + II.SoftVars.All, "saved saved", "saved");
				Do While InStr(.SoftVars.All, "saved saved"): II.SoftVars.All = Replace(.SoftVars.All, "saved saved", "saved"): Loop;
				asaSaved.KeyValueDelimiter = " =";
			}
			public As Mangle			{
				try
{;

				sOrig = strIn;
				LenOrig = Len(sOrig);
				bPreviousWasAnUnderscore = true;
				sOut = string.Empty;

				if ( InStr(strIn, "_") > 0 )
				{;
				For CurrChar = 1 To LenOrig;
				sChar = Mid$(strIn, CurrChar, 1);
				switch sChar;
				Case "0" To "9";
				sOut = sOut + sChar;
				bPreviousWasAnUnderscore = false;

				Case "A" To "Z", "a" To "z";
				if ( bPreviousWasAnUnderscore )
				{;
				sOut = sOut + UCase(sChar);
				}
				else
				{;
				sOut = sOut + LCase(sChar);
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As ISadAddin_ExecuteSoftCommandInline			{
				// TODO: Rewrite try/catch and/or goto. ErrorHandler;

				lProcType;

				// 254      Dim Decision                As VbMsgBoxResult;

				// Dim sDefault                As String;

				if ( ! MySadCommands(sInlineSoftCommandName + "*I") Is null )
				{
;
				if ( ! MySadCommands(sInlineSoftCommandName + "*I").IsInline )
				{
 return; // ???;
				switch sInlineSoftCommandName;
				Case "MODULECONTENTS", "CONTENTS", "MEMBERS";
				sResults = string.Empty;
				foreach( var CurrMember in mvbInst.ActiveCodePane.CodeModule.Members );
				switch CurrMember.Type;
				Case vbext_mt_Const     ':   sResults = sResults + CurrMember.Name + "=CONSTANT=" + X + "$$$$";
				Case vbext_mt_Variable  ':   sResults = sResults + CurrMember.Name + "=VARIABLE=" + X + "=" + X + "=" + X + "=$$$$";
				Case }
				else
				{;
				switch CurrMember.Type;
				Case vbext_mt_Event:    sResults = sResults + CurrMember.Name + "=EVENT==";
				Case vbext_mt_Method:   sResults = sResults + CurrMember.Name + "=METHOD==";
				Case vbext_mt_Property: sT = mvbInst.ActiveCodePane.CodeModule.ProcOfLine(CurrMember.CodeLocation, lProcType);
				sResults = sResults + CurrMember.Name + "=PROPERTY=" + Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", true, string.Empty) + "=";
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


				// If Now() > 36300 Then;
				//    MsgBox "The Slice and Dice evaluation has expired. Thank you for participating." & Chr(13) & Chr(13) & "Please download the latest evaluation from: http://www.sliceanddice.com/VB5CodeWalker";
				//    return; // ???;
				// End If;

				// If UCase(oParent.Version) <> UCase(MySadCommands.Attributes("Version")) Then;
				//    If MsgBox("Version mismatch:" & Chr(13) & Chr(9) & "SADAddin.NewCommands = " & MySadCommands.Attributes("Version") & Chr(13) & Chr(9) & "Slice and Dice = " & oParent.Version & Chr(13) & Chr(13) & "Continue loading DLL ?", vbYesNo, "*** WARNING - DLL Version mismatch **") = vbNo Then;
				//       return; // ???;
				//    End If;
				// End If;

				Parent = oParent;
				MySadCommands.Parent = oParent;

				mvbInst = vbInst;


				MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath + "SADAddin.txt");
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
				sMessage = sMessage + vbTab + "Sandal:       sadAddin.NewCommands" + vbNewLine;
				if ( Erl <> 0 )
				{
 sMessage = sMessage + vbTab + "Sandal Line:  " + Erl + vbnewLine;
				sMessage = sMessage + vbTab + "Error Number: " + ex + vbNewLine;
				sMessage = sMessage + vbTab + "Description:  " + Err.Description + vbNewLine;

				sMessage = sMessage + vbNewLine + vbTab + "SoftCode Parameters (Resolved):" + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "sVar1 = " + sVar1 + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "sVar2 = " + sVar2 + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "sVar3 = " + sVar3 + vbNewLine;
				sMessage = sMessage + vbTab + vbTab + "sVar4 = " + sVar4 + vbNewLine;

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
