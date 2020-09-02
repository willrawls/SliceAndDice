using System;

namespace MetX.SliceAndDice
{
		public class NewCommands
		{

			public  Parent;
			public  MySadCommands;
			public string sVar1;
			public string sVar2;
			public string sVar3;
			public string sVar4;
			public Evaluator ktEval;
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

			// '    Dim Template                As CTemplate
'    Dim Area                    As Long
'    Dim CurrSet                 As Long
'    Dim bInlineCommandExecuted  As Boolean
'    Dim lParameterCount         As Long
'    Dim sDefault                As String
'    Dim sT                      As String
'    Dim sVar()                  As String
'
'    Dim CurrParam               As Long
'    Dim ParameterCount          As Long
'
'    ReDim sVar(1 To 5) As String
'
'    If Not MySadCommands(sInlineSoftCommandName & "*I") Is Nothing Then
'       If Not MySadCommands(sInlineSoftCommandName & "*I").IsInline Then return; // ???
'
'       ParameterCount = Parent.lTokenCount(sParameters, ",")
'
'       If ParameterCount > 0 Then
'          If ParameterCount < 5 Then
'             ReDim sVar(1 To 5) As String
'          Else
'             ReDim sVar(1 To ParameterCount) As String
'          End If
'          For CurrParam = 1 To ParameterCount
'              sVar(CurrParam) = Parent.sGetToken(sParameters, CurrParam, ",")
'              If Len(SoftVars(sVar(CurrParam))) Then sVar(CurrParam) = SoftVars(sVar(CurrParam))
'          Next CurrParam
'       End If
'
'       Select Case sInlineSoftCommandName
'              Case "DOSOMETHING"
'                 ' Do something and return results through the string 'sResults'
'                   sResults = vbNullString
'
'                   ISadAddin_ExecuteSoftCommandInline = true
'
'            ' More inline commands go here
'        End Select
'    End If
'
'    return; // ???
'
'    Dim X As CInsertionInfo
'
'    Set X = New CInsertionInfo
'    With X
'         .AllParameters = sParameters
'         .Result = sResults
'         Set .SoftVars = SoftVars
'         .SoftCommandName = sInlineSoftCommandName
'         .CurrentLineToProcess = "-Inline Substition-"
'    End With
'
'    ErrorsOcurred X, sVar(1), sVar(2), sVar(3), sVar(4), sVar(5)
'
'    Set X.SoftVars = Nothing
'    Set X = Nothing
'
'    Resume EH_SADAddin_ISadAddin_ExecuteSoftCommandInline_Continue
'    Resume
;
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

				MySadCommands.Add("KTEval",.Syntax = "[Result As SoftVar]=[Expression As SoftVarOrString]";
				MySadCommands.Add("KTEval",.OneLineDescription = "Significantly more advanced Eval command";
				MySadCommands.Add("KTEval",.LongDescription = "'Copyright� 1999, Tretyakov Konstantin" + vbNewLine +
                    "'_____________________________________________________" + vbNewLine +
                    "'This is the 'Evaluator' class: it inputs a string" + vbNewLine +
                    "'like ""2+2"" or ""2+4*sin(3.4)^2-8*arccos(0.55)"", etc" + vbNewLine +
                    "'_____________________________________________________" + vbNewLine +
                    "'You may use the code for free, if you give me credit." + vbNewLine +
                    "'if ( you modify it or make your own program with it," + vbNewLine +
                    "'I would VERY APPRECIATE, if you mail me it (or better-" + vbNewLine +
                    "'a link to it)" + vbNewLine +
                    "'On the whole - just do not stamp your name on what you haven't" + vbNewLine +
                    "'done quite alone." + vbNewLine +
                    "'This code was written totally by me, and 'it took me about" + vbNewLine +
                    "'2 days to code it (and about a year" + vbNewLine +
                    "'-that is,from the very moment I got interested in programming-" + vbNewLine +
                    "'I spent dreaming of having such a thing)" + vbNewLine +
                    "" + vbNewLine +
                    "'(BTW this code seems to be quite unique-" + vbNewLine +
                    "'I searched all over the Internet for such, but NOONE" + vbNewLine +
                    "'is giving the source for such things)" + vbNewLine +
                    "'______________________________________________________" + vbNewLine +
                    "'Yours Sincerely, Konstantin Tretyakov (kt_ee@yahoo.com)" + vbNewLine;
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
				// sVar3 = Parent.sGetToken(II.AllParameters, 3, ","): If Len(.SoftVars(sVar3)) Then sVar3 = .SoftVars(sVar3);
				// sVar4 = Parent.sGetToken(II.AllParameters, 4, ","): If Len(.SoftVars(sVar4)) Then sVar4 = .SoftVars(sVar4);

				// On Error Resume Next;

				switch UCase$(.SoftCommandName);
				Case "KTEVAL";
				//  Do something here;
				ktEval = new Evaluator();
				if ( ! ktEval Is null )
				{
;
				II.SoftVars(II.Result) = ktEval.Evaluate(II.Expression, true) + string.Empty;
			}
			public As ISadAddin_ExecuteSoftCommandInline			{
				// On Error GoTo EH_SADAddin_ISadAddin_ExecuteSoftCommandInline;
				// EH_SADAddin_ISadAddin_ExecuteSoftCommandInline_Continue:;
				// EH_SADAddin_ISadAddin_ExecuteSoftCommandInline:;
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


				// MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath & "sadKTEval.txt");
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
				sMessage = sMessage + vbTab + "Sandal:       sadKTEval.NewCommands" + vbNewLine;
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
