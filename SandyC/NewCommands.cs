using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
        public class NewCommands
        {

        public object mvbInst;
        public object Parent;
        public object asaSaved;
        public object MySadCommands;
        public As Const;
        public object bEOLAtEndOfLine;
        public object lWrapLength;
        public object lThisWrap;
        public object lLineOffset;
        public object Decision;
        public object sWordWrapped;
        public object sToWrap;
        public object Token1;
        public object Token2;
        public object SoftVar1;
        public object SoftVar2;
        public object sOperator;
        public object bPreviousWasAnUnderscore;
        public object sOrig;
        public object CurrChar;
        public object LenOrig;
        public object sChar;
        public object sOut;
        public object Template;
        public object Area;
        public object lParamCount;
        public object CurrProject;
        public object CurrModule;
        public object CurrMember;
        public object sT;
        public object sVar1;
        public object sVar2;
        public object sVar3;
        public object ProcType;
        public CInsertionInfo X;
        public object fh;
        public object sMessage;


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
        // 
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
            // ON ERROR RESUME NEXT
        };
            }

            public void Class_Terminate()
            {
                MySadCommands = null;
            }

            public As ISadAddin_ExecuteExternal            {
                //    ';
            }

            public As ISadAddin_ExecuteSoftCommand            {
                // TODO: Rewrite try/catch and/or goto. EH_SADAddin_ISadAddin_ExecuteSoftCommand;




                if ( ! MySadCommands(.SoftCommandName + "*C") Is null )
            {;
                if ( MySadCommands(.SoftCommandName + "*C").IsInline )
            {
 return; // ???;

                switch .SoftCommandName.ToUpper();
                Case "SAVEVARS";
                asaSaved = new SliceAndDice.CAssocArray();
                asaSaved.All = II.SoftVars.All;

                Case "RESTOREVARS";
                if ( ! asaSaved Is null )
            {;
                //                        '   MsgBox "You must execute the ~~SaveVars command first (See F1 help) before using the ~~RestoreVars command.";
                //                        'Else;
                asaSaved.KeyValueDelimiter = " saved=";
                II.SoftVars.All = Replace(asaSaved.All + vbNewLine + II.SoftVars.All, "saved saved", "saved");
                .SoftVars.All, "saved saved"): II.SoftVars.All = Replace(.SoftVars.All, "saved saved".Contains("saved"): Loop) {;
                asaSaved.KeyValueDelimiter = " =";
            }

            public As Mangle            {
                try
{;

                sOrig = strIn;
                LenOrig = Len(sOrig);
                bPreviousWasAnUnderscore = true;
                sOut = string.Empty;

                if ( InStr(strIn, "_") > 0 )
            {;
                for(var CurrChar = 1; CurrChar < LenOrig; CurrChar++)  {;
                sChar = strIn, CurrChar.Substring( 1);
                switch sChar;
                Case "0" To "9";
                sOut +=  sChar;
                bPreviousWasAnUnderscore = false;
                ;
                Case "A" To "Z", "a" To "z";
                if ( bPreviousWasAnUnderscore )
            {;
                sOut +=  sChar.ToUpper();
                }
            else
            {;
                sOut +=  LCase(sChar);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As ISadAddin_ExecuteSoftCommandInline            {
                // TODO: Rewrite try/catch and/or goto. ErrorHandler;

                ;

                // 254      Dim Decision                As VbMsgBoxResult;

                //   'Dim sDefault                As string;

                if ( ! MySadCommands(sInlineSoftCommandName + "*I") Is null )
            {;
                if ( ! MySadCommands(sInlineSoftCommandName + "*I").IsInline )
            {
 return; // ???;
                switch sInlineSoftCommandName;
                Case "MODULECONTENTS", "CONTENTS", "MEMBERS";
                sResults = string.Empty;
                foreach( var CurrMember in mvbInst.ActiveCodePane.CodeModule.Members );
                switch CurrMember.Type;
                Case vbext_mt_Const     ':   sResults +=  CurrMember.Name + "=CONSTANT=" + X + "$$$$";
                Case vbext_mt_Variable  ':   sResults +=  CurrMember.Name + "=VARIABLE=" + X + "=" + X + "=" + X + "=$$$$";
                Case }
            else
            {;
                switch CurrMember.Type;
                Case vbext_mt_Event:    sResults +=  CurrMember.Name + "=EVENT==";
                Case vbext_mt_Method:   sResults +=  CurrMember.Name + "=METHOD==";
                Case vbext_mt_Property: sT = mvbInst.ActiveCodePane.CodeModule.ProcOfLine(CurrMember.CodeLocation, lProcType);
                sResults +=  CurrMember.Name + "=PROPERTY=" + Switch(lProcType = 0, "PROC", lProcType = 1, "LET", lProcType = 2, "SET", lProcType = 3, "GET", true, string.Empty) + "=";
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


                //    'If Now() > 36300 Then;
                //    '   MsgBox "The Slice and Dice evaluation has expired. Thank you for participating." + Chr(13) + Chr(13) + "Please download the latest evaluation from: http://www.sliceanddice.com/VB5CodeWalker";
                //    '   return; // ???;
                //    'End If;

                //    'If UCase(oParent.Version) <> UCase(MySadCommands.Attributes("Version")) Then;
                //    '   If MsgBox("Version mismatch:" + Chr(13) + Chr(9) + "SADAddin.NewCommands = " + MySadCommands.Attributes("Version") + Chr(13) + Chr(9) + "Slice and Dice = " + oParent.Version + Chr(13) + Chr(13) + "Continue loading DLL ?", vbYesNo, "*** WARNING - DLL Version mismatch **") = vbNo Then;
                //    '      return; // ???;
                //    '   End If;
                //    'End If;

                Parent = oParent;
                MySadCommands.Parent = oParent;
                ;
                mvbInst = vbInst;


                MySadCommands.All = Parent.sFileContents(Parent.TemplateDatabasePath + "SADAddin.txt");
                ISadAddin_Startup = (ex = 0);

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
                sMessage +=  vbTab + "Sandal:       sadAddin.NewCommands" + vbNewLine;
                if ( Erl <> 0 )
            {
 sMessage == sMessage + vbTab + "Sandal Line:  " + Erl + vbnewLine;
                sMessage +=  vbTab + "Error Number: " + ex + vbNewLine;
                sMessage +=  vbTab + "Description:  " + Err.Description + vbNewLine;

                sMessage +=  vbNewLine + vbTab + "SoftCode Parameters (Resolved):" + vbNewLine;
                sMessage +=  vbTab + vbTab + "sVar1 = " + sVar1 + vbNewLine;
                sMessage +=  vbTab + vbTab + "sVar2 = " + sVar2 + vbNewLine;
                sMessage +=  vbTab + vbTab + "sVar3 = " + sVar3 + vbNewLine;
                sMessage +=  vbTab + vbTab + "sVar4 = " + sVar4 + vbNewLine;
                ;
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
