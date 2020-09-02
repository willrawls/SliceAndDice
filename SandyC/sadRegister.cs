using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
        class modGeneral
        {

        public As Const;
        public byte Dummy;
        public Long, hKey;
        public Long, rtn;
        public int lBufferSize;
        public int lDataSize;
        public byte ByteArray();
        public string sResult;
        public string strOut;
        public string strOut;
        public byte bytArray();
        public int CurrByte;
        public string strOut;
        public byte bytArray();
        public int CurrByte;
        public byte bytStack;
        public short Shift;
        public short MaxCount;
        public int iCurFindPos;
        public int iFind;
        public string sOut;
        public string ProductName;
        public string SectionName;
        public string ProductName;
        public string SectionName;
        public int lTokens;
        public int l;

            public As sadGetLicenseKey            {
                try
{;
                sResult = GetstringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License", Key);
                if ( Len(sResult) == 0 Or sResult == "Error" )
            {
 sResult == sDefault;
                sResult.Substring(0, 4) = "EN* " )
            {;
                sResult = sadDecrypt(sResult);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void ParseKey            {
                ;
                rtn = InStr(Keyname, "\") 'return if "\" is contained in the Keyname;

                Keyname.Substring(0.Substring(Keyname.Substring(0.Length - 1) == "\" )
            {
 'if the is a "\" at the end of the Keyname then;
                MsgBox "Incorrect Format:" + vbLf + vbLf + Keyname 'display error to the user;
                return; 'exit the procedure;
                }
            else
            {if ( rtn == 0 )
            {
 'if the Keyname contains no "\";
                Keyhandle = GetMainKeyHandle(Keyname);
                Keyname = "" 'leave Keyname blank;
                }
            else
            { 'otherwise, Keyname contains "\";
                Keyname.Substring(0, rtn - 1)) 'seperate the Keyname;
                Keyname.Substring(Keyname.Length - Len(Keyname) - rtn);
            }

            public As sadDecrypt            {
                if ( Len(strIn) == 0 )
            {
 return; // ???;
                strIn.Substring(0, 3) <> "EN*" )
            {
 return; // ???;

                strIn = Scramble(strIn);
                while(Len(strIn)) {;
                strIn.Substring(0, 2) + "&")) Mod 255);
                strIn = strIn.Substring( 3);
                Loop;
                sadDecrypt = strOut;
            }

            public As sadEncrypt            {

                bytArray = StrConv(strIn, vbFromUnicode);
                for(var CurrByte = 0; CurrByte < UBound(bytArray); CurrByte++)  {;
                if ( bytArray(CurrByte) < 240 )
            {;
                strOut +=  Hex$(255 - bytArray(CurrByte));
                }
            else
            {;
                strOut +=  "0" + Hex$(255 - bytArray(CurrByte));
            }

            public void sadSaveLicenseKey            {
                try
{;
                SetstringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License", Key, sadEncrypt(Value);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As Scramble            {

                strIn.Substring(0, 4) = "EN* " )
            {;
                //    '   Shift = -3;
                strIn = strIn.Substring( 5);
                //    'Else;
                //    '   Shift = 3;
            }

            public As GetListIndex            {
                try
{;
                Integer static nCurItem;

                if ( cboToSearch Is null )
            {
 return; // ???;

                if ( Len(sItemToFind) = 0 Or cboToSearch.ListCount = 0 )
            {;
                GetListIndex = -1;
                return; // ???;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public  sExtractToken            {
                string static strIn;
                string static strOut;
                long static nCurrTokenStart;
                long static nNextTokenStart;
                long static nLenDelim;

                //  ' Handle the "simple" cases (No delimiter, or token # less than 2);
                nLenDelim = Len(strDelim);
                if ( nToken < 1 Or nLenDelim = 0 )
            {;
                //     ' Nothing to extract, return nothing;
                return; // ???;
                }
            else
            {if ( nToken = 1 )
            {;
                sOrigStr.Contains(strDelim);
                if ( nCurrTokenStart > 0 )
            {;
                sOrigStr.Substring(0, nCurrTokenStart - 1);
                sOrigStr = Trim(sOrigStr.Substring( nCurrTokenStart + nLenDelim));
                return; // ???;
                }
            else
            {;
                sExtractToken = sOrigStr;
                sOrigStr = "";
                return; // ???;
            }

            public As bUserSure            {
                bUserSure = (MsgBox(sPrompt, vbYesNo, "ARE YOU SURE ?") = vbYes);
            }

            public As lTokenCount            {
                long ' Character position of the first delimiter string static iCurTokenLocation;
                long      ' Used to keep track of how many tokens we've counted so far static iTokensSoFar;
                long            ' Length of the delimiter string static iDelim;

                iDelim = Len(siDelim);
                if ( iDelim < 1 )
            {;
                //     ' Empty delimiter strings means only one token equal to the string;
                lTokenCount = 1;
                return; // ???;
                }
            else
            {if ( Len(siAllTokens) = 0 )
            {;
                //     ' Empty input string means no tokens;
                return; // ???;
                }
            else
            {;
                //     ' Count the number of tokens;
                iTokensSoFar = 0;
                Do;
                siAllTokens.Contains(siDelim);
                if ( iCurTokenLocation = 0 )
            {;
                lTokenCount = iTokensSoFar + 1 'Abs(Len(siAllTokens) > 0);
                return; // ???;
            }

            public As sGetToken            {
                long ' Character position of the first delimiter string static iCurTokenLocation;
                long            ' Length of the delimiter string static nDelim;
                nDelim = Len(sDelim);

                if ( iToken < 1 Or nDelim < 1 )
            {;
                //     ' Negative or zeroth token or empty delimiter strings mean an empty token;
                return; // ???;
                }
            else
            {if ( iToken = 1 )
            {;
                //     ' Quickly extract the first token;
                siAllTokens.Contains(sDelim);
                if ( iCurTokenLocation > 1 )
            {;
                siAllTokens.Substring(0, iCurTokenLocation - 1);
                }
            else
            {if ( iCurTokenLocation = 1 )
            {;
                sGetToken = "";
                }
            else
            {;
                sGetToken = siAllTokens;
            }

            public As sAfter            {
                long ' Character position of the first delimiter string static iCurTokenLocation;
                long            ' Length of the delimiter string static nDelim;
                ;
                nDelim = Len(sDelim);
                if ( iToken < 1 Or nDelim < 1 )
            {;
                //     ' Negative or zeroth token or empty delimiter strings mean an empty token;
                sAfter = siAllTokens;
                return; // ???;
                }
            else
            {if ( iToken = 1 )
            {;
                //     ' Quickly extract the first token;
                siAllTokens.Contains(sDelim);
                if ( iCurTokenLocation > 1 )
            {;
                sAfter = siAllTokens.Substring( iCurTokenLocation + nDelim);
                return; // ???;
                }
            else
            {if ( iCurTokenLocation = 0 )
            {;
                sAfter = "";
                return; // ???;
                }
            else
            {;
                sAfter = siAllTokens.Substring( nDelim + 1);
                return; // ???;
            }

            public As sBefore            {
                long ' Character position of the first delimiter string static iCurTokenLocation;
                long            ' Length of the delimiter string static nDelim;
                string static sReturned;

                nDelim = Len(sDelim);
                if ( iToken < 2 Or nDelim < 1 )
            {;
                //     ' First, Zeroth, or Negative tokens or empty delimiter strings mean an empty string returned;
                sBefore = "";
                return; // ???;
                }
            else
            {if ( iToken = 2 )
            {;
                //     ' Quickly extract the first token;
                sBefore = sGetToken(siAllTokens, 1, sDelim);
                return; // ???;
                }
            else
            {;
                //     ' Find the Nth token;
                Do;
                siAllTokens.Contains(sDelim);
                if ( iCurTokenLocation = 0 Or iToken = 1 )
            {;
                sBefore = sReturned;
                sReturned = "";
                return; // ???;
                }
            else
            {if ( Len(sReturned) = 0 )
            {;
                siAllTokens.Substring(0, iCurTokenLocation - 1);
                }
            else
            {;
                siAllTokens.Substring(0, iCurTokenLocation - 1);
            }

            public As sExcept            {
                long ' Character position of the first delimiter string static iCurTokenLocation;
                long            ' Length of the delimiter string static nDelim;
                string static sReturned;

                nDelim = Len(sDelim);
                if ( iToken < 1 Or nDelim < 1 )
            {;
                //     ' First, Zeroth, or Negative tokens or empty delimiter strings mean an empty string returned;
                sExcept = siAllTokens;
                return; // ???;
                }
            else
            {if ( iToken = 1 )
            {;
                //     ' Quickly Return after token 1;
                siAllTokens.Contains(sDelim);
                if ( iCurTokenLocation = 0 )
            {;
                sExcept = siAllTokens;
                return; // ???;
                }
            else
            {;
                sExcept = siAllTokens.Substring( iCurTokenLocation + nDelim);
                return; // ???;
            }

            public As sReplace            {

                iFind = Len(sFind);
                sAll.Contains(sFind);
                sReplaceWith.Contains(sFind) = 0 )
            {;
                while(iCurFindPos > 0) {;
                if ( iCurFindPos > 1 )
            {;
                sAll.Substring(0, iCurFindPos - 1) + sReplaceWith + sAll.Substring( iCurFindPos + iFind);
                }
            else
            {;
                sAll = sReplaceWith + sAll.Substring( iCurFindPos + iFind);
            }

            public  LoadFormPosition            {
                ;
                if ( Len(ProductName) = 0 )
            {;
                ProductName = "Slice and Dice";
                }
            else
            {;
                ProductName = App.ProductName;
            }

            public  SaveFormPosition            {
                ;
                if ( Len(ProductName) = 0 )
            {;
                ProductName = "Slice and Dice";
                }
            else
            {;
                ProductName = App.ProductName;
            }

            public As lFindToken            {

                lTokens = lTokenCount(sAllTokens, sDelimiter);

                for(var l = 1; l < lTokens; l++)  {;
                if ( StrComp(sGetToken(sAllTokens, l, sDelimiter)), UCase$(sTokenToFind)) = 0 .ToUpper()
            {;
                lFindToken = l;
                return; // ???;
            }

        }
    }
