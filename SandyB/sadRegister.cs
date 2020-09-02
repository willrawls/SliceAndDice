using System;

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

			public As sadGetLicenseKey			{
				try
{;
				sResult = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License", Key);
				if ( Len(sResult) = 0 Or sResult = "Error" )
				{
 sResult = sDefault;
				if ( Left$(sResult, 4) = "EN* " )
				{;
				sResult = sadDecrypt(sResult);
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void ParseKey			{

				rtn = InStr(Keyname, "\") 'return if "\" is contained in the Keyname;

				if ( Left$(Keyname, 5) <> "HKEY_" Or Right$(Keyname, 1) = "\" )
				{
 'if the is a "\" at the end of the Keyname then;
				MsgBox "Incorrect Format:" + vbLf + vbLf + Keyname 'display error to the user;
				return; 'exit the procedure;
				}
				else
				{if ( rtn = 0 )
				{
 'if the Keyname contains no "\";
				Keyhandle = GetMainKeyHandle(Keyname);
				Keyname = "" 'leave Keyname blank;
				}
				else
				{ 'otherwise, Keyname contains "\";
				Keyhandle = GetMainKeyHandle(Left$(Keyname, rtn - 1)) 'seperate the Keyname;
				Keyname = Right$(Keyname, Len(Keyname) - rtn);
			}
			public As sadDecrypt			{
				if ( Len(strIn) = 0 )
				{
 return; // ???;
				if ( Left$(strIn, 3) <> "EN*" )
				{
 return; // ???;

				strIn = Scramble(strIn);
				Do While Len(strIn);
				strOut = strOut + Chr$((255 - Val("+H" + Left$(strIn, 2) + "+")) Mod 255);
				strIn = Mid$(strIn, 3);
				Loop;
				sadDecrypt = strOut;
			}
			public As sadEncrypt			{

				bytArray = StrConv(strIn, vbFromUnicode);
				For CurrByte = 0 To UBound(bytArray);
				if ( bytArray(CurrByte) < 240 )
				{;
				strOut = strOut + Hex$(255 - bytArray(CurrByte));
				}
				else
				{;
				strOut = strOut + "0" + Hex$(255 - bytArray(CurrByte));
			}
			public void sadSaveLicenseKey			{
				try
{;
				SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License", Key, sadEncrypt(Value);
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As Scramble			{

				if ( Left$(strIn, 4) = "EN* " )
				{;
				//    Shift = -3;
				strIn = Mid$(strIn, 5);
				// Else;
				//    Shift = 3;
			}
			public As GetListIndex			{
				try
{;
				static nCurItem As Integer;

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
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public  sExtractToken			{
				static strIn As String;
				static strOut As String;
				static nCurrTokenStart As Long;
				static nNextTokenStart As Long;
				static nLenDelim As Long;

				//  Handle the "simple" cases (No delimiter, or token # less than 2);
				nLenDelim = Len(strDelim);
				if ( nToken < 1 Or nLenDelim = 0 )
				{;
				//  Nothing to extract, return nothing;
				return; // ???;
				}
				else
				{if ( nToken = 1 )
				{;
				nCurrTokenStart = InStr(sOrigStr, strDelim);
				if ( nCurrTokenStart > 0 )
				{;
				sExtractToken = Left$(sOrigStr, nCurrTokenStart - 1);
				sOrigStr = Trim(Mid$(sOrigStr, nCurrTokenStart + nLenDelim));
				return; // ???;
				}
				else
				{;
				sExtractToken = sOrigStr;
				sOrigStr = "";
				return; // ???;
			}
			public As bUserSure			{
				bUserSure = (MsgBox(sPrompt, vbYesNo, "ARE YOU SURE ?") = vbYes);
			}
			public As lTokenCount			{
				static iCurTokenLocation As Long ' Character position of the first delimiter string;
				static iTokensSoFar As Long      ' Used to keep track of how many tokens we've counted so far;
				static iDelim As Long            ' Length of the delimiter string;

				iDelim = Len(siDelim);
				if ( iDelim < 1 )
				{;
				//  Empty delimiter strings means only one token equal to the string;
				lTokenCount = 1;
				return; // ???;
				}
				else
				{if ( Len(siAllTokens) = 0 )
				{;
				//  Empty input string means no tokens;
				return; // ???;
				}
				else
				{;
				//  Count the number of tokens;
				iTokensSoFar = 0;
				Do;
				iCurTokenLocation = InStr(siAllTokens, siDelim);
				if ( iCurTokenLocation = 0 )
				{;
				lTokenCount = iTokensSoFar + 1 'Abs(Len(siAllTokens) > 0);
				return; // ???;
			}
			public As sGetToken			{
				static iCurTokenLocation As Long ' Character position of the first delimiter string;
				static nDelim As Long            ' Length of the delimiter string;
				nDelim = Len(sDelim);

				if ( iToken < 1 Or nDelim < 1 )
				{;
				//  Negative or zeroth token or empty delimiter strings mean an empty token;
				return; // ???;
				}
				else
				{if ( iToken = 1 )
				{;
				//  Quickly extract the first token;
				iCurTokenLocation = InStr(siAllTokens, sDelim);
				if ( iCurTokenLocation > 1 )
				{;
				sGetToken = Left$(siAllTokens, iCurTokenLocation - 1);
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
			public As sAfter			{
				static iCurTokenLocation As Long ' Character position of the first delimiter string;
				static nDelim As Long            ' Length of the delimiter string;

				nDelim = Len(sDelim);
				if ( iToken < 1 Or nDelim < 1 )
				{;
				//  Negative or zeroth token or empty delimiter strings mean an empty token;
				sAfter = siAllTokens;
				return; // ???;
				}
				else
				{if ( iToken = 1 )
				{;
				//  Quickly extract the first token;
				iCurTokenLocation = InStr(siAllTokens, sDelim);
				if ( iCurTokenLocation > 1 )
				{;
				sAfter = Mid$(siAllTokens, iCurTokenLocation + nDelim);
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
				sAfter = Mid$(siAllTokens, nDelim + 1);
				return; // ???;
			}
			public As sBefore			{
				static iCurTokenLocation As Long ' Character position of the first delimiter string;
				static nDelim As Long            ' Length of the delimiter string;
				static sReturned As String;

				nDelim = Len(sDelim);
				if ( iToken < 2 Or nDelim < 1 )
				{;
				//  First, Zeroth, or Negative tokens or empty delimiter strings mean an empty string returned;
				sBefore = "";
				return; // ???;
				}
				else
				{if ( iToken = 2 )
				{;
				//  Quickly extract the first token;
				sBefore = sGetToken(siAllTokens, 1, sDelim);
				return; // ???;
				}
				else
				{;
				//  Find the Nth token;
				Do;
				iCurTokenLocation = InStr(siAllTokens, sDelim);
				if ( iCurTokenLocation = 0 Or iToken = 1 )
				{;
				sBefore = sReturned;
				sReturned = "";
				return; // ???;
				}
				else
				{if ( Len(sReturned) = 0 )
				{;
				sReturned = Left$(siAllTokens, iCurTokenLocation - 1);
				}
				else
				{;
				sReturned = sReturned + sDelim + Left$(siAllTokens, iCurTokenLocation - 1);
			}
			public As sExcept			{
				static iCurTokenLocation As Long ' Character position of the first delimiter string;
				static nDelim As Long            ' Length of the delimiter string;
				static sReturned As String;

				nDelim = Len(sDelim);
				if ( iToken < 1 Or nDelim < 1 )
				{;
				//  First, Zeroth, or Negative tokens or empty delimiter strings mean an empty string returned;
				sExcept = siAllTokens;
				return; // ???;
				}
				else
				{if ( iToken = 1 )
				{;
				//  Quickly Return after token 1;
				iCurTokenLocation = InStr(siAllTokens, sDelim);
				if ( iCurTokenLocation = 0 )
				{;
				sExcept = siAllTokens;
				return; // ???;
				}
				else
				{;
				sExcept = Mid$(siAllTokens, iCurTokenLocation + nDelim);
				return; // ???;
			}
			public As sReplace			{

				iFind = Len(sFind);
				iCurFindPos = InStr(sAll, sFind);
				if ( InStr(sReplaceWith, sFind) = 0 )
				{;
				Do While iCurFindPos > 0;
				if ( iCurFindPos > 1 )
				{;
				sAll = Left$(sAll, iCurFindPos - 1) + sReplaceWith + Mid$(sAll, iCurFindPos + iFind);
				}
				else
				{;
				sAll = sReplaceWith + Mid$(sAll, iCurFindPos + iFind);
			}
			public  LoadFormPosition			{

				if ( Len(ProductName) = 0 )
				{;
				ProductName = "Slice and Dice";
				}
				else
				{;
				ProductName = App.ProductName;
			}
			public  SaveFormPosition			{

				if ( Len(ProductName) = 0 )
				{;
				ProductName = "Slice and Dice";
				}
				else
				{;
				ProductName = App.ProductName;
			}
			public As lFindToken			{

				lTokens = lTokenCount(sAllTokens, sDelimiter);

				For l = 1 To lTokens;
				if ( StrComp(UCase$(sGetToken(sAllTokens, l, sDelimiter)), UCase$(sTokenToFind)) = 0 )
				{;
				lFindToken = l;
				return; // ???;
			}
		}
}
