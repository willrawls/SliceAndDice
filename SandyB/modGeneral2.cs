using System;

namespace MetX.SliceAndDice
{
		class modGeneral
		{

			public As Const;
			public As Const;
			public As Const;
			public As Const;

			public void ExtendListView			{
				Long style;
				Long lReturned;

				style = SendMessage(hWndListView, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0);
				style = style Or LVS_EX_FULLROWSELECT;
				lReturned = SendMessage(hWndListView, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, style);
			}
			public As iTokenCount			{
				static iCurTokenLocation As Long ' Character position of the first delimiter string;
				static iTokensSoFar As Integer      ' Used to keep track of how many tokens we've counted so far;
				static iDelim As Integer            ' Length of the delimiter string;

				iDelim = Len(siDelim);
				if ( iDelim < 1 )
				{;
				//  Empty delimiter strings means only one token equal to the string;
				iTokenCount = 1;
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
				iTokenCount = iTokensSoFar + 1 'Abs(Len(siAllTokens) > 0);
				return; // ???;
				};
				iTokensSoFar = iTokensSoFar + 1;
				siAllTokens = Mid$(siAllTokens, iCurTokenLocation + iDelim);
				Loop;
				};
			}
			public As sAfter			{
				static iCurTokenLocation As Long ' Character position of the first delimiter string;
				static nDelim As Integer            ' Length of the delimiter string;

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
				};
				}
				else
				{;
				//  Find the Nth token;
				Do;
				iCurTokenLocation = InStr(siAllTokens, sDelim);
				if ( iCurTokenLocation = 0 )
				{;
				return; // ???;
				}
				else
				{;
				siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim);
				};
				iToken = iToken - 1;
				Loop Until iToken = 1;

				//  Extract the Nth token (Which is the next token at this point);
				iCurTokenLocation = InStr(siAllTokens, sDelim);
				if ( iCurTokenLocation > 0 )
				{;
				sAfter = Mid$(siAllTokens, iCurTokenLocation + nDelim);
				return; // ???;
				}
				else
				{;
				return; // ???;
				};
				};
			}
			public As sGetToken			{
				static iCurTokenLocation As Long ' Character position of the first delimiter string;
				static nDelim As Integer            ' Length of the delimiter string;
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
				};
				return; // ???;
				}
				else
				{;
				//  Find the Nth token;
				Do;
				iCurTokenLocation = InStr(siAllTokens, sDelim);
				if ( iCurTokenLocation = 0 )
				{;
				return; // ???;
				}
				else
				{;
				siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim);
				};
				iToken = iToken - 1;
				Loop Until iToken = 1;

				//  Extract the Nth token (Which is the next token at this point);
				iCurTokenLocation = InStr(siAllTokens, sDelim);
				if ( iCurTokenLocation > 0 )
				{;
				sGetToken = Left$(siAllTokens, iCurTokenLocation - 1);
				return; // ???;
				}
				else
				{;
				sGetToken = siAllTokens;
				return; // ???;
				};
				};
			}
			public void SetListIndex			{
				cboToSearch.ListIndex = GetListIndex(cboToSearch, sItemToFind);
			}
			public As GetListIndex			{
				static nCurItem As Integer;

				if ( Len(sItemToFind) = 0 Or cboToSearch.ListCount = 0 )
				{;
				GetListIndex = -1;
				return; // ???;
				};

				sItemToFind = UCase$(sItemToFind);

				For nCurItem = 0 To cboToSearch.ListCount - 1;
				if ( StrComp(UCase$(cboToSearch.List(nCurItem)), sItemToFind) = 0 )
				{;
				GetListIndex = nCurItem;
				return; // ???;
				};
				};

			}
			public As sNormalize			{
				sNormalize = sReplace(sReplace(sLine, Chr$(13) + Chr$(10), "%$%EOL%$%"), Chr$(9), "%$%TAB%$%");
			}
			public As sReplace			{
				Long iCurFindPos;
				Integer iFind;

				iFind = Len(sFind);
				iCurFindPos = InStr(sAll, sFind);
				Do While iCurFindPos > 0;
				if ( iCurFindPos > 1 )
				{;
				sAll = Left$(sAll, iCurFindPos - 1) + sReplaceWith + Mid$(sAll, iCurFindPos + iFind);
				}
				else
				{;
				sAll = sReplaceWith + Mid$(sAll, iCurFindPos + iFind);
				};
				iCurFindPos = InStr(sAll, sFind);
				Loop;
				sReplace = sAll;
			}
			public As sDenormalize			{
				sDenormalize = sReplace(sReplace(sLine, "%$%EOL%$%", Chr$(13) + Chr$(10)), "%$%TAB%$%", Chr$(9));
			}
			public As sBefore			{
				static iCurTokenLocation As Long ' Character position of the first delimiter string;
				static nDelim As Integer            ' Length of the delimiter string;
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
				};
				siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim);
				iToken = iToken - 1;
				Loop;
				};
			}
		}
}
