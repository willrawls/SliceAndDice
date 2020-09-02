using System;

namespace MetX.SliceAndDice
{
		class modGeneral
		{

			public As CollectFormData			{
				Long lCurrForm;
				Long lCurrField;
				Object CurrForm;
				Object CurrField;
				String sOut;
				try
{;
				For lCurrForm = 0 To wb.Document.Forms.length - 1;
				CurrForm = wb.Document.Forms(lCurrForm);
				For lCurrField = 0 To CurrForm.length - 1;
				CurrField = CurrForm(lCurrField);
				sOut = sOut + lCurrForm + "." + CurrField.Name + "=" + CurrField.Value + "+";
				};
				};
				CurrField = null;
				CurrForm = null;

				CollectFormData = sOut;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As ListToString			{
				Long CurrItem;
				String sOut;

				if ( bMoveBackward )
				{;
				For CurrItem = lstToRead.ListCount - 1 To 0 Step -1;
				sOut = sOut + lstToRead.List(CurrItem) + sDelimiter;
				};
				}
				else
				{;
				For CurrItem = 0 To lstToRead.ListCount - 1;
				sOut = sOut + lstToRead.List(CurrItem) + sDelimiter;
				};
				};

				ListToString = sOut;
			}
			public  LoadFormPosition			{
				String ProductName;

				if ( Len(ProductName) = 0 )
				{;
				ProductName = "Your Product Name Here";
				}
				else
				{;
				ProductName = App.ProductName;
				};

				if ( GetSetting(ProductName, frmToActOn.Name, "Position Saved", false) )
				{;
				frmToActOn.Left = GetSetting(ProductName, frmToActOn.Name, "Form Position Left", frmToActOn.Left);
				frmToActOn.Top = GetSetting(ProductName, frmToActOn.Name, "Form Position Top", frmToActOn.Top);
				frmToActOn.Width = GetSetting(ProductName, frmToActOn.Name, "Form Position Width", frmToActOn.Width);
				frmToActOn.Height = GetSetting(ProductName, frmToActOn.Name, "Form Position Height", frmToActOn.Height);
				}
				else
				{if ( bAutoCenter )
				{;
				frmToActOn.Left = (Screen.Width - frmToActOn.Width) / 2;
				frmToActOn.Top = (Screen.Height - frmToActOn.Height) / 2;
				};
			}
			public void Main()
			{
				;
			}
			public  SaveFormPosition			{
				String ProductName;

				if ( Len(ProductName) = 0 )
				{;
				ProductName = "Your Product Name Here";
				}
				else
				{;
				ProductName = App.ProductName;
				};

				SaveSetting ProductName, frmToActOn.Name, "Position Saved", true;
				SaveSetting ProductName, frmToActOn.Name, "Form Position Left", frmToActOn.Left;
				SaveSetting ProductName, frmToActOn.Name, "Form Position Top", frmToActOn.Top;
				SaveSetting ProductName, frmToActOn.Name, "Form Position Width", frmToActOn.Width;
				SaveSetting ProductName, frmToActOn.Name, "Form Position Height", frmToActOn.Height;
			}
			public void StringToList			{
				Long CurrItem;
				String sEntry;

				if ( bClearFirst )
				{
 lstToFill.Clear;
				Do While InStr(1, sContents, sDelimiter, vbTextCompare);
				sEntry = Left$(sContents, InStr(1, sContents, sDelimiter, vbTextCompare) - 1);
				lstToFill.AddItem sEntry;
				sContents = Mid$(sContents, InStr(1, sContents, sDelimiter, vbTextCompare) + Len(sDelimiter));
				Loop;
			}
			public As LogError			{
				Long fh;
				String sMessage;

				fh = FreeFile;
				Open "ERRORLOG.TXT" For Append As #fh;
				sMessage = "***** Error " + Format(lError, "00000") + " at: " + Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM");
				sMessage = sMessage + Chr(13) + "  *** Module:         " + sModuleName;
				sMessage = sMessage + Chr(13) + "  *** Procedure:      " + sProcName;
				sMessage = sMessage + Chr(13) + "  *** Description:    " + sErrorMsg;
				Print #fh, sMessage;
				sMessage = sMessage + Chr(13) + Chr(13) + Chr(9) + "Continue after error ? (No to exit program)";
				if ( MsgBox(sMessage, vbYesNo) = vbNo )
				{;
				Print #fh, "  *** Program shut down by user after error.";
				ShutDownNicely;
				}
				else
				{;
				Print #fh, "  *** Program continued by user after error.";
				};
				Close #fh;

			}
			public void ShutDownNicely()
			{
				try
{;
				//  Close all objects, forms, handles, etc. here;
				if ( frmBrowser.Visible )
				{
 frmBrowser.Hide;
				// If frmMain.Visible Then frmMain.Hide;
				// frmMain.IEBrowsers.Clear;

				// End;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As sGetToken			{
				Long iCurTokenLocation;
				Integer nDelim;
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
				iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare);
				if ( iCurTokenLocation > 1 )
				{;
				sGetToken = Left$(sAllTokens, iCurTokenLocation - 1);
				}
				else
				{if ( iCurTokenLocation = 1 )
				{;
				sGetToken = "";
				}
				else
				{;
				sGetToken = sAllTokens;
				};
				return; // ???;
				}
				else
				{;
				//  Find the Nth token;
				Do;
				iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare);
				if ( iCurTokenLocation = 0 )
				{;
				return; // ???;
				}
				else
				{;
				sAllTokens = Mid$(sAllTokens, iCurTokenLocation + nDelim);
				};
				iToken = iToken - 1;
				Loop Until iToken = 1;

				//  Extract the Nth token (Which is the next token at this point);
				iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare);
				if ( iCurTokenLocation > 0 )
				{;
				sGetToken = Left$(sAllTokens, iCurTokenLocation - 1);
				return; // ???;
				}
				else
				{;
				sGetToken = sAllTokens;
				return; // ???;
				};
				};
			}
			public As sAfter			{
				Long iCurTokenLocation;
				Integer nDelim;

				nDelim = Len(sDelim);
				if ( iToken < 1 Or nDelim < 1 )
				{;
				//  Negative or zeroth token or empty delimiter strings mean an empty token;
				sAfter = sAllTokens;
				return; // ???;
				}
				else
				{if ( iToken = 1 )
				{;
				//  Quickly extract the first token;
				iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare);
				if ( iCurTokenLocation > 1 )
				{;
				sAfter = Mid$(sAllTokens, iCurTokenLocation + nDelim);
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
				sAfter = Mid$(sAllTokens, nDelim + 1);
				return; // ???;
				};
				}
				else
				{;
				//  Find the Nth token;
				Do;
				iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare);
				if ( iCurTokenLocation = 0 )
				{;
				return; // ???;
				}
				else
				{;
				sAllTokens = Mid$(sAllTokens, iCurTokenLocation + nDelim);
				};
				iToken = iToken - 1;
				Loop Until iToken = 1;

				//  Extract the Nth token (Which is the next token at this point);
				iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare);
				if ( iCurTokenLocation > 0 )
				{;
				sAfter = Mid$(sAllTokens, iCurTokenLocation + nDelim);
				return; // ???;
				}
				else
				{;
				return; // ???;
				};
				};
			}
			public As sBefore			{
				Long iCurTokenLocation;
				Integer nDelim;
				String sReturned;

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
				sBefore = sGetToken(sAllTokens, 1, sDelim);
				return; // ???;
				}
				else
				{;
				//  Find the Nth token;
				Do;
				iCurTokenLocation = InStr(1, sAllTokens, sDelim, vbTextCompare);
				if ( iCurTokenLocation = 0 Or iToken = 1 )
				{;
				sBefore = sReturned;
				sReturned = "";
				return; // ???;
				}
				else
				{if ( Len(sReturned) = 0 )
				{;
				sReturned = Left$(sAllTokens, iCurTokenLocation - 1);
				}
				else
				{;
				sReturned = sReturned + sDelim + Left$(sAllTokens, iCurTokenLocation - 1);
				};
				sAllTokens = Mid$(sAllTokens, iCurTokenLocation + nDelim);
				iToken = iToken - 1;
				Loop;
				};
			}
			public As lFindToken			{
				Long lTokens;
				Long l;

				lTokens = lTokenCount(sAllTokens, sDelimiter);

				For l = 1 To lTokens;
				if ( StrComp(UCase$(Left$(sGetToken(sAllTokens, l, sDelimiter), Len(sTokenToFind))), UCase$(sTokenToFind)) = 0 )
				{;
				lFindToken = l;
				return; // ???;
				};
				Next;

				lFindToken = 0;
			}
			public As lTokenCount			{
				Long iCurTokenLocation;
				Integer iTokensSoFar;
				Integer iDelim;

				iDelim = Len(siDelim);
				if ( iDelim < 1 )
				{;
				//  Empty delimiter strings means only one token equal to the string;
				lTokenCount = 1;
				return; // ???;
				}
				else
				{if ( Len(sAllTokens) = 0 )
				{;
				//  Empty input string means no tokens;
				return; // ???;
				}
				else
				{;
				//  Count the number of tokens;
				iTokensSoFar = 0;
				Do;
				iCurTokenLocation = InStr(1, sAllTokens, siDelim, vbTextCompare);
				if ( iCurTokenLocation = 0 )
				{;
				lTokenCount = iTokensSoFar + 1 'Abs(Len(sAllTokens) > 0);
				return; // ???;
				};
				iTokensSoFar = iTokensSoFar + 1;
				sAllTokens = Mid$(sAllTokens, iCurTokenLocation + iDelim);
				Loop;
				};
			}
			public As bUserSure			{
				bUserSure = (MsgBox(sPrompt, vbYesNo, "ARE YOU SURE ?") = vbYes);
			}
		}
}
