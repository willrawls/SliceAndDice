using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
        class modGeneral
        {
            public As CollectFormData            {
                ;
                ;
                ;
                ;
                ;
                try
{;
                for(var lCurrForm = 0; lCurrForm < wb.Document.Forms.length - 1; lCurrForm++)  {;
                CurrForm = wb.Document.Forms(lCurrForm);
                for(var lCurrField = 0; lCurrField < CurrForm.length - 1; lCurrField++)  {;
                CurrField = CurrForm(lCurrField);
                sOut +=  lCurrForm + "." + CurrField.Name + "=" + CurrField.Value + "&";
                } // lCurrField;
                } // lCurrForm;
                CurrField = null;
                CurrForm = null;
                ;
                CollectFormData = sOut;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As ListTostring            {
                ;
                ;

                if ( bMoveBackward )
            {;
                for(var CurrItem = lstToRead.ListCount - 1; CurrItem < 0 Step -1; CurrItem++)  {;
                sOut +=  lstToRead.List(CurrItem) + sDelimiter;
                } // CurrItem;
                }
            else
            {;
                for(var CurrItem = 0; CurrItem < lstToRead.ListCount - 1; CurrItem++)  {;
                sOut +=  lstToRead.List(CurrItem) + sDelimiter;
                } // CurrItem;
                };

                ListTostring = sOut;
            }

            public  LoadFormPosition            {
                ;
                ;
                if ( Len(ProductName) = 0 )
            {;
                ProductName = "Your Product Name Here";
                }
            else
            {;
                ProductName = App.ProductName;
                };
                ;
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

            }

            public  SaveFormPosition            {
                ;
                ;
                if ( Len(ProductName) = 0 )
            {;
                ProductName = "Your Product Name Here";
                }
            else
            {;
                ProductName = App.ProductName;
                };
                ;
                SaveSetting ProductName, frmToActOn.Name, "Position Saved", true;
                SaveSetting ProductName, frmToActOn.Name, "Form Position Left", frmToActOn.Left;
                SaveSetting ProductName, frmToActOn.Name, "Form Position Top", frmToActOn.Top;
                SaveSetting ProductName, frmToActOn.Name, "Form Position Width", frmToActOn.Width;
                SaveSetting ProductName, frmToActOn.Name, "Form Position Height", frmToActOn.Height;
            }

            public void stringToList            {
                ;
                ;

                if ( bClearFirst )
            {
 lstToFill.Clear;
                1, sContents, sDelimiter.Contains(vbTextCompare)) {;
                1, sContents, sDelimiter.Contains(vbTextCompare) - 1);
                lstToFill.AddItem(sEntry);
                1, sContents.Contains(sDelimiter.Substring( vbTextCompare) + Len(sDelimiter));
                Loop;
            }

            public As LogError            {
                ;
                ;

                fh = FreeFile;
                #fh     Open "ERRORLOG.TXT" For Append;
                sMessage = "***** Error " + Format(lError, "00000") + " at: " + Format(Now(), "MM/DD/YYYY HH:MM:SS AM/PM");
                sMessage +=  Chr(13) + "  *** Module:         " + sModuleName;
                sMessage +=  Chr(13) + "  *** Procedure:      " + sProcName;
                sMessage +=  Chr(13) + "  *** Description:    " + sErrorMsg;
                Print #fh, sMessage;
                sMessage +=  Chr(13) + Chr(13) + Chr(9) + "Continue after error ? (No to exit program)";
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
                //  ' Close all objects, forms, handles, etc. here;
                if ( frmBrowser.Visible )
            {
 frmBrowser.Hide;
                //   'If frmMain.Visible Then frmMain.Hide;
                //   'frmMain.IEBrowsers.Clear;
                ;
                //    'End;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As sGetToken            {
                ;
                ;
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
                1, sAllTokens, sDelim.Contains(vbTextCompare);
                if ( iCurTokenLocation > 1 )
            {;
                sAllTokens.Substring(0, iCurTokenLocation - 1);
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
                //     ' Find the Nth token;
                Do;
                1, sAllTokens, sDelim.Contains(vbTextCompare);
                if ( iCurTokenLocation = 0 )
            {;
                return; // ???;
                }
            else
            {;
                sAllTokens = sAllTokens.Substring( iCurTokenLocation + nDelim);
                };
                iToken +=  1;
                Loop Until iToken = 1;

                //     ' Extract the Nth token (Which is the next token at this point);
                1, sAllTokens, sDelim.Contains(vbTextCompare);
                if ( iCurTokenLocation > 0 )
            {;
                sAllTokens.Substring(0, iCurTokenLocation - 1);
                return; // ???;
                }
            else
            {;
                sGetToken = sAllTokens;
                return; // ???;
                };
                };
            }

            public As sAfter            {
                ;
                ;
                ;
                nDelim = Len(sDelim);
                if ( iToken < 1 Or nDelim < 1 )
            {;
                //     ' Negative or zeroth token or empty delimiter strings mean an empty token;
                sAfter = sAllTokens;
                return; // ???;
                }
            else
            {if ( iToken = 1 )
            {;
                //     ' Quickly extract the first token;
                1, sAllTokens, sDelim.Contains(vbTextCompare);
                if ( iCurTokenLocation > 1 )
            {;
                sAfter = sAllTokens.Substring( iCurTokenLocation + nDelim);
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
                sAfter = sAllTokens.Substring( nDelim + 1);
                return; // ???;
                };
                }
            else
            {;
                //     ' Find the Nth token;
                Do;
                1, sAllTokens, sDelim.Contains(vbTextCompare);
                if ( iCurTokenLocation = 0 )
            {;
                return; // ???;
                }
            else
            {;
                sAllTokens = sAllTokens.Substring( iCurTokenLocation + nDelim);
                };
                iToken +=  1;
                Loop Until iToken = 1;

                //     ' Extract the Nth token (Which is the next token at this point);
                1, sAllTokens, sDelim.Contains(vbTextCompare);
                if ( iCurTokenLocation > 0 )
            {;
                sAfter = sAllTokens.Substring( iCurTokenLocation + nDelim);
                return; // ???;
                }
            else
            {;
                return; // ???;
                };
                };
            }

            public As sBefore            {
                ;
                ;
                ;

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
                sBefore = sGetToken(sAllTokens, 1, sDelim);
                return; // ???;
                }
            else
            {;
                //     ' Find the Nth token;
                Do;
                1, sAllTokens, sDelim.Contains(vbTextCompare);
                if ( iCurTokenLocation = 0 Or iToken = 1 )
            {;
                sBefore = sReturned;
                sReturned = "";
                return; // ???;
                }
            else
            {if ( Len(sReturned) = 0 )
            {;
                sAllTokens.Substring(0, iCurTokenLocation - 1);
                }
            else
            {;
                sAllTokens.Substring(0, iCurTokenLocation - 1);
                };
                sAllTokens = sAllTokens.Substring( iCurTokenLocation + nDelim);
                iToken +=  1;
                Loop;
                };
            }

            public As lFindToken            {
                ;
                ;

                lTokens = lTokenCount(sAllTokens, sDelimiter);

                for(var l = 1; l < lTokens; l++)  {;
                sGetToken(sAllTokens, l, sDelimiter), Len(sTokenToFind))).Substring(0, sTokenToFind)) = 0 .ToUpper()
            {;
                lFindToken = l;
                return; // ???;
                };
                Next;

                lFindToken = 0;
            }

            public As lTokenCount            {
                ;
                ;
                ;

                iDelim = Len(siDelim);
                if ( iDelim < 1 )
            {;
                //     ' Empty delimiter strings means only one token equal to the string;
                lTokenCount = 1;
                return; // ???;
                }
            else
            {if ( Len(sAllTokens) = 0 )
            {;
                //     ' Empty input string means no tokens;
                return; // ???;
                }
            else
            {;
                //     ' Count the number of tokens;
                iTokensSoFar = 0;
                Do;
                1, sAllTokens, siDelim.Contains(vbTextCompare);
                if ( iCurTokenLocation = 0 )
            {;
                lTokenCount = iTokensSoFar + 1 'Abs(Len(sAllTokens) > 0);
                return; // ???;
                };
                iTokensSoFar +=  1;
                sAllTokens = sAllTokens.Substring( iCurTokenLocation + iDelim);
                Loop;
                };
            }

            public As bUserSure            {
                bUserSure = (MsgBox(sPrompt, vbYesNo, "ARE YOU SURE ?") = vbYes);
            }

        }
    }
