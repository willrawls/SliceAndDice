using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
        class modGeneral
        {

        public As Const;
        public As Const;
        public As Const;
        public As Const;

            public void ExtendListView            {
                ;
                ;

                style = SendMessage(hWndListView, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0);
                style += r LVS_EX_FULLROWSELECT;
                lReturned = SendMessage(hWndListView, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, style);
            }

            public As iTokenCount            {
                long ' Character position of the first delimiter string     static iCurTokenLocation;
                Integer      ' Used to keep track of how many tokens we've counted so far     static iTokensSoFar;
                Integer            ' Length of the delimiter string     static iDelim;

                iDelim = Len(siDelim);
                if ( iDelim < 1 )
            {;
                //     ' Empty delimiter strings means only one token equal to the string;
                iTokenCount = 1;
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
                iTokenCount = iTokensSoFar + 1 'Abs(Len(siAllTokens) > 0);
                return; // ???;
                };
                iTokensSoFar +=  1;
                siAllTokens = siAllTokens.Substring( iCurTokenLocation + iDelim);
                Loop;
                };
            }

            public As sAfter            {
                long ' Character position of the first delimiter string     static iCurTokenLocation;
                Integer            ' Length of the delimiter string     static nDelim;
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
                };
                }
            else
            {;
                //     ' Find the Nth token;
                Do;
                siAllTokens.Contains(sDelim);
                if ( iCurTokenLocation = 0 )
            {;
                return; // ???;
                }
            else
            {;
                siAllTokens = siAllTokens.Substring( iCurTokenLocation + nDelim);
                };
                iToken +=  1;
                Loop Until iToken = 1;

                //     ' Extract the Nth token (Which is the next token at this point);
                siAllTokens.Contains(sDelim);
                if ( iCurTokenLocation > 0 )
            {;
                sAfter = siAllTokens.Substring( iCurTokenLocation + nDelim);
                return; // ???;
                }
            else
            {;
                return; // ???;
                };
                };
            }

            public As sGetToken            {
                long ' Character position of the first delimiter string     static iCurTokenLocation;
                Integer            ' Length of the delimiter string     static nDelim;
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
                };
                return; // ???;
                }
            else
            {;
                //     ' Find the Nth token;
                Do;
                siAllTokens.Contains(sDelim);
                if ( iCurTokenLocation = 0 )
            {;
                return; // ???;
                }
            else
            {;
                siAllTokens = siAllTokens.Substring( iCurTokenLocation + nDelim);
                };
                iToken +=  1;
                Loop Until iToken = 1;

                //     ' Extract the Nth token (Which is the next token at this point);
                siAllTokens.Contains(sDelim);
                if ( iCurTokenLocation > 0 )
            {;
                siAllTokens.Substring(0, iCurTokenLocation - 1);
                return; // ???;
                }
            else
            {;
                sGetToken = siAllTokens;
                return; // ???;
                };
                };
            }

            public void SetListIndex            {
                cboToSearch.ListIndex = GetListIndex(cboToSearch, sItemToFind);
            }

            public As GetListIndex            {
                Integer     static nCurItem;

                if ( Len(sItemToFind) = 0 Or cboToSearch.ListCount = 0 )
            {;
                GetListIndex = -1;
                return; // ???;
                };

                sItemToFind = sItemToFind.ToUpper();

                for(var nCurItem = 0; nCurItem < cboToSearch.ListCount - 1; nCurItem++)  {;
                if ( StrComp(cboToSearch.List(nCurItem)), sItemToFind) = 0 .ToUpper()
            {;
                GetListIndex = nCurItem;
                return; // ???;
                };
                } // nCurItem;

            }

            public As sNormalize            {
                sNormalize = sReplace(sReplace(sLine, Chr$(13) + Chr$(10), "%$%EOL%$%"), Chr$(9), "%$%TAB%$%");
            }

            public As sReplace            {
                ;
                ;

                iFind = Len(sFind);
                sAll.Contains(sFind);
                while(iCurFindPos > 0) {;
                if ( iCurFindPos > 1 )
            {;
                sAll.Substring(0, iCurFindPos - 1) + sReplaceWith + sAll.Substring( iCurFindPos + iFind);
                }
            else
            {;
                sAll = sReplaceWith + sAll.Substring( iCurFindPos + iFind);
                };
                sAll.Contains(sFind);
                Loop;
                sReplace = sAll;
            }

            public As sDenormalize            {
                sDenormalize = sReplace(sReplace(sLine, "%$%EOL%$%", Chr$(13) + Chr$(10)), "%$%TAB%$%", Chr$(9));
            }

            public As sBefore            {
                long ' Character position of the first delimiter string     static iCurTokenLocation;
                Integer            ' Length of the delimiter string     static nDelim;
                string     static sReturned;

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
                };
                siAllTokens = siAllTokens.Substring( iCurTokenLocation + nDelim);
                iToken +=  1;
                Loop;
                };
            }

        }
    }
