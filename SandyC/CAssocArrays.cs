using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
        public class CAssocArrays
        {

        public Dictionary<string,string>() mCol;


                public CAssocArray Item
    {
        get
        {
        // TODO: Rewrite try/catch and/or goto. EH_CAssocArrays_Item
         Item = mCol(sIndexKey);
        EH_CAssocArrays_Item_Continue:
        return; // ???
        EH_CAssocArrays_Item:
         Item = Add(                                                                                                                                                                                                                                                                                                         sIndexKey);
        goto EH_CAssocArrays_Item_Continue;
        }

    }


                public int Count
    {
        get
        {
        Count = mCol.Count;
        }

    }


                public IUnknown NewEnum
    {
        get
        {
         NewEnum = mCol.[_NewEnum];
        }

    }



            public As Add            {
                ;

                if ( Len(sKey) = 0 )
            {;
                Err.Raise vbObjectError + 0, "CAssocArrays", "Tryed to Add(                                                                                                   an Assoc Array without a key.");
                };

                objNewMember = new CAssocArray();
                objNewMember.Section = sKey;
                mCol.Add(                                                                                                   objNewMember, sKey);
                Add(                                                                                                   = objNewMember);
                objNewMember = null;
            }

            public void Clear()
            {
                mCol = null;
                mCol = new Dictionary<string,string>();
            }

            public void Load            {
                if ( Len(sFilename) = 0 )
            {;
                Err.Raise vbObjectError + 2, "CAssocArrays_LoadAll", "Tryed to load w/o a filename.";
                };

                ;
                ;
                ;

                fh = FreeFile;

                if ( bClearFirst = true )
            {;
                Clear;
                };
                ;
                #fh     Open sFilename For Input Access Read;
                Do Until EOF(fh);
                Line Input #fh, sLine;
                if ( Len(sLine) = 0 )
            {;
                //             ' Skip it;
                }
            else
sLine.Substring(0, 1) = "[" )
            {;
                CurAssocArray = Add(                                                                                                                                                                                                      sLine, 2.Substring( Len(sLine) - 2));

                Line Input #fh, sLine;
                CurAssocArray.ItemDelimiter = sAfter(sDenormalize(sLine), 1, "=");
                Line Input #fh, sLine;
                CurAssocArray.KeyValueDelimiter = sAfter(sDenormalize(sLine), 1, "=");
                Line Input #fh, sLine;
                CurAssocArray.FieldDelimiter = sAfter(sDenormalize(sLine), 1, "=");

                }
            else
            {;
                CurAssocArray.Add(                                                                                                   sGetToken(sLine, 1, "="), sAfter(sLine, 1, "="));
                };
                Loop;
                Close #fh;
                CurAssocArray = null;
            }

            public void Remove            {
                mCol.Remove sIndexKey;
            }

            public void Save            {
                if ( Len(sFilename) = 0 )
            {;
                Err.Raise vbObjectError + 1, "CAssocArrays_SaveAll", "Tryed to save w/o a filename.";
                };

                ;
                ;
                ;
                ;

                fh = FreeFile;
                #fh     Open sFilename For Output Access Write;
                foreach( var CurAssocArray in mCol;

                sOldItem = CurAssocArray.ItemDelimiter;
                sOldKV = CurAssocArray.KeyValueDelimiter;
                Print #fh, "[" + CurAssocArray.Section + "]";
                Print #fh, "Delimiter Item=" + sNormalize(.ItemDelimiter);
                Print #fh, "Delimiter Key Value=" + sNormalize(.KeyValueDelimiter);
                Print #fh, "Delimiter Field=" + sNormalize(.FieldDelimiter);
                CurAssocArray.ItemDelimiter = Chr$(13) + Chr$(10);
                CurAssocArray.KeyValueDelimiter = "=";
                Print #fh, CurAssocArray.All;
                CurAssocArray.ItemDelimiter = sOldItem;
                CurAssocArray.KeyValueDelimiter = sOldKV;

                } // CurAssocArray;
                Close #fh;
                CurAssocArray = null;
            }

            public void Class_Initialize()
            {
                Clear;
            }

            public void Class_Terminate()
            {
                mCol = null;
            }

        }
    }
