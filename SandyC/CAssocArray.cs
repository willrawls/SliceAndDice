using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
        public class CAssocArray
        {

        public object mCol;
        public object CurItem;
        public object Section;
        public object ItemDelimiter;
        public string KeyValueDelimiter;
        public object FieldDelimiter;


            
    /*
        ********************************************************************************;
        Class Module      CAssocArray;
        ';
        Filename          CAssocArray.cls;
        ';
        Copyright         1998 by Firm Solutions;
                          All rights reserved.;
        ';
        Author            William M. Rawls;
                          Firm Solutions;
        ';
        Created On        4/30/1998 1:23 PM;
        ';
        Description;
        ';
           The Reality Matrix, Dimention 2 of 3;
              "Associative array" like abilities;
        ';
           What's "assosiative array" like abilities mean ?;
              Email = wrawls@firmsolutions.com to find out more.;
           Web page you ask ?;
              URL = http://www.firmsolutions.com/RealityMatrix.html
           Why does this read like an e-mail ?;
              Because = It does;
        ';
        Revisions;
        ';
        <RevisionDate>, <RevisedBy>;
        <Description of Revision>;
        ';
        4/30/1998, William M. Rawls;
freeware. Use at your own risk.         Entered into public domain;
        ';
        ********************************************************************************;
        ;
    */

    public string All
    {
        get
        {
        string static sAllKeyValues;
        foreach( var CurItem in mCol )
        sAllKeyValues +=  CurItem.Key + KeyValueDelimiter + CurItem.Value + ItemDelimiter;
        } // CurItem
        CurItem = null;
        All = sAllKeyValues;
        sAllKeyValues = "";
        }

        set
        {
        string static sT;
        string static sKey;
        string static sValue;
        Clear;
        while(Len(value) > 0) {
        sT = sGetToken(value, 1, ItemDelimiter);
        sT.Contains(KeyValueDelimiter) > 0 )
            {

        sKey = sGetToken(sT, 1, KeyValueDelimiter);
        sValue = sGetToken(sT, 2, KeyValueDelimiter);
        Add(                                                                                                                                                                                                      sKey, sValue);
        }
            else
            {
        Add(                                                                                                                                                                                                      sT);
        }
        value = sAfter(value, 1, ItemDelimiter);
        Loop;
        sT = "";
        sKey = "";
        sValue = "";
        }

    }


                public CAssocItem Item
    {
        get
        {
        // TODO: Rewrite try/catch and/or goto. EH_CAssocArray_Item
         Item = mCol(sIndexKey);
        EH_CAssocArray_Item_Continue:
        return; // ???
        EH_CAssocArray_Item:
         Item = Add(                                                                                                                                                                                                                                                                                                         sIndexKey);
        goto EH_CAssocArray_Item_Continue;
        }

    }


                public int Count
    {
        get
        {
        Count = mCol.Count;
        }

    }


                public object Keys
    {
        get
        {
        string static sKeys;
        string sDelimiter ;
        sDelimiter = ItemDelimiter;
        foreach( var CurItem in mCol )
        sKeys +=  sDelimiter + CurItem.Key;
        } // CurItem
        Keys = sKeys.Substring( 2);
        sKeys = "";
        }

        set
        {
        long TokenCount ;
        long CurToken ;
        string sCurToken ;
        this.Clear;
        TokenCount = iTokenCount(value, this.ItemDelimiter);
        for(var CurToken = 1; CurToken < TokenCount; CurToken++)  {
        this.Add(                                                                                                                                                                                                      sGetToken(value, CurToken, ItemDelimiter));
        } // CurToken
        }

    }


                public IUnknown NewEnum
    {
        get
        {
         NewEnum = mCol.[_NewEnum];
        }

    }



            public void Clear()
            {
                mCol = null;
                mCol = new Dictionary<string,string>();
            }

            public As Add            {
                CurItem = new CAssocItem();

                Key.Contains(this.KeyValueDelimiter) > 0 )
            {;
                CurItem.Key = sGetToken(Key, 1, this.KeyValueDelimiter);
                CurItem.Value = sAfter(Key, 1, this.KeyValueDelimiter);
                }
            else
            {;
                CurItem.Key = Key;
                CurItem.Value = Value;
                };

                mCol.Add(                                                                                                   CurItem, Key);
                Add(                                                                                                   = CurItem);
                CurItem = null;
            }

            public As Column            {
                string     static sColumn;

                if ( Len(sDelimiter) == 0 )
            {
 sDelimiter == ItemDelimiter;

                foreach( var CurItem in mCol;
                sColumn +=  sDelimiter + CurItem.F(iCol, FieldDelimiter);
                } // CurItem;
                CurItem = null;

                Column = sColumn.Substring( 2);
                sColumn = "";
            }

            public void FillList            {
                if ( bClearList )
            {
 lstCtrl.Clear;
                ;
                if ( iColumn = 0 )
            {;
                foreach( var CurItem in mCol;
                lstCtrl.AddItem(CurItem.Key);
                } // CurItem;
                }
            else
            {;
                foreach( var CurItem in mCol;
                lstCtrl.AddItem(CurItem.F(CLng(iColumn), FieldDelimiter));
                } // CurItem;
                };
                CurItem = null;
                ;
                if ( StrComp(sItemToChoose, "*NONE*") <> 0 )
            {;
                SetListIndex(lstCtrl, sItemToChoose);
                };
            }

            public void FillListView            {
                ListView     static lvwX;
                ListItem     static NewItem;
                Integer     static SubItems;
                Integer     static CurSubItem;
                string     static sT;

                lvwX = lvwCtrl;
                lvwX.ListItems.Clear;
                if ( bFullLine )
            {
 ExtendListView lvwX.hWnd;

                foreach( var CurItem in mCol;

                SubItems = iTokenCount(.Value, FieldDelimiter);
                if ( SubItems > 0 )
            {;
                sT = sGetToken(.Value, 1, FieldDelimiter);
                NewItem = lvwX.ListItems.Add(                                                                                                                                                                                                      , "Key " + CurItem.Key, CurItem.Key) ', sT, sT);
                for(var CurSubItem = 2; CurSubItem < SubItems; CurSubItem++)  {;
                NewItem.SubItems(CurSubItem - 1) = sGetToken(.Value, CurSubItem, FieldDelimiter);
                } // CurSubItem;
                }
            else
            {;
                NewItem = lvwX.ListItems.Add(                                                                                                                                                                                                      , CurItem.Key, CurItem.Key);
                };

                } // CurItem;
                CurItem = null;
                NewItem = null;
                lvwX = null;
                sT = "";
            }

            public void FillListViewColumns            {
                ListView     static lvwX;
                lvwX = lvwCtrl;

                lvwX.ColumnHeaders.Clear;
                foreach( var CurItem in mCol;
                lvwX.ColumnHeaders.Add(                                                                                                   , CurItem.Key, CurItem.Key, Val(CurItem.Value));
                } // CurItem;
                CurItem = null;

                lvwX = null;
            }

            public void FillTreeNode            {
                ;
                ;
                ;
                ;
                ;
                ;
                ;
                ;

                ;
                ;

                nChildDelimiter = Len(ChildDelimiter);
                nEndChildDelimiter = Len(EndChildDelimiter);

                tvwX = tvwToFill;
                nodX = nodCur;
                foreach( var CurItem in mCol;
                sKey = sGetToken(CurItem.Key, 1, IconDelimiter);
                sIcon = sGetToken(CurItem.Key, 2, IconDelimiter);
                sTag = sGetToken(sKey, 2, TagDelimiter);
                sKey = sGetToken(sKey, 1, TagDelimiter);
                if ( Len(sIcon) == 0 )
            {
 sIcon == sImage;

                if ( nChildDelimiter > 0 )
            {;
                sKey.Substring(0, nChildDelimiter) = ChildDelimiter )
            {;
                sKey = sKey.Substring( nChildDelimiter + 1);
                if ( ! PrevNode Is null )
            {;
                if ( UBound(NodeStack) > 0 )
            {;
                ;
                }
            else
            {;
                ;
                };
                NodeStack(UBound(NodeStack)) = nodX;
                nodX = PrevNode;
                };
                };
                };
                try
{;
                if ( nodX Is null )
            {;
                PrevNode = tvwX.Nodes.Add(                                                                                                                                                                                                      , , sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter));
                }
            else
            {;
                PrevNode = tvwX.Nodes.Add(                                                                                                                                                                                                      nodX.Key, tvwChild, nodX.Key + "_" + sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter));
                };

                PrevNode.ExpandedImage = sGetToken(sIcon, 1, EndChildDelimiter);
                PrevNode.Expanded = bExpanded;
                PrevNode.Tag = sTag;

                ;
                if ( nEndChildDelimiter > 0 )
            {;
                sIcon.Substring(sIcon.Length - nEndChildDelimiter) = EndChildDelimiter) {;
                sIcon.Substring(0, Len(sIcon) - nEndChildDelimiter);
                nodX = NodeStack(UBound(NodeStack));
                if ( UBound(NodeStack) > 1 )
            {;
                ;
                }
            else
            {;
                ;
                };
                Loop;
                };
                } // CurItem;
                CurItem = null;
                PrevNode = null;
                nodX = null;
                tvwX = null;
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void ListViewToAll            {
                Integer     static CurSubItem;
                string     static sAll;

                ListView     static lvwX;
                ListItem     static CurListItem;

                lvwX = lvwCtrl;
                Clear;
                sAll = "";
                foreach( var CurListItem in lvwX.ListItems;

                sAll +=  ItemDelimiter + CurListItem.Key + KeyValueDelimiter + CurListItem.Icon;
                for(var CurSubItem = 1; CurSubItem < lvwX.ColumnHeaders.Count - 1; CurSubItem++)  {;
                sAll +=  FieldDelimiter + CurListItem.SubItems(CurSubItem);
                } // CurSubItem;
                sAll = sAll.Substring( Len(FieldDelimiter));

                } // CurListItem;
                All = sAll.Substring( Len(ItemDelimiter) + 1);
                CurListItem = null;
                lvwX = null;
                sAll = "";
            }

            public void RecordsetToAll            {
                ;
                this.Clear;
                ;
                Do Until rst.EOF;

                for(var CurField = 1; CurField < rst.Fields.Count - 1; CurField++)  {;
                this.Add(                                                                                                                                                                                                      rst.Fields(0)).Value +=  sReplace("" + rst.Fields(CurField), this.FieldDelimiter, ".") + this.FieldDelimiter;
                } // CurField;
                .Value.Substring(0, Len(.Value) - 1);

                rst.MoveNext;
                Loop;
            }

            public void Remove            {
                mCol.Remove sIndexKey;
            }

            public void TreeToAll            {
                ;

                ;
                ;

                tvwX = tvwToGet;
                Clear;
                foreach( var CurNode in tvwX.Nodes;

                if ( CurNode.Parent Is null )
            {;
                if ( Len(.Tag) = 0 )
            {;
                sAll +=  ItemDelimiter + CurNode.Text + IconDelimiter + CurNode.Image;
                }
            else
            {;
                sAll +=  ItemDelimiter + CurNode.Text + TagDelimiter + CurNode.Tag + IconDelimiter + CurNode.Image;
                };
                TreeToAll_AddChildren sAll, CurNode;
                };

                } // CurNode;
                All = sAll.Substring( 3);
                sAll = "";
                CurNode = null;
                tvwX = null;
            }

            public void TreeToAll_AddChildren            {
                ;
                CurChild = CurNode.Child;
                if ( ! CurChild Is null )
            {;
                if ( Len(CurChild.Tag) = 0 )
            {;
                sAll +=  ItemDelimiter + ChildDelimiter + CurChild.Text + IconDelimiter + CurChild.Image;
                }
            else
            {;
                sAll +=  ItemDelimiter + ChildDelimiter + CurChild.Text + TagDelimiter + CurChild.Tag + IconDelimiter + CurChild.Image;
                };
                ;
                if ( ! CurChild.Child Is null )
            {;
                TreeToAll_AddChildren sAll, CurChild;
                };
                ;
                CurChild = CurChild.Next;
                ;
                Do Until CurChild Is null;
                if ( Len(CurChild.Tag) = 0 )
            {;
                sAll +=  ItemDelimiter + CurChild.Text + IconDelimiter + CurChild.Image;
                }
            else
            {;
                sAll +=  ItemDelimiter + CurChild.Text + TagDelimiter + CurChild.Tag + IconDelimiter + CurChild.Image;
                };
                if ( ! CurChild.Child Is null )
            {;
                TreeToAll_AddChildren sAll, CurChild;
                };
                CurChild = CurChild.Next;
                Loop;
                sAll +=  EndChildDelimiter;
                };
                CurChild = null;
            }

            public void Class_Initialize()
            {
                mCol = new Dictionary<string,string>();
                ItemDelimiter = "~";
                KeyValueDelimiter = "=";
                FieldDelimiter = ",";
            }

            public void Class_Terminate()
            {
                mCol = null;
            }

        }
    }
