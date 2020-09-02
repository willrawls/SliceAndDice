using System;

namespace MetX.SliceAndDice
{
		public class CAssocArray
		{

			public  mCol;
			public  CurItem;
			public  Section;
			public  ItemDelimiter;
			public string KeyValueDelimiter;
			public  FieldDelimiter;

			// ' ********************************************************************************
' Class Module      CAssocArray
'
' Filename          CAssocArray.cls
'
' Copyright         1998 by Firm Solutions
'                   All rights reserved.
'
' Author            William M. Rawls
'                   Firm Solutions
'
' Created On        4/30/1998 1:23 PM
'
' Description
'
'    The Reality Matrix, Dimention 2 of 3
'       "Associative array" like abilities
'
'    What's "assosiative array" like abilities mean ?
'       Email = wrawls@firmsolutions.com to find out more.
'    Web page you ask ?
'       URL = http://www.firmsolutions.com/RealityMatrix.html
'    Why does this read like an e-mail ?
'       Because = It does
'
' Revisions
'
' <RevisionDate>, <RevisedBy>
' <Description of Revision>
'
' 4/30/1998, William M. Rawls
' Entered into public domain as freeware. Use at your own risk.
'
' ********************************************************************************
;
			public string All
			{
				get
				{
					static sAllKeyValues As String;
					foreach( var CurItem in mCol );
					sAllKeyValues = sAllKeyValues + CurItem.Key + KeyValueDelimiter + .Value + ItemDelimiter;
					End With;
					}
					CurItem = null;
					All = sAllKeyValues;
					sAllKeyValues = "";
				}

				set
				{
					static sT As String;
					static sKey As String;
					static sValue As String;
					Clear;
					Do While Len(value) > 0;
					sT = sGetToken(value, 1, ItemDelimiter);
					if ( InStr(sT, KeyValueDelimiter) > 0 )
				{
					sKey = sGetToken(sT, 1, KeyValueDelimiter);
					sValue = sGetToken(sT, 2, KeyValueDelimiter);
					Add sKey, sValue;
					}
				else
				{
					Add sT;
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
					// TODO: Rewrite try/catch and/or goto. EH_CAssocArray_Item;
					Item = mCol(sIndexKey);
					EH_CAssocArray_Item_Continue:;
					return; // ???;
					EH_CAssocArray_Item:;
					Item = Add(sIndexKey);
					Resume EH_CAssocArray_Item_Continue;
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
					static sKeys As String;
					String sDelimiter;
					sDelimiter = ItemDelimiter;
					foreach( var CurItem in mCol );
					sKeys = sKeys + sDelimiter + CurItem.Key;
					}
					Keys = Mid$(sKeys, 2);
					sKeys = "";
				}

				set
				{
					Long TokenCount;
					Long CurToken;
					String sCurToken;
					Me.Clear;
					TokenCount = iTokenCount(value, Me.ItemDelimiter);
					For CurToken = 1 To TokenCount;
					Me.Add sGetToken(value, CurToken, ItemDelimiter);
					}
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
				mCol = new Dictionary<string,string>()();
			}
			public As Add			{
				CurItem = new CAssocItem();

				if ( Len(Value) = 0 And InStr(Key, Me.KeyValueDelimiter) > 0 )
				{;
				CurItem.Key = sGetToken(Key, 1, Me.KeyValueDelimiter);
				CurItem.Value = sAfter(Key, 1, Me.KeyValueDelimiter);
				}
				else
				{;
				CurItem.Key = Key;
				CurItem.Value = Value;
				};
				End With;
				mCol.Add CurItem, Key;
				Add = CurItem;
				CurItem = null;
			}
			public As Column			{
				static sColumn As String;

				if ( Len(sDelimiter) = 0 )
				{
 sDelimiter = ItemDelimiter;

				foreach( var CurItem in mCol );
				sColumn = sColumn + sDelimiter + CurItem.F(iCol, FieldDelimiter);
				};
				CurItem = null;

				Column = Mid$(sColumn, 2);
				sColumn = "";
			}
			public void FillList			{
				if ( bClearList )
				{
 lstCtrl.Clear;

				if ( iColumn = 0 )
				{;
				foreach( var CurItem in mCol );
				lstCtrl.AddItem CurItem.Key;
				};
				}
				else
				{;
				foreach( var CurItem in mCol );
				lstCtrl.AddItem CurItem.F(CLng(iColumn), FieldDelimiter);
				};
				};
				CurItem = null;

				if ( StrComp(sItemToChoose, "*NONE*") <> 0 )
				{;
				SetListIndex lstCtrl, sItemToChoose;
				};
			}
			public void FillListView			{
				static lvwX As ListView;
				static NewItem As ListItem;
				static SubItems As Integer;
				static CurSubItem As Integer;
				static sT As String;

				lvwX = lvwCtrl;
				lvwX.ListItems.Clear;
				if ( bFullLine )
				{
 ExtendListView lvwX.hWnd;

				foreach( var CurItem in mCol );

				SubItems = iTokenCount(.Value, FieldDelimiter);
				if ( SubItems > 0 )
				{;
				sT = sGetToken(.Value, 1, FieldDelimiter);
				NewItem = lvwX.ListItems.Add(, "Key " + CurItem.Key, .Key) ', sT, sT);
				For CurSubItem = 2 To SubItems;
				NewItem.SubItems(CurSubItem - 1) = sGetToken(.Value, CurSubItem, FieldDelimiter);
				};
				}
				else
				{;
				NewItem = lvwX.ListItems.Add(, CurItem.Key, .Key);
				};
				End With;
				};
				CurItem = null;
				NewItem = null;
				lvwX = null;
				sT = "";
			}
			public void FillListViewColumns			{
				static lvwX As ListView;
				lvwX = lvwCtrl;

				lvwX.ColumnHeaders.Clear;
				foreach( var CurItem in mCol );
				lvwX.ColumnHeaders.Add , CurItem.Key, CurItem.Key, Val(CurItem.Value);
				};
				CurItem = null;
				End With;
				lvwX = null;
			}
			public void FillTreeNode			{
				Node nodX;
				Node PrevNode;
				TreeView tvwX;
				String sKey;
				String sTag;
				String sIcon;
				Integer nChildDelimiter;
				Integer nEndChildDelimiter;

				Node NodeStack();
				0) NodeStack(0;

				nChildDelimiter = Len(ChildDelimiter);
				nEndChildDelimiter = Len(EndChildDelimiter);

				tvwX = tvwToFill;
				nodX = nodCur;
				foreach( var CurItem in mCol );
				sKey = sGetToken(CurItem.Key, 1, IconDelimiter);
				sIcon = sGetToken(CurItem.Key, 2, IconDelimiter);
				sTag = sGetToken(sKey, 2, TagDelimiter);
				sKey = sGetToken(sKey, 1, TagDelimiter);
				if ( Len(sIcon) = 0 )
				{
 sIcon = sImage;

				if ( nChildDelimiter > 0 )
				{;
				if ( Left$(sKey, nChildDelimiter) = ChildDelimiter )
				{;
				sKey = Mid$(sKey, nChildDelimiter + 1);
				if ( ! PrevNode Is null )
				{
;
				if ( UBound(NodeStack) > 0 )
				{;
				To Preserve;
				}
				else
				{;
				1) NodeStack(1;
				};
				NodeStack(UBound(NodeStack)) = nodX;
				nodX = PrevNode;
				};
				};
				};
				try
{;
				if ( nodX Is null )
				{
;
				PrevNode = tvwX.Nodes.Add(, , sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter));
				}
				else
				{;
				PrevNode = tvwX.Nodes.Add(nodX.Key, tvwChild, nodX.Key + "_" + sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sKey, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter), sGetToken(sIcon, 1, EndChildDelimiter));
				};

				PrevNode.ExpandedImage = sGetToken(sIcon, 1, EndChildDelimiter);
				PrevNode.Expanded = bExpanded;
				PrevNode.Tag = sTag;
				End With;

				if ( nEndChildDelimiter > 0 )
				{;
				Do While Right$(sIcon, nEndChildDelimiter) = EndChildDelimiter;
				sIcon = Left$(sIcon, Len(sIcon) - nEndChildDelimiter);
				nodX = NodeStack(UBound(NodeStack));
				if ( UBound(NodeStack) > 1 )
				{;
				To Preserve;
				}
				else
				{;
				0) NodeStack(0;
				};
				Loop;
				};
				};
				CurItem = null;
				PrevNode = null;
				nodX = null;
				tvwX = null;
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void ListViewToAll			{
				static CurSubItem As Integer;
				static sAll As String;

				static lvwX As ListView;
				static CurListItem As ListItem;

				lvwX = lvwCtrl;
				Clear;
				sAll = "";
				foreach( var CurListItem in lvwX.ListItems );

				sAll = sAll + ItemDelimiter + CurListItem.Key + KeyValueDelimiter + .Icon;
				For CurSubItem = 1 To lvwX.ColumnHeaders.Count - 1;
				sAll = sAll + FieldDelimiter + CurListItem.SubItems(CurSubItem);
				};
				sAll = Mid$(sAll, Len(FieldDelimiter));
				End With;
				};
				All = Mid$(sAll, Len(ItemDelimiter) + 1);
				CurListItem = null;
				lvwX = null;
				sAll = "";
			}
			public void RecordsetToAll			{
				Long CurField;
				Me.Clear;

				Do Until rst.EOF;

				For CurField = 1 To rst.Fields.Count - 1;
				Me.Add(rst.Fields(0)).Value = Me.Add(rst.Fields(0)).Value + sReplace("" + rst.Fields(CurField), Me.FieldDelimiter, ".") + Me.FieldDelimiter;
				};
				Me.Add(rst.Fields(0)).Value = Left(.Value, Len(.Value) - 1);
				End With;
				rst.MoveNext;
				Loop;
			}
			public void Remove			{
				mCol.Remove sIndexKey;
			}
			public void TreeToAll			{
				String sAll;

				Node CurNode;
				TreeView tvwX;

				tvwX = tvwToGet;
				Clear;
				foreach( var CurNode in tvwX.Nodes );

				if ( CurNode.Parent Is null )
				{
;
				if ( Len(.Tag) = 0 )
				{;
				sAll = sAll + ItemDelimiter + CurNode.Text + IconDelimiter + .Image;
				}
				else
				{;
				sAll = sAll + ItemDelimiter + CurNode.Text + TagDelimiter + .Tag + IconDelimiter + .Image;
				};
				TreeToAll_AddChildren sAll, CurNode;
				};
				End With;
				};
				All = Mid$(sAll, 3);
				sAll = "";
				CurNode = null;
				tvwX = null;
			}
			public void TreeToAll_AddChildren			{
				Node CurChild;
				CurChild = CurNode.Child;
				if ( ! CurChild Is null )
				{
;
				if ( Len(CurChild.Tag) = 0 )
				{;
				sAll = sAll + ItemDelimiter + ChildDelimiter + CurChild.Text + IconDelimiter + CurChild.Image;
				}
				else
				{;
				sAll = sAll + ItemDelimiter + ChildDelimiter + CurChild.Text + TagDelimiter + CurChild.Tag + IconDelimiter + CurChild.Image;
				};

				if ( ! CurChild.Child Is null )
				{
;
				TreeToAll_AddChildren sAll, CurChild;
				};

				CurChild = CurChild.Next;

				Do Until CurChild Is null;
				if ( Len(CurChild.Tag) = 0 )
				{;
				sAll = sAll + ItemDelimiter + CurChild.Text + IconDelimiter + CurChild.Image;
				}
				else
				{;
				sAll = sAll + ItemDelimiter + CurChild.Text + TagDelimiter + CurChild.Tag + IconDelimiter + CurChild.Image;
				};
				if ( ! CurChild.Child Is null )
				{
;
				TreeToAll_AddChildren sAll, CurChild;
				};
				CurChild = CurChild.Next;
				Loop;
				sAll = sAll + EndChildDelimiter;
				};
				CurChild = null;
			}
			public void Class_Initialize()
			{
				mCol = new Dictionary<string,string>()();
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
