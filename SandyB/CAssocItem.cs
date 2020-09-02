using System;

namespace MetX.SliceAndDice
{
		public class CAssocItem
		{

			public string Key;
			public string m_sValue;

			// ' ********************************************************************************
' Class Module      CAssocItem
'
' Filename          CAssocItem.cls
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
' The Reality Matrix, Dimention 3 of 3
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
			public object Value
			{
				get
				{
					Value = m_sValue;
				}

				set
				{
					m_sValue = value;
				}

			}

			// '
' Retrieves the Nth delimited token from the value.
' If delimiter is ommited, then a space is assumed as the delimiter.
' NOTE: sGetToken required for proper use
'
;
			public string F
			{
				get
				{
					F = sGetToken(Value, Index, sDelimiter);
				}

			}

		}
}
