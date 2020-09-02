using System;

namespace MetX.SliceAndDice
{
		class modCrypt
		{

			public As Const;
			public byte Dummy;
			public int CurrByte;
			public string sOut;
			public string strOut;
			public string strOut;
			public byte bytArray();
			public int CurrByte;
			public string strOut;
			public byte bytArray();
			public byte bytStack;
			public int CurrByte;
			public int MaxCount;

			public As LocalGenerated			{

				For CurrByte = 1 To Length;
				sOut = sOut + (CLng(Rnd() * 10) Mod 10);
				};

				LocalGenerated = sOut;
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
			public As Scramble			{

				if ( Left$(strIn, 4) = "EN* " )
				{;
				strIn = Mid$(strIn, 5);
			}
		}
}
