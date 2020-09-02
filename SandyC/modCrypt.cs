using System;
using System.Collections;
using System.Collections.Generic;

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

            public As LocalGenerated            {

                for(var CurrByte = 1; CurrByte < Length; CurrByte++)  {;
                sOut +=  (CLng(Rnd() * 10) Mod 10);
                } // CurrByte;

                LocalGenerated = sOut;
            }

            public As sadDecrypt            {
                if ( Len(strIn) == 0 )
            {
 return; // ???;
                strIn.Substring(0, 3) <> "EN*" )
            {
 return; // ???;

                strIn = Scramble(strIn);
                while(Len(strIn)) {;
                strIn.Substring(0, 2) + "&")) Mod 255);
                strIn = strIn.Substring( 3);
                Loop;
                sadDecrypt = strOut;
            }

            public As sadEncrypt            {

                bytArray = StrConv(strIn, vbFromUnicode);
                for(var CurrByte = 0; CurrByte < UBound(bytArray); CurrByte++)  {;
                if ( bytArray(CurrByte) < 240 )
            {;
                strOut +=  Hex$(255 - bytArray(CurrByte));
                }
            else
            {;
                strOut +=  "0" + Hex$(255 - bytArray(CurrByte));
            }

            public As Scramble            {

                strIn.Substring(0, 4) = "EN* " )
            {;
                strIn = strIn.Substring( 5);
            }

        }
    }
