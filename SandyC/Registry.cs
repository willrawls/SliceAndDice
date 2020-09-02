using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
        class Registry
        {

        public Long, hKey;
        public Long, rtn;
        public int lBufferSize;
        public int lDataSize;
        public byte ByteArray();
        public string sResult;
        public string OrigKey;

            public void sadSaveLicenseKey            {
                try
{;
                SetstringValue "HKEY_LOCAL_MACHINE" + gsBS + "SOFTWARE" + gsBS + "Zion Systems" + gsBS + "License", Key, sadEncrypt(Value);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As sadGetLicenseKey            {
                try
{;
                sResult = GetstringValue("HKEY_LOCAL_MACHINE" + gsBS + "SOFTWARE" + gsBS + "Zion Systems" + gsBS + "License", Key);
                if ( Len(sResult) == 0 Or sResult == "Error" )
            {
 sResult == sDefault;
                sResult.Substring(0, 4) = "EN* " )
            {;
                sResult = sadDecrypt(sResult);
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public void ParseKey            {

                Keyname.Contains(gsBS)                        'return if gsBS is contained in the Keyname;

                Keyname.Substring(0.Substring(Keyname.Substring(0.Length - 1) == gsBS )
            {
    'if the is a gsBS at the end of the Keyname then;
                MsgBox "Incorrect Format:" + gs2EOL + Keyname 'display error to the user;
                return;                                      'exit the procedure;
                }
            else
            {if ( rtn == 0 )
            {
                               'if the Keyname contains no gsBS;
                Keyhandle = GetMainKeyHandle(Keyname);
                Keyname = string.Empty                        'leave Keyname blank;
                }
            else
            {                                              'otherwise, Keyname contains gsBS;
                Keyname.Substring(0, rtn - 1))    'seperate the Keyname;
                Keyname.Substring(Keyname.Length - Len(Keyname) - rtn);
            }

        }
    }
