using System;

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

			public void sadSaveLicenseKey			{
				try
{;
				SetStringValue "HKEY_LOCAL_MACHINE" + gsBS + "SOFTWARE" + gsBS + "Zion Systems" + gsBS + "License", Key, sadEncrypt(Value);
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As sadGetLicenseKey			{
				try
{;
				sResult = GetStringValue("HKEY_LOCAL_MACHINE" + gsBS + "SOFTWARE" + gsBS + "Zion Systems" + gsBS + "License", Key);
				if ( Len(sResult) = 0 Or sResult = "Error" )
				{
 sResult = sDefault;
				if ( Left$(sResult, 4) = "EN* " )
				{;
				sResult = sadDecrypt(sResult);
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public void ParseKey			{

				rtn = InStr(Keyname, gsBS)                        'return if gsBS is contained in the Keyname;

				if ( Left$(Keyname, 5) <> "HKEY_" Or Right$(Keyname, 1) = gsBS )
				{
    'if the is a gsBS at the end of the Keyname then;
				MsgBox "Incorrect Format:" + gs2EOL + Keyname 'display error to the user;
				return;                                      'exit the procedure;
				}
				else
				{if ( rtn = 0 )
				{
                               'if the Keyname contains no gsBS;
				Keyhandle = GetMainKeyHandle(Keyname);
				Keyname = string.Empty                        'leave Keyname blank;
				}
				else
				{                                              'otherwise, Keyname contains gsBS;
				Keyhandle = GetMainKeyHandle(Left$(Keyname, rtn - 1))    'seperate the Keyname;
				Keyname = Right$(Keyname, Len(Keyname) - rtn);
			}
		}
}
