using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
    /// <summary>
    ///   //  **********************************************************
//  API call to determin where the user's Windows directory is
//  **********************************************************
//  -------------------------------------------------
//  Calls the windows API to get the windows directory
//  -------------------------------------------------
.
    /// </summary>
    class modVB5AddInTool
    {
    public  sGetWindowsDir$()
    {
        ;
        ;
        sT = string$(145, 0)              ' Size Buffer;
        x = GetWindowsDirectory(sT, 145)  ' Make API Call;
        sT.Substring(0, x)                 ' Trim Buffer;
        sT.Substring(sT.Length - 1) <> "\" )
            {
      ' Add(                                                                                                   \ if necessary);
        sGetWindowsDir = sT + "\";
        }
            else
            {;
        sGetWindowsDir = sT;
        };
    }

    }
    }
