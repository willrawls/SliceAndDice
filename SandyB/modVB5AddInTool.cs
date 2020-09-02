using System;

namespace MetX.SliceAndDice
{
		class modVB5AddInTool
		{

			public  sGetWindowsDir$()
			{
				Integer x;
				String sT;

				sT = String$(145, 0)              ' Size Buffer;
				x = GetWindowsDirectory(sT, 145)  ' Make API Call;
				sT = Left$(sT, x)                 ' Trim Buffer;

				if ( Right$(sT, 1) <> "\" )
				{
      ' Add \ if necessary;
				sGetWindowsDir = sT + "\";
				}
				else
				{;
				sGetWindowsDir = sT;
				};
			}
		}
}
