using System;

namespace MetX.SliceAndDice
{
		class modGeneral
		{

			public string NewMessageText;

			public As StringToClipboard			{
				try
{;
				if ( Len(sTextToPutOnClipboard) = 0 )
				{
 StringToClipboard = true: return; // ???;


				Clipboard.Clear;

				if ( ex = 0 )
				{;
				Clipboard.SetText sTextToPutOnClipboard, vbCFText;
				if ( ex = 0 )
				{;
				StringToClipboard = true;
				}
				else
				{;
				MsgBox "Error " + ex + ") Putting Text onto the Clipboard. Error Description = " + Err.Description, , "sadSoftCodeWmr modGeneral.StringToClipboard (line " + Erl + ")";
				}
        catch(Exception e)
        {
            /* ON ERROR RESUME NEXT (ish) */
        }
;
			}
			public As GetFileList			{
				Screen.MousePointer = vbHourglass;
				DoEvents;
				if ( Right$(sStartingDirectory, 1) <> "\" )
				{
 sStartingDirectory = sStartingDirectory + "\";
				GetFileList = FindFiles(sStartingDirectory, sFilePattern);
				Screen.MousePointer = vbDefault;
			}
			public As FindFiles			{
				String null_character;
				String dirs();
				Long num_dirs;
				String sub_dir;
				String file_name;
				Integer i;
				String txt;
				Long search_handle;
				WIN32_FIND_DATA file_data;

				//  ASCII character 0 terminates strings.;
				null_character = Chr$(0);

				//  Search for matching files in this directory.;
				//  Get the first matching file.;
				search_handle = FindFirstFile(
        sStartingDirectory + sFilePattern, file_data);
				if ( search_handle <> INVALID_HANDLE_VALUE )
				{;
				//  Save this file's name.;
				Do While GetLastError <> ERROR_NO_MORE_FILES;
				file_name = file_data.cFileName;
				file_name = Left$(file_name,
                InStr(file_name, null_character) - 1);
				if ( file_name <> "." And file_name <> ".." )
				{;
				//  Add the file to the return value.;
				txt = txt + sStartingDirectory + file_name + vbCrLf;
				};

				//  Get the next file.;
				FindNextFile search_handle, file_data;
				Loop;

				//  Close the file search hanlde.;
				FindClose search_handle;
				};

				//  Get this directory's subdirectories.;
				//  Get the first subdirectory.;
				search_handle = FindFirstFile(
        sStartingDirectory + "*.*", file_data);
				if ( search_handle <> INVALID_HANDLE_VALUE )
				{;
				//  Save this file's name.;
				Do While GetLastError <> ERROR_NO_MORE_FILES;
				//  Save the subdirectory name.;
				if ( file_data.dwFileAttributes And DDL_DIRECTORY )
				{;
				file_name = file_data.cFileName;
				file_name = Left$(file_name,
                    InStr(file_name, null_character) - 1);
				if ( file_name <> "." And file_name <> ".." )
				{;
				num_dirs = num_dirs + 1;
				ReDim Preserve dirs(1 To num_dirs);
				dirs(num_dirs) = sStartingDirectory + file_name + "\";
				};
				};

				//  Get the next file.;
				FindNextFile search_handle, file_data;
				Loop;

				//  Close the file search hanlde.;
				FindClose search_handle;
				};

				//  Recursively search the subdirectories.;
				For i = 1 To num_dirs;
				//  Add this subdirectory's matching files;
				//  to the result string.;
				txt = txt + FindFiles(dirs(i), sFilePattern);
				};

				//  Return the string we have built.;
				FindFiles = txt;
			}
		}
}
