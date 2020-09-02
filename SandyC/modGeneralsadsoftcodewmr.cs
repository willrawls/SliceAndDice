using System;
using System.Collections;
using System.Collections.Generic;

namespace MetX.SliceAndDice
{
        class modGeneral
        {

        public string NewMessageText;

            public As stringToClipboard            {
                try
{;
                if ( Len(sTextToPutOnClipboard) == 0 )
            {
 stringToClipboard == true: return; // ???;


                Clipboard.Clear;

                if ( ex = 0 )
            {;
                Clipboard.SetText sTextToPutOnClipboard, vbCFText;
                if ( ex = 0 )
            {;
                stringToClipboard = true;
                }
            else
            {;
                MsgBox "Error " + ex + ") Putting Text onto the Clipboard. Error Description = " + Err.Description, , "sadSoftCodeWmr modGeneral.stringToClipboard (line " + Erl + ")";
                }
        catch(Exception e)
        {
            // ON ERROR RESUME NEXT
        };
            }

            public As GetFileList            {
                Screen.MousePointer = vbHourglass;
                DoEvents;
                sStartingDirectory.Substring(sStartingDirectory.Length - 1) <> "\" )
            {
 sStartingDirectory == sStartingDirectory + "\";
                GetFileList = FindFiles(sStartingDirectory, sFilePattern);
                Screen.MousePointer = vbDefault;
            }

            public As FindFiles            {
                ;
                ;
                ;
                ;
                ;
                ;
                ;
                ;
                ;

                //    ' ASCII character 0 terminates strings.;
                null_character = Chr$(0);

                //    ' Search for matching files in this directory.;
                //    ' Get the first matching file.;
                search_handle = FindFirstFile(
        sStartingDirectory + sFilePattern, file_data);
                if ( search_handle <> INVALID_HANDLE_VALUE )
            {;
                //        ' Save this file's name.;
                while(GetLastError <> ERROR_NO_MORE_FILES) {;
                file_name = file_data.cFileName;
                file_name.Substring(0, file_name.Contains(null_character) - 1);
                if ( file_name <> "." And file_name <> ".." )
            {;
                //                ' Add the file to the return value.;
                txt +=  sStartingDirectory + file_name + vbCrLf;
                };

                //            ' Get the next file.;
                FindNextFile search_handle, file_data;
                Loop;

                //        ' Close the file search hanlde.;
                FindClose search_handle;
                };

                //    ' Get this directory's subdirectories.;
                //    ' Get the first subdirectory.;
                search_handle = FindFirstFile(
        sStartingDirectory + "*.*", file_data);
                if ( search_handle <> INVALID_HANDLE_VALUE )
            {;
                //        ' Save this file's name.;
                while(GetLastError <> ERROR_NO_MORE_FILES) {;
                //            ' Save the subdirectory name.;
                if ( file_data.dwFileAttributes And DDL_DIRECTORY )
            {;
                file_name = file_data.cFileName;
                file_name.Substring(0, file_name.Contains(null_character) - 1);
                if ( file_name <> "." And file_name <> ".." )
            {;
                num_dirs +=  1;
                ReDim Preserve dirs(1 To num_dirs);
                dirs(num_dirs) = sStartingDirectory + file_name + "\";
                };
                };

                //            ' Get the next file.;
                FindNextFile search_handle, file_data;
                Loop;

                //        ' Close the file search hanlde.;
                FindClose search_handle;
                };

                //    ' Recursively search the subdirectories.;
                for(var i = 1; i < num_dirs; i++)  {;
                //        ' Add this subdirectory's matching files;
                //        ' to the result string.;
                txt +=  FindFiles(dirs(i), sFilePattern);
                } // i;

                //    ' Return the string we have built.;
                FindFiles = txt;
            }

        }
    }
