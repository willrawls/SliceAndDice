Attribute VB_Name = "modGeneral"
Option Explicit

#If Win32 Then
    Private Const MAX_PATH = 260
    Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
    End Type
    Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
    End Type
    
    Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
    Private Declare Function GetLastError Lib "kernel32" () As Long
    
    Private Const ERROR_NO_MORE_FILES = 18&
    Private Const INVALID_HANDLE_VALUE = -1
    Private Const DDL_DIRECTORY = &H10
#End If

Public NewMessageText As String

Public Function StringToClipboard(ByVal sTextToPutOnClipboard As String) As Boolean
479      On Error Resume Next
480      If Len(sTextToPutOnClipboard) = 0 Then StringToClipboard = True: Exit Function

481      Err.Clear
482      Clipboard.Clear

483      If Err.Number = 0 Then
484         Clipboard.SetText sTextToPutOnClipboard, vbCFText
485         If Err.Number = 0 Then
486            StringToClipboard = True
487         Else
488            MsgBox "Error " & Err.Number & ") Putting Text onto the Clipboard. Error Description = " & Err.Description, , "sadSoftCodeWmr modGeneral.StringToClipboard (line " & Erl & ")"
489         End If
490      Else
            MsgBox "Error " & Err.Number & ") Putting Text onto the Clipboard. Error Description = " & Err.Description, , "sadSoftCodeWmr modGeneral.StringToClipboard (line " & Erl & ")"
492      End If
End Function

Public Function GetFileList(ByVal sStartingDirectory As String, ByVal sFilePattern As String, Optional ByVal sItemDelimiter As String = vbNewLine) As String
    Screen.MousePointer = vbHourglass
        DoEvents
        If Right$(sStartingDirectory, 1) <> "\" Then sStartingDirectory = sStartingDirectory & "\"
        GetFileList = FindFiles(sStartingDirectory, sFilePattern)
    Screen.MousePointer = vbDefault
End Function


Public Function FindFiles(ByVal sStartingDirectory As String, ByVal sFilePattern As String) As String
    Dim null_character As String
    Dim dirs() As String
    Dim num_dirs As Long
    Dim sub_dir As String
    Dim file_name As String
    Dim i As Integer
    Dim txt As String
    Dim search_handle As Long
    Dim file_data As WIN32_FIND_DATA

    ' ASCII character 0 terminates strings.
    null_character = Chr$(0)

    ' Search for matching files in this directory.
    ' Get the first matching file.
    search_handle = FindFirstFile( _
        sStartingDirectory & sFilePattern, file_data)
    If search_handle <> INVALID_HANDLE_VALUE Then
        ' Save this file's name.
        Do While GetLastError <> ERROR_NO_MORE_FILES
            file_name = file_data.cFileName
            file_name = Left$(file_name, _
                InStr(file_name, null_character) - 1)
            If file_name <> "." And file_name <> ".." Then
                ' Add the file to the return value.
                txt = txt & sStartingDirectory & file_name & vbCrLf
            End If

            ' Get the next file.
            FindNextFile search_handle, file_data
        Loop

        ' Close the file search hanlde.
        FindClose search_handle
    End If

    ' Get this directory's subdirectories.
    ' Get the first subdirectory.
    search_handle = FindFirstFile( _
        sStartingDirectory & "*.*", file_data)
    If search_handle <> INVALID_HANDLE_VALUE Then
        ' Save this file's name.
        Do While GetLastError <> ERROR_NO_MORE_FILES
            ' Save the subdirectory name.
            If file_data.dwFileAttributes And DDL_DIRECTORY Then
                file_name = file_data.cFileName
                file_name = Left$(file_name, _
                    InStr(file_name, null_character) - 1)
                If file_name <> "." And file_name <> ".." Then
                    num_dirs = num_dirs + 1
                    ReDim Preserve dirs(1 To num_dirs)
                    dirs(num_dirs) = sStartingDirectory & file_name & "\"
                End If
            End If

            ' Get the next file.
            FindNextFile search_handle, file_data
        Loop

        ' Close the file search hanlde.
        FindClose search_handle
    End If

    ' Recursively search the subdirectories.
    For i = 1 To num_dirs
        ' Add this subdirectory's matching files
        ' to the result string.
        txt = txt & FindFiles(dirs(i), sFilePattern)
    Next i

    ' Return the string we have built.
    FindFiles = txt
End Function

