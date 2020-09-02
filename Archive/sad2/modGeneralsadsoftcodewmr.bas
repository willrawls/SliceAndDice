Attribute VB_Name = "Order"
Option Explicit
Option Compare Text

Public RootTruth    As String

Private Crc32Table(255) As Long
Private colCache    As Collection
Private DataArray   As Variant
Private b1() As Byte
Private b2() As Byte

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
    search_handle = FindFirstFile(sStartingDirectory & "*.*", file_data)
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


Public Function CreateAllDirs(ByVal ToCheck As String) As Boolean
On Error GoTo errCreateAllDirs
  ' Recursively creates non-existing directories from the given path name
    Dim ToCreate    As String
    Dim CurrIndex   As Long

    Do While Not DoesDirectoryExist(ToCheck)
     ' Loop and parse through the string
       CurrIndex = InStr(CurrIndex + 1, ToCheck & IIf(Right$(ToCheck, 1) = "\", "", "\"), "\")
       If CurrIndex = 0 Then Exit Do

     ' Create the path name
       ToCreate = Mid$(ToCheck, 1, CurrIndex - 1)
       MkDir ToCreate
    Loop
    CreateAllDirs = True

ContinueAfterError:
    Exit Function

errCreateAllDirs:
    Select Case Err.Number
        Case 75 ' Ignore errors for existing paths
             Resume Next
        Case Else
             Resume ContinueAfterError
    End Select
End Function

Public Function DoesDirectoryExist(ByVal ToCheck As String) As Boolean
    DoesDirectoryExist = (Not (Len(Dir$(ToCheck)) > 0)) And (Len(Dir$(ToCheck, vbDirectory)) > 0)
End Function

Public Function FileList(ByVal sSpec As String) As Variant
    Dim nFiles     As Long
    Dim sFile      As String
    Dim vFileList  As Variant

    ReDim vFileList(1 To 10) As Variant
    sFile = Dir$(sSpec)

    Do While Len(sFile)
       nFiles = nFiles + 1
       vFileList(nFiles) = sFile
       If nFiles = UBound(vFileList) Then
          ReDim Preserve vFileList(1 To UBound(vFileList) + 10) As Variant
       End If
       sFile = Dir
    Loop

    If nFiles Then
       ReDim Preserve vFileList(1 To nFiles) As Variant
    End If

    FileList = vFileList
End Function

Public Sub StringToFile(ByVal sPathAndFilename As String, ByRef sToFile, Optional ByVal bStoreIfBlank As Boolean = True, Optional ByVal bAppend As Boolean = False)
On Error Resume Next
    Dim fh As Long

    If bStoreIfBlank Or Len(Trim$(sToFile)) > 0 Then
        fh = FreeFile
        If Not bAppend Then
            Open sPathAndFilename For Output Access Write As #fh
        Else
            Open sPathAndFilename For Append Access Write As #fh
        End If
             Print #fh, sToFile
        Close #fh
    End If
End Sub

Public Function FileToString(ByVal sPathAndFilename As String) As String
On Error Resume Next
    Dim fh As Long

    If Len(Dir$(sPathAndFilename)) Then
       fh = FreeFile
       Open sPathAndFilename For Input Access Read As #fh
            FileToString = Input(LOF(fh), fh)
       Close #fh
    End If
End Function

Public Function FileToArray(ByVal sPathAndFilename As String) As Variant
On Error Resume Next
    Dim fh As Long

    If Len(Dir$(sPathAndFilename)) Then
       fh = FreeFile
       Open sPathAndFilename For Input Access Read As #fh
            FileToArray = Split(Input(LOF(fh), fh), " ")
       Close #fh
    End If
End Function

Public Function FilesToArray(ByVal sPathAndFileSpec As String) As Variant
    Dim FileNames    As Variant
    Dim FileArrays   As Variant
    Dim CurrFile     As Variant
    Dim FileCount    As Long

    FileNames = FileList(sPathAndFileSpec)
    ReDim FileArrays(0 To UBound(FileNames)) As Variant

    FileCount = 0
    For Each CurrFile In FileNames
        FileCount = FileCount + 1
        FileArrays(FileCount) = Array(CurrFile, FileToArray(CurrFile))
    Next CurrFile

    FilesToArray = FileArrays
End Function

Public Function TemporaryFile(Optional ByVal FileExtention As String = "tmp") As String
    TemporaryFile = Environ$("temp")
    If Len(TemporaryFile) = 0 Then TemporaryFile = App.Path
    If Right$(TemporaryFile, 1) <> "\" Then TemporaryFile = TemporaryFile & "\"

    TemporaryFile = Left$(Replace$(vbNullString & Rnd, ".", "-"), 8) & "." & FileExtention
End Function

Public Function VariantToBytes(ByRef vData As Variant) As Byte()
On Error Resume Next
    Dim sFilename As String

    sFilename = TemporaryFile
    If VariantToFile(sFilename, vData) Then
       VariantToBytes = FileToBytes(sFilename)
    End If
    Kill sFilename
End Function

Public Function BytesToVariant(ByRef bytX() As Byte) As Variant
On Error Resume Next
On Error GoTo 0
    Dim sFilename As String

    sFilename = TemporaryFile
    If BytesToFile(sFilename, bytX) Then
       BytesToVariant = FileToVariant(sFilename)
    End If
    Kill sFilename
End Function

Public Function FileToBytes(sFilename As String) As Byte()
On Error Resume Next
    Dim fh      As Long
    Dim bytX()  As Byte

    fh = FreeFile
    Open sFilename For Binary As #fh
         bytX = InputB(LOF(fh), fh)
    Close #fh
    FileToBytes = bytX
End Function

Public Function BytesToFile(sFilename As String, ByRef bytX() As Byte) As Boolean
On Error Resume Next
On Error GoTo 0
    Dim fh      As Long

    fh = FreeFile
    Err.Clear
    Open sFilename For Binary As #fh
         Put #fh, , bytX
    Close #fh
    BytesToFile = (Err.Number = 0)
    Err.Clear
End Function

Public Function FileToVariant(sFilename As String) As Variant
On Error Resume Next
On Error GoTo 0
    Dim fh    As Long
    Dim vData As Variant

    fh = FreeFile
    Open sFilename For Binary As #fh
         Get #fh, , vData
    Close #fh
    FileToVariant = vData
End Function

Public Function VariantToFile(ByVal sFilename As String, ByRef vData As Variant) As Boolean
On Error Resume Next
    Dim fh As Long
    Dim x As Byte
    
    Kill sFilename
    Err.Clear

On Error GoTo EH_VarToFile
    fh = FreeFile
    Open sFilename For Binary As #fh 'Len = 1024
         Put #fh, , vData
    Close #fh
    VariantToFile = (Err.Number = 0)
    Err.Clear

EH_VarToFile_Continue:
    Exit Function

EH_VarToFile:
    MsgBox "VarToFile " & Err.Number & ": " & Err.Description
    Resume EH_VarToFile_Continue
    
    Resume
End Function

Public Property Get Awareness() As String
    Awareness = RootTruth & "Systems\Chaotic\Awareness\"
End Property

Public Property Get Categorization() As String
    Categorization = RootTruth & "Systems\Chaotic\Categorization\"
End Property

Public Property Get Knowledge() As String
    Knowledge = RootTruth & "Systems\Chaotic\Knowledge\"
End Property

Public Property Get Manipulation() As String
    Manipulation = RootTruth & "Systems\Chaotic\Manipulation\"
End Property

Public Property Get InBuddha() As String
    InBuddha = RootTruth & "Patterns\Chaotic\Buddha\"
End Property

Public Property Get InAristotle() As String
    InAristotle = RootTruth & "Patterns\Chaotic\Aristotle\"
End Property

Public Property Get InLuxian() As String
    InLuxian = RootTruth & "Patterns\Chaotic\Luxian\"
End Property

Public Property Get InZen() As String
    InZen = RootTruth & "Patterns\Chaotic\Zen\"
End Property

Private Sub Class_Initialize()
    On Error Resume Next
    Dim lCrc32Value As Long

    Set colCache = New Collection
    
    RootTruth = "E:\Truth Labz\"

    lCrc32Value = InitCrc32()
End Sub

Private Sub Class_Terminate()
    Set colCache = New Collection
End Sub

Public Function ArraySearch(ByVal ToFind As String, ByRef ToSearch As Variant) As Long
    DataArray = ToSearch

    ArraySearch = BinarySearch(LBound(DataArray), UBound(DataArray), ToFind)

    ToSearch = DataArray
End Function

Private Function BinarySearch(ByRef Lower As Long, ByRef Upper As Long, Target As String) As Long
    Dim Middle As Long
    Middle = Fix(Lower + Upper) / 2

    If StrComp(DataArray(Middle), Target) = 0 Then
        BinarySearch = Middle
        Exit Function
    End If

    If Lower >= Upper Then
        BinarySearch = -1
        Exit Function
    End If

    If StrComp(Target, DataArray(Middle)) < 0 Then
        Upper = Middle + 1
    Else
        Lower = Middle + 1
    End If

    BinarySearch = BinarySearch(Lower, Upper, Target)
End Function

Public Function IsAbbreviation(ByVal ToCheck As String, ByVal AbbreviationPattern As String) As Boolean
    ' Return True if ToCheck abbreviates AbbreviationPattern
    '
    ' This function returns True if a test string in ToCheck
    ' abbreviates a master string in AbbreviationPattern...
    '
    ' The REQUIRED characters of AbbreviationPattern should be in
    ' upper case (the required characters of AbbreviationPattern
    ' must be alphabetic.) The optional characters of
    ' AbbreviationPattern must be nonalphabetic or lower case.
    '
    ' ToCheck may be in any combination of upper and lower
    ' case. Its beginning must match the required characters
    ' of the longstring, and characters after the beginning of
    ' the input string must match the corresponding characters
    ' of the longstring.
    '
    ' For example, IsAbbreviation("b", "Barf"), IsAbbreviation("kamunkle",
    ' "KAmunkle"), and IsAbbreviation("crunk", "CRUnk") will return True.
    ' IsAbbreviation("Jungly", "WUngly"), IsAbbreviation("ROO", "ROOBUGNA"),
    ' and IsAbbreviation("RUGnafoo", "RUGna") return False.

    Dim lngIndex1 As Long
    Dim lngLength As Long
    Dim lngMinAbbrev As Long
    Dim strCacheKey As String

    IsAbbreviation = False
    lngMinAbbrev = -1
    strCacheKey = "K" & AbbreviationPattern                ' Long string may be a number

    On Error Resume Next
    lngMinAbbrev = colCache(strCacheKey)
    On Error GoTo 0

    If lngMinAbbrev = -1 Then
        ' Longstring not in cache: determine mininum abbreviation
        ' and add to cache: boot the oldest entry if at limit

        For lngIndex1 = 1 To Len(AbbreviationPattern)
            If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid$(AbbreviationPattern, lngIndex1, 1)) = 0 Then
                Exit For
            End If
        Next lngIndex1

        If lngIndex1 = 1 Then
            ' MsgBox "Error in longstring passed to IsAbbreviation: longstring contains no upper-case leading characters: returning False"
            Exit Function
        End If

        lngMinAbbrev = lngIndex1 - 1

        If colCache.Count = 256 Then
            colCache.Remove 1
        End If
        colCache.Add lngMinAbbrev, strCacheKey
    End If

    lngLength = Len(ToCheck)
    If lngLength >= lngMinAbbrev Then
        If lngLength <= Len(AbbreviationPattern) Then
            IsAbbreviation = UCase$(ToCheck) = Mid$(UCase$(AbbreviationPattern), 1, lngLength)
        End If
    End If
End Function

Public Function CountInStr(ByRef ToSearch As String, ByRef ToFind As String, Optional ByVal bIgnoreCase As Boolean = True) As Long
    Dim CurIndex As Long

    If Len(ToSearch) = 0 Then Exit Function

    If bIgnoreCase Then
        CountInStr = UBound(Split(ToSearch, ToFind, , vbTextCompare))
        Exit Function
    End If

    Do: CurIndex = InStr(CurIndex + 1, ToSearch, ToFind, vbBinaryCompare)
        CountInStr = CountInStr + (CurIndex <> 0)
    Loop While CurIndex <> 0
End Function

Public Function GetTag(SourceString As String, Tag As String) As String
    If InStr(SourceString, "<" & Tag & ">") = 0 Then Exit Function
    GetTag = Mid$(SourceString, InStr(SourceString, "<" & Tag & ">"), InStr(SourceString, "</" & Tag & ">") + Len("</" & Tag & ">") - 1)
End Function

Public Function GetTagText(SourceString As String, Tag As String) As String
    If InStr(SourceString, "<" & Tag & ">") = 0 Then GetTagText = vbNullString
    GetTagText = Mid$(SourceString, InStr(SourceString, "<" & Tag & ">") + Len("<" & Tag & ">"), (InStr(SourceString, "</" & Tag & ">")) - (InStr(SourceString, "<" & Tag & ">") + Len("<" & Tag & ">")))
End Function

Public Function CutTag(SourceString As String, Tag As String) As String
    If InStr(SourceString, "<" & Tag & ">") = 0 Then Exit Function
    CutTag = Left$(SourceString, InStr(SourceString, "<" & Tag & ">") - 1) & Mid$(SourceString, InStrRev(SourceString, "</" & Tag & ">") + Len("</" & Tag & ">"))
End Function

Public Function InStrLike(Optional Start, Optional String1, Optional String2, Optional intCompareMethod As VbCompareMethod = vbTextCompare) As Variant
    On Error GoTo err_InStrLike
    Dim intPos      As Integer
    Dim intLength   As Integer
    Dim strBuffer   As String
    Dim blnFound    As Boolean
    Dim varReturn   As Variant

    If Not IsNumeric(Start) And IsMissing(String2) Then
        String2 = String1
        String1 = Start
        Start = 1
    End If

    If IsNull(String1) Or IsNull(String2) Then
        varReturn = Null
        GoTo exit_InStrLike
    End If

    If Left$(String2, 1) = "*" Then
        'Err.Raise vbObjectError + 2600, "InStrLike", "Comparison mask cannot start With '*' since a start position cannot be determined."
        Exit Function
    End If

    For intPos = Start To Len(String1) - Len(String2) + 1
        If InStr(1, String2, "*", vbTextCompare) Then
            For intLength = 1 To Len(String1) - intPos + 1
                strBuffer = Mid$(String1, intPos, intLength)
                If strBuffer Like String2 Then
                    blnFound = True
                    GoTo done
                End If
            Next intLength
        Else
            strBuffer = Mid$(String1, intPos, Len(String2))
            If strBuffer Like String2 Then
                blnFound = True
                GoTo done
            End If
        End If
    Next intPos

done:
    If blnFound = False Then
        varReturn = 0
    Else
        varReturn = intPos
    End If

exit_InStrLike:
    InStrLike = varReturn
    Exit Function

err_InStrLike:

    Select Case Err.Number
        Case Else
            varReturn = Null
            MsgBox Err.Description, vbCritical, "Error #" & Err.Number & "(InStrLike)"
            GoTo exit_InStrLike
    End Select
End Function

'   --------------- Begin CRC32

'// Then all we have to do is writing pu
'     blic functions like these...
Private Function InitCrc32(Optional ByVal Seed As Long = &HEDB88320, Optional ByVal Precondition As Long = &HFFFFFFFF) As Long
    '// Declare counter variable iBytes, cou
    '     nter variable iBits, value variables lCr
    '     c32 and lTempCrc32
    Dim iBytes As Integer, iBits As Integer, lCrc32 As Long, lTempCrc32 As Long
    '// Turn on error trapping
    On Error Resume Next
    '// Iterate 256 times


    For iBytes = 0 To 255
        '// Initiate lCrc32 to counter variable
        lCrc32 = iBytes
        '// Now iterate through each bit in coun
        '     ter byte


        For iBits = 0 To 7
            '// Right shift unsigned long 1 bit
            lTempCrc32 = lCrc32 And &HFFFFFFFE
            lTempCrc32 = lTempCrc32 \ &H2
            lTempCrc32 = lTempCrc32 And &H7FFFFFFF
            '// Now check if temporary is less than
            '     zero and then mix Crc32 checksum with Se
            '     ed value


            If (lCrc32 And &H1) <> 0 Then
                lCrc32 = lTempCrc32 Xor Seed
            Else
                lCrc32 = lTempCrc32
            End If
        Next
        '// Put Crc32 checksum value in the hold
        '     ing array
        Crc32Table(iBytes) = lCrc32
    Next
    '// After this is done, set function val
    '     ue to the precondition value
    InitCrc32 = Precondition
End Function
'// The function above is the initializi
'     ng function, now we have to write the co
'     mputation function


Public Function AddCrc32(ByVal Item As String, ByVal Crc32 As Long) As Long
    '// Declare following variables
    Dim bCharValue As Byte, iCounter As Integer, lIndex As Long
    Dim lAccValue As Long, lTableValue As Long
    '// Turn on error trapping
    On Error Resume Next
    '// Iterate through the string that is t
    '     o be checksum-computed


    For iCounter = 1 To Len(Item)
        '// Get ASCII value for the current char
        '     acter
        bCharValue = Asc(Mid$(Item, iCounter, 1))
        '// Right shift an Unsigned Long 8 bits
        lAccValue = Crc32 And &HFFFFFF00
        lAccValue = lAccValue \ &H100
        lAccValue = lAccValue And &HFFFFFF
        '// Now select the right adding value fr
        '     om the holding table
        lIndex = Crc32 And &HFF
        lIndex = lIndex Xor bCharValue
        lTableValue = Crc32Table(lIndex)
        '// Then mix new Crc32 value with previo
        '     us accumulated Crc32 value
        Crc32 = lAccValue Xor lTableValue
    Next
    '// Set function value the the new Crc32
    '     checksum
    AddCrc32 = Crc32
End Function
'// At last, we have to write a function
'     so that we can get the Crc32 checksum va
'     lue at any time


Public Function GetCrc32(ByVal Crc32 As Long) As Long
    '// Turn on error trapping
    On Error Resume Next
    '// Set function to the current Crc32 va
    '     lue
    GetCrc32 = Crc32 Xor &HFFFFFFFF
End Function
'// To Test the Routines Above...


'// This is the command that you would use to compute your own string
Public Function Compute(ToGet As String) As String
    Dim lCrc32Value As Long
    On Error Resume Next
    lCrc32Value = InitCrc32()
    lCrc32Value = AddCrc32(ToGet, lCrc32Value)
    Compute = Hex$(GetCrc32(lCrc32Value))
End Function

'   --------------- End CRC32

Public Function Simil(String1 As String, String2 As String) As Double
    Dim l1 As Long
    Dim l2 As Long
    Dim l  As Long
    Dim r   As Double

    If StrComp(String1, String2) = 0 Then
        r = 1
    Else
        l1 = Len(String1)
        l2 = Len(String2)

        If l1 = 0 Or l2 = 0 Then
           r = 0
        Else
           ReDim b1(1 To l1): ReDim b2(1 To l2)


           For l = 1 To l1
               b1(l) = Asc(Mid$(String1, l, 1))
           Next


           For l = 1 To l2
               b2(l) = Asc(Mid$(String2, l, 1))
           Next
            r = SubSim(1, l1, 1, l2) / (l1 + l2) * 2
        End If
    End If

    Simil = r
    Erase b1
    Erase b2
End Function

Private Function SubSim(st1 As Long, end1 As Long, st2 As Long, end2 As Long) As Long
    Dim c1  As Long
    Dim c2  As Long
    Dim ns1 As Long
    Dim ns2 As Long
    Dim i   As Long
    Dim max As Long

    If st1 > end1 Or st2 > end2 Or st1 <= 0 Or st2 <= 0 Then Exit Function

    For c1 = st1 To end1
        For c2 = st2 To end2
            i = 0
            Do Until b1(c1 + i) <> b2(c2 + i)
                i = i + 1
                If i > max Then
                    ns1 = c1
                    ns2 = c2
                    max = i
                End If
                If c1 + i > end1 Or c2 + i > end2 Then Exit Do
            Loop
        Next
    Next
    
    max = max + SubSim(ns1 + max, end1, ns2 + max, end2)
    max = max + SubSim(st1, ns1 - 1, st2, ns2 - 1)

    SubSim = max
End Function

Public Function sMassage(ByVal sToMassage As String, Optional ByVal sReplacement As String) As String
    Dim vInvalidChars   As Variant
    Dim CurrChar        As Variant

    For Each CurrChar In vInvalidChars
        sToMassage = Replace$(sToMassage, CurrChar, sReplacement)
    Next CurrChar
    
    sMassage = sToMassage
End Function

Public Function sGinsu(ByVal xmlFields As String, ByVal xmlValues As String, ByVal xmlWrapper As String, Optional ByVal FieldDelimiter As String = ";", Optional ByVal ValueDelimiter As String = vbNewLine) As String
On Error Resume Next
    Dim Result        As HConcat
    Dim FieldList     As Variant
    Dim ValueList     As Variant
    Dim CurrField     As Variant
    Dim CurrValue     As Long
    Dim MaxValue      As Long

    Set Result = New HConcat

    If Len(xmlWrapper) Then Result.Data = "<" & xmlWrapper & ">" & vbNewLine

    If Len(FieldDelimiter) = 0 Then FieldDelimiter = ";"
    If Len(ValueDelimiter) = 0 Then ValueDelimiter = vbNewLine
    
    FieldList = Split(xmlFields, FieldDelimiter)
    ValueList = Split(xmlValues, ValueDelimiter)

    CurrValue = -1
    MaxValue = UBound(ValueList)

    Do While CurrValue < MaxValue
       For Each CurrField In FieldList
           CurrValue = CurrValue + 1
           If CurrValue <= MaxValue Then
              Result.Concat vbTab & "<" & CurrField & ">" & ValueList(CurrValue) & "</" & CurrField & ">" & vbNewLine
           Else
              Result.Concat vbTab & "<" & CurrField & "></" & CurrField & ">" & vbNewLine
           End If
       Next CurrField
       Result.Concat "</" & xmlWrapper & ">" & vbNewLine
       If CurrValue < MaxValue Then Result.Concat "<" & xmlWrapper & ">" & vbNewLine
    Loop

    sGinsu = Result.Data
    Set Result = Nothing
End Function

Public Function UniqueWords(ByVal SortedWords As Variant) As Variant
    Dim Uniques         As Variant
    Dim TotalWords      As Long
    Dim TotalUniques    As Long
    Dim CurrentWord     As Long

    TotalWords = UBound(SortedWords) - 1
    ReDim Uniques(0 To (TotalWords \ 3)) As Variant

    For CurrentWord = 1 To TotalWords
        If StrComp(SortedWords(CurrentWord), SortedWords(CurrentWord + 1)) Then
           If TotalUniques > UBound(Uniques) Then
              ReDim Preserve Uniques(0 To TotalUniques + 1000) As Variant
           End If
           Uniques(TotalUniques) = SortedWords(CurrentWord)
           TotalUniques = TotalUniques + 1
        End If
    Next CurrentWord
    If TotalUniques > 0 Then
        TotalUniques = TotalUniques - 1
        ReDim Preserve Uniques(0 To TotalUniques) As Variant
        UniqueWords = Uniques
    Else
        UniqueWords = SortedWords
    End If
End Function

Public Sub QuickSort(ByRef arr As Variant, Optional numEls As Variant, Optional descending As Boolean)
    Dim leftStk(32)     As Long
    Dim rightStk(32)    As Long
    Dim leftNdx         As Long
    Dim rightNdx        As Long
    Dim SP              As Long
    Dim i               As Long
    Dim j               As Long

    Dim value           As Variant
    Dim temp            As Variant

  ' account for optional arguments
    If IsMissing(numEls) Then numEls = UBound(arr)

  ' init pointers
    leftNdx = LBound(arr)
    rightNdx = numEls

  ' init stack
    SP = 1
    leftStk(SP) = leftNdx
    rightStk(SP) = rightNdx

    Do
        If rightNdx > leftNdx Then
            value = arr(rightNdx)
            i = leftNdx - 1
            j = rightNdx
    
          ' find the pivot item
            If descending Then
                Do
                    Do: i = i + 1: Loop Until arr(i) <= value
                    Do: j = j - 1: Loop Until j = leftNdx Or arr(j) >= value
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                Loop Until j <= i
            Else
                Do
                    Do: i = i + 1: Loop Until arr(i) >= value
                    Do: j = j - 1: Loop Until j = leftNdx Or arr(j) <= value
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                Loop Until j <= i
            End If

          ' swap found items
            temp = arr(j)
            arr(j) = arr(i)
            arr(i) = arr(rightNdx)
            arr(rightNdx) = temp

          ' push on the stack the pair of pointers that differ most
            SP = SP + 1
            If (i - leftNdx) > (rightNdx - i) Then
                leftStk(SP) = leftNdx
                rightStk(SP) = i - 1
                leftNdx = i + 1
            Else
                leftStk(SP) = i + 1
                rightStk(SP) = rightNdx
                rightNdx = i - 1
            End If

        Else
          ' pop a new pair of pointers off the stacks
            leftNdx = leftStk(SP)
            rightNdx = rightStk(SP)
            SP = SP - 1
            If SP = 0 Then Exit Do
        End If
    Loop
End Sub

Public Function Cubize(ByVal ToCube As String, ByVal ItemLength As Long) As Variant
    Dim StrLen          As Long
    Dim CounterA        As Long
    Dim StartCount      As Long
    Dim LenDiff         As Long
    Dim LenTest         As Single
    Dim vtReturnArray   As Variant

    StrLen = Len(ToCube)                                 'get the length of our string
    LenTest = StrLen / ItemLength                          ' initialize LenTest

    StartCount = 1                                         'initialize StartCount
    If LenTest > Int(LenTest) Then LenTest = Int(LenTest) + 1    'we only want To work With Longs
    ReDim vtReturnArray(0 To LenTest - 1)                          'Size our array

    For CounterA = 0 To LenTest - 1
    vtReturnArray(CounterA) = Mid$(ToCube, StartCount, ItemLength)    'fill our array
    StartCount = StartCount + ItemLength               'increment our starting point in the Mid$ Function
    Next CounterA

    Cubize = vtReturnArray
End Function

Function WordCube(ByVal i_strToCube As String, ByVal i_lMaxLength) As Variant
    Dim strWork             As String
    Dim lCurrReturnIndex    As Long
    Dim vtReturnArray       As Variant

    If i_lMaxLength >= Len(i_strToCube) Then
        WordCube = Array(i_strToCube, "")
    Else
        vtReturnArray = Cubize(i_strToCube, i_lMaxLength)

        For lCurrReturnIndex = 0 To UBound(vtReturnArray)
            If (Right$(vtReturnArray(lCurrReturnIndex), 1) <> " ") And (lCurrReturnIndex <> UBound(vtReturnArray)) Then
                'Didn't break at the end of a word, fix it.
                strWork = vtReturnArray(lCurrReturnIndex)
                vtReturnArray(lCurrReturnIndex) = Split(strWork)

                strWork = vtReturnArray(lCurrReturnIndex)(UBound(vtReturnArray(lCurrReturnIndex)))
                vtReturnArray(lCurrReturnIndex)(UBound(vtReturnArray(lCurrReturnIndex))) = ""
                strWork = strWork & vtReturnArray(lCurrReturnIndex + 1)
                'BUGBUG: Works only for 2 items which is fine for now
                vtReturnArray(lCurrReturnIndex + 1) = Left$(strWork, i_lMaxLength)

                strWork = Trim$(Join(vtReturnArray(lCurrReturnIndex)))
                vtReturnArray(lCurrReturnIndex) = strWork
            Else
                'Broke at the end of a word or this is the last item, just trim it
                vtReturnArray(lCurrReturnIndex) = Trim$(vtReturnArray(lCurrReturnIndex))
            End If
        Next
        WordCube = vtReturnArray
    End If
End Function

Public Function vQuickSort(ByVal arr As Variant, Optional ByVal descending As Boolean) As Variant
    QuickSort arr, , descending
    vQuickSort = arr
End Function
