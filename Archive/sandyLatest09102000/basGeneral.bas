Attribute VB_Name = "modGeneral"

Option Explicit
' ================================================================================
' Class Module      modGeneral
'
' Filename          modGeneral.cls
'
' Author            William M. Rawls
'
' Created On        9/3/1997 8:00 pm
'
' Description
'
' General functions
'
' ================================================================================

' ========================
' Publicly available stuff
' ========================
' True if processing is occurring that should cause any cascading events to exit immediately (search for gbProcessing to see impact)
Public gbProcessing As Boolean

' True if the user cancel processing while doing an insertion
Public gbCancelInsertion As Boolean

' True if this copy is an evaluation and the designated evaluation period has ended.
Public gbEvaluationHasExpired As Boolean

' ================================================
' Publicly available constant strings
'   Call InitPublic() to set at beginning of program
' Why ? These strings are very common in VB
'   and using the Publicly available
' ================================================
Public gsEolTab   As String                           ' vbNewLine & vbTab
Public gs2EOL     As String                           ' vbNewLine & vbNewLine
Public gs2EOLTab  As String                           ' gs2EOLTab

Public Const gsE  As String = "="
Public Const gsA  As String = "'"
Public Const gsC  As String = ","
Public Const gsP  As String = "."
Public Const gsS  As String = " "
Public Const gsQ  As String = """"
Public Const gsSC As String = ";"
Public Const gsBS As String = "\"

Public Const gsBO As String = "{": Public Const gsBC As String = "}"
Public Const gsPO As String = "(": Public Const gsPC As String = ")"

Public Const gsFindBO = "Find{"
Public Const gsSelectFrom As String = "SELECT * FROM "
Public Const gsWhere              As String = " WHERE "

Public Const gsSoftVarDelimiter               As String = "%%"
Public Const gsSoftCmdDelimiter               As String = "~~"
Public Const gsInlineCmdDelimiter             As String = "::"
Public Const gsNormalizeDelimiter             As String = "%$%"
Public Const gsCategoryTemplateDelimiter      As String = " - "
Public Const gsSpecialLineItemDelimiter       As String = "**"
Public Const gsQueuedInsertionDelimiter       As String = "~!~!"

Public Const gsSliceAndDice       As String = "Slice and Dice"
Public Const gsTemplate           As String = "Template"
Public Const gsCategory           As String = "Category"

Public Const gsLast               As String = "Last"


' ===================================
' ====== BrowseForFolder stuff ======
' ===================================
Private Type BrowseInfo
    hWndOwner         As Long
    pIDLRoot          As Long
    pszDisplayName    As String
    lpszTitle         As String
    ulFlags           As Long
    lpfnCallback      As Long
    lParam            As Long
    iImage            As Long
End Type

Private Const BIF_RETURNONLYFSDIRS As Long = 1
Private Const MAX_PATH As Long = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

' ==========================================================
' API call to determin where the user's Windows directory is
' ==========================================================
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' *****************************************************************
' API call to determin where the user's Windows System directory is
' *****************************************************************
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

Public Function sMassage(ByVal sToMassage As String, Optional ByVal sReplacement As String = "_") As String
    Dim vInvalidChars As Variant
    Dim CurrChar As Variant

    vInvalidChars = Split("/,\,:,*,?,"",<,>,|", ",")

    For Each CurrChar In vInvalidChars
        sToMassage = Replace$(sToMassage, CurrChar, sReplacement)
    Next CurrChar
    
    sMassage = sToMassage
End Function

' -------------------------------------------------
' Calls the windows API to get the windows directory
' -------------------------------------------------
Public Property Get WindowsDirectory() As String
1        Dim X              As Integer
2        Dim sT             As String
         Static sWindowsDir As String

         If Len(sWindowsDir) Then
            WindowsDirectory = sWindowsDir
            Exit Property
         End If

3        sT = String$(145, 0)                              ' Size Buffer
4        X = GetWindowsDirectory(sT, 145)                  ' Make API Call
5        sT = Left$(sT, X)                                 ' Trim Buffer

6        If Right$(sT, 1) <> gsBS Then                     ' Add \ if necessary
7            WindowsDirectory = sT & gsBS
             sWindowsDir = sT & gsBS
8        Else
9            WindowsDirectory = sT
             sWindowsDir = sT
10       End If
End Property

Public Function FileExists(sFilename As String) As Boolean
11       On Error Resume Next
12       Err.Clear
13       FileExists = Len(Dir$(sFilename)) > 0
14       Err.Clear
End Function

Public Function GetListIndex(cboToSearch As Control, ByVal sItemToFind As String) As Integer
15       On Error Resume Next
16       Static nCurItem As Integer

17       If cboToSearch Is Nothing Then Exit Function

18       If Len(sItemToFind) = 0 Or cboToSearch.ListCount = 0 Then
19           GetListIndex = -1
20           Exit Function
21       End If

22       sItemToFind = UCase$(sItemToFind)

23       For nCurItem = 0 To cboToSearch.ListCount - 1
24           If StrComp(UCase$(cboToSearch.List(nCurItem)), sItemToFind) = 0 Then
25               GetListIndex = nCurItem
26               Exit Function
27           End If
28       Next nCurItem
End Function

Public Function EnumFiles(ByVal sPath As String, Optional ByVal sMask As String = "SAD*.dll", Optional ByVal Attr As VbFileAttribute = vbNormal) As String
29       Dim CurrFile As String
30       Dim sFileList   As String
         Dim sToRemove   As String ' Only used if hunting for directories and not files
         Dim ListToSift  As Variant
         Dim CurrSuspect As Variant
         Dim ExtraItems  As Variant
         Dim CurrExtra   As Variant
         
         Dim vM()        As Variant

31       If Len(sPath) = 0 Then sPath = CurDir
32       If Right$(sPath, 1) <> gsBS Then sPath = sPath & gsBS

33       CurrFile = Dir$(sPath & sMask, Attr)
34       sFileList = vbNullString

35       Do While Len(CurrFile)
36           sFileList = sFileList & CurrFile & gsSC
37           CurrFile = Dir
38       Loop

         If (Attr = vbDirectory) And (Len(sFileList) > 0) Then
          ' Repeat, but only look for files,
            CurrFile = Dir$(sPath & sMask)
            Do While Len(CurrFile)
               sToRemove = sToRemove & CurrFile & gsSC
               CurrFile = Dir
            Loop
            
          ' Merge the 2 lists,
            ListToSift = Split(Replace$("<<" & Join$(Split(sFileList, gsSC), ">>" & gsSC & "<<") & ">>", gsSC & "<<>>", vbNullString) _
                             & gsSC _
                             & Replace$("<<" & Join$(Split(sToRemove, gsSC), ">>" & gsSC & "<<") & ">>", gsSC & "<<>>", vbNullString), ">>;<<")

            ReDim vM(0 To (UBound(ExtraItems) - UBound(ListToSift)))            ' # files = # found - # directories
            
            ExtraItems = Split(Replace$("<<" & Join$(Split(ExtraItems, gsSC), ">>;<<") & ">>", gsSC & "<<>>", vbNullString), ">>;<<")

            sFileList = vbNullString
            For Each CurrExtra In ExtraItems
                ListToSift = Filter(ListToSift, CurrExtra, False)
            Next CurrExtra
            sFileList = Replace$(Join$(ListToSift, gsE & vbNewLine), vbNewLine & ".=" & vbNewLine, vbNewLine & ".=Current Directory" & vbNewLine)
        
            
          '   excluding all from list 1 that are in list 2,
          '   also, include descriptions for "." and ".."
          ' leaving only a purely inclusive and walkable list of directories !
          
         End If

39       If sMask = "SAD*.dll" Then
40           EnumFiles = Replace(sFileList, ".dll;", ".NewCommands=Load" & vbNewLine)
41       Else
42           EnumFiles = sFileList
43       End If

End Function

Public Function LogError(ByVal sModuleName As String, ByVal sProcName As String, lError As Long, ByVal sErrorMsg As String, ByVal LineNumber As Long, Optional ByVal sQuestion As String) As Boolean
44       Dim fh       As Long
45       Dim sMessage As String

46       fh = FreeFile

47       sMessage = "***** Error Occured At: " & Format$(Now(), "MM/DD/YYYY HH:MM:SS AM/PM")
48       sMessage = sMessage & vbNewLine & "  Module    : " & sModuleName
49       sMessage = sMessage & vbNewLine & "  Procedure : " & sProcName
50       sMessage = sMessage & vbNewLine & "  Error     : " & lError
51       If LineNumber > 0 Then
52           sMessage = sMessage & vbNewLine & "  Last Line # " & LineNumber
53       End If
54       sMessage = sMessage & vbNewLine & "  Details   :" & vbNewLine & sErrorMsg
55       If Len(sQuestion) > 0 Then
56           sMessage = sMessage & vbNewLine & vbNewLine & vbTab & sQuestion
57           LogError = bUserSure(sMessage, "SLICE AND DICE ERROR TRAP")
58       Else
59           MsgBox sMessage
60           LogError = True
61       End If

62       Open "sadErrors.log" For Append As #fh
63       Print #fh, sMessage
64       If LogError Then
65           If Len(sQuestion) Then
66               Print #fh, "***** " & gsSliceAndDice & " was continued by user after error."
67           Else
68               Print #fh, "***** " & gsSliceAndDice & " automatically continued after error."
69               LogError = False                          ' For compatibility sake
70           End If
71       Else
72           Print #fh, "***** User decided not to continue the " & gsSliceAndDice & " operation after error."
73       End If
74       Close #fh

       ' Stop
End Function

Public Function FindInCollection(colToFindIn As Object, sToFind As String) As Object
75       Dim CurItem As Object                             'VBComponent
76       For Each CurItem In colToFindIn
77           If StrComp(CurItem.Name, sToFind, vbTextCompare) = 0 Then
78               Set FindInCollection = CurItem
79               Exit Function
80           End If
81       Next CurItem
82       Set FindInCollection = Nothing
End Function

Public Function sFileContents(ByVal sPathAndFilename As String) As String
83       On Error Resume Next
84       Dim fh As Long
85       If Len(Dir$(sPathAndFilename)) Then
86           fh = FreeFile
87           Open sPathAndFilename For Input Access Read As #fh
88           sFileContents = Input$(LOF(fh), fh)
89           Close #fh
90       End If
End Function

Public Function sGetGUID(ByVal sProgID As String) As String
91       On Error Resume Next
92       sGetGUID = GetStringValue("HKEY_CLASSES_ROOT" & gsBS & sProgID & gsBS & "CLSID", vbNullString)
End Function

Public Function sChoose(sChoices As String, Optional ByVal sDelimiter As String = gsSC, Optional ByVal sDefault As String)
93       On Error GoTo EH_Wizard_sChoose
94       Static bInHereAlready As Boolean
95       If bInHereAlready Then Exit Function
96       bInHereAlready = True

97       If Len(sDelimiter) = 0 Then sDelimiter = gsSC

98       Dim frmX As New frmListSelect
99       With frmX
100          .Initialize sChoices, sDelimiter, sDefault
101          .ZOrder
102          .Show vbModal
103          sChoose = .Choice
104      End With

105 EH_Wizard_sChoose_Continue:
106      bInHereAlready = False
107      Exit Function

108 EH_Wizard_sChoose:
109      LogError "modGeneral", "sChoose", Err.Number, Err.Description, Erl
110      Resume EH_Wizard_sChoose_Continue

111      Resume
End Function

Public Function FindListIndex(lvwToSearch As Control, sToFind As String) As Long
112      Dim i As Long

113      If Len(sToFind) > 0 Then
114          For i = 0 To lvwToSearch.ListCount - 1
115              If StrComp(lvwToSearch.List(i), sToFind, vbTextCompare) = 0 Then
116                  FindListIndex = i
117                  Exit Function
118              End If
119          Next i
120      End If
121      FindListIndex = -1
End Function

Public Sub SetListIndex(cboToSearch As Control, sToFind As String)
122      Dim nIndex As Long

123      nIndex = FindListIndex(cboToSearch, sToFind)
124      If nIndex > -1 Then cboToSearch.ListIndex = nIndex
End Sub

Public Sub SetListViewIndex(lvwToSearch As ListView, sToFind As String)
'Dim nIndex As Long

125      FindListViewIndex lvwToSearch, sToFind, True

    'nIndex = FindListViewIndex(lvwToSearch, sToFind, True)
    'If nIndex > -1 Then lvwToSearch.ListIndex = nIndex
End Sub

Public Function FindListViewIndex(lvwToSearch As ListView, sToFind As String, Optional ByVal bSelectOnFind As Boolean) As Long
126      Dim i As Long
127      Dim CurrItem As ListItem

128      If Len(sToFind) > 0 Then
129          i = 0
130          For Each CurrItem In lvwToSearch.ListItems
131              i = i + 1
132              If StrComp(CurrItem.Text, sToFind, vbTextCompare) = 0 Then
133                  FindListViewIndex = i
134                  If bSelectOnFind Then
135                      CurrItem.Selected = True
136                  End If
137                  Exit Function
138              End If
139          Next CurrItem
140      End If
141      FindListViewIndex = -1
End Function

' On Return:
'    sExtractToken =
'       Token # nToken
'    sOrigStr =
'       Token # (1 through nToken-1) & (nToken+1 through nTokenCount(sOrigStr))
'
' Notes:
'    Yes there are a lot of "Exit Function" statements...
'       That's to make the function return as quickly as possible
'
Public Function sExtractToken(ByRef sOrigStr As String, Optional ByVal nToken As Integer = 1, Optional ByVal strDelim As String = gsS)
142      Static strIn As String
143      Static strOut As String
144      Static nCurrTokenStart As Long
145      Static nNextTokenStart As Long
146      Static nLenDelim As Long

    ' Handle the "simple" cases (No delimiter, or token # less than 2)
147      nLenDelim = Len(strDelim)
148      If nToken < 1 Or nLenDelim = 0 Then
        ' Nothing to extract, return nothing
149          Exit Function
150      ElseIf nToken = 1 Then
151          nCurrTokenStart = InStr(sOrigStr, strDelim)
152          If nCurrTokenStart > 0 Then
153              sExtractToken = Left$(sOrigStr, nCurrTokenStart - 1)
154              sOrigStr = Trim$(Mid$(sOrigStr, nCurrTokenStart + nLenDelim))
155              Exit Function
156          Else
157              sExtractToken = sOrigStr
158              sOrigStr = vbNullString
159              Exit Function
160          End If
161      End If

    ' Find the start of then nToken'th Token
162      strIn = sOrigStr: strOut = vbNullString
163      nToken = nToken - 1
164      Do Until nToken = 0
165          nCurrTokenStart = InStr(strIn, strDelim)
166          If nCurrTokenStart = 0 Or Len(strIn) = 0 Then Exit Function
167          strOut = strOut & Left$(strIn, nCurrTokenStart - 1)
168          strIn = Mid$(strIn, nCurrTokenStart + nLenDelim)

        ' Check to see if this is the one the calling function is looking for
169          nToken = nToken - 1
170      Loop

    ' Now we're at the point" & gsWhere & "the token sought for resides
171      nCurrTokenStart = InStr(strIn, strDelim)
172      If nCurrTokenStart > 0 Then
173          If nCurrTokenStart > 1 Then
174              sExtractToken = Left$(strIn, nCurrTokenStart - 1)
175          Else
176              sExtractToken = vbNullString
177          End If
        ' Rewrite the original string without the last token
178          sOrigStr = Trim$(strOut & Mid$(strIn, nCurrTokenStart))
179          Exit Function
180      Else
181          sExtractToken = strIn
182          sOrigStr = Trim$(strOut)
183          Exit Function
184      End If
End Function

' ================================================================================
' Name              BrowseForFolder
'
' Parameters
'      hWndOwner                     (I)  Window handle of owner
'      sPrompt                       (I)  Browse window caption
'
' Description
'
' Allows the user to "browse" for a directory (32 bit only!!!)
'
' ================================================================================
Public Function BrowseForFolder(ByVal hWndOwner As Long, ByVal sPrompt As String) As String
185      On Error GoTo ErrorHandler
186      Dim iNull    As Long
187      Dim lpIDList As Long
188      Dim lResult  As Long
189      Dim sPath    As String
190      Dim udtBI    As BrowseInfo

191      With udtBI
192          .hWndOwner = hWndOwner
193          .lpszTitle = sPrompt
194          .ulFlags = BIF_RETURNONLYFSDIRS
195      End With

196      lpIDList = SHBrowseForFolder(udtBI)
197      If lpIDList Then
198          sPath = String$(MAX_PATH, 0)
199          lResult = SHGetPathFromIDList(lpIDList, sPath)
200          Call CoTaskMemFree(lpIDList)
201          iNull = InStr(sPath, vbNullChar)
202          If iNull Then
203              sPath = Left$(sPath, iNull - 1)
204          End If
205      End If

206 Done:
207      BrowseForFolder = sPath
208      Exit Function

209 ErrorHandler:
210      LogError "modGeneral", "BrowseForFolder", Err.Number, Err.Description, Erl
211      Resume Done

212      Resume
End Function


' ================================================================================
' Name              bUserSure
'
' Parameters
'       sPrompt                      (I)  Opt. Question to ask the user
'                                         Default = "Are you sure this is what you want to do ?"
' Description
'
' Returns true if the user selects "Yes" from the MsgBox displayed
'
' ================================================================================
Public Function bUserSure(Optional ByVal sPrompt As String = "Are you sure this is what you want to do ?", Optional ByVal sTitle As String = "ARE YOU SURE ?") As Boolean
213      bUserSure = (MsgBox(sPrompt, vbYesNo, sTitle) = vbYes)
End Function

' ================================================================================
' Name              NextNegativeUnique
'
' Parameters
'      None
'
' Description
'
' Used to return a unique negative number. Numbers are unique to the current
' program session only.
'
' ================================================================================
Public Function NextNegativeUnique() As Long
214      Static lNextSerial As Long
215      lNextSerial = lNextSerial - 1
216      NextNegativeUnique = lNextSerial
End Function

' ================================================================================
' Synopsis          Returns the number of tokens as delimited by siDelim
'
' Parameters
'
'   siAllTokens                 (I) Required. The string containing all the tokens
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = gsS
' Description
'  For the following:
'    siAllTokens         sDelim  Returns       Notes
'   "William M Rawls"    gsS     3             "William", "M", and "Rawls"
'   "William M Rawls"    "iam"   2             "Will" and " M Rawls"
'   "William M Rawls"    vbNullString      1             No delimiter? String has one token,
'                                              "William M Rawls"
'   "1.00.05"            "."     3             "1", "00", and "05"
' ================================================================================
Public Function lTokenCount(ByVal siAllTokens As String, Optional ByVal siDelim As String = gsS) As Long
217      Static lCurTokenLocation As Long                  ' Character position of the first delimiter string
218      Static lTokensSoFar As Long                       ' Used to keep track of how many tokens we've counted so far
219      Static lDelim As Long                             ' Length of the delimiter string

220      lDelim = Len(siDelim)
221      If lDelim < 1 Then
        ' Empty delimiter strings means only one token equal to the string
222          lTokenCount = 1
223          Exit Function
224      ElseIf Len(siAllTokens) = 0 Then
        ' Empty input string means no tokens
225          Exit Function
226      Else
        ' Count the number of tokens
227          lTokensSoFar = 0
228          Do
229              lCurTokenLocation = InStr(siAllTokens, siDelim)
230              If lCurTokenLocation = 0 Then
231                  lTokenCount = lTokensSoFar + 1        'Abs(Len(siAllTokens) > 0)
232                  Exit Function
233              End If
234              lTokensSoFar = lTokensSoFar + 1
235              siAllTokens = Mid$(siAllTokens, lCurTokenLocation + lDelim)
236          Loop
237      End If
End Function
' ================================================================================
' Name              nz
'
' Parameters
'      vData                         (O)  Variant to test for NULL
'       sDefault                     (O)  Opt. On NULL this string is returned
'                                         Default = vbNullString
' Description
'
' Returns sDefault if the variant is NULL, otherwise it returns the Variant
'
' ================================================================================
Public Function nZ(ByRef vData As Variant, Optional sDefault As String = vbNullString) As String
238      If IsNull(vData) Then
239          nZ = sDefault
240      Else
241          nZ = vData
242      End If
End Function

Public Function sDenormalize(sLine As String) As String
243      sDenormalize = Replace(Replace(sLine, gsNormalizeDelimiter & "EOL" & gsNormalizeDelimiter, vbNewLine), gsNormalizeDelimiter & "TAB" & gsNormalizeDelimiter, vbTab)
End Function

' ================================================================================
' Synopsis          Returns the Nth Token from siAllTokens delimited by sDelim
'
' Parameters
'
'   siAllTokens                 (I) Required. The string containing all the tokens
'   iToken                      (I) Optional. The index of the token to return
'                                   DEFAULT = 1
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = gsS
' Description
'  For the following:
'    siAllTokens         iToken   sDelim  Returns       Notes
'   "William M Rawls"    1       gsS     "William"      First word
'   "William M Rawls"    2       gsS     "M"            Second word
'   "William M Rawls"    3       gsS     "Rawls"        Third word
'   "William M Rawls"    4       gsS     vbNullString             No forth word
'   "William M Rawls"    0       gsS     vbNullString             Zeroth token is always empty
'   "William M Rawls"   -1       gsS     vbNullString             Negative tokesn always empty
'   "William M Rawls"    1       vbNullString      vbNullString             No delimiter ? Token empty
' ================================================================================
Public Function sGetToken(ByVal siAllTokens As String, Optional ByVal iToken As Long = 1, Optional ByVal sDelim As String = gsS) As String
244      Static iCurTokenLocation As Long                  ' Character position of the first delimiter string
245      Static nDelim As Long                             ' Length of the delimiter string
246      nDelim = Len(sDelim)

247      If iToken < 1 Or nDelim < 1 Then
        ' Negative or zeroth token or empty delimiter strings mean an empty token
248          Exit Function
249      ElseIf iToken = 1 Then
        ' Quickly extract the first token
250          iCurTokenLocation = InStr(siAllTokens, sDelim)
251          If iCurTokenLocation > 1 Then
252              sGetToken = Left$(siAllTokens, iCurTokenLocation - 1)
253          ElseIf iCurTokenLocation = 1 Then
254              sGetToken = vbNullString
255          Else
256              sGetToken = siAllTokens
257          End If
258          Exit Function
259      Else
        ' Find the Nth token
260          Do
261              iCurTokenLocation = InStr(siAllTokens, sDelim)
262              If iCurTokenLocation = 0 Then
263                  Exit Function
264              Else
265                  siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
266              End If
267              iToken = iToken - 1
268          Loop Until iToken = 1

        ' Extract the Nth token (Which is the next token at this point)
269          iCurTokenLocation = InStr(siAllTokens, sDelim)
270          If iCurTokenLocation > 0 Then
271              sGetToken = Left$(siAllTokens, iCurTokenLocation - 1)
272              Exit Function
273          Else
274              sGetToken = siAllTokens
275              Exit Function
276          End If
277      End If
End Function
' ================================================================================
' Synopsis          Returns everything AFTER the Nth Token from siAllTokens delimited by sDelim
'
' Parameters
'
'   siAllTokens                 (I) Required. The string containing all the tokens
'   iToken                      (I) Optional. The index of the token to use as an "after" ref
'                                   DEFAULT = 1
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = gsS
' Description
'  For the following:
'    siAllTokens         iToken   sDelim  Returns           Notes
'   "William M Rawls"    1       gsS     "M Rawls"          After the first word
'   "William M Rawls"    2       gsS     "Rawls"            After the second word
'   "William M Rawls"    3       gsS     vbNullString                 After the third word (nothing)
'   "William M Rawls"    0       gsS     "William M Rawls"  After zeroth token is always the input string
'   "William M Rawls"   -1       gsS     "William M Rawls"  Negative tokens act same as zero
'   "William M Rawls"    1       vbNullString      "William M Rawls"  Same as one
' ================================================================================
Public Function sAfter(ByVal siAllTokens As String, Optional ByVal iToken As Long = 1, Optional ByVal sDelim As String = gsS) As String
278      Static iCurTokenLocation As Long                  ' Character position of the first delimiter string
279      Static nDelim As Long                             ' Length of the delimiter string

280      nDelim = Len(sDelim)
281      If iToken < 1 Or nDelim < 1 Then
        ' Negative or zeroth token or empty delimiter strings mean an empty token
282          sAfter = siAllTokens
283          Exit Function
284      ElseIf iToken = 1 Then
        ' Quickly extract the first token
285          iCurTokenLocation = InStr(siAllTokens, sDelim)
286          If iCurTokenLocation > 1 Then
287              sAfter = Mid$(siAllTokens, iCurTokenLocation + nDelim)
288              Exit Function
289          ElseIf iCurTokenLocation = 0 Then
290              sAfter = vbNullString
291              Exit Function
292          Else
293              sAfter = Mid$(siAllTokens, nDelim + 1)
294              Exit Function
295          End If
296      Else
        ' Find the Nth token
297          Do
298              iCurTokenLocation = InStr(siAllTokens, sDelim)
299              If iCurTokenLocation = 0 Then
300                  Exit Function
301              Else
302                  siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
303              End If
304              iToken = iToken - 1
305          Loop Until iToken = 1

        ' Extract the Nth token (Which is the next token at this point)
306          iCurTokenLocation = InStr(siAllTokens, sDelim)
307          If iCurTokenLocation > 0 Then
308              sAfter = Mid$(siAllTokens, iCurTokenLocation + nDelim)
309              Exit Function
310          Else
311              Exit Function
312          End If
313      End If
End Function

' ================================================================================
' Synopsis          Returns everything BEFORE the Nth Token from siAllTokens delimited by sDelim
'
' Parameters
'
'   siAllTokens                 (I) Required. The string containing all the tokens
'   iToken                      (I) Optional. The index of the token to use as a "before" ref
'                                   DEFAULT = 2
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = gsS (Space)
' Description
'  For the following:
'    siAllTokens         iToken   sDelim  Returns           Notes
'   "William M Rawls"    1       gsS     vbNullString                 Before the first word (nothing)
'   "William M Rawls"    2       gsS     "William"          Before the second word
'   "William M Rawls"    3       gsS     "William M"        Before the third word
'   "William M Rawls"    0       gsS     vbNullString                 Before zeroth token (nothing)
'   "William M Rawls"   -1       gsS     vbNullString                 Negative tokens act same as zero
'   "William M Rawls"    1       vbNullString      vbNullString                 Same as one
' ================================================================================
Public Function sBefore(ByVal siAllTokens As String, Optional ByVal iToken As Long = 2, Optional ByVal sDelim As String = gsS) As String
314      Static iCurTokenLocation As Long                  ' Character position of the first delimiter string
315      Static nDelim As Long                             ' Length of the delimiter string
316      Static sReturned As String

317      nDelim = Len(sDelim)
318      If iToken < 2 Or nDelim < 1 Then
        ' First, Zeroth, or Negative tokens or empty delimiter strings mean an empty string returned
319          sBefore = vbNullString
320          Exit Function
321      ElseIf iToken = 2 Then
        ' Quickly extract the first token
322          sBefore = sGetToken(siAllTokens, 1, sDelim)
323          Exit Function
324      Else
        ' Find the Nth token
325          Do
326              iCurTokenLocation = InStr(siAllTokens, sDelim)
327              If iCurTokenLocation = 0 Or iToken = 1 Then
328                  sBefore = sReturned
329                  sReturned = vbNullString
330                  Exit Function
331              ElseIf Len(sReturned) = 0 Then
332                  sReturned = Left$(siAllTokens, iCurTokenLocation - 1)
333              Else
334                  sReturned = sReturned & sDelim & Left$(siAllTokens, iCurTokenLocation - 1)
335              End If
336              siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
337              iToken = iToken - 1
338          Loop
339      End If
End Function
' ================================================================================
' Synopsis          Returns everything EXCEPT the Nth Token from siAllTokens delimited by sDelim
'
' Parameters
'
'   siAllTokens                 (I) Required. The string containing all the tokens
'   iToken                      (I) Optional. The index of the token to exclude
'                                   DEFAULT = 1
'   siDelim                     (I) Optional. The delimiter string that separates
'                                   the tokens. DEFAULT = gsS (Space)
' Description
'  For the following:
'    siAllTokens         iToken   sDelim  Returns           Notes
'   "William M Rawls"    1       gsS     "M Rawls"          After 1st token
'   "William M Rawls"    2       gsS     "William Rawls"    1st and 3rd token
'   "William M Rawls"    3       gsS     "William M"        Before the third word
'   "William M Rawls"    0       gsS     "William M Rawls"  Everything except 0th token (everything)
'   "William M Rawls"   -1       gsS     vbNullString                 Negative tokens act same as zero
'   "William M Rawls"    1       vbNullString      "William M Rawls"  Same as zero
' ================================================================================
Public Function sExcept(ByVal siAllTokens As String, Optional ByVal iToken As Long = 1, Optional ByVal sDelim As String = gsS) As String
340      Static iCurTokenLocation As Long                  ' Character position of the first delimiter string
341      Static nDelim As Long                             ' Length of the delimiter string
342      Static sReturned As String

343      nDelim = Len(sDelim)
344      If iToken < 1 Or nDelim < 1 Then
        ' First, Zeroth, or Negative tokens or empty delimiter strings mean an empty string returned
345          sExcept = siAllTokens
346          Exit Function
347      ElseIf iToken = 1 Then
        ' Quickly Return after token 1
348          iCurTokenLocation = InStr(siAllTokens, sDelim)
349          If iCurTokenLocation = 0 Then
350              sExcept = siAllTokens
351              Exit Function
352          Else
353              sExcept = Mid$(siAllTokens, iCurTokenLocation + nDelim)
354              Exit Function
355          End If
356      Else
        ' Find the Nth token
357          Do
358              iCurTokenLocation = InStr(siAllTokens, sDelim)
359              If iToken = 1 Then
360                  If iCurTokenLocation > 0 Then
361                      sExcept = sReturned & sDelim & Mid$(siAllTokens, iCurTokenLocation + nDelim)
362                  Else
363                      sExcept = sReturned
364                  End If
365                  sReturned = vbNullString
366                  Exit Function
367              ElseIf iCurTokenLocation = 0 Then
368                  sExcept = sReturned & sDelim & siAllTokens
369                  sReturned = vbNullString
370                  Exit Function
371              ElseIf Len(sReturned) = 0 Then
372                  sReturned = Left$(siAllTokens, iCurTokenLocation - 1)
373              Else
374                  sReturned = sReturned & sDelim & Left$(siAllTokens, iCurTokenLocation - 1)
375              End If
376              siAllTokens = Mid$(siAllTokens, iCurTokenLocation + nDelim)
377              iToken = iToken - 1
378          Loop
379      End If
End Function
' ================================================================================
' Name              SliceAndDice.modGeneral_sInsertSpaces
'
' Parameters
'      sToInsertInto                 (I)  String.
'
' Returns
'      String                        .
'
' Description
'
' Inserts spaces into a string. Spaces are inserted at each capital letter.
' Common words not to be capitalized are handled automatically (a, an, the, or,
' of). The sub string "ID" is handled specially and is removed. This function is
' excellent for labels.
'
' ================================================================================
Public Function sInsertSpaces(ByVal sToInsertInto As String) As String
380      Dim bytOriginal() As Byte
381      Dim sWithSpaces As String
382      Dim nUpper As Long
383      Dim nCurrent As Long
384      Dim nA As Byte
385      Dim nZ As Byte

386      bytOriginal = StrConv(sToInsertInto, vbFromUnicode)
387      nUpper = UBound(bytOriginal)

388      For nCurrent = 0 To nUpper
389          If bytOriginal(nCurrent) >= 65 And bytOriginal(nCurrent) <= 90 And nCurrent <> 0 Then
390              sWithSpaces = sWithSpaces & gsS & Chr$(bytOriginal(nCurrent))
391          Else
392              sWithSpaces = sWithSpaces & Chr$(bytOriginal(nCurrent))
393          End If
394      Next nCurrent

395      sInsertSpaces = Replace(Replace(Replace(Replace(Replace(Replace(sWithSpaces, " Of ", " of "), " The ", " the "), " A ", " a "), " An ", " an "), " I D", vbNullString), "  ", gsS)
End Function

Public Function sNormalize(sLine As String) As String
396      sNormalize = Replace(Replace(sLine, vbNewLine, gsNormalizeDelimiter & "EOL" & gsNormalizeDelimiter), vbTab, gsNormalizeDelimiter & "TAB" & gsNormalizeDelimiter)
End Function

Public Function sTableToPropertyName(ByVal sTableName As String) As String
397      sTableToPropertyName = Replace(Replace(Replace(Replace(sTableName, gsS, "_"), "*", "_"), "-", "_"), gsP, "__")
End Function

Public Sub BrowseTo(sURL As String)
398      On Error Resume Next
399      Static WinVer As String
400      Static WebBrowserCommand As String

401      If Len(WinVer) = 0 Then
402          WinVer = GetStringValue("HKEY_LOCAL_MACHINE" & gsBS & "SOFTWARE" & gsBS & "Microsoft" & gsBS & "Windows" & gsBS & "CurrentVersion", "Version")
403          If WinVer = "Error" Then
404              WinVer = GetStringValue("HKEY_LOCAL_MACHINE" & gsBS & "SOFTWARE" & gsBS & "Microsoft" & gsBS & "Windows NT" & gsBS & "CurrentVersion", "CurrentVersion")
405          End If
406          If WinVer = "Windows 98" Then
407              WebBrowserCommand = "start "
408          Else
409              WebBrowserCommand = GetStringValue("HKEY_CLASSES_ROOT" & gsBS & "htmlfile" & gsBS & "shell" & gsBS & "open" & gsBS & "command", vbNullString) & gsS
410              If WebBrowserCommand = "Error " Then
411                  WebBrowserCommand = gsQ & "C:" & gsBS & "PROGRA~1" & gsBS & "Plus!" & gsBS & "MICROS~1" & gsBS & "iexplore.exe -nohome "
412              End If
413          End If
414      End If
415      Shell WebBrowserCommand & sURL, vbNormalFocus
End Sub

' ================================================================================
' Name              zn
'
' Parameters
'      sData                         (I)  String to test for empty
'
' Description
'
' If the Variant passed in is an empty string, this procedure returns NULL,
' otherwise it returns the string passed.
'
' ================================================================================
Public Function zn(sData As String) As Variant
416      If Len(sData) = 0 Then zn = Null Else zn = sData
End Function

' ================================================================================
' Name              Class_Initialize
'
' Parameters
'      None
'
' Description
'
' Initializes the Public strings and the first Negative returned value to -2
'
' ================================================================================
Public Sub InitPublic()
417      gsEolTab = vbNewLine & vbTab
418      gs2EOL = vbNewLine & vbNewLine
419      gs2EOLTab = gs2EOL & vbTab

420      Call NextNegativeUnique                           ' Sets first time to -2 vs -1 (since -1 is usually used to indicated that a new negative # is needed)
End Sub

Public Function LoadFormPosition(frmToActOn As Form, Optional ByVal bAutoCenter = True, Optional ByVal bRemeberWidth As Boolean = True, Optional ByVal sSectionName As String)
421      Dim ProductName As String
422      Dim SectionName As String

423      If Len(App.ProductName) = 0 Then
424          ProductName = gsSliceAndDice
425      Else
426          ProductName = App.ProductName
427      End If

428      With frmToActOn
429          If Len(sSectionName) Then
430              SectionName = sSectionName
431          Else
432              SectionName = .Name
433          End If
434          If GetSetting(ProductName, SectionName, "Position Saved", False) Then
435              .Left = GetSetting(ProductName, SectionName, "Form Position Left", .Left)
436              .Top = GetSetting(ProductName, SectionName, "Form Position Top", .Top)
437              If bRemeberWidth Then .Width = GetSetting(ProductName, SectionName, "Form Position Width", .Width)
438              .Height = GetSetting(ProductName, SectionName, "Form Position Height", .Height)
439          ElseIf bAutoCenter Then
440              .Left = (Screen.Width - .Width) / 2
441              .Top = (Screen.Height - .Height) / 2
442          End If

        ' Ensure it'll fit on the screen (screen resolution change ?)
443          If .Left > Screen.Width Then .Left = 0
444          If .Top > Screen.Height Then .Top = 0
445          If .Left + .Width > Screen.Width Then .Width = Screen.Width - .Left
446          If bRemeberWidth Then
447              If .Top + .Height > Screen.Height Then .Height = Screen.Height - .Top
448          End If
449      End With
End Function


Public Function SaveFormPosition(frmToActOn As Form, Optional ByVal sSectionName As String)
450      Dim ProductName As String
451      Dim SectionName As String

452      If Len(App.ProductName) = 0 Then
453          ProductName = gsSliceAndDice
454      Else
455          ProductName = App.ProductName
456      End If

457      With frmToActOn
458          If Len(sSectionName) Then
459              SectionName = sSectionName
460          Else
461              SectionName = .Name
462          End If
463          SaveSetting ProductName, SectionName, "Position Saved", True
464          SaveSetting ProductName, SectionName, "Form Position Left", .Left
465          SaveSetting ProductName, SectionName, "Form Position Top", .Top
466          SaveSetting ProductName, SectionName, "Form Position Width", .Width
467          SaveSetting ProductName, SectionName, "Form Position Height", .Height
468      End With
End Function

Public Function lFindToken(ByVal sAllTokens As String, ByVal sTokenToFind As String, Optional ByVal sDelimiter As String = gsS) As Long
469      Dim lTokens As Long
470      Dim l       As Long

471      lTokens = lTokenCount(sAllTokens, sDelimiter)

472      For l = 1 To lTokens
473          If StrComp(UCase$(sGetToken(sAllTokens, l, sDelimiter)), UCase$(sTokenToFind)) = 0 Then
474              lFindToken = l
475              Exit Function
476          End If
477      Next l

478      lFindToken = 0
End Function

Public Function StringToClipboard(ByVal sTextToPutOnClipboard As String) As Boolean
479      On Error Resume Next
480      If Len(sTextToPutOnClipboard) = 0 Then StringToClipboard = True: Exit Function

481      Err.Clear
482      Clipboard.Clear

483      If Err.Number = 0 Then
484          Clipboard.SetText sTextToPutOnClipboard, vbCFText
485          If Err.Number = 0 Then
486              StringToClipboard = True
487          Else
488              LogError "modGeneral", "StringToClipboard", Err.Number, "Error putting Text onto the Clipboard. Error Description = " & Err.Description, Erl
489          End If

490      Else
491          LogError "modGeneral", "StringToClipboard", Err.Number, "Error putting Text onto the Clipboard. Error Description = " & Err.Description, Erl
492      End If
End Function

Public Function SaveToFile(ByVal sFilename As String, ByVal sContents As String) As Boolean
109      Dim fh As Long
On Error Resume Next
110      If Len(sFilename) = 0 Then Exit Function

         Err.Clear
111         fh = FreeFile
112         Open sFilename For Output Access Write As #fh
113         Print #fh, sContents
114         Close #fh
         SaveToFile = (Err.Number = 0)
End Function


' --------------------------------------------------------
' Calls the windows API to get the windows\SYSTEM directory
' --------------------------------------------------------
Public Property Get WindowsSystemDirectory() As String
    Dim X               As Integer
    Dim sT              As String
    Static sSystemDir   As String

    If Len(sSystemDir) Then
       WindowsSystemDirectory = sSystemDir
       Exit Property
    End If

    sT = String(145, 0)                 ' Size Buffer
    X = GetSystemDirectory(sT, 145)      ' Make API Call
    sT = Left(sT, X)                 ' Trim Buffer

    If Right(sT, 1) <> "\" Then         ' Add \ if necessary
       WindowsSystemDirectory = sT + "\"
       sSystemDir = sT + "\"
    Else
       WindowsSystemDirectory = sT
       sSystemDir = sT
    End If
End Property

