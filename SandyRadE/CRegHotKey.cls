VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const IDHOT_SNAPWINDOW = -1                   '/* SHIFT-PRINTSCRN  */
Private Const IDHOT_SNAPDESKTOP = -2                  '/* PRINTSCRN        */
Public Enum EHKModifiers
    MOD_ALT = &H1&
    MOD_CONTROL = &H2&
    MOD_SHIFT = &H4&
    MOD_WIN = &H8&
End Enum
Private Const WM_HOTKEY = &H312&
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long) As Long
Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
' To Report API errors:
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_RESTORE = &HF120
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOW = 5

' Implementation
Private Type tHotKeyInfo
    sName As String
    sAtomName As String
    sEventAtomName As String
    lID As Long
    eKey As KeyCodeConstants
    eModifiers As EHKModifiers
End Type
Private m_tAtoms() As tHotKeyInfo
Private m_iAtomCount As Long
Private m_hWnd As Long

Public Event HotKeyPress(ByVal sName As String, ByVal eModifiers As EHKModifiers, ByVal eKey As KeyCodeConstants)

Implements SSubTimer.ISubclass

Public Sub Attach(ByVal hwndA As Long)
1        Clear
2        If (hwndA <> 0) Then
3            m_hWnd = hwndA
4            AttachMessage Me, m_hWnd, WM_HOTKEY
5        End If
End Sub
Public Sub Clear()
6        On Error Resume Next
7        Dim i As Long
8        Dim iMax As Long

    ' Remove all hot keys and atoms:
9        iMax = m_iAtomCount
10       For i = iMax To 1 Step -1
11           UnregisterKey m_tAtoms(i).sEventAtomName
12       Next i

    ' Stop subclassing:
13       If (m_hWnd <> 0) Then
14           DetachMessage Me, m_hWnd, WM_HOTKEY
15           m_hWnd = 0
16       End If
End Sub
Public Sub RegisterKey( _
      ByVal sName As String, _
      ByVal eKey As KeyCodeConstants, _
      ByVal eModifiers As EHKModifiers _
   )
17       Dim lID As Long
18       Dim lErr As Long
19       Dim lR As Long
20       Dim sError As String
21       Dim sMsg As String
22       Dim i As Long
23       Dim sAtomName As String
24       Dim sShortAtomName As String

25       sShortAtomName = Left$(sName, 30)
    ' Check for valid user name:
    'If Len(sName) > 32 Then
    '   Err.Raise vbObjectError + 1048 + 3, App.EXEName & ".cRegHotKey", "Key Name too long (max 32 characters)."
    '   Exit Sub
    'Else
26       For i = 1 To m_iAtomCount
27           If (m_tAtoms(i).sEventAtomName = sName) Then
28               Err.Raise vbObjectError + 1048 + 4, App.EXEName & ".cRegHotKey", "The Key Name '" & sName & "' is already registered."
29               Exit Sub
30           End If
31       Next i
    'End If

    ' Modify the user supplied name to get a more random system name:
32       sAtomName = sShortAtomName & "_" & App.EXEName & "_" & GetTickCount()
33       If (Len(sAtomName) > 254) Then
34           sAtomName = Left$(sAtomName, 254)
35       End If

    ' Create a new atom:
36       lID = GlobalAddAtom(sAtomName)
37       If (lID = 0) Then
38           lErr = Err.LastDllError
39           sError = WinError(lErr)
40           sMsg = "Failed to add GlobalAtom"
41           If Len(sError) Then
42               sMsg = sMsg & " [" & sError & "]"
43           End If
44           Err.Raise vbObjectError + 1048 + 2, App.EXEName & ".cRegHotKey", sMsg
45       Else
        ' We have added the atom, now try to Register the
        ' key:
46           lR = RegisterHotKey(m_hWnd, lID, eModifiers, eKey)
47           If (lR = 0) Then
48               lErr = Err.LastDllError
            ' Remove the atom:
49               GlobalDeleteAtom lID
            ' Raise the error:
50               WinError lErr
51               sError = WinError(lErr)
52               If Right$(sError, 2) = vbNewLine Then sError = Left$(sError, Len(sError) - 2)
53               sMsg = "Failed to Register Hot Key"
54               If Len(sError) Then
55                   sMsg = sMsg & " [" & sError & "]"
56               End If
57               Err.Raise vbObjectError + 1048 + 3, App.EXEName & ".cRegHotKey", sMsg
58           Else
            ' Succeeded in adding the hot key:
59               m_iAtomCount = m_iAtomCount + 1
60               ReDim Preserve m_tAtoms(1 To m_iAtomCount) As tHotKeyInfo
61               With m_tAtoms(m_iAtomCount)
62                   .sName = sShortAtomName
63                   .sAtomName = sAtomName
64                   .sEventAtomName = sName
65                   .lID = lID
66                   .eModifiers = eModifiers
67                   .eKey = eKey
68               End With
69           End If

70       End If
End Sub
Public Sub UnregisterKey( _
      ByVal sName As String _
   )
71       Dim lIndex As Long
72       Dim i As Long
73       lIndex = AtomIndex(sName)
74       If (lIndex > 0) Then
        ' Unregister the key:
75           UnregisterHotKey m_hWnd, m_tAtoms(lIndex).lID
        ' Unregister the atom:
76           GlobalDeleteAtom m_tAtoms(lIndex).lID
        ' Remove from internal array:
77           If (m_iAtomCount > 1) Then
78               For i = lIndex To m_iAtomCount - 1
79                   LSet m_tAtoms(lIndex) = m_tAtoms(lIndex + 1)
80               Next i
81               m_iAtomCount = m_iAtomCount - 1
82               ReDim Preserve m_tAtoms(1 To m_iAtomCount) As tHotKeyInfo
83           Else
84               m_iAtomCount = 0
85               Erase m_tAtoms
86           End If
87       End If
End Sub

Private Property Get AtomIndex(ByVal sName As String) As Long
88       Dim i As Long
89       For i = 1 To m_iAtomCount
90           If (m_tAtoms(i).sEventAtomName = sName) Then
91               AtomIndex = i
92               Exit Property
93           End If
94       Next i
95       Err.Raise vbObjectError + 1048 + 1, App.EXEName & ".cRegHotKey", "No hot key registered under the name '" & sName & gsA
End Property

Private Function WinError(ByVal lLastDLLError As Long) As String
96       Dim sBuff As String
97       Dim lCount As Long

    ' Return the error message associated with LastDLLError:
98       sBuff = String$(256, 0)
99       lCount = FormatMessage( _
            FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
            0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
100      If lCount Then
101          WinError = Left$(sBuff, lCount)
102      End If

End Function

Public Sub RestoreAndActivate(ByVal hwnd As Long)
103      If (IsWindowVisible(hwnd) = 0) Then
104          ShowWindow hwnd, SW_SHOW
105      End If
106      If (IsIconic(hwnd) <> 0) Then
107          SendMessageByLong hwnd, WM_SYSCOMMAND, SC_RESTORE, 0
108      End If
109      SetForegroundWindow hwnd
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer.EMsgResponse)
' ...
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer.EMsgResponse
110      ISubclass_MsgResponse = emrPreprocess
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
111      Dim i        As Long
112      Dim lIndex   As Long

113      If (iMsg = WM_HOTKEY) Then
        ' Interpret the hotkey.  wParam is the id, the
        ' loword of lParam is the key modifier and the
        ' hiword of lParam is the key code:
        Select Case wParam
            Case IDHOT_SNAPWINDOW
114                  RaiseEvent HotKeyPress("Window Snapshot", MOD_SHIFT, vbKeySnapshot)
115              Case IDHOT_SNAPDESKTOP
116                  RaiseEvent HotKeyPress("Desktop Snapshot", 0, vbKeySnapshot)
117              Case Else
                ' Try to find id:
118                  For i = 1 To m_iAtomCount
119                      If (m_tAtoms(i).lID = wParam) Then
120                          lIndex = i
121                          Exit For
122                      End If
123                  Next i
124                  If (lIndex <> 0) Then
125                      RaiseEvent HotKeyPress(m_tAtoms(lIndex).sEventAtomName, m_tAtoms(lIndex).eModifiers, m_tAtoms(lIndex).eKey)
126                  Else
                    ' What has happened?
127                      RaiseEvent HotKeyPress("Unknown HotKey", (lParam And &HFFFF&), (lParam \ &H10000))
128                  End If
129          End Select
130      End If
End Function
