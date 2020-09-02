VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Sandle - Process C/C++ Header"
   ClientHeight    =   6345
   ClientLeft      =   420
   ClientTop       =   570
   ClientWidth     =   8610
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8610
      TabIndex        =   6
      Top             =   5850
      Width           =   8610
      Begin VB.CheckBox chkProcessIncludes 
         Caption         =   "Process #include <x.h> files"
         Height          =   525
         Left            =   4980
         TabIndex        =   15
         Top             =   0
         Width           =   2565
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   135
         TabIndex        =   10
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1350
         TabIndex        =   9
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         Height          =   375
         Left            =   3780
         TabIndex        =   8
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton cmdPickFile 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   2565
         TabIndex        =   7
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.PictureBox picRight 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5850
      Left            =   2790
      ScaleHeight     =   5850
      ScaleWidth      =   5820
      TabIndex        =   13
      Top             =   0
      Width           =   5820
      Begin MSComctlLib.ListView lvwContents 
         Height          =   5640
         Left            =   0
         TabIndex        =   14
         Top             =   15
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9948
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "imlIcons"
         SmallIcons      =   "imlIcons"
         ColHdrIcons     =   "imlIcons"
         ForeColor       =   -2147483640
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Name"
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Value"
            Text            =   "Value"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5850
      Left            =   0
      ScaleHeight     =   5850
      ScaleWidth      =   2640
      TabIndex        =   11
      Top             =   0
      Width           =   2640
      Begin MSComctlLib.TreeView tvwHierarchy 
         Height          =   5310
         Left            =   15
         TabIndex        =   12
         Top             =   15
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   9366
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   265
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imlIcons"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imlIcons 
         Left            =   45
         Top             =   5190
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483648
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":000C
               Key             =   "File"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":045E
               Key             =   "Check"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":08B0
               Key             =   "Property"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":0D02
               Key             =   "Property - Number 1"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":101C
               Key             =   "Property - Number 2"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":1336
               Key             =   "Property - Number 3"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":1650
               Key             =   "Diamond"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":1AA2
               Key             =   "Clock"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":1EF4
               Key             =   "Plus"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2346
               Key             =   "Minus"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2798
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2BEA
               Key             =   "Object"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":303C
               Key             =   "Method"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":348E
               Key             =   "Variable"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":38E0
               Key             =   "Constant"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   3735
         Index           =   3
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   6588
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "imlIcons"
         SmallIcons      =   "imlIcons"
         ColHdrIcons     =   "imlIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   3735
         Index           =   2
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   6588
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "imlIcons"
         SmallIcons      =   "imlIcons"
         ColHdrIcons     =   "imlIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   3735
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   6588
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "imlIcons"
         SmallIcons      =   "imlIcons"
         ColHdrIcons     =   "imlIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Parent As NewCommands

Public asaDefines As New SliceAndDice.CAssocArray
Public asaTemp As New SliceAndDice.CAssocArray

Private Sub AddNode(ByVal sText As String, ByVal sIcon As String, Optional ByVal sParent As String, Optional ByVal sKey As String, Optional ByVal sTag As String, Optional ByVal bExpanded As Boolean = True)
1    On Error GoTo 0
2        If Len(sText) = 0 Or Len(sIcon) = 0 Then Exit Sub
3        If Len(sKey) = 0 Then sKey = sText

4        If Len(sParent) Then
5           With tvwHierarchy.Nodes.Add(sParent, tvwChild, sKey, sText, sIcon, sIcon)
6                .ExpandedImage = sIcon
7                .Expanded = bExpanded
8                .BackColor = lvwContents.BackColor
9                .ForeColor = lvwContents.ForeColor
10               .Tag = sTag
11          End With
12       Else
13          With tvwHierarchy.Nodes.Add(, , sKey, sText, sIcon, sIcon)
14               .ExpandedImage = sIcon
15               .Expanded = bExpanded
16               .BackColor = lvwContents.BackColor
17               .ForeColor = lvwContents.ForeColor
18               .Tag = sTag
19          End With
20       End If
End Sub

Private Sub ProcessFile(sFilename As String, Optional ByVal bClearFirst As Boolean = False)
21       Dim sContents As String
22       Dim sLine As String
23       Dim LineNumber As Long
24       Dim sRootName As String
25       Dim bInAComment As Boolean
26       Dim sEOL As String
    
27       Dim sCurrNode As String
28       Dim sObject As String
29       Dim sScope As String
30       Dim sLastMethod As String
    
31       Dim sEnum As String
32       Dim sStruct As String
33       Dim sUnion As String
34       Dim lBraceCount As Long

35       Form_Load
36       With Parent.Parent
37           sContents = .sFileContents(sFilename)
38           Screen.MousePointer = vbHourglass
39           If Len(sContents) Then
40              sEOL = Chr(13)
41              sRootName = .sGetToken(sFilename, .lTokenCount(sFilename, "\"), "\")
42              If bClearFirst Then
43                 lvwContents.ListItems.Clear
44                 tvwHierarchy.Nodes.Clear
45              End If
46              AddNode sRootName, "File"
47              AddNode "Objects", "Object", sRootName
48              AddNode "Constants", "Constant", sRootName, , , False
49              AddNode "Defines", "Constant", sRootName, , "asaDefines", False
50              AddNode "Enums", "Constant", sRootName
51              AddNode "Structs", "Constant", sRootName
52              AddNode "Union", "Constant", sRootName
53              AddNode "Unknowns", "Constant", sRootName, , , False
    
         ' Make processing a little faster by getting rid of crap up front
54              sContents = Replace(sContents, Chr(10), "")
55              sContents = Replace(sContents, Chr(10), "")
56              sContents = Replace(sContents, Chr(9), "     ")
    
57              Do While Len(sContents)
58                 LineNumber = LineNumber + 1
59                 sLine = .sGetToken(sContents, 1, sEOL)
60                 sContents = .sAfter(sContents, 1, sEOL)
    
            ' Immediate line changes
61                 sLine = .sBefore(sLine, 2, "//")
    
            ' Filter comments
62                 If bInAComment Then
63                    If InStr(sLine, "*/") Then
64                       sLine = .sAfter(sLine, 1, "*/")
65                       bInAComment = False
66                    End If
67                 End If
68                 Do While InStr(sLine, "/*") And Not bInAComment
69                    If InStr(sLine, "*/") Then
70                       sLine = .sBefore(sLine, 2, "/*") & .sAfter(sLine, 1, "*/")
71                    Else
72                       bInAComment = True
73                    End If
74                 Loop
    
            ' Filter pre-processor superfulous statements
75                 If InStr(LCase(sLine), "#include") And (chkProcessIncludes.Value <> 0) Then
76                    sLine = ""
77                 ElseIf InStr(LCase(sLine), "#if") Or InStr(LCase(sLine), "#elseif") Or InStr(LCase(sLine), "#endif") Or InStr(LCase(sLine), "#pragma") Then
78                    sLine = ""
79                 End If

            ' Determine object endings
80                 If Len(sObject) Then
81                    lBraceCount = lBraceCount + (.lTokenCount(sLine, "{") - 1) - (.lTokenCount(sLine, "}") - 1)
82                    If lBraceCount <= 0 Then sObject = "": sLine = ""
83                    If InStr(LCase(sLine), "public:") Then sScope = "Public":    sLine = .sAfter(sLine, 1, "public:")
84                    If InStr(LCase(sLine), "private:") Then sScope = "Private":  sLine = .sAfter(sLine, 1, "priavte:")
85                    If InStr(LCase(sLine), "protected:") Then sScope = "Friend": sLine = .sAfter(sLine, 1, "protected:")
86                    If StrComp(Trim(sLine), "{") = 0 Then sLine = ""
87                 End If

            ' Determine object beginnings
88                 sLine = Trim(sLine)
89                 If Left(LCase(sLine), 6) = "class " Then
90                    If StrComp(Right(sLine, 1), ";") <> 0 Then
91                       sObject = Trim(.sBefore(.sAfter(sLine), 2, "{"))
92                       If InStr(sObject, ":") Then
93                          sObject = Replace(Replace(sObject, "public ", ""), "  ", " ")
94                          sObject = .sBefore(sObject, 2, ":") & " (Implements " & .sAfter(sObject, 1, ":") & ")"
95                       End If
96                       AddNode sObject, "Object", "Objects"
97                       AddNode "Properties", "Property", sObject, sObject & "_Properties", , False
98                       AddNode "Methods", "Method", sObject, sObject & "_Methods", , False
99                       lBraceCount = .lTokenCount(sLine, "{") - 1
100                      sScope = "Private"
101                   End If
102                   sLine = ""
103                End If

            ' Finale post processing
104                sLine = Trim(Replace(sLine, "  ", " "))
105                If StrComp(Right(sLine, 1), ";") = 0 Then sLine = Left(sLine, Len(sLine) - 1)
    
            ' Determine which area it fits into and place it there
106                If Len(sLine) > 0 And Not bInAComment Then
107                   If Len(sObject) Then
108                      If Len(sLastMethod) Then
109                         tvwHierarchy.Nodes(sLastMethod).Text = tvwHierarchy.Nodes(sLastMethod).Text & " " & sLine
110                         If InStr(sLine, ")") Then
111                            sLastMethod = ""
112                         End If
113                      ElseIf InStr(sLine, "(") Then
114                         If InStr(sLine, "~" & Trim(.sGetToken(sObject, 1, "(")) & "(") Then
115                            AddNode sScope & " Class_Terminate", "Method", sObject & "_Methods", "Line " & LineNumber, , False
116                         ElseIf InStr(sLine, Trim(.sGetToken(sObject, 1, "(")) & "(") Then
117                            AddNode sScope & " Class_Initialize", "Method", sObject & "_Methods", "Line " & LineNumber, , False
118                         Else
119                            AddNode sScope & " " & MassReplace(sLine), "Method", sObject & "_Methods", "Line " & LineNumber, , False
120                         End If
121                         If InStr(sLine, ")") = 0 Then
122                            sLastMethod = "Line " & LineNumber
123                         End If
124                      Else
125                         AddNode sScope & " " & MassReplace(sLine), "Property", sObject & "_Properties", "Line " & LineNumber, , False
126                      End If
127                   ElseIf Left(sLine, 1) = "#" Then
128                      If InStr(LCase(sLine), "define") Then
129                         sLine = Replace(sLine, "  ", " ")
130                         sLine = Trim(Mid(sLine, InStr(LCase(sLine), "define") + 6))
131                         asaDefines(.sGetToken(sLine)) = .sAfter(sLine)
132                         If Len(asaDefines(.sGetToken(sLine))) = 0 Then asaDefines(.sGetToken(sLine)) = .sGetToken(sLine)
                      'AddNode sLine, "Constant", "Defines", "Line " & LineNumber, , False
133                      Else
134                         AddNode sLine, "Constant", "Unknowns", "Line " & LineNumber, , False
135                      End If
136                   Else
137                      AddNode sLine, "Constant", "Unknowns", "Line " & LineNumber
138                   End If
139                 End If
140             Loop
141          End If
142          Screen.MousePointer = vbDefault
143      End With
End Sub

Private Function MassReplace(sLine) As String
144      Dim CurrItem    As CAssocItem
145      Dim sOut        As String

146      sOut = sLine
147      For Each CurrItem In asaDefines
148          If Len(CurrItem.Key) > 0 Then
149             sOut = Replace(sOut, CurrItem.Key, CurrItem.Value)
150          End If
151      Next CurrItem
152      MassReplace = Replace(sOut, "!!!", "")

End Function

Private Sub chkProcessIncludes_Click()
153      SaveSetting App.ProductName, "Last", "Process Includes", chkProcessIncludes.Value
End Sub


Private Sub cmdCancel_Click()
154      Form_Unload 0
155      Hide
End Sub

Private Sub cmdGenerate_Click()
156      Form_Unload 0
157      MsgBox "Generation would occur here"
158      Hide
End Sub

Private Sub cmdOK_Click()
159      Form_Unload 0
160      Hide
End Sub

Private Sub cmdPickFile_Click()
161      Dim sFilename As String
162      sFilename = Parent.Parent.sChooseFile(, , "C Header|*.h|C++ Header|*.hpp|All Files|*.*")
163      If Len(sFilename) Then ProcessFile sFilename, True
End Sub


Private Sub Form_Load()
164      LoadFormPosition Me
165      Form_Resize
166      With asaDefines
167           .Clear
168           .Item("&") = "!!!"
169           .Item("virtual ") = "!!!"
170           .Item("char *") = "String "
171           .Item("LPCTSTR ") = "String "
172           .Item("BSTR ") = "String "
173           .Item("void *") = "Long "
174           .Item("int ") = "Long "
175           .Item("short") = "Integer"
       '.Item("long ") = "Long "
176           .Item("char ") = "Byte "
177           .Item("") = " "
178           .Item("") = " "
179           .Item("") = " "
180           .Item("operator==") = "CompareForEquality "
181           .Item("operator!=") = "CompareForInequality "
182           .Item("_exports") = "!!!"
183           .Item(": public") = " - Implements "
184           .Item("const ") = "!!!"
185           .Item("afx_msg ") = "!!!"
186           .Item(" const") = "!!!"
187           .Item("(void)") = "()"
188           .Item("void ") = "Sub "
189           .Item(" void") = "!!!"
190           .Item("(String ") = "(ByVal String "
191           .Item(", String ") = ", ByVal String "
192      End With
193      chkProcessIncludes.Value = GetSetting(App.ProductName, "Last", "Process Includes", 0)
End Sub

Private Sub Form_Resize()
194  On Error Resume Next
195     picLeft.Width = ScaleWidth * 0.55
196     picRight.Width = ScaleWidth - picLeft.Width - 100
197     tvwHierarchy.Move 30, 30, picLeft.Width - 30, picLeft.Height - 30
198     lvwContents.Move 30, 30, picRight.Width - 30, picRight.Height - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
199      SaveFormPosition Me
End Sub

Private Sub tvwHierarchy_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
200      Dim CurrItem As SliceAndDice.CAssocItem
201      Dim CurrNode As Node
    
202      If Button = 1 And Shift = 0 Then
203         Set CurrNode = Nothing
204         Set CurrNode = tvwHierarchy.HitTest(x, y)
205         If Not CurrNode Is Nothing Then
206             If lvwContents.Tag <> UCase(CurrNode.Tag) Then
207                lvwContents.ListItems.Clear
              Select Case UCase(CurrNode.Tag)
                     Case "ASADEFINES"
208                            For Each CurrItem In asaDefines
209                                With lvwContents.ListItems.Add(, , CurrItem.Key, "Constant", "Constant")
210                                     .SubItems(1) = CurrItem.Value
211                                End With
212                            Next CurrItem
213                            lvwContents.Tag = UCase(CurrNode.Tag)
214                End Select
215             End If
216         End If
217         Set CurrNode = Nothing
218      End If
End Sub

