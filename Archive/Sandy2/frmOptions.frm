VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Sandle - Process C/C++ Header"
   ClientHeight    =   6345
   ClientLeft      =   6090
   ClientTop       =   1650
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
On Error GoTo 0
    If Len(sText) = 0 Or Len(sIcon) = 0 Then Exit Sub
    If Len(sKey) = 0 Then sKey = sText

    If Len(sParent) Then
       With tvwHierarchy.Nodes.Add(sParent, tvwChild, sKey, sText, sIcon, sIcon)
            .ExpandedImage = sIcon
            .Expanded = bExpanded
            .BackColor = lvwContents.BackColor
            .ForeColor = lvwContents.ForeColor
            .Tag = sTag
       End With
    Else
       With tvwHierarchy.Nodes.Add(, , sKey, sText, sIcon, sIcon)
            .ExpandedImage = sIcon
            .Expanded = bExpanded
            .BackColor = lvwContents.BackColor
            .ForeColor = lvwContents.ForeColor
            .Tag = sTag
       End With
    End If
End Sub

Private Sub ProcessFile(sFilename As String, Optional ByVal bClearFirst As Boolean = False)
    Dim sContents As String
    Dim sLine As String
    Dim LineNumber As Long
    Dim sRootName As String
    Dim bInAComment As Boolean
    Dim sEOL As String
    
    Dim sCurrNode As String
    Dim sObject As String
    Dim sScope As String
    Dim sLastMethod As String
    
    Dim sEnum As String
    Dim sStruct As String
    Dim sUnion As String
    Dim lBraceCount As Long

    Form_Load
    With Parent.Parent
        sContents = .sFileContents(sFilename)
        Screen.MousePointer = vbHourglass
        If Len(sContents) Then
           sEOL = Chr(13)
           sRootName = .sGetToken(sFilename, .lTokenCount(sFilename, "\"), "\")
           If bClearFirst Then
              lvwContents.ListItems.Clear
              tvwHierarchy.Nodes.Clear
           End If
           AddNode sRootName, "File"
           AddNode "Objects", "Object", sRootName
           AddNode "Constants", "Constant", sRootName, , , False
           AddNode "Defines", "Constant", sRootName, , "asaDefines", False
           AddNode "Enums", "Constant", sRootName
           AddNode "Structs", "Constant", sRootName
           AddNode "Union", "Constant", sRootName
           AddNode "Unknowns", "Constant", sRootName, , , False
    
         ' Make processing a little faster by getting rid of crap up front
           sContents = .sReplace(sContents, Chr(10), "")
           sContents = .sReplace(sContents, Chr(10), "")
           sContents = .sReplace(sContents, Chr(9), "     ")
    
           Do While Len(sContents)
              LineNumber = LineNumber + 1
              sLine = .sGetToken(sContents, 1, sEOL)
              sContents = .sAfter(sContents, 1, sEOL)
    
            ' Immediate line changes
              sLine = .sBefore(sLine, 2, "//")
    
            ' Filter comments
              If bInAComment Then
                 If InStr(sLine, "*/") Then
                    sLine = .sAfter(sLine, 1, "*/")
                    bInAComment = False
                 End If
              End If
              Do While InStr(sLine, "/*") And Not bInAComment
                 If InStr(sLine, "*/") Then
                    sLine = .sBefore(sLine, 2, "/*") & .sAfter(sLine, 1, "*/")
                 Else
                    bInAComment = True
                 End If
              Loop
    
            ' Filter pre-processor superfulous statements
              If InStr(LCase(sLine), "#include") And (chkProcessIncludes.Value <> 0) Then
                 sLine = ""
              ElseIf InStr(LCase(sLine), "#if") Or InStr(LCase(sLine), "#elseif") Or InStr(LCase(sLine), "#endif") Or InStr(LCase(sLine), "#pragma") Then
                 sLine = ""
              End If

            ' Determine object endings
              If Len(sObject) Then
                 lBraceCount = lBraceCount + (.lTokenCount(sLine, "{") - 1) - (.lTokenCount(sLine, "}") - 1)
                 If lBraceCount <= 0 Then sObject = "": sLine = ""
                 If InStr(LCase(sLine), "public:") Then sScope = "Public":    sLine = .sAfter(sLine, 1, "public:")
                 If InStr(LCase(sLine), "private:") Then sScope = "Private":  sLine = .sAfter(sLine, 1, "priavte:")
                 If InStr(LCase(sLine), "protected:") Then sScope = "Friend": sLine = .sAfter(sLine, 1, "protected:")
                 If StrComp(Trim(sLine), "{") = 0 Then sLine = ""
              End If

            ' Determine object beginnings
              sLine = Trim(sLine)
              If Left(LCase(sLine), 6) = "class " Then
                 If StrComp(Right(sLine, 1), ";") <> 0 Then
                    sObject = Trim(.sBefore(.sAfter(sLine), 2, "{"))
                    If InStr(sObject, ":") Then
                       sObject = .sReplace(.sReplace(sObject, "public ", ""), "  ", " ")
                       sObject = .sBefore(sObject, 2, ":") & " (Implements " & .sAfter(sObject, 1, ":") & ")"
                    End If
                    AddNode sObject, "Object", "Objects"
                    AddNode "Properties", "Property", sObject, sObject & "_Properties", , False
                    AddNode "Methods", "Method", sObject, sObject & "_Methods", , False
                    lBraceCount = .lTokenCount(sLine, "{") - 1
                    sScope = "Private"
                 End If
                 sLine = ""
              End If

            ' Finale post processing
              sLine = Trim(.sReplace(sLine, "  ", " "))
              If StrComp(Right(sLine, 1), ";") = 0 Then sLine = Left(sLine, Len(sLine) - 1)
    
            ' Determine which area it fits into and place it there
              If Len(sLine) > 0 And Not bInAComment Then
                 If Len(sObject) Then
                    If Len(sLastMethod) Then
                       tvwHierarchy.Nodes(sLastMethod).Text = tvwHierarchy.Nodes(sLastMethod).Text & " " & sLine
                       If InStr(sLine, ")") Then
                          sLastMethod = ""
                       End If
                    ElseIf InStr(sLine, "(") Then
                       If InStr(sLine, "~" & Trim(.sGetToken(sObject, 1, "(")) & "(") Then
                          AddNode sScope & " Class_Terminate", "Method", sObject & "_Methods", "Line " & LineNumber, , False
                       ElseIf InStr(sLine, Trim(.sGetToken(sObject, 1, "(")) & "(") Then
                          AddNode sScope & " Class_Initialize", "Method", sObject & "_Methods", "Line " & LineNumber, , False
                       Else
                          AddNode sScope & " " & MassReplace(sLine), "Method", sObject & "_Methods", "Line " & LineNumber, , False
                       End If
                       If InStr(sLine, ")") = 0 Then
                          sLastMethod = "Line " & LineNumber
                       End If
                    Else
                       AddNode sScope & " " & MassReplace(sLine), "Property", sObject & "_Properties", "Line " & LineNumber, , False
                    End If
                 ElseIf Left(sLine, 1) = "#" Then
                    If InStr(LCase(sLine), "define") Then
                       sLine = .sReplace(sLine, "  ", " ")
                       sLine = Trim(Mid(sLine, InStr(LCase(sLine), "define") + 6))
                       asaDefines(.sGetToken(sLine)) = .sAfter(sLine)
                       If Len(asaDefines(.sGetToken(sLine))) = 0 Then asaDefines(.sGetToken(sLine)) = .sGetToken(sLine)
                      'AddNode sLine, "Constant", "Defines", "Line " & LineNumber, , False
                    Else
                       AddNode sLine, "Constant", "Unknowns", "Line " & LineNumber, , False
                    End If
                 Else
                    AddNode sLine, "Constant", "Unknowns", "Line " & LineNumber
                 End If
               End If
           Loop
        End If
        Screen.MousePointer = vbDefault
    End With
End Sub

Private Function MassReplace(sLine) As String
    Dim CurrItem As CAssocItem
    Dim sOut As String
    sOut = sLine
    For Each CurrItem In asaDefines
        If Len(CurrItem.Key) > 0 Then
           sOut = Parent.Parent.sReplace(sOut, CurrItem.Key, CurrItem.Value)
        End If
    Next CurrItem
    MassReplace = Parent.Parent.sReplace(sOut, "!!!", "")
    
End Function

Private Sub chkProcessIncludes_Click()
    SaveSetting App.ProductName, "Last", "Process Includes", chkProcessIncludes.Value
End Sub


Private Sub cmdCancel_Click()
    Form_Unload 0
    Hide
End Sub

Private Sub cmdGenerate_Click()
    Form_Unload 0
    MsgBox "Generation would occur here"
    Hide
End Sub

Private Sub cmdOK_Click()
    Form_Unload 0
    Hide
End Sub

Private Sub cmdPickFile_Click()
    Dim sFilename As String
    sFilename = Parent.Parent.sChooseFile(, , "C Header|*.h|C++ Header|*.hpp|All Files|*.*")
    If Len(sFilename) Then ProcessFile sFilename, True
End Sub


Private Sub Form_Load()
    LoadFormPosition Me
    Form_Resize
    With asaDefines
         .Clear
         .Item("&") = "!!!"
         .Item("virtual ") = "!!!"
         .Item("char *") = "String "
         .Item("LPCTSTR ") = "String "
         .Item("BSTR ") = "String "
         .Item("void *") = "Long "
         .Item("int ") = "Long "
         .Item("short") = "Integer"
       '.Item("long ") = "Long "
         .Item("char ") = "Byte "
         .Item("") = " "
         .Item("") = " "
         .Item("") = " "
         .Item("operator==") = "CompareForEquality "
         .Item("operator!=") = "CompareForInequality "
         .Item("_exports") = "!!!"
         .Item(": public") = " - Implements "
         .Item("const ") = "!!!"
         .Item("afx_msg ") = "!!!"
         .Item(" const") = "!!!"
         .Item("(void)") = "()"
         .Item("void ") = "Sub "
         .Item(" void") = "!!!"
         .Item("(String ") = "(ByVal String "
         .Item(", String ") = ", ByVal String "
    End With
    chkProcessIncludes.Value = GetSetting(App.ProductName, "Last", "Process Includes", 0)
End Sub

Private Sub Form_Resize()
On Error Resume Next
   picLeft.Width = ScaleWidth * 0.55
   picRight.Width = ScaleWidth - picLeft.Width - 100
   tvwHierarchy.Move 30, 30, picLeft.Width - 30, picLeft.Height - 30
   lvwContents.Move 30, 30, picRight.Width - 30, picRight.Height - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPosition Me
End Sub

Private Sub tvwHierarchy_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim CurrItem As SliceAndDice.CAssocItem
    Dim CurrNode As Node
    
    If Button = 1 And Shift = 0 Then
       Set CurrNode = Nothing
       Set CurrNode = tvwHierarchy.HitTest(x, y)
       If Not CurrNode Is Nothing Then
           If lvwContents.Tag <> UCase(CurrNode.Tag) Then
              lvwContents.ListItems.Clear
              Select Case UCase(CurrNode.Tag)
                     Case "ASADEFINES"
                          For Each CurrItem In asaDefines
                              With lvwContents.ListItems.Add(, , CurrItem.Key, "Constant", "Constant")
                                   .SubItems(1) = CurrItem.Value
                              End With
                          Next CurrItem
                          lvwContents.Tag = UCase(CurrNode.Tag)
              End Select
           End If
       End If
       Set CurrNode = Nothing
    End If
End Sub

