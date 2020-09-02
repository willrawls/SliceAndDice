VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmRegister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Slice and Dice Online Registration"
   ClientHeight    =   2475
   ClientLeft      =   6240
   ClientTop       =   4275
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   525
      Left            =   30
      TabIndex        =   3
      Top             =   1860
      Width           =   3855
   End
   Begin VB.TextBox txtInvoiceNumber 
      Height          =   300
      Left            =   1620
      TabIndex        =   1
      ToolTipText     =   "Enter the Invoice number given to you during step 1 here."
      Top             =   780
      Width           =   2205
   End
   Begin VB.CommandButton cmdStepTwo 
      Caption         =   "Step 2: Inform Central Server of Invoice Number"
      Enabled         =   0   'False
      Height          =   525
      Left            =   30
      TabIndex        =   2
      Top             =   1260
      Width           =   3855
   End
   Begin VB.CommandButton cmdStepOne 
      Caption         =   "Step 1: Secure Ordering / Payment Online"
      Height          =   525
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3855
   End
   Begin InetCtlsObjects.Inet inetRegister 
      Index           =   0
      Left            =   3210
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      RequestTimeout  =   100
   End
   Begin VB.Label lblInvoiceNumber 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Invoice Number from Step 1:"
      Height          =   390
      Left            =   210
      TabIndex        =   4
      Top             =   720
      Width           =   1245
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CurrentStage As Long
Public Parent As NewCommands

Public Sub SubmitTemplate()
1    On Error Resume Next
2        Dim sCategory As String
3        Dim sTemplate As String

4        With Parent.Parent
5             If Len(.CurrentTemplateNameAndCategory) > 0 Then
6                .GetCategoryAndName .CurrentTemplateNameAndCategory, sCategory, sTemplate
7                If Not .SliceAndDice(sCategory).Templates(sTemplate) Is Nothing Then
8                   With .SliceAndDice(sCategory).Templates(sTemplate)
                    
9                   End With
10               End If
11            End If
12       End With
End Sub


Public Function GetCentralUpdateInfo(Optional ByVal bFetchNewFiles As Boolean = False) As Boolean
13       Dim sResponse As String
14       Dim asaX As CAssocArray
15       Dim CurrItem As CAssocItem
16   On Error Resume Next
17       Screen.MousePointer = vbHourglass
18           sResponse = GetURL("http://www.sliceanddice.com/central.update")
19       Screen.MousePointer = vbDefault
20       If Len(sResponse) = 0 Then
21          If bUserSure("The Central Server Update Information cannot be acceessed right now." & vbCr & vbTab & "Continue with current settings ?") Then
22             GetCentralUpdateInfo = True
23          End If
24       Else
25          sResponse = sReplace(sResponse, vbCrLf, "")
26          If InStr(sResponse, "$$$$") = 0 Then
27             If bUserSure("The Central Server Update Information cannot be acceessed right now." & vbCr & vbTab & "Continue with current settings ?") Then
28                GetCentralUpdateInfo = True
29             End If
30          End If
31          Set asaX = New CAssocArray
32              asaX.ItemDelimiter = "$$$$"
33              asaX.All = sResponse
34              For Each CurrItem In asaX
35                  sadSaveLicenseKey CurrItem.Key, CurrItem.Value
36                  If bFetchNewFiles Then
                  
37                  End If
38              Next CurrItem
39          Set asaX = Nothing
40          GetCentralUpdateInfo = True
41       End If
End Function

Public Function GetFile(ByVal sURL As String, ByVal sFilename As String) As Boolean
42   On Error Resume Next
43       Dim fh As Long
44       Dim b() As Byte
45       Dim strURL As String

46       If Len(sURL) = 0 Or Len(sFilename) = 0 Then Exit Function

47       If InStr(sFilename, "\") = 0 Then sFilename = App.Path & "\" & sFilename

48       If Len(Dir(sFilename)) > 0 Then
49          Kill sFilename
50       End If

   'Load inetRegister(0)
51           inetRegister(0).RequestTimeout = 60
52           b() = inetRegister(0).OpenURL(sURL, icByteArray)
   'Unload inetRegister(0)

53       fh = FreeFile
54       Open sFilename For Binary Access Write As #fh
55            Put #fh, , b()
56       Close #fh
End Function

Public Function GetURL(ByVal sURL As String) As String
57   On Error Resume Next
   'Load inetRegister(0)
58           inetRegister(0).RequestTimeout = 60
59           GetURL = inetRegister(0).OpenURL(sURL)
   'Unload inetRegister(0)
End Function

Public Sub PostURL(ByVal sURL As String, ByVal sData)
60   On Error Resume Next
   'Load inetRegister(0)
61           inetRegister(0).RequestTimeout = 60
62           inetRegister(0).Execute sURL, "POST", sData
   'Unload inetRegister(0)
End Sub

Private Sub cmdDone_Click()
63       Hide
End Sub

Private Sub cmdStepOne_Click()
64   On Error Resume Next
65       If Len(txtInvoiceNumber) > 0 Then
66          If bUserSure("It appears you have already ordered Slice and Dice because there is an Invoice number." & vbCr & vbTab & "Would you like to continue to the online ordering system ?") Then
67             CurrentStage = 2
68             sadSaveLicenseKey "Current Stage", 2
69             Shell "start http://www.sliceanddice.com/register.html", vbNormalFocus
70          End If
71       ElseIf bUserSure("You will now be directed to the web based Slice and Dice ordering system." & vbCr & "After ordering, you'll be given an invoice number. Write it down." & vbCr & "Come back to this form and enter the Invoice number in the" & vbCr & "'Invoice Number from Step 1' text box" & vbCr & "and press the 'Step 2' button." & vbCr & vbTab & "Would you like to go to the web site now ?") Then
72          CurrentStage = 2
73          sadSaveLicenseKey "Current Stage", 2
74          Shell "start http://www.sliceanddice.com/register.html", vbNormalFocus
75       End If
End Sub

Private Sub cmdStepTwo_Click()
76       Dim InvoiceNumber           As String
77       Dim DecryptedInvoiceNumber  As String
78       Dim sEncryptedRegKey        As String
79       Dim sRegKey                 As String
80       Dim sResponse               As String
81       Dim CurrValue               As Long
82       Dim bOkaySoFar              As Boolean
83       Dim asaX                    As CAssocArray
84       Dim fh                      As Long
    
85       Dim Value08                 As String
86       Dim ProductID               As Long
87       Dim NumberOfLicenses        As Long
88       Dim LicensesRemaining       As Long
    
89   On Error Resume Next
    
90       InvoiceNumber = txtInvoiceNumber.Text
    
91       If Now - CVDate(sadGetLicenseKey("Last Updated", CDbl(Now))) > 7 Then
92          bOkaySoFar = GetCentralUpdateInfo
93       Else
94          bOkaySoFar = True
95       End If
96       If bOkaySoFar Then
97           sadSaveLicenseKey "Invoice Number", InvoiceNumber
98           sRegKey = ""
99           For CurrValue = 1 To 14
100              sRegKey = sRegKey & "Value" & Format(CurrValue, "00") & "=" & sadGetLicenseKey("Value" & Format(CurrValue, "00"), "?") & "$$$$"
101          Next CurrValue
102          sEncryptedRegKey = Scramble(Mid$(sadEncrypt(sRegKey), 5))
    
103          Screen.MousePointer = vbHourglass
104              sResponse = GetURL(sadGetLicenseKey("BackDoor", "http://205.179.61.237/SliceAndDiceFAQ/SliceAndDiceFAQ.asp?BackDoor") & "=2&InvoiceNumber=" & InvoiceNumber & "&RegKey=" & sEncryptedRegKey & "&Random=" & CLng(Rnd * 1000000))
105          Screen.MousePointer = vbDefault

106          If Len(sResponse) = 0 Then
          ' Attempt a manual registration
107              DecryptedInvoiceNumber = sadInvoiceDecrypt(InvoiceNumber)
108              If Left$(DecryptedInvoiceNumber, 7) = "VBXTRAS" Then
109                 ProductID = Val(Mid$(DecryptedInvoiceNumber, 8, 2))
110                 Value08 = Format(sadGetLicenseKey("Value03", "?"), "00000.00000")
111                 Value08 = Left$(Value08, 5) & "-999-" & Mid$(Value08, 7)
112                 NumberOfLicenses = IIf(Val(Mid$(DecryptedInvoiceNumber, 8, 2)) = 2, 2, 1) * IIf(Val(Mid$(DecryptedInvoiceNumber, 10, 3)) < 1, 1, Val(Mid$(DecryptedInvoiceNumber, 10, 3)))
113                 If NumberOfLicenses < 1 Then NumberOfLicenses = 1
114                 LicensesRemaining = NumberOfLicenses
               'Country = sCountry
115              ElseIf Left$(DecryptedInvoiceNumber, 10) = "COMPSOURCE" Then
116                 ProductID = Val(Mid$(DecryptedInvoiceNumber, 11, 2))
117                 Value08 = Format(sadGetLicenseKey("Value03", "?"), "00000.00000")
118                 Value08 = Left$(Value08, 5) & "-999-" & Mid$(Value08, 7)
119                 If Len(DecryptedInvoiceNumber) = 21 Then
                    'COMPSOURCE02001000001
                    ' 1-10
                    '11-12
                    '13-15
                    '16-21
120                    NumberOfLicenses = IIf(Val(Mid$(DecryptedInvoiceNumber, 11, 2)) = 2, 2, 1) * IIf(Val(Mid$(DecryptedInvoiceNumber, 13, 3)) < 1, 1, Val(Mid$(DecryptedInvoiceNumber, 13, 3)))
121                    If NumberOfLicenses < 1 Then NumberOfLicenses = 1
122                 ElseIf Len(DecryptedInvoiceNumber) = 20 Then
                    'COMPSOURCE0201000004
                    ' 1-10
                    '11-12
                    '13-14
                    '15-20
123                    NumberOfLicenses = IIf(Val(Mid$(DecryptedInvoiceNumber, 11, 2)) = 2, 2, 1) * IIf(Val(Mid$(DecryptedInvoiceNumber, 13, 2)) < 1, 1, Val(Mid$(DecryptedInvoiceNumber, 13, 2)))
124                    If NumberOfLicenses < 1 Then NumberOfLicenses = 1
125                 Else
126                    NumberOfLicenses = 1
127                 End If
128                 LicensesRemaining = NumberOfLicenses
               'Country = sCountry
129             Else
             ' Manual registration failed, report error.
130                 MsgBox "Unable to communicate with the central server. Please try again later." & vbCr & vbTab & "If the problem persists, please call 1-888-311-6876."
131                 Exit Sub
132             End If
133             sResponse = "Valid$$$$" & Scramble(Mid$(sadEncrypt("Value01=" & CDbl(sadGetLicenseKey("Value01", "?")) & "$$$$" & "Value03=" & CDbl(sadGetLicenseKey("Value03", "?")) & "$$$$" & "Value05=" & sadGetLicenseKey("Value05", "?") & "$$$$" & "Value08=" & Value08 & "$$$$" & "Value10=" & sadGetLicenseKey("Value10", "?") & "$$$$" & "Value12=" & sadGetLicenseKey("Value12", "?") & "$$$$" & "Value14=" & sadGetLicenseKey("Value14", "?")), 5))
134          End If

135             fh = FreeFile
136             If Len(Dir(App.Path & "\sadkey.txt")) Then
137                Kill App.Path & "\sadkey.txt"
138             End If
139             Open App.Path & "\sadkey.txt" For Output Access Write As #fh
140                  Print #fh, sResponse
141             Close #fh
           Select Case UCase$(sGetToken(sResponse, 1, "$$$$"))
                  Case "NEW RECORD CREATED"
142                         MsgBox "The central server got your Invoice number. However, payment information has not been validated at the Central Server." & vbCr & vbCr & "As soon as your payment information is processed and the Central Server is informed of your payment you'll be able to come back and do step 3 to validate your registration. If you have entered your Invoice number incorrectly, proper registration may not occur. Just correct the Invoice number and press the Step 2 button again to update the Central Server.", vbInformation
143                    Case "UNABLE TO PROCESS REQUEST"
144                         MsgBox "The central server got your Invoice number. However, some problem on the server side has prevented processing. The Central Server administrator is looking into the problem. Please try to send your request later. If the problem continues, please call 1-888-311-6876.", vbCritical
145                    Case "PAYMENT NOT RECEIVED"
146                         MsgBox "The central server already has your information on file. However, payment information has not been validated at the Central Server." & vbCr & vbCr & "As soon as your payment information is processed and the Central Server is informed of your payment you'll be able to come back and do step 2 to validate your registration. If you have entered your Invoice number incorrectly, proper registration may not occur. Just correct the Invoice number and press the Step 2 button again to update the Central Server.", vbInformation
147                    Case "INVOICE NUMBER UPDATED"
148                         MsgBox "Your Invoice Number on record with the Central Server has been updated." & vbCr & "As soon as your payment information is processed and the Central Server is informed of your payment you'll be able to come back and do step 3 to validate your registration. If you have entered your Invoice number incorrectly, proper registration may not occur. Just correct the Invoice number and press the Step 2 button again to update the Central Server.", vbInformation
149                    Case "INVALID"
150                         MsgBox "The Central Server has not been informed of your payment (yet). Please make sure your Invoice number is correct. If after 5 business days your account is not updated, call 1-888-311-6876 or write billing@sliceanddice.com for a prompt response.", vbInformation
151                    Case "ACCOUNT DISABLED"
152                         MsgBox "The Central Server indicates that this account has been disabled." & vbCr & vbCr & "Reason: " & sAfter(sResponse, 1, "$$$$") & vbCr & vbCr & "Please call 1-888-311-6876 to resolve this matter.", vbExclamation
153                    Case "VALID", "OUT OF LICENSES"
154                         If UCase$(sGetToken(sResponse, 1, "$$$$")) = "OUT OF LICENSES" Then
155                            MsgBox "The Central Server says you have run out of licenses for registration." & vbCr & "However, in the sprit of good faith you will be allowed to continue using this product as a full version and we will assume your hard drive has crashed, you are legally moving your software between computers, you have upgraded your computer or hard drive, etc.", vbCritical
156                         End If
157                         sResponse = sadDecrypt("EN* " & Scramble(sAfter(sResponse, 1, "$$$$")))
158                         If lTokenCount(sResponse, "$$$$") < 7 Then
159                            MsgBox "Invalid authorization obtained from server. Try again later. Call 1-888-311-6876 if the problem continues.", vbCritical
160                         Else
161                            Set asaX = New CAssocArray
162                                asaX.ItemDelimiter = "$$$$"
163                                asaX.All = sResponse
164                                CurrValue = 0
165                                bOkaySoFar = True
166                                Do
167                                     CurrValue = CurrValue + 1
                                   Select Case CurrValue
                                          Case 1, 3, 5, 8, 10, 12, 14
168                                                 bOkaySoFar = (Len(asaX("Value" & Format(CurrValue, "00"))) > 0)
169                                     End Select
170                                Loop While bOkaySoFar And CurrValue < 14
171                                If bOkaySoFar Then
172                                   bOkaySoFar = IsDate(CVDate(asaX("Value01"))) And IsDate(CVDate(asaX("Value03")))
173                                End If
174                                If bOkaySoFar Then
175                                   bOkaySoFar = (Format(CLng(sGetToken(asaX("Value03"), 1, ".")), "00000") = sGetToken(asaX("Value08"), 1, "-")) And (sGetToken(Format(CVDate(asaX("Value03")), ".00000"), 2, ".") = sGetToken(asaX("Value08"), 3, "-"))
176                                End If
177                                If Not bOkaySoFar Then
178                                   MsgBox "Invalid authorization obtained from server. Try again later. Call 1-888-311-6876 if the problem continues.", vbCritical
179                                Else
180                                   sadSaveLicenseKey "Value01", asaX("Value01")
                                'sadSaveLicenseKey "Value02", asaX("Value02")
181                                   sadSaveLicenseKey "Value03", asaX("Value03")
                                'sadSaveLicenseKey "Value04", asaX("Value04")
182                                   sadSaveLicenseKey "Value05", asaX("Value05")
                                'sadSaveLicenseKey "Value06", asaX("Value06")
                                'sadSaveLicenseKey "Value07", asaX("Value07")
183                                   sadSaveLicenseKey "Value08", asaX("Value08")
                                'sadSaveLicenseKey "Value09", asaX("Value09")
184                                   sadSaveLicenseKey "Value10", asaX("Value10")
                                'sadSaveLicenseKey "Value11", asaX("Value11")
185                                   sadSaveLicenseKey "Value12", asaX("Value12")
                                'sadSaveLicenseKey "Value13", asaX("Value13")
186                                   sadSaveLicenseKey "Value14", asaX("Value14")
187                                   SaveSetting "API Viewer", "Options", "AllowCopy", ""
188                                   fh = FreeFile
189                                   If Len(Dir(App.Path & "\sadkey.reg")) Then '
190                                     Kill App.Path & "\sadkey.reg"
191                                   End If
192                                   Open App.Path & "\sadkey.reg" For Output Access Write As #fh
193                                        Print #fh, "REGEDIT4"
194                                        Print #fh, ""
195                                        Print #fh, "[HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems]"
196                                        Print #fh, ""
197                                        Print #fh, "[HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License]"
198                                        Print #fh, """Value01""=""" & sadEncrypt(asaX("Value01")) & """"
199                                        Print #fh, """Value03""=""" & sadEncrypt(asaX("Value03")) & """"
200                                        Print #fh, """Value05""=""" & sadEncrypt(asaX("Value05")) & """"
201                                        Print #fh, """Value08""=""" & sadEncrypt(asaX("Value08")) & """"
202                                        Print #fh, """Value10""=""" & sadEncrypt(asaX("Value10")) & """"
203                                        Print #fh, """Value12""=""" & sadEncrypt(asaX("Value12")) & """"
204                                        Print #fh, """Value14""=""" & sadEncrypt(asaX("Value14")) & """"
205                                        Print #fh, """Invoice Number""=""" & sadEncrypt(InvoiceNumber) & """"
206                                        Print #fh, ""
207                                   Close #fh

208                                   If Val(asaX("RegistrationsLeft")) > 0 Then
209                                      MsgBox "Thank you for registering Slice and Dice. You have " & asaX("RegistrationsLeft") & " licenses left that can be installed on other machines. If you need to transfer a license, please call 1-888-311-6876.", vbCritical
210                                      Parent.Parent.ShowSplashScreen
211                                      Parent.MySadCommands.Attributes("Registered").Value = "True"
212                                      Set Parent = Nothing
213                                      Unload Me
214                                      Exit Sub
215                                   Else
216                                      MsgBox "Thank you for registering Slice and Dice!" & vbCr & "Be sure to check for product updates every few weeks as new and exciting extentions are being made.", vbInformation
217                                      Parent.Parent.ShowSplashScreen
218                                      Parent.MySadCommands.Attributes("Registered").Value = "True"
219                                      Set Parent = Nothing
220                                      Unload Me
221                                      Exit Sub
222                                   End If
223                               End If
224                            asaX.Clear
225                            Set asaX = Nothing
226                         End If
227                    Case Else
228                         MsgBox "The central server returned an invalid response and maybe down. Please try again later." & vbCr & vbTab & "If the problem persists, please call 1-888-311-6876."
229             End Select
230          End If
End Sub

Public Function sadInvoiceDecrypt(ByVal sInvoiceNumber As String) As String
231  On Error Resume Next
232      Dim strOut      As String
233      Dim bytArray()  As Byte
234      Dim CurrByte    As Long
235      Dim ValueLen    As Long
236      Dim OffsetLen   As Long
237      Dim CharLoc     As Long
238      Dim StartAt     As Long
239      Dim CurrOffset  As Long
240      Dim CheckSum    As Long
241      Dim CheckValue  As Long
242      Const Offsets   As String = "615243516784259045218002180248620684102579462315787815795168911248961534896127811596154329617581123589402160548"
243      Const Values    As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"

244      If Len(sInvoiceNumber) = 0 Then Exit Function

245      ValueLen = Len(Values)
246      OffsetLen = Len(Offsets)
247      sInvoiceNumber = Replace(Replace(Replace(UCase$(sInvoiceNumber), "O", "0"), ".", "O"), "-", "")
248      StartAt = Val(Right$(sInvoiceNumber, 1) & Left$(sInvoiceNumber, 1))
249      sInvoiceNumber = Mid$(sInvoiceNumber, 2, Len(sInvoiceNumber) - 2)
250      CheckValue = InStr(Values, Right$(sInvoiceNumber, 1))
251      sInvoiceNumber = Left$(sInvoiceNumber, Len(sInvoiceNumber) - 1)
252      CurrOffset = StartAt
253      bytArray = StrConv(sInvoiceNumber, vbFromUnicode)
254      strOut = ""

255      For CurrByte = 0 To UBound(bytArray)
256          CharLoc = InStr(Values, Chr$(bytArray(CurrByte)))
257          If CharLoc < 1 Or CharLoc > ValueLen Then Exit Function
258          CharLoc = CharLoc - Val(Mid$(Offsets, CurrOffset, 1))
259          If CharLoc < 1 Then CharLoc = ValueLen + CharLoc
260          CheckSum = (CheckSum + Asc(Mid$(Values, CharLoc, 1))) Mod ValueLen
261          strOut = strOut & Mid$(Values, CharLoc, 1)
262          CurrOffset = CurrOffset + 1
263          If CurrOffset > OffsetLen Then CurrOffset = 1
264      Next CurrByte

265      If Left$(strOut, 1) = "F" And Right$(strOut, 1) = "S" Then
266         If CheckSum < 1 Then CheckSum = 1
267         strOut = Mid$(strOut, 2, Len(strOut) - 2)
268         If CheckSum = CheckValue Then
269            sadInvoiceDecrypt = strOut
270         Else
271            sadInvoiceDecrypt = ""
272         End If
273      Else
274         sadInvoiceDecrypt = ""
275      End If
End Function


Private Sub Form_Load()
276  On Error Resume Next
277      Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
278      CurrentStage = sadGetLicenseKey("Current Stage", 1)
279      txtInvoiceNumber = sadGetLicenseKey("Invoice Number", "")
280      If Len(txtInvoiceNumber) > 10 Then
281         cmdStepTwo.Enabled = True
282      End If
283      CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License"
End Sub


Private Sub txtInvoiceNumber_Change()
284  On Error Resume Next
285      If Len(txtInvoiceNumber) > 15 Then
286         cmdStepTwo.Enabled = True
287      Else
288         cmdStepTwo.Enabled = False
289      End If
End Sub


