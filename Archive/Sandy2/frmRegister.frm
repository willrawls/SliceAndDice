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
On Error Resume Next
    Dim sCategory As String
    Dim sTemplate As String

    With Parent.Parent
         If Len(.CurrentTemplateNameAndCategory) > 0 Then
            .GetCategoryAndName .CurrentTemplateNameAndCategory, sCategory, sTemplate
            If Not .SliceAndDice(sCategory).Templates(sTemplate) Is Nothing Then
               With .SliceAndDice(sCategory).Templates(sTemplate)
                    
               End With
            End If
         End If
    End With
End Sub


Public Function GetCentralUpdateInfo(Optional ByVal bFetchNewFiles As Boolean = False) As Boolean
    Dim sResponse As String
    Dim asaX As CAssocArray
    Dim CurrItem As CAssocItem
On Error Resume Next
    Screen.MousePointer = vbHourglass
        sResponse = GetURL("http://www.SandySupport.com/central.update")
    Screen.MousePointer = vbDefault
    If Len(sResponse) = 0 Then
       If bUserSure("The Central Server Update Information cannot be acceessed right now." & Chr(13) & Chr(9) & "Continue with current settings ?") Then
          GetCentralUpdateInfo = True
       End If
    Else
       sResponse = Replace(sResponse, Chr(13) & Chr(10), "")
       If InStr(sResponse, "$$$$") = 0 Then
          If bUserSure("The Central Server Update Information cannot be acceessed right now." & Chr(13) & Chr(9) & "Continue with current settings ?") Then
             GetCentralUpdateInfo = True
          End If
       End If
       Set asaX = New CAssocArray
           asaX.ItemDelimiter = "$$$$"
           asaX.All = sResponse
           For Each CurrItem In asaX
               sadSaveLicenseKey CurrItem.Key, CurrItem.Value
               If bFetchNewFiles Then
                  
               End If
           Next CurrItem
       Set asaX = Nothing
       GetCentralUpdateInfo = True
    End If
End Function

Public Function GetFile(ByVal sURL As String, ByVal sFilename As String) As Boolean
On Error Resume Next
    Dim fh As Long
    Dim b() As Byte
    Dim strURL As String

    If Len(sURL) = 0 Or Len(sFilename) = 0 Then Exit Function

    If InStr(sFilename, "\") = 0 Then sFilename = App.Path & "\" & sFilename

    If Len(Dir(sFilename)) > 0 Then
       Kill sFilename
    End If

   'Load inetRegister(0)
        inetRegister(0).RequestTimeout = 120
        b() = inetRegister(0).OpenURL(sURL, icByteArray)
   'Unload inetRegister(0)

    fh = FreeFile
    Open sFilename For Binary Access Write As #fh
         Put #fh, , b()
    Close #fh
End Function

Public Function GetURL(ByVal sURL As String) As String
On Error Resume Next
   'Load inetRegister(0)
        inetRegister(0).RequestTimeout = 120
        GetURL = inetRegister(0).OpenURL(sURL)
   'Unload inetRegister(0)
End Function

Public Sub PostURL(ByVal sURL As String, ByVal sData)
On Error Resume Next
   'Load inetRegister(0)
        inetRegister(0).RequestTimeout = 120
        inetRegister(0).Execute sURL, "POST", sData
   'Unload inetRegister(0)
End Sub

Private Sub cmdDone_Click()
    Hide
End Sub

Private Sub cmdStepOne_Click()
On Error Resume Next
    If Len(txtInvoiceNumber) > 0 Then
       If bUserSure("It appears you have already ordered Slice and Dice because there is an Invoice number." & Chr(13) & Chr(9) & "Would you like to continue to the online ordering system ?") Then
          CurrentStage = 2
          sadSaveLicenseKey "Current Stage", 2
          Shell "start http://www.SandySupport.com/register.html", vbNormalFocus
       End If
    ElseIf bUserSure("You will now be directed to the web based Slice and Dice ordering system." & Chr(13) & "After ordering, you'll be given an invoice number. Write it down." & Chr(13) & "Come back to this form and enter the Invoice number in the" & Chr(13) & "'Invoice Number from Step 1' text box" & Chr(13) & "and press the 'Step 2' button." & Chr(13) & Chr(9) & "Would you like to go to the web site now ?") Then
       CurrentStage = 2
       sadSaveLicenseKey "Current Stage", 2
       Shell "start http://www.SandySupport.com/register.html", vbNormalFocus
    End If
End Sub

Private Sub cmdStepTwo_Click()
    Dim sRegKey As String
    Dim sResponse As String
    Dim CurrValue As Long
    Dim bOkaySoFar As Boolean
    Dim asaX As CAssocArray
    Dim fh As Long
    
On Error Resume Next
    
    If Now - CVDate(sadGetLicenseKey("Last Updated", CDbl(Now))) > 7 Then
       bOkaySoFar = GetCentralUpdateInfo
    Else
       bOkaySoFar = True
    End If
    If bOkaySoFar Then
        sadSaveLicenseKey "Invoice Number", txtInvoiceNumber
        sRegKey = ""
        For CurrValue = 1 To 14
            sRegKey = sRegKey & "Value" & Format(CurrValue, "00") & "=" & sadGetLicenseKey("Value" & Format(CurrValue, "00"), "?") & "$$$$"
        Next CurrValue
        sRegKey = Scramble(Mid(sadEncrypt(sRegKey), 5))
    
        Screen.MousePointer = vbHourglass
            sResponse = GetURL(sadGetLicenseKey("BackDoor", "http://209.196.104.22/SliceAndDiceFAQ/SliceAndDiceFAQ.asp?BackDoor") & "=2&InvoiceNumber=" & txtInvoiceNumber & "&RegKey=" & sRegKey & "&Random=" & CLng(Rnd * 1000000))
        Screen.MousePointer = vbDefault

        If Len(sResponse) = 0 Then
           MsgBox "Unable to communicate with the central server. Please try again later." & Chr(13) & Chr(9) & "If the problem persists, please call 1-888-311-6876."
        Else
           fh = FreeFile
           If Len(Dir(App.Path & "\s2kkey.txt")) Then
              Kill App.Path & "\s2kkey.txt"
           End If
           Open App.Path & "\s2kkey.txt" For Output Access Write As #fh
                Print #fh, sResponse
           Close #fh
           Select Case UCase(sGetToken(sResponse, 1, "$$$$"))
                  Case "NEW RECORD CREATED"
                       MsgBox "The central server got your Invoice number. However, payment information has not been validated at the Central Server." & Chr(13) & Chr(13) & "As soon as your payment information is processed and the Central Server is informed of your payment you'll be able to come back and do step 3 to validate your registration. If you have entered your Invoice number incorrectly, proper registration may not occur. Just correct the Invoice number and press the Step 2 button again to update the Central Server.", vbInformation
                  Case "UNABLE TO PROCESS REQUEST"
                       MsgBox "The central server got your Invoice number. However, some problem on the server side has prevented processing. The Central Server administrator is looking into the problem. Please try to send your request later. If the problem continues, please call 1-888-311-6876.", vbCritical
                  Case "PAYMENT NOT RECEIVED"
                       MsgBox "The central server already has your information on file. However, payment information has not been validated at the Central Server." & Chr(13) & Chr(13) & "As soon as your payment information is processed and the Central Server is informed of your payment you'll be able to come back and do step 2 to validate your registration. If you have entered your Invoice number incorrectly, proper registration may not occur. Just correct the Invoice number and press the Step 2 button again to update the Central Server.", vbInformation
                  Case "INVOICE NUMBER UPDATED"
                       MsgBox "Your Invoice Number on record with the Central Server has been updated." & Chr(13) & "As soon as your payment information is processed and the Central Server is informed of your payment you'll be able to come back and do step 3 to validate your registration. If you have entered your Invoice number incorrectly, proper registration may not occur. Just correct the Invoice number and press the Step 2 button again to update the Central Server.", vbInformation
                  Case "INVALID"
                       MsgBox "The Central Server has not been informed of your payment (yet). Please make sure your Invoice number is correct. If after 5 business days your account is not updated, call 1-888-311-6876 or write billing@SandySupport.com for a prompt response.", vbInformation
                  Case "ACCOUNT DISABLED"
                       MsgBox "The Central Server indicates that this account has been disabled." & Chr(13) & Chr(13) & "Reason: " & sAfter(sResponse, 1, "$$$$") & Chr(13) & Chr(13) & "Please call 1-888-311-6876 to resolve this matter.", vbExclamation
                  Case "VALID", "OUT OF LICENSES"
                       If UCase(sGetToken(sResponse, 1, "$$$$")) = "OUT OF LICENSES" Then
                          MsgBox "The Central Server says you have run out of licenses for registration." & Chr(13) & "However, in the sprit of good faith you will be allowed to continue using this product as a full version and we will assume your hard drive has crashed, you are legally moving your software between computers, you have upgraded your computer or hard drive, etc.", vbCritical
                       End If
                       sResponse = sadDecrypt("EN* " & Scramble(sAfter(sResponse, 1, "$$$$")))
                       If lTokenCount(sResponse, "$$$$") < 7 Then
                          MsgBox "Invalid authorization obtained from server. Try again later. Call 1-888-311-6876 if the problem continues.", vbCritical
                       Else
                          Set asaX = New CAssocArray
                              asaX.ItemDelimiter = "$$$$"
                              asaX.All = sResponse
                              CurrValue = 0
                              bOkaySoFar = True
                              Do
                                   CurrValue = CurrValue + 1
                                   Select Case CurrValue
                                          Case 1, 3, 5, 8, 10, 12, 14
                                               bOkaySoFar = (Len(asaX("Value" & Format(CurrValue, "00"))) > 0)
                                   End Select
                              Loop While bOkaySoFar And CurrValue < 14
                              If bOkaySoFar Then
                                 bOkaySoFar = IsDate(CVDate(asaX("Value01"))) And IsDate(CVDate(asaX("Value03")))
                              End If
                              If bOkaySoFar Then
                                 bOkaySoFar = (Format(CLng(sGetToken(asaX("Value03"), 1, ".")), "00000") = sGetToken(asaX("Value08"), 1, "-")) And (sGetToken(Format(CVDate(asaX("Value03")), ".00000"), 2, ".") = sGetToken(asaX("Value08"), 3, "-"))
                              End If
                              If Not bOkaySoFar Then
                                 MsgBox "Invalid authorization obtained from server. Try again later. Call 1-888-311-6876 if the problem continues.", vbCritical
                              Else
                                 sadSaveLicenseKey "Value01", asaX("Value01")
                                'sadSaveLicenseKey "Value02", asaX("Value02")
                                 sadSaveLicenseKey "Value03", asaX("Value03")
                                'sadSaveLicenseKey "Value04", asaX("Value04")
                                 sadSaveLicenseKey "Value05", asaX("Value05")
                                'sadSaveLicenseKey "Value06", asaX("Value06")
                                'sadSaveLicenseKey "Value07", asaX("Value07")
                                 sadSaveLicenseKey "Value08", asaX("Value08")
                                'sadSaveLicenseKey "Value09", asaX("Value09")
                                 sadSaveLicenseKey "Value10", asaX("Value10")
                                'sadSaveLicenseKey "Value11", asaX("Value11")
                                 sadSaveLicenseKey "Value12", asaX("Value12")
                                'sadSaveLicenseKey "Value13", asaX("Value13")
                                 sadSaveLicenseKey "Value14", asaX("Value14")
                                 SaveSetting "API Viewer", "Options", "AllowCopy", ""
                                 fh = FreeFile
                                 If Len(Dir(App.Path & "\sadkey.reg")) Then '
                                   Kill App.Path & "\sadkey.reg"
                                 End If
                                 Open App.Path & "\sadkey.reg" For Output Access Write As #fh
                                      Print #fh, "REGEDIT4"
                                      Print #fh, ""
                                      Print #fh, "[HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems]"
                                      Print #fh, ""
                                      Print #fh, "[HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License]"
                                      Print #fh, """Value01""=""" & sadEncrypt(asaX("Value01")) & """"
                                      Print #fh, """Value03""=""" & sadEncrypt(asaX("Value03")) & """"
                                      Print #fh, """Value05""=""" & sadEncrypt(asaX("Value05")) & """"
                                      Print #fh, """Value08""=""" & sadEncrypt(asaX("Value08")) & """"
                                      Print #fh, """Value10""=""" & sadEncrypt(asaX("Value10")) & """"
                                      Print #fh, """Value12""=""" & sadEncrypt(asaX("Value12")) & """"
                                      Print #fh, """Value14""=""" & sadEncrypt(asaX("Value14")) & """"
                                      Print #fh, """Invoice Number""=""" & sadEncrypt(txtInvoiceNumber) & """"
                                      Print #fh, ""
                                 Close #fh

                                 If Val(asaX("RegistrationsLeft")) > 0 Then
                                    MsgBox "Thank you for registering Slice and Dice. You have " & asaX("RegistrationsLeft") & " licenses left that can be installed on other machines. If you need to transfer a license, please call 1-888-311-6876.", vbCritical
                                    Parent.Parent.ShowSplashScreen
                                    Parent.MySadCommands.Attributes("Registered").Value = "True"
                                    Set Parent = Nothing
                                    Unload Me
                                    Exit Sub
                                 Else
                                    MsgBox "Thank you for registering Slice and Dice!" & Chr(13) & "Be sure to check for product updates every few weeks as new and exciting extentions are being made.", vbInformation
                                    Parent.Parent.ShowSplashScreen
                                    Parent.MySadCommands.Attributes("Registered").Value = "True"
                                    Set Parent = Nothing
                                    Unload Me
                                    Exit Sub
                                 End If
                             End If
                          asaX.Clear
                          Set asaX = Nothing
                       End If
                  Case Else
                       MsgBox "The central server returned an invalid response and maybe down. Please try again later." & Chr(13) & Chr(9) & "If the problem persists, please call 1-888-311-6876."
           End Select
        End If
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    CurrentStage = sadGetLicenseKey("Current Stage", 1)
    txtInvoiceNumber = sadGetLicenseKey("Invoice Number", "")
    If Len(txtInvoiceNumber) > 10 Then
       cmdStepTwo.Enabled = True
    End If
    CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Zion Systems\License"
End Sub


Private Sub txtInvoiceNumber_Change()
On Error Resume Next
    If Len(txtInvoiceNumber) > 15 Then
       cmdStepTwo.Enabled = True
    Else
       cmdStepTwo.Enabled = False
    End If
End Sub


