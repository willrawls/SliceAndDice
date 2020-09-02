VERSION 5.00
Begin {17016CEE-E118-11D0-94B8-00A0C91110ED} SliceAndDiceFAQ 
   ClientHeight    =   7350
   ClientLeft      =   2130
   ClientTop       =   2520
   ClientWidth     =   15075
   _ExtentX        =   26591
   _ExtentY        =   12965
   MajorVersion    =   0
   MinorVersion    =   8
   StateManagementType=   1
   ASPFileName     =   ""
   DIID_WebClass   =   "{12CBA1F6-9056-11D1-8544-00A024A55AB0}"
   DIID_WebClassEvents=   "{12CBA1F5-9056-11D1-8544-00A024A55AB0}"
   TypeInfoCookie  =   285
   BeginProperty WebItems {193556CD-4486-11D1-9C70-00C04FB987DF} 
      WebItemCount    =   3
      BeginProperty WebItem1 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "Category"
         DISPID          =   1280
         Template        =   "Category1.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{3FEC8E94-DE6B-11D2-BA4F-0080C8C222EC}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "D:\Data\Programming\Projects - Firm Solutions\HTMLSandy\Category.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
      BeginProperty WebItem2 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "Categorys"
         DISPID          =   1281
         Template        =   "Categorys1.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{3FEC8DD2-DE6B-11D2-BA4F-0080C8C222EC}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "D:\Data\Programming\Projects - Firm Solutions\HTMLSandy\Categorys.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
      BeginProperty WebItem3 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "PatchList"
         DISPID          =   1282
         Template        =   "PatchList1.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{51D56CD5-DF68-11D2-BA4F-0080C8C222EC}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "D:\Data\Programming\Projects - Firm Solutions\HTMLSandy\PatchList.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
   EndProperty
   NameInURL       =   "SliceAndDiceFAQ"
End
Attribute VB_Name = "SliceAndDiceFAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Public Function sCodeAreaToHTML(sCodeArea As String) As String
    sCodeAreaToHTML = "<font color=""#000000"" size=""2"" face = ""System"">" & sReplace(sReplace(sReplace(sCodeArea, Chr(13) & Chr(10), "<BR>"), Chr(9), "     "), " ", "&nbsp;") & "</font>"
End Function

Private Sub BackdoorSandyRegistration_ProcessTag(ByVal TagName As String, TagContents As String, SendTags As Boolean)
    Select Case UCase(TagName)
           Case "WC@REQUEST"
                TagContents = Request(TagContents)
    End Select
End Sub

Public Function sCountryDateOut(ByVal dtIn As Date, sCountry As String) As String
    Select Case UCase(sCountry)
           Case "GERMANY"
           Case Else ' "USA"
                sCountryDateOut = dtIn
    End Select
End Function

Private Sub Category_ProcessTag(ByVal TagName As String, TagContents As String, SendTags As Boolean)
    Dim CurrTemplate As Object 'SliceAndDice.CTemplate
    Dim sT As String
    Dim sCategory As String
On Error Resume Next
    Select Case UCase(TagContents)
           Case "LIST CATEGORY"
                sCategory = Request("CategoryName")
                For Each CurrTemplate In Sandy.Categorys(sCategory).Templates
                    sT = sT & "<TR><TD><A HREF=""SliceAndDiceFAQ.asp?CategoryName=" & sCategory & "&TemplateName=" & CurrTemplate.ShortTemplateName & """>" & CurrTemplate.ShortTemplateName & "</A></TD></TR>" & Chr(13) & Chr(10)
                Next CurrTemplate
                TagContents = sT
    End Select
End Sub

Private Sub Categorys_ProcessTag(ByVal TagName As String, TagContents As String, SendTags As Boolean)
    Dim CurrCategory As Object 'SliceAndDice.CCategory
    Dim sT As String
    
    Select Case UCase(TagContents)
           Case "LIST CATEGORYS"
                For Each CurrCategory In Sandy.Categorys
                    sT = sT & "<TR><TD><A HREF=""SliceAndDiceFAQ.asp?CategoryName=" & CurrCategory.Key & """>" & CurrCategory.Key & "</A></TD><TD>" & CurrCategory.Templates.Count & "</TD></TR>" & Chr(13) & Chr(10)
                Next CurrCategory
                TagContents = sT
    End Select
End Sub

Private Sub PatchList_ProcessTag(ByVal TagName As String, TagContents As String, SendTags As Boolean)
    Dim CurrCategory As Object 'SliceAndDice.CCategory
    Dim CurrTemplate As Object 'SliceAndDice.CTemplate
    Dim asaDates As Object 'SliceAndDice.CAssocArray
    Dim sT As String
    Dim sCategory As String

    Set asaDates = CreateObject("SliceAndDice.CAssocArray")

On Error Resume Next
    Select Case UCase(TagContents)
           Case "LIST PATCHDATES"
                asaDates.Clear
                asaDates.AddInOrder = True
                For Each CurrCategory In Sandy.Categorys
                    For Each CurrTemplate In CurrCategory.Templates
                        asaDates("" & Format(CLng(CurrTemplate.DateModified), "00000")) = "<A HREF=""SliceAndDiceFAQ.asp?GeneratePatchFile=" & Format(CurrTemplate.DateModified, "Mmmm D, YYYY HH:\0\0:\0\0 AM/PM") & """>" & Format(CurrTemplate.DateModified, "Mmmm D, YYYY HH AM/PM") & "</A>"
                    Next CurrTemplate
                Next CurrCategory
                asaDates.KeyValueDelimiter = "</TD><TD>"
                asaDates.ItemDelimiter = "</TD></TR>" & Chr(13) & Chr(10) & "<TR><TD>"
                TagContents = "<TR><TD>Serial Date</TD><TD>Get all changes since this date:</TD></TR><TR><TD>" & asaDates.All & "</TD></TR>"
    End Select
End Sub

Private Sub RegistrationFeedback_ProcessTag(ByVal TagName As String, TagContents As String, SendTags As Boolean)
On Error Resume Next
    Select Case UCase(TagName)
           Case "WC@REQUEST"
                TagContents = Request.Form(TagContents)
    End Select
End Sub

Private Sub WebClass_Initialize()
    If Sandy Is Nothing Then
       Set Sandy = CreateObject("SliceAndDice.CSliceAndDice")
    End If
End Sub


Private Sub WebClass_Start()
On Error Resume Next
    Dim PatchFilename As String
    Dim sLine As String
    Dim sRegKey As String
    Dim sCountry As String
    Dim sValid As String
    Dim InvoiceNumber As String
    
    Dim fh As Long
    Dim CurrValue As Long
    Dim DeltaDate As Date
    Dim bOkaySoFar As Boolean
    Dim CurrItem As Variant
    Dim asaX As Object 'SliceAndDice.CAssocArray
    Dim rst As DAO.Recordset

    If Not IsSandyLoadedYet Then
       Sandy.Load "D:\Data\Programming\Projects - Firm Solutions\SliceAndDice\SliceAndDiceCTL.mdb"
       Set db = OpenDatabase("D:\Data\Programming\Projects - Firm Solutions\HTMLSandy\SandyRegistration.mdb", False, False)
       IsSandyLoadedYet = True
    End If

    Response.Clear
    Response.Expires = 0
    Response.Buffer = True

    If Len(Request("BackDoor")) Then
       If Not (Val(Request.QueryString.Item("BackDoor")) > 0 And Len(Request.QueryString.Item("BackDoor")) = 1) Then
          Response.Redirect "http://www.sliceanddice.com/"
          Exit Sub
       End If
       fh = FreeFile
       Open "D:\Data\Programming\UnprocessedSandyRegFiles\" & sReplace(Format(Now, "000000.000000"), ".", "-") & "-" & Format(CLng(Rnd * 1000), "0000") & ".srk" For Output Access Write As #fh
            Print #fh, "[Sandy Registration]"
            For Each CurrItem In Request.QueryString
                Print #fh, CurrItem & "=" & sReplace(sReplace("" & Request.QueryString.Item(CurrItem), Chr(13) & Chr(10), "%$%EOL%$%"), Chr(9), "%$%TAB%$%")
            Next CurrItem
       Close #fh
       If Request.QueryString.Item("BackDoor") = "2" Then
          If Len(Request.QueryString.Item("RegKey")) = 0 Or Len(Request.QueryString.Item("InvoiceNumber")) = 0 Then
             Response.Redirect "http://www.sliceanddice.com/"
             Exit Sub
          End If
          sRegKey = sadDecrypt("EN* " & Scramble(Request.QueryString.Item("RegKey")))
          If lTokenCount(sRegKey, "$$$$") > 14 Then
             Set asaX = CreateObject("SliceAndDice.CAssocArray")
                 asaX.ItemDelimiter = "$$$$"
                 asaX.All = sRegKey
                 CurrValue = 0
                 For CurrValue = 1 To 14
                      bOkaySoFar = (Len(asaX("Value" & Format(CurrValue, "00"))) > 0)
                      If Not bOkaySoFar Then Exit For
                 Next CurrValue

                 If bOkaySoFar Then
                    sCountry = "USA"
                  
                  ' Take care of international month names (backward compatible)
                    If InStr(asaX("Value03"), "Mai") > 0 Or InStr(asaX("Value01"), "Mai") > 0 Then
                       asaX("Value01") = sReplace(asaX("Value01"), "Mai", "May")
                       asaX("Value03") = sReplace(asaX("Value03"), "Mai", "May")
                       sCountry = "Germany"
                    End If

                    If IsDate(CVDate(asaX("Value03"))) Then
                       Err.Clear
                       InvoiceNumber = sadInvoiceDecrypt(Request.QueryString("InvoiceNumber"))
                       If Len("" & InvoiceNumber) > 0 Then

                          If Left(InvoiceNumber, 7) = "VBXTRAS" Then
                             Set rst = db.OpenRecordset("select * from RegKeyVBXTRAS where InvoiceNumber='" & InvoiceNumber & "'", dbOpenDynaset)
                          ElseIf Left(InvoiceNumber, 12) = "VBCodeDotCom" Then
                             Set rst = db.OpenRecordset("select * from RegKeyVBCodeDotCom where InvoiceNumber='" & InvoiceNumber & "'", dbOpenDynaset)
                          ElseIf Left(InvoiceNumber, 10) = "COMPSOURCE" Then
                             Set rst = db.OpenRecordset("select * from RegKeyCompSource where InvoiceNumber='" & InvoiceNumber & "'", dbOpenDynaset)
                          ElseIf Left(InvoiceNumber, 17) = "SANDYVBCODEDOTCOM" Then
                             Set rst = db.OpenRecordset("select * from RegKeyVBCodeDotCom where InvoiceNumber='SANDYVBCODEDOTCOM'", dbOpenDynaset)
                          Else
                             Set rst = db.OpenRecordset("select * from RegKey where InvoiceNumber='" & InvoiceNumber & "'", dbOpenDynaset)
                          End If
                       Else
                          Set rst = db.OpenRecordset("select * from RegKey where Value03=#" & Format(CVDate(asaX("Value03")), "Mmmm D YYYY H:NN:SS AM/PM") & "# and Value05='" & asaX("Value05") & "'", dbOpenDynaset)
                       End If
                       If Err = 0 Then
                          If rst.RecordCount > 0 Then
                             If rst!AccountDisabled Then
                                Response.Write "Account Disabled$$$$" & rst!ReasonForDisable
                             ElseIf rst!PaymentReceived Then
                                Err.Clear
                                With rst
                                     .Edit
                                        !LicensesRemaining = !LicensesRemaining - 1
                                        !LastRequest = Now
                                        sValid = Scramble(Mid(sadEncrypt("Value01=" & CDbl(!Value01) & "$$$$" & "Value03=" & CDbl(!Value03) & "$$$$" & "Value05=" & !Value05 & "$$$$" & "Value08=" & !Value08 & "$$$$" & "Value10=" & !Value10 & "$$$$" & "Value12=" & !Value12 & "$$$$" & "Value14=" & !Value14), 5))
                                        
                                        If !LicensesRemaining <= -2 Then
                                           !AccountDisabled = True
                                           !ReasonForDisable = "Excessive registration. Please call Firm Solutions at 1-888-311-6876 to resolve this problem."
                                           .Update
                                           Response.Write "Account Disabled$$$$" & !ReasonForDisable
                                        ElseIf !LicensesRemaining < 0 Then
                                           !DateOfLastSuccess = Now
                                           .Update
                                           Response.Write "Out of Licenses$$$$" & sValid
                                        Else
                                           !DateOfLastSuccess = Now
                                           .Update
                                           Response.Write "Valid$$$$" & sValid
                                        End If
                                End With
                             Else
                                If Len(InvoiceNumber) Then
                                   If InvoiceNumber <> rst!InvoiceNumber Then
                                      With rst
                                           .Edit
                                              !InvoiceNumber = InvoiceNumber
                                              !LastRequest = Now
                                           .Update
                                      End With
                                      Response.Write "Invoice Number Updated"
                                   Else
                                      Response.Write "Payment not received"
                                   End If
                                Else
                                   Response.Write "Payment not received"
                                End If
                             End If
                             rst.Close
                          Else
                             rst.Close
                             Err.Clear
                             If Left(InvoiceNumber, 7) = "VBXTRAS" Then
                                Set rst = db.OpenRecordset("RegKeyVBXTRAS", dbOpenTable)
                             ElseIf Left(InvoiceNumber, 12) = "VBCodeDotCom" Then
                                Set rst = db.OpenRecordset("RegKeyVBCodeDotCom", dbOpenTable)
                             ElseIf Left(InvoiceNumber, 10) = "COMPSOURCE" Then
                                Set rst = db.OpenRecordset("RegKeyCOMPSOURCE", dbOpenTable)
                             ElseIf Left(InvoiceNumber, 17) = "SANDYVBCODEDOTCOM" Then
                                Set rst = db.OpenRecordset("RegKeyVBXTRAS", dbOpenTable)
                             Else
                                Set rst = db.OpenRecordset("RegKey", dbOpenTable)
                             End If
                             With rst
                                  .AddNew
                                    !Value01 = CVDate(asaX("Value01"))
                                    !Value03 = CVDate(asaX("Value03"))
                                    !Value05 = asaX("Value05")
                                    !Value08 = asaX("Value08")
                                    !Value10 = asaX("Value10")
                                    !Value12 = asaX("Value12")
                                    !Value14 = asaX("Value14")
                                    If Len(InvoiceNumber) Then
                                       !InvoiceNumber = InvoiceNumber
                                    Else
                                       !InvoiceNumber = "SADVBCODEDOTCOM"
                                    End If
                                    If Left(InvoiceNumber, 7) = "VBXTRAS" Then
                                       !ProductID = Val(Mid(InvoiceNumber, 8, 2))
                                       !Value08 = Format(!Value03, "00000.00000")
                                       !Value08 = Left(!Value08, 5) & "-" & Format(!RegKeyID Mod 999, "000") & "-" & Mid(!Value08, 7)
                                       !LastRequest = Now
                                       !PaymentReceived = True
                                       !DateOfLastSuccess = Now
                                       !DateOfActivation = Now
                                       !NumberOfLicenses = IIf(Val(Mid(InvoiceNumber, 8, 2)) = 2, 2, 1) * IIf(Val(Mid(InvoiceNumber, 10, 3)) < 1, 1, Val(Mid(InvoiceNumber, 10, 3)))
                                       If !NumberOfLicenses < 1 Then !NumberOfLicenses = 1
                                       !LicensesRemaining = !NumberOfLicenses
                                       !Country = sCountry
                                    ElseIf Left(InvoiceNumber, 10) = "COMPSOURCE" Then
                                       !ProductID = Val(Mid(InvoiceNumber, 11, 2))
                                       !Value08 = Format(!Value03, "00000.00000")
                                       !Value08 = Left(!Value08, 5) & "-" & Format(!RegKeyID Mod 999, "000") & "-" & Mid(!Value08, 7)
                                       !LastRequest = Now
                                       !PaymentReceived = True
                                       !DateOfLastSuccess = Now
                                       !DateOfActivation = Now
                                       If Len(InvoiceNumber) = 21 Then
                                            'COMPSOURCE02001000001
                                            ' 1-10
                                            '11-12
                                            '13-15
                                            '16-21
                                          !NumberOfLicenses = IIf(Val(Mid(InvoiceNumber, 11, 2)) = 2, 2, 1) * IIf(Val(Mid(InvoiceNumber, 13, 3)) < 1, 1, Val(Mid(InvoiceNumber, 13, 3)))
                                          If !NumberOfLicenses < 1 Then !NumberOfLicenses = 1
                                       ElseIf Len(InvoiceNumber) = 20 Then
                                            'COMPSOURCE0201000004
                                            ' 1-10
                                            '11-12
                                            '13-14
                                            '15-20
                                          !NumberOfLicenses = IIf(Val(Mid(InvoiceNumber, 11, 2)) = 2, 2, 1) * IIf(Val(Mid(InvoiceNumber, 13, 2)) < 1, 1, Val(Mid(InvoiceNumber, 13, 2)))
                                          If !NumberOfLicenses < 1 Then !NumberOfLicenses = 1
                                       Else
                                          !NumberOfLicenses = 1
                                       End If
                                       !LicensesRemaining = !NumberOfLicenses
                                       !Country = sCountry
                                    ElseIf Left(InvoiceNumber, 12) = "VBCodeDotCom" Or Left(InvoiceNumber, 17) = "SANDYVBCODEDOTCOM" Then
                                       !ProductID = 5
                                       !Value08 = Format(!Value03, "00000.00000")
                                       !Value08 = Left(!Value08, 5) & "-" & Format(!RegKeyID Mod 999, "000") & "-" & Mid(!Value08, 7)
                                       !LastRequest = Now
                                       !PaymentReceived = True
                                       !DateOfLastSuccess = Now
                                       !DateOfActivation = Now
                                       !NumberOfLicenses = 1
                                       !LicensesRemaining = 1
                                       !Country = sCountry
                                    Else
                                       !LastRequest = Now
                                       !NumberOfLicenses = 0
                                       !LicensesRemaining = 0
                                       !Country = sCountry
                                    End If
                                  sValid = Scramble(Mid(sadEncrypt("Value01=" & CDbl(!Value01) & "$$$$" & "Value03=" & CDbl(!Value03) & "$$$$" & "Value05=" & !Value05 & "$$$$" & "Value08=" & !Value08 & "$$$$" & "Value10=" & !Value10 & "$$$$" & "Value12=" & !Value12 & "$$$$" & "Value14=" & !Value14), 5))
                                  .Update
                             End With
                             rst.Close
                             If Err <> 0 Then
                                Response.Write "Unable to process request"
                             Else
                                Response.Write "Valid$$$$" & sValid
                             End If
                          End If
                       Else
                          Response.Write "Unable to process request"
                       End If
                       Set rst = Nothing
                    Else
                       Response.Write "Invalid"
                    End If
                 Else
                    Response.Write "Invalid"
                 End If
             Set asaX = Nothing
          Else
             Response.Write "Invalid"
          End If
       Else
          With Response
               .Redirect "http://www.icatmall.com/sliceanddicereg"
               '.Write "<HTML>"
               '.Write "<HEAD><meta http-equiv=""refresh"" content=""1; url=http://www.icatmall.com/sliceanddicereg""></HEAD>"
               '.Write "<BODY background=""http://www.redshift.com/~jon-tom/images/cool_tile.gif"" bgColor=#ffffff>"
               '.Write "    <CENTER>"
               '.Write "        <H1>Welcome<BR></H1>"
               '.Write "        <H3>To the Slice and Dice Online Registration System.</H3>"
               '.Write "<BR><HR><BR>"
               '.Write "<H3>Please click <A HREF=""http://www.icatmall.com/sliceanddicereg"">here</A> to continue (You'll be redirected automatically)."
               '.Write "</BODY>"
               '.Write "</HTML>"
          End With
       End If
      'BackdoorSandyRegistration.WriteTemplate

'    ElseIf Len(Request("LogRegisteration")) Then
'       fh = FreeFile
'       Open "D:\Data\Programming\UnprocessedSandyRegFiles\" & sReplace(Format(Now, "000000.000000"), ".", "-") & "-" & Format(CLng(Rnd * 1000), "0000") & ".srg" For Output Access Write As #fh
'            Print #fh, "[Sandy Registration]"
'            For Each CurrItem In Request.Form
'                Print #fh, CurrItem & "=" & sReplace(sReplace("" & Request.Form(CurrItem), Chr(13) & Chr(10), "%$%EOL%$%"), Chr(9), "%$%TAB%$%")
'            Next CurrItem
'       Close #fh
'       RegistrationFeedback.WriteTemplate
    ElseIf Len(Request("GeneratePatchFile")) Then
       DeltaDate = CVDate(Request("GeneratePatchFile"))
       
       PatchFilename = "MDBPatch" & sReplace(Format(DeltaDate, "00000.00"), ".", "-") & ".sad"
       With Response
            .Write "<HTML><BODY>"
            If Sandy.GenerateDeltaPatchFile(DeltaDate, App.Path & IIf(Right(App.Path, 1) <> "\", "\", "") & PatchFilename) Then
               .Write "<A HREF=""" & PatchFilename & """>Click here to download the Sandy Patch File</A> for all changes since " & Request("GeneratePatchFile")
            Else
               .Write "Unable to generate a patch file for that date. Team Server may be busy with another request. Please try again later."
            End If
            .Write "</BODY></HTML>"
       End With
    ElseIf Len(Request("List")) Then
       PatchList.WriteTemplate
    ElseIf Len(Request("CategoryName")) And Len(Request("TemplateName")) = 0 Then
       Category.WriteTemplate
    ElseIf Len(Request("TemplateName")) <> 0 And Len(Request("CategoryName")) <> 0 Then
       With Response
            .Write "<html>"
            .Write "<body>"
            
            With Sandy.Categorys(Request("CategoryName")).Templates(Request("TemplateName"))
                 Response.Write "<h1><font face=""Arial"">Contents of Template: " & .Key & "</font></h1>"
                 If Len(.memoCodeAtCursor) Then
                    Response.Write "<HR><HR><H2>At Cursor</H2><HR>"
                    Response.Write "<BLOCKQUOTE>" & sCodeAreaToHTML(.memoCodeAtCursor) & "</BLOCKQUOTE>"
                 End If
                 If Len(.memoCodeAtTop) Then
                    Response.Write "<HR><HR><H2>At Top</H2><HR>"
                    Response.Write "<BLOCKQUOTE>" & sCodeAreaToHTML(.memoCodeAtTop) & "</BLOCKQUOTE>"
                 End If
                 If Len(.memoCodeAtBottom) Then
                    Response.Write "<HR><HR><H2>At Bottom</H2><HR>"
                    Response.Write "<BLOCKQUOTE>" & sCodeAreaToHTML(.memoCodeAtBottom) & "</BLOCKQUOTE>"
                 End If
            End With
            .Write "</body>"
            .Write "</html>"
       End With
    ElseIf Len(Request("CTL")) <> 0 Then
       Categorys.WriteTemplate
    Else
       Response.Redirect "http://www.sliceanddice.com"
    End If
End Sub

