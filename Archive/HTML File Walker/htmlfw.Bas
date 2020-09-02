Attribute VB_Name = "HTMLFW"
Option Explicit

' -Title(Testing 123)                                           Sets the page's title
' -Dir(W:\Graphics\backgrounds)                                 Sets the path to walk
' -Quiet                                                        No visual interface 8)
' -Output(W:\Graphics\backgrounds\Folder.htm)                   Send results of walk to the named file
' -Build                                                        Use current settings to build the page
' -Put                                                          Place the page on HTTP site
' -ExploreLocal                                                 Cause browser to appear with outputed file's contents.

Public gsEOL As String
Public gsEOLTab As String
Public gs2EOL As String
Public gsTAB As String
Public sLinePrefix As String
Public sLineSuffix As String
Public sOut As String

Public Const gsTimeOut As String = "10"
Public Const gsHTMLFileWalkerHomePage As String = "http://www.firmsolutions.com/HTMLFileWalker.html"

Public gbQuiet As Boolean
Public gbInterruptBuild As Long
Public gbPageGenerated As Boolean

Public gsStartPath As String


Public Sub Main()
    With frmMain
         .Show
         If Len(.txtTitle) = 0 And Len(.txtOutFilename) = 0 And Len(.txtStartPath) = 0 Then
            .cmdReset_Click
         End If
    End With
End Sub

Public Function sCapitalize(ByVal sValue As String)
    Static i As Long
    i = Len(sValue)
    If i > 1 Then
       sCapitalize = UCase(Left(sValue, 1)) & LCase(Mid(sValue, 2))
    ElseIf i = 1 Then
       sCapitalize = UCase(sValue)
    Else
       sCapitalize = ""
    End If
End Function

Public Function sGetExtension(ByVal sPath As String)
    Static sT As String
    sT = sGetFilename(sPath)
    If InStr(sT, ".") > 0 Then
       sGetExtension = sGetToken(sT, nTokens(sT, "."), ".")
      'sGetExtension = Mid(sT, InStr(sT, ".") + 1)
    Else
       sGetExtension = ""
    End If
End Function

Public Function sGetFilename(ByVal sPath As String)
    
    Do While InStr(sPath, "\")
       sPath = Mid(sPath, InStr(sPath, "\") + 1)
    Loop
    sGetFilename = sPath
End Function

Public Function sGetPath(ByVal sPath As String)
    Static i As Long
    Static j As Long

    i = 1
    j = InStr(sPath, "\")
    Do While j
       i = i + j
       j = InStr(Mid(sPath, i), "\")
    Loop
    If i = 1 Then
       sGetPath = ""
    Else
       sGetPath = Left(sPath, i - 1)
    End If
End Function
