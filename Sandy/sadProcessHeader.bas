Attribute VB_Name = "modGeneral"
Option Explicit

Public Function LoadFormPosition(frmToActOn As Form, Optional ByVal bAutoCenter = True)
1        Dim ProductName As String
    
2        If Len(ProductName) = 0 Then
3           ProductName = "Your Product Name Here"
4        Else
5           ProductName = App.ProductName
6        End If
    
7        If GetSetting(ProductName, frmToActOn.Name, "Position Saved", False) Then
8           frmToActOn.Left = GetSetting(ProductName, frmToActOn.Name, "Form Position Left", frmToActOn.Left)
9           frmToActOn.Top = GetSetting(ProductName, frmToActOn.Name, "Form Position Top", frmToActOn.Top)
10          frmToActOn.Width = GetSetting(ProductName, frmToActOn.Name, "Form Position Width", frmToActOn.Width)
11          frmToActOn.Height = GetSetting(ProductName, frmToActOn.Name, "Form Position Height", frmToActOn.Height)
12       ElseIf bAutoCenter Then
13          frmToActOn.Left = (Screen.Width - frmToActOn.Width) / 2
14          frmToActOn.Top = (Screen.Height - frmToActOn.Height) / 2
15       End If
End Function

Public Function SaveFormPosition(frmToActOn As Form)
16       Dim ProductName As String
    
17       If Len(ProductName) = 0 Then
18          ProductName = "Your Product Name Here"
19       Else
20          ProductName = App.ProductName
21       End If
    
22       SaveSetting ProductName, frmToActOn.Name, "Position Saved", True
23       SaveSetting ProductName, frmToActOn.Name, "Form Position Left", frmToActOn.Left
24       SaveSetting ProductName, frmToActOn.Name, "Form Position Top", frmToActOn.Top
25       SaveSetting ProductName, frmToActOn.Name, "Form Position Width", frmToActOn.Width
26       SaveSetting ProductName, frmToActOn.Name, "Form Position Height", frmToActOn.Height
End Function

