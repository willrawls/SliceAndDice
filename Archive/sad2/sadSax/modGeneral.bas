Attribute VB_Name = "modGeneral"
Option Explicit

Public Function LoadFormPosition(frmToActOn As Form, Optional ByVal bAutoCenter = True)
    Dim ProductName As String
    
    If Len(ProductName) = 0 Then
       ProductName = "Your Product Name Here"
    Else
       ProductName = App.ProductName
    End If
    
    If GetSetting(ProductName, frmToActOn.Name, "Position Saved", False) Then
       frmToActOn.Left = GetSetting(ProductName, frmToActOn.Name, "Form Position Left", frmToActOn.Left)
       frmToActOn.Top = GetSetting(ProductName, frmToActOn.Name, "Form Position Top", frmToActOn.Top)
       frmToActOn.Width = GetSetting(ProductName, frmToActOn.Name, "Form Position Width", frmToActOn.Width)
       frmToActOn.Height = GetSetting(ProductName, frmToActOn.Name, "Form Position Height", frmToActOn.Height)
    ElseIf bAutoCenter Then
       frmToActOn.Left = (Screen.Width - frmToActOn.Width) / 2
       frmToActOn.Top = (Screen.Height - frmToActOn.Height) / 2
    End If
End Function

Public Function SaveFormPosition(frmToActOn As Form)
    Dim ProductName As String
    
    If frmToActOn.WindowState <> vbNormal Then Exit Function
    
    If Len(ProductName) = 0 Then
       ProductName = "Your Product Name Here"
    Else
       ProductName = App.ProductName
    End If
    
    SaveSetting ProductName, frmToActOn.Name, "Position Saved", True
    SaveSetting ProductName, frmToActOn.Name, "Form Position Left", frmToActOn.Left
    SaveSetting ProductName, frmToActOn.Name, "Form Position Top", frmToActOn.Top
    SaveSetting ProductName, frmToActOn.Name, "Form Position Width", frmToActOn.Width
    SaveSetting ProductName, frmToActOn.Name, "Form Position Height", frmToActOn.Height
End Function

