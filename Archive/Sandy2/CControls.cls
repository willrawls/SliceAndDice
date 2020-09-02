VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyControls As VBIDE.VBControls

Implements SandySupport.IControls

Private Sub Class_Terminate()
    Set MyControls = Nothing
End Sub

Private Function IControls_Add(ByVal sProgID As String) As SandySupport.IControl
    Dim X As IControl

    Set X = New CControl
        Set X.Link = MyControls.Add(sProgID)
        Set IControls_Add = X
    Set X = Nothing
End Function

Private Property Get IControls_Count() As Long
    IControls_Count = MyControls.Count
End Property

Private Property Get IControls_Item(vIndex As Variant, Optional ByVal ControlIndex As Integer) As SandySupport.IControl
    Dim X As IControl

    Set X = New CControl
        If IsMissing(ControlIndex) Then
           Set X.Link = MyControls.Item(vIndex).ContainedVBControls
        Else
           Set X.Link = MyControls.Item(vIndex, ControlIndex).ContainedVBControls
        End If
        Set IControls_Item = X
    Set X = Nothing
End Property

Private Property Set IControls_Link(ByVal RHS As Object)
    Set MyControls = RHS
End Property

Private Property Get IControls_Link() As Object
    Set IControls_Link = MyControls
End Property


Private Property Get IControls_NewEnum() As stdole.IUnknown
    Set IControls_NewEnum = MyControls.[NewEnum]
End Property
