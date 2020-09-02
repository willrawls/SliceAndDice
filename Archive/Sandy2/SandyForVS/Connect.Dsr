VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9945
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   17542
   _Version        =   393216
   Description     =   "Slice and Dice 2 for Visual Studio"
   DisplayName     =   "Sandy 2 for VS"
   AppName         =   "Microsoft Development Environment"
   AppVer          =   "Microsoft Development Environment 6.0 (User)"
   LoadName        =   "None"
   RegLocation     =   "HKEY_LOCAL_MACHINE\Software\Microsoft\VisualStudio\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private SandyIDE As CSandyIDE

Private Sub AddinInstance_Initialize()
    '
End Sub

Private Sub AddinInstance_OnAddInsUpdate(custom() As Variant)
    '
End Sub

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)
    '
End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
On Error Resume Next
    Set SandyIDE = CreateObject("SandyForVS.CSandyIDE")
    If SandyIDE Is Nothing Then
       MsgBox "Failed to create a 'SandyForVS.CSandyIDE' object. Can't start Slice and Dice."
       Exit Sub
    End If

    SandyIDE.OnConnection Application
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
On Error Resume Next
    If Not SandyIDE Is Nothing Then
       SandyIDE.OnDisconnection
       Set SandyIDE = Nothing
    End If
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    '
End Sub

Private Sub AddinInstance_Terminate()
    '
End Sub

