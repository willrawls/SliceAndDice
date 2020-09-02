VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyMember As VBIDE.Member

Implements SandySupport.IMember

Private Sub Class_Terminate()
On Error Resume Next
    Set MyMember = Nothing
End Sub


Private Property Let IMember_Description(ByVal RHS As String)
On Error Resume Next
    MyMember.Description = RHS
End Property

Private Property Get IMember_Description() As String
On Error Resume Next
    IMember_Description = MyMember.Description
End Property

Private Property Let IMember_Hidden(ByVal RHS As Boolean)
On Error Resume Next
    MyMember.Hidden = RHS
End Property

Private Property Get IMember_Hidden() As Boolean
On Error Resume Next
    IMember_Hidden = MyMember.Hidden
End Property


Private Property Set IMember_Link(ByVal RHS As Object)
On Error Resume Next
    Set MyMember = RHS
End Property

Private Property Get IMember_Link() As Object
On Error Resume Next
    Set IMember_Link = MyMember
End Property

Private Property Get IMember_MemberType() As Long
On Error Resume Next
    IMember_MemberType = MyMember.Type
End Property

Private Property Get IMember_Name() As String
On Error Resume Next
    IMember_Name = MyMember.Name
End Property


Private Property Let IMember_StandardMethod(ByVal RHS As Long)
On Error Resume Next
    MyMember.StandardMethod = RHS
End Property

Private Property Get IMember_StandardMethod() As Long
On Error Resume Next
    IMember_StandardMethod = MyMember.StandardMethod
End Property


