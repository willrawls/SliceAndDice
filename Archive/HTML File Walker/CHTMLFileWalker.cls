VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHTMLFileWalker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public sTitle As String
Public sDir As String
Public bQuiet As Boolean
Public sOutput As String

Private m_sTitle As String
Private m_sDir As String
Private m_Quiet As Boolean
Private m_sOutputFilename As String




Public Function Build() As Boolean
    frmMain.HandleCommandLineActions "-Build"
End Function


Public Function ExploreLocal() As Boolean
    frmMain.HandleCommandLineActions "-ExploreLocal"
End Function


Public Function PutFileToFTP() As Boolean
    frmMain.HandleCommandLineActions "-Put"
End Function


Public Property Get Title() As String
       Title = m_sTitle
End Property

Public Property Let Title(sData As String)
       m_sTitle = sData
       frmMain.HandleCommandLineActions "-Build"
End Property

Public Property Get Dir() As String
       Dir = m_sDir
End Property

Public Property Let Dir(sData As String)
       m_sDir = sData
End Property

Public Property Get Quiet() As Boolean
       Quiet = m_Quiet
End Property

Public Property Let Quiet(vData As Boolean)
       m_Quiet = vData
End Property

Public Property Get OutputFilename() As String
       OutputFilename = m_sOutputFilename
End Property

Public Property Let OutputFilename(sData As String)
       m_sOutputFilename = sData
    frmMain.HandleCommandLineActions vbNullString
End Property

