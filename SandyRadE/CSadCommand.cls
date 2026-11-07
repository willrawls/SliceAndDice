VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSadCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sandy - Soft Command"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public SoftCommandName As String
Public Aliases As String
Public Parameters As String
Public Syntax As String

Public Results As String
Public Errors As String

Public OneLineDescription As String
Public LongDescription As String
Public Comments As String
Public SeeAlso As String
Public Examples As String
Public HelpFile As String
Public HelpTopic As Long

Public IsInline As Boolean
Public Index As Long

