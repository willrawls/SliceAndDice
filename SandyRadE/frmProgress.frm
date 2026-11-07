VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Progress Indicator"
   ClientHeight    =   1605
   ClientLeft      =   5670
   ClientTop       =   4710
   ClientWidth     =   6030
   ControlBox      =   0   'False
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pbrProgress 
      Align           =   1  'Align Top
      Height          =   525
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   1050
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   926
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar pbrProgress 
      Align           =   1  'Align Top
      Height          =   525
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   525
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   926
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar pbrProgress 
      Align           =   1  'Align Top
      Height          =   525
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   926
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
