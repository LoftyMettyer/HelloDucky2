VERSION 5.00
Object = "{66DD2720-DB90-4D94-963B-369CC9DC8BF8}#1.0#0"; "COAWF_TabPage.ocx"
Begin VB.Form frmTestHarness 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   900
      TabIndex        =   1
      Top             =   225
      Width           =   1680
   End
   Begin COAWF_TabPage.COASD_TabPage objTabPages 
      Height          =   1545
      Left            =   765
      TabIndex        =   0
      Top             =   945
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   2725
   End
End
Attribute VB_Name = "frmTestHarness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()



objTabPages.Tabs.Clear
objTabPages.AddTabPage "hhdh"
objTabPages.Tabpage(1).Caption = "hhhe"

objTabPages.AddTabPage "kshdfkjsfh"
objTabPages.Tabpage(2).Caption = "kshdfkjsfh"



End Sub
