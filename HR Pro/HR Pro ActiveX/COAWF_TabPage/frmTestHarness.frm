VERSION 5.00
Object = "{66DD2720-DB90-4D94-963B-369CC9DC8BF8}#5.1#0"; "COAWF_TabPage.ocx"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Munipulaite"
      Height          =   600
      Left            =   3195
      TabIndex        =   2
      Top             =   225
      Width           =   1185
   End
   Begin COAWFTabPage.COAWF_TabPage objTabPages 
      Height          =   1500
      Left            =   630
      TabIndex        =   1
      Top             =   945
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   2646
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   420
      Left            =   900
      TabIndex        =   0
      Top             =   225
      Width           =   1680
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

objTabPages.AddTabPage "kshdfkjsfh"



End Sub

Private Sub Command2_Click()


objTabPages.AddTabPage "page3"

'MsgBox objTabPages.GetCaptions

MsgBox objTabPages.hWnd



End Sub
