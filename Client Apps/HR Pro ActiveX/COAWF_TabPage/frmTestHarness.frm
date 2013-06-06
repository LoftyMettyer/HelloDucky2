VERSION 5.00
Object = "{66DD2720-DB90-4D94-963B-369CC9DC8BF8}#5.4#0"; "COAWF_TabPage.ocx"
Begin VB.Form frmTestHarness 
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   420
      Left            =   6300
      TabIndex        =   5
      Top             =   1845
      Width           =   555
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000080FF&
      Height          =   825
      Left            =   4500
      ScaleHeight     =   765
      ScaleWidth      =   1575
      TabIndex        =   4
      Top             =   3420
      Width           =   1635
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Dock"
      Height          =   600
      Left            =   4680
      TabIndex        =   3
      Top             =   1710
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Munipulaite"
      Height          =   600
      Left            =   3195
      TabIndex        =   2
      Top             =   225
      Width           =   1185
   End
   Begin COAWFTabPage.COAWF_TabPage objTabPages 
      Height          =   3435
      Left            =   630
      TabIndex        =   1
      Top             =   945
      Width           =   3615
      _extentx        =   6376
      _extenty        =   6059
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

Private Sub Command3_Click()
  '+ objTabPages.ClientLeft
'  For Each ctlPictureBox In objTabContainer
    Picture1.Move objTabPages.Left + objTabPages.ClientLeft, objTabPages.Top + objTabPages.ClientTop, _
      objTabPages.clientWidth, objTabPages.ClientHeight
'  Next ctlPictureBox



End Sub


Private Sub Command4_Click()

  objTabPages.Caption = "hello;there;ducky"

End Sub

Private Sub objTabPages_Click()
MsgBox "click"
End Sub

Private Sub objTabPages_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MsgBox "mousedown"
End Sub
