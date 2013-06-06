VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.1#0"; "CODEJO~2.OCX"
Begin VB.Form frmHiddenStyle 
   Caption         =   "Hidden Form To Apply Styles"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComctlLib.ListView ListView1 
      Height          =   2310
      Left            =   3780
      TabIndex        =   2
      Top             =   2895
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   4075
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin SSActiveTreeView.SSTree SSTree1 
      Height          =   2295
      Left            =   960
      TabIndex        =   1
      Top             =   2880
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4048
      _Version        =   65536
      Indentation     =   570
      PictureBackgroundUseMask=   0   'False
      HasFont         =   0   'False
      HasMouseIcon    =   0   'False
      HasPictureBackground=   0   'False
      ImageList       =   "(None)"
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   120
      Top             =   120
      _Version        =   851969
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hidden Form to Apply Styles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
End
Attribute VB_Name = "frmHiddenStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
End Sub

Private Sub Form_Load()
  LoadSkin Me, Me.SkinFramework1
End Sub

