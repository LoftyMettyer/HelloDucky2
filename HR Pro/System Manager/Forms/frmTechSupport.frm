VERSION 5.00
Object = "{AD837810-DD1E-44E0-97C5-854390EA7D3A}#3.2#0"; "COA_Navigation.ocx"
Begin VB.Form frmTechSupport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Support"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5002
   Icon            =   "frmTechSupport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmTechSupport.frx":000C
   ScaleHeight     =   2355
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Contacts for Support :"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   100
      Width           =   6375
      Begin COANavigation.COA_Navigation navSupport 
         Height          =   210
         Left            =   1350
         TabIndex        =   6
         Top             =   1110
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   370
         Caption         =   "navSupport"
         DisplayType     =   0
         NavigateIn      =   0
         NavigateTo      =   ""
         InScreenDesigner=   0   'False
         ColumnID        =   0
         ColumnName      =   ""
         Selected        =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         FontSize        =   8.25
         FontStrikethrough=   0   'False
         FontUnderline   =   -1  'True
         ForeColor       =   16711680
         BackColor       =   -2147483633
         NavigateOnSave  =   0   'False
      End
      Begin COANavigation.COA_Navigation navEmail 
         Height          =   210
         Left            =   1350
         TabIndex        =   7
         Top             =   750
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   370
         Caption         =   "navEmail"
         DisplayType     =   0
         NavigateIn      =   0
         NavigateTo      =   ""
         InScreenDesigner=   0   'False
         ColumnID        =   0
         ColumnName      =   ""
         Selected        =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         FontSize        =   8.25
         FontStrikethrough=   0   'False
         FontUnderline   =   -1  'True
         ForeColor       =   16711680
         BackColor       =   -2147483633
         NavigateOnSave  =   0   'False
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone :"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   5
         Top             =   400
         Width           =   1020
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email :"
         Height          =   215
         Index           =   3
         Left            =   195
         TabIndex        =   4
         Top             =   750
         Width           =   600
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web site :"
         Height          =   215
         Index           =   4
         Left            =   195
         TabIndex        =   3
         Top             =   1110
         Width           =   870
      End
      Begin VB.Label lblTel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+44 (0) 01582 714800"
         Height          =   195
         Left            =   1350
         TabIndex        =   2
         Top             =   405
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   5160
      TabIndex        =   0
      Top             =   1800
      Width           =   1320
   End
End
Attribute VB_Name = "frmTechSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
  UnLoad Me
End Sub


Private Sub Form_Load()
  
  lblTel.Caption = GetSystemSetting("Support", "Telephone No", "")
  
  navEmail.Caption = GetSystemSetting("Support", "Email", "")
  navEmail.NavigateTo = "mailto:" & navEmail.Caption
  
  navSupport.Caption = GetSystemSetting("Support", "Webpage", "")
  navSupport.NavigateTo = navSupport.Caption

End Sub


Private Sub Form_Resize()
  DisplayApplication
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case KeyCode = vbKeyEscape
    UnLoad Me
End Select
End Sub


