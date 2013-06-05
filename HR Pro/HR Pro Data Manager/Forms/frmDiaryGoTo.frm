VERSION 5.00
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Begin VB.Form frmDiaryGoTo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Go To Specified Date"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1031
   Icon            =   "frmDiaryGoTo.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGoToDate 
      Caption         =   "&Go to Date"
      Default         =   -1  'True
      Height          =   400
      Left            =   735
      TabIndex        =   1
      Top             =   850
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   400
      Left            =   2010
      TabIndex        =   2
      Top             =   850
      Width           =   1200
   End
   Begin GTMaskDate.GTMaskDate cboDate 
      Height          =   315
      Left            =   1785
      TabIndex        =   0
      Top             =   300
      Width           =   1440
      _Version        =   65537
      _ExtentX        =   2540
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      NullText        =   "__/__/____"
      BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSelect      =   -1  'True
      MaskCentury     =   2
      SpinButtonEnabled=   0   'False
      BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ToolTips        =   0   'False
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date To Display :"
      Height          =   195
      Index           =   1
      Left            =   195
      TabIndex        =   3
      Top             =   360
      Width           =   1500
   End
End
Attribute VB_Name = "frmDiaryGoTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cboDate_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF2 Then
    cboDate.DateValue = Date
  End If

End Sub

Private Sub cboDate_LostFocus()

'  If IsNull(cboDate.DateValue) And Not _
'     IsDate(cboDate.DateValue) And _
'     cboDate.Text <> "  /  /" Then
'
'     MsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     cboDate.DateValue = Null
'     cboDate.SetFocus
'     Exit Sub
'  End If

  'MH20020423 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate cboDate

End Sub


Private Sub cmdClose_Click()
  Unload Me
End Sub


Private Sub cmdGoToDate_Click()

  On Local Error GoTo LocalErr

  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  'Double check valid date in case lost focus is missed
  '(by pressing enter for default button lostfocus is not run)
  If ValidateGTMaskDate(cboDate) = False Then
    Exit Sub
  ElseIf IsDate(cboDate.DateValue) = False Then
    MsgBox "Please enter a valid date", vbExclamation
    cboDate.SetFocus
    Exit Sub
  End If
  
  gobjDiary.DiaryEventID = 0
  gobjDiary.DateSelected = cboDate.DateValue
  Unload Me
  Exit Sub

LocalErr:
  MsgBox "Please enter a valid date", vbExclamation
  cboDate.SetFocus

End Sub


Private Sub Form_Activate()
    
  With cboDate
    .SetFocus
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  
End Sub

Private Sub Form_Load()
  'JPD 20041118 Fault 8231
  UI.FormatGTDateControl cboDate
  
  'Call SetDateComboFormat(cboDate)
  cboDate.DateValue = Date
End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmDiaryGoTo = Nothing
End Sub

