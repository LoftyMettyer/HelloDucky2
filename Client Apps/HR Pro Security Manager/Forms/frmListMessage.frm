VERSION 5.00
Begin VB.Form frmListMessage 
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5370
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1043
   Icon            =   "frmListMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      Height          =   400
      Left            =   3960
      TabIndex        =   1
      Top             =   1320
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   400
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   1200
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Height          =   400
      Left            =   2520
      TabIndex        =   0
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Frame fraList 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2250
      Begin VB.ListBox lstList 
         Height          =   405
         IntegralHeight  =   0   'False
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   1680
      End
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmListMessage.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblMSG 
      Height          =   585
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmListMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FORM_START_WIDTH = 6500

Private Const FORM_START_HEIGHT = 3500


Private intResponse As VbMsgBoxResult
Public Function AddToList(pstrAddString As String) As Boolean
  lstList.AddItem pstrAddString
  AddToList = True
End Function
Public Property Get ListCount() As Integer
  ListCount = lstList.ListCount
End Property

Public Sub ResetList()
  Do While (lstList.ListCount > 0)
    lstList.RemoveItem (0)
  Loop
End Sub


Public Property Get Response() As VbMsgBoxResult
  Response = intResponse
End Property


Public Function ShowMessage(pstrCaption As String, pstrMessage As String)
  Me.Caption = pstrCaption
  SetFormIcon
  Me.lblMSG.Caption = pstrMessage
  Me.Show vbModal
End Function

Private Function PrintUsage() As Boolean
  
  Dim objPrintDef As clsPrintDef
  Dim blnOK As Boolean
  Dim sKey As String
  Dim iLoop As Integer
  
  Set objPrintDef = New clsPrintDef

  Screen.MousePointer = vbHourglass
  
  blnOK = (objPrintDef.IsOK)
  
  If blnOK Then
    With objPrintDef
      If .PrintStart(True) Then
        .PrintHeader Me.Caption & " usage"
  
        For iLoop = 0 To lstList.ListCount - 1
          .PrintNormal lstList.List(iLoop)
        Next iLoop
        
        .PrintEnd
      End If
    End With
  End If

End Function

Private Sub ResizeForm()
  
  Const OFFSET_LEFT = 160
  Const OFFSET_RIGHT = 160
  Const OFFSET_TOP = 160
  Const OFFSET_BOTTOM = 160
  Const MSG_OFFSET = 240
  Const BUTTON_OFFSET = 95
  Const MIN_FORM_HEIGHT = 3375
  Const MIN_FORM_WIDTH = 6435
  
  If Me.Height < MIN_FORM_HEIGHT Then Me.Height = MIN_FORM_HEIGHT
  If Me.Width < MIN_FORM_WIDTH Then Me.Width = MIN_FORM_WIDTH
  
  lblMSG.Width = Me.ScaleWidth - imgIcon(0).Width - (3 * MSG_OFFSET)
  lblMSG.Top = 300
  lblMSG.Left = 960
  lblMSG.Height = 800
  imgIcon(0).Top = MSG_OFFSET
  imgIcon(0).Left = MSG_OFFSET
  imgIcon(0).Width = 480
  imgIcon(0).Height = imgIcon(0).Width
  
  cmdPrint.Top = Me.ScaleHeight - BUTTON_OFFSET - cmdPrint.Height
  cmdPrint.Left = lstList.Left
  cmdYes.Top = Me.ScaleHeight - BUTTON_OFFSET - cmdYes.Height
  cmdYes.Left = Me.ScaleWidth - BUTTON_OFFSET - OFFSET_RIGHT - (2 * cmdYes.Width)
  cmdNo.Top = cmdYes.Top
  cmdNo.Left = cmdYes.Left + cmdYes.Width + BUTTON_OFFSET

  fraList.Top = lblMSG.Top + lblMSG.Height + 120
  fraList.Left = OFFSET_LEFT
  fraList.Width = Me.ScaleWidth - OFFSET_LEFT - OFFSET_RIGHT
  fraList.Height = cmdPrint.Top - BUTTON_OFFSET - fraList.Top
  lstList.Top = OFFSET_TOP
  lstList.Left = OFFSET_LEFT
  lstList.Width = fraList.Width - OFFSET_LEFT - OFFSET_RIGHT
  lstList.Height = fraList.Height - OFFSET_TOP - OFFSET_BOTTOM

End Sub

Private Sub SetFormIcon()
  


End Sub



Private Sub cmdNo_Click()
  intResponse = vbNo
  Unload Me
End Sub

Private Sub cmdYes_Click()
  intResponse = vbYes
  Unload Me
End Sub

Private Sub cmdPrint_Click()
  PrintUsage
End Sub



Private Sub Form_Load()
  Me.Width = FORM_START_WIDTH
  Me.Height = FORM_START_HEIGHT
  
    ' Get rid of the icon off the form
  Me.Icon = Nothing
  SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW Or WS_EX_DLGMODALFRAME
  
End Sub

Private Sub Form_Resize()
  ResizeForm
End Sub





