VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmCustomReportsPreview 
   Caption         =   "Custom Report"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1132
   Icon            =   "frmCustomReportsPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9285
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7965
      TabIndex        =   1
      Top             =   6075
      Width           =   1200
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "&Output..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6630
      TabIndex        =   0
      Top             =   6075
      Width           =   1200
   End
   Begin SSDataWidgets_B.SSDBGrid grdOutput 
      Height          =   5880
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   9075
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      Col.Count       =   0
      DividerType     =   0
      BevelColorFrame =   0
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      BalloonHelp     =   0   'False
      RowNavigation   =   1
      MaxSelectedRows =   0
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      TabNavigation   =   1
      _ExtentX        =   16007
      _ExtentY        =   10372
      _StockProps     =   79
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid_Selected 
      Height          =   465
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   6345
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      Col.Count       =   0
      DividerType     =   0
      BevelColorFrame =   0
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      BalloonHelp     =   0   'False
      RowNavigation   =   1
      MaxSelectedRows =   0
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      TabNavigation   =   1
      _ExtentX        =   11192
      _ExtentY        =   820
      _StockProps     =   79
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid grdBreak 
      Height          =   1200
      Left            =   -500
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   400
      _Version        =   196617
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      Col.Count       =   0
      DividerType     =   0
      BevelColorFrame =   0
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      BalloonHelp     =   0   'False
      RowNavigation   =   1
      MaxSelectedRows =   0
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      TabNavigation   =   1
      _ExtentX        =   706
      _ExtentY        =   2117
      _StockProps     =   79
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCustomReportsPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjParent As clsCustomReportsRUN
Private mstrErrorString As String
Private mblnReportHasSummaryInfo As Boolean
Private mstrBaseTable As String

Private Const dblFORM_MINWIDTH = 9400
Private Const dblFORM_MINHEIGHT = 7000

Public Property Let BaseTable(sBaseTable As String)
  mstrBaseTable = sBaseTable
End Property

Public Property Get BaseTable() As String
  BaseTable = mstrBaseTable
End Property

Public Property Let HasSummaryInfo(ByVal blnNewValue As Boolean)
  mblnReportHasSummaryInfo = blnNewValue
End Property

Public Property Get ErrorMessage() As String
  ErrorMessage = mstrErrorString
End Property

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdOutput_Click()
  mobjParent.OutputReport True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()
  Me.Height = GetPCSetting("CustomReportPreview", "Height", dblFORM_MINHEIGHT)
  Me.Width = GetPCSetting("CustomReportPreview", "Width", dblFORM_MINWIDTH)
    
  Hook Me.hwnd, dblFORM_MINWIDTH, dblFORM_MINHEIGHT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  SavePCSetting "CustomReportPreview", "Height", Me.Height
  SavePCSetting "CustomReportPreview", "Width", Me.Width
  Unload frmOutputOptions
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
  
  Const lngGap As Long = 120

  If Me.WindowState = vbMinimized Then Exit Sub
  
  cmdClose.Top = Me.ScaleHeight - (cmdClose.Height + lngGap)
  cmdClose.Left = Me.ScaleWidth - (cmdClose.Width + lngGap)
  cmdOutput.Top = cmdClose.Top
  cmdOutput.Left = cmdClose.Left - (cmdOutput.Width + lngGap)

  Me.grdOutput.Move lngGap, lngGap, Me.ScaleWidth - (lngGap * 2), cmdClose.Top - (lngGap * 2)

  ' Get rid of the icon off the form
  RemoveIcon Me

End Sub

Public Property Let Parent(ByVal objNewValue As clsCustomReportsRUN)
  Set mobjParent = objNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hwnd
End Sub

