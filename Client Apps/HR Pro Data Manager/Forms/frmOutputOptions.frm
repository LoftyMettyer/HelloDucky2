VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOutputOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Output Options"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1049
   Icon            =   "frmOutputOptions.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOutputDestination 
      Caption         =   "Output Destination(s) :"
      Height          =   4080
      Left            =   2740
      TabIndex        =   12
      Top             =   120
      Width           =   6550
      Begin VB.TextBox txtEmailAttachAs 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3495
         TabIndex        =   24
         Tag             =   "0"
         Top             =   2985
         Width           =   2865
      End
      Begin VB.TextBox txtFilename 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3495
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   1275
         Width           =   2565
      End
      Begin VB.ComboBox cboPrinterName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3495
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   765
         Width           =   2865
      End
      Begin VB.ComboBox cboSaveExisting 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3495
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1680
         Width           =   2865
      End
      Begin VB.TextBox txtEmailSubject 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3495
         TabIndex        =   20
         Top             =   2580
         Width           =   2865
      End
      Begin VB.TextBox txtEmailGroup 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3495
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   2175
         Width           =   2565
      End
      Begin VB.CommandButton cmdFilename 
Caption = "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   6060
         TabIndex        =   18
         Top             =   1275
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdEmailGroup 
Caption = "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   6060
         TabIndex        =   17
         Top             =   2175
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CheckBox chkDestination 
         Caption         =   "Displa&y output on screen"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   375
         Value           =   1  'Checked
         Width           =   3105
      End
      Begin VB.CheckBox chkDestination 
         Caption         =   "Send to &printer"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   825
         Width           =   1605
      End
      Begin VB.CheckBox chkDestination 
         Caption         =   "Save to &file"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1335
         Width           =   1455
      End
      Begin VB.CheckBox chkDestination 
         Caption         =   "Send as &email"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   2235
         Width           =   1515
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "Attach as :"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   2025
         TabIndex        =   30
         Top             =   3045
         Width           =   1200
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         Caption         =   "File name :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2025
         TabIndex        =   29
         Top             =   1335
         Width           =   1185
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "Email subject :"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   2025
         TabIndex        =   28
         Top             =   2640
         Width           =   1440
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "Email group :"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   2025
         TabIndex        =   27
         Top             =   2235
         Width           =   1335
      End
      Begin VB.Label lblSave 
         AutoSize        =   -1  'True
         Caption         =   "If existing file :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2025
         TabIndex        =   26
         Top             =   1740
         Width           =   1485
      End
      Begin VB.Label lblPrinter 
         AutoSize        =   -1  'True
         Caption         =   "Printer location :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2025
         TabIndex        =   25
         Top             =   825
         Width           =   1590
      End
   End
   Begin VB.Frame fraOutputFormat 
      Caption         =   "Output Format :"
      Height          =   3240
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2500
      Begin VB.OptionButton optOutputFormat 
         Caption         =   "Excel P&ivot Table"
         Height          =   195
         Index           =   6
         Left            =   200
         TabIndex        =   11
         Top             =   2800
         Width           =   1900
      End
      Begin VB.OptionButton optOutputFormat 
         Caption         =   "Excel Char&t"
         Height          =   195
         Index           =   5
         Left            =   200
         TabIndex        =   10
         Top             =   2400
         Width           =   1900
      End
      Begin VB.OptionButton optOutputFormat 
         Caption         =   "E&xcel Worksheet"
         Height          =   195
         Index           =   4
         Left            =   200
         TabIndex        =   9
         Top             =   2000
         Width           =   1900
      End
      Begin VB.OptionButton optOutputFormat 
         Caption         =   "&Word Document"
         Height          =   195
         Index           =   3
         Left            =   200
         TabIndex        =   8
         Top             =   1600
         Width           =   1900
      End
      Begin VB.OptionButton optOutputFormat 
         Caption         =   "&HTML Document"
         Height          =   195
         Index           =   2
         Left            =   200
         TabIndex        =   7
         Top             =   1200
         Width           =   1900
      End
      Begin VB.OptionButton optOutputFormat 
         Caption         =   "CS&V File"
         Height          =   195
         Index           =   1
         Left            =   200
         TabIndex        =   6
         Top             =   800
         Width           =   1900
      End
      Begin VB.OptionButton optOutputFormat 
         Caption         =   "D&ata Only"
         Height          =   195
         Index           =   0
         Left            =   200
         TabIndex        =   5
         Top             =   400
         Width           =   1900
      End
   End
   Begin VB.Frame fraOutputPage 
      Caption         =   "Data Range :"
      Height          =   800
      Left            =   120
      TabIndex        =   2
      Top             =   3400
      Width           =   2500
      Begin VB.ComboBox cboPageBreak 
         Height          =   315
         ItemData        =   "frmOutputOptions.frx":0D90
         Left            =   200
         List            =   "frmOutputOptions.frx":0D92
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   2150
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   6825
      TabIndex        =   0
      Top             =   4300
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8100
      TabIndex        =   1
      Top             =   4300
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOutputOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objOutputDef As clsOutputDef
Private mlngFormat As Long
Private mblnCancelled As Boolean

Private Sub Form_Load()
  
  mlngFormat = 0
  mblnCancelled = True
  
  Set objOutputDef = New clsOutputDef
  objOutputDef.ParentForm = Me
  objOutputDef.PopulateCombos True, True, True

End Sub

Public Sub ShowFormats(blnData As Boolean, blnCSV As Boolean, blnHTML As Boolean, _
  blnWord As Boolean, blnExcel As Boolean, blnChart As Boolean, blnPivot As Boolean)
  
  objOutputDef.ShowFormats blnData, blnCSV, blnHTML, blnWord, blnExcel, blnChart, blnPivot

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set objOutputDef = Nothing
End Sub

Public Property Get Cancelled() As Variant
  Cancelled = mblnCancelled
End Property

Public Property Let PageRange(blnPageRange As Boolean)
  fraOutputPage.Visible = blnPageRange
  fraOutputFormat.Height = IIf(blnPageRange, 3250, 4080)
  If blnPageRange Then
    EnableCombo cboPageBreak, True
    If cboPageBreak.ListCount > 0 Then
      cboPageBreak.ListIndex = 0
    End If
  End If
End Property

Private Sub optOutputFormat_Click(Index As Integer)
  objOutputDef.FormatClick Index, False
End Sub

Private Sub chkDestination_Click(Index As Integer)
  objOutputDef.DestinationClick Index
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  If objOutputDef.ValidDestination Then
    mlngFormat = objOutputDef.GetSelectedFormatIndex
    mblnCancelled = False
    Me.Hide
  End If
End Sub

Public Property Get Format() As Long
  Format = mlngFormat
End Property

