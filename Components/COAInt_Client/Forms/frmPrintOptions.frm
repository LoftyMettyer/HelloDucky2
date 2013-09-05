VERSION 5.00
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmPrintOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Options"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrintOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOrient 
      Caption         =   "Orientation :"
      Height          =   1755
      Left            =   100
      TabIndex        =   4
      Top             =   1005
      Width           =   3400
      Begin VB.Frame fraPortrait 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   1305
         Index           =   1
         Left            =   1780
         TabIndex        =   28
         Top             =   260
         Width           =   1050
         Begin VB.Image imgPortrait 
            Height          =   1305
            Left            =   0
            Picture         =   "frmPrintOptions.frx":000C
            Top             =   0
            Width           =   1050
         End
      End
      Begin VB.Frame fraPortrait 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   1305
         Index           =   0
         Left            =   1860
         TabIndex        =   27
         Top             =   350
         Width           =   1050
      End
      Begin VB.Frame fraLandscape 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   1035
         Index           =   1
         Left            =   1470
         TabIndex        =   26
         Top             =   360
         Width           =   1575
         Begin VB.Image imgLandscape 
            Height          =   1035
            Left            =   0
            Picture         =   "frmPrintOptions.frx":0C1F
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.OptionButton optLand 
         Caption         =   "&Landscape"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "Po&rtrait"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1000
      End
      Begin VB.Frame fraLandscape 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   1035
         Index           =   0
         Left            =   1560
         TabIndex        =   25
         Top             =   450
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5840
      TabIndex        =   24
      Top             =   4185
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   400
      Left            =   4520
      TabIndex        =   23
      Top             =   4185
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options :"
      Height          =   3030
      Left            =   3640
      TabIndex        =   16
      Top             =   1005
      Width           =   3385
      Begin VB.CheckBox chkGridLines 
         Caption         =   "Grid lines"
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   2265
         Value           =   1  'Checked
         Width           =   2130
      End
      Begin VB.CheckBox chkShading 
         Caption         =   "Shading and colours"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   2655
         Value           =   1  'Checked
         Width           =   2130
      End
      Begin VB.CheckBox chkCollate 
         Caption         =   "Collate Copies"
         Enabled         =   0   'False
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   810
         Value           =   1  'Checked
         Width           =   1550
      End
      Begin VB.CheckBox chkHeadingsEveryPage 
         Caption         =   "Headings on every page"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   1890
         Value           =   1  'Checked
         Width           =   2500
      End
      Begin COASpinner.COA_Spinner ASRSpinner1 
         Height          =   315
         Left            =   2460
         TabIndex        =   18
         Top             =   315
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
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
         MaximumValue    =   9999
         MinimumValue    =   1
         Text            =   "1"
      End
      Begin VB.Label lblCopies 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of C&opies :"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   375
         Width           =   1680
      End
      Begin VB.Image imgCollateFalse 
         Height          =   630
         Left            =   1675
         Picture         =   "frmPrintOptions.frx":182C
         Top             =   840
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Image imgCollateTrue 
         Height          =   810
         Left            =   1720
         Picture         =   "frmPrintOptions.frx":1D28
         Top             =   855
         Width           =   1410
      End
   End
   Begin VB.Frame fraMargins 
      Caption         =   "Margins (cm) :"
      Height          =   1200
      Left            =   100
      TabIndex        =   7
      Top             =   2835
      Width           =   3400
      Begin VB.TextBox txtMargin 
         Height          =   315
         Index           =   0
         Left            =   705
         TabIndex        =   9
         Text            =   "2.5"
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txtMargin 
         Height          =   315
         Index           =   1
         Left            =   2425
         TabIndex        =   11
         Text            =   "2.5"
         Top             =   285
         Width           =   585
      End
      Begin VB.TextBox txtMargin 
         Height          =   315
         Index           =   2
         Left            =   705
         TabIndex        =   13
         Text            =   "2.5"
         Top             =   750
         Width           =   585
      End
      Begin VB.TextBox txtMargin 
         Height          =   315
         Index           =   3
         Left            =   2425
         TabIndex        =   15
         Text            =   "2.5"
         Top             =   735
         Width           =   585
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top :"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblBottom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom :"
         Height          =   195
         Left            =   1585
         TabIndex        =   10
         Top             =   345
         Width           =   750
      End
      Begin VB.Label lblLeft 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left :"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   810
         Width           =   450
      End
      Begin VB.Label lblRight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right :"
         Height          =   195
         Left            =   1585
         TabIndex        =   14
         Top             =   795
         Width           =   570
      End
   End
   Begin VB.Frame fraDefault 
      Caption         =   "Printer :"
      Height          =   855
      Left            =   100
      TabIndex        =   0
      Top             =   60
      Width           =   6905
      Begin VB.ComboBox cboPrinter 
         Height          =   315
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   315
         Width           =   4305
      End
      Begin VB.CommandButton cmdSetDefault 
         Caption         =   "Set As &Default"
         Height          =   315
         Left            =   5340
         TabIndex        =   3
         Top             =   315
         Width           =   1400
      End
      Begin VB.Label lblUsePrinter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printer :"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   375
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmPrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnGridlines As Boolean
Private mblnPortrait As Boolean
Private mlngCopies As Long
Private mblnCollateCopies As Boolean
Private mblnShading As Boolean
Private mblnHeadingsOnEveryPage As Boolean

Private mintMarginTop As Integer
Private mintMarginBottom As Integer
Private mintMarginLeft As Integer
Private mintMarginRight As Integer

Private mblnCancelled As Boolean
Private mblnLoading As Boolean
Private mbPrintingExpression As Boolean

Const LOCALE_SYSTEM_DEFAULT = &H800
Const LOCALE_USER_DEFAULT = &H400
Const LOCALE_SDATE = &H1D            ' date separator
Const LOCALE_SSHORTDATE = &H1F       ' short date format string
Const LOCALE_SDECIMAL = &HE          ' decimal separator
Const LOCALE_STHOUSAND = &HF         ' thousand separator
Const LOCALE_IMEASURE = &HD          ' Measurement System

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private mstrDefaultPrinterName As String

Public Property Get MarginTop() As Integer
  MarginTop = mintMarginTop
End Property
Public Property Get MarginBottom() As Integer
  MarginBottom = mintMarginBottom
End Property
Public Property Get MarginLeft() As Integer
  MarginLeft = mintMarginLeft
End Property
Public Property Get MarginRight() As Integer
  MarginRight = mintMarginRight
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Get PrintGrid() As Boolean
  PrintGrid = mblnGridlines
End Property

Public Property Get PrintPortrait() As Boolean
  PrintPortrait = mblnPortrait
End Property

Public Property Get PrintCopies() As Long
  PrintCopies = mlngCopies
End Property

Public Property Get HeadingsOnEveryPage() As Boolean
  HeadingsOnEveryPage = mblnHeadingsOnEveryPage
End Property

Public Property Get Shading() As Boolean
  Shading = mblnShading
End Property

Public Property Get CollateCopies() As Boolean
  CollateCopies = mblnCollateCopies
End Property

Private Sub ShowOrientationPreview(pfPortrait As Boolean)
  Dim fraTempFrame As Frame
  
  For Each fraTempFrame In fraPortrait
    fraTempFrame.Visible = pfPortrait
  Next fraTempFrame
  
  For Each fraTempFrame In fraLandscape
    fraTempFrame.Visible = Not pfPortrait
  Next fraTempFrame
  
End Sub

Private Sub cboPrinter_Click()

  On Error GoTo ErrorTrap

  Dim objPrinter As Printer
        
  If Not mblnLoading Then
    ' Set the printer to be what they've selected
    For Each objPrinter In Printers
      If objPrinter.DeviceName = cboPrinter.Text Then
        Set Printer = objPrinter
        Exit For
      End If
    Next
  End If
  
TidyUpAndExit:
  Exit Sub
ErrorTrap:

End Sub

Private Sub chkCollate_Click()

  If chkCollate.Value Then
    Me.imgCollateTrue.Visible = True
    Me.imgCollateFalse.Visible = False
  Else
    Me.imgCollateTrue.Visible = False
    Me.imgCollateFalse.Visible = True
  End If
  
End Sub
Private Sub cmdPrint_Click()

  On Error GoTo ErrorTrap

  mblnCancelled = False
  mblnPortrait = (optPortrait.Value = True)
  mlngCopies = ASRSpinner1.Value
  mblnCollateCopies = (chkCollate = vbChecked)
  mblnGridlines = (chkGridLines = vbChecked)
  mblnShading = (chkShading = vbChecked)
  mblnHeadingsOnEveryPage = (chkHeadingsEveryPage = vbChecked)
  
  If GetSystemMeasurement = "us" Then
    mintMarginTop = Val(txtMargin(0).Text)
    mintMarginBottom = Val(txtMargin(1).Text)
    mintMarginLeft = Val(txtMargin(2).Text)
    mintMarginRight = Val(txtMargin(3).Text)
  Else
    mintMarginTop = Val(txtMargin(0).Text) * 10
    mintMarginBottom = Val(txtMargin(1).Text) * 10
    mintMarginLeft = Val(txtMargin(2).Text) * 10
    mintMarginRight = Val(txtMargin(3).Text) * 10
  End If
  
  Me.Hide
  
TidyUpAndExit:
  Exit Sub
ErrorTrap:

End Sub

Private Sub ASRSpinner1_Change()
  chkCollate.Enabled = (ASRSpinner1.Value > 1 And Not mbPrintingExpression)
End Sub

Private Sub cmdCancel_Click()
  mblnCancelled = True
  Me.Hide
End Sub

Private Sub cmdSetDefault_Click()

  mstrDefaultPrinterName = cboPrinter.Text
  SaveSetting "HR Pro", "Printer", "DeviceName", mstrDefaultPrinterName
  
End Sub

Private Sub Form_Activate()
  cmdPrint.SetFocus
End Sub

Private Sub Form_Load()

  On Error GoTo ErrorTrap

  mblnLoading = True

  ' Retrieve the current default printer and set the combo
  Dim objPrinter As Printer
  Dim intCurrentPrinter As Integer
  Dim bFound As Boolean
  
  mstrDefaultPrinterName = GetSetting("HR Pro", "Printer", "DeviceName", "")
  For Each objPrinter In Printers
    cboPrinter.AddItem objPrinter.DeviceName
    
    'JDM - 07/11/01 - Fault 3102 - Causes crash with some printers - why? I dunno...
    'cboPrinter.ItemData(cboPrinter.NewIndex) = objPrinter.hdc

    If LCase(objPrinter.DeviceName) = LCase(mstrDefaultPrinterName) Then
      Set Printer = objPrinter
      bFound = True
    End If
  Next objPrinter

  If bFound Then
    cboPrinter.Text = mstrDefaultPrinterName
  End If

  If GetSystemMeasurement = "us" Then
    fraMargins.Caption = "Margins (inches) :"
    txtMargin(0).Text = "1"
    txtMargin(1).Text = "1"
    txtMargin(2).Text = "1"
    txtMargin(3).Text = "1"
  End If
  
  ' Get rid of the icon off the form
  RemoveIcon Me
   
  mblnLoading = False

TidyUpAndExit:
  Exit Sub
ErrorTrap:

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  On Error GoTo ErrorTrap

  If UnloadMode = vbFormControlMenu Then
    mblnCancelled = True
    Cancel = True
    Me.Hide
  End If

TidyUpAndExit:
  Exit Sub
ErrorTrap:

End Sub

Private Sub optLand_Click()
  ShowOrientationPreview False
End Sub

Private Sub optPortrait_Click()
  ShowOrientationPreview True
End Sub

Private Sub txtMargin_GotFocus(Index As Integer)

  With txtMargin(Index)
    .SetFocus
    .SelStart = 0
    .SelLength = Len(txtMargin(Index).Text)
  End With
  
End Sub

Private Sub txtMargin_LostFocus(Index As Integer)

  On Error GoTo ErrorTrap

  If GetSystemMeasurement = "us" Then
  
    If IsNumeric(txtMargin(Index).Text) Then
    
      If Val(txtMargin(Index).Text) > 5 Then txtMargin(Index).Text = 5
      If Val(txtMargin(Index).Text) < 0.5 Then txtMargin(Index).Text = 0.5
      GoTo TidyUpAndExit
      
    ElseIf txtMargin(Index).Text = "" Then
      
      txtMargin(Index).Text = 0.5
      GoTo TidyUpAndExit
      
    End If
    
    With txtMargin(Index)
      MsgBox "Margins must be numeric values between 0.5 and 5 inches", vbInformation + vbOKOnly, "Print Options Error"
      .SetFocus
      .SelStart = 0
      .SelLength = Len(txtMargin(Index).Text)
    End With
  
  Else
  
    If IsNumeric(txtMargin(Index).Text) Then
    
      If Val(txtMargin(Index).Text) > 10 Then txtMargin(Index).Text = 10
      If Val(txtMargin(Index).Text) < 1 Then txtMargin(Index).Text = 1
      GoTo TidyUpAndExit
      
    ElseIf txtMargin(Index).Text = "" Then
      
      txtMargin(Index).Text = 1
      GoTo TidyUpAndExit
      
    End If
    
    With txtMargin(Index)
      MsgBox "Margins must be numeric values between 1 and 10 centimetres", vbInformation + vbOKOnly, "Print Options Error"
      .SetFocus
      .SelStart = 0
      .SelLength = Len(txtMargin(Index).Text)
    End With
    
  End If
  
TidyUpAndExit:
  Exit Sub
ErrorTrap:

End Sub

Function GetSystemMeasurement() As String

  On Error GoTo ErrorTrap
  
  ' Return the system measurement (metric or us).
  
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IMEASURE, sBuffer, 99)
  GetSystemMeasurement = Left(sBuffer, lngLength - 1)
  
  If GetSystemMeasurement = 1 Then
    GetSystemMeasurement = "us"
  Else
    GetSystemMeasurement = "metric"
  End If
  
TidyUpAndExit:
  Exit Function
ErrorTrap:

End Function

' Sets the default settings for printing an expression
Public Sub PrintDefinition()

  mbPrintingExpression = True

  'Expressions should default to portrait
  optLand.Value = False
  optPortrait.Value = True
  
  ' Collate copies only apply to reports
  chkCollate.Value = vbUnchecked
  chkCollate.Enabled = False
  
  ' Grid lines only apply to reports
  chkGridLines.Value = vbUnchecked
  chkGridLines.Enabled = False
  
  ' Shading only apply to reports
  chkShading.Value = vbUnchecked
  chkShading.Enabled = False
  
  ' Headings only apply to reports
  chkHeadingsEveryPage.Value = vbUnchecked
  chkHeadingsEveryPage.Enabled = False
  
  ' Don't touch the margins
  txtMargin(0).Enabled = False
  txtMargin(1).Enabled = False
  txtMargin(2).Enabled = False
  txtMargin(3).Enabled = False

End Sub

