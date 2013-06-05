VERSION 5.00
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmPrintOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Options"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1052
   Icon            =   "frmPrintOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDefault 
      Caption         =   "Printer :"
      Height          =   855
      Left            =   105
      TabIndex        =   23
      Top             =   60
      Width           =   6705
      Begin VB.CommandButton cmdSetDefault 
         Caption         =   "Set As &Default"
         Height          =   315
         Left            =   5250
         TabIndex        =   1
         Top             =   315
         Width           =   1380
      End
      Begin VB.ComboBox cboPrinter 
         Height          =   315
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   4305
      End
      Begin VB.Label lblUsePrinter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printer :"
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   375
         Width           =   585
      End
   End
   Begin VB.Frame fraMargins 
      Caption         =   "Margins (cm) :"
      Height          =   1200
      Left            =   105
      TabIndex        =   18
      Top             =   2835
      Width           =   3300
      Begin VB.TextBox txtMargin 
         Height          =   315
         Index           =   3
         Left            =   2325
         TabIndex        =   7
         Text            =   "2.5"
         Top             =   735
         Width           =   585
      End
      Begin VB.TextBox txtMargin 
         Height          =   315
         Index           =   2
         Left            =   705
         TabIndex        =   6
         Text            =   "2.5"
         Top             =   750
         Width           =   585
      End
      Begin VB.TextBox txtMargin 
         Height          =   315
         Index           =   1
         Left            =   2325
         TabIndex        =   5
         Text            =   "2.5"
         Top             =   285
         Width           =   585
      End
      Begin VB.TextBox txtMargin 
         Height          =   315
         Index           =   0
         Left            =   705
         TabIndex        =   4
         Text            =   "2.5"
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblRight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right :"
         Height          =   195
         Left            =   1485
         TabIndex        =   22
         Top             =   795
         Width           =   480
      End
      Begin VB.Label lblLeft 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left :"
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   810
         Width           =   390
      End
      Begin VB.Label lblBottom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom :"
         Height          =   195
         Left            =   1485
         TabIndex        =   20
         Top             =   345
         Width           =   615
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top :"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options :"
      Height          =   3030
      Left            =   3540
      TabIndex        =   16
      Top             =   1005
      Width           =   3285
      Begin VB.CheckBox chkHeadingsEveryPage 
         Caption         =   "Headings on every page"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   1890
         Value           =   1  'Checked
         Width           =   2490
      End
      Begin VB.CheckBox chkCollate 
         Caption         =   "Collate Copies"
         Enabled         =   0   'False
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   810
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkShading 
         Caption         =   "Shading and colours"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   2655
         Value           =   1  'Checked
         Width           =   2130
      End
      Begin VB.CheckBox chkGridLines 
         Caption         =   "Grid lines"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   2265
         Value           =   1  'Checked
         Width           =   2130
      End
      Begin COASpinner.COA_Spinner ASRSpinner1 
         Height          =   315
         Left            =   2415
         TabIndex        =   8
         Top             =   315
         Width           =   735
         _ExtentX        =   1296
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
      Begin VB.Image imgCollateTrue 
         Height          =   810
         Left            =   1710
         Picture         =   "frmPrintOptions.frx":000C
         Top             =   855
         Width           =   1410
      End
      Begin VB.Image imgCollateFalse 
         Height          =   630
         Left            =   1665
         Picture         =   "frmPrintOptions.frx":0509
         Top             =   840
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lblCopies 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of C&opies :"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   375
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   400
      Left            =   4320
      TabIndex        =   13
      Top             =   4185
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5640
      TabIndex        =   14
      Top             =   4185
      Width           =   1200
   End
   Begin VB.Frame fraOrient 
      Caption         =   "Orientation :"
      Height          =   1755
      Left            =   105
      TabIndex        =   15
      Top             =   1005
      Width           =   3285
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
            Picture         =   "frmPrintOptions.frx":0A05
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
            Picture         =   "frmPrintOptions.frx":1618
            Top             =   0
            Width           =   1575
         End
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
      Begin VB.OptionButton optPortrait 
         Caption         =   "Po&rtrait"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   990
      End
      Begin VB.OptionButton optLand 
         Caption         =   "&Landscape"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   1230
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

'Private objDefPrinter As cSetDfltPrinter

Private mfDenyCollate As Boolean
Public Property Let DenyCollate(pfNewValue As Boolean)
  mfDenyCollate = pfNewValue
    
  If mfDenyCollate Then
    chkCollate.Value = vbUnchecked
  End If

  chkCollate.Enabled = Not mfDenyCollate

End Property

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

Private Sub cboPrinter_Click()

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmPrintOptions.cboPrinter_Click()"

'TM20020829 Fault 4356
'  Dim objPrinter As Printer
'
'  If Not mblnLoading Then
'    ' Set the printer to be what they've selected
'    For Each objPrinter In Printers
'      If objPrinter.DeviceName = cboPrinter.Text Then
'        objDefPrinter.SetPrinterAsDefault cboPrinter.Text
'        Set Printer = objPrinter
'        Exit For
'      End If
'    Next
'
'    ' Flag to reset this printer
'    gblnResetPrinterDefaultBack = True
'  End If
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

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

  Dim objDefPrinter As cSetDfltPrinter
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmPrintOptions.cmdPrint_Click()"

  mblnCancelled = False
  mblnPortrait = (optPortrait.Value = True)
  mlngCopies = ASRSpinner1.Value
  mblnCollateCopies = (chkCollate = vbChecked)
  mblnGridlines = (chkGridLines = vbChecked)
  mblnShading = (chkShading = vbChecked)
  mblnHeadingsOnEveryPage = (chkHeadingsEveryPage = vbChecked)
  
  If UI.GetSystemMeasurement = "us" Then
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
  
  '******************************************************************************
  'TM20020829 Fault 4356 - set the default printer and use the
  '                        'gblnResetPrinterDefaultBack' flag to reset to
  '                        original printer when 'objDefPrinter' is killed.
  
  If Not mblnLoading Then

    'MH20030922 Fault 6124 (Q257688)
    Set objDefPrinter = New cSetDfltPrinter
    objDefPrinter.SetPrinterAsDefault cboPrinter.Text
    Set objDefPrinter = Nothing

    ' Flag to reset this printer
    gblnResetPrinterDefaultBack = True
  End If
  
  '******************************************************************************
  
  Me.Hide
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub ASRSpinner1_Change()
  chkCollate.Enabled = ((ASRSpinner1.Value > 1) _
    And (Not mbPrintingExpression) _
    And (Not mfDenyCollate))
  
End Sub

Private Sub cmdCancel_Click()
  mblnCancelled = True
  'Me.Hide
  
  Unload Me

End Sub

Private Sub cmdSetDefault_Click()

  gstrDefaultPrinterName = cboPrinter.Text
  SavePCSetting "Printer", "DeviceName", gstrDefaultPrinterName

  '******************************************************************************
  'TM20020828 Fault 1432 - set the default printer using the APIs in the
  '                        cSetDfltPrinter class.

  Dim bDefaultPrinterSet As Boolean
  Dim objDefPrinter As cSetDfltPrinter

  Set objDefPrinter = New cSetDfltPrinter
  bDefaultPrinterSet = objDefPrinter.SetPrinterAsDefault(gstrDefaultPrinterName)
  Set objDefPrinter = Nothing

  '******************************************************************************

End Sub

Private Sub Form_Activate()
  cmdPrint.SetFocus
End Sub

Private Sub ShowOrientationPreview(pfPortrait As Boolean)
  Dim fraTempFrame As Frame
  
  For Each fraTempFrame In fraPortrait
    fraTempFrame.Visible = pfPortrait
  Next fraTempFrame
  
  For Each fraTempFrame In fraLandscape
    fraTempFrame.Visible = Not pfPortrait
  Next fraTempFrame
  
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

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmPrintOptions.Form_Load()"

  mblnLoading = True

  ' Retrieve the current default printer and set the combo
  Dim objPrinter As Printer
  Dim intCurrentPrinter As Integer
  Dim bFound As Boolean

  'Set objDefPrinter = New cSetDfltPrinter
  
  ''TM20020828 Fault 1432
  'Printer.TrackDefault = True
  'gstrDefaultPrinterName = Printer.DeviceName
  'SavePCSetting "Printer", "DeviceName", gstrDefaultPrinterName
  
  For Each objPrinter In Printers
    cboPrinter.AddItem objPrinter.DeviceName
    
    'JDM - 07/11/01 - Fault 3102 - Causes crash with some printers - why? I dunno...
    'cboPrinter.ItemData(cboPrinter.NewIndex) = objPrinter.hdc
    
    If LCase(objPrinter.DeviceName) = LCase(Printer.DeviceName) Then
      'objDefPrinter.SetPrinterAsDefault (gstrDefaultPrinterName)
      'Set Printer = objPrinter
      bFound = True
      cboPrinter.ListIndex = cboPrinter.NewIndex
    End If
  Next objPrinter


'  If bFound Then
'    cboPrinter.Text = gstrDefaultPrinterName
'  End If

  If UI.GetSystemMeasurement = "us" Then
    fraMargins.Caption = "Margins (inches) :"
    txtMargin(0).Text = "1"
    txtMargin(1).Text = "1"
    txtMargin(2).Text = "1"
    txtMargin(3).Text = "1"
  End If

  mblnLoading = False

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmPrintOptions.Form_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode)

  If UnloadMode = vbFormControlMenu Then
    mblnCancelled = True
    Cancel = True
    Me.Hide
    
    Unload Me

  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

'Private Sub Form_Unload(Cancel As Integer)
'  'Set objDefPrinter = Nothing
'End Sub

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
  gobjErrorStack.PushStack "frmPrintOptions.txtMargin_LostFocus(Index)", Array(Index)

  If UI.GetSystemMeasurement = "us" Then
  
    If IsNumeric(txtMargin(Index).Text) Then
    
      If Val(txtMargin(Index).Text) > 5 Then txtMargin(Index).Text = 5
      If Val(txtMargin(Index).Text) < 0.5 Then txtMargin(Index).Text = 0.5
      GoTo TidyUpAndExit
      
    ElseIf txtMargin(Index).Text = "" Then
      
      txtMargin(Index).Text = 0.5
      GoTo TidyUpAndExit
      
    End If
    
    With txtMargin(Index)
      COAMsgBox "Margins must be numeric values between 0.5 and 5 inches", vbInformation + vbOKOnly, "Print Options Error"
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
      COAMsgBox "Margins must be numeric values between 1 and 10 centimetres", vbInformation + vbOKOnly, "Print Options Error"
      .SetFocus
      .SelStart = 0
      .SelLength = Len(txtMargin(Index).Text)
    End With
    
  End If
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

' Sets the default settings for printing an expression
Public Sub PrintDefinition()

  mbPrintingExpression = True

  'Expressions should default to portrait
  optLand.Value = False
  optPortrait.Value = True
  
  ' Collate copies only apply to reports
  DenyCollate = True
  
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


