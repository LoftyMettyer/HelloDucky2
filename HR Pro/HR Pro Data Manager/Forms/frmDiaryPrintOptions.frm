VERSION 5.00
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Begin VB.Form frmDiaryPrintOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Diary Print"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1080
   Icon            =   "frmDiaryPrintOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3660
      TabIndex        =   2
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2340
      TabIndex        =   1
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Frame fraDateRangeDef 
      Caption         =   "Date Range :"
      Height          =   2550
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4740
      Begin VB.OptionButton optDate 
         Caption         =   "&Range"
         Height          =   285
         Index           =   1
         Left            =   195
         TabIndex        =   4
         Top             =   690
         Width           =   1050
      End
      Begin VB.TextBox txtDateExpr 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "0"
         Top             =   1890
         Width           =   2325
      End
      Begin VB.TextBox txtDateExpr 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   8
         Tag             =   "0"
         Top             =   1485
         Width           =   2325
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Cu&stom"
         Height          =   210
         Index           =   2
         Left            =   200
         TabIndex        =   5
         Top             =   1155
         Width           =   1275
      End
      Begin VB.OptionButton optDate 
         Caption         =   "C&urrent View"
         Height          =   195
         Index           =   0
         Left            =   200
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   3660
      End
      Begin VB.CommandButton cmdExprDate 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4050
         Picture         =   "frmDiaryPrintOptions.frx":000C
         TabIndex        =   6
         Top             =   1485
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.CommandButton cmdExprDate 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4050
         Picture         =   "frmDiaryPrintOptions.frx":0084
         TabIndex        =   7
         Top             =   1890
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin GTMaskDate.GTMaskDate cboManualDate 
         Height          =   315
         Index           =   0
         Left            =   1500
         TabIndex        =   13
         Top             =   675
         Width           =   1395
         _Version        =   65537
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         Enabled         =   0   'False
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
         BackColor       =   -2147483633
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
         CalSelForeColor =   -2147483643
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
      Begin GTMaskDate.GTMaskDate cboManualDate 
         Height          =   315
         Index           =   1
         Left            =   3240
         TabIndex        =   14
         Top             =   675
         Width           =   1395
         _Version        =   65537
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         Enabled         =   0   'False
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
         BackColor       =   -2147483633
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
         CalSelForeColor =   -2147483643
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
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2955
         TabIndex        =   12
         Top             =   735
         Width           =   195
      End
      Begin VB.Label lblCustomDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date :"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   11
         Top             =   1950
         Width           =   990
      End
      Begin VB.Label lblCustomDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date :"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   10
         Top             =   1545
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmDiaryPrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miDateType As ReturnPrintDateType
Private madRangeDates(1) As Date
Private malngCustomDateIDs(1) As Long
Private mbCancelled As Boolean
Private mbEnableDefault As Boolean

Public Property Let EnableDefault(ByVal pbAllow As Boolean)
   mbEnableDefault = pbAllow
End Property

Public Property Get NoRecordsMessage() As String

  Select Case miDateType
    Case RETURN_DEFAULT
    NoRecordsMessage = "There is no data in the current view to print."
    Case RETURN_MANUAL
    NoRecordsMessage = "There is no data in the selected date range to print."
    Case RETURN_CALCULATION
    NoRecordsMessage = "There is no data in the calculated date range to print."
  End Select

End Property

' Returns the start date of a manual entry
Public Property Get RangeStartDate() As Date
  If miDateType = RETURN_MANUAL Then
    RangeStartDate = madRangeDates(0)
  Else
    RangeStartDate = vbNull
  End If
End Property

' Returns the end date of a manual entry
Public Property Get RangeEndDate() As Date
  If miDateType = RETURN_MANUAL Then
    RangeEndDate = madRangeDates(1)
  Else
    RangeEndDate = vbNull
  End If
End Property

' Returns the end date expression id
Public Property Get EndDateExpressionID() As Long
  If miDateType = RETURN_CALCULATION Then
    EndDateExpressionID = malngCustomDateIDs(1)
  Else
    EndDateExpressionID = 0
  End If
End Property

' Returns the end date expression id
Public Property Get StartDateExpressionID() As Long
  If miDateType = RETURN_CALCULATION Then
    StartDateExpressionID = malngCustomDateIDs(0)
  Else
    StartDateExpressionID = 0
  End If
End Property

' Returns the date type selected
Public Property Get ReturnDateType() As ReturnPrintDateType
  ReturnDateType = miDateType
End Property

' Was form cancelled?
Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Private Sub cboManualDate_Change(Index As Integer)
  'madRangeDates(Index) = cboManualDate(Index).DateValue
End Sub

'Private Sub cboManualDate_DblClick(Index As Integer)
'  'NHRD23072003 Fault 6295
'  cboManualDate(Index).DateValue = Date
'End Sub

Private Sub cboManualDate_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  'NHRD23072003 Fault 6295
  If KeyCode = vbKeyF2 Then
    cboManualDate(Index).DateValue = Date
  End If
End Sub

Private Sub cboManualDate_LostFocus(Index As Integer)

  On Error Resume Next

  ValidateGTMaskDate cboManualDate(Index)   'MH20030509 Fault 6291
  madRangeDates(Index) = cboManualDate(Index).DateValue

End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()

  Dim iCount As Integer
  Dim bValid As Boolean
  
  ' Validate that we have valid dates
  bValid = True
  Select Case miDateType
    Case RETURN_MANUAL
      'MH20030905 Fault 6291
      If Not ValidateGTMaskDate(cboManualDate(0)) Then
        Exit Sub
      ElseIf Not IsValidDate(cboManualDate(0)) Then
        COAMsgBox "Please enter a valid date", vbExclamation
        cboManualDate(0).SetFocus
        Exit Sub
      End If
      If Not ValidateGTMaskDate(cboManualDate(1)) Then
        Exit Sub
      ElseIf Not IsValidDate(cboManualDate(1)) Then
        COAMsgBox "Please enter a valid date", vbExclamation
        cboManualDate(1).SetFocus
        Exit Sub
      End If

    Case RETURN_CALCULATION
      For iCount = 0 To 1
        If malngCustomDateIDs(iCount) = 0 Then
          bValid = False
        End If
      Next iCount
  End Select
  
  If bValid Then
    ' Only save appropriate ones - to avoid hitting the database
    SaveUserSetting "diaryprint", "rangetype", miDateType
    
    Select Case miDateType
      Case RETURN_DEFAULT
        SaveUserSetting "diaryprint", "startvalue", ""
        SaveUserSetting "diaryprint", "endvalue", ""
      Case RETURN_MANUAL
        SaveUserSetting "diaryprint", "startvalue", madRangeDates(0)
        SaveUserSetting "diaryprint", "endvalue", madRangeDates(1)
      Case RETURN_CALCULATION
        SaveUserSetting "diaryprint", "startvalue", malngCustomDateIDs(0)
        SaveUserSetting "diaryprint", "endvalue", malngCustomDateIDs(1)
    End Select
  
    mbCancelled = False
    Unload Me
  Else
    COAMsgBox "Please complete date selection", vbExclamation, Me.Caption
  End If
End Sub

Private Sub Form_Initialize()
  mbCancelled = True    'MH20030905 Fault 6508
  mbEnableDefault = True
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
  'JPD 20041118 Fault 8231
  UI.FormatGTDateControl cboManualDate(0)
  UI.FormatGTDateControl cboManualDate(1)

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub optDate_Click(Index As Integer)

  Dim bCustomDate As Boolean
  Dim bRange As Boolean
  Dim iCount As Integer

  bCustomDate = (Index = 2)
  bRange = (Index = 1)

  ' Custom settings
  For iCount = 0 To 1
    cmdExprDate(iCount).Enabled = bCustomDate
    lblCustomDate(iCount).Enabled = bCustomDate
    If Not bCustomDate Then
      txtDateExpr(iCount).Text = vbNullString
      malngCustomDateIDs(iCount) = 0
    End If
  Next iCount
  
  ' Range settings
  For iCount = 0 To 1
    If Not bRange Then
      cboManualDate(iCount).Text = vbNullString
    End If
    
    cboManualDate(iCount).Enabled = bRange
    cboManualDate(iCount).BackColor = IIf(bRange, vbWhite, vbButtonFace)
  Next iCount
  lblToDate.Enabled = bRange    'MH20030509

  miDateType = Index

End Sub

Private Sub cmdExprDate_Click(Index As Integer)
  
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression
  
  fOK = True
  
  Set objExpression = New clsExprExpression
  With objExpression
    
    fOK = .Initialise(0, malngCustomDateIDs(Index), giEXPR_RECORDINDEPENDANTCALC, giEXPRVALUE_DATE)
    If fOK Then

      Do
        .SelectExpression True
        
        'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
        ' when all we need is the expression return type (time saving measure).
        .ValidateExpression True, True
        
        fOK = (.ReturnType = giEXPRVALUE_DATE)
        
        If Not fOK Then
          COAMsgBox "This calculation does not return a date value.", vbExclamation, Me.Caption
        End If
      Loop While Not fOK

      If fOK Then
        txtDateExpr(Index).Text = .Name
        malngCustomDateIDs(Index) = .ExpressionID
      End If
    
    End If
  End With
  
  Set objExpression = Nothing

End Sub

Public Sub Initialise()

  ' Only save appropriate ones - to avoid hitting the database
  miDateType = GetUserSetting("diaryprint", "rangetype", 0)
  miDateType = IIf(Not mbEnableDefault And miDateType = RETURN_DEFAULT, RETURN_MANUAL, miDateType)
  optDate(miDateType).Value = True
  optDate(0).Enabled = mbEnableDefault
  
  Select Case miDateType
    Case RETURN_MANUAL
      madRangeDates(0) = GetUserSetting("diaryprint", "startvalue", vbNull)
      madRangeDates(1) = GetUserSetting("diaryprint", "endvalue", vbNull)
      'NHRD22072003 Fault 6294
      cboManualDate(0).Text = "" 'madRangeDates(0)
      cboManualDate(1).Text = "" 'madRangeDates(1)
      
    Case RETURN_CALCULATION
      malngCustomDateIDs(0) = GetUserSetting("diaryprint", "startvalue", 0)
      malngCustomDateIDs(1) = GetUserSetting("diaryprint", "endvalue", 0)
      txtDateExpr(0).Text = GetExpressionName(malngCustomDateIDs(0))
      txtDateExpr(1).Text = GetExpressionName(malngCustomDateIDs(1))
      
  End Select

End Sub

Private Function GetExpressionName(lngID As Long) As String

  Dim objExpr As clsExprExpression
  Dim strExpressionName As String
  Dim bOK As Boolean

  On Local Error GoTo LocalErr

  Set objExpr = New clsExprExpression
  With objExpr
    bOK = .Initialise(0, lngID, giEXPR_RECORDINDEPENDANTCALC, giEXPRVALUE_DATE)
    If bOK Then
      .ConstructExpression
      GetExpressionName = .ComponentDescription
    Else
      GetExpressionName = ""
    End If
  End With
  Set objExpr = Nothing

Exit Function

LocalErr:
  GetExpressionName = "<Unknown Expression>"
End Function


Private Function IsValidDate(dtTemp As GTMaskDate.GTMaskDate) As Boolean

  Dim dtNewDate As Variant

  On Error GoTo ExitSub
  
  IsValidDate = False
      
  If Not ValidateGTMaskDate(dtTemp) Then
    Exit Function
  End If
  
  dtNewDate = dtTemp.DateValue
  If Not IsNull(dtNewDate) Then
    If IsDate(dtNewDate) Then
      With frmDiary.mvwViewbyMonth
        
        If DateDiff("d", dtNewDate, .MinDate) <= 0 And _
           DateDiff("d", dtNewDate, .MaxDate) >= 0 Then
          IsValidDate = True
        End If
      
      End With
    End If
  End If

ExitSub:

End Function


Private Sub PrintStuff()


End Sub

