VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmDiaryDuplicate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Repeat/Copy Diary Entry"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
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
   Icon            =   "frmDiaryDuplicate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCopy 
      Caption         =   "Copy To :"
      Height          =   1400
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
      Begin GTMaskDate.GTMaskDate cboCopyDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
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
      Begin MSMask.MaskEdBox mskCopyTime 
         Height          =   315
         Left            =   1200
         TabIndex        =   17
         Top             =   800
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Time :"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   855
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame fraRepeat 
      Caption         =   "Repeat :"
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkWeekends 
         Caption         =   "Include &weekends"
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   1300
         Width           =   2010
      End
      Begin VB.ComboBox cboIntervalPeriod 
         Height          =   315
         ItemData        =   "frmDiaryDuplicate.frx":000C
         Left            =   2280
         List            =   "frmDiaryDuplicate.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   800
         Width           =   1185
      End
      Begin COASpinner.COA_Spinner spnIntervalAmount 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   800
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaximumValue    =   99
         MinimumValue    =   1
         Text            =   "1"
      End
      Begin GTMaskDate.GTMaskDate cboUntilDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   300
         Width           =   1530
         _Version        =   65537
         _ExtentX        =   2699
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
         BackStyle       =   0  'Transparent
         Caption         =   "Until :"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Interval :"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   860
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3000
      TabIndex        =   16
      Top             =   3480
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1725
      TabIndex        =   15
      Top             =   3480
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtTitle 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin VB.OptionButton optRepeat 
         Caption         =   "&Repeat Entry"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.OptionButton optCopy 
         Caption         =   "Copy &Entry"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmDiaryDuplicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsTables As Recordset

Private Sub cboCopyDate_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF2 Then
    cboCopyDate.DateValue = Date
  End If

End Sub

Private Sub cboCopyDate_LostFocus()

'  If IsNull(cboCopyDate.DateValue) And Not _
'     IsDate(cboCopyDate.DateValue) And _
'     cboCopyDate.Text <> "  /  /" Then
'
'     COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     cboCopyDate.DateValue = Null
'     cboCopyDate.SetFocus
'     Exit Sub
'  End If

  'MH20020423 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate cboCopyDate
  
End Sub


Private Sub cboUntilDate_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF2 Then
    cboUntilDate.DateValue = Date
  End If

End Sub

Private Sub cboUntilDate_LostFocus()

'  If IsNull(cboUntilDate.DateValue) And Not _
'     IsDate(cboUntilDate.DateValue) And _
'     cboUntilDate.Text <> "  /  /" Then
'
'     COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     cboUntilDate.DateValue = Null
'     cboUntilDate.SetFocus
'     Exit Sub
'  End If

  'MH20020423 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate cboUntilDate

End Sub


Private Sub cmdCancel_Click()
  Unload Me
End Sub


Private Sub cmdOK_Click()

  Dim strTitle As String
  Dim dtDate As Date
  Dim strTime As String
  Dim strNotes As String
  Dim intAlarm As Integer
  Dim strAccess As String
  Dim lngOriginalID As Long
  Dim intCount As Integer
  Dim strIntervalPeriod As String
  Dim lngDateDiff As Long
  
  With rsTables
    strTitle = .Fields("EventTitle")
    'dtDate = Format(.Fields("EventDate"), "Short Date")
    'strTime = .Fields("EventTime")
    dtDate = CDate(.Fields("EventDate").Value)
    strTime = DiaryFormat(.Fields("EventDate"), "hh:nn")
    strNotes = .Fields("EventNotes")
    intAlarm = .Fields("Alarm")
    strAccess = .Fields("Access")

    'Store the ID of the original source record,
    'if zero then the current record is the source
    lngOriginalID = .Fields("CopiedFromID")
    If CLng(lngOriginalID) = 0 Then
      lngOriginalID = gobjDiary.DiaryEventID
    End If
  End With
  'Reset the diary pointer so that new
  'record(s) with be made
  gobjDiary.DiaryEventID = 0

  Select Case cboIntervalPeriod.ListIndex
  Case 1
    strIntervalPeriod = "ww"    'weeks
  Case 2
    strIntervalPeriod = "m"     'months
  Case 3
    strIntervalPeriod = "yyyy"  'years
  Case Else
    strIntervalPeriod = "d"     'days
  End Select
       
  'NHRD29102004 Fault 9248
  If cboUntilDate.DateValue < (DateAdd(strIntervalPeriod, spnIntervalAmount.Value, dtDate)) Then
    COAMsgBox "Interval Period is greater than the Until Date." & vbCrLf & vbCrLf & "Decrease the Interval Period or Increase the Until Date.", vbExclamation, "Interval Period"
        gobjProgress.CloseProgress
        Screen.MousePointer = vbDefault
    Exit Sub
  End If

  If optCopy.Value = True Then

    'If IsDate(cboCopyDate.Text) = False Then
    If ValidateGTMaskDate(cboCopyDate) = False Then
      Exit Sub
    ElseIf IsValidDate(cboCopyDate.DateValue) = False Then
      COAMsgBox "Please enter a valid date", vbExclamation
      cboCopyDate.SetFocus
      Exit Sub
    ElseIf FormatTime(Me.mskCopyTime) = vbNullString Then
      COAMsgBox "Please enter a valid time", vbExclamation
      mskCopyTime.SetFocus
      Exit Sub
    End If
    
    
    'dtDate = CDate(cboCopyDate.Text)
    dtDate = cboCopyDate.DateValue
    Call gobjDiary.PutRecord(strTitle, dtDate, DiaryFormat(mskCopyTime.Text, "hh:nn"), _
                      strNotes, intAlarm, strAccess, lngOriginalID)
  
    intCount = 1
  
  Else

    'If IsDate(cboUntilDate.Text) = False Then
    If ValidateGTMaskDate(cboUntilDate) = False Then
      Exit Sub
    ElseIf IsValidDate(cboUntilDate.DateValue) = False Then
      COAMsgBox "Please enter a valid date", vbExclamation
      cboUntilDate.SetFocus
      Exit Sub
    End If
    
    'MH20001025 fault 1211
    'Hide form to stop any buttons being pressed etc. !
    Me.Hide
    frmDiary.Enabled = False
    Screen.MousePointer = vbHourglass
  
    With gobjProgress
      '.AviFile = App.Path & "\videos\diary.avi"
      .AVI = dbDiary
      .MainCaption = "Diary"
      .NumberOfBars = 1
      .Caption = "Duplicate Diary Entries"
      .Time = False
      .Cancel = True
      .OpenProgress
      .Bar1Caption = "Duplicating : " & Trim(strTitle)
    End With
  
'    Select Case cboIntervalPeriod.ListIndex
'    Case 1
'      strIntervalPeriod = "ww"    'weeks
'    Case 2
'      strIntervalPeriod = "m"     'months
'    Case 3
'      strIntervalPeriod = "yyyy"  'years
'    Case Else
'      strIntervalPeriod = "d"     'days
'    End Select
  
    lngDateDiff = DateDiff(strIntervalPeriod, dtDate, cboUntilDate.DateValue) + 1
    gobjProgress.Bar1MaxValue = lngDateDiff

    dtDate = DateAdd(strIntervalPeriod, spnIntervalAmount.Value, dtDate)
    
    Do While lngDateDiff >= 0 And Not gobjProgress.Cancelled
  
      If (Weekday(dtDate) <> vbSaturday And Weekday(dtDate) <> vbSunday) Or _
          chkWeekends.Value = vbChecked Then
        Call gobjDiary.PutRecord(strTitle, dtDate, strTime, _
                  strNotes, intAlarm, strAccess, lngOriginalID)
        
        intCount = intCount + 1
      End If
  
      dtDate = DateAdd(strIntervalPeriod, spnIntervalAmount.Value, dtDate)
      
      lngDateDiff = DateDiff("d", dtDate, cboUntilDate.DateValue)
      gobjProgress.UpdateProgress False
      DoEvents
  
    Loop
  
    gobjProgress.CloseProgress
    Screen.MousePointer = vbDefault

  End If

  frmDiary.Enabled = True
  COAMsgBox "Diary entry duplicated " & CStr(intCount) & " time(s).", vbInformation
  Unload Me

End Sub

Private Sub Form_Load()
  
'  Call SetDateComboFormat(cboCopyDate)
'  Call SetDateComboFormat(cboUntilDate)
  
  cboIntervalPeriod.ListIndex = 0
  'Call optCopy_Click
  Call optRepeat_Click

  'JPD 20041118 Fault 8231
  UI.FormatGTDateControl cboCopyDate
  UI.FormatGTDateControl cboUntilDate
  
  Set rsTables = gobjDiary.GetCurrentRecord
  txtTitle.Text = rsTables.Fields("EventTitle")
  'cboCopyDate.Text = CDate(rsTables.Fields("EventDate"))
  'mskCopyTime.Text = rsTables.Fields("EventTime")
  cboCopyDate.DateValue = CDate(rsTables.Fields("EventDate"))
  mskCopyTime.Mask = "99" & UI.GetSystemTimeSeparator & "99"
  mskCopyTime.Text = Format(rsTables.Fields("EventDate"), "hh:nn")

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


'Private Sub mskCopyTime_LostFocus()
'
'  Dim strTime As String
'
'  strTime = FormatTime(mskCopyTime.Text)
'  'If strTime <> vbNullString Then
'  '  mskCopyTime.Text = strTime
'  'End If
'
'End Sub

Private Sub optCopy_Click()
  fraCopy.Visible = True
  fraRepeat.Visible = False
End Sub

Private Sub optRepeat_Click()
  fraCopy.Visible = False
  fraRepeat.Visible = True
End Sub


Private Function IsValidDate(dtNewDate As Variant) As Boolean

  On Error GoTo ExitSub
  
  IsValidDate = False
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

