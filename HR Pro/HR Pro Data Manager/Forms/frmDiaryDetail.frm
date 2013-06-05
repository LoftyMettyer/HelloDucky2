VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Begin VB.Form frmDiaryDetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Diary Event"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1028
   Icon            =   "frmDiaryDetail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print..."
      Height          =   400
      Left            =   9480
      TabIndex        =   18
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "Notes : "
      Height          =   2565
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   9180
      Begin VB.TextBox txtNotes 
         Height          =   2040
         Left            =   195
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   315
         Width           =   8805
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   9480
      TabIndex        =   9
      Top             =   720
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   9480
      TabIndex        =   8
      Top             =   240
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Diary Event :"
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   9180
      Begin VB.OptionButton optAccess 
         Caption         =   "&Hidden"
         Height          =   195
         Index           =   2
         Left            =   6000
         TabIndex        =   6
         Top             =   1575
         Width           =   1200
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "&Read Only"
         Height          =   195
         Index           =   1
         Left            =   6000
         TabIndex        =   5
         Top             =   1200
         Width           =   1425
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "Read / &Write"
         Height          =   195
         Index           =   0
         Left            =   6000
         TabIndex        =   4
         Top             =   810
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.TextBox txtOwner 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   15
         Top             =   315
         Width           =   3000
      End
      Begin VB.CheckBox chkAlarm 
         Caption         =   "&Alarm"
         Height          =   195
         Left            =   1620
         TabIndex        =   3
         Top             =   1560
         Width           =   1050
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   1620
         MaxLength       =   255
         TabIndex        =   0
         Top             =   315
         Width           =   3000
      End
      Begin MSMask.MaskEdBox mskTime 
         Height          =   315
         Left            =   1620
         TabIndex        =   2
         Top             =   1080
         Width           =   700
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin GTMaskDate.GTMaskDate cboDate 
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   690
         Width           =   1485
         _Version        =   65537
         _ExtentX        =   2619
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
         Caption         =   "Access :"
         Height          =   195
         Index           =   5
         Left            =   5100
         TabIndex        =   16
         Top             =   810
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Owner :"
         Height          =   195
         Index           =   4
         Left            =   5100
         TabIndex        =   14
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Title :"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   11
         Top             =   365
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Time :"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   13
         Top             =   1125
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   12
         Top             =   765
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmDiaryDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngTimeStamp As Long
Private blnActivated As Boolean

Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Changed(blnChanged As Boolean)
  cmdOK.Enabled = blnChanged
End Property

Private Function CheckNull(varInput As Variant) As String
  CheckNull = IIf(IsNull(varInput), vbNullString, Trim(varInput))
End Function

Private Sub cboDate_Change()
  Me.Changed = True
End Sub

Private Sub cboDate_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF2 Then
    cboDate.DateValue = Date
  End If

End Sub

Private Sub cboDate_LostFocus()

'  If IsNull(cboDate.DateValue) And Not IsDate(cboDate.DateValue) And cboDate.Text <> "  /  /" Then
'    COAMsgBox "You have entered an invalid date.", vbOKOnly & vbExclamation, "Diary Detail"
'    cboDate.DateValue = Null
'    cboDate.SetFocus
'    Exit Sub
'  End If

  'MH20020423 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate cboDate

End Sub

Private Sub chkAlarm_Click()
  Me.Changed = True
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub


Private Sub cmdOK_Click()
  If SaveChanges Then
    Unload Me
  End If
End Sub


Private Sub Form_Activate()

  Dim lngDiaryEventID As Long
  Dim rsTables As Recordset
  Dim strSQL As String
  Dim strAccess As String
  Dim lngColumnID As Long
  Dim intCount As Integer
  Dim dtEventDate As Date

  If blnActivated Then
    Exit Sub
  End If
  blnActivated = True


  lngDiaryEventID = gobjDiary.DiaryEventID
  mskTime.Mask = "99" & UI.GetSystemTimeSeparator & "99"
  
  If lngDiaryEventID = 0 Then
    
    'New record therefore set up defaults
    cboDate.DateValue = gobjDiary.DateSelected
    If Format(cboDate.Text, DateFormat) = Format(Date, DateFormat) Then
      mskTime.Text = Format(Time, "hh:mm")  'Default to current time if today
      'mskTime.Text = DiaryFormat(Time, "hh:mm")
    Else
      mskTime.Text = "00" & UI.GetSystemTimeSeparator & "00"
    End If

    txtTitle.Text = vbNullString
    txtNotes.Text = vbNullString
    chkAlarm.Value = vbChecked
    txtOwner.Text = gsUserName
    optAccess(0) = True   'Default to Read/Write for new entry
    mlngTimeStamp = 0

    txtTitle.SetFocus

  Else
  
    Set rsTables = gobjDiary.GetCurrentRecord
    
    With rsTables
    
      If .BOF And .EOF Then
        Exit Sub    'No Records
      End If

      dtEventDate = IIf(Not IsNull(.Fields("EventDate")), CDate(.Fields("EventDate")), Date)
      cboDate.DateValue = dtEventDate

      mskTime.Text = Format(dtEventDate, "hh:nn")
      'mskTime.Text = DiaryFormat(dtEventDate, "hh:nn")

      txtTitle.Text = CheckNull(.Fields("EventTitle"))
      chkAlarm.Value = Abs(CInt(.Fields("Alarm")) <> 0)
      txtOwner.Text = CheckNull(.Fields("UserName"))
      lngColumnID = IIf(Not IsNull(.Fields("ColumnID")), Val(.Fields("ColumnID")), 0)
      mlngTimeStamp = IIf(Not IsNull(.Fields("intTimeStamp")), Val(.Fields("intTimeStamp")), 0)

      'MH20001013 Fault 829 Show Column details in notes for system events
      If lngColumnID > 0 And Not IsNull(.Fields("ColumnValue")) Then
        txtNotes.Text = Replace(datGeneral.GetColumnName(lngColumnID), "_", " ") & _
                        " : " & _
                        Format(.Fields("ColumnValue"), DateFormat)
      Else
        txtNotes.Text = CheckNull(.Fields("EventNotes").Value)
      End If

      Select Case .Fields("Access")
      Case "RO"
        optAccess(1) = True
      Case "HD"
        optAccess(2) = True
      Case Else
        optAccess(0) = True
      End Select
    
    End With


    If gobjDiary.WriteAccess = False Then
      'Viewing somebody else's read only diary event
      'or a system diary entry
      'Call DisableControls(Me)
      'txtNotes.Enabled = True
      'txtNotes.Locked = True
      'txtNotes.ForeColor = vbApplicationWorkspace

      ControlsDisableAll Me

      cmdPrint.Enabled = True
      cmdCancel.Enabled = True
      cmdCancel.SetFocus

      'Ensure that this flag is shown correctly
      '(DisableControls sub changes opt buttons!)
      optAccess(1).Value = True

      'Don't allow access to alarm checkbox if system event in the future
      If gobjDiary.AlarmAccess Then
        chkAlarm.Enabled = True
        chkAlarm.ForeColor = vbWindowText
      End If

    ElseIf gobjDiary.EventOwner = False Then
      'If viewing somebodys diary event then do not
      'allow user to change access
      For intCount = 0 To 2
        optAccess(intCount).Enabled = False
      Next

    End If
  
  End If

  Me.Changed = False

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
  'Call SetDateComboFormat(cboDate)
  blnActivated = False

  'JPD 20041118 Fault 8231
  UI.FormatGTDateControl cboDate
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  If Me.Changed = True And Me.Visible Then
    
    Select Case COAMsgBox("You have changed the current diary event. Save changes ?", vbQuestion + vbYesNoCancel, "Diary Event")
    Case vbYes
      If Not SaveChanges Then
        Cancel = True
      End If
      Exit Sub
    Case vbCancel
      Cancel = True
      Exit Sub
    End Select
  End If

  Set frmDiaryDetail = Nothing

End Sub


'Private Sub DisableControls(Frm As Form)
'
'  'Disable all controls on the form
'  'apart from Labels, Frames and TabStrips
'
'  On Local Error Resume Next
'
'  Dim ctlTemp As Control
'  For Each ctlTemp In Frm.Controls
'    If Not (TypeOf ctlTemp Is Label) And _
'       Not (TypeOf ctlTemp Is Frame) And _
'       Not (TypeOf ctlTemp Is TabStrip) Then
'          ctlTemp.Enabled = False
'          ctlTemp.BackColor = vbButtonFace
'    End If
'  Next
'  Frm.cmdCancel.Enabled = True
'
'End Sub

'Private Sub mskTime_Change()
'  Me.Changed = True
'End Sub

Private Sub mskTime_KeyUp(KeyCode As Integer, Shift As Integer)
  Me.Changed = True
End Sub

Private Sub mskTime_GotFocus()
  With mskTime
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

'Private Sub mskTime_LostFocus()
'
'  Dim strTime As String
'
'  strTime = FormatTime(mskTime.Text)
'  If strTime <> vbNullString Then
'    mskTime.Text = strTime
'  End If
'
'End Sub

Private Sub optAccess_Click(Index As Integer)
  Me.Changed = True
End Sub

Private Sub txtNotes_Change()
  Me.Changed = True
End Sub

Private Sub txtNotes_GotFocus()
  With txtNotes
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  cmdOK.Default = False
End Sub

Private Sub txtNotes_LostFocus()
  cmdOK.Default = True
End Sub
Private Sub txtTitle_Change()
  Me.Changed = True
End Sub

Private Sub txtTitle_GotFocus()
  With txtTitle
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub



Private Sub cmdPrint_Click()

  Dim objPrintDef As clsPrintDef

  Set objPrintDef = New HRProDataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      If .PrintStart(False) Then
        .PrintHeader "Diary Event : " & txtTitle.Text
        .PrintNormal "Event Date : " & Format(cboDate.DateValue, DateFormat)
        .PrintNormal "Event Time : " & mskTime.Text
        .PrintNormal "Alarmed : " & IIf(chkAlarm.Value, "Yes", "No")
        .PrintNormal
        
        .PrintNormal "Owner : " & txtOwner.Text
        
        If optAccess(0).Value = True Then
          .PrintNormal "Access : Read / Write"
        ElseIf optAccess(1).Value = True Then
          .PrintNormal "Access : Read Only"
        ElseIf optAccess(2).Value = True Then
          .PrintNormal "Access : Hidden"
        End If
        .PrintNormal
        
        '---------
        
        .PrintTitle "Notes"
        .PrintNonBold Replace(txtNotes, Chr(10), "")
        .PrintEnd
        .PrintConfirm "Diary Event", Me.Caption
      End If
  
    End With
  
  End If
    
  Set objPrintDef = Nothing

Exit Sub

LocalErr:
  COAMsgBox "Printing Diary Event Failed", vbExclamation

End Sub


Private Function IsValidDate(dtInput As Variant) As Boolean

  On Local Error GoTo ExitSub

  IsValidDate = False
  If IsDate(dtInput) Then
    'Next line can raise an error (but will be trapped)
    'This will check if date is within monthview.mindate and maxdate
    gobjDiary.DateSelected = dtInput
    IsValidDate = True
  End If

ExitSub:

End Function


Private Function IsInCurrentFilter() As Boolean

  Dim blnPast As Boolean
  
  IsInCurrentFilter = True
  
  With gobjDiary
  
    If .FilterEventType = 1 Then  'Viewing System Events
      IsInCurrentFilter = False
      Exit Function
    End If

    If (.FilterAlarmStatus = 1 And chkAlarm.Value = vbUnchecked) Or _
       (.FilterAlarmStatus = 2 And chkAlarm.Value = vbChecked) Then
      IsInCurrentFilter = False
      Exit Function
    End If

    blnPast = (DateDiff("n", CDate(Format(cboDate.DateValue, DateFormat) & " " & mskTime.Text), Now) > 0)
    If (.FilterPastPresent = 1 And Not blnPast) Or _
       (.FilterPastPresent = 2 And blnPast) Then
      IsInCurrentFilter = False
      Exit Function
    End If

  End With

End Function


Private Function SaveChanges() As Boolean

  Dim lngDiaryEventID As Long
  Dim strSQL As String
  Dim strAccess As String
  Dim intCount As Integer
  
  Dim blnContinueSave As Boolean
  Dim blnSaveAsNew As Boolean

  SaveChanges = False

  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  'Double check valid date in case lost focus is missed
  '(by pressing enter for default button lostfocus is not run)
  '''If IsDate(cboDate.DateValue) = False Then
  ''If IsValidDate(cboDate.DateValue) = False Then
  
  If ValidateGTMaskDate(cboDate) = False Then
    Exit Function

  ElseIf IsValidDate(cboDate.DateValue) = False Then
    COAMsgBox "Please enter a valid date for this event", vbExclamation
    cboDate.SetFocus
    Exit Function

  ElseIf FormatTime(mskTime) = vbNullString Then
    COAMsgBox "Please enter a valid time for this event", vbExclamation
    mskTime.SetFocus
    Exit Function

  ElseIf Trim$(txtTitle) = vbNullString Then
    COAMsgBox "Please enter a title for this event", vbExclamation
    txtTitle.SetFocus
    Exit Function

  End If

  gobjDiary.DateSelected = cboDate.DateValue


  Call UtilityDefAmended("ASRSysDiaryEvents", "DiaryEventsID", gobjDiary.DiaryEventID, mlngTimeStamp, _
        blnContinueSave, blnSaveAsNew, "diary event")
  
  If blnContinueSave = False Then
    Exit Function
  ElseIf blnSaveAsNew Then
    txtOwner = gsUserName
    For intCount = 0 To 2
      optAccess(0).Enabled = True
    Next
    gobjDiary.DiaryEventID = 0
  End If
  
  
  'For intCount = 0 To 2
  '  If optAccess(intCount) = True Then
  '    intAccess = intCount
  '    Exit For
  '  End If
  'Next

  strAccess = "RW"
  If optAccess(1) = True Then
    strAccess = "RO"
  ElseIf optAccess(2) = True Then
    strAccess = "HD"
  End If
  
  'TM20020520 Fault 2267 - have now moved the formatting of the string to the PutRecord
  'function of the clsDiary.
'  Dim sNotestext As String
'  sNotestext = Replace(txtNotes.Text, "'", "''")
'  sNotestext = Left(sNotestext, 7000)
  
  Call gobjDiary.PutRecord(txtTitle.Text, cboDate.DateValue, _
    mskTime.Text, txtNotes.Text, chkAlarm.Value, strAccess, 0)
  
'  Call gobjDiary.PutRecord(txtTitle.Text, cboDate.DateValue, _
'    mskTime.Text, sNotestext, chkAlarm.Value, strAccess, 0)

  If Not IsInCurrentFilter Then
    If Not gobjDiary.ViewingAlarms Then
      COAMsgBox "The event saved does not satisfy the current filter.", vbInformation, App.ProductName
    End If
    gobjDiary.DiaryEventID = 0
  End If

  gobjDiary.DateSelected = cboDate.DateValue
  gobjDiary.RefreshDiaryData
  Me.Changed = False

  SaveChanges = True

End Function

