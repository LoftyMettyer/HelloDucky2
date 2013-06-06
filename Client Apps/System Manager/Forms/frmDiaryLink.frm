VERSION 5.00
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmDiaryLink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Diary Link"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5014
   Icon            =   "frmDiaryLink.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2200
      Index           =   1
      Left            =   100
      TabIndex        =   8
      Top             =   1800
      Width           =   6300
      Begin VB.ComboBox cboDateLinkDirection 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmDiaryLink.frx":000C
         Left            =   4050
         List            =   "frmDiaryLink.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   1300
      End
      Begin VB.ComboBox cboDateLinkOffsetPeriod 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmDiaryLink.frx":0029
         Left            =   2640
         List            =   "frmDiaryLink.frx":0039
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   300
         Width           =   1300
      End
      Begin VB.CheckBox chkCheckLeavingDate 
         Caption         =   "Do not diary events after the employee &leaving date"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   1060
         Value           =   1  'Checked
         Width           =   4800
      End
      Begin VB.CheckBox chkDiaryReminder 
         Caption         =   "&Alarmed Event"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   760
         Value           =   1  'Checked
         Width           =   2250
      End
      Begin COASpinner.COA_Spinner spnDateLinkOffset 
         Height          =   315
         Left            =   1740
         TabIndex        =   10
         Top             =   300
         Width           =   795
         _ExtentX        =   1402
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
         MaximumValue    =   999
         Text            =   "0"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Offset :"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Link Details :"
      Height          =   1600
      Index           =   0
      Left            =   100
      TabIndex        =   0
      Top             =   120
      Width           =   6300
      Begin VB.TextBox txtDiaryComment 
         Height          =   315
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   2
         Top             =   300
         Width           =   4395
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "..."
         Height          =   315
         Left            =   5835
         TabIndex        =   5
         Top             =   700
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox txtFilter 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   700
         Width           =   4100
      End
      Begin GTMaskDate.GTMaskDate cboEffectiveDate 
         Height          =   315
         Left            =   1740
         TabIndex        =   7
         Top             =   1080
         Width           =   1500
         _Version        =   65537
         _ExtentX        =   2646
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
         Text            =   "  /  /"
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
            Name            =   "MS Sans Serif"
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
         Caption         =   "Effective Date :"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   1155
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Filter :"
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   760
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3900
      TabIndex        =   15
      Top             =   4150
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5200
      TabIndex        =   16
      Top             =   4150
      Width           =   1200
   End
End
Attribute VB_Name = "frmDiaryLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Properties.
Private mlngTableID As Long
Private gfCancelled As Boolean
Private gsDiaryComment As String
Private giDiaryOffset As Integer
Private giDiaryPeriod As TimePeriods
Private gfDiaryReminder As Boolean
Private glDiaryFilterID As Long
Private gdtDiaryEffectiveDate As Date
Private gfCheckLeavingDate As Boolean

' Flag to see if any changes have been made by the user
Private mblnChanged As Boolean

' Form handling variables.
Private mfLoading As Boolean

Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Let TableID(ByVal lngNewValue As Long)
  mlngTableID = lngNewValue
End Property


Private Sub cboDateLinkOffsetPeriod_Click()

  If Not mfLoading Then mblnChanged = True

End Sub

Private Sub cboDateLinkDirection_Click()
  If Not mfLoading Then mblnChanged = True
  
End Sub

Private Sub cboEffectiveDate_Change()

  If Not mfLoading Then mblnChanged = True

End Sub

Private Sub cboEffectiveDate_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF2 Then
    cboEffectiveDate.DateValue = Date
  End If

End Sub

Private Sub cboEffectiveDate_LostFocus()

  If IsNull(cboEffectiveDate.DateValue) And Len(Trim(Replace(cboEffectiveDate.Text, UI.GetSystemDateSeparator, ""))) = 0 Then
     MsgBox "You must enter a date.", vbOKOnly + vbExclamation, App.Title
     cboEffectiveDate.SetFocus
     Exit Sub
  End If
  
'  If IsNull(cboEffectiveDate.DateValue) And Not _
'     IsDate(cboEffectiveDate.DateValue) And _
'     cboEffectiveDate.Text <> "  /  /" Then
'
'     MsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, App.Title
'     cboEffectiveDate.DateValue = Null
'     cboEffectiveDate.SetFocus
'     Exit Sub
'  End If

  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  ValidateGTMaskDate cboEffectiveDate

End Sub

Private Sub chkCheckLeavingDate_Click()
  If Not mfLoading Then mblnChanged = True
End Sub

Private Sub chkDiaryReminder_Click()
  If Not mfLoading Then mblnChanged = True
End Sub

Private Sub cmdCancel_Click()
  ' Mark the form as being cancelled.
  gfCancelled = True

  ' Unload the form.
  UnLoad Me

End Sub

Private Sub cmdFilter_Click()

  ' Display the 'Where Clause' expression selection form.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim objExpr As CExpression

  fOK = True

  ' Instantiate an expression object.
  Set objExpr = New CExpression

  With objExpr
    ' Set the properties of the expression object.
    .Initialise mlngTableID, txtFilter.Tag, giEXPR_LINKFILTER, giEXPRVALUE_LOGIC

    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      txtFilter.Tag = .ExpressionID
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
    Else
      ' Check in case the original expression has been deleted.
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      If txtFilter.Text = vbNullString Then
        txtFilter.Tag = 0
      End If
    End If
  
  End With


TidyUpAndExit:
  Set objExpr = Nothing
  If Not fOK Then
    MsgBox "Error changing filter ID.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Private Sub cmdOk_Click()

  'MH20020424 Fault 3760
  '(Avoid changing 01/13/2002 to 13/01/2002)
  'Double check valid date in case lost focus is missed
  '(by pressing enter for default button lostfocus is not run)
  If ValidateGTMaskDate(cboEffectiveDate) = False Then
    Exit Sub
  End If
  
  ' Validate the diary comment.
  If Len(Trim(txtDiaryComment.Text)) = 0 Then
    MsgBox "A comment must be entered.", vbOKOnly + vbExclamation, Application.Name
    txtDiaryComment.SetFocus
    Exit Sub
  End If
  
  If IsNull(cboEffectiveDate.DateValue) Then
    MsgBox "You must enter a date.", vbOKOnly + vbExclamation, App.Title
    cboEffectiveDate.SetFocus
    Exit Sub
  End If
  
  gsDiaryComment = Trim(txtDiaryComment.Text)
  
  'giDiaryOffset = Val(spnDateLinkOffset.Text)
  giDiaryOffset = spnDateLinkOffset.value * IIf(cboDateLinkDirection.ListIndex = 1, 1, -1)
  
  If cboDateLinkOffsetPeriod.ListIndex >= 0 Then
    giDiaryPeriod = cboDateLinkOffsetPeriod.ItemData(cboDateLinkOffsetPeriod.ListIndex)
  Else
    giDiaryPeriod = 0
  End If
  gfDiaryReminder = (chkDiaryReminder.value = vbChecked)
  glDiaryFilterID = txtFilter.Tag
  gdtDiaryEffectiveDate = cboEffectiveDate.DateValue
  gfCheckLeavingDate = (chkCheckLeavingDate.value = vbChecked)

  ' Mark the form as not being cancelled.
  gfCancelled = False
  'mlngTableID.ChangedDiaryLink = True

  Application.ChangedDiaryLink = True

  ' Unload the form.
  UnLoad Me
  
End Sub

Private Sub Form_Activate()

  mfLoading = False
  mblnChanged = False
  
  chkCheckLeavingDate.Enabled = EnableCheckLeavingDate
  If chkCheckLeavingDate.Enabled = False Then
    chkCheckLeavingDate.value = vbUnchecked
  End If

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
  
  mfLoading = True
  
  ' Populate the diary period combo.
  cboDateLinkOffsetPeriod_Initialize

  'JPD 20041115 Fault 8970
  UI.FormatGTDateControl cboEffectiveDate
  
  If Application.AccessMode <> accFull And _
     Application.AccessMode <> accSupportMode Then
        ControlsDisableAll Me
        cmdFilter.Enabled = True
  End If

  ' Position the form.
  UI.frmAtCenterOfParent Me, frmSysMgr

  mfLoading = False

End Sub

Private Sub cboDateLinkOffsetPeriod_refresh()
  ' Refresh the diary period combo.
  Dim iLoop As Integer
  
  With cboDateLinkOffsetPeriod
    
    If spnDateLinkOffset.value > 0 Then
    
      ' Select the current period.
      For iLoop = 0 To .ListCount - 1
        If .ItemData(iLoop) = giDiaryPeriod Then
          .ListIndex = iLoop
          Exit Sub
        End If
      Next iLoop
  
    End If
  
    ' The current period is not in the combo so
    ' select the top combo item.
    '.ListIndex = 0
  
  End With

End Sub

Private Sub cboDateLinkOffsetPeriod_Initialize()
  ' Populate the diary period combo.
     
  ' Add an item for each period.
  cboDateLinkOffsetPeriod.Clear
  AddItemToComboBox cboDateLinkOffsetPeriod, "Day(s)", iTimePeriodDays
  AddItemToComboBox cboDateLinkOffsetPeriod, "Week(s)", iTimePeriodWeeks
  AddItemToComboBox cboDateLinkOffsetPeriod, "Month(s)", iTimePeriodMonths
  AddItemToComboBox cboDateLinkOffsetPeriod, "Year(s)", iTimePeriodYears
  

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  Dim pintAnswer As Integer
  
  If UnloadMode = vbFormControlMenu Then
    gfCancelled = True
  End If
  
  If gfCancelled = True Then
    
    If mblnChanged = True And cmdOK.Enabled Then
      
      pintAnswer = MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
        
      If pintAnswer = vbYes Then
        cmdOk_Click
      ElseIf pintAnswer = vbCancel Then
        Cancel = True
        Exit Sub
      End If
    
    End If
  
  End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub txtDiaryComment_Change()
  
  If Not mfLoading Then mblnChanged = True

End Sub

Private Sub txtDiaryComment_GotFocus()
  ' Select the whole string.
  UI.txtSelText

End Sub

Public Property Get DiaryPeriod() As TimePeriods
  ' Return the diary period.
  DiaryPeriod = giDiaryPeriod
  
End Property

Public Property Let DiaryPeriod(ByVal piNewValue As TimePeriods)
  ' Update the global variable.
  giDiaryPeriod = piNewValue
  
  ' Update the Diary Period combobox.
  cboDateLinkOffsetPeriod_refresh

End Property

Public Property Get DiaryComment() As String
  ' Return the diary comment.
  DiaryComment = gsDiaryComment
  
End Property

Public Property Let DiaryComment(ByVal psNewValue As String)
  ' Update the global variable.
  gsDiaryComment = psNewValue

  ' Update the Diary Comment textbox.
  txtDiaryComment.Text = gsDiaryComment

End Property

Public Property Get DiaryOffset() As Integer
  ' Return the diary offset.
  DiaryOffset = giDiaryOffset
  
End Property

Public Property Let DiaryOffset(ByVal piNewValue As Integer)
  ' Update the Diary reminder checkbox.
  giDiaryOffset = piNewValue
  
  ' Update the Diary Offset spinner.
  If Abs(giDiaryOffset) > 0 Then
    cboDateLinkDirection.ListIndex = IIf(giDiaryOffset < 0, 0, 1)
  Else
    cboDateLinkDirection.ListIndex = -1
  End If
  spnDateLinkOffset.value = Abs(giDiaryOffset)
  'SetComboItem cboDateLinkOffsetPeriod, .DatePeriod

End Property


Public Property Get DiaryReminder() As Boolean
  ' Return the diary reminder flag.
  DiaryReminder = gfDiaryReminder
  
End Property

Public Property Let DiaryReminder(ByVal pfNewValue As Boolean)
  ' Update the global variable.
  gfDiaryReminder = pfNewValue
  
  ' Update the Diary reminder checkbox.
  chkDiaryReminder.value = IIf(gfDiaryReminder, 1, 0)

End Property


Public Property Get CheckLeavingDate() As Boolean
  CheckLeavingDate = gfCheckLeavingDate
End Property

Public Property Let CheckLeavingDate(ByVal blnNewValue As Boolean)
  gfCheckLeavingDate = blnNewValue
  chkCheckLeavingDate.value = IIf(gfCheckLeavingDate, vbChecked, vbUnchecked)
End Property


Public Property Get Cancelled() As Boolean
  ' Return the Cancelled flag.
  Cancelled = gfCancelled
  
End Property


Public Property Get FilterID() As Long
  FilterID = glDiaryFilterID
End Property

Public Property Let FilterID(ByVal lngNewValue As Long)
  glDiaryFilterID = lngNewValue
  
  txtFilter.Tag = glDiaryFilterID
  txtFilter.Text = GetExpressionName(glDiaryFilterID)
End Property

Public Property Get EffectiveDate() As Date
  EffectiveDate = gdtDiaryEffectiveDate
End Property

Public Property Let EffectiveDate(ByVal dtNewValue As Date)
  gdtDiaryEffectiveDate = dtNewValue
  
  cboEffectiveDate.DateValue = gdtDiaryEffectiveDate
End Property


Private Sub txtFilter_Change()

  If Not mfLoading Then mblnChanged = True

End Sub


Private Function EnableCheckLeavingDate() As Boolean

  Dim lngPersonnelTableID As Long
  Dim lngLeavingDateID As Long
  Dim strSQL As String
  Dim blnChildOfPers As Boolean

  
  EnableCheckLeavingDate = False

  'MH20011025 Fault 3034
  If Application.AccessMode <> accFull And _
     Application.AccessMode <> accSupportMode Then
        Exit Function
  End If
  
  ' Check if Leaving Date column in module setup
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Get the Personnel table ID.
    .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
    If .NoMatch Then
      lngPersonnelTableID = 0
    Else
      lngPersonnelTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    If lngPersonnelTableID = 0 Then
      Exit Function
    End If


    ' Get the Leaving Date column ID.
    .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LEAVINGDATE
    If .NoMatch Then
      Exit Function
    End If
    lngLeavingDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)

  End With


  'If not personnel table then check if child of personnel
  If mlngTableID <> lngPersonnelTableID Then

    With recTabEdit
      .Index = "idxName"
      
      If Not (.BOF And .EOF) Then
        .MoveFirst
      End If
      
      ' Add all child tables to the listbox.
      ' NB. Do not add deleted tables, and do not add the listbox's base table.
      Do While Not .EOF()
        If (Not !Deleted) And _
          (!TableID <> lngPersonnelTableID) And _
          (!TableType = iTabChild) Then

          recRelEdit.Index = "idxParentID"
          recRelEdit.Seek "=", lngPersonnelTableID, mlngTableID
          blnChildOfPers = (Not recRelEdit.NoMatch)

        End If
        
        .MoveNext
      Loop

    End With

  End If

  EnableCheckLeavingDate = ((mlngTableID = lngPersonnelTableID Or blnChildOfPers) And lngLeavingDateID > 0)

End Function


Private Sub spnDateLinkOffset_Change()
  
  Dim blnOffset As Boolean
  
  blnOffset = (spnDateLinkOffset.value > 0)
  
  cboDateLinkOffsetPeriod.Enabled = blnOffset
  cboDateLinkOffsetPeriod.BackColor = IIf(blnOffset, vbWindowBackground, vbButtonFace)
  
  cboDateLinkDirection.Enabled = blnOffset
  cboDateLinkDirection.BackColor = IIf(blnOffset, vbWindowBackground, vbButtonFace)
  
  If blnOffset Then
    If cboDateLinkOffsetPeriod.ListIndex < 0 Then
      cboDateLinkOffsetPeriod.ListIndex = 0
    End If
    If cboDateLinkDirection.ListIndex < 0 Then
      cboDateLinkDirection.ListIndex = 0
    End If
  Else
    cboDateLinkOffsetPeriod.ListIndex = -1
    cboDateLinkDirection.ListIndex = -1
  End If

  If Not mfLoading Then mblnChanged = True

End Sub

