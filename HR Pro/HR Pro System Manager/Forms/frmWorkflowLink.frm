VERSION 5.00
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmWorkflowLink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Workflow Link"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
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
   HelpContextID   =   5077
   Icon            =   "frmWorkflowLink.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLinkTypeDetails 
      Caption         =   "Date Related Link :"
      Height          =   2200
      Index           =   2
      Left            =   3000
      TabIndex        =   18
      Top             =   1800
      Width           =   4800
      Begin VB.ComboBox cboDateLinkOffsetPeriod 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmWorkflowLink.frx":000C
         Left            =   1950
         List            =   "frmWorkflowLink.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   700
         Width           =   1300
      End
      Begin VB.ComboBox cboDateLinkDirection 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmWorkflowLink.frx":003C
         Left            =   3360
         List            =   "frmWorkflowLink.frx":004C
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   700
         Width           =   1300
      End
      Begin VB.ComboBox cboDateLinkColumn 
         Height          =   315
         Left            =   1000
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   300
         Width           =   3660
      End
      Begin COASpinner.COA_Spinner spnDateLinkOffset 
         Height          =   315
         Left            =   1000
         TabIndex        =   22
         Top             =   700
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
      Begin VB.Label lblDateLinkOffset 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Offset :"
         Height          =   195
         Left            =   200
         TabIndex        =   21
         Top             =   760
         Width           =   570
      End
      Begin VB.Label lblDateLinkColumn 
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Left            =   195
         TabIndex        =   19
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.Frame fraLinkTypeDetails 
      Caption         =   "Record Related Link :"
      Height          =   2200
      Index           =   1
      Left            =   2280
      TabIndex        =   14
      Top             =   1800
      Width           =   4800
      Begin VB.CheckBox chkRecordLinkRecord 
         Caption         =   "&Insert Record"
         Height          =   240
         Index           =   0
         Left            =   200
         TabIndex        =   15
         Top             =   360
         Width           =   1890
      End
      Begin VB.CheckBox chkRecordLinkRecord 
         Caption         =   "&Update Record"
         Height          =   240
         Index           =   1
         Left            =   200
         TabIndex        =   16
         Top             =   1160
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.CheckBox chkRecordLinkRecord 
         Caption         =   "D&elete Record"
         Height          =   240
         Index           =   2
         Left            =   200
         TabIndex        =   17
         Top             =   760
         Width           =   1890
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3900
      TabIndex        =   23
      Top             =   4150
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5200
      TabIndex        =   24
      Top             =   4150
      Width           =   1200
   End
   Begin VB.Frame fraLinkType 
      Caption         =   "Link Type :"
      Height          =   2200
      Left            =   100
      TabIndex        =   8
      Top             =   1800
      Width           =   1380
      Begin VB.OptionButton optLinkType 
         Caption         =   "D&ate"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   11
         Tag             =   "2"
         Top             =   1160
         Width           =   700
      End
      Begin VB.OptionButton optLinkType 
         Caption         =   "&Record"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Tag             =   "1"
         Top             =   760
         Width           =   900
      End
      Begin VB.OptionButton optLinkType 
         Caption         =   "Co&lumn"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Tag             =   "0"
         Top             =   360
         Value           =   -1  'True
         Width           =   990
      End
   End
   Begin VB.Frame fraLinkTypeDetails 
      Caption         =   "Column Related Link :"
      Height          =   2200
      Index           =   0
      Left            =   1600
      TabIndex        =   12
      Top             =   1800
      Width           =   4800
      Begin VB.ListBox lstColumnLinkColumns 
         Height          =   1860
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   220
         Width           =   4545
      End
   End
   Begin VB.Frame fraLinkDetails 
      Caption         =   "Link Details :"
      Height          =   1600
      Left            =   100
      TabIndex        =   0
      Top             =   120
      Width           =   6300
      Begin VB.ComboBox cboWorkflow 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   4420
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
      Begin VB.CommandButton cmdFilter 
         Caption         =   "..."
         Height          =   315
         Left            =   5835
         TabIndex        =   5
         Top             =   700
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin GTMaskDate.GTMaskDate cboEffectiveDate 
         Height          =   315
         Left            =   1740
         TabIndex        =   7
         Top             =   1095
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
      Begin VB.Label lblWorkflowName 
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label lblFilter 
         AutoSize        =   -1  'True
         Caption         =   "Filter :"
         Height          =   195
         Left            =   200
         TabIndex        =   3
         Top             =   760
         Width           =   465
      End
      Begin VB.Label lblEffectiveDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date :"
         Height          =   195
         Left            =   200
         TabIndex        =   6
         Top             =   1160
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmWorkflowLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjWorkflowLink As clsWorkflowTriggeredLink
Private mblnCancelled As Boolean
Private mblnReadOnly As Boolean
Private mblnLoading As Boolean

Private miLinkType As WorkflowTriggerLinkType

Public Enum WFTriggerRelatedRecord
  WFRELATEDRECORD_INSERT = 0
  WFRELATEDRECORD_UPDATE = 1
  WFRELATEDRECORD_DELETE = 2
End Enum
Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property


Private Sub cboWorkflow_refresh()
  
  ' Initialise the Workflow combo(s)
  Dim iListIndex As Integer

  iListIndex = 0

  ' Clear the combo, and add '<None>' items.
  cboWorkflow.Clear

  ' Add items to the combo for each trigger-initiated workflow based on the current table
  ' that has not been deleted.
  With recWorkflowEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If (Not !Deleted) _
        And (!InitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED) _
        And (!BaseTable = mobjWorkflowLink.TableID) Then

        cboWorkflow.AddItem !Name
        cboWorkflow.ItemData(cboWorkflow.NewIndex) = !id

        If !id = mobjWorkflowLink.WorkflowID Then
          iListIndex = cboWorkflow.NewIndex
        End If
      End If
      
      .MoveNext
    Loop
  End With

  With cboWorkflow
    If .ListCount = 0 Then
      .AddItem "<None>"
      .ItemData(.NewIndex) = 0
      .ListIndex = .NewIndex
      .Enabled = False
    Else
      .Enabled = Not mblnReadOnly
      .ListIndex = iListIndex
    End If
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
  End With

End Sub


Private Sub cboDateLinkOffsetPeriod_refresh()
  
  ' Initialise the Date Offset Period combo
  'Dim iListIndex As Integer

  'iListIndex = 0

  ' Clear the combo, and add '<None>' items.
  With cboDateLinkOffsetPeriod
    .Clear

    ' Add items to the combo for each possible period
    .AddItem "Day(s)"
    .ItemData(.NewIndex) = WORKFLOWTRIGGERLINKOFFESTPERIOD_DAY
    'If mobjWorkflowLink.DateOffsetPeriod = .ItemData(.NewIndex) Then
    '  iListIndex = .NewIndex
    'End If
  
    .AddItem "Week(s)"
    .ItemData(.NewIndex) = WORKFLOWTRIGGERLINKOFFESTPERIOD_WEEK
    'If mobjWorkflowLink.DateOffsetPeriod = .ItemData(.NewIndex) Then
    '  iListIndex = .NewIndex
    'End If
  
    .AddItem "Month(s)"
    .ItemData(.NewIndex) = WORKFLOWTRIGGERLINKOFFESTPERIOD_MONTH
    'If mobjWorkflowLink.DateOffsetPeriod = .ItemData(.NewIndex) Then
    '  iListIndex = .NewIndex
    'End If
  
    .AddItem "Year(s)"
    .ItemData(.NewIndex) = WORKFLOWTRIGGERLINKOFFESTPERIOD_YEAR
    'If mobjWorkflowLink.DateOffsetPeriod = .ItemData(.NewIndex) Then
    '  iListIndex = .NewIndex
    'End If

    .ListIndex = -1 'iListIndex
  End With


End Sub


Private Sub cboDateLinkDirection_refresh()
  
  With cboDateLinkDirection
    .Clear
    .AddItem "Before"
    .AddItem "After"
    .ListIndex = -1
  End With


End Sub



Public Property Get WorkflowLink() As clsWorkflowTriggeredLink
  Set WorkflowLink = mobjWorkflowLink
  
End Property


Public Function PopulateControls() As Boolean

  Dim lngCount As Long
  Dim fOK As Boolean
  
  fOK = True
  mblnLoading = True
  
  With mobjWorkflowLink
    cboWorkflow_refresh
    cboDateLinkOffsetPeriod_refresh
    cboDateLinkDirection_refresh
    
    PopulateAvailable
    
    txtFilter.Tag = .FilterID
    txtFilter.Text = GetExpressionName(txtFilter.Tag)

    cboEffectiveDate.DateValue = .EffectiveDate
    
    miLinkType = .LinkType
    optLinkType(.LinkType).value = True

    DisplayLinkTypeFrame
    
    Select Case .LinkType
      Case WORKFLOWTRIGGERLINKTYPE_COLUMN
        ' Column list configured in the 'PopulateAvailable' method, called above
        
      Case WORKFLOWTRIGGERLINKTYPE_RECORD
        chkRecordLinkRecord(WFRELATEDRECORD_INSERT).value = IIf(.RecordInsert, vbChecked, vbUnchecked)
        chkRecordLinkRecord(WFRELATEDRECORD_UPDATE).value = IIf(.RecordUpdate, vbChecked, vbUnchecked)
        chkRecordLinkRecord(WFRELATEDRECORD_DELETE).value = IIf(.RecordDelete, vbChecked, vbUnchecked)

      Case WORKFLOWTRIGGERLINKTYPE_DATE
        
        'MH20090722 HRPRO-123
        'spnDateLinkOffset.value = .DateOffset
        cboDateLinkDirection.ListIndex = IIf(.DateOffset < 0, 0, 1)
        spnDateLinkOffset.value = Abs(.DateOffset)
        SetComboItem cboDateLinkOffsetPeriod, .DateOffsetPeriod
        
    End Select
  End With

  If mblnReadOnly Then
    ControlsDisableAll Me
  End If
  
  fOK = (cboWorkflow.ListCount > 0)
  If Not fOK Then
    MsgBox "Unable to set up any Workflow links on this table as there are no 'triggered' Workflows based on this table.", vbCritical
  End If

  mblnLoading = False

  Changed = Not fOK
  PopulateControls = fOK

End Function


Private Sub PopulateAvailable()
  Dim iDateLinkColumnIndex As Integer
  Dim lngCount As Long
    
  iDateLinkColumnIndex = 0
  
  lstColumnLinkColumns.Clear
  cboDateLinkColumn.Clear

  With recColEdit
    .Index = "idxName"
    .Seek ">=", mobjWorkflowLink.TableID

    If Not .NoMatch Then
      ' Add items to the listview for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        If !TableID <> mobjWorkflowLink.TableID Then
          Exit Do
        End If

        If (Not !Deleted) And _
          (!columnType <> giCOLUMNTYPE_LINK) And _
          (!columnType <> giCOLUMNTYPE_SYSTEM) And _
          (!DataType <> dtVARBINARY) And _
          (!DataType <> dtLONGVARBINARY) Then

          lstColumnLinkColumns.AddItem (!ColumnName)
          lstColumnLinkColumns.ItemData(lstColumnLinkColumns.NewIndex) = !ColumnID

          If miLinkType = WORKFLOWTRIGGERLINKTYPE_COLUMN Then
            For lngCount = 1 To mobjWorkflowLink.LinkColumns.Count
              If mobjWorkflowLink.LinkColumns(lngCount).ColumnID = !ColumnID Then
                lstColumnLinkColumns.Selected(lstColumnLinkColumns.NewIndex) = True
                Exit For
              End If
            Next lngCount
          End If
          
          If !DataType = sqlDate Then
            cboDateLinkColumn.AddItem !ColumnName
            cboDateLinkColumn.ItemData(cboDateLinkColumn.NewIndex) = !ColumnID
            
            If !ColumnID = mobjWorkflowLink.DateColumnID Then
              iDateLinkColumnIndex = cboDateLinkColumn.NewIndex
            End If
          End If
        End If

        .MoveNext
      Loop
    End If
  End With
  
  If cboDateLinkColumn.ListCount = 0 Then
    cboDateLinkColumn.AddItem "<None>"
    cboDateLinkColumn.ItemData(cboDateLinkColumn.NewIndex) = 0
    cboDateLinkColumn.ListIndex = 0
  Else
    cboDateLinkColumn.ListIndex = iDateLinkColumnIndex
  End If
  
End Sub



Public Property Get Changed() As Boolean
  Changed = cmdOk.Enabled
End Property


Public Property Let Changed(ByVal blnNewValue As Boolean)
  If Not mblnLoading Then
    cmdOk.Enabled = blnNewValue And Not mblnReadOnly
  End If
End Property


Public Property Let WorkflowLink(ByVal objNewValue As clsWorkflowTriggeredLink)
  Set mobjWorkflowLink = objNewValue
  
End Property


Private Sub cboDateLinkColumn_Click()
  Changed = True

End Sub


Private Sub cboDateLinkDirection_Click()
  Changed = True

End Sub

Private Sub cboDateLinkOffsetPeriod_Click()
  Changed = True

End Sub

Private Sub cboEffectiveDate_Change()
  Changed = True

End Sub

Private Sub cboEffectiveDate_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    cboEffectiveDate.DateValue = Date
  End If

End Sub

Private Sub cboEffectiveDate_LostFocus()

  If IsEmpty(cboEffectiveDate.DateValue) Then
     MsgBox "You must enter a date.", vbOKOnly + vbExclamation, App.Title
  End If
  
  ValidateGTMaskDate cboEffectiveDate

End Sub


Private Sub cboWorkflow_Click()
  Changed = True

End Sub


Private Sub chkRecordLinkRecord_Click(Index As Integer)
  Changed = True

End Sub


Private Sub cmdCancel_Click()
  If Me.Changed Then
    Select Case MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
    Case vbYes
      cmdOK_Click
      Exit Sub
    Case vbCancel
      Exit Sub
    End Select
  End If

  Me.Hide

End Sub

Private Sub cmdFilter_Click()
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim objExpr As CExpression

  fOK = True

  ' Instantiate an expression object.
  Set objExpr = New CExpression

  With objExpr
    ' Set the properties of the expression object.
    .Initialise mobjWorkflowLink.TableID, txtFilter.Tag, giEXPR_LINKFILTER, giEXPRVALUE_LOGIC

    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      txtFilter.Tag = .ExpressionID
      txtFilter.Text = GetExpressionName(txtFilter.Tag)
      Changed = True
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

Private Sub cmdOK_Click()
  If ValidDefinition = False Then
    Exit Sub
  End If

  SaveDefinition
  mblnCancelled = False
  Me.Hide

End Sub

Private Function ValidDefinition() As Boolean
  Dim fInvalidDate As Boolean
  
  ValidDefinition = False

  If cboWorkflow.ListCount = 0 Then
    MsgBox "No Workflow selected.", vbExclamation, Me.Caption
    Exit Function
  Else
    If cboWorkflow.ItemData(cboWorkflow.ListIndex) = 0 Then
      MsgBox "No Workflow selected.", vbExclamation, Me.Caption
      Exit Function
    End If
  End If
  
  fInvalidDate = IsNull(cboEffectiveDate.DateValue)
  If Not fInvalidDate Then
    fInvalidDate = IsEmpty(cboEffectiveDate.DateValue)
  End If
  
  If fInvalidDate Then
    MsgBox "No Effective Date defined.", vbExclamation, Me.Caption
    
    If cboEffectiveDate.Enabled Then
      cboEffectiveDate.SetFocus
    End If
    Exit Function
  End If

  If ValidateGTMaskDate(cboEffectiveDate) = False Then
    Exit Function
  End If

  Select Case miLinkType
    Case WORKFLOWTRIGGERLINKTYPE_COLUMN
      If lstColumnLinkColumns.SelCount = 0 Then
        MsgBox "No Columns selected.", vbExclamation, Me.Caption
        
        If lstColumnLinkColumns.Enabled Then
          lstColumnLinkColumns.SetFocus
        End If
        Exit Function
      End If
  
    Case WORKFLOWTRIGGERLINKTYPE_RECORD
      If chkRecordLinkRecord(WFRELATEDRECORD_INSERT).value = vbUnchecked _
        And chkRecordLinkRecord(WFRELATEDRECORD_UPDATE).value = vbUnchecked _
        And chkRecordLinkRecord(WFRELATEDRECORD_DELETE).value = vbUnchecked Then
      
        MsgBox "No Record action selected.", vbExclamation, Me.Caption
        
        If chkRecordLinkRecord(WFRELATEDRECORD_INSERT).Enabled Then
          chkRecordLinkRecord(WFRELATEDRECORD_INSERT).SetFocus
        End If
        Exit Function
      End If
            
    Case WORKFLOWTRIGGERLINKTYPE_DATE
      If (cboDateLinkColumn.ListCount = 0) Or (cboDateLinkColumn.ListIndex < 0) Then
        MsgBox "No Date Column selected.", vbExclamation, Me.Caption
        Exit Function
      Else
        If cboDateLinkColumn.ItemData(cboDateLinkColumn.ListIndex) = 0 Then
          MsgBox "No Date Column selected.", vbExclamation, Me.Caption
          Exit Function
        End If
      End If
  End Select

  ValidDefinition = True

End Function



Private Function SaveDefinition() As Boolean

  Dim lngCount As Long

  With mobjWorkflowLink
    .WorkflowID = cboWorkflow.ItemData(cboWorkflow.ListIndex)
    .FilterID = Val(txtFilter.Tag)
    .EffectiveDate = cboEffectiveDate.DateValue
    
    .LinkType = miLinkType
    
    .ClearColumns
    If miLinkType = WORKFLOWTRIGGERLINKTYPE_COLUMN Then
      For lngCount = 0 To lstColumnLinkColumns.ListCount - 1
        If lstColumnLinkColumns.Selected(lngCount) Then
          .AddColumn lstColumnLinkColumns.ItemData(lngCount)
        End If
      Next
    End If
    
    .RecordInsert = (miLinkType = WORKFLOWTRIGGERLINKTYPE_RECORD) _
      And (chkRecordLinkRecord(WFRELATEDRECORD_INSERT).value = vbChecked)
    .RecordUpdate = (miLinkType = WORKFLOWTRIGGERLINKTYPE_RECORD) _
      And (chkRecordLinkRecord(WFRELATEDRECORD_UPDATE).value = vbChecked)
    .RecordDelete = (miLinkType = WORKFLOWTRIGGERLINKTYPE_RECORD) _
      And (chkRecordLinkRecord(WFRELATEDRECORD_DELETE).value = vbChecked)
      
    If (miLinkType = WORKFLOWTRIGGERLINKTYPE_DATE) And (cboDateLinkColumn.ListIndex >= 0) Then
      .DateColumnID = cboDateLinkColumn.ItemData(cboDateLinkColumn.ListIndex)
      
      'MH20090722 HRPRO-123
      '.DateOffset = spnDateLinkOffset.value
      .DateOffset = spnDateLinkOffset.value * IIf(cboDateLinkDirection.ListIndex = 1, 1, -1)
      
      If cboDateLinkOffsetPeriod.ListIndex >= 0 Then
        .DateOffsetPeriod = cboDateLinkOffsetPeriod.ItemData(cboDateLinkOffsetPeriod.ListIndex)
      Else
        .DateOffsetPeriod = 0
      End If
    Else
      .DateColumnID = 0
      .DateOffset = 0
      .DateOffsetPeriod = WORKFLOWTRIGGERLINKOFFESTPERIOD_DAY
    End If
  End With

  SaveDefinition = True

End Function


Private Sub Form_Load()
  Dim fraTemp As Frame
  
  mblnCancelled = True
  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)

  UI.FormatGTDateControl cboEffectiveDate

  For Each fraTemp In fraLinkTypeDetails
    fraTemp.Left = fraLinkTypeDetails(0).Left
    fraTemp.Top = fraLinkTypeDetails(0).Top
  Next fraTemp
  Set fraTemp = Nothing
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    cmdCancel_Click
    Cancel = True
  End If

End Sub


Private Sub lstColumnLinkColumns_ItemCheck(Item As Integer)
  If (mblnReadOnly) And (Not mblnLoading) Then
    lstColumnLinkColumns.Selected(Item) = Not lstColumnLinkColumns.Selected(Item)
  End If
  Changed = True

End Sub

Private Sub optLinkType_Click(piIndex As Integer)
  ' Set the link type property.
  miLinkType = piIndex
  Changed = True
  
  DisplayLinkTypeFrame

End Sub


Private Sub DisplayLinkTypeFrame()
  Dim fraTemp As Frame
  
  ' Display only the frame that defines the selected link type.
  For Each fraTemp In fraLinkTypeDetails
    fraTemp.Left = fraLinkTypeDetails(0).Left
    fraTemp.Top = fraLinkTypeDetails(0).Top
    
    fraTemp.Visible = (fraTemp.Index = miLinkType)
  Next fraTemp
  Set fraTemp = Nothing
  
End Sub


'MH20090722 HRPRO-123
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

  Changed = True

End Sub

