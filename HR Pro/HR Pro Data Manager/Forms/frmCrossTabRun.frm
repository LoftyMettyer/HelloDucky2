VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCrossTabRun 
   Caption         =   "Cross Tabs"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1025
   Icon            =   "frmCrossTabRun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleMode       =   0  'User
   ScaleWidth      =   9540.714
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   780
      Top             =   6975
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrossTabRun.frx":000C
            Key             =   "AbsenceBreakdown"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrossTabRun.frx":042F
            Key             =   "StabilityIndex"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrossTabRun.frx":086F
            Key             =   "Turnover"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrossTabRun.frx":0CB0
            Key             =   "CrossTabs"
         EndProperty
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      CausesValidation=   0   'False
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9240
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      Col.Count       =   0
      stylesets.count =   1
      stylesets(0).Name=   "Highlight"
      stylesets(0).ForeColor=   -2147483634
      stylesets(0).BackColor=   -2147483635
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmCrossTabRun.frx":10A9
      DefColWidth     =   3528
      BevelColorHighlight=   -2147483643
      MultiLine       =   0   'False
      ActiveCellStyleSet=   "Highlight"
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      MaxSelectedRows =   1
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns(0).Width=   3528
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      TabNavigation   =   1
      _ExtentX        =   16298
      _ExtentY        =   9340
      _StockProps     =   79
      Caption         =   "SSDBGrid1"
      ForeColor       =   0
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
   Begin VB.Frame fraIntersection 
      Caption         =   "Intersection :"
      Height          =   1500
      Left            =   90
      TabIndex        =   1
      Top             =   5430
      Width           =   5925
      Begin VB.CheckBox chkThousandSeparators 
         Caption         =   "Use 1000 &separators (,)"
         Height          =   330
         Left            =   3465
         TabIndex        =   7
         Top             =   1095
         Width           =   2430
      End
      Begin VB.CheckBox chkPercentageOfPage 
         Caption         =   "Percentage of &Page"
         Height          =   225
         Left            =   3465
         TabIndex        =   5
         Top             =   555
         Width           =   2385
      End
      Begin VB.TextBox txtIntersectionCol 
         Height          =   315
         Left            =   1170
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   270
         Width           =   2085
      End
      Begin VB.CheckBox chkSuppressZeros 
         Caption         =   "Suppress &Zeros"
         Height          =   255
         Left            =   3465
         TabIndex        =   6
         Top             =   840
         Width           =   2040
      End
      Begin VB.CheckBox chkPercentage 
         Caption         =   "Percentage of &Type"
         Height          =   195
         Left            =   3465
         TabIndex        =   4
         Top             =   270
         Width           =   2385
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmCrossTabRun.frx":10C5
         Left            =   1170
         List            =   "frmCrossTabRun.frx":10C7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   675
         Width           =   2115
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Left            =   225
         TabIndex        =   15
         Top             =   315
         Width           =   810
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
         Height          =   195
         Left            =   225
         TabIndex        =   16
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame fraPage 
      Caption         =   "Page :"
      Height          =   1500
      Left            =   6135
      TabIndex        =   8
      Top             =   5430
      Width           =   3240
      Begin VB.TextBox txtPageBreakCol 
         Height          =   315
         Left            =   1140
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   270
         Width           =   1980
      End
      Begin VB.ComboBox cboPageBreak 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   675
         Width           =   1980
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value :"
         Height          =   195
         Left            =   225
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblPageBreak 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   315
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "O&utput..."
      Height          =   400
      Left            =   6885
      TabIndex        =   13
      Top             =   7005
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   400
      Left            =   8175
      TabIndex        =   14
      Top             =   7005
      Width           =   1200
   End
End
Attribute VB_Name = "frmCrossTabRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'JDM - 22/06/01 - Moved definition of CrossTabType to modHRPro

Public mblnLoading As Boolean
'Private gblnBatchMode As Boolean
Private CrossTabPrint As clsPrintGrid
Private datData As HRProDataMgr.clsDataAccess          'DataAccess Class
Private fOK As Boolean
Private mstrErrorMessage As String
Private lngFileNum As Integer
Private mblnUserCancelled As Boolean
Private mblnNoRecords As Boolean
Private mlngCrossTabType As CrossTabType
Private mstrTempTableName As String
Private mcmdUpdateRecDescs As ADODB.Command
Private frmBreakDown As frmCrossTabCellBreakDown

Private Const HOR As Integer = 0  'Horizontal
Private Const VER As Integer = 1  'Vertical
Private Const PGB As Integer = 2  'Page Break
Private Const INS As Integer = 3  'Intersection

Private Const TYPECOUNT As Integer = 0
Private Const TYPEAVERAGE As Integer = 1
Private Const TYPEMAXIMUM As Integer = 2
Private Const TYPEMINIMUM As Integer = 3
Private Const TYPETOTAL As Integer = 4

Private mstrBaseTable As String
Private mlngBaseTableID As Long
Private rsCrossTabData As Recordset
Private mblnIntersection As Boolean
Private mblnPageBreak As Boolean
Private mblnShowAllPagesTogether As Boolean
Private mdtReportStartDate As Date
Private mdtReportEndDate As Date
    
Private mstrCrossTabName As String
Private mlngIntersectionType As Long
Private mblnShowPercentage As Boolean
Private mblnPercentageofPage As Boolean
Private mblnSuppressZeros As Boolean
Private mbThousandSeparators As Boolean
Private mlngRecordDescExprID As Long
Private mstrPicklistFilter As String

Private mstrSQLSelect As String
Private mstrSQLFrom As String
Private mstrSQLJoin As String
Private mstrSQLWhere As String
  
Private mstrIntersectionMask As String
Private mlngIntersectionDecimals As Long
Private mdblPercentageFactor As Double
Private mlngFirstColWidth As Long

'The column headings are held in this array of variants.
'The reason that they are in variants and not a two dimensional array of
'strings is because the three headings arrays can have different upper
'bound limits.
'NOTE: To reference these you will need to use the following syntax:
'
' mvarHeadings(1)(3) - this will reference array no. 1, element 3
'
Private mvarHeadings(2) As Variant
Private mvarSearches(2) As Variant

Private mdblHorTotal() As Double
Private mdblVerTotal() As Double
Private mdblPgbTotal() As Double
Private mdblPageTotal() As Double
Private mdblGrandTotal() As Double

Private mdblDataArray() As Double
Private mstrOutput() As String

Private mstrType() As String      'e.g. mstrtype(TYPETOTAL) = for example: "Total"
Private mlngColID(3) As Long
Private mstrColName(3) As String  'e.g. mstrColName(INS) = "Salary" (the name of the intersection field)
Private mlngColDataType(3) As String
Private mstrColCase(3) As String
Private mstrFormat(3) As String
Private mdblMin(2) As Double
Private mdblMax(2) As Double
Private mdblStep(2) As Double

Private mlngMinFormWidth As Long
Private mlngMinFormHeight As Long

Private mblnInvalidPicklistFilter As Boolean

'Private mblnCustomDates As Boolean
'Private mlngStartDateExprID As Long
'Private mlngEndDateExprID As Long
Private mlngPicklistFilterID As Long
Private mstrPicklistFilterType As String
Private mlngVerCol As Long
Private mlngHorCol As Long
Private mlngPicklistID As Long
Private mlngFilterID As Long
Private msAbsenceBreakdownTypes As String
Private mblnIncludeNewStarters As Boolean
Private mlngHorColID As Long
Private mblnPreviewOnScreen As Boolean
Private mlngOutputFormat As Long
Private mblnOutputScreen As Boolean
Private mblnOutputPrinter As Boolean
Private mstrOutputPrinterName As String
Private mblnOutputSave As Boolean
Private mlngOutputSaveExisting As Long
Private mblnOutputEmail As Boolean
Private mlngOutputEmailAddr As Long
Private mstrOutputEmailSubject As String
Private mstrOutputEmailAttachAs As String
Private mstrOutputFilename As String
Private mblnChkPicklistFilter As String

' Array holding the User Defined functions that are needed for this report
Private mastrUDFsRequired() As String


Private Function IsRecordSelectionValid() As Boolean
  Dim sSQL As String
  Dim lCount As Long
  Dim rsTemp As Recordset
  Dim iResult As RecordSelectionValidityCodes

' Base Table First
  If mlngFilterID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngFilterID)
    Select Case iResult
      Case REC_SEL_VALID_DELETED
        mstrErrorMessage = "The base table filter used in this definition has been deleted by another user."
      Case REC_SEL_VALID_INVALID
        mstrErrorMessage = "The base table filter used in this definition is invalid."
      Case REC_SEL_VALID_HIDDENBYOTHER
        If Not gfCurrentUserIsSysSecMgr Then
          mstrErrorMessage = "The base table filter used in this definition has been made hidden by another user."
        End If
    End Select
  ElseIf mlngPicklistID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, mlngPicklistID)
    Select Case iResult
      Case REC_SEL_VALID_DELETED
        mstrErrorMessage = "The base table picklist used in this definition has been deleted by another user."
      Case REC_SEL_VALID_INVALID
        mstrErrorMessage = "The base table picklist used in this definition is invalid."
      Case REC_SEL_VALID_HIDDENBYOTHER
        If Not gfCurrentUserIsSysSecMgr Then
          mstrErrorMessage = "The base table picklist used in this definition has been made hidden by another user."
        End If
    End Select
  End If

  IsRecordSelectionValid = (Len(mstrErrorMessage) = 0)
  
End Function



Private Function CheckIfLeaver(dtTemp As Variant) As Boolean

  CheckIfLeaver = False
  If Not IsNull(dtTemp) Then
    'CheckIfLeaver = (DateDiff("d", dtTemp, mdtReportEndDate) > 0)
    CheckIfLeaver = (DateDiff("d", dtTemp, mdtReportEndDate) >= 0)
  End If

End Function


Private Function SQLEmployedAtStartOfReport(strColStart As String, strColLeaving As String) As String

  SQLEmployedAtStartOfReport = _
    "(Datediff(d,'" & FormatDateSQL(mdtReportStartDate) & _
    "'," & strColStart & ") <= 0 OR " & strColStart & " IS NULL)" & _
    " AND " & _
    "(Datediff(d,'" & FormatDateSQL(mdtReportStartDate) & _
    "'," & strColLeaving & ") >= 0 OR " & strColLeaving & " IS NULL)"

End Function
      
      
Private Function SQLLeaversBetweenStartAndEnd(strColStart As String, strColLeaving As String) As String

  SQLLeaversBetweenStartAndEnd = _
    "Datediff(d,'" & FormatDateSQL(mdtReportEndDate) & _
    "'," & strColLeaving & ") <= 0"
  
  If mblnIncludeNewStarters Then
    SQLLeaversBetweenStartAndEnd = _
      "(Datediff(d,'" & FormatDateSQL(mdtReportEndDate) & _
      "'," & strColStart & ") <= 0 OR " & strColStart & " IS NULL) AND " & _
      SQLLeaversBetweenStartAndEnd
  Else
    SQLLeaversBetweenStartAndEnd = _
      SQLEmployedAtStartOfReport(strColStart, strColLeaving) & " AND " & _
      SQLLeaversBetweenStartAndEnd
  End If

End Function
   
    
Private Function SQLEmployedAtEndOfReport(strColStart As String, strColLeaving As String)

  SQLEmployedAtEndOfReport = _
    "(Datediff(d,'" & FormatDateSQL(mdtReportEndDate) & _
    "'," & strColStart & ") <= 0 OR " & strColStart & " IS NULL) AND " & _
    "(Datediff(d,'" & FormatDateSQL(mdtReportEndDate) & _
    "'," & strColLeaving & ") > 0 OR " & strColLeaving & " IS NULL)"

End Function
        
        
Private Function SQLOneYearServiceAtEndOfReport(strColStart As String, strColLeaving As String)

  SQLOneYearServiceAtEndOfReport = _
    "(Datediff(d,'" & FormatDateSQL(mdtReportStartDate) & _
    "'," & strColStart & ") <= 0 OR " & strColStart & " IS NULL) AND " & _
    "(Datediff(d,'" & FormatDateSQL(mdtReportEndDate) & _
    "'," & strColLeaving & ") > 0 OR " & strColLeaving & " IS NULL)"

End Function


Private Function HTMLText(strInput As String) As String
  
  HTMLText = Replace(strInput, "<", "&LT;")
  HTMLText = Replace(HTMLText, ">", "&GT;")
  HTMLText = Replace(HTMLText, vbTab, "</TD><TD>")
  HTMLText = Replace(HTMLText, "<TD></TD>", "<TD>&nbsp;</TD>")
  If Left(HTMLText, 5) = "</TD>" Then
    HTMLText = "&nbsp;" & HTMLText
  End If
  If Right(HTMLText, 4) = "<TD>" Then
    HTMLText = HTMLText & "&nbsp;"
  End If

End Function

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property

Public Function AbsenceBreakdownExecuteReport(lngPersonnelID As Long)

'(dtStartDate As Date, dtEndDate As Date _
    , lngHorColID As Long, lngVerColID As Long, lngPicklistID As Long, lngFilterID As Long _
    , lngPersonnelID As Long, asIncludedTypes() As String)
  
  'NHRD09042002 Fault 3322 - Code Added
  'mlngPickListID = lngPicklistID
  'mlngFilterID = lngFilterID

  Set datData = New HRProDataMgr.clsDataAccess
  mblnLoading = True
  
  ReDim mastrUDFsRequired(0)
  
  mlngCrossTabType = cttAbsenceBreakdown
  fOK = True
  
  If gblnBatchMode = True Then
    Call GetReportConfig("AbsenceBreakdown")
  End If
  AbsenceBreakdownRetreiveDefinition lngPersonnelID

  gobjEventLog.AddHeader eltStandardReport, mstrCrossTabName
  
  If fOK Then InitialiseProgressBar
  If Progress Then UDFFunctions mastrUDFsRequired, True
  If Progress Then CreateTempTable
  If Progress Then UDFFunctions mastrUDFsRequired, True
  If Progress Then AbsenceBreakdownRunStoredProcedure
  If Progress Then AbsenceBreakdownGetHeadingsAndSearches
  If Progress Then BuildTypeArray
  If Progress Then AbsenceBreakdownBuildDataArrays
  If Progress Then PopulateCombos
  If Progress Then PrepareForms
  If Progress Then CreateGridColumns
  If Progress Then PopulateGrid

  If Progress Then
    If gblnBatchMode Or Not mblnPreviewOnScreen Then
      fOK = OutputReport(False)
    End If
  End If
  
  
  mblnLoading = False

  Call OutputJobStatus

  AbsenceBreakdownExecuteReport = fOK

End Function

Private Property Let InvalidPicklistFilter(bValid As Boolean)
  mblnInvalidPicklistFilter = bValid
End Property

Public Property Get InvalidPicklistFilter() As Boolean
  InvalidPicklistFilter = mblnInvalidPicklistFilter
End Property

Public Property Get ErrorString() As String

  ErrorString = mstrErrorMessage
  
End Property

Private Sub cboPageBreak_Click()
  If mblnLoading Then Exit Sub
  PopulateGrid
  cboPageBreak.SetFocus
End Sub

Private Sub cboType_Click()
  If mblnLoading Then Exit Sub
  
  PopulateGrid
  'NHRD27102004 Fault 8975
  SSDBGrid1.Scroll SSDBGrid1.Cols, 0
  cboType.SetFocus
End Sub

Private Sub chkPercentage_Click()
  
  Dim blnEnabled As Boolean
  
  If mblnLoading Then Exit Sub
  
  blnEnabled = (chkPercentage.Value = vbChecked And cboPageBreak.Enabled = True)
  chkPercentageOfPage.Enabled = blnEnabled
  If Not blnEnabled Then
    mblnLoading = True
    chkPercentageOfPage.Value = vbUnchecked
    mblnLoading = False
  End If

  PopulateGrid
  chkPercentage.SetFocus

End Sub


Private Sub chkPercentageOfPage_Click()
  If mblnLoading Then Exit Sub
  PopulateGrid
  chkPercentageOfPage.SetFocus
End Sub

Private Sub chkSuppressZeros_Click()
  If mblnLoading Then Exit Sub
  Call PopulateGrid
  chkSuppressZeros.SetFocus
End Sub

Private Sub chkThousandSeparators_Click()

  If mblnLoading Then Exit Sub
  Call PopulateGrid
  chkThousandSeparators.SetFocus

End Sub

Private Sub cmdOK_Click()
  Unload Me
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
  If Me.Visible And Me.Enabled Then
    cmdOK.SetFocus
  End If
End Sub


Private Function Progress() As Boolean
  
  'This needs to be here, otherwise the progress bar will continue to the end
  'rather than cancelling immediately
  If fOK = False Then
    Progress = False
    Exit Function
  End If
  
  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    fOK = False
  End If
  
  gobjProgress.UpdateProgress gblnBatchMode
  Progress = fOK

End Function

Public Function ExecuteCrossTab(lngCrossTabID As Long) As Boolean
  
  Set datData = New HRProDataMgr.clsDataAccess
  mblnLoading = True

  mlngCrossTabType = cttNormal  'Not turnover nor stability report !
  mblnShowAllPagesTogether = False

  fOK = True

  ReDim mastrUDFsRequired(0)

  ' RH 16/02/01 - This must be done AFTER the main definition has been retrieved
  '               otherwise mstrCrossTabName will be blank !
  'gobjEventLog.AddHeader eltCrossTab, mstrCrossTabName

  Call RetreiveDefinition(lngCrossTabID)
  
  gobjEventLog.AddHeader eltCrossTab, mstrCrossTabName

  If fOK Then Call InitialiseProgressBar
  If Progress Then Call UDFFunctions(mastrUDFsRequired, True)
  If Progress Then Call CreateTempTable
  If Progress Then Call UDFFunctions(mastrUDFsRequired, False)
  If Progress Then Call GetHeadingsAndSearches
  If Progress Then Call BuildTypeArray
  If Progress Then Call BuildDataArrays
  If Progress Then Call PopulateCombos
  If Progress Then Call PrepareForms
  
  If mblnShowAllPagesTogether Then
    If Progress Then Call TurnoverCreateGridColumns
  Else
    If Progress Then Call CreateGridColumns
  End If

  If Progress Then Call PopulateGrid

  mblnLoading = False
  
  If Progress Then
    If gblnBatchMode Or Not mblnPreviewOnScreen Then
      fOK = OutputReport(False)
    End If
  End If
  
  Call UtilUpdateLastRun(utlCrossTab, lngCrossTabID)
  Call OutputJobStatus

  ExecuteCrossTab = fOK

End Function


Private Sub OutputJobStatus()

  mblnUserCancelled = _
    (mblnUserCancelled Or gobjProgress.Cancelled Or _
    (InStr(LCase(mstrErrorMessage), "cancelled by user") > 0))

  With gobjEventLog
  
    If fOK Or mblnNoRecords Then
      .ChangeHeaderStatus elsSuccessful
      
      If mblnNoRecords Then
        mstrErrorMessage = "Completed successfully." & vbCrLf & _
                          mstrErrorMessage
        .AddDetailEntry mstrErrorMessage
        fOK = True
      Else
        mstrErrorMessage = "Completed successfully."
      End If

    ElseIf mblnUserCancelled Then
      .ChangeHeaderStatus elsCancelled
      mstrErrorMessage = "Cancelled by user."
    Else
      'Only details records for failures !
      .AddDetailEntry mstrErrorMessage
      .ChangeHeaderStatus elsFailed
      mstrErrorMessage = "Failed." & vbCrLf & vbCrLf & mstrErrorMessage
    End If
  
  End With
  
  'mstrErrorMessage = _
    IIf(mlngCrossTabType = cttNormal, "Cross Tab : ", "") & _
    mstrCrossTabName & " " & mstrErrorMessage
  If mlngCrossTabType = cttNormal Then
    mstrErrorMessage = "Cross Tab : '" & mstrCrossTabName & "' " & mstrErrorMessage
  Else
    mstrErrorMessage = mstrCrossTabName & " " & mstrErrorMessage
  End If


  If Not gblnBatchMode Then
    If gobjProgress.Visible Then
      gobjProgress.CloseProgress
    End If
    If (fOK = False) Or (mblnNoRecords = True) Or (Not mblnPreviewOnScreen) Then
      MsgBox mstrErrorMessage, IIf(fOK Or mblnNoRecords, vbInformation, vbExclamation) + vbOKOnly, Me.Caption
    End If
  Else
    gobjProgress.ResetBar2
  End If

End Sub


Private Sub InitialiseProgressBar()

Dim strUtilityAction As String
  
  With gobjProgress
    '.AviFile = App.Path & "\videos\crosstab.avi"
      .AVI = dbText
      Select Case mlngCrossTabType
        Case cttNormal
          strUtilityAction = "Cross Tab"
        Case cttTurnover
          strUtilityAction = "Turnover Report"
        Case cttStability
          strUtilityAction = "Stability Index"
        Case cttAbsenceBreakdown
          strUtilityAction = "Absence Breakdown"
        Case Else
          strUtilityAction = "Cross Tab"
      End Select
      
    If Not gblnBatchMode Then
      .CloseProgress
      .NumberOfBars = 1
      .Caption = Me.Caption
      .Time = False
      .Cancel = True
      .Bar1Value = 0
      .Bar1MaxValue = IIf(mlngCrossTabType = cttNormal, 12, 17)
      .Bar1Caption = IIf(mlngCrossTabType = cttNormal, "Cross Tab : " & mstrCrossTabName, Me.Caption)
      .MainCaption = strUtilityAction
      .OpenProgress
    Else
      .ResetBar2
      .Bar2MaxValue = IIf(mlngCrossTabType = cttNormal, 12, 17)
      .Bar2Caption = IIf(mlngCrossTabType = cttNormal, "Cross Tab : " & mstrCrossTabName, Me.Caption)
      .MainCaption = strUtilityAction
    End If
  End With
  

End Sub

Private Sub BuildTypeArray()

  On Error GoTo LocalErr
  
  If mblnIntersection Then
    
    ReDim mstrType(4) As String
    mstrType(TYPECOUNT) = "Count"
    mstrType(TYPEAVERAGE) = "Average"
    mstrType(TYPEMAXIMUM) = "Maximum"
    mstrType(TYPEMINIMUM) = "Minimum"
    mstrType(TYPETOTAL) = "Total"
  
  Else

      ReDim mstrType(0) As String
      mstrType(TYPECOUNT) = "Count"
 
  End If

Exit Sub

LocalErr:
  mstrErrorMessage = "Error building type array"
  fOK = False
  
End Sub

Private Sub AbsenceBreakdownRunStoredProcedure()

  ' Purpose : To re-jig the selected records from the normal cross tab so it can be used in the standard
  '           crosstab output.

  On Error GoTo LocalErr
  
  Dim sSQL As String
  Dim iAffectedRecords As Integer
  Dim lngAffectedRecords As Long
  Dim strReportStartDate As String
  Dim strReportEndDate As String

  ' Convert dates to string to pass to stored procedure
  'TM24062004 Fault 8799 - we don't need to pass the full date AND time through.
'  strReportStartDate = Format(mdtReportStartDate, "YYYY/MM/DD 00:00:00.000")
'  strReportEndDate = Format(mdtReportEndDate, "YYYY/MM/DD 00:00:00.000")
  strReportStartDate = Format(mdtReportStartDate, "YYYY/MM/DD")
  strReportEndDate = Format(mdtReportEndDate, "YYYY/MM/DD")
  

  'Fire off the stored procedure to turn the current data into something the crosstab code will like.
  sSQL = "EXECUTE sp_ASR_AbsenceBreakdown_Run '" + strReportStartDate + "','" + strReportEndDate + "','" + mstrTempTableName + "'"
  datData.ExecuteSql sSQL

  ' Check that records exist (in case there's no working pattern and such like)
  Dim rsCrossTabData As Recordset
  Set rsCrossTabData = New Recordset
  rsCrossTabData.ActiveConnection = gADOCon
  rsCrossTabData.Open "Select * From " & mstrTempTableName, , adOpenKeyset, adLockOptimistic, adCmdText

  If rsCrossTabData.EOF Then
    mstrErrorMessage = "No records meet selection criteria."
    mblnNoRecords = True
    fOK = False
  End If

  ' Fault 2422 - Switch days into language of client machine (server independant)
  ' JDM - 19/06/01 - Fault 2472 - Whoops, missed out some error checking...
  If fOK Then
    With rsCrossTabData
      .MoveFirst
      Do While Not .EOF And Not gobjProgress.Cancelled

        If .Fields("Day_Number") < 8 Then
          .Fields("HOR") = WeekdayName(.Fields("Day_Number"), False, vbMonday)
        End If

        .MoveNext
      Loop

    End With
  End If

  Exit Sub

LocalErr:
  mstrErrorMessage = "Error running stored procedure in database"
  fOK = False
  
End Sub

Private Function GetGroupNumber(strValue As String, Index As Integer)

  'This returns which array element a particular value should be added to
  'Examples:
  '
  ' value = null, Minimum = 16, Maximum = 70, Step = 5
  '    GetGroupNumber = 0
  '
  ' value = 11, Minimum = 16, Maximum = 70, Step = 5
  '    GetGroupNumber = 1
  '
  ' value = 18, Minimum = 16, Maximum = 70, Step = 5
  '    GetGroupNumber = 2
  '
  ' value = 26, Minimum = 16, Maximum = 70, Step = 5
  '    GetGroupNumber = 4
  '
  ' value = 92, Minimum = 16, Maximum = 70, Step = 5
  '    GetGroupNumber = 13

  On Error GoTo LocalErr

  Dim dblValue As Double
  Dim lngCount As Long
  Dim dblLoop As Double
  Dim rsTemp As Recordset

  GetGroupNumber = 0
  'GetGroupNumber = IIf(strValue = vbNullString, 0, -1)

  'Non range column
  If mdblMin(Index) = 0 And mdblMax(Index) = 0 Then

    For lngCount = 0 To UBound(mvarHeadings(Index))

      Select Case mlngColDataType(Index)
      Case sqlDate
        If mvarHeadings(Index)(lngCount) = Format(strValue, DateFormat) Then
          GetGroupNumber = lngCount
          Exit For
        End If

      Case sqlNumeric, sqlInteger
        If UCase(mvarHeadings(Index)(lngCount)) = datGeneral.ConvertNumberForDisplay(Format(strValue, mstrFormat(Index))) Then
          GetGroupNumber = lngCount
          Exit For
        End If

      Case Else
        'MH20021018 Fault 4532 & 4533
        If LCase(mvarHeadings(Index)(lngCount)) = LCase(Trim(strValue)) Then
          GetGroupNumber = lngCount
          Exit For
        End If

      End Select

    Next
  
  Else    'Numeric ranges...
    
    dblValue = Val(strValue)
    If strValue = vbNullString Then
      GetGroupNumber = 0
      Exit Function
    ElseIf dblValue < mdblMin(Index) Then
      GetGroupNumber = 1
      Exit Function
    End If
  
    lngCount = 1
    For dblLoop = mdblMin(Index) To mdblMax(Index) Step mdblStep(Index)
      lngCount = lngCount + 1
      'If dblValue >= dblLoop And dblValue <= dblLoop + mdblStep(Index) Then
      If dblValue >= dblLoop And dblValue < dblLoop + mdblStep(Index) Then
        GetGroupNumber = lngCount
        Exit Function
      End If
    Next
    GetGroupNumber = lngCount + 1

  End If

'If GetGroupNumber = -1 Then
'  GetGroupNumber = 0
'  'Stop
'  'mstrErrorMessage = "Error grouping data <" & strValue & ">"
'  'fOK = False
'End If

Exit Function

LocalErr:
  mstrErrorMessage = "Error grouping data <" & strValue & ">"
  fOK = False
  
End Function


Public Sub RetreiveDefinition(lngCrossTabID As Long)

  On Error GoTo LocalErr

  Dim rsCrossTabDefinition As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT ASRSysCrossTab.*, " & _
           "'TableName' = ASRSysTables.TableName, " & _
           "'RecordDescExprID' = ASRSysTables.RecordDescExprID, " & _
           "'IntersectionColName' = ASRSysColumns.ColumnName, " & _
           "'IntersectionDecimals' = ASRSysColumns.Decimals " & _
           "FROM ASRSysCrossTab " & _
           "JOIN ASRSysTables ON ASRSysCrossTab.TableID = ASRSysTables.TableID " & _
           "LEFT OUTER JOIN ASRSysColumns ON ASRSysCrossTab.IntersectionColID = ASRSysColumns.ColumnID " & _
           "WHERE CrossTabID = " & CStr(lngCrossTabID)
  
  Set rsCrossTabDefinition = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  If rsCrossTabDefinition.BOF And rsCrossTabDefinition.EOF Then
    Set rsCrossTabDefinition = Nothing
    mstrErrorMessage = "This definition has been deleted by another user"
    fOK = False
    Exit Sub
  End If


  With rsCrossTabDefinition
    'NHRD09042002 Fault 3322 - Code Added
    mblnChkPicklistFilter = !PrintFilterHeader
    mlngFilterID = !FilterID
    mlngPicklistID = !PicklistID
    
    mlngBaseTableID = !TableID
    mstrBaseTable = !TableName
    mlngRecordDescExprID = !RecordDescExprID
    mstrCrossTabName = !Name
    
    mlngIntersectionType = !IntersectionType
    mblnShowPercentage = !Percentage
    mblnPercentageofPage = !PercentageofPage
    mblnSuppressZeros = !SuppressZeros
    mbThousandSeparators = !ThousandSeparators
    
    mlngColID(HOR) = !HorizontalColID
    mdblMin(HOR) = Val(!HorizontalStart)
    mdblMax(HOR) = Val(!HorizontalStop)
    mdblStep(HOR) = Val(!HorizontalStep)
    mstrColName(HOR) = datGeneral.GetColumnName(mlngColID(HOR))
    mlngColDataType(HOR) = datGeneral.GetDataType(mlngBaseTableID, mlngColID(HOR))
    mstrFormat(HOR) = GetFormat(mlngColID(HOR))
    
    
    mlngColID(VER) = !VerticalColID
    mdblMin(VER) = Val(!VerticalStart)
    mdblMax(VER) = Val(!VerticalStop)
    mdblStep(VER) = Val(!VerticalStep)
    mstrColName(VER) = datGeneral.GetColumnName(mlngColID(VER))
    mlngColDataType(VER) = datGeneral.GetDataType(mlngBaseTableID, mlngColID(VER))
    mstrFormat(VER) = GetFormat(mlngColID(VER))
    
    
    mlngColID(PGB) = !PageBreakColID
    mblnPageBreak = (mlngColID(PGB) > 0)
    If mblnPageBreak Then
      mstrColName(PGB) = datGeneral.GetColumnName(mlngColID(PGB))
      mlngColDataType(PGB) = datGeneral.GetDataType(mlngBaseTableID, mlngColID(PGB))
      mstrFormat(PGB) = GetFormat(mlngColID(PGB))
      mdblMin(PGB) = Val(!PageBreakStart)
      mdblMax(PGB) = Val(!PageBreakStop)
      mdblStep(PGB) = Val(!PageBreakStep)
    End If
    
    
    mblnIntersection = (!IntersectionColID > 0)
    If mblnIntersection Then
      mlngColID(INS) = !IntersectionColID
      mstrColName(INS) = !IntersectionColName
      'mstrIntersectionMask = String$(20, "#") & "0." & _
                             String$(CLng(!IntersectionDecimals), "0")
'      mstrIntersectionMask = String$(20, "#") & "0"
      mstrIntersectionMask = GetFormat(!IntersectionColID)
      
      mlngIntersectionDecimals = !IntersectionDecimals
      'If CLng(mlngIntersectionDecimals) > 0 Then
        'mstrIntersectionMask = mstrIntersectionMask & _
            UI.GetSystemDecimalSeparator & String$(CLng(!IntersectionDecimals), "0")
      '  mstrIntersectionMask = mstrIntersectionMask & _
            "." & String$(CLng(mlngIntersectionDecimals), "0")
      'End If
    End If
    
    mstrPicklistFilter = GetPicklistFilterSelect(!PicklistID, !FilterID)
    
    
    mlngOutputFormat = !OutputFormat
    mblnOutputScreen = !OutputScreen
    mblnOutputPrinter = !OutputPrinter
    mstrOutputPrinterName = !OutputPrinterName
    mblnOutputSave = !OutputSave
    mlngOutputSaveExisting = !OutputSaveExisting
    mblnOutputEmail = !OutputEmail
    mlngOutputEmailAddr = !OutputEmailAddr
    mstrOutputEmailSubject = !OutputEmailSubject
    mstrOutputEmailAttachAs = IIf(IsNull(!OutputEmailAttachAs), vbNullString, !OutputEmailAttachAs)
    mstrOutputFilename = !OutputFilename

    mblnPreviewOnScreen = (!OutputPreview Or (mlngOutputFormat = fmtDataOnly And mblnOutputScreen))
    
    
    If fOK = False Then
      Exit Sub
    End If
    
  End With

  If Not IsRecordSelectionValid Then
    fOK = False
    Exit Sub
  End If

TidyAndExit:
  Set rsCrossTabDefinition = Nothing

Exit Sub

LocalErr:
  mstrErrorMessage = "Error reading Cross Tab definition"
  fOK = False
  Resume TidyAndExit

End Sub


Private Function GetFormat(lngColumnID As Long) As String

  Dim rsTemp As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT DataType, Size, Decimals, Use1000Separator FROM ASRSysColumns Where ColumnID = " & CStr(lngColumnID)
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  
  Select Case rsTemp!DataType
  Case sqlNumeric
  
    'JDM - 18/06/03 - Fault 6073 - added isnull
    If IIf(IsNull(rsTemp!Use1000Separator), False, rsTemp!Use1000Separator) Then
      If rsTemp!Decimals > 0 Then
        GetFormat = "#,0." & String(rsTemp!Decimals, "0")
      Else
        GetFormat = "#,0"
      End If
    Else
    
      GetFormat = String$(rsTemp!Size - 1, "#") & "0"
      If rsTemp!Decimals > 0 Then
        'GetFormat = GetFormat & UI.GetSystemDecimalSeparator & String$(rsTemp!Decimals, "0")
        GetFormat = GetFormat & "." & String$(rsTemp!Decimals, "0")
      End If
    End If

    
  Case sqlInteger
    GetFormat = String$(9, "#") & "0"
    
  End Select
  
  Set rsTemp = Nothing

End Function


Public Sub GetHeadingsAndSearches()

  Dim strHeading() As String
  Dim strSearch() As String
  Dim lngLoop As Long
  
  
  On Error GoTo LocalErr
  
  For lngLoop = 0 To 2
    
    ReDim strHeading(0) As String
    ReDim strSearch(0) As String
    
    If lngLoop = 2 And mblnPageBreak = False Then
      'When no page break field is specified
      strHeading(0) = "<None>"
    
    Else
      GetHeadingsAndSearchesForColumns lngLoop, strHeading(), strSearch()

    End If
    
    'Store each array in an array of variants (an array in an array!)
    mvarHeadings(lngLoop) = strHeading
    mvarSearches(lngLoop) = strSearch

  Next

Exit Sub

LocalErr:
  mstrErrorMessage = "Error building headings and search arrays"
  fOK = False
  
End Sub


Private Sub GetHeadingsAndSearchesForColumns(lngLoop As Long, strHeading() As String, strSearch() As String)

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strFieldValue As String
  Dim lngCount As Long
  Dim dblGroup As Double
  Dim dblGroupMax As Double
  Dim dblUnit As Double
  Dim strColumnName As String
  Dim strWhereEmpty As String


  strColumnName = Choose(lngLoop + 1, "Hor", "Ver", "Pgb")
  
  If mdblMin(lngLoop) = 0 And mdblMax(lngLoop) = 0 Then
    
    lngCount = 0

    strWhereEmpty = strColumnName & " IS NULL"
    If mlngColDataType(lngLoop) <> sqlNumeric And _
       mlngColDataType(lngLoop) <> sqlInteger And _
       mlngColDataType(lngLoop) <> sqlBoolean Then
      strWhereEmpty = strWhereEmpty & _
          " OR RTrim(" & strColumnName & ") = ''"
    End If

    'MH20010327 Always add <empty> to see if that helps problems
    '''Check for Empty
    ''strSQL = "SELECT DISTINCT " & strColumnName & _
    ''         " FROM " & mstrTempTableName & _
    ''         " WHERE " & strWhereEmpty
    ''Set rsTemp = datData.OpenRecordset(strSQL, adOpenDynamic, adLockReadOnly)
    
    ''If Not (rsTemp.BOF And rsTemp.EOF) Then

    ' Don't put in empty clauses if we're running an absence breakdown
    If mlngCrossTabType <> cttAbsenceBreakdown Then
        ReDim Preserve strHeading(lngCount) As String
        ReDim Preserve strSearch(lngCount) As String
        strHeading(lngCount) = "<Empty>"
        strSearch(lngCount) = strWhereEmpty
        lngCount = lngCount + 1
    End If
    ''End If
    
    If mlngCrossTabType = cttAbsenceBreakdown And strColumnName = "Hor" Then
     strSQL = "SELECT DISTINCT " & strColumnName & " ,Day_Number, DisplayOrder" & _
             " FROM " & mstrTempTableName & _
             " ORDER BY DisplayOrder"
    Else
      strSQL = "SELECT DISTINCT " & strColumnName & _
             " FROM " & mstrTempTableName & _
             " ORDER BY " & strColumnName
    End If

    Set rsTemp = datData.OpenRecordset(strSQL, adOpenDynamic, adLockReadOnly)

    With rsTemp
      .MoveFirst
      Do While Not .EOF And Not gobjProgress.Cancelled
        
        'MH20010213 Had to make this change so that working pattern would work
        'The field has spaces at the begining
        'strFieldValue = IIf(IsNull(.Fields(0).Value), vbNullString, Trim(.Fields(0).Value))
        'If strFieldValue <> vbNullString Then
        strFieldValue = IIf(IsNull(.Fields(0).Value), vbNullString, .Fields(0).Value)

        If Trim(strFieldValue) <> vbNullString Or mlngCrossTabType = cttAbsenceBreakdown Then
          ReDim Preserve strHeading(lngCount) As String
          ReDim Preserve strSearch(lngCount) As String
          
          Select Case mlngColDataType(lngLoop)
          Case sqlDate
            strHeading(lngCount) = Format(.Fields(0).Value, DateFormat)
            strSearch(lngCount) = strColumnName & " = '" & FormatDateSQL(.Fields(0).Value) & "'"

          Case sqlBoolean
            strHeading(lngCount) = IIf(.Fields(0).Value, "True", "False")
            strSearch(lngCount) = strColumnName & " = " & IIf(.Fields(0).Value, "1", "0")

          Case sqlNumeric, sqlInteger
            'strHeading(lngCount) = .Fields(0).Value
            'strSearch(lngCount) = strColumnName & " = " & .Fields(0).Value

            strHeading(lngCount) = datGeneral.ConvertNumberForDisplay(Format(.Fields(0).Value, mstrFormat(lngLoop)))
            strSearch(lngCount) = strColumnName & " = " & datGeneral.ConvertNumberForSQL(.Fields(0).Value)

          Case Else
            strHeading(lngCount) = FormatString(.Fields(0).Value)
            strSearch(lngCount) = "ltrim(rtrim(" & strColumnName & ")) = '" & Trim(Replace(strFieldValue, "'", "''")) & "'"

          End Select

          lngCount = lngCount + 1

'        Else
'
'          ReDim Preserve strHeading(lngCount) As String
'          ReDim Preserve strSearch(lngCount) As String
'          strHeading(lngCount) = "<Empty>"
'          strSearch(lngCount) = strColumnName & " IS NULL"
'
'          If mlngColDataType(lngLoop) <> sqlNumeric And _
'             mlngColDataType(lngLoop) <> sqlInteger And _
'             mlngColDataType(lngLoop) <> sqlBoolean Then
'            strSearch(lngCount) = strSearch(lngCount) & _
'                " OR RTrim(" & strColumnName & ") = ''"
'          End If
'
'          lngCount = lngCount + 1
'
        End If

        .MoveNext
      Loop
    End With

  Else
    
    ReDim Preserve strHeading(1) As String
    ReDim Preserve strSearch(1) As String
    
    'First element of range for null values...
    strHeading(0) = "<Empty>"
    strSearch(0) = strColumnName & " IS NULL"

    'Second element of range for those less than minimum value of range...
    strHeading(1) = " < " & datGeneral.ConvertNumberForDisplay(Format(mdblMin(lngLoop), mstrFormat(lngLoop)))
    'MH20010411 Fault 1978 Convert to int stops overflow error !
    'strSearch(1) = "Convert(int," & strColumnName & ") < " & datGeneral.ConvertNumberForSQL(mdblMin(lngLoop))
    strSearch(1) = "Convert(float," & strColumnName & ") < " & datGeneral.ConvertNumberForSQL(mdblMin(lngLoop))

    dblUnit = GetSmallestUnit(lngLoop)

    If mdblStep(lngLoop) = 0 Then
      mstrErrorMessage = "Step value for " & strColumnName & " column cannot be zero"
      fOK = False
      Exit Sub
    End If

    lngCount = 2
    For dblGroup = mdblMin(lngLoop) To mdblMax(lngLoop) Step mdblStep(lngLoop)
      ReDim Preserve strHeading(lngCount) As String
      ReDim Preserve strSearch(lngCount) As String
      dblGroupMax = dblGroup + mdblStep(lngLoop) - dblUnit
      strHeading(lngCount) = " " & datGeneral.ConvertNumberForDisplay(Format(dblGroup, mstrFormat(lngLoop))) & _
                             IIf(dblGroupMax <> dblGroup, " - " & datGeneral.ConvertNumberForDisplay(Format(dblGroupMax, mstrFormat(lngLoop))), "")
      'MH20010411 Fault 1978 Convert to int stops overflow error !
      'strSearch(lngCount) = "Convert(int," & strColumnName & ") BETWEEN " & _
                            datGeneral.ConvertNumberForSQL(dblGroup) & " AND " & datGeneral.ConvertNumberForSQL(dblGroupMax)
      strSearch(lngCount) = "Convert(float," & strColumnName & ") BETWEEN " & _
                            datGeneral.ConvertNumberForSQL(dblGroup) & " AND " & datGeneral.ConvertNumberForSQL(dblGroupMax)

      lngCount = lngCount + 1
    Next

    ReDim Preserve strHeading(lngCount) As String
    ReDim Preserve strSearch(lngCount) As String
    'Last element of range for those more than maximum value of range...
    strHeading(lngCount) = "> " & datGeneral.ConvertNumberForDisplay(Format(dblGroup - dblUnit, mstrFormat(lngLoop)))
    'MH20010411 Fault 1978 Convert to int stops overflow error !
    'strSearch(lngCount) = "Convert(int," & strColumnName & ") > " & datGeneral.ConvertNumberForSQL(dblGroup - dblUnit)
    strSearch(lngCount) = "Convert(float," & strColumnName & ") > " & datGeneral.ConvertNumberForSQL(dblGroup - dblUnit)
    
    lngCount = lngCount + 1
  End If

End Sub


Private Function GetSmallestUnit(lngLoop As Long) As Double
  
  'e.g. mstrFormat(lngLoop) = 0.0,   GetSmallestUnit = 0.1
  '     mstrFormat(lngLoop) = 0.000, GetSmallestUnit = 0.001
  '     mstrFormat(lngLoop) = #0,    GetSmallestUnit = 1
  '     mstrFormat(lngLoop) = #0.00, GetSmallestUnit = 0.01
  
'  Dim intFound As Integer
'
'  intFound = InStr(mstrFormat(lngLoop), UI.GetSystemDecimalSeparator)
'  If intFound > 0 Then
'    GetSmallestUnit = Mid$(mstrFormat(lngLoop), intFound, Len(mstrFormat(lngLoop)) - intFound) & "1"
'  Else
'    GetSmallestUnit = 1
'  End If

  Dim strTemp As String
  Dim intFound As Integer
  
  intFound = InStr(mstrFormat(lngLoop), ".")
  If intFound > 0 Then
    strTemp = Mid$(mstrFormat(lngLoop), intFound, Len(mstrFormat(lngLoop)) - intFound) & "1"
    GetSmallestUnit = datGeneral.ConvertNumberForDisplay(strTemp)
  Else
    GetSmallestUnit = 1
  End If

End Function


Private Sub BuildDataArrays()

  Dim strTempValue As String
  Dim dblThisIntersectionVal As Double
  
  Dim lngCol As Long
  Dim lngRow As Long
  Dim lngPage As Long
  Dim lngNumCols As Long
  Dim lngNumRows As Long
  Dim lngNumPages As Long

  On Error GoTo LocalErr

  lngNumCols = UBound(mvarHeadings(0))
  lngNumRows = UBound(mvarHeadings(1))
  lngNumPages = IIf(mblnPageBreak, UBound(mvarHeadings(2)), 0)

  ReDim mdblDataArray(lngNumCols, lngNumRows, lngNumPages, 4) As Double
  ReDim mdblHorTotal(lngNumCols, lngNumPages, 4) As Double
  ReDim mdblVerTotal(lngNumRows, lngNumPages, 4) As Double
  ReDim mdblPgbTotal(lngNumCols, lngNumRows + 1, 4) As Double   '+1 for totals !
  ReDim mdblPageTotal(lngNumPages, 4) As Double
  ReDim mdblGrandTotal(4) As Double

  With rsCrossTabData

    .MoveFirst
    Do While Not .EOF And Not gobjProgress.Cancelled

      strTempValue = IIf(Not IsNull(.Fields("HOR")), .Fields("HOR"), vbNullString)
      lngCol = GetGroupNumber(strTempValue, HOR)

      If Not IsNull(.Fields("VER")) Then
        strTempValue = FormatString(.Fields("VER"))
      Else
        strTempValue = vbNullString
      End If
      
      lngRow = GetGroupNumber(strTempValue, VER)

      If mblnPageBreak Then
        strTempValue = IIf(Not IsNull(.Fields("PGB")), .Fields("PGB"), vbNullString)
        lngPage = GetGroupNumber(strTempValue, PGB)
      Else
        lngPage = 0
      End If

    'Count
      mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT) = mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT) + 1
      mdblHorTotal(lngCol, lngPage, TYPECOUNT) = mdblHorTotal(lngCol, lngPage, TYPECOUNT) + 1
      mdblVerTotal(lngRow, lngPage, TYPECOUNT) = mdblVerTotal(lngRow, lngPage, TYPECOUNT) + 1
      mdblPgbTotal(lngCol, lngRow, TYPECOUNT) = mdblPgbTotal(lngCol, lngRow, TYPECOUNT) + 1
      mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT) = mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT) + 1
      mdblPageTotal(lngPage, TYPECOUNT) = mdblPageTotal(lngPage, TYPECOUNT) + 1
      mdblGrandTotal(TYPECOUNT) = mdblGrandTotal(TYPECOUNT) + 1

      'If mblnIntersection And IsNull(.Fields(.Fields.Count - 1)) = False Then
      If mblnIntersection Then

        If IsNull(.Fields("INS")) Then
          dblThisIntersectionVal = 0
        Else
          dblThisIntersectionVal = Val(datGeneral.ConvertNumberForSQL(.Fields("INS")))
        End If
        
      'Total
        mdblDataArray(lngCol, lngRow, lngPage, TYPETOTAL) = mdblDataArray(lngCol, lngRow, lngPage, TYPETOTAL) + dblThisIntersectionVal
        mdblHorTotal(lngCol, lngPage, TYPETOTAL) = mdblHorTotal(lngCol, lngPage, TYPETOTAL) + dblThisIntersectionVal
        mdblVerTotal(lngRow, lngPage, TYPETOTAL) = mdblVerTotal(lngRow, lngPage, TYPETOTAL) + dblThisIntersectionVal
        mdblPgbTotal(lngCol, lngRow, TYPETOTAL) = mdblPgbTotal(lngCol, lngRow, TYPETOTAL) + dblThisIntersectionVal
        mdblPgbTotal(lngCol, lngNumRows + 1, TYPETOTAL) = mdblPgbTotal(lngCol, lngNumRows + 1, TYPETOTAL) + dblThisIntersectionVal
        mdblPageTotal(lngPage, TYPETOTAL) = mdblPageTotal(lngPage, TYPETOTAL) + dblThisIntersectionVal
        mdblGrandTotal(TYPETOTAL) = mdblGrandTotal(TYPETOTAL) + dblThisIntersectionVal
  
      'Average
        mdblDataArray(lngCol, lngRow, lngPage, TYPEAVERAGE) = mdblDataArray(lngCol, lngRow, lngPage, TYPETOTAL) / mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT)
        mdblHorTotal(lngCol, lngPage, TYPEAVERAGE) = mdblHorTotal(lngCol, lngPage, TYPETOTAL) / mdblHorTotal(lngCol, lngPage, TYPECOUNT)
        mdblVerTotal(lngRow, lngPage, TYPEAVERAGE) = mdblVerTotal(lngRow, lngPage, TYPETOTAL) / mdblVerTotal(lngRow, lngPage, TYPECOUNT)
        mdblPgbTotal(lngCol, lngRow, TYPEAVERAGE) = mdblPgbTotal(lngCol, lngRow, TYPETOTAL) / mdblPgbTotal(lngCol, lngRow, TYPECOUNT)
        mdblPgbTotal(lngCol, lngNumRows + 1, TYPEAVERAGE) = mdblPgbTotal(lngCol, lngNumRows + 1, TYPETOTAL) / mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT)
        mdblPageTotal(lngPage, TYPEAVERAGE) = mdblPageTotal(lngPage, TYPETOTAL) / mdblPageTotal(lngPage, TYPECOUNT)
        mdblGrandTotal(TYPEAVERAGE) = mdblGrandTotal(TYPETOTAL) / mdblGrandTotal(TYPECOUNT)
      
      'Minimum
        If dblThisIntersectionVal < mdblDataArray(lngCol, lngRow, lngPage, TYPEMINIMUM) Or mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT) = 1 Then
          mdblDataArray(lngCol, lngRow, lngPage, TYPEMINIMUM) = dblThisIntersectionVal
          If dblThisIntersectionVal < mdblHorTotal(lngCol, lngPage, TYPEMINIMUM) Or mdblHorTotal(lngCol, lngPage, TYPECOUNT) = 1 Then
            mdblHorTotal(lngCol, lngPage, TYPEMINIMUM) = dblThisIntersectionVal
          End If
          If dblThisIntersectionVal < mdblVerTotal(lngRow, lngPage, TYPEMINIMUM) Or mdblVerTotal(lngRow, lngPage, TYPECOUNT) = 1 Then
            mdblVerTotal(lngRow, lngPage, TYPEMINIMUM) = dblThisIntersectionVal
          End If
          If dblThisIntersectionVal < mdblPgbTotal(lngCol, lngRow, TYPEMINIMUM) Or mdblPgbTotal(lngCol, lngRow, TYPECOUNT) = 1 Then
            mdblPgbTotal(lngCol, lngRow, TYPEMINIMUM) = dblThisIntersectionVal
          End If
          If dblThisIntersectionVal < mdblPgbTotal(lngCol, lngNumRows + 1, TYPEMINIMUM) Or mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT) = 1 Then
            mdblPgbTotal(lngCol, lngNumRows + 1, TYPEMINIMUM) = dblThisIntersectionVal
          End If
          If dblThisIntersectionVal < mdblPageTotal(lngPage, TYPEMINIMUM) Or mdblPageTotal(lngPage, TYPECOUNT) = 1 Then
            mdblPageTotal(lngPage, TYPEMINIMUM) = dblThisIntersectionVal
            If dblThisIntersectionVal < mdblGrandTotal(TYPEMINIMUM) Or mdblGrandTotal(TYPECOUNT) = 1 Then
              mdblGrandTotal(TYPEMINIMUM) = dblThisIntersectionVal
            End If
          End If
        End If
  
      'Maximum
        If dblThisIntersectionVal > mdblDataArray(lngCol, lngRow, lngPage, TYPEMAXIMUM) Or mdblDataArray(lngCol, lngRow, lngPage, TYPECOUNT) = 1 Then
          mdblDataArray(lngCol, lngRow, lngPage, TYPEMAXIMUM) = dblThisIntersectionVal
          If dblThisIntersectionVal > mdblHorTotal(lngCol, lngPage, TYPEMAXIMUM) Or mdblHorTotal(lngCol, lngPage, TYPECOUNT) = 1 Then
            mdblHorTotal(lngCol, lngPage, TYPEMAXIMUM) = dblThisIntersectionVal
          End If
          If dblThisIntersectionVal > mdblVerTotal(lngRow, lngPage, TYPEMAXIMUM) Or mdblVerTotal(lngRow, lngPage, TYPECOUNT) = 1 Then
            mdblVerTotal(lngRow, lngPage, TYPEMAXIMUM) = dblThisIntersectionVal
          End If
          If dblThisIntersectionVal > mdblPgbTotal(lngCol, lngRow, TYPEMAXIMUM) Or mdblPgbTotal(lngCol, lngRow, TYPECOUNT) = 1 Then
            mdblPgbTotal(lngCol, lngRow, TYPEMAXIMUM) = dblThisIntersectionVal
          End If
          If dblThisIntersectionVal > mdblPgbTotal(lngCol, lngNumRows + 1, TYPEMAXIMUM) Or mdblPgbTotal(lngCol, lngNumRows + 1, TYPECOUNT) = 1 Then
            mdblPgbTotal(lngCol, lngNumRows + 1, TYPEMAXIMUM) = dblThisIntersectionVal
          End If
          If dblThisIntersectionVal > mdblPageTotal(lngPage, TYPEMAXIMUM) Or mdblPageTotal(lngPage, TYPECOUNT) = 1 Then
            mdblPageTotal(lngPage, TYPEMAXIMUM) = dblThisIntersectionVal
            If dblThisIntersectionVal > mdblGrandTotal(TYPEMAXIMUM) Or mdblGrandTotal(TYPECOUNT) = 1 Then
              mdblGrandTotal(TYPEMAXIMUM) = dblThisIntersectionVal
            End If
          End If
        End If

      End If

      .MoveNext
    Loop

  End With

  Set rsCrossTabData = Nothing

Exit Sub

LocalErr:
  mstrErrorMessage = "Error processing data" & _
      IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString)
  fOK = False

End Sub

Private Sub AbsenceBreakdownBuildDataArrays()

  Dim strTempValue As String
  Dim dblThisIntersectionVal As Double
  
  Dim lngCol As Long
  Dim lngRow As Long
  Dim lngPage As Long
  Dim lngNumCols As Long
  Dim lngNumRows As Long
  Dim lngNumPages As Long


  On Error GoTo LocalErr

  lngNumCols = UBound(mvarHeadings(0))
  lngNumRows = UBound(mvarHeadings(1))
  lngNumPages = IIf(mblnPageBreak, UBound(mvarHeadings(2)), 0)

  ReDim mdblDataArray(lngNumCols, lngNumRows, lngNumPages, 4) As Double
  ReDim mdblHorTotal(lngNumCols, lngNumPages, 4) As Double
  ReDim mdblVerTotal(lngNumRows, lngNumPages, 4) As Double
  ReDim mdblPgbTotal(lngNumCols, lngNumRows + 1, 4) As Double   '+1 for totals !
  ReDim mdblPageTotal(lngNumPages, 4) As Double
  ReDim mdblGrandTotal(4) As Double

  ' Because the stored procedure has run we need to requery the recordset
  rsCrossTabData.Requery

  With rsCrossTabData

    .MoveFirst
    Do While Not .EOF And Not gobjProgress.Cancelled

      strTempValue = IIf(Not IsNull(.Fields("HOR")), .Fields("HOR"), vbNullString)
      lngCol = GetGroupNumber(strTempValue, HOR)

      strTempValue = IIf(Not IsNull(.Fields("VER")), .Fields("VER"), vbNullString)
      lngRow = GetGroupNumber(strTempValue, VER)

      'Count
      mdblDataArray(lngCol, lngRow, 0, TYPECOUNT) = mdblDataArray(lngCol, lngRow, 0, TYPECOUNT) + IIf(Not IsNull(.Fields("VALUE")), .Fields("VALUE"), 143)
      mdblHorTotal(lngCol, 0, TYPECOUNT) = mdblHorTotal(lngCol, 0, TYPECOUNT) + 1
      mdblVerTotal(lngRow, 0, TYPECOUNT) = mdblVerTotal(lngRow, 0, TYPECOUNT) + 1

      mdblDataArray(lngCol, lngRow, 0, TYPETOTAL) = mdblDataArray(lngCol, lngRow, 0, TYPETOTAL) + IIf(Not IsNull(.Fields("VALUE")), .Fields("VALUE"), 143)
      mdblHorTotal(lngCol, 0, TYPETOTAL) = mdblHorTotal(lngCol, 0, TYPETOTAL) + IIf(Not IsNull(.Fields("VALUE")), .Fields("VALUE"), 0)
      mdblVerTotal(lngRow, 0, TYPETOTAL) = mdblVerTotal(lngRow, 0, TYPETOTAL) + IIf(Not IsNull(.Fields("VALUE")), .Fields("VALUE"), 0)

      mdblDataArray(lngCol, lngRow, lngPage, TYPEAVERAGE) = mdblDataArray(lngCol, lngRow, lngPage, TYPEAVERAGE) + IIf(Not IsNull(.Fields("VALUE")), .Fields("VALUE"), 143)
      mdblHorTotal(lngCol, lngPage, TYPEAVERAGE) = mdblHorTotal(lngCol, lngPage, TYPEAVERAGE) + IIf(Not IsNull(.Fields("VALUE")), .Fields("VALUE"), 0)
      mdblVerTotal(lngRow, lngPage, TYPEAVERAGE) = mdblVerTotal(lngRow, lngPage, TYPEAVERAGE) + IIf(Not IsNull(.Fields("VALUE")), .Fields("VALUE"), 0)

      .MoveNext
    Loop

  End With

  Set rsCrossTabData = Nothing

Exit Sub

LocalErr:
  mstrErrorMessage = "Error processing data" & _
      IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString)
  fOK = False

End Sub



Private Sub TurnoverBuildDataArrays()

  Const STAFF As Integer = 0
  Const LEAVERS As Integer = 1
  Const TURNOVER As Integer = 2
  Const lngTYPE As Long = 0
  
  Dim dblThisIntersectionVal As Double
  Dim strTempValue As String
  Dim dblTempValue As Double
  
  Dim lngCol As Long
  Dim lngRow As Long
  Dim lngPage As Long
  Dim lngNumCols As Long
  Dim lngNumRows As Long
  Dim lngNumPages As Long
  
  Dim dtLeavingDate As Date
  Dim blnAtStart As Boolean
  Dim blnLeaver As Boolean
  
  On Error GoTo LocalErr
  
  lngNumCols = 3
  lngNumRows = UBound(mvarHeadings(1))
  lngNumPages = IIf(mblnPageBreak, UBound(mvarHeadings(PGB)), 0)
  
  ReDim mdblDataArray(lngNumCols, lngNumRows, lngNumPages, lngTYPE) As Double
  ReDim mdblHorTotal(lngNumCols, lngNumPages, lngTYPE) As Double
  ReDim mdblVerTotal(lngNumCols, lngNumRows, lngTYPE) As Double
  ReDim mdblPgbTotal(lngNumCols, lngNumRows + 1, lngTYPE) As Double
  ReDim mdblPageTotal(lngNumPages, lngTYPE) As Double

  With rsCrossTabData
  
    .MoveFirst
    Do While Not .EOF

      strTempValue = IIf(Not IsNull(.Fields("VER")), .Fields("VER"), vbNullString)
      lngRow = GetGroupNumber(strTempValue, VER)
    
      If mblnPageBreak Then
        strTempValue = IIf(Not IsNull(.Fields("PGB")), .Fields("PGB"), vbNullString)
        lngPage = GetGroupNumber(strTempValue, PGB)
      Else
        lngPage = 0
      End If

      blnAtStart = True
      If Not IsNull(.Fields("StartDate").Value) Then
        'blnAtStart = (DateDiff("d", .Fields("StartDate").Value, mdtReportStartDate) > 0)
        blnAtStart = (DateDiff("d", .Fields("StartDate").Value, mdtReportStartDate) >= 0)
      End If
      blnLeaver = CheckIfLeaver(.Fields("LeavingDate").Value)

      'If blnAtStart Then
      '  Debug.Print .Fields("ID").Value
      '  Debug.Print blnLeaver
      '  Stop
      'End If

    'Number of Staff column
    '(add 1 if employeed at end of report else add 0.5 if leaving between start and end dates)
      dblTempValue = IIf(blnAtStart, 0.5, 0) + IIf(Not blnLeaver, 0.5, 0)
      
      mdblDataArray(STAFF, lngRow, lngPage, lngTYPE) = mdblDataArray(STAFF, lngRow, lngPage, lngTYPE) + dblTempValue
      mdblHorTotal(STAFF, lngPage, lngTYPE) = mdblHorTotal(STAFF, lngPage, lngTYPE) + dblTempValue
      mdblVerTotal(STAFF, lngRow, lngTYPE) = mdblVerTotal(STAFF, lngRow, lngTYPE) + dblTempValue
      mdblPgbTotal(STAFF, lngRow, lngTYPE) = mdblPgbTotal(STAFF, lngRow, lngTYPE) + dblTempValue
      mdblPgbTotal(STAFF, lngNumRows + 1, lngTYPE) = mdblPgbTotal(STAFF, lngNumRows + 1, lngTYPE) + dblTempValue
      
    'Leavers
      If blnLeaver And (blnAtStart Or mblnIncludeNewStarters) Then
        mdblDataArray(LEAVERS, lngRow, lngPage, lngTYPE) = mdblDataArray(LEAVERS, lngRow, lngPage, lngTYPE) + 1
        mdblHorTotal(LEAVERS, lngPage, lngTYPE) = mdblHorTotal(LEAVERS, lngPage, lngTYPE) + 1
        mdblVerTotal(LEAVERS, lngRow, lngTYPE) = mdblVerTotal(LEAVERS, lngRow, lngTYPE) + 1
        mdblPgbTotal(LEAVERS, lngRow, lngTYPE) = mdblPgbTotal(LEAVERS, lngRow, lngTYPE) + 1
        mdblPgbTotal(LEAVERS, lngNumRows + 1, lngTYPE) = mdblPgbTotal(LEAVERS, lngNumRows + 1, lngTYPE) + 1
      End If

    'Turnover (Turnover = Staff / Leavers)
      If mdblDataArray(STAFF, lngRow, lngPage, lngTYPE) > 0 Then
        mdblDataArray(TURNOVER, lngRow, lngPage, lngTYPE) = _
          mdblDataArray(LEAVERS, lngRow, lngPage, lngTYPE) / mdblDataArray(STAFF, lngRow, lngPage, lngTYPE)
      End If
      If mdblHorTotal(STAFF, lngPage, lngTYPE) > 0 Then
        mdblHorTotal(TURNOVER, lngPage, lngTYPE) = _
          mdblHorTotal(LEAVERS, lngPage, lngTYPE) / mdblHorTotal(STAFF, lngPage, lngTYPE)
      End If
      If mdblVerTotal(STAFF, lngRow, lngTYPE) > 0 Then
        mdblVerTotal(TURNOVER, lngRow, lngTYPE) = _
          mdblVerTotal(LEAVERS, lngRow, lngTYPE) / mdblVerTotal(STAFF, lngRow, lngTYPE)
      End If
      If mdblPgbTotal(STAFF, lngRow, lngTYPE) > 0 Then
        mdblPgbTotal(TURNOVER, lngRow, lngTYPE) = _
          mdblPgbTotal(LEAVERS, lngRow, lngTYPE) / mdblPgbTotal(STAFF, lngRow, lngTYPE)
      End If
      If mdblPgbTotal(STAFF, lngNumRows + 1, lngTYPE) > 0 Then
        mdblPgbTotal(TURNOVER, lngNumRows + 1, lngTYPE) = _
          mdblPgbTotal(LEAVERS, lngNumRows + 1, lngTYPE) / mdblPgbTotal(STAFF, lngNumRows + 1, lngTYPE)
      End If

      .MoveNext
    Loop
  
  End With

  Set rsCrossTabData = Nothing

Exit Sub

LocalErr:
  mstrErrorMessage = "Error processing data" & _
      IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString)
  fOK = False

End Sub


Private Sub StabilityBuildDataArrays()

  Const ONEYEARAGO As Integer = 0
  Const ONEYEARSERVICE As Integer = 1
  Const STABILITY As Integer = 2
  Const lngTYPE As Long = 0
  
  Dim dblThisIntersectionVal As Double
  Dim strTempValue As String
  
  Dim lngCol As Long
  Dim lngRow As Long
  Dim lngPage As Long
  Dim lngNumCols As Long
  Dim lngNumRows As Long
  Dim lngNumPages As Long
  
  Dim dtLeavingDate As Date
  Dim blnLeaver As Boolean
  
  On Error GoTo LocalErr
  
  lngNumCols = 3
  lngNumRows = UBound(mvarHeadings(1))
  lngNumPages = IIf(mblnPageBreak, UBound(mvarHeadings(PGB)), 0)
  
  ReDim mdblDataArray(lngNumCols, lngNumRows, lngNumPages, lngTYPE) As Double
  ReDim mdblHorTotal(lngNumCols, lngNumPages, lngTYPE) As Double
  ReDim mdblVerTotal(lngNumCols, lngNumRows, lngTYPE) As Double
  ReDim mdblPgbTotal(lngNumCols, lngNumRows + 1, lngTYPE) As Double
  ReDim mdblPageTotal(lngNumPages, lngTYPE) As Double

  With rsCrossTabData
  
    .MoveFirst
    Do While Not .EOF

      strTempValue = IIf(Not IsNull(.Fields("VER")), .Fields("VER"), vbNullString)
      lngRow = GetGroupNumber(strTempValue, VER)
    
      If mblnPageBreak Then
        strTempValue = IIf(Not IsNull(.Fields("PGB")), .Fields("PGB"), vbNullString)
        lngPage = GetGroupNumber(strTempValue, PGB)
      Else
        lngPage = 0
      End If

      blnLeaver = CheckIfLeaver(.Fields("LeavingDate").Value)
      
    
    'Number of Staff column a year ago
    '(add 1 if employeed at end of report else add 0.5 if leaving between start and end dates)
      mdblDataArray(ONEYEARAGO, lngRow, lngPage, lngTYPE) = mdblDataArray(ONEYEARAGO, lngRow, lngPage, lngTYPE) + 1
      mdblHorTotal(ONEYEARAGO, lngPage, lngTYPE) = mdblHorTotal(ONEYEARAGO, lngPage, lngTYPE) + 1
      mdblVerTotal(ONEYEARAGO, lngRow, lngTYPE) = mdblVerTotal(ONEYEARAGO, lngRow, lngTYPE) + 1
      mdblPgbTotal(ONEYEARAGO, lngRow, lngTYPE) = mdblPgbTotal(ONEYEARAGO, lngRow, lngTYPE) + 1
      mdblPgbTotal(ONEYEARAGO, lngNumRows + 1, lngTYPE) = mdblPgbTotal(ONEYEARAGO, lngNumRows + 1, lngTYPE) + 1
      
    'Leavers
      mdblDataArray(ONEYEARSERVICE, lngRow, lngPage, lngTYPE) = mdblDataArray(ONEYEARSERVICE, lngRow, lngPage, lngTYPE) + _
          IIf(blnLeaver, 0, 1)
      mdblHorTotal(ONEYEARSERVICE, lngPage, lngTYPE) = mdblHorTotal(ONEYEARSERVICE, lngPage, lngTYPE) + _
          IIf(blnLeaver, 0, 1)
      mdblVerTotal(ONEYEARSERVICE, lngRow, lngTYPE) = mdblVerTotal(ONEYEARSERVICE, lngRow, lngTYPE) + _
          IIf(blnLeaver, 0, 1)
      mdblPgbTotal(ONEYEARSERVICE, lngRow, lngTYPE) = mdblPgbTotal(ONEYEARSERVICE, lngRow, lngTYPE) + _
          IIf(blnLeaver, 0, 1)
      mdblPgbTotal(ONEYEARSERVICE, lngNumRows + 1, lngTYPE) = mdblPgbTotal(ONEYEARSERVICE, lngNumRows + 1, lngTYPE) + _
          IIf(blnLeaver, 0, 1)

    'Turnover (Turnover = Staff / Leavers)
      If mdblDataArray(ONEYEARAGO, lngRow, lngPage, lngTYPE) > 0 Then
        mdblDataArray(STABILITY, lngRow, lngPage, lngTYPE) = _
          (mdblDataArray(ONEYEARSERVICE, lngRow, lngPage, lngTYPE) / mdblDataArray(ONEYEARAGO, lngRow, lngPage, lngTYPE))
      End If
      If mdblHorTotal(ONEYEARAGO, lngPage, lngTYPE) > 0 Then
        mdblHorTotal(STABILITY, lngPage, lngTYPE) = _
          (mdblHorTotal(ONEYEARSERVICE, lngPage, lngTYPE) / mdblHorTotal(ONEYEARAGO, lngPage, lngTYPE))
      End If
      If mdblVerTotal(ONEYEARAGO, lngRow, lngTYPE) > 0 Then
        mdblVerTotal(STABILITY, lngRow, lngTYPE) = _
          (mdblVerTotal(ONEYEARSERVICE, lngRow, lngTYPE) / mdblVerTotal(ONEYEARAGO, lngRow, lngTYPE))
      End If
      If mdblPgbTotal(ONEYEARAGO, lngRow, lngTYPE) > 0 Then
        mdblPgbTotal(STABILITY, lngRow, lngTYPE) = _
          (mdblPgbTotal(ONEYEARSERVICE, lngRow, lngTYPE) / mdblPgbTotal(ONEYEARAGO, lngRow, lngTYPE))
      End If
      If mdblPgbTotal(ONEYEARAGO, lngNumRows + 1, lngTYPE) > 0 Then
        mdblPgbTotal(STABILITY, lngNumRows + 1, lngTYPE) = _
          (mdblPgbTotal(ONEYEARSERVICE, lngNumRows + 1, lngTYPE) / mdblPgbTotal(ONEYEARAGO, lngNumRows + 1, lngTYPE))
      End If

      .MoveNext
    Loop
  
  End With

  Set rsCrossTabData = Nothing

Exit Sub

LocalErr:
  mstrErrorMessage = "Error processing data" & _
      IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString)
  fOK = False

End Sub


Private Function GetPercentageFactor(lngPage As Long, lngTYPE As Long)

  'mdblPercentageFactor will be used in FORMATCELL, if required
  mdblPercentageFactor = 0
  If chkPercentage.Value = vbChecked Then
    If chkPercentageOfPage = vbChecked Then
      If mdblPageTotal(lngPage, lngTYPE) > 0 Then
        mdblPercentageFactor = 1 / mdblPageTotal(lngPage, lngTYPE)
      End If
    Else
      If mdblGrandTotal(lngTYPE) > 0 Then
        mdblPercentageFactor = 1 / mdblGrandTotal(lngTYPE)
      End If
    End If
  End If

End Function


Public Sub PopulateGrid()
  BuildOutputStrings cboPageBreak.ItemData(Me.cboPageBreak.ListIndex)
  OutputGrid
End Sub


Public Sub BuildOutputStrings(lngSinglePage As Long)

  Const strDelim As String = vbTab
  Dim strTempDelim As String

  Dim lngNumCols As Long
  Dim lngNumRows As Long
  Dim lngNumPages As Long
  
  Dim lngCol As Long
  Dim lngRow As Long
  Dim lngPage As Long
  Dim lngTYPE As Long
  Dim lngPointer As Long
  
  Dim sngAverage As Single
  Dim iAverageColumn As Integer
  
  On Error GoTo LocalErr
  
  lngNumCols = UBound(mvarHeadings(HOR))
  lngNumRows = UBound(mvarHeadings(VER))
  lngNumPages = IIf(mblnPageBreak, UBound(mvarHeadings(PGB)), 0)
  iAverageColumn = lngNumCols - 1

  ' JDM - 22/06/01 - Fault 2476 - Display totals instead
  If mlngCrossTabType <> cttAbsenceBreakdown Then
    lngTYPE = cboType.ItemData(cboType.ListIndex)
  Else
    lngTYPE = TYPETOTAL
  End If
  
  'mdblPercentageFactor will be used in FORMATCELL, if required
  Call GetPercentageFactor(lngSinglePage, lngTYPE)
  
  ReDim mstrOutput(lngNumRows + 2)
  
  'Add First Column details (Vertical headings)
  mstrOutput(0) = strDelim & mstrOutput(0)
  For lngRow = 0 To lngNumRows
    mstrOutput(lngRow + 1) = Trim(mvarHeadings(VER)(lngRow)) & strDelim & mstrOutput(lngRow + 1)
  Next
  mstrOutput(lngNumRows + 2) = _
    IIf(mlngCrossTabType = cttNormal, cboType.Text, "Total") & _
    strDelim & mstrOutput(lngNumRows + 2)
  
  If mblnShowAllPagesTogether Then

    'Now add the main row data
    For lngPage = 0 To lngNumPages
      For lngCol = 0 To lngNumCols

        strTempDelim = IIf(lngCol < lngNumCols Or lngPage < lngNumPages, strDelim, "")

        mstrOutput(0) = mstrOutput(0) & _
          Trim(mvarHeadings(0)(lngCol)) & strTempDelim


        For lngRow = 0 To lngNumRows
          mstrOutput(lngRow + 1) = mstrOutput(lngRow + 1) & _
              FormatCell(mdblDataArray(lngCol, lngRow, lngPage, lngTYPE), lngCol) & strTempDelim
        Next

        mstrOutput(lngNumRows + 2) = mstrOutput(lngNumRows + 2) & _
          FormatCell(mdblHorTotal(lngCol, lngPage, lngTYPE), lngCol) & strTempDelim
      Next
    Next
    
    
    If mblnPageBreak Then
      For lngCol = 0 To lngNumCols
        mstrOutput(0) = mstrOutput(0) & strDelim & _
            Trim(mvarHeadings(0)(lngCol))
        
        For lngRow = 0 To lngNumRows + 1
          mstrOutput(lngRow + 1) = mstrOutput(lngRow + 1) & strDelim & _
              FormatCell(mdblPgbTotal(lngCol, lngRow, lngTYPE), lngCol)
        Next
      Next
    End If

  Else
    'Now add the main row data
    For lngCol = 0 To lngNumCols
      mstrOutput(0) = mstrOutput(0) & Trim(mvarHeadings(0)(lngCol)) & IIf(lngCol <> lngNumCols, strDelim, "")
      For lngRow = 0 To lngNumRows
        mstrOutput(lngRow + 1) = mstrOutput(lngRow + 1) & FormatCell(mdblDataArray(lngCol, lngRow, lngSinglePage, lngTYPE)) & IIf(lngCol <> lngNumCols, strDelim, "")
      Next
      
        ' JDM - 10/09/2003 - Fault 7048 - Make the average column not total up.
        If mlngCrossTabType = cttAbsenceBreakdown And lngCol = iAverageColumn Then
          sngAverage = mdblHorTotal(lngCol - 1, lngSinglePage, TYPETOTAL) / mdblHorTotal(lngCol, lngSinglePage, TYPECOUNT)
          mstrOutput(lngNumRows + 2) = mstrOutput(lngNumRows + 2) & FormatCell(sngAverage) & IIf(lngCol <> lngNumCols, strDelim, "")
        Else
          mstrOutput(lngNumRows + 2) = mstrOutput(lngNumRows + 2) & FormatCell(mdblHorTotal(lngCol, lngSinglePage, lngTYPE)) & IIf(lngCol <> lngNumCols, strDelim, "")
        End If
    Next

    'Add the last column details (Vertical totals)
    'If mlngCrossTabType <> cttAbsenceBreakdown Then
    If mlngCrossTabType = cttNormal Then
      mstrOutput(0) = mstrOutput(0) & strDelim & cboType.Text
      For lngRow = 0 To lngNumRows
        mstrOutput(lngRow + 1) = mstrOutput(lngRow + 1) & strDelim & _
                                   FormatCell(mdblVerTotal(lngRow, lngSinglePage, lngTYPE))
      Next
      mstrOutput(lngNumRows + 2) = mstrOutput(lngNumRows + 2) & strDelim & _
                                   FormatCell(mdblPageTotal(lngSinglePage, lngTYPE))
    End If
  End If

Exit Sub

LocalErr:
  mstrErrorMessage = "Error building output strings"
  fOK = False

End Sub


Private Sub CreateGridColumns()

Dim lngCount As Long
Dim strExprId As Variant

With SSDBGrid1
    .Redraw = False
    .Caption = mstrCrossTabName

    'NHRD09042002 Fault 3322 - Code Added
    If mlngCrossTabType = cttAbsenceBreakdown Then
    'If mlngCrossTabType <> cttNormal Then
        'Define the header of the report
        .Caption = .Caption & _
            " (" & Format(mdtReportStartDate, DateFormat) & _
            " - " & Format(mdtReportEndDate, DateFormat) & ")"

        If mblnChkPicklistFilter Then
            'this will add the picklist or filter name into the SSDBGRID control
            'If for some reason they are both empty then the original caption
            'assigned above will be used
            'If mlngFilterID = 0 And mlngPicklistID = 0 Then
            If mlngPicklistFilterID = 0 Then
                'No picklist or Filter selected so no header to attach
                .Caption = .Caption & " (No Picklist Or Filter Selected)"
            Else
                If mstrPicklistFilterType = "F" Then
                    'Get Filter Name
                    strExprId = GetExprField(mlngPicklistFilterID, "NAME")
                    'Alter heading
                    .Caption = .Caption & "(Base Table Filter : " & strExprId & ")"
                
                ElseIf mstrPicklistFilterType = "P" Then
                    'Get Picklist Name
                    strExprId = GetPickListField(mlngPicklistFilterID, "NAME")
                    'Alter heading
                    .Caption = .Caption & " (Base Table Picklist : " & strExprId & ")"
                End If
            End If
    'end proto
        End If
    Else
      'Define the header of the report
        If mblnChkPicklistFilter Then
            'this will add the picklist or filter name into the SSDBGRID control
            'If for some reason they are both empty then the original caption
            'assigned above will be used
      
            If mlngFilterID = 0 And mlngPicklistID = 0 Then
                'Alter heading
                .Caption = .Caption & " (No Picklist Or Filter Selected)"
            Else
                If mlngFilterID > 0 Then
                    'Get Filter Name
                    strExprId = GetExprField(mlngFilterID, "NAME")
                    'Alter heading
                    .Caption = .Caption & " (Base Table Filter : " & strExprId & ")"
                ElseIf mlngPicklistID > 0 Then
                    'Get Picklist Name
                    strExprId = GetPickListField(mlngPicklistID, "NAME")
                    'Alter heading
                    .Caption = .Caption & " (Base Table Picklist : " & strExprId & ")"
                End If
            End If
            
        End If
    
    End If

    'MH20030902 Fault 6761
    .Caption = Replace(.Caption, "&", "&&")


    'Add column for row headings
    .Columns.RemoveAll
    .Columns.Add .Columns.Count
    With .Columns(.Columns.Count - 1)
      .Caption = vbNullString
      .Locked = True
      .Style = ssStyleButton
      .ButtonsAlways = True
      .BackColor = vbButtonFace   'required for printing !
    End With
    
    'Add main columns
    For lngCount = 0 To UBound(mvarHeadings(HOR))
      .Columns.Add .Columns.Count
      With .Columns(.Columns.Count - 1)
        .Caption = Trim(mvarHeadings(HOR)(lngCount))
        .Locked = True
        .Alignment = ssCaptionAlignmentRight
        .CaptionAlignment = ssColCapAlignCenter
        
        If mlngCrossTabType = cttAbsenceBreakdown Then
          .Width = 1000
        End If
      End With
    Next
    
    'Add total column
    If Not mlngCrossTabType = cttAbsenceBreakdown Then
      .Columns.Add .Columns.Count
      With .Columns(.Columns.Count - 1)
        .Caption = cboType.Text
        .Locked = True
        .Alignment = ssCaptionAlignmentRight
        .CaptionAlignment = ssColCapAlignCenter
        If mlngCrossTabType = cttAbsenceBreakdown Then
          .Width = 1000
        End If
      End With
    End If
    
    .SplitterPos = 1
    .SplitterVisible = False
    .Redraw = True
End With

End Sub

Private Sub OutputGrid(Optional blnRestoreCursorPosition As Boolean = True)

  Dim lngCount As Long
  Dim lngTempRow As Long
  Dim lngFirstRow As Long
  
  Dim lngTempCol As Long
  Dim lngVisibleCols As Long
  
  Dim lngTempGrp As Long
  Dim lngVisibleGrps As Long
  Dim blnRedraw As Boolean

  With SSDBGrid1

    lngTempRow = IIf(.Row > 0, .Row, 0)
    lngFirstRow = .AddItemRowIndex(.FirstRow)
    
    If mblnShowAllPagesTogether Then
      lngVisibleGrps = .VisibleGrps
      lngTempGrp = IIf(.Grp > 1, .Grp, 1)
    End If
    lngVisibleCols = .VisibleCols
    lngTempCol = IIf(.Col > 1, .Col, 1)
    
    blnRedraw = .Redraw
    If blnRedraw And .Visible Then
      .Redraw = False
    End If
    
    .RemoveAll
    If Not mblnShowAllPagesTogether Then
      .Columns(.Columns.Count - 1).Caption = cboType.Text
    End If
    For lngCount = 1 To UBound(mstrOutput)
      .AddItem mstrOutput(lngCount)
    Next

    On Error Resume Next
    
    'If blnRestoreCursorPosition Then
    If (blnRestoreCursorPosition And Me.Visible) Or Not mblnPreviewOnScreen Then

      .FirstRow = .AddItemBookmark(lngFirstRow)
      .Row = lngTempRow
      '.Redraw = True

      If mlngCrossTabType = cttNormal Then
        lngCount = lngTempCol - Int(lngVisibleCols / 2)
        If lngCount < 1 Then
          .LeftCol = 1
        ElseIf lngCount > .Cols - Int(lngVisibleCols / 2) Then
          .LeftCol = .Cols - lngVisibleCols
        Else
          .LeftCol = lngCount
        End If
      Else
        lngCount = lngTempGrp - Int(lngVisibleGrps / 2) + 1
        If lngCount < 0 Then
          .LeftGrp = 0
        ElseIf lngCount > .Groups.Count - Int(lngVisibleGrps / 2) Then
          .LeftGrp = .Groups.Count - lngVisibleGrps
        Else
          .LeftGrp = lngCount
        End If
        .Grp = lngTempCol
      End If
      .Col = lngTempCol

    End If

    'MH20030924 Faults 7046 & 7047
    'If blnRedraw Then
    If blnRedraw Or gblnBatchMode Then
      .Redraw = True
    End If

  End With

End Sub


Private Function FormatCell(ByVal dblCellValue As Double, Optional lngHOR As Long) As String

  Dim strMask As String
  
  On Error GoTo LocalErr
  
  strMask = vbNullString
  FormatCell = vbNullString
  
  
  If dblCellValue <> 0 Or chkSuppressZeros <> vbChecked Then

    If mlngCrossTabType <> cttNormal Then
 
      If mlngCrossTabType = cttAbsenceBreakdown Then
        'NHRD22092004 Fault 7999 Changed it so doubles would show two decimal places
        'strMask = String$(20, "#") & "0.0#"
        strMask = String$(20, "#") & "0.00"
      Else
        strMask = String$(20, "#") & "0"

        If lngHOR = 2 Then
          'strMask = String$(20, "#") & "0" & UI.GetSystemDecimalSeparator & "00%"
          strMask = String$(20, "#") & "0.00%"
        ElseIf lngHOR = 0 And mlngCrossTabType = cttTurnover Then
          'strMask = String$(20, "#") & "0" & UI.GetSystemDecimalSeparator & "0"
          strMask = String$(20, "#") & "0.0"
        End If
      End If
    
    Else
      
      ' 1000 seperators
      'strMask = String$(20, "#") & "0"
      strMask = IIf(chkThousandSeparators.Value = vbChecked, "#,0", "#0")
      
      If chkPercentage = vbChecked Then
        'If percentage
        dblCellValue = dblCellValue * mdblPercentageFactor
        'strMask = strMask & UI.GetSystemDecimalSeparator & "00%"
        strMask = strMask & ".00%"
      
      ElseIf cboType.ItemData(cboType.ListIndex) > 0 Then
        'if not count then
        'value should be displayed as per field definition
        'strMask = mstrIntersectionMask
        If mlngIntersectionDecimals > 0 Then
          strMask = strMask & "." & String(mlngIntersectionDecimals, "0")
        End If

      End If
  
    End If

    If strMask <> vbNullString Then
      FormatCell = Format(dblCellValue, strMask)
    End If

  End If


Exit Function

LocalErr:
  mstrErrorMessage = "Error formatting data"
  fOK = False

End Function


Private Sub Form_Load()

  Set frmBreakDown = New frmCrossTabCellBreakDown

  mlngMinFormWidth = Me.Width
  mlngMinFormHeight = Me.Height

  ' Set height and width to last saved. Form is centred on screen
  Me.Width = GetPCSetting(gsDatabaseName & "\CrossTab", "Width", mlngMinFormWidth)
'  If Me.Width > Screen.Width Then
'    Me.Width = Screen.Width
'  End If

  Me.Height = GetPCSetting(gsDatabaseName & "\CrossTab", "Height", mlngMinFormHeight)
'  If Me.Height > Screen.Height Then
'    Me.Height = Screen.Height
'  End If

 Hook Me.hWnd, mlngMinFormWidth, mlngMinFormHeight

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Screen.MousePointer <> vbDefault Then
    Screen.MousePointer = vbDefault
  End If

End Sub

Private Sub Form_Resize()

  Const lngGap As Long = 100
  Dim lngLeft As Long
  Dim lngTop As Long
  Dim lngWidth As Long
  Dim lngHeight As Long

  'JPD 20030908 Fault 5756
  DisplayApplication
  
  If Me.WindowState = vbMinimized Then
    Exit Sub
  End If

  'Check the minimum form size
'  If Me.Width < mlngMinFormWidth Then
'    Me.Width = mlngMinFormWidth
'  End If
'
'  If Me.Height < mlngMinFormHeight Then
'    Me.Height = mlngMinFormHeight
'  End If

  
  'Position the command buttons...
  lngTop = Me.ScaleHeight - (cmdOK.Height + lngGap)
  
  lngLeft = Me.ScaleWidth - (cmdOK.Width + lngGap)
  cmdOK.Move lngLeft, lngTop

  lngLeft = lngLeft - (cmdOutput.Width + lngGap)
  cmdOutput.Move lngLeft, lngTop


  If mlngCrossTabType = cttNormal Then
    'Position the frames...
    lngTop = lngTop - (fraIntersection.Height + lngGap)
    fraIntersection.Move lngGap, lngTop

    lngLeft = lngGap + fraIntersection.Width + lngGap
    lngWidth = Me.ScaleWidth - (lngLeft + lngGap)

    'MH20030918 Fault 4522
    If lngWidth < (txtPageBreakCol.Left + txtPageBreakCol.Width + lngGap) Then
      lngWidth = txtPageBreakCol.Left + txtPageBreakCol.Width + lngGap
    End If

    fraPage.Move lngLeft, lngTop, lngWidth

  Else
    
    'fraIntersection.Visible = False
    Me.lblColumn.Visible = False
    Me.chkPercentage.Visible = False
    Me.chkPercentageOfPage.Visible = False
    Me.lblType.Top = Me.lblColumn.Top
    Me.chkSuppressZeros.Top = Me.lblColumn.Top
    Me.txtIntersectionCol.Text = mstrCrossTabName
    Me.cboType.Visible = False

    fraIntersection.Height = Me.cboType.Height + (lngGap * 4)
    lngTop = lngTop - (fraIntersection.Height + lngGap)
    lngWidth = (Me.ScaleWidth - (lngGap * 2)) - 20

    fraIntersection.Move lngGap, lngTop, lngWidth
    fraPage.Visible = False
 
  End If


  'Position the Grid...
  lngWidth = (Me.ScaleWidth - (lngGap * 2)) - 20
  lngHeight = (lngTop - (lngGap * 2)) - 20
  SSDBGrid1.Move lngGap, lngGap, lngWidth, lngHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)

  On Local Error Resume Next

  ' Save the window size ready to recall next time user views the event log
  SavePCSetting gsDatabaseName & "\CrossTab", "Height", Me.Height
  SavePCSetting gsDatabaseName & "\CrossTab", "Width", Me.Width

  If Not gblnBatchMode Then
    If Not mcmdUpdateRecDescs Is Nothing Then
      If mcmdUpdateRecDescs.State = adStateExecuting Then
        mcmdUpdateRecDescs.Cancel
      End If
      Set mcmdUpdateRecDescs = Nothing
    End If
  End If

  'TM20020531 Fault 3756
'  If Trim(mstrTempTableName) <> "" Then
'    datData.ExecuteSql "IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & _
'                       "DROP TABLE " & mstrTempTableName
'  End If
  datGeneral.DropUniqueSQLObject mstrTempTableName, 3
  
  Unload frmBreakDown
  Set frmBreakDown = Nothing
  
  Unload frmOutputOptions
  Set frmOutputOptions = Nothing

  Unhook Me.hWnd
  
End Sub


Private Sub PrepareForms()

  On Error GoTo LocalErr
  
  With txtIntersectionCol
    .Text = IIf(mblnIntersection, Replace(mstrColName(INS), "_", " "), "<None>")
    .Enabled = False
    .BackColor = vbButtonFace
    lblColumn.Enabled = mblnIntersection
    lblType.Enabled = mblnIntersection
  End With

  With txtPageBreakCol
    .Text = IIf(mblnPageBreak, Replace(mstrColName(PGB), "_", " "), "<None>")
    .Enabled = False
    .BackColor = vbButtonFace
    lblPageBreak.Enabled = mblnPageBreak
    lblValue.Enabled = mblnPageBreak
  End With
  
  
  chkPercentage.Value = IIf(mblnShowPercentage, vbChecked, vbUnchecked)
  chkPercentageOfPage.Value = IIf(mblnPercentageofPage, vbChecked, vbUnchecked)
  chkPercentageOfPage.Enabled = (mblnPageBreak = True And mblnShowPercentage = True)
  chkSuppressZeros.Value = IIf(mblnSuppressZeros, vbChecked, vbUnchecked)
  chkThousandSeparators.Value = IIf(mbThousandSeparators, vbChecked, vbUnchecked)

  frmBreakDown.Initialise Me
  
  With frmBreakDown

    ' JDM - 22/06/01 - Ensure drilldown knows what type of report is being run
    .ReportMode = mlngCrossTabType

    .Caption = Me.Caption & " Cell Breakdown"
    If mlngCrossTabType <> cttNormal Then
      If mlngCrossTabType = cttAbsenceBreakdown Then
        .lblHorizontalFieldName = "Day :"
      Else
        .lblHorizontalFieldName = "Employee selection :"
      End If
    Else
      .lblHorizontalFieldName = Replace(mstrColName(0), "_", " ") & " :"
    End If
    
    If mlngCrossTabType = cttAbsenceBreakdown Then
      .lblVerticalFieldName = "Type : "
      .lblPageBreakFieldName.Visible = False
      .cboValue(2).Visible = False
    Else
      .lblVerticalFieldName = Replace(mstrColName(1), "_", " ") & " :"
      .lblPageBreakFieldName = txtPageBreakCol.Text & " :"
    End If

    With .SSDBGrid1
      
      .Columns.RemoveAll
      .Columns.Add 0
      .Columns(0).Caption = mstrBaseTable
      
      If mlngCrossTabType <> cttNormal Then
        .Columns(0).Width = .Width - 2400
        .Columns.Add 1
        .Columns(1).Caption = Replace(gsPersonnelStartDateColumnName, "_", " ")
        .Columns(1).Width = 1200
        .Columns.Add 2
        If mlngCrossTabType = cttAbsenceBreakdown Then
          .Columns.Add 3
          .Columns(1).Caption = "Start Date"
          .Columns(2).Caption = "End Date"
          .Columns(3).Caption = "Days Taken"
        Else
          .Columns(2).Caption = Replace(gsPersonnelLeavingDateColumnName, "_", " ")
        End If
        .Columns(2).Width = 1200

      ElseIf mblnIntersection Then
        .Columns(0).Width = .Width - 1200
        .Columns.Add 1
        .Columns(1).Caption = Replace(mstrColName(INS), "_", " ")
        .Columns(1).Width = 1200
        .Columns(1).Alignment = ssCaptionAlignmentRight
        .Columns(1).CaptionAlignment = ssColCapAlignLeftJustify
      
      End If

    End With

  End With


'MH20021203
'  frmOutputOptions.Initialise Me, True
'
'  With frmOutputOptions
'
'    .optOutput(mintDefaultOutput).Value = True
'
'    If mintDefaultOutput = 1 Then
'      .cboExportTo.ListIndex = mintDefaultExportTo
'      .chkSave = IIf(mblnDefaultSave, vbChecked, vbUnchecked)
'      .txtFilename.Text = mstrDefaultSaveAs
'      .chkCloseApplication = IIf(mblnDefaultCloseApp, vbChecked, vbUnchecked)
'    End If
'
'    .optDataRange(0).Value = True
'    If Not mblnPageBreak Then
'      .cboPageBreak.BackColor = vbButtonFace
'      .cboPageBreak.Enabled = False
'      .optDataRange(1).Enabled = False
'      .optDataRange(0).Enabled = False
'    End If
'
'  End With

  ' Get rid of the icon off the form
  RemoveIcon Me

Exit Sub

LocalErr:
  mstrErrorMessage = "Error populating forms"
  fOK = False

End Sub


Private Sub PopulateCombos()

  On Error GoTo LocalErr
  
  PopulateComboWithArray cboPageBreak, mvarHeadings(2), False, False
  PopulateComboWithArray cboType, mstrType(), False, False
  SetComboItem cboType, mlngIntersectionType

  If gblnBatchMode Then
    Exit Sub
  End If
  
  
  With frmBreakDown
    
    If mlngCrossTabType = cttNormal Then
      PopulateComboWithArray .cboValue(HOR), mvarHeadings(HOR), True, Not mblnShowAllPagesTogether
    Else

      If mlngCrossTabType = cttAbsenceBreakdown Then
        'MH20040211 Fault 8036
        'PopulateComboWithArray .cboValue(HOR), mvarHeadings(HOR), True, Not mblnShowAllPagesTogether
        PopulateComboWithArray .cboValue(HOR), mvarHeadings(HOR), True, False
      Else

    
      With .cboValue(HOR)
        .AddItem "Staff at " & Format(mdtReportStartDate, DateFormat)
        .ItemData(.NewIndex) = 0
        
        If mlngCrossTabType = cttTurnover Then
          .AddItem "Leavers between " & Format(mdtReportStartDate, DateFormat) & _
                             " - " & Format(mdtReportEndDate, DateFormat)
          .ItemData(.NewIndex) = 1
        
          .AddItem "Staff at " & Format(mdtReportEndDate, DateFormat)
          .ItemData(.NewIndex) = 2
        Else
          .AddItem "Employed for report duration"
          .ItemData(.NewIndex) = 1
        
          .AddItem "Leavers between " & Format(mdtReportStartDate, DateFormat) & _
                                " - " & Format(mdtReportEndDate, DateFormat)
          .ItemData(.NewIndex) = 2
        
        End If

      End With

    End If
    
    End If

    PopulateComboWithArray .cboValue(VER), mvarHeadings(VER), True, True

    If mblnPageBreak Then
      PopulateComboWithArray .cboValue(PGB), mvarHeadings(PGB), False, mblnShowAllPagesTogether
    Else
      .cboValue(PGB).Enabled = (.cboValue(PGB).ListCount > 1)
      .cboValue(PGB).BackColor = IIf(.cboValue(PGB).Enabled, vbWindowBackground, vbButtonFace)
    End If
  
  End With

'MH20021203
'  PopulateComboWithArray frmOutputOptions.cboPageBreak, mvarHeadings(PGB), False, False
'  With frmOutputOptions.cboPageBreak
'    .Enabled = (.Enabled And Not mblnShowAllPagesTogether)
'  End With

Exit Sub

LocalErr:
  mstrErrorMessage = "Error populating combo boxes" & vbCrLf & Err.Description
  fOK = False

End Sub


Private Sub PopulateComboWithArray(cboOutput As ComboBox, varArray As Variant, blnBypassDisable As Boolean, blnAddAllOption)

  Dim lngCount As Long

  On Error GoTo LocalErr
  
  With cboOutput
    .Clear
    For lngCount = 0 To UBound(varArray)
      .AddItem Trim(varArray(lngCount))
      .ItemData(.NewIndex) = lngCount
    Next
    
    If blnAddAllOption Then
      .AddItem "<All>"
      .ItemData(.NewIndex) = lngCount
    End If

    .ListIndex = 0
    .Enabled = (.ListCount > 1 Or blnBypassDisable)
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
  End With

Exit Sub

LocalErr:
  mstrErrorMessage = "Error populating combo box <" & cboOutput.Name & ">"
  fOK = False

End Sub

Private Sub SSDBGrid1_Click()
  'This should avoid getting two cells highlighted at once
  SSDBGrid1.Refresh
End Sub

Private Sub SSDBGrid1_DblClick()
  
  Dim lngHOR As Long
  Dim lngVER As Long
  Dim lngPGB As Long
  Dim lngEndtime As Long
  
  On Error GoTo LocalErr

  If mblnLoading Then Exit Sub
  
' JIRA-381  - Why check - they should already be processed... Causes error when record description is blank.
'  If Not CheckIfGotRecDescs Then
'    Exit Sub
'  End If
  
  
  With SSDBGrid1
  
    mblnLoading = True
    .Redraw = False
    
    If mblnShowAllPagesTogether Then
      lngVER = SSDBGrid1.AddItemRowIndex(SSDBGrid1.Bookmark)
      lngHOR = (SSDBGrid1.Col - 1) Mod (UBound(mvarHeadings(HOR)) + 1)
      lngPGB = (SSDBGrid1.Col - 1) \ (UBound(mvarHeadings(HOR)) + 1)
    Else
      lngVER = SSDBGrid1.AddItemRowIndex(SSDBGrid1.Bookmark)
      lngHOR = SSDBGrid1.Col - 1
      lngPGB = cboPageBreak.ItemData(cboPageBreak.ListIndex)
    End If
  
  End With
  
  If lngHOR >= 0 And lngVER >= 0 Then
  
    If PopulateCellBreakdown2(lngHOR, lngVER, lngPGB) Then
      With frmBreakDown
        .lblType = cboType.Text & " :"
        mblnLoading = True
        SetComboItem .cboValue(HOR), lngHOR
        SetComboItem .cboValue(VER), lngVER
        SetComboItem .cboValue(PGB), lngPGB
        SetComboItem cboPageBreak, lngPGB
      End With
      SSDBGrid1.Redraw = True
      mblnLoading = False
      frmBreakDown.Show vbModal
    End If

  End If

  SSDBGrid1.Redraw = True
  mblnLoading = False

Exit Sub

LocalErr:
  mstrErrorMessage = "Error viewing data"
  SSDBGrid1.Redraw = True
  mblnLoading = False
  fOK = False

End Sub


Private Function GetRecordDesc(lngRecordID As Long)

  ' Return TRUE if the user has been granted the given permission.
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  On Error GoTo LocalErr
  
  If mlngRecordDescExprID < 1 Then
    GetRecordDesc = "Record Description Undefined"
    Exit Function
  End If
  
  
  ' Check if the user can create New instances of the given category.
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "dbo.sp_ASRExpr_" & mlngRecordDescExprID
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("Result", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("RecordID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = lngRecordID

    cmADO.Execute

    GetRecordDesc = .Parameters(0).Value
  End With
  Set cmADO = Nothing

Exit Function

LocalErr:
  mstrErrorMessage = "Error reading record description" & vbCr & _
                  "(ID = " & CStr(lngRecordID) & ", Record Description = " & CStr(mlngRecordDescExprID)
  fOK = False

End Function


Public Sub BreakdownComboClick(Index As Integer)
  
  Dim lngHOR As Long
  Dim lngVER As Long
  Dim lngPGB As Long
  
  If mblnLoading Then Exit Sub

  On Error GoTo LocalErr

  mblnLoading = True
    
  With frmBreakDown
  
    'Store new cell reference
    lngHOR = .cboValue(0).ItemData(.cboValue(0).ListIndex)
    lngVER = .cboValue(1).ItemData(.cboValue(1).ListIndex)
    If mblnPageBreak Then
      lngPGB = .cboValue(2).ItemData(.cboValue(2).ListIndex)
    End If

    'Change current cell highlighted
    With SSDBGrid1
      .Redraw = False

      .SelBookmarks.RemoveAll
      .Bookmark = .AddItemBookmark(lngVER)
      
      If mblnShowAllPagesTogether Then
        '.Col = (lngHOR + 1) + (lngPGB * (UBound(mvarHeadings(PGB))))
        .Col = (lngHOR + 1) + (lngPGB * (UBound(mvarHeadings(HOR)) + 1))
      Else
        .Col = lngHOR + 1
      End If
      .Redraw = True

    End With
    
    SetComboItem cboPageBreak, CLng(lngPGB)

    'If page changed repopulate grid
    'If Index = 2 Then
      Call PopulateGrid
    'End If
    
    'Reload cell breakdown
    Call PopulateCellBreakdown2(lngHOR, lngVER, lngPGB)

  End With
    
  mblnLoading = False

Exit Sub

LocalErr:
  MsgBox Err.Description, vbCritical, Me.Caption
  mblnLoading = False
  SSDBGrid1.Redraw = True

End Sub

Private Sub SSDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

  With SSDBGrid1
    Select Case KeyCode
    Case vbKeyLeft
      If .Col > 1 Then
        .Col = .Col - 1
      End If
      KeyCode = 0
    Case vbKeyRight
      If .Col < .Columns.Count - 1 Then
        .Col = .Col + 1
      End If
      KeyCode = 0
    End Select
  End With

End Sub


Private Sub SSDBGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' JPD 26/7/00
  Screen.MousePointer = vbArrow

End Sub

'Private Sub SSDBGrid1_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B.ssPrintInfo)
'  Call CrossTabPrint.PrintInitialise(ssPrintInfo)
'End Sub
'
'
'Private Sub OutputHTML()
'
'  Dim lngTempPage As Long
'  Dim lngPage As Long
'  Dim lngCount As Long
'  Dim strFilename As String
'  Dim blnTempSuppressZeros As Boolean
'  Dim lngErrorTime As Long
'  Dim lngTime As Long
'
'
'  On Error GoTo LocalErr
'
'
'  'Remember the current page so that it can be restored after printing
'  lngTempPage = cboPageBreak.ItemData(cboPageBreak.ListIndex)
'
'  With frmOutputOptions
'
'    If .txtFilename.Text = "" Then
'      strFilename = GetTmpFName
'      strFilename = Left$(strFilename, Len(strFilename) - 3) & "htm"
'    Else
'      strFilename = .txtFilename.Text
'    End If
'    lngFileNum = FreeFile
'
'    Open strFilename For Output Access Write As #lngFileNum
'    Print #lngFileNum, "<HTML><BODY><CENTER>"
'
'    '25/07/2001 MH Fault 2241
'    'If .optDataRange(1).Value = True Then
'    If .optDataRange(1).Value = True Or mblnShowAllPagesTogether Then
'
'      'single page
'      lngPage = .cboPageBreak.ItemData(.cboPageBreak.ListIndex)
'      Call BuildOutputStrings(lngPage)
'      Call OutputHTMLPage(lngPage)
'
'    Else
'      'All Pages
'      For lngCount = 0 To UBound(mvarHeadings(PGB))
'        Call BuildOutputStrings(lngCount)
'        Call OutputHTMLPage(lngCount)
'
'        If lngCount < UBound(mvarHeadings(PGB)) Then
'          Print #lngFileNum, "</TABLE><BR><HR>"
'        End If
'      Next
'
'    End If
'
'    Print #lngFileNum, "</TABLE>" & _
'            "<P></P><HR><div align=" & Chr(34) & "center" & Chr(34) & ">" & vbCrLf & _
'            "<table width=" & Chr(34) & "99%" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & ">" & vbCrLf & _
'            "<tr>" & vbCrLf & _
'            "<td width=" & Chr(34) & "50%" & Chr(34) & "><font face=" & Chr(34) & "Verdana, Arial, Helvetica, sans-serif" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & ">Created on " & Format(Now, DateFormat) & " at " & Format(Now, "hh:nn") & " by " & gsUserName & "</font></td>" & vbCrLf & _
'            "<td width=" & Chr(34) & "48%" & Chr(34) & ">" & vbCrLf & _
'            "<div align=" & Chr(34) & "right" & Chr(34) & "><font face=" & Chr(34) & "Verdana, Arial, Helvetica, sans-serif" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & ">Page &lt;n/a&gt;</font></div>" & vbCrLf & _
'            "</td>" & vbCrLf & _
'            "</tr>" & vbCrLf & _
'            "</table>" & vbCrLf & _
'            "</div>"
'
'
'    Print #lngFileNum, "</CENTER></BODY></HTML>"
'    Close #lngFileNum
'
'  End With
'
'  'Only open browser if required...
'  If frmOutputOptions.chkCloseApplication = False Then
'    Call ShellExecute(0&, vbNullString, strFilename, vbNullString, vbNullString, 0)
'  End If
'
'Exit Sub
'
'LocalErr:
'
'  If Err.Number = 76 Then
'    On Local Error Resume Next
'    mstrErrorMessage = "Error exporting to HTML"
'    fOK = False
'  ElseIf lngErrorTime = 0 Then
'    lngErrorTime = Timer
'    Resume 0
'  ElseIf Timer - lngErrorTime < 2 Then
'    Resume 0
'  Else
'    On Local Error Resume Next
'    mstrErrorMessage = "Error exporting to HTML"
'    fOK = False
'
'  End If
'  Err.Clear
'
'End Sub
'
'Private Sub OutputHTMLPage(lngPage As Long)
'
'  Dim strTemp As String
'  Dim lngCount As Long
'  Dim strHeadings As String
'
'  Print #lngFileNum, "<BR>"
'
'  'MH20010824 Fault 2474
'  'Print #lngFileNum, "<B>" & HTMLText(mstrCrossTabName) & "</B>"
'  Print #lngFileNum, "<B>" & HTMLText(Replace(SSDBGrid1.Caption, "&&", "&")) & "</B>"
'
'  If mblnPageBreak And Not mblnShowAllPagesTogether Then
'    strTemp = mstrColName(PGB) & " : " & Trim(mvarHeadings(PGB)(lngPage))
'    Print #lngFileNum, "</CENTER>"
'    Print #lngFileNum, HTMLText(strTemp)
'    Print #lngFileNum, "<CENTER>"
'  Else
'    Print #lngFileNum, "<BR>"
'  End If
'
'  Print #lngFileNum, "<BR>"
'  Print #lngFileNum, "<TABLE border=1 cellspacing=1 cellpadding=3 bordercolorlight=#000000 bordercolordark=#000000>"
'
'  If mblnShowAllPagesTogether And mblnPageBreak Then
'    strHeadings = "<TR align=center><TD border=1 rowspan=2 bgcolor=#CCCCCC>&nbsp;</TD>"
'    For lngCount = 0 To UBound(mvarHeadings(PGB))
'      strHeadings = strHeadings & _
'          "<TD colspan=3 bgcolor=#CCCCCC>" & _
'          Replace(Replace(Trim(mvarHeadings(PGB)(lngCount)), "<", "&LT;"), ">", "&GT;") & _
'          "</TD>"
'    Next
'    Print #lngFileNum, strHeadings & "<TD colspan=3 bgcolor=#CCCCCC>Total</TD></TR>"
'
'    Print #lngFileNum, "<TD bgcolor=#CCCCCC>" & _
'                       Replace(HTMLText(Mid(mstrOutput(0), 2)), "<TD>", "<TD bgcolor=#CCCCCC>") & _
'                       "</TD></TR>"
'  Else
'    Print #lngFileNum, "<TD bgcolor=#CCCCCC>" & _
'                       Replace(HTMLText(mstrOutput(0)), "<TD>", "<TD bgcolor=#CCCCCC>") & _
'                       "</TD></TR>"
'  End If
'
'  For lngCount = 1 To UBound(mstrOutput)
'    Print #lngFileNum, "<TD bgcolor=#CCCCCC>" & _
'                       HTMLText(mstrOutput(lngCount)) & _
'                       "</TD></TR>"
'  Next
'
'End Sub


'Private Function GetFilterPicklistName(lngExprID As Long) As Long
''NHRD 25/03/2002 This function returns the Name of a specified ExprId ID
''I want to improve at a later date to return the Picklist Name too.
'
'Dim rsTemp As Recordset
'
'If lngExprID > 0 Then
'
'Else
'
'End If
'
'
''open the ASRSysExpressions table
''Retrieve the Name
''Close table object reference
'
'End Function

Private Function GetPicklistFilterSelect(lngPicklistID As Long, lngFilterID As Long) As String

  Dim rsTemp As Recordset

  If lngPicklistID > 0 Then

    mstrErrorMessage = IsPicklistValid(lngPicklistID)
    If mstrErrorMessage <> vbNullString Then
      InvalidPicklistFilter = True
      fOK = False
      Exit Function
    End If

    'Get List of IDs from Picklist
    Set rsTemp = datData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & lngPicklistID, adOpenForwardOnly, adLockReadOnly)
    fOK = Not (rsTemp.BOF And rsTemp.EOF)

    If Not fOK Then
      mstrErrorMessage = "The base table picklist contains no records."
    Else
      GetPicklistFilterSelect = vbNullString
      Do While Not rsTemp.EOF
        GetPicklistFilterSelect = GetPicklistFilterSelect & _
            IIf(Len(GetPicklistFilterSelect) > 0, ", ", "") & rsTemp.Fields(0)
        rsTemp.MoveNext
      Loop
    End If

    rsTemp.Close
    Set rsTemp = Nothing

  ElseIf lngFilterID > 0 Then

    mstrErrorMessage = IsFilterValid(lngFilterID)
    If mstrErrorMessage <> vbNullString Then
      InvalidPicklistFilter = True
      fOK = False
      Exit Function
    End If

    'Get list of IDs from Filter
    fOK = datGeneral.FilteredIDs(lngFilterID, GetPicklistFilterSelect)

    ' Generate any UDFs that are used in this filter
    If fOK Then
      datGeneral.FilterUDFs lngFilterID, mastrUDFsRequired()
    End If

    If Not fOK Then
      ' Permission denied on something in the filter.
      mstrErrorMessage = "You do not have permission to use the '" & datGeneral.GetFilterName(lngFilterID) & "' filter."
    End If

  End If

End Function


Public Function PopulateCellBreakdown2(lngHOR As Long, lngVER As Long, lngPGB As Long) As Boolean

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim objColumnPrivileges As CColumnPrivileges
  Dim strOutput As String

  Dim strColumnName() As String
  Dim strWhere As String
  Dim intMax As Integer
  

  On Error GoTo LocalErr

  Screen.MousePointer = vbHourglass
  PopulateCellBreakdown2 = False
  mblnLoading = True
  
  
  strWhere = vbNullString
  If lngHOR <= UBound(mvarSearches(HOR)) Then
    strWhere = _
      IIf(strWhere = vbNullString, " WHERE ", strWhere & " AND ") & _
      "(" & mvarSearches(HOR)(lngHOR) & ")"
  End If

  If lngVER <= UBound(mvarSearches(VER)) Then
    strWhere = _
      IIf(strWhere = vbNullString, " WHERE ", strWhere & " AND ") & _
      "(" & mvarSearches(VER)(lngVER) & ")"
  End If

  If mblnPageBreak Then
    If lngPGB <= UBound(mvarSearches(PGB)) Then
      strWhere = _
        IIf(strWhere = vbNullString, " WHERE ", strWhere & " AND ") & _
        "(" & mvarSearches(PGB)(lngPGB) & ")"
    End If
  End If
  
  
  strSQL = "SELECT * FROM " & mstrTempTableName & _
           strWhere & " ORDER BY "

  Select Case cboType.ItemData(cboType.ListIndex)
  Case TYPEMINIMUM: strSQL = strSQL & "Ins, "
  Case TYPEMAXIMUM: strSQL = strSQL & "Ins DESC, "
  End Select
  strSQL = strSQL & "RecDesc"

  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)


  With frmBreakDown.SSDBGrid1

    .Redraw = False
    .RemoveAll
    
    With rsTemp
      If Not .EOF Then .MoveFirst
      Do While Not .EOF
        'ID Column
        strOutput = .Fields("RecDesc").Value

        If mlngCrossTabType <> cttNormal Then

          If mlngCrossTabType = cttAbsenceBreakdown Then
            strOutput = strOutput & vbTab

            ' Add absence start date
              If IsNull(.Fields("Start_Date").Value) Then
                strOutput = strOutput & vbTab
              Else
                strOutput = strOutput & Format(.Fields("Start_Date").Value, DateFormat) & vbTab
              End If

            ' Add absence end date
              If IsNull(.Fields("End_Date").Value) Then
                strOutput = strOutput & vbTab
              Else
                strOutput = strOutput & Format(.Fields("End_Date").Value, DateFormat) & vbTab
              End If

            ' Add occurences
              If IsNull(.Fields("Value").Value) Then
                strOutput = strOutput & vbTab
              Else
                'MH20040127 Fault 7092 - Round average to 2 decimal places
                'strOutput = strOutput & .Fields("Value").Value & vbTab
                strOutput = strOutput & Format(.Fields("Value").Value, "0.00") & vbTab
              End If

          Else
            strOutput = strOutput & vbTab
          
            'Add start date
            If IsNull(.Fields("StartDate").Value) Then
              strOutput = strOutput & vbTab
            Else
              strOutput = strOutput & Format(.Fields("StartDate").Value, DateFormat) & vbTab
            End If
            
            'Add leaving date
            If IsNull(.Fields("LeavingDate").Value) Then
              strOutput = strOutput & vbTab
            Else
              strOutput = strOutput & Format(.Fields("LeavingDate").Value, DateFormat) & vbTab
            End If
          End If
        
        ElseIf mblnIntersection Then
          'add intersection
          If Not IsNull(.Fields("Ins").Value) Then
            strOutput = strOutput & vbTab & Format(.Fields("Ins"), mstrIntersectionMask)
          End If

        End If

        frmBreakDown.SSDBGrid1.AddItem strOutput
        .MoveNext
      Loop
    End With

    .Redraw = True
    Call SizeBreakdownColumns
    Screen.MousePointer = vbDefault
  
  End With
  
  frmBreakDown.txtCellValue = IIf(mlngCrossTabType = cttNormal, SSDBGrid1.ActiveCell.Text, frmBreakDown.SSDBGrid1.Rows)
  mblnLoading = False
  PopulateCellBreakdown2 = True

Exit Function

LocalErr:
  mblnLoading = False
  frmBreakDown.SSDBGrid1.Redraw = True
  Screen.MousePointer = vbDefault
  'Never see breakdown if running in batch mode so no need to check if batch
  MsgBox "Error reading Breakdown" & vbCrLf & Err.Description, vbExclamation, Me.Caption
  Err.Clear

End Function


Public Function TurnoverExecuteReport()
  
  Set datData = New HRProDataMgr.clsDataAccess
  mblnLoading = True
  
  'NHRD09042002 Fault 3322 - Code Added
  mlngCrossTabType = cttTurnover
  
  ReDim mastrUDFsRequired(0)
  
  fOK = True

  If gblnBatchMode = True Then
    Call GetReportConfig("Turnover")
  End If
  Call TurnoverRetreiveDefinition

  gobjEventLog.AddHeader eltStandardReport, mstrCrossTabName

  If fOK Then Call InitialiseProgressBar
  If Progress Then Call UDFFunctions(mastrUDFsRequired, True)
  If Progress Then Call CreateTempTable
  If Progress Then Call UDFFunctions(mastrUDFsRequired, False)
  If Progress Then Call TurnoverGetHeadingsAndSearches
  If Progress Then Call BuildTypeArray
  If Progress Then Call TurnoverBuildDataArrays
  If Progress Then Call PopulateCombos
  If Progress Then Call PrepareForms
  If Progress Then Call TurnoverCreateGridColumns
  If Progress Then Call PopulateGrid

  If Progress Then
    If gblnBatchMode Or Not mblnPreviewOnScreen Then
      fOK = OutputReport(False)
    End If
  End If
  
  mblnLoading = False

  Call OutputJobStatus

  TurnoverExecuteReport = fOK

End Function


Public Function TurnoverStabilityReport()
  
  Dim strReportType As String
  
  Set datData = New HRProDataMgr.clsDataAccess
  mblnLoading = True
  
  ReDim mastrUDFsRequired(0)
  
  mlngCrossTabType = cttStability
  fOK = True
  
  If gblnBatchMode = True Then
    Call GetReportConfig("Stability")
  End If
  Call TurnoverRetreiveDefinition
  
  gobjEventLog.AddHeader eltStandardReport, mstrCrossTabName
  
  If fOK Then Call InitialiseProgressBar
  If Progress Then Call UDFFunctions(mastrUDFsRequired, True)
  If Progress Then Call CreateTempTable
  If Progress Then Call UDFFunctions(mastrUDFsRequired, False)
  If Progress Then Call TurnoverGetHeadingsAndSearches
  If Progress Then Call BuildTypeArray
  If Progress Then Call StabilityBuildDataArrays
  If Progress Then Call PopulateCombos
  If Progress Then Call PrepareForms
  If Progress Then Call TurnoverCreateGridColumns
  If Progress Then Call PopulateGrid

  If Progress Then
    If gblnBatchMode Or Not mblnPreviewOnScreen Then
      fOK = OutputReport(False)
    End If
  End If

  mblnLoading = False

  Call OutputJobStatus

  TurnoverStabilityReport = fOK

End Function


Private Sub TurnoverCreateGridColumns()

  Dim intColumn As Integer
  Dim intGroup As Integer
  Dim lngNumPages As Long
  'NHRD09042002 Fault 3322 - variable Added
  Dim strExprId As Variant
  
  On Local Error GoTo LocalErr

  With SSDBGrid1
    .Redraw = False
    
    'Ascertain whether the user has decided to include Picklist/Filter info
    'mblnChkPicklistFilter = frmCrossTabStdRpts.chkPrintFilterHeader.Value
    'mblnChkPicklistFilter = (GetUserSetting("Turnover", "PrintFilterHeader", 0) = 1)
    
    If mblnChkPicklistFilter = True Then
      'this will add the picklist or filter name into the SSDBGRID control
      'unless there is no value then this is indicated instead.
      'Work it out here so we don't have to put it into each CASE branch below
      If mstrPicklistFilterType = "F" Then
          strExprId = GetExprField(mlngPicklistFilterID, "NAME")
          strExprId = " (Base Table Filter: " & strExprId & ")"
      ElseIf mstrPicklistFilterType = "P" Then
          strExprId = GetPickListField(mlngPicklistFilterID, "NAME")
          strExprId = " (Base Table Picklist: " & strExprId & ")"
      Else
          strExprId = " (No Picklist or Filter Selected)"
      End If
    End If
    
    Select Case mlngCrossTabType
    Case cttTurnover
    
      .Caption = _
          "Turnover Report by " & Replace(mstrColName(VER), "_", " ") & _
          IIf(mblnPageBreak, " and " & Replace(mstrColName(PGB), "_", " "), "") & _
          " (" & Format(mdtReportStartDate, DateFormat) & _
          " - " & Format(mdtReportEndDate, DateFormat) & ")" & _
          IIf(mblnChkPicklistFilter, strExprId, vbNullString)
    
    Case cttStability

      .Caption = _
          "Stability Index Report by " & Replace(mstrColName(VER), "_", " ") & _
          IIf(mblnPageBreak, " and " & Replace(mstrColName(PGB), "_", " "), "") & _
          " (" & Format(mdtReportStartDate, DateFormat) & _
          " - " & Format(mdtReportEndDate, DateFormat) & ")" & _
          IIf(mblnChkPicklistFilter, strExprId, vbNullString)

    End Select
    
    
    'MH20030902 Fault 6761
    .Caption = Replace(.Caption, "&", "&&")
    
    
    .Groups.RemoveAll
    
    .Groups.Add 0
    .Groups(0).Columns.Add 0
    With .Groups(0).Columns(0)
      .Caption = vbNullString
      .Locked = True
      .Style = ssStyleButton
      .ButtonsAlways = True
      .BackColor = vbButtonFace   'required for printing !
    End With

    lngNumPages = IIf(mblnPageBreak, UBound(mvarHeadings(PGB)) + 1, 0)

    For intGroup = 0 To lngNumPages
      .Groups.Add intGroup + 1
      
      With .Groups(intGroup + 1)
        For intColumn = 0 To UBound(mvarHeadings(HOR))
          .Columns.Add intColumn
        Next
        For intColumn = 0 To UBound(mvarHeadings(HOR))
          .Columns(intColumn).Width = 1350
          .Columns(intColumn).Locked = True
          .Columns(intColumn).Caption = Trim(mvarHeadings(HOR)(intColumn))
          .Columns(intColumn).Alignment = ssCaptionAlignmentRight
          .Columns(intColumn).CaptionAlignment = ssColCapAlignCenter
        Next

        .Width = (UBound(mvarHeadings(HOR)) + 1) * 1350

        If mblnPageBreak Then
          If intGroup = UBound(mvarHeadings(PGB)) + 1 Then
            .Caption = "Total"
          Else
            .Caption = Trim(mvarHeadings(PGB)(intGroup))
          End If
        Else
          SSDBGrid1.GroupHeadLines = 0
        End If


      End With
    
    Next

    .SplitterPos = 1
    .SplitterVisible = False
    .Redraw = True
  
  End With

Exit Sub

LocalErr:
  MsgBox "Error creating grid columns" & vbCrLf & "(" & Err.Description & ")", vbCritical, Me.Caption

End Sub


Public Sub TurnoverGetHeadingsAndSearches()

  Dim strColumnName As String
  
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strHeading() As String
  Dim strSearch() As String
  Dim strFieldValue As String
  Dim lngLoop As Long
  Dim lngCount As Long
  Dim dblGroup As Double
  Dim dblGroupMax As Double
  Dim dblUnit As Double

  On Error GoTo LocalErr
  
  For lngLoop = 0 To 2
    
    If lngLoop = 0 Then
      'When no page break field is specified
      ReDim strHeading(2) As String
      ReDim strSearch(2) As String

      If mlngCrossTabType = cttTurnover Then
        strHeading(0) = "Staff"
        strHeading(1) = "Leavers"
        strHeading(2) = "Turnover"
        
        strSearch(0) = SQLEmployedAtStartOfReport("startdate", "leavingdate")
        strSearch(1) = SQLLeaversBetweenStartAndEnd("startdate", "leavingdate")
        strSearch(2) = SQLEmployedAtEndOfReport("startdate", "leavingdate")
      
      Else
        strHeading(0) = "Staff At Start"
        strHeading(1) = "Remaining Staff"
        strHeading(2) = "Stability Index"

        strSearch(0) = SQLEmployedAtStartOfReport("startdate", "leavingdate")
        strSearch(1) = SQLOneYearServiceAtEndOfReport("startdate", "leavingdate")
        strSearch(2) = SQLLeaversBetweenStartAndEnd("startdate", "leavingdate")
      
      End If
    
    Else
      'MH20070301 Fault 11867
      ReDim strHeading(0) As String
      ReDim strSearch(0) As String
      
      If mblnPageBreak Or lngLoop <> 2 Then
        'ReDim strHeading(0) As String
        'ReDim strSearch(0) As String
        GetHeadingsAndSearchesForColumns lngLoop, strHeading(), strSearch()
      End If
    End If
    
    mvarHeadings(lngLoop) = strHeading
    mvarSearches(lngLoop) = strSearch

  Next

Exit Sub

LocalErr:
  mstrErrorMessage = "Error building headings and search arrays"
  fOK = False
  
End Sub
Public Sub AbsenceBreakdownGetHeadingsAndSearches()

  Dim strHeading() As String
  Dim strSearch() As String
  Dim lngLoop As Long
  
  
  On Error GoTo LocalErr
  
  For lngLoop = 0 To 2
    
    ReDim strHeading(0) As String
    ReDim strSearch(0) As String
    
    If lngLoop = 2 And mblnPageBreak = False Then
      'When no page break field is specified
      strHeading(0) = "<None>"
    Else
      GetHeadingsAndSearchesForColumns lngLoop, strHeading(), strSearch()
    End If

    
    'Store each array in an array of variants (an array in an array!)
    mvarHeadings(lngLoop) = strHeading
    mvarSearches(lngLoop) = strSearch

  Next

Exit Sub

LocalErr:
  mstrErrorMessage = "Error building headings and search arrays"
  fOK = False
  
End Sub

Private Sub TurnoverRetreiveDefinition()

  Dim strReportType As String
  'Dim lngID As Long
  'Dim lngHorCol As Long
  'Dim lngVerCol As Long
  Dim lngExprID As Long

  Select Case mlngCrossTabType
  Case cttTurnover
    mstrCrossTabName = "Turnover Report"
    strReportType = "Turnover"
  Case cttStability
    mstrCrossTabName = "Stability Index Report"
    strReportType = "Stability"
  End Select
  Me.Caption = mstrCrossTabName

  'If mblnCustomDates Then
  '  mdtReportStartDate = datGeneral.GetValueForRecordIndependantCalc(mlngStartDateExprID)
  '  mdtReportEndDate = datGeneral.GetValueForRecordIndependantCalc(mlngEndDateExprID)
  'Else
  '  mdtReportEndDate = DateAdd("d", Day(Date) * -1, Date)
  '  mdtReportStartDate = DateAdd("d", 1, DateAdd("yyyy", -1, mdtReportEndDate))
  'End If

  mlngBaseTableID = glngPersonnelTableID
  mstrBaseTable = gsPersonnelTableName
  mlngRecordDescExprID = datGeneral.GetRecDescExprID(mlngBaseTableID)

  Select Case mstrPicklistFilterType
  Case "A"
    mstrPicklistFilter = vbNullString
  Case "P"
    mstrPicklistFilter = GetPicklistFilterSelect(mlngPicklistFilterID, 0)
  Case "F"
    mstrPicklistFilter = GetPicklistFilterSelect(0, mlngPicklistFilterID)
  End Select

  
  If fOK = False Then
    Exit Sub
  End If
  
  
  mlngColID(HOR) = glngPersonnelStartDateID
  mstrColName(HOR) = gsPersonnelStartDateColumnName
  
  'TM20020426 Fault 3239 - store the data type of the horizontal column.
  'mlngColDataType(HOR) = datGeneral.GetDataType(mlngBaseTableID, glngPersonnelStartDateID)
  'mlngColDataType(HOR) = datGeneral.GetDataType(mlngBaseTableID, lngHorColID)

  mstrFormat(HOR) = GetFormat(mlngColID(HOR))

  If mlngVerCol = 0 Then
    mstrErrorMessage = "The " & strReportType & " Report has not been configured.  Please set this up in the Report Configuration under the Administration menu."
    fOK = False
    Exit Sub
  End If
  
  mlngColID(VER) = mlngVerCol
  mstrColName(VER) = datGeneral.GetColumnName(mlngVerCol)
  mlngColDataType(VER) = datGeneral.GetDataType(mlngBaseTableID, mlngVerCol)
  mstrFormat(VER) = GetFormat(mlngColID(VER))

  'Map the Horizontal column to the page break column !
  '(This is easier as the cross tab page break column
  ' and the turnover horizontal are both optional !)
  mblnPageBreak = (mlngHorCol > 0)
  If mblnPageBreak Then
    mlngColID(PGB) = mlngHorCol
    mstrColName(PGB) = datGeneral.GetColumnName(mlngHorCol)
    mlngColDataType(PGB) = datGeneral.GetDataType(mlngBaseTableID, mlngHorCol)
    mstrFormat(PGB) = GetFormat(mlngColID(PGB))
  End If

  mblnIntersection = False
  mlngIntersectionDecimals = 1
  mblnShowAllPagesTogether = True

  mblnPreviewOnScreen = (mblnPreviewOnScreen Or (mlngOutputFormat = fmtDataOnly And mblnOutputScreen))

End Sub

Private Sub AbsenceBreakdownRetreiveDefinition(lngPersonnelID As Long)

'(dtStartDate As Date, dtEndDate As Date _
  , lngHorColID As Long, lngVerColID As Long, lngPicklistID As Long, lngFilterID As Long _
  , lngPersonnelID As Long, pastrIncludedTypes() As String)

  'Dim objAbsBreakdown As clsAbsenceBreakdown
  Dim rsType As Recordset
  Dim strSQL As String
  Dim strType As String
  
  Dim iCount As Integer
  Dim lngExprID As Long
  Dim lngID As Long
  Dim lngHorColID As Long
  Dim lngVerColID As Long
  'Dim strReportType As String

  'mstrCrossTabName = "Absence Breakdown"
   mstrCrossTabName = "Absence Breakdown Report"
  Me.Caption = mstrCrossTabName
  
  If Not fOK Then
    Exit Sub
  End If
  
'  strSQL = "SELECT * " & _
'           "FROM " & gsAbsenceTypeTableName & " " & _
'           "ORDER BY " & gsAbsenceTypeTypeColumnName
'  Set rsType = datGeneral.GetReadOnlyRecords(strSQL)
'
'  Set objAbsBreakdown = New clsAbsenceBreakdown
'  objAbsBreakdown.ReportType = "AbsenceBreakdown"
'
'  msAbsenceBreakdownTypes = vbNullString
'  Do Until rsType.EOF
'
'    strType = rsType.Fields(gsAbsenceTypeTypeColumnName).Value
'    If objAbsBreakdown.CheckIfAbsenceTypeSelected(strType) = True Then
'      msAbsenceBreakdownTypes = _
'        IIf(msAbsenceBreakdownTypes <> vbNullString, msAbsenceBreakdownTypes & ", ", "") & _
'        "'" & Replace(strType, "'", "''") & "'"
'    End If
'
'    rsType.MoveNext
'
'  Loop
'  rsType.Close
'  Set rsType = Nothing

  If msAbsenceBreakdownTypes <> vbNullString Then
    msAbsenceBreakdownTypes = UCase(msAbsenceBreakdownTypes)
  End If

  'mdtReportStartDate = objAbsBreakdown.StartDate
  'mdtReportEndDate = objAbsBreakdown.EndDate
  'If mblnCustomDates Then
  '  mdtReportStartDate = datGeneral.GetValueForRecordIndependantCalc(mlngStartDateExprID)
  '  mdtReportEndDate = datGeneral.GetValueForRecordIndependantCalc(mlngEndDateExprID)
  'Else
  '  mdtReportEndDate = DateAdd("d", Day(Date) * -1, Date)
  '  mdtReportStartDate = DateAdd("d", 1, DateAdd("yyyy", -1, mdtReportEndDate))
  'End If
  
  'Set objAbsBreakdown = Nothing
  
  
  If glngPersonnelTableID = 0 Then
    mstrErrorMessage = "Personnel module setup has not been completed."
    fOK = False
    Exit Sub
  End If
  
  mlngBaseTableID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE))
  If mlngBaseTableID = 0 Then
    mstrErrorMessage = "Absence module setup has not been completed."
    fOK = False
    Exit Sub
  End If


  mstrBaseTable = datGeneral.GetTableName(mlngBaseTableID)
  mlngRecordDescExprID = datGeneral.GetRecDescExprID(mlngBaseTableID)
  
  
  
  ' Load the appropraite records
  If lngPersonnelID > 0 Then
    mstrPicklistFilter = CStr(lngPersonnelID)
  Else
    Select Case mstrPicklistFilterType
    Case "A"
      mstrPicklistFilter = vbNullString
    Case "P"
      mstrPicklistFilter = GetPicklistFilterSelect(mlngPicklistFilterID, 0)
    Case "F"
      mstrPicklistFilter = GetPicklistFilterSelect(0, mlngPicklistFilterID)
    End Select
  End If
  
  'Ascertain whether the user has decided to include Picklist/Filter info
  'mblnChkPicklistFilter = (GetUserSetting("AbsenceBreakdown", "PrintFilterHeader", "0") = "1")

  If fOK = False Then
    Exit Sub
  End If

  lngHorColID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTSESSION))
  lngVerColID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE))

  mlngColID(HOR) = lngHorColID
  mstrColName(HOR) = datGeneral.GetColumnName(lngHorColID)
  mlngColDataType(HOR) = datGeneral.GetDataType(mlngBaseTableID, lngHorColID)
  mstrFormat(HOR) = GetFormat(mlngColID(HOR))
  
  mlngColID(VER) = lngVerColID
  mstrColName(VER) = datGeneral.GetColumnName(lngVerColID)
  mlngColDataType(VER) = datGeneral.GetDataType(mlngBaseTableID, lngVerColID)
  mstrFormat(VER) = GetFormat(mlngColID(VER))

  mblnIntersection = False
  'MH20050105 Fault 9567
  'mlngIntersectionDecimals = 1
  mlngIntersectionDecimals = 2
  mblnShowAllPagesTogether = False
  
  
  'mlngOutputFormat = GetUserSetting("AbsenceBreakdown", "Format", 0)
  'mblnOutputScreen = GetUserSetting("AbsenceBreakdown", "Screen", 1)
  'mblnOutputPrinter = GetUserSetting("AbsenceBreakdown", "Printer", 0)
  'mstrOutputPrinterName = GetUserSetting("AbsenceBreakdown", "PrinterName", vbNullString)
  'mblnOutputSave = GetUserSetting("AbsenceBreakdown", "Save", 0)
  'mlngOutputSaveExisting = GetUserSetting("AbsenceBreakdown", "SaveExisting", -1)
  'mblnOutputEmail = GetUserSetting("AbsenceBreakdown", "Email", 0)
  'mlngOutputEmailAddr = GetUserSetting("AbsenceBreakdown", "EmailAddr", 0)
  'mstrOutputEmailSubject = GetUserSetting("AbsenceBreakdown", "EmailSubject", vbNullString)
  'mstrOutputFilename = GetUserSetting("AbsenceBreakdown", "FileName", vbNullString)

  mblnPreviewOnScreen = (mblnPreviewOnScreen Or (mlngOutputFormat = fmtDataOnly And mblnOutputScreen))

End Sub

Public Sub SizeBreakdownColumns()

  Dim lngWidth As Long
  Dim intCount As Integer

  With frmBreakDown.SSDBGrid1

    'If .Rows > 11 Then
    If .Rows > 10 Then
      If .ScrollBars = ssScrollBarsNone Then
        .Columns(0).Width = .Columns(0).Width - 240
        .ScrollBars = ssScrollBarsVertical
      End If
      lngWidth = .Width - 240
    Else
      If .ScrollBars = ssScrollBarsVertical Then
        .Columns(0).Width = .Columns(0).Width + 240
        .ScrollBars = ssScrollBarsNone
      End If
      lngWidth = .Width
    End If

    For intCount = .Cols - 1 To 1 Step -1
    
      If mlngCrossTabType = cttStability Or cttTurnover Then
        .Columns(intCount).Width = 1400
      End If
      
      lngWidth = lngWidth - .Columns(intCount).Width
    Next
    .Columns(0).Width = lngWidth

  End With

End Sub
'Public Sub AbsenceBreakdownSizeColumns()
'
'  Dim lngWidth As Long
'  Dim intCount As Integer
'
'  With frmCrossTabRun.SSDBGrid1
'
''    lngWidth = (.Width - .Columns(1).Width - 5) / (.Columns.Count - 1)
''
'    For intCount = 1 To .Columns.Count - 1
'      .Columns(intCount).Width = 1000
'    Next intCount
'
'  End With
'
'End Sub




Private Sub GetSQL2(strCol() As String)
  
  Dim objTableView As CTablePrivilege
  Dim objColumnPrivileges As CColumnPrivileges
  Dim sRealSource As String
  Dim sSource As String
  Dim lngCount As Long
  Dim fColumnOK As Boolean
  Dim alngTableViews() As Long
  Dim iNextIndex As Integer
  Dim fFound As Boolean

  Dim sCaseStatement As String
  'Dim sWhereColumn As String
  Dim strSelectedRecords As String
  'Dim sWhereIDs As String
  'Dim blnOK As Boolean
  Dim iCount As Integer
  Dim strCode As String

  On Error GoTo LocalErr
  
  fOK = True
  ReDim alngTableViews(2, 0)
  
  mstrSQLFrom = gcoTablePrivileges.Item(mstrBaseTable).RealSource
  mstrSQLSelect = vbNullString
  mstrSQLJoin = vbNullString
  ReDim asViews(0)


  For lngCount = 0 To UBound(strCol, 2)

    Set objColumnPrivileges = GetColumnPrivileges(mstrBaseTable)
    fColumnOK = objColumnPrivileges.IsValid(strCol(1, lngCount))
    If fColumnOK Then
      fColumnOK = objColumnPrivileges.Item(strCol(1, lngCount)).AllowSelect
      
      If fColumnOK Then
        fColumnOK = gcoTablePrivileges.Item(mstrBaseTable).AllowSelect
      End If
    
    End If
    
    Set objColumnPrivileges = Nothing
    
    If fColumnOK Then
      ' The column can be read from the base table/view, or directly from a parent table.
      ' Add the column to the column list.

'      If strSelectedRecords = vbNullString And mstrPicklistFilter <> vbNullString Then
'
'        If mlngCrossTabType = cttAbsenceBreakdown Then
'          strSelectedRecords = mstrSQLFrom & ".ID_" & Trim(Str(glngPersonnelTableID)) & " IN (" & mstrPicklistFilter & ")"
'        Else
'          strSelectedRecords = mstrSQLFrom & ".ID IN (" & mstrPicklistFilter & ")"
'        End If
'
'      End If

      mstrSQLSelect = mstrSQLSelect & _
        IIf(Len(mstrSQLSelect) > 0, ", ", "") & _
        mstrSQLFrom & "." & strCol(1, lngCount) & _
        " AS '" & strCol(2, lngCount) & "'"

    Else

      ReDim asViews(0)
      For Each objTableView In gcoTablePrivileges.Collection
          
        'Loop thru all of the views for this table where the user has select access
        If (Not objTableView.IsTable) And _
          (objTableView.TableID = mlngBaseTableID) And _
          (objTableView.AllowSelect) Then
            
          sSource = objTableView.ViewName

          ' Get the column permission for the view.
          Set objColumnPrivileges = GetColumnPrivileges(sSource)

          If objColumnPrivileges.IsValid(strCol(1, lngCount)) Then
            If objColumnPrivileges.Item(strCol(1, lngCount)).AllowSelect Then
              ' Add the view info to an array to be put into the column list or order code below.
              iNextIndex = UBound(asViews) + 1
              ReDim Preserve asViews(iNextIndex)
              asViews(iNextIndex) = sSource


              '=== This is the join code section ===
              ' Add the view to the Join code.
              ' Check if the view has already been added to the join code.
              fFound = False
              For iNextIndex = 1 To UBound(alngTableViews, 2)
                If alngTableViews(2, iNextIndex) = objTableView.ViewID Then
                  fFound = True
                  Exit For
                End If
              Next iNextIndex

              If Not fFound Then
                ' The view has not yet been added to the join code, so add it to the array and the join code.
                ' (also include the picklist info)

                iNextIndex = UBound(alngTableViews, 2) + 1
                ReDim Preserve alngTableViews(2, iNextIndex)
                alngTableViews(1, iNextIndex) = 1
                alngTableViews(2, iNextIndex) = objTableView.ViewID

                mstrSQLJoin = mstrSQLJoin & vbCrLf & _
                  " LEFT OUTER JOIN " & sSource & _
                  " ON " & mstrSQLFrom & ".ID = " & sSource & ".ID"

                strSelectedRecords = strSelectedRecords & _
                  IIf(strSelectedRecords <> vbNullString, " OR ", vbNullString) & _
                  mstrSQLFrom & ".ID IN (SELECT ID FROM " & sSource & ")"

'                'If mstrPicklistFilter <> vbNullString Then
'                  strSelectedRecords = strSelectedRecords & _
'                      IIf(strSelectedRecords <> vbNullString, " OR ", vbNullString) & "(" & _
'                      IIf(mstrPicklistFilter <> vbNullString, sSource & ".ID IN (" & mstrPicklistFilter & ") AND ", vbNullString) & _
'                      sSource & ".ID > 0)"
'                'End If
                
              End If
            End If
            '=== End of Join Code ===
              
              
            Set objColumnPrivileges = Nothing
          End If

        End If
      Next objTableView
      Set objTableView = Nothing

      ' The current user does have permission to 'read' the column through a/some view(s) on the
      ' table.
      If UBound(asViews) = 0 Then
        fOK = False
        'MH20010716 Fault 2497
        'If its the ID column they they don't have any access to the table.
        'mstrErrorMessage = "You do not have permission to see the column '" & strCol(1, lngCount) & "' " & _
                            "either directly or through any views." & vbCrLf
        mstrErrorMessage = "You do not have permission to see the " & _
                            IIf(strCol(1, lngCount) = "ID", "table '" & mstrBaseTable, "column '" & strCol(1, lngCount)) & _
                            "' either directly or through any views." & vbCrLf
        Exit Sub
      Else

'MH20071106 Fault 12585
''        ' Add the column to the column list.
''        sCaseStatement = "CASE"
''        sWhereColumn = vbNullString
''        For iNextIndex = 1 To UBound(asViews)
''          sCaseStatement = sCaseStatement & _
''            " WHEN NOT " & asViews(iNextIndex) & "." & strCol(1, lngCount) & " IS NULL THEN " & asViews(iNextIndex) & "." & strCol(1, lngCount) & vbCrLf
''        Next iNextIndex
''
''        If Len(sCaseStatement) > 0 Then
''          sCaseStatement = sCaseStatement & _
''            " ELSE NULL END AS " & _
''            "'" & strCol(2, lngCount) & "'"
''
''          mstrSQLSelect = mstrSQLSelect & _
''            IIf(Len(mstrSQLSelect) > 0, ", ", "") & vbCrLf & _
''            sCaseStatement
''        End If
      
        ' Add the column to the column list.
        If UBound(asViews) = 1 Then
          mstrSQLSelect = mstrSQLSelect & _
            IIf(Len(mstrSQLSelect) > 0, ", ", "") & vbCrLf & _
            asViews(1) & "." & strCol(1, lngCount) & " AS '" & strCol(2, lngCount) & "'"
        Else
          sCaseStatement = ""
          For iNextIndex = 1 To UBound(asViews)
            sCaseStatement = sCaseStatement & _
              IIf(sCaseStatement <> "", vbCrLf & " , ", "") & _
              asViews(iNextIndex) & "." & strCol(1, lngCount)
          Next iNextIndex
  
          If Len(sCaseStatement) > 0 Then
            mstrSQLSelect = mstrSQLSelect & _
              IIf(Len(mstrSQLSelect) > 0, ", ", "") & vbCrLf & _
              "COALESCE(" & sCaseStatement & ")" & vbCrLf & _
              "AS '" & strCol(2, lngCount) & "'"
          End If
        End If
      
      End If
    End If
  Next

  Select Case mlngCrossTabType
  Case cttAbsenceBreakdown
      If msAbsenceBreakdownTypes <> vbNullString Then
        mstrSQLWhere = mstrSQLWhere & _
            IIf(mstrSQLWhere <> vbNullString, " AND ", " WHERE ") & _
            "(UPPER(" + gsAbsenceTypeColumnName + ") IN (" & msAbsenceBreakdownTypes & "))"
      End If
      
      'MH20060619 Fault
      mstrSQLWhere = mstrSQLWhere & _
          IIf(mstrSQLWhere <> vbNullString, " AND ", " WHERE ") & _
          "( " & gsAbsenceStartDateColumnName & _
          " <= CONVERT(datetime, '" & FormatDateSQL(mdtReportEndDate) + "'))" & _
          "And (" & gsAbsenceEndDateColumnName & _
          " >= CONVERT(datetime, '" & FormatDateSQL(mdtReportStartDate) + "') OR " & gsAbsenceEndDateColumnName & " IS NULL)"
          '" >= CONVERT(datetime, '" & FormatDateSQL(mdtReportStartDate) + "'))"
  
  'MH20040113 Fault 7234 - Disregard records outside of the report period...
  Case cttStability, cttTurnover

    ' JDM - Fault 8137 - Problems if accessing through views only
    If UBound(asViews) = 0 Then
      mstrSQLWhere = mstrSQLWhere & _
          IIf(mstrSQLWhere <> vbNullString, " AND ", " WHERE ") & _
          "(Datediff(d,'" & FormatDateSQL(IIf(mblnIncludeNewStarters, mdtReportEndDate, mdtReportStartDate)) & _
          "'," & gsPersonnelStartDateColumnName & ") <= 0 OR " & gsPersonnelStartDateColumnName & " IS NULL)" & _
          " AND " & _
          "(Datediff(d,'" & FormatDateSQL(mdtReportStartDate) & _
          "'," & gsPersonnelLeavingDateColumnName & ") >= 0 OR " & gsPersonnelLeavingDateColumnName & " IS NULL)"
    Else

      ' Loop through accessible views building the code
      strCode = ""
      For iCount = 1 To UBound(asViews)
        strCode = strCode & _
            IIf(strCode <> vbNullString, vbCrLf & " AND " & vbCrLf & "(", " (") & _
            "(Datediff(d,'" & FormatDateSQL(IIf(mblnIncludeNewStarters, mdtReportEndDate, mdtReportStartDate)) & _
            "'," & asViews(iCount) & "." & gsPersonnelStartDateColumnName & ") <= 0 OR " & asViews(iCount) & "." & gsPersonnelStartDateColumnName & " IS NULL)" & _
            " AND " & _
            "(Datediff(d,'" & FormatDateSQL(mdtReportStartDate) & _
            "'," & asViews(iCount) & "." & gsPersonnelLeavingDateColumnName & ") >= 0 OR " & asViews(iCount) & "." & gsPersonnelLeavingDateColumnName & " IS NULL))"
      Next iCount
      mstrSQLWhere = mstrSQLWhere & IIf(mstrSQLWhere <> vbNullString, " AND (", " WHERE (") & strCode & ")"

    End If

  End Select

  If strSelectedRecords <> vbNullString Then
    mstrSQLWhere = mstrSQLWhere & _
        IIf(mstrSQLWhere <> vbNullString, " AND ", " WHERE ") & _
        "(" & strSelectedRecords & ")"
  End If

Exit Sub

LocalErr:
  mstrErrorMessage = "Error retrieving data"
  fOK = False

End Sub


Private Sub CreateTempTable()

  Dim strColumn() As String
  Dim strSQL As String
  Dim lngMax As Long

  On Error GoTo LocalErr

  lngMax = 2
  ReDim strColumn(2, lngMax) As String

  strColumn(1, 0) = "ID"
  strColumn(2, 0) = "ID"

  strColumn(1, 1) = mstrColName(HOR)
  strColumn(2, 1) = "Hor"

  strColumn(1, 2) = mstrColName(VER)
  strColumn(2, 2) = "Ver"

  If mblnPageBreak Then
    lngMax = lngMax + 1
    ReDim Preserve strColumn(2, lngMax) As String

    strColumn(1, lngMax) = mstrColName(PGB)
    strColumn(2, lngMax) = "Pgb"
  End If

  If mblnIntersection Then
    lngMax = lngMax + 1
    ReDim Preserve strColumn(2, lngMax) As String

    strColumn(1, lngMax) = mstrColName(INS)
    strColumn(2, lngMax) = "Ins"
  End If

  If mlngCrossTabType <> cttNormal Then
    
    If mlngCrossTabType = cttAbsenceBreakdown Then
      lngMax = lngMax + 7
      ReDim Preserve strColumn(2, lngMax) As String

      strColumn(1, lngMax) = gsAbsenceDurationColumnName
      strColumn(2, lngMax) = "Value"

      strColumn(1, lngMax - 4) = gsAbsenceStartDateColumnName
      strColumn(2, lngMax - 4) = "Start_Date"

      strColumn(1, lngMax - 3) = gsAbsenceStartSessionColumnName
      strColumn(2, lngMax - 3) = "Start_Session"

      strColumn(1, lngMax - 2) = gsAbsenceEndDateColumnName
      strColumn(2, lngMax - 2) = "End_Date"

      strColumn(1, lngMax - 1) = gsAbsenceEndSessionColumnName
      strColumn(2, lngMax - 1) = "End_Session"

      strColumn(1, lngMax - 5) = "ID_" + Trim(Str(glngPersonnelTableID))
      strColumn(2, lngMax - 5) = "Personnel_ID"

      strColumn(1, lngMax - 6) = gsAbsenceDurationColumnName  ' Used to hold the day number (1=Mon, 2=Tues etc.)
      strColumn(2, lngMax - 6) = "Day_Number"

    Else
      lngMax = lngMax + 2
      ReDim Preserve strColumn(2, lngMax) As String

      strColumn(1, lngMax - 1) = gsPersonnelStartDateColumnName
      strColumn(2, lngMax - 1) = "StartDate"

      strColumn(1, lngMax) = gsPersonnelLeavingDateColumnName
      strColumn(2, lngMax) = "LeavingDate"


    End If

  End If


  Call GetSQL2(strColumn())
  If fOK = False Then
    Exit Sub
  End If
  
  mstrTempTableName = datGeneral.UniqueSQLObjectName("ASRSysTempCrossTab", 3)
  mstrSQLSelect = mstrSQLSelect & ", " & _
    "space(255) as 'RecDesc' INTO " & mstrTempTableName
  
  
  'MH20071106 Fault 12585
  If mstrPicklistFilter <> vbNullString Then
    mstrSQLWhere = mstrSQLWhere & _
      IIf(mstrSQLWhere <> vbNullString, " AND ", " WHERE ")
    
    If mlngCrossTabType = cttAbsenceBreakdown Then
      mstrSQLWhere = mstrSQLWhere & _
        mstrSQLFrom & ".ID_" & Trim(Str(glngPersonnelTableID)) & " IN (" & mstrPicklistFilter & ")"
    Else
      mstrSQLWhere = mstrSQLWhere & _
        mstrSQLFrom & ".ID IN (" & mstrPicklistFilter & ")"
    End If
  End If
  
  
  strSQL = "SELECT " & mstrSQLSelect & vbCrLf & _
           " FROM " & mstrSQLFrom & vbCrLf & _
           mstrSQLJoin & vbCrLf & _
           mstrSQLWhere
  
  
  'MH20010327 Seems that it might be moving on pass this line of code too
  'quickly so I've tried returning the number of rows effected to make
  'sure that it completes fully
  'datData.Execute strSQL
  datData.ExecuteSqlReturnAffected strSQL

  'Dim tt As Double
  'tt = Timer + 2
  'Do While Timer < tt
  '  DoEvents
  'Loop

'  Set mcmdUpdateRecDescs = New ADODB.Command
'  mcmdUpdateRecDescs.ActiveConnection = gADOCon
'  mcmdUpdateRecDescs.CommandText = strSQL
'  mcmdUpdateRecDescs.CommandTimeout = 0
'  mcmdUpdateRecDescs.Execute , , adAsyncExecute
'
'  Do While mcmdUpdateRecDescs.State = adStateExecuting
'    DoEvents
'  Loop
'
'  Set mcmdUpdateRecDescs = Nothing
  
  
  strSQL = "SELECT * FROM " & mstrTempTableName
  
'  If mlngCrossTabType <> cttNormal Then
'    If mlngCrossTabType <> cttAbsenceBreakdown Then
'      strSQL = strSQL & " WHERE " & _
'        SQLEmployedAtStartOfReport("startdate", "leavingdate")
'    End If
'  End If
  
'  Select Case mlngCrossTabType
'  Case cttStability
'    strSQL = strSQL & " WHERE " & _
'      SQLEmployedAtStartOfReport("startdate", "leavingdate")
'  Case cttTurnover
'    strSQL = strSQL & " WHERE " & _
'      "(Datediff(d,'" & Format(mdtReportStartDate, "mm/dd/yyyy") & _
'      "',leavingdate) > 0 OR leavingdate IS NULL)"
'  End Select
  
  
  Set rsCrossTabData = New Recordset
  rsCrossTabData.ActiveConnection = gADOCon
  rsCrossTabData.Properties("Preserve On Commit") = True
  rsCrossTabData.Properties("Preserve On Abort") = True
  rsCrossTabData.Open strSQL, , adOpenKeyset, adLockReadOnly, adCmdText

  If rsCrossTabData.EOF Then
    mstrErrorMessage = "No records meet selection criteria."
    mblnNoRecords = True
    fOK = False
  End If

  
  'Check if we might need record description...
  If Not gblnBatchMode And mlngRecordDescExprID > 0 Then
    Set mcmdUpdateRecDescs = New ADODB.Command
    mcmdUpdateRecDescs.ActiveConnection = gADOCon
    mcmdUpdateRecDescs.CommandText = "EXEC dbo.sp_ASRCrossTabsRecDescs '" & mstrTempTableName & "', " & CStr(mlngRecordDescExprID)
    mcmdUpdateRecDescs.CommandTimeout = 0
    
    ' JDM - Fault 2421 - Must have record description calculated NOW.
    'If mlngCrossTabType = cttAbsenceBreakdown Then
      mcmdUpdateRecDescs.Execute
    'Else
    '  mcmdUpdateRecDescs.Execute , , adAsyncExecute
    'End If
  End If

Exit Sub

LocalErr:
  mstrErrorMessage = "Error retrieving data" & vbCrLf & _
                      "(" & Err.Description & ")"
  fOK = False

End Sub


Private Function CheckIfGotRecDescs(Optional strWhere As String)

  Dim rsTemp As Recordset
  Dim strSQL As String

  'Check if we've got all of the record descriptions...
  'And that we are not still trying to work them out
  CheckIfGotRecDescs = False
  
  If mlngRecordDescExprID = 0 Then
    MsgBox "Unable to show cell breakdown details as no record description " & _
           "has been set up for the '" & mstrBaseTable & "' table.", vbInformation, Me.Caption
    Exit Function
  End If


  strSQL = "SELECT Count(ID) FROM " & mstrTempTableName & _
           IIf(strWhere = vbNullString, " WHERE ", strWhere & " AND ") & _
           "RecDesc = space(255)"
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  
  If rsTemp.Fields(0).Value > 0 Then

    'gobjProgress.AviFile = App.Path & "\videos\crosstab.avi"
    gobjProgress.AVI = dbText
    gobjProgress.MainCaption = "Cross Tab"
    gobjProgress.Bar1Caption = "Getting record descriptions..."
    gobjProgress.Bar1MaxValue = rsTemp.Fields(0).Value
    gobjProgress.Caption = Me.Caption
    gobjProgress.NumberOfBars = 1
    gobjProgress.OpenProgress
    gobjProgress.Time = False

    strSQL = "SELECT Count(ID) FROM " & mstrTempTableName & _
             IIf(strWhere = vbNullString, " WHERE ", strWhere & " AND ") & _
             "RecDesc = space(255)"

    Do While rsTemp.Fields(0).Value > 0
      If gobjProgress.Cancelled Then
        gobjProgress.CloseProgress
        Exit Function
      End If
      
      DoEvents
      Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
      gobjProgress.Bar1Value = gobjProgress.Bar1MaxValue - rsTemp.Fields(0).Value
      gobjProgress.UpdateProgress
    Loop
    
    If Not gblnBatchMode Then
      gobjProgress.CloseProgress
    End If
  End If

  rsTemp.Close
  CheckIfGotRecDescs = True

End Function

Private Sub cmdOutput_Click()
  OutputReport True
End Sub


Private Function OutputReport(blnPrompt As Boolean) As Boolean

  Dim objOutput As clsOutputRun
  Dim objColumn As clsColumn
  
  Dim rsTemp As Recordset
  Dim strSQL As String

  Dim strTemp As String
  Dim lngPage As Long
  Dim lngSelectedPage As Long
  Dim lngIndex As Long
  Dim strOutput() As String
  Dim strPageValue As String

  Dim lngNumCols As Long
  Dim lngNumRows As Long
  Dim lngCol As Integer
  Dim lngRow As Integer
  Dim lngTYPE As Long
  Dim lngGroupNum As Long
  Dim blnThousandSep As Boolean
  
  Dim lngColumn0Width As Long
  Dim lngTextWidth As Long
  Dim lngExcelDataType As Long

  Dim strDefTitle As String

  blnThousandSep = (chkThousandSeparators.Value = vbChecked)


  Set objOutput = New clsOutputRun


  If Not mblnShowAllPagesTogether And blnPrompt And mblnPageBreak Then
    With objOutput.cboPageBreak
      .Clear
      .AddItem "<All>"
      .ItemData(.NewIndex) = -1

      For lngIndex = 0 To UBound(mvarHeadings(PGB))
        .AddItem mvarHeadings(PGB)(lngIndex)
        .ItemData(.NewIndex) = lngIndex
      Next

    End With
  End If

  If objOutput.SetOptions _
      (blnPrompt, mlngOutputFormat, mblnOutputScreen, _
      mblnOutputPrinter, mstrOutputPrinterName, _
      mblnOutputSave, mlngOutputSaveExisting, _
      mblnOutputEmail, mlngOutputEmailAddr, mstrOutputEmailSubject, _
      mstrOutputEmailAttachAs, mstrOutputFilename) Then

    If Not gblnBatchMode Then
      If mlngCrossTabType = cttNormal Then
        objOutput.OpenProgress "Cross Tab", mstrCrossTabName, UBound(mvarHeadings(PGB)) + 2
      Else
        objOutput.OpenProgress mstrCrossTabName, vbNullString, UBound(mvarHeadings(PGB)) + 2
      End If
    End If

    If objOutput.GetFile Then

      If objOutput.Format = fmtExcelPivotTable Then 'And mlngCrossTabType = cttNormal Then

        objOutput.PivotSuppressBlanks = (chkSuppressZeros.Value = vbChecked)
        'AE20071018 Fault #12540
        objOutput.PivotDataFunction = cboType.Text
        
        objOutput.AddColumn " ", sqlVarChar, 0
        For lngIndex = 0 To UBound(mvarHeadings(HOR))
          strTemp = mvarHeadings(HOR)(lngIndex)
          objOutput.AddColumn strTemp, sqlNumeric, mlngIntersectionDecimals, blnThousandSep
        Next
        objOutput.AddColumn cboType.Text, sqlNumeric, mlngIntersectionDecimals, blnThousandSep

        strSQL = "SELECT HOR as 'Horizontal', VER as 'Vertical'" & _
            IIf(mblnPageBreak, ", PGB as 'Page Break'", vbNullString) & _
            ", RecDesc as 'Record Description'" & _
            IIf(mblnIntersection, ", Ins as 'Intersection'", vbNullString) & _
            IIf(mlngCrossTabType = cttAbsenceBreakdown, ", Value as 'Duration'", vbNullString) & _
            " FROM " & mstrTempTableName

        If mlngCrossTabType = cttAbsenceBreakdown Then
          strSQL = strSQL & _
              " WHERE NOT HOR IN ('Total','Count','Average')"

        ElseIf mblnPageBreak Then
          If objOutput.cboPageBreak.ListIndex > 0 Then
            strSQL = strSQL & " WHERE " & mvarSearches(PGB)(objOutput.cboPageBreak.ListIndex - 1)
          End If
          strSQL = strSQL & " ORDER BY PGB"
        End If
        
        Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)


        With rsTemp
          If Not mblnPageBreak Then
            lngRow = 1
            ReDim strOutput(.Fields.Count - 1, 0)
            For lngCol = 0 To .Fields.Count - 1
              strOutput(lngCol, 0) = .Fields(lngCol).Name
            Next
          End If

          .MoveFirst
          Do While Not .EOF
            
            If mblnPageBreak Then
              If strPageValue <> .Fields("Page Break").Value Then

                If strPageValue <> vbNullString Then
                  objOutput.AddPage Replace(SSDBGrid1.Caption, "&&", "&"), strPageValue
                  objOutput.DataArray strOutput
                End If
                strPageValue = .Fields("Page Break").Value
                
                lngRow = 1
                ReDim strOutput(.Fields.Count - 1, 0)
                For lngCol = 0 To .Fields.Count - 1
                  strOutput(lngCol, 0) = .Fields(lngCol).Name
                Next

              End If
            End If
            
            ReDim Preserve strOutput(.Fields.Count - 1, lngRow)
            For lngCol = 0 To .Fields.Count - 1

              'MH20070226 Fault 11961
              'If lngCol <= UBound(mvarHeadings) Then
              If lngCol < 2 Or (lngCol = 2 And mblnPageBreak) Then

                lngGroupNum = GetGroupNumber(CStr(IIf(IsNull(.Fields(lngCol).Value), vbNullString, .Fields(lngCol).Value)), lngCol)
                strOutput(lngCol, lngRow) = mvarHeadings(lngCol)(lngGroupNum)
              Else
                strOutput(lngCol, lngRow) = .Fields(lngCol).Value
              End If
            Next
            lngRow = lngRow + 1
            .MoveNext
          Loop
        End With

        objOutput.AddPage Replace(SSDBGrid1.Caption, "&&", "&"), IIf(strPageValue <> vbNullString, strPageValue, mstrCrossTabName)
        objOutput.DataArray strOutput

      Else

        'MH20040218 Fault
        lngExcelDataType = IIf(chkPercentage.Value = vbUnchecked, sqlNumeric, sqlUnknown)
        
        objOutput.HeaderCols = 1
        objOutput.PageTitles = mblnPageBreak
        If Not mblnShowAllPagesTogether Then
          objOutput.AddColumn " ", sqlVarChar, 0
          For lngIndex = 0 To UBound(mvarHeadings(HOR))
            strTemp = mvarHeadings(HOR)(lngIndex)

            'MH20040129 Fault 7998
            If mlngCrossTabType = cttAbsenceBreakdown And lngIndex = 8 Then
              objOutput.AddColumn strTemp, lngExcelDataType, 2, blnThousandSep
            Else
              objOutput.AddColumn strTemp, lngExcelDataType, mlngIntersectionDecimals, blnThousandSep
            End If

          Next
          objOutput.AddColumn cboType.Text, lngExcelDataType, mlngIntersectionDecimals, blnThousandSep
    
          
          If mblnPageBreak Then
            With objOutput.cboPageBreak
              lngSelectedPage = -1
              If .ListIndex <> -1 Then
                lngSelectedPage = .ItemData(.ListIndex)
              End If
            End With
          
            SSDBGrid1.Redraw = False

            lngColumn0Width = SSDBGrid1.Columns(0).Width
            For lngPage = 0 To UBound(mvarHeadings(PGB))
              If lngSelectedPage = -1 Or lngSelectedPage = lngPage Then
                objOutput.AddPage Replace(SSDBGrid1.Caption, "&&", "&"), CStr(mvarHeadings(PGB)(lngPage))
                BuildOutputStrings lngPage
                OutputGrid

                If objOutput.Format = fmtDataOnly Then
                  With SSDBGrid1.Columns(0)
                    .Caption = mstrColName(PGB) & " : " & Trim(mvarHeadings(PGB)(lngPage))
                    lngTextWidth = Me.TextWidth(.Caption) + 180
                    .Width = IIf(lngTextWidth > lngColumn0Width, lngTextWidth, lngColumn0Width)
                  End With
                End If

                objOutput.DataGrid SSDBGrid1
              End If
              If Not gblnBatchMode Then
                If gobjProgress.Cancelled Then
                  Exit For
                End If
                gobjProgress.UpdateProgress gblnBatchMode
              End If
            Next
      
            If objOutput.Format = fmtDataOnly Then
              With SSDBGrid1.Columns(0)
                .Caption = vbNullString
                .Width = lngColumn0Width
              End With
            End If
            SSDBGrid1.Redraw = True
          Else
            objOutput.AddPage Replace(SSDBGrid1.Caption, "&&", "&"), mstrCrossTabName
            objOutput.DataGrid SSDBGrid1
          End If
    
        Else
          objOutput.AddColumn " ", sqlVarChar, 0
          For lngPage = 0 To UBound(mvarHeadings(PGB)) + 1
            For lngIndex = 0 To UBound(mvarHeadings(HOR))
              strTemp = mvarHeadings(HOR)(lngIndex)
              
              If mlngCrossTabType = cttAbsenceBreakdown Then
                objOutput.AddColumn strTemp, sqlNumeric, 1, blnThousandSep
              Else
                Select Case lngIndex
                Case 0
                  If mlngCrossTabType = cttTurnover Then
                    objOutput.AddColumn strTemp, sqlNumeric, 1, blnThousandSep
                  Else
                    objOutput.AddColumn strTemp, sqlNumeric, 0, blnThousandSep
                  End If
                Case 1
                  objOutput.AddColumn strTemp, sqlNumeric, 0, blnThousandSep
                Case 2
                  objOutput.AddColumn strTemp, sqlUnknown, 2, blnThousandSep
                End Select
              End If
    
            Next
            If Not gblnBatchMode Then
              If gobjProgress.Cancelled Then
                Exit For
              End If
            End If
          Next
    
          objOutput.AddPage Replace(SSDBGrid1.Caption, "&&", "&"), mstrCrossTabName    'CStr(mvarHeadings(PGB)(lngPage))
          objOutput.DataGrid SSDBGrid1
          gobjProgress.UpdateProgress gblnBatchMode
    
        End If
      End If
    End If
    
    SSDBGrid1.Redraw = True
    
    mblnUserCancelled = objOutput.UserCancelled
    If Not gblnBatchMode Then
      gobjProgress.CloseProgress
    End If
    objOutput.Complete

    mstrErrorMessage = objOutput.ErrorMessage
    fOK = (mstrErrorMessage = vbNullString)
  
  Else
    blnPrompt = (blnPrompt And Not objOutput.UserCancelled)
    mstrErrorMessage = objOutput.ErrorMessage
    fOK = (mstrErrorMessage = vbNullString)
  
  End If

  If blnPrompt Then
    gobjProgress.CloseProgress

    If mlngCrossTabType = cttNormal Then
      strDefTitle = "Cross Tab : '" & mstrCrossTabName & "'"
    Else
      strDefTitle = mstrCrossTabName
    End If

    'MH20040302 Fault 8143
    'Not ideal but the only way to prevent a runtime error was a doevents
    DoEvents

    If fOK Then
      MsgBox strDefTitle & " output complete.", _
          vbInformation, Me.Caption
    Else
      MsgBox strDefTitle & " output failed." & vbCrLf & vbCrLf & mstrErrorMessage, _
          vbExclamation, Me.Caption
    End If

  End If

  SSDBGrid1.Redraw = True
  
  Set objOutput = Nothing
  OutputReport = fOK

End Function


Public Property Get PreviewOnScreen() As Boolean
  PreviewOnScreen = (fOK And mblnPreviewOnScreen And Not gblnBatchMode And Not mblnNoRecords)
End Property


Public Sub GetReportConfig(strReportType As String)

  'Dim objAbsBreakdown As clsAbsenceBreakdown

  Dim rsType As Recordset
  Dim strSQL As String
  Dim strType As String
  
  Dim lngStartDateExprID As Long
  Dim lngEndDateExprID As Long

  strSQL = "SELECT * " & _
           "FROM " & gsAbsenceTypeTableName & " " & _
           "ORDER BY " & gsAbsenceTypeTypeColumnName
  Set rsType = datGeneral.GetReadOnlyRecords(strSQL)

  'Set objAbsBreakdown = New clsAbsenceBreakdown
  'objAbsBreakdown.ReportType = "AbsenceBreakdown"

  msAbsenceBreakdownTypes = vbNullString
  Do Until rsType.EOF

    strType = rsType.Fields(gsAbsenceTypeTypeColumnName).Value
    If GetSystemSetting("AbsenceBreakdown", "Absence Type " & strType, vbNullString) = "1" Then
      msAbsenceBreakdownTypes = _
        IIf(msAbsenceBreakdownTypes <> vbNullString, msAbsenceBreakdownTypes & ", ", "") & _
        "'" & Replace(strType, "'", "''") & "'"
    End If
    
    rsType.MoveNext

  Loop
  rsType.Close
  Set rsType = Nothing

  If msAbsenceBreakdownTypes <> vbNullString Then
    msAbsenceBreakdownTypes = UCase(msAbsenceBreakdownTypes)
  End If
  
  If GetSystemSetting(strReportType, "Custom Dates", "0") = "1" Then
    lngStartDateExprID = GetSystemSetting(strReportType, "Start Date", 0)
    mstrErrorMessage = IsCalcValid(lngStartDateExprID)
    If mstrErrorMessage <> vbNullString Then
      fOK = False
      Exit Sub
    End If
    mdtReportStartDate = datGeneral.GetValueForRecordIndependantCalc(lngStartDateExprID)
    
    lngEndDateExprID = GetSystemSetting(strReportType, "End Date", 0)
    mstrErrorMessage = IsCalcValid(lngEndDateExprID)
    If mstrErrorMessage <> vbNullString Then
      fOK = False
      Exit Sub
    End If
    mdtReportEndDate = datGeneral.GetValueForRecordIndependantCalc(lngEndDateExprID)

    'MH20030911 Fault 5991
    If DateDiff("d", mdtReportStartDate, mdtReportEndDate) < 0 Then
      mstrErrorMessage = "The report end date is before the report start date."
      fOK = False
      Exit Sub
    End If

  Else
    mdtReportEndDate = DateAdd("d", Day(Date) * -1, Date)
    mdtReportStartDate = DateAdd("d", 1, DateAdd("yyyy", -1, mdtReportEndDate))
  End If

  mlngPicklistFilterID = GetSystemSetting(strReportType, "ID", "0")
  mstrPicklistFilterType = GetSystemSetting(strReportType, "Type", "A")
  mlngVerCol = GetSystemSetting(strReportType, "VerColID", 0)
  mlngHorCol = GetSystemSetting(strReportType, "HorColID", 0)
  mlngOutputFormat = GetSystemSetting(strReportType, "Format", 0)
  mblnOutputScreen = (GetSystemSetting(strReportType, "Screen", 1) = 1)
  mblnOutputPrinter = (GetSystemSetting(strReportType, "Printer", 0) = 1)
  mstrOutputPrinterName = GetSystemSetting(strReportType, "PrinterName", vbNullString)
  mblnOutputSave = (GetSystemSetting(strReportType, "Save", 0) = 1)
  mlngOutputSaveExisting = GetSystemSetting(strReportType, "SaveExisting", -1)
  mblnOutputEmail = (GetSystemSetting(strReportType, "Email", 0) = 1)
  mlngOutputEmailAddr = GetSystemSetting(strReportType, "EmailAddr", 0)
  mstrOutputEmailSubject = GetSystemSetting(strReportType, "EmailSubject", vbNullString)
  mstrOutputEmailAttachAs = GetSystemSetting(strReportType, "EmailAttachAs", vbNullString)
  mstrOutputFilename = GetSystemSetting(strReportType, "FileName", vbNullString)
  mblnPreviewOnScreen = (GetSystemSetting(strReportType, "Preview", True) Or (mlngOutputFormat = fmtDataOnly And mblnOutputScreen))
  mblnChkPicklistFilter = (GetSystemSetting(strReportType, "PrintFilterHeader", 0) = 1)

End Sub


Public Sub SetAbsenceBreakdownParameters( _
          dtStartDate As Date, _
          dtEndDate As Date, _
          lngPicklistFilterID As Long, _
          strPicklistFilterType As String, _
          sAbsenceBreakdownTypes As String)

  mdtReportStartDate = dtStartDate
  mdtReportEndDate = dtEndDate
  mlngPicklistFilterID = lngPicklistFilterID
  mstrPicklistFilterType = strPicklistFilterType
  msAbsenceBreakdownTypes = sAbsenceBreakdownTypes

End Sub


Public Sub SetTurnoverParameters( _
          dtStartDate As Date, _
          dtEndDate As Date, _
          lngPicklistFilterID As Long, _
          strPicklistFilterType As String, _
          lngVerCol As Long, _
          lngHorCol As Long, _
          blnIncludeNewStarters As Boolean)

  mdtReportStartDate = dtStartDate
  mdtReportEndDate = dtEndDate
  mlngPicklistFilterID = lngPicklistFilterID
  mstrPicklistFilterType = strPicklistFilterType
  mlngVerCol = lngVerCol
  mlngHorCol = lngHorCol
  mblnIncludeNewStarters = blnIncludeNewStarters

End Sub


Public Sub SetOutputParameters( _
          lngOutputFormat As Long, _
          blnOutputScreen As Boolean, _
          blnOutputPrinter As Boolean, _
          strOutputPrinterName As String, _
          blnOutputSave As Boolean, _
          lngOutputSaveExisting As Long, _
          blnOutputEmail As Boolean, _
          lngOutputEmailAddr As Long, _
          strOutputEmailSubject As String, _
          strOutputEmailAttachAs As String, _
          strOutputFilename As String, _
          blnPreviewOnScreen As Boolean, _
          blnChkPicklistFilter As Boolean)

  mlngOutputFormat = lngOutputFormat
  mblnOutputScreen = blnOutputScreen
  mblnOutputPrinter = blnOutputPrinter
  mstrOutputPrinterName = strOutputPrinterName
  mblnOutputSave = blnOutputSave
  mlngOutputSaveExisting = lngOutputSaveExisting
  mblnOutputEmail = blnOutputEmail
  mlngOutputEmailAddr = lngOutputEmailAddr
  mstrOutputEmailSubject = strOutputEmailSubject
  mstrOutputEmailAttachAs = strOutputEmailAttachAs
  mstrOutputFilename = strOutputFilename
  mblnPreviewOnScreen = blnPreviewOnScreen
  mblnChkPicklistFilter = blnChkPicklistFilter

End Sub


Private Function FormatDateSQL(dtTemp As Date) As String
  FormatDateSQL = Replace(Format(dtTemp, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/")
End Function

Private Function FormatString(ByVal sHeading As String) As String

  Dim sReturnValue As String

  sReturnValue = Left(Trim(Replace(sHeading, Chr(32), "")), 255)
  sReturnValue = Replace(sReturnValue, Chr(9), "")
  sReturnValue = Replace(sReturnValue, Chr(10), "")
  sReturnValue = Replace(sReturnValue, Chr(13), "")

  FormatString = sReturnValue

End Function

