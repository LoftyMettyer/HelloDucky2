VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmMatchRun 
   Caption         =   "Match Report"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1075
   Icon            =   "frmMatchRun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   7125
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   400
      Left            =   9360
      TabIndex        =   1
      Top             =   6600
      Width           =   1200
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "O&utput..."
      Height          =   400
      Left            =   8040
      TabIndex        =   0
      Top             =   6600
      Width           =   1200
   End
   Begin SSDataWidgets_B.SSDBGrid grdOutput 
      Height          =   6345
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   10425
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      Col.Count       =   0
      DividerType     =   0
      BevelColorFrame =   0
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
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
      _ExtentX        =   18389
      _ExtentY        =   11192
      _StockProps     =   79
      BackColor       =   -2147483643
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
Attribute VB_Name = "frmMatchRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmOutput As frmOutputOptions
Private mblnUserCancelled As Boolean
Private mstrTempTableName As String

Private mcolColDetails As Collection
Private mcolRelations As Collection
Private frmBreakDown As frmMatchRunBreakDown
Private mlngTableViews() As Long
Private mstrExcelFormats() As String
Private fOK As Boolean
'Private gblnBatchMode As Boolean
Private mstrErrorMessage As String
Private mblnNoRecords As Boolean
Private mbDefinitionOwner As Boolean
Private alngSourceTables() As Long

Private mlngMatchReportID As Long
Private mlngMatchReportType As MatchReportType
Private mstrName As String
Private mstrRecordSelectionName As String
Private mstrDescription As String
Private mlngNumRecords As Long
Private mblnEqualGrade As Boolean
Private mblnReportingStructure As Boolean

Private mlngScoreMode As Long
Private mblnScoreCheck As Boolean
Private mlngScoreLimit As Long

Private mlngTable1ID As Long
Private mstrTable1Name As String
Private mstrTable1RealSource As String
Private mlngTable1RecDescExprID As Long
Private mlngTable1AllRecords As Long
Private mlngTable1PickListID As Long
Private mlngTable1FilterID As Long
Private mstrTable1Where As String

Private mlngTable2ID As Long
Private mstrTable2Name As String
Private mstrTable2RealSource As String
Private mlngTable2RecDescExprID As Long
Private mlngTable2AllRecords As Long
Private mlngTable2PickListID As Long
Private mlngTable2FilterID As Long
Private mstrTable2Where As String

Private mstrSQL As String
Private mstrSQLWhere As String
Private mstrSQLGroupBy As String
Private mstrSQLOrderBy As String
Private mstrSQLMatchScore As String

Private mcolSQLSelect As Collection
Private mcolSQLJoin As Collection
Private mcolSQLWhere As Collection
Private mcolSQLOrderBy As Collection
Private mcolSQLMatchScore As Collection
'Private mcolSQLOrderBy As Collection

Private mblnPreviewOnScreen As Boolean

'New Default Output Variables
Private mlngOutputFormat As Long
Private mblnOutputScreen As Boolean
Private mblnOutputPrinter As Boolean
Private mstrOutputPrinterName As String
Private mblnOutputSave As Boolean
Private mlngOutputSaveExisting As Long
'Private mlngOutputSaveFormat As Long ' May need in future
Private mblnOutputEmail As Boolean
Private mlngOutputEmailAddr As Long
Private mstrOutputEmailSubject As String
Private mstrOutputEmailAttachAs As String
'Private mlngOutputEmailFileFormat As Long ' May need in future
Private mstrOutputFileName As String
Private mblnChkPicklistFilter As Boolean 'might not need
Private mstrOutputTitlePage As String
Private mstrOutputReportPackTitle As String
Private mstrOutputOverrideFilter As String
Private mblnOutputTOC As Boolean
Private mblnOutputCoverSheet As Boolean
Private mlngOverrideFilterID As Long

' Array holding the User Defined functions that are needed for this report
Private mastrUDFsRequired() As String


Public Property Let MatchReportID(lngNewValue As Long)
  mlngMatchReportID = lngNewValue
End Property

Public Property Let MatchReportType(lngMatchReportType As MatchReportType)
  
  mlngMatchReportType = lngMatchReportType

  Select Case mlngMatchReportType
  Case mrtSucession
    Me.Caption = "Succession Planning"
    Me.HelpContextID = 1077
  Case mrtCareer
    Me.Caption = "Career Progression"
    Me.HelpContextID = 1079
  End Select

End Property

Public Property Get ErrorString() As String
  ErrorString = mstrErrorMessage
End Property

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property


Public Function RunMatchReport(Optional plngTableID As Long, Optional plngRecordID As Long) As Boolean
Dim strUtilityName As String

  On Error GoTo RunMatchReport_ERROR

  fOK = True
  Screen.MousePointer = vbHourglass

  If frmBreakDown Is Nothing Then
    Set frmBreakDown = New frmMatchRunBreakDown
  End If
  
  If fOK Then fOK = GetMatchReportDefinition

  strUtilityName = "Match Report"

  Select Case mlngMatchReportType
  Case mrtNormal
    gobjEventLog.AddHeader eltMatchReport, mstrName
  Case mrtSucession
    gobjEventLog.AddHeader eltSuccessionPlanning, mstrName
    strUtilityName = "Succession Planning Report"
  Case mrtCareer
    gobjEventLog.AddHeader eltCareerProgression, mstrName
    strUtilityName = "Career Progression Report"
  End Select

  If fOK Then
    With gobjProgress
      '.AviFile = App.Path & "\videos\report.avi"
      .AVI = dbText
      .MainCaption = strUtilityName
      If Not gblnBatchMode Then
        .NumberOfBars = 1
        .Caption = Me.Caption
        .Time = False
        .Cancel = True
        .Bar1Caption = Me.Caption & " : " & mstrName
        .OpenProgress
      Else
        .ResetBar2
        .Bar2Caption = Me.Caption & " : " & mstrName
      End If
    End With
  End If

  If fOK Then fOK = GetDetailsRecordsets
  If fOK Then fOK = GetRelationRecordsets
  If fOK Then fOK = CheckModuleSetupPermissions
  If fOK Then fOK = GetDataRecordset(plngTableID, plngRecordID)

  If fOK Then fOK = InitialiseFormBreakdown
  If fOK Then fOK = PopulateGridMain
  
  RemoveTemporarySQLObjects
  
  If fOK Then
    If gblnBatchMode Or Not mblnPreviewOnScreen Then
      fOK = OutputReport(False)
    End If
  End If
  
  Call UtilUpdateLastRun(utlMatchReport, mlngMatchReportID)
  mblnUserCancelled = (InStr(LCase(mstrErrorMessage), "cancelled by user") > 0)

  If mblnNoRecords Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful
    gobjEventLog.AddDetailEntry mstrErrorMessage
    mstrErrorMessage = "Completed successfully." & vbCrLf & mstrErrorMessage
    fOK = True
  ElseIf fOK Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful
    mstrErrorMessage = "Completed successfully."
  ElseIf mblnUserCancelled Then
    gobjEventLog.ChangeHeaderStatus elsCancelled
    mstrErrorMessage = "Cancelled by user."
  Else
    'Only details records for failures !
    gobjEventLog.AddDetailEntry mstrErrorMessage
    gobjEventLog.ChangeHeaderStatus elsFailed
    mstrErrorMessage = "Failed." & vbCrLf & vbCrLf & mstrErrorMessage
  End If

  mstrErrorMessage = Me.Caption & " : '" & mstrName & "' " & mstrErrorMessage

  If Not gblnBatchMode Then
    If gobjProgress.Visible Then gobjProgress.CloseProgress
    If (fOK = False) Or (mblnNoRecords = True) Or (Not mblnPreviewOnScreen) Then
      COAMsgBox mstrErrorMessage, IIf(fOK, vbInformation, vbExclamation) + vbOKOnly, Me.Caption
    End If
  End If

  Screen.MousePointer = vbDefault
  RunMatchReport = fOK

  Exit Function

RunMatchReport_ERROR:
  fOK = False
  RunMatchReport = False
  mstrErrorMessage = "Error whilst running this definition." & vbCrLf & Err.Description
  Resume Next

End Function


Private Function GetDetailsRecordsets() As Boolean

  On Error GoTo GetDetailsRecordsets_ERROR

  Dim objColumn As clsColumn
  Dim rsMatchReportsDetails As Recordset
  Dim strTempSQL As String
  Dim intTemp As Integer

  
  strTempSQL = _
      "SELECT ASRSysMatchReportDetails.*, " & _
      "       ASRSysTables.TableID, ASRSysTables.TableName," & _
      "       ASRSysColumns.ColumnName, ASRSysColumns.DataType " & _
      "FROM ASRSysMatchReportDetails " & _
      "LEFT OUTER JOIN   ASRSysColumns on ASRSysMatchReportDetails.colexprid = ASRSysColumns.columnid" & _
      "                  and ASRSysMatchReportDetails.ColType = 'C' " & _
      "LEFT OUTER JOIN   ASRSysTables on ASRSysColumns.TableID = ASRSysTables.TableID " & _
      "WHERE  MatchReportID = " & CStr(mlngMatchReportID) & " " & _
      "ORDER BY [ColSequence]"
  
  Set rsMatchReportsDetails = datGeneral.GetReadOnlyRecords(strTempSQL)

  Set mcolColDetails = New Collection

  Set objColumn = New clsColumn
  objColumn.ColType = "C"
  objColumn.TableID = mlngTable1ID
  objColumn.TableName = mstrTable1Name
  objColumn.ColumnName = "ID"
  objColumn.Hidden = True
  objColumn.Heading = "ID1"
  mcolColDetails.Add objColumn, "ID1"

  If mlngTable2ID > 0 Then
    Set objColumn = New clsColumn
    objColumn.ColType = "C"
    objColumn.TableID = mlngTable2ID
    objColumn.TableName = mstrTable2Name
    objColumn.ColumnName = "ID"
    objColumn.Hidden = True
    objColumn.Heading = "ID2"
    mcolColDetails.Add objColumn, "ID2"
  End If
  
  
  With rsMatchReportsDetails
    If .BOF And .EOF Then
      GetDetailsRecordsets = False
      mstrErrorMessage = "No columns found in the specified definition." & vbCrLf & "Please remove this definition and create a new one."
      Exit Function
    End If

    intTemp = 0
    Do Until .EOF
      intTemp = intTemp + 1
      
      ReDim Preserve mstrExcelFormats(intTemp)

      Set objColumn = New clsColumn
      
      objColumn.ColType = !ColType
      objColumn.ID = !ColExprID
      objColumn.Size = !ColSize
      objColumn.DecPlaces = !ColDecs
      objColumn.Heading = !ColHeading
      objColumn.Sequence = !ColSequence
      objColumn.SortSeq = !SortOrderSeq
      objColumn.SortDir = !SortOrderDirection
      objColumn.ThousandSeparator = datGeneral.DoesColumnUseSeparators(!ColExprID)
      
      If objColumn.ColType = "C" Then
        objColumn.TableID = !TableID
        objColumn.TableName = !TableName
        objColumn.ColumnName = !ColumnName
        objColumn.DataType = !DataType
        
        Select Case CLng(!DataType)
        Case sqlNumeric, sqlInteger
          objColumn.IsNumeric = True

          If objColumn.DecPlaces > 0 Then
            If objColumn.DecPlaces > 127 Then
              mstrExcelFormats(intTemp) = "0." & String(127, "0")
            Else
              mstrExcelFormats(intTemp) = "0." & String(objColumn.Size, "0")
            End If
          Else
            If objColumn.Size > 0 Then
              mstrExcelFormats(intTemp) = "0"
            Else
              mstrExcelFormats(intTemp) = "General"
            End If
          End If

        Case sqlDate
          mstrExcelFormats(intTemp) = DateFormat
        Case Else
          mstrExcelFormats(intTemp) = "@"
        End Select

      Else
        'MH20010307
        Dim objExpr As DataMgr.clsExprExpression
        Set objExpr = New clsExprExpression

        objExpr.ExpressionID = CLng(!ColExprID)
        objExpr.ConstructExpression
        
        'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
        ' when all we need is the expression return type (time saving measure).
        objExpr.ValidateExpression True, True  'MH20010508

        objColumn.TableID = objExpr.BaseTableID
        objColumn.TableName = objExpr.BaseTableName
        objColumn.IsNumeric = True    'Match score has to be numeric
        objColumn.DataType = sqlNumeric
        
        
        Select Case objExpr.ReturnType
        Case giEXPRVALUE_NUMERIC, giEXPRVALUE_BYREF_NUMERIC
          If objColumn.DecPlaces > 0 Then
            If objColumn.DecPlaces > 127 Then
              mstrExcelFormats(intTemp) = "0." & String(127, "0")
            Else
              mstrExcelFormats(intTemp) = "0." & String(objColumn.DecPlaces, "0")
            End If
          Else
            If objColumn.Size > 0 Then
              mstrExcelFormats(intTemp) = "0"
            Else
              mstrExcelFormats(intTemp) = "General"
            End If
          End If

        Case giEXPRVALUE_DATE, giEXPRVALUE_BYREF_DATE
          mstrExcelFormats(intTemp) = DateFormat
        Case Else
          mstrExcelFormats(intTemp) = "@"
        End Select

        Set objExpr = Nothing

      End If

      mcolColDetails.Add objColumn, objColumn.ColType & CStr(objColumn.ID)
     
     .MoveNext
    Loop
  .MoveFirst
  End With

  GetDetailsRecordsets = True
  Exit Function

GetDetailsRecordsets_ERROR:

  GetDetailsRecordsets = False
  mstrErrorMessage = "Error whilst retrieving the details recordsets'." & vbCrLf & Err.Description

End Function

Private Function GetRelationRecordsets() As Boolean

  On Error GoTo GetRelationRecordsets_ERROR

  Dim objRelation As clsMatchRelation
  Dim objColumn As clsColumn
  Dim objBreakdownCols As Collection
  Dim rsMatchReportsDetails As Recordset
  Dim rsMatchBreakdownColumns As Recordset
  Dim strTempSQL As String
  Dim pintNextIndex As Integer
  Dim intTemp As Integer
  
  
  strTempSQL = _
      "SELECT ASRSysMatchReportTables.*," & _
      "       a.TableName as Table1Name, b.TableName as Table2Name " & _
      "FROM   ASRSysMatchReportTables " & _
      "JOIN   ASRSysTables a on ASRSysMatchReportTables.Table1ID = a.TableID " & _
      "LEFT OUTER JOIN   ASRSysTables b on ASRSysMatchReportTables.Table2ID = b.TableID " & _
      "WHERE  MatchReportID = " & CStr(mlngMatchReportID) & _
      "ORDER BY ASRSysMatchReportTables.MatchRelationID"

  Set rsMatchReportsDetails = datGeneral.GetReadOnlyRecords(strTempSQL)
  If rsMatchReportsDetails.BOF And rsMatchReportsDetails.EOF Then
    mstrErrorMessage = "Cannot load the table relation information for this definition."
    GetRelationRecordsets = False
    rsMatchReportsDetails.Close
    Set rsMatchReportsDetails = Nothing
    Exit Function
  End If
  
  
  Set mcolRelations = New Collection
  ReDim mlngTableViews(2, 0)
  
  With rsMatchReportsDetails
    Do While Not .EOF
  
      Set objRelation = New clsMatchRelation
  
      objRelation.Table1ID = !Table1ID
      objRelation.Table1Name = !Table1Name


      If Not TablePermission(!Table1ID) Then
        mstrErrorMessage = "You do not have permission to read the '" & !Table1Name & "' table either directly or through any views."
        GetRelationRecordsets = False
        Exit Function
      End If

      If !Table2ID > 0 Then
        If Not TablePermission(!Table2ID) Then
          mstrErrorMessage = "You do not have permission to read the '" & !Table2Name & "' table either directly or through any views."
          GetRelationRecordsets = False
          Exit Function
        End If
      End If


      objRelation.Table1RealSource = gcoTablePrivileges.Item(objRelation.Table1Name).RealSource
      AddToJoinArray 0, !Table1ID
      
      objRelation.Table2ID = !Table2ID
      If objRelation.Table2ID > 0 Then
        objRelation.Table2Name = !Table2Name
      
        'If gcoTablePrivileges.Item(objRelation.Table2Name).AllowSelect = False Then
        '  mstrErrorMessage = "You do not have read permission for the '" & objRelation.Table2Name & "' table."
        '  GetRelationRecordsets = False
        '  Exit Function
        'End If
      
        objRelation.Table2RealSource = gcoTablePrivileges.Item(objRelation.Table2Name).RealSource
        AddToJoinArray 0, !Table2ID
      End If
      
      objRelation.RequiredExprID = !RequiredExprID
      objRelation.PreferredExprID = !PreferredExprID
      objRelation.MatchScoreID = !MatchScoreExprID
      
      
      
      strTempSQL = "SELECT   ASRSysMatchReportBreakdown.*, " & _
                   "         ASRSysTables.TableID, ASRSysTables.TableName, " & _
                   "         ASRSysColumns.ColumnName, ASRSysColumns.DataType " & _
                   "FROM     ASRSysMatchReportBreakdown " & _
                   "JOIN     ASRSysMatchReportTables " & _
                   "ON       ASRSysMatchReportBreakdown.MatchRelationID = ASRSysMatchReportTables.MatchRelationID " & _
                   "LEFT OUTER JOIN ASRSysColumns " & _
                   "ON       ASRSysMatchReportBreakdown.ColExprID = ASRSysColumns.ColumnID AND ASRSysMatchReportBreakdown.ColType = 'C' " & _
                   "LEFT OUTER JOIN ASRSysTables " & _
                   "ON       ASRSysColumns.TableID = ASRSysTables.TableID " & _
                   "WHERE    ASRSysMatchReportBreakdown.MatchReportID = " & CStr(mlngMatchReportID) & " " & _
                   "AND      ASRSysMatchReportTables.Table1ID = " & CStr(objRelation.Table1ID) & " " & _
                   "ORDER BY ColSequence"

      Set rsMatchBreakdownColumns = datGeneral.GetReadOnlyRecords(strTempSQL)

      Set objBreakdownCols = New Collection
      
      With rsMatchBreakdownColumns
        Do While Not .EOF
          Set objColumn = New clsColumn
          
          objColumn.ColType = !ColType
          objColumn.ID = !ColExprID
          objColumn.Size = !ColSize
          objColumn.DecPlaces = !ColDecs
          objColumn.Heading = !ColHeading
          'objcolumn.Sequence = !ColSequence
          'objColumn.SortSeq = !SortOrderSeq
          'objColumn.SortDir = !SortOrderDirection
          objColumn.ThousandSeparator = datGeneral.DoesColumnUseSeparators(!ColExprID)

          If objColumn.ColType = "C" Then
            objColumn.DataType = !DataType
            objColumn.TableID = !TableID
            objColumn.TableName = !TableName
            objColumn.ColumnName = !ColumnName
            objColumn.IsNumeric = (objColumn.DataType = sqlInteger Or sqlNumeric)
          Else
            objColumn.DataType = sqlNumeric
            objColumn.IsNumeric = True
          End If

          objBreakdownCols.Add objColumn, objColumn.ColType & CStr(objColumn.ID)

          .MoveNext
        Loop
      End With

      objRelation.BreakdownColumns = objBreakdownCols
      mcolRelations.Add objRelation, "T" & CStr(objRelation.Table1ID)

      Set objBreakdownCols = Nothing
      Set objRelation = Nothing

      .MoveNext
    Loop
  End With

  GetRelationRecordsets = True

TidyUpAndExit:
  Exit Function

GetRelationRecordsets_ERROR:
  GetRelationRecordsets = False
  mstrErrorMessage = "Error whilst retrieving the relation recordsets" & vbCrLf & Err.Description
  Resume TidyUpAndExit

End Function


Private Function GetDataRecordset(plngTableID As Long, plngRecordID As Long) As Boolean

  Dim rsTemp As Recordset
  Dim strReportingStructure As String
  Dim strMainTable As String
  
  On Local Error GoTo GetDataRecordset_ERROR

  fOK = True
  
  mstrTable1RealSource = gcoTablePrivileges.Item(mstrTable1Name).RealSource
  If mlngTable2ID > 0 Then
    mstrTable2RealSource = gcoTablePrivileges.Item(mstrTable2Name).RealSource
  End If
  ReDim alngSourceTables(2, 0)

  If fOK Then fOK = GenerateSQLWhere(plngTableID, plngRecordID)
  If fOK Then fOK = GenerateSQLMatchScore
  If fOK Then fOK = GenerateSQLSelect
  If fOK Then fOK = GenerateSQLJoin
  If fOK Then fOK = GenerateSQLOrderBy
  If fOK Then fOK = UDFFunctions(mastrUDFsRequired, True)

  If fOK Then
    mstrTempTableName = GetTempTable
  End If

  If fOK = False Then
    Exit Function
  End If

  If mlngMatchReportType = mrtNormal Then
'MH20050104 Fault 9550
'    mstrSQL = "SELECT ID FROM " & mstrTable1Name &
    mstrSQL = "SELECT ID FROM " & mstrTable1RealSource & _
              IIf(mstrTable1Where <> vbNullString, " WHERE " & mstrTable1Where, vbNullString)
  Else
'MH20050104 Fault 9550
'    mstrSQL = "SELECT ID FROM " & mstrTable2Name &
    mstrSQL = "SELECT ID FROM " & mstrTable2RealSource & _
              IIf(mstrTable2Where <> vbNullString, " WHERE " & mstrTable2Where, vbNullString)
  End If
  
  Set rsTemp = datGeneral.GetReadOnlyRecords(mstrSQL)

  If gblnBatchMode Then
    gobjProgress.Bar2MaxValue = rsTemp.RecordCount
  Else
    gobjProgress.Bar1MaxValue = rsTemp.RecordCount
  End If

  Do While Not rsTemp.EOF
      
      If fOK Then
        
        'Reporting Structure
        If mlngMatchReportType <> mrtNormal Then
          strReportingStructure = GetReportingStructure(rsTemp.Fields(0).Value)
        End If
        
        
        mstrSQL = _
          "INSERT INTO [" & gsUserName & "].[" & mstrTempTableName & "]" & _
          " SELECT " & IIf(mlngNumRecords > 0, "TOP " & CStr(mlngNumRecords) & " ", vbNullString) & _
          mcolSQLSelect("T0") & vbCrLf & _
          " FROM " & mstrTable1RealSource & vbCrLf & _
          mcolSQLJoin("T0")
        
        Select Case mlngMatchReportType
        Case mrtNormal
          mstrSQL = mstrSQL & _
                " WHERE " & mstrTable1RealSource & ".ID = " & CStr(rsTemp.Fields(0).Value) & vbCrLf
        Case mrtSucession
          mstrSQL = mstrSQL & _
                " WHERE " & mstrTable1RealSource & ".ID = " & CStr(GetJobTableID(rsTemp.Fields(0).Value)) & vbCrLf
        Case mrtCareer
          mstrSQL = mstrSQL & _
                " WHERE " & mstrTable2RealSource & ".ID = " & CStr(rsTemp.Fields(0).Value) & vbCrLf
        End Select
        
        If mstrSQLWhere <> vbNullString Then
          mstrSQL = mstrSQL & " AND " & mstrSQLWhere & vbCrLf
        End If

        If strReportingStructure <> vbNullString Then
          mstrSQL = mstrSQL & " AND " & strReportingStructure & vbCrLf
        End If
        
        
        If mlngMatchReportType = mrtNormal Then
          mstrSQL = mstrSQL & _
                  mstrSQLGroupBy & vbCrLf & _
                  IIf(mstrTable2Where <> vbNullString, " HAVING " & mstrTable2Where & vbCrLf, "")

          If mblnScoreCheck Then
            mstrSQL = mstrSQL & _
                  IIf(mstrTable2Where <> vbNullString, " AND ", " HAVING ") & _
                  mcolSQLMatchScore("T0") & _
                  IIf(mlngScoreMode = 0, " >= ", " <= ") & _
                  CStr(mlngScoreLimit) & vbCrLf
          End If

        Else
          mstrSQL = mstrSQL & _
                  mstrSQLGroupBy & vbCrLf & _
                  IIf(mstrTable1Where <> vbNullString, " HAVING " & mstrTable1Where & vbCrLf, "")

          If mblnScoreCheck Then
            mstrSQL = mstrSQL & _
                  IIf(mstrTable1Where <> vbNullString, " AND ", " HAVING ") & _
                  mcolSQLMatchScore("T0") & _
                  IIf(mlngScoreMode = 0, " >= ", " <= ") & _
                  CStr(mlngScoreLimit) & vbCrLf
          End If

        End If

        'MH20030606
        'Still need order in case we are doing TOP X records
        mstrSQL = mstrSQL & _
            " ORDER BY " & mcolSQLMatchScore("T0") & _
            IIf(mlngScoreMode = 0, " DESC", vbNullString)

        datGeneral.ExecuteSql mstrSQL, mstrErrorMessage

      End If

      If gobjProgress.Cancelled Then
        mstrErrorMessage = "Cancelled by user."
        mblnUserCancelled = True
      End If
      If mstrErrorMessage <> vbNullString Then
        
        'MH20060103 Bodge fix to ignore warning about nulls...
        If mstrErrorMessage <> "Warning: Null value is eliminated by an aggregate or other SET operation." Then
          GetDataRecordset = False
          Exit Function
        End If
      
      End If

      gobjProgress.UpdateProgress gblnBatchMode
      rsTemp.MoveNext

  Loop

  fOK = UDFFunctions(mastrUDFsRequired, False)

  rsTemp.Close
  Set rsTemp = Nothing

  GetDataRecordset = fOK


TidyUpAndExit:
  Exit Function

GetDataRecordset_ERROR:
  GetDataRecordset = False
  mstrErrorMessage = "Error retrieving data" & vbCrLf & Err.Description
  Resume TidyUpAndExit

End Function


Public Function GetRecordsetBreakdown(lngTableID As Long, lngRecord1ID As Long, lngRecord2ID As Long)

  Dim objRelation As clsMatchRelation

  Set objRelation = mcolRelations("T" & CStr(lngTableID))


'MH20030909
  mstrSQL = "SELECT " & mcolSQLSelect("T" & CStr(lngTableID)) & vbCrLf & _
            "FROM " & objRelation.Table1RealSource & vbCrLf

  If lngTableID = mlngTable1ID Then

    mstrSQL = mstrSQL & _
        mcolSQLJoin("T" & CStr(lngTableID)) & _
        " WHERE " & objRelation.Table1RealSource & ".ID = " & CStr(lngRecord1ID)
    If objRelation.Table2ID > 0 Then
      mstrSQL = mstrSQL & _
          " AND " & objRelation.Table2RealSource & ".ID = " & CStr(lngRecord2ID)
    End If


  Else
    If objRelation.Table2ID > 0 Then
      mstrSQL = mstrSQL & _
        mcolSQLJoin("T" & CStr(lngTableID)) & _
        " AND " & objRelation.Table1RealSource & ".ID_" & CStr(mlngTable1ID) & " = " & CStr(lngRecord1ID) & _
        " AND " & objRelation.Table2RealSource & ".ID_" & CStr(mlngTable2ID) & " = " & CStr(lngRecord2ID)
    End If

    mstrSQL = mstrSQL & _
      " WHERE " & objRelation.Table1RealSource & ".ID_" & CStr(mlngTable1ID) & " = " & CStr(lngRecord1ID)
    If frmBreakDown.chkAllRecords.Value = vbChecked And mlngTable2ID > 0 Then
      mstrSQL = mstrSQL & _
        " OR " & objRelation.Table2RealSource & ".ID_" & CStr(mlngTable2ID) & " = " & CStr(lngRecord2ID)
    End If

  End If
'  mstrSQL = "SELECT " & mcolSQLSelect("T" & CStr(lngTableID)) & vbCrLf & _
'            "FROM " & mstrTable1RealSource & vbCrLf & _
'            "CROSS JOIN " & mstrTable2RealSource & vbCrLf
'
'
'  If lngTableID <> mlngTable1ID Then
'    mstrSQL = mstrSQL & _
'      " LEFT OUTER JOIN " & objRelation.Table1RealSource & " ON " & objRelation.Table1RealSource & ".ID_" & CStr(mlngTable1ID) & " = " & CStr(lngRecord1ID) & _
'      " LEFT OUTER JOIN " & objRelation.Table2RealSource & " ON " & objRelation.Table2RealSource & ".ID_" & CStr(mlngTable2ID) & " = " & CStr(lngRecord2ID)
'
'    If frmBreakDown.chkAllRecords.Value <> vbChecked And mlngTable2ID > 0 Then
'      mstrSQL = mstrSQL & " AND " & mcolSQLJoin("T" & CStr(lngTableID))
'    End If
'  End If
'
'  mstrSQL = mstrSQL & _
'      " WHERE " & mstrTable1RealSource & ".ID = " & CStr(lngRecord1ID)
'  If objRelation.Table2ID > 0 Then
'    mstrSQL = mstrSQL & _
'        " AND " & mstrTable2RealSource & ".ID = " & CStr(lngRecord2ID)
'  End If

  mstrSQL = mstrSQL & _
          " ORDER BY " & mcolSQLOrderBy("T" & CStr(lngTableID))

  GetRecordsetBreakdown = fOK

End Function


Private Function GenerateSQLMatchScore() As Boolean

  Dim objColumn As clsColumn
  Dim mobjColumnPrivileges As CColumnPrivileges
  Dim mobjTableView As CTablePrivilege
  Dim objRelation As clsMatchRelation
  
  Dim blnOK As Boolean
  Dim pblnColumnOK As Boolean
  Dim iLoop1 As Integer
  Dim pblnNoSelect As Boolean
  Dim pblnFound As Boolean
  Dim lngScoreExpr As Long
  Dim strOutput As String
  
  Dim pintLoop As Integer
  Dim pstrColumnCode As String
  Dim pstrColumnCount As String
  Dim pstrSource As String
  Dim pintNextIndex As Integer
  Dim strRealSource1 As String
  Dim strRealSource2 As String
  
  Dim sFilterCode As String
  Dim sCalcCode As String
  Dim objCalcExpr As clsExprExpression
  Dim lngCount As Long
  
  
  pstrColumnCode = vbNullString
  pstrColumnCount = vbNullString
  lngCount = 0
  Set mcolSQLMatchScore = New Collection
  
  For Each objRelation In mcolRelations
    
    sCalcCode = vbNullString
    
    If objRelation.MatchScoreID > 0 Then
    
      Set objCalcExpr = New clsExprExpression

      blnOK = objCalcExpr.Initialise(objRelation.Table1ID, objRelation.MatchScoreID, giEXPR_MATCHSCOREEXPRESSION, giEXPRVALUE_LOGIC, objRelation.Table2ID)

      If blnOK Then
        blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sCalcCode, True)
        If blnOK And gbEnableUDFFunctions Then
          blnOK = objCalcExpr.UDFCalculationCode(alngSourceTables, mastrUDFsRequired(), True)
        End If
      End If

      If blnOK Then
        ' Add the required views to the JOIN code.
        For iLoop1 = 1 To UBound(alngSourceTables, 2)
          AddToJoinArray alngSourceTables(1, iLoop1), alngSourceTables(2, iLoop1)
        Next iLoop1
      Else
        ' Permission denied on something in the calculation.
        mstrErrorMessage = "You do not have permission to use the match score expression."
        GenerateSQLMatchScore = False
        Exit Function
      End If
      Set objCalcExpr = Nothing


'      sCalcCode = GetCalcCode(objRelation.Table1ID, objRelation.Table2ID, objRelation.MatchScoreID, giEXPR_MATCHSCOREEXPRESSION, giEXPRVALUE_LOGIC)
'      If sCalcCode = vbNullString Then
'        ' Permission denied on something in the calculation.
'        mstrErrorMessage = "You do not have permission to use a match score calculation."
'        GenerateSQLMatchScore = False
'        Exit Function
'      End If


If ASRDEVELOPMENT Then
  sCalcCode = Replace(sCalcCode, Chr(10), " ")
  sCalcCode = Replace(sCalcCode, Chr(11), " ")
  sCalcCode = Replace(sCalcCode, Chr(12), " ")
  sCalcCode = Replace(sCalcCode, Chr(13), " ")
  
  Do While InStr(sCalcCode, "  ") > 0
    sCalcCode = Replace(sCalcCode, "  ", " ")
  Loop
End If

      If sCalcCode <> vbNullString Then
        sCalcCode = "isnull(" & sCalcCode & ",0)"
      End If

    Else
      'If objRelation.Table2ID <> mlngTable2ID Then
        'sCalcCode = "case when " & objRelation.Table2RealSource & ".ID > 0 then 100 else 0 end"
        sCalcCode = "100"
      'End If

    End If
    
    If sCalcCode <> vbNullString Then

      Set objCalcExpr = New clsExprExpression

      If objRelation.PreferredExprID > 0 Then

        'If objRelation.Table2ID > 0 Then
        '  blnOK = objCalcExpr.Initialise(objRelation.Table2ID, objRelation.PreferredExprID, giEXPR_MATCHJOINEXPRESSION, giEXPRVALUE_LOGIC, objRelation.Table1ID)
        'Else
          blnOK = objCalcExpr.Initialise(objRelation.Table1ID, objRelation.PreferredExprID, giEXPR_MATCHJOINEXPRESSION, giEXPRVALUE_LOGIC, objRelation.Table2ID)
        'End If

        If blnOK Then
          blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sFilterCode, True)
          'blnOK = objCalcExpr.RuntimeFilterCode(sFilterCode, True, False)
          
          ' JDM - 06/08/2003 - Fault 6224 - Not generating UDFs for fields
          If blnOK And gbEnableUDFFunctions Then
            blnOK = objCalcExpr.UDFCalculationCode(alngSourceTables, mastrUDFsRequired(), True)
          End If
        End If
  
        If blnOK Then
          ' Add the required views to the JOIN code.
          For iLoop1 = 1 To UBound(alngSourceTables, 2)
            AddToJoinArray alngSourceTables(1, iLoop1), alngSourceTables(2, iLoop1)
          Next iLoop1
        Else
          ' Permission denied on something in the calculation.
          mstrErrorMessage = "You do not have permission to use the required match expression."
          GenerateSQLMatchScore = False
          Exit Function
        End If
    
        If sFilterCode <> vbNullString Then
          'sCalcCode = " CASE WHEN " & sFilterCode & " THEN " & sCalcCode & " ELSE 0 END "
          sCalcCode = " CASE WHEN " & sFilterCode & " = 1 THEN " & sCalcCode & " ELSE 0 END "
          'Stop
        End If
        Set objCalcExpr = Nothing
      
      End If
      
      'End If


      strRealSource1 = objRelation.Table1RealSource
      
      If objRelation.Table2ID > 0 Then
        strRealSource2 = objRelation.Table2RealSource

        sCalcCode = "case when " & _
                strRealSource1 & ".ID > 0 and " & _
                strRealSource2 & ".ID > 0 then " & _
                sCalcCode & _
                " else 0 end"
      End If

      If objRelation.Table1ID = mlngTable1ID Then
        pstrColumnCount = pstrColumnCount & _
          IIf(pstrColumnCount <> vbNullString, "+", "") & _
          "1"

        pstrColumnCode = pstrColumnCode & _
          IIf(pstrColumnCode <> vbNullString, "+", "") & _
          "max(cast(" & sCalcCode & " as float))"
      
      ElseIf objRelation.Table2ID = 0 And mcolRelations.Count = 1 Then
        pstrColumnCount = pstrColumnCount & _
          IIf(pstrColumnCount <> vbNullString, "+", "") & _
          "1"

        pstrColumnCode = pstrColumnCode & _
          IIf(pstrColumnCode <> vbNullString, "+", "") & _
          "cast(sum(" & sCalcCode & ") as float)"

      Else
        pstrColumnCount = pstrColumnCount & _
          IIf(pstrColumnCount <> vbNullString, "+", "") & _
          "count(distinct " & strRealSource1 & ".ID)"
      
        'pstrColumnCode = pstrColumnCode & _
          IIf(pstrColumnCode <> vbNullString, "+", "") & _
          "cast((sum(" & sCalcCode & ") * count(distinct " & strRealSource1 & ".ID) / " & _
          "case when sum(case when " & strRealSource1 & ".ID > 0 then 1 else 0 end) = 0 then 1 else sum(case when " & strRealSource1 & ".ID > 0 then 1 else 0 end) end) as float)"
        pstrColumnCode = pstrColumnCode & _
          IIf(pstrColumnCode <> vbNullString, "+", "") & _
          "cast((sum(" & sCalcCode & ") * count(distinct " & strRealSource1 & ".ID) / " & _
          "case when sum(case when " & strRealSource1 & ".ID > 0 then 1 else 0 end) = 0 then 1 else cast(sum(case when " & strRealSource1 & ".ID > 0 then 1 else 0 end) as float) end) as float)"
        
      End If

      mcolSQLMatchScore.Add sCalcCode, "T" & CStr(objRelation.Table1ID)

    End If

  Next

  If pstrColumnCount <> "1" And mlngTable2ID > 0 Then
    strOutput = "((" & pstrColumnCode & ") / " & "case when " & pstrColumnCount & " = 0 then 1 else " & pstrColumnCount & " end)"
  Else
    strOutput = pstrColumnCode
  End If
  mcolSQLMatchScore.Add strOutput, "T0"

  GenerateSQLMatchScore = True

End Function


Private Function GenerateSQLSelect() As Boolean

  Dim objRelation As clsMatchRelation
  Dim strOutput As String
  
  Set mcolSQLSelect = Nothing
  Set mcolSQLSelect = New Collection

  Set mcolSQLOrderBy = Nothing
  Set mcolSQLOrderBy = New Collection

  mstrErrorMessage = vbNullString
  GenerateSQLSelect = False
  
  mstrSQLGroupBy = vbNullString

  GetSelectStatement mcolColDetails, 0, ""
  If mstrErrorMessage <> vbNullString Then
    Exit Function
  End If
  
  
  For Each objRelation In mcolRelations
    GetSelectStatement objRelation.BreakdownColumns, objRelation.Table1ID, objRelation.Table1RealSource
    If mstrErrorMessage <> vbNullString Then
      Exit Function
    End If

    'mcolSQLSelect.Add strOutput, "T" & CStr(objRelation.Table1ID)
  Next
  
  GenerateSQLSelect = True

End Function


Private Function GetSelectStatement(colColumns As Collection, lngTableID As Long, strTable1RealSource As String) As String
  
  Dim objColumn As clsColumn
  Dim mobjColumnPrivileges As CColumnPrivileges
  Dim mobjTableView As CTablePrivilege
  Dim objCalcExpr As clsExprExpression
  
  Dim blnOK As Boolean
  Dim pblnColumnOK As Boolean
  Dim pblnNoSelect As Boolean
  Dim pstrColumnCode As String
  Dim pstrSource As String
  Dim pintNextIndex As Integer
  Dim strRealSource As String
  Dim sCalcCode As String
  Dim blnBooleanColumn As Boolean
  
  Dim strSQLSelect As String
  Dim strSQLOrderBy As String
  Dim strOrderColumn As String
  
  strSQLSelect = vbNullString
  strSQLOrderBy = vbNullString
  If strTable1RealSource <> vbNullString Then
    strSQLOrderBy = "case when " & strTable1RealSource & ".ID is null then 1 else 0 end"
  End If


  ' Set flags with their starting values
  blnOK = True
  pblnNoSelect = False
  On Local Error GoTo LocalErr


  strSQLSelect = vbNullString
  pstrColumnCode = vbNullString

  For Each objColumn In colColumns

    ' If its a COLUMN then...
    If objColumn.ColType = "C" Then
      
      ' Check permission on that column
      Set mobjColumnPrivileges = GetColumnPrivileges(objColumn.TableName)
      blnBooleanColumn = mobjColumnPrivileges.Item(objColumn.ColumnName).DataType = sqlBoolean

      pblnColumnOK = gcoTablePrivileges.Item(objColumn.TableName).AllowSelect
      
      'MH20040422 Fault 8267
      'If pblnColumnOK Then
      If pblnColumnOK Or objColumn.ColumnName = "ID" Then
        strRealSource = gcoTablePrivileges.Item(objColumn.TableName).RealSource
        pblnColumnOK = mobjColumnPrivileges.IsValid(objColumn.ColumnName)
        If pblnColumnOK Then
          pblnColumnOK = mobjColumnPrivileges.Item(objColumn.ColumnName).AllowSelect
        End If
      End If
      
      If pblnColumnOK Then
        pstrColumnCode = strRealSource & "." & Trim(objColumn.ColumnName)

        AddToJoinArray 0, objColumn.TableID
      Else

        ' this column cannot be read direct. If its from a parent, try parent views
        ' Loop thru the views on the table, seeing if any have read permis for the column

        pblnNoSelect = True
        ReDim mstrViews(0)
        pstrColumnCode = vbNullString
        
        For Each mobjTableView In gcoTablePrivileges.Collection
          If (Not mobjTableView.IsTable) And _
          (mobjTableView.TableID = objColumn.TableID) And _
          (mobjTableView.AllowSelect) Then
            
            pstrSource = mobjTableView.ViewName
            strRealSource = gcoTablePrivileges.Item(pstrSource).RealSource
            
            ' Get the column permission for the view
            Set mobjColumnPrivileges = GetColumnPrivileges(pstrSource)
            
            ' If we can see the column from this view
            If mobjColumnPrivileges.IsValid(objColumn.ColumnName) Then
              If mobjColumnPrivileges.Item(objColumn.ColumnName).AllowSelect Then
                pstrColumnCode = pstrColumnCode & _
                    " WHEN NOT " & pstrSource & "." & objColumn.ColumnName & " IS NULL THEN " & pstrSource & "." & objColumn.ColumnName
                AddToJoinArray 1, mobjTableView.ViewID
              End If
            End If
          End If

        Next mobjTableView

        Set mobjTableView = Nothing

        ' Does the user have select permission thru ANY views ?
        ' If we cant see a column, then get outta here
        If pstrColumnCode = vbNullString Then
          strSQLSelect = vbNullString
          mstrErrorMessage = "You do not have permission to see the column '" & objColumn.ColumnName & "' either directly or through any views."
          Exit Function
        
        Else
          pstrColumnCode = "CASE" & pstrColumnCode & " ELSE NULL END"

        End If
        
        If Not blnOK Then
          strSQLSelect = vbNullString
          Exit Function
        End If
      
      End If

      
      'MH20040422 Fault 8285
      'If mobjColumnPrivileges.Item(objColumn.ColumnName).DataType = sqlBoolean Then
      If blnBooleanColumn Then
        pstrColumnCode = "(case when " & pstrColumnCode & " = 1 then 'Y' else 'N' end)"
      End If

      If lngTableID = 0 Then
        mstrSQLGroupBy = mstrSQLGroupBy & _
          IIf(mstrSQLGroupBy <> vbNullString, ", ", "") & pstrColumnCode
      End If

      strOrderColumn = pstrColumnCode

      'pstrColumnCode = pstrColumnCode & " AS '" & objColumn.TableName & objColumn.ColumnName & "'"
      pstrColumnCode = pstrColumnCode & " AS '" & Replace(objColumn.Heading, "'", "''") & "'"
    Else
      pstrColumnCode = vbCrLf & mcolSQLMatchScore("T" & CStr(lngTableID))
      strOrderColumn = mcolSQLMatchScore("T" & CStr(lngTableID))
      pstrColumnCode = pstrColumnCode & _
          " AS 'Match_Score'"
    
    End If

    If mlngMatchReportType <> mrtNormal Then
      objColumn.SQL = pstrColumnCode
    End If

    strSQLSelect = strSQLSelect & _
        IIf(strSQLSelect <> vbNullString, ", ", "") & pstrColumnCode
        
    strSQLOrderBy = strSQLOrderBy & _
        IIf(strSQLOrderBy <> vbNullString, ", ", "") & strOrderColumn
    
  Next
  
  If lngTableID = 0 And mstrSQLGroupBy <> vbNullString Then
    mstrSQLGroupBy = "GROUP BY " & mstrSQLGroupBy & vbCrLf
  End If

  mcolSQLSelect.Add strSQLSelect, "T" & CStr(lngTableID)
  mcolSQLOrderBy.Add strSQLOrderBy, "T" & CStr(lngTableID)

Exit Function

LocalErr:
  fOK = False
  mstrErrorMessage = "Error whilst generating SQL Select statement." & vbCrLf & Err.Description

End Function

Private Function GenerateSQLJoin() As Boolean

  On Error GoTo GenerateSQLJoin_ERROR

  Dim pobjTableView As CTablePrivilege
  Dim objChildTable As CTablePrivilege
  Dim objRelation As clsMatchRelation
  Dim objCalcExpr As clsExprExpression

  Dim strOutputMain As String
  Dim strOutputBaseBreakdown As String
  Dim strOutputChildBreakdown As String
  Dim strOutputGrade As String

  Dim pintLoop As Integer
  Dim pintLoop1 As Integer
  Dim sCalcCode As String
  Dim sTemp As String
  Dim strRealSource As String
  Dim blnFound As Boolean
  
  Dim strESelect As String
  Dim strEJoin As String
  Dim strPSelect As String
  Dim strPJoin As String
  Dim strViewIDs As String
  Dim strArray() As String
  Dim lngIndex As Long

  Dim blnChildOf1 As Boolean
  Dim blnChildOf2 As Boolean


  Set mcolSQLJoin = New Collection
  strOutputMain = vbNullString

  If mlngTable2ID > 0 Then
    strOutputBaseBreakdown = "CROSS JOIN " & mstrTable2RealSource
  End If


  If mlngMatchReportType <> mrtNormal Then
    strRealSource = gcoTablePrivileges.Item(gstrGradeTableName).RealSource

    GetSelectAndJoinForColumn glngPersonnelTableID, gsPersonnelTableName, gsPersonnelGradeColumnName, strESelect, strEJoin, strViewIDs
    If strESelect = vbNullString Then
      mstrErrorMessage = "You do not have permission to see the column '" & gsPersonnelTableName & "." & gsPersonnelGradeColumnName & "' either directly or through any views."
      GenerateSQLJoin = False
      Exit Function
    End If
    
    strArray = Split(strViewIDs, " ")
    For lngIndex = 1 To UBound(strArray)
      AddToJoinArray 1, CLng(strArray(lngIndex))
    Next


    GetSelectAndJoinForColumn glngPostTableID, gstrPostTableName, gstrPostGradeColumnName, strPSelect, strPJoin, strViewIDs
    If strPSelect = vbNullString Then
      mstrErrorMessage = "You do not have permission to see the column '" & gstrPostTableName & "." & gstrPostGradeColumnName & "' either directly or through any views."
      GenerateSQLJoin = False
      Exit Function
    End If

    strArray = Split(strViewIDs, " ")
    For lngIndex = 1 To UBound(strArray)
      AddToJoinArray 1, CLng(strArray(lngIndex))
    Next

    strOutputGrade = _
          " LEFT OUTER JOIN " & strRealSource & " ASRSys" & gsPersonnelTableName & gstrGradeTableName & _
          " ON (" & strESelect & ") = " & "ASRSys" & gsPersonnelTableName & gstrGradeTableName & "." & gstrGradeColumnName & vbCrLf & _
          " LEFT OUTER JOIN " & strRealSource & " ASRSys" & gstrPostTableName & gstrGradeTableName & _
          " ON (" & strPSelect & ") = " & "ASRSys" & gstrPostTableName & gstrGradeTableName & "." & gstrGradeColumnName & vbCrLf

  End If


  For pintLoop = 1 To UBound(mlngTableViews, 2)

    ' Get the table/view object from the id stored in the array
    If mlngTableViews(1, pintLoop) = 0 Then
      Set pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
    Else
      Set pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
    End If


    If pobjTableView.TableID = mlngTable1ID Then
      
      strOutputBaseBreakdown = strOutputBaseBreakdown & _
            " LEFT OUTER JOIN " & pobjTableView.RealSource & _
            " ON " & mstrTable1RealSource & ".ID = " & pobjTableView.RealSource & ".ID" & vbCrLf

    ElseIf pobjTableView.TableID = mlngTable2ID Then
            
      strOutputBaseBreakdown = strOutputBaseBreakdown & _
            " LEFT OUTER JOIN " & pobjTableView.RealSource & _
            " ON " & mstrTable2RealSource & ".ID = " & pobjTableView.RealSource & ".ID" & vbCrLf

    Else
    
      blnChildOf1 = datGeneral.IsAChildOf(pobjTableView.TableID, mlngTable1ID)
      blnChildOf2 = datGeneral.IsAChildOf(pobjTableView.TableID, mlngTable2ID)
    
      If blnChildOf1 And blnChildOf2 Then
        mstrErrorMessage = "Cannot use the '" & pobjTableView.TableName & "' table as it is a child table of both the '" & mstrTable1Name & "' and the '" & mstrTable2Name & "' tables."
        GenerateSQLJoin = False
        Exit Function

      ElseIf blnChildOf1 Then
        
        strOutputMain = strOutputMain & _
              " LEFT OUTER JOIN " & pobjTableView.RealSource & _
              " ON " & mstrTable1RealSource & ".ID = " & pobjTableView.RealSource & ".ID_" & CStr(mlngTable1ID) & vbCrLf
  
      ElseIf blnChildOf2 Then
  
        blnFound = False
        For Each objRelation In mcolRelations
          If objRelation.Table2ID = pobjTableView.TableID Then
            blnFound = True
            Exit For
          End If
        Next
        
        sCalcCode = vbNullString
  
        If blnFound Then
  
          If objRelation.RequiredExprID > 0 Then
            If objRelation.Table2ID > 0 Then
              sCalcCode = sCalcCode & _
                mcolSQLWhere("T" & CStr(objRelation.Table2ID)) & " = 1 "
            Else
              sCalcCode = sCalcCode & _
                mcolSQLWhere("T" & CStr(objRelation.Table1ID)) & " = 1 "
            End If
          
          End If
    
    
    If ASRDEVELOPMENT Then
      sCalcCode = Replace(sCalcCode, Chr(10), " ")
      sCalcCode = Replace(sCalcCode, Chr(11), " ")
      sCalcCode = Replace(sCalcCode, Chr(12), " ")
      sCalcCode = Replace(sCalcCode, Chr(13), " ")
      Do While InStr(sCalcCode, "  ")
        sCalcCode = Replace(sCalcCode, "  ", " ")
      Loop
    End If
  
          If objRelation.RequiredExprID > 0 Then
            strOutputMain = strOutputMain & _
                  " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & _
                  "(" & mstrTable2RealSource & ".ID = " & pobjTableView.RealSource & ".ID_" & CStr(mlngTable2ID) & vbCrLf & _
                  IIf(sCalcCode <> vbNullString, " AND " & sCalcCode, "") & ")" & vbCrLf
          End If
  
          If objRelation.PreferredExprID > 0 Then
            Set objCalcExpr = New clsExprExpression
            fOK = objCalcExpr.Initialise(objRelation.Table1ID, objRelation.PreferredExprID, giEXPR_MATCHJOINEXPRESSION, giEXPRVALUE_LOGIC, objRelation.Table2ID)
            If fOK Then
              fOK = objCalcExpr.RuntimeFilterCode(sTemp, True, False)
              If fOK And gbEnableUDFFunctions Then
                fOK = objCalcExpr.UDFFilterCode(mastrUDFsRequired(), True, False)
              End If
            End If
    
            If fOK Then
              For pintLoop1 = 1 To UBound(alngSourceTables, 2)
                AddToJoinArray Val(alngSourceTables(1, pintLoop1)), Val(alngSourceTables(2, pintLoop1))
              Next
            Else
              mstrErrorMessage = "You do not have permission to use the preferred match expression."
              Exit Function
            End If
            Set objCalcExpr = Nothing
    
            sCalcCode = sCalcCode & _
              IIf(sCalcCode <> vbNullString, " AND ", vbNullString) & sTemp & " = 1 "
    
          End If
  
  If ASRDEVELOPMENT Then
    sCalcCode = Replace(sCalcCode, Chr(10), " ")
    sCalcCode = Replace(sCalcCode, Chr(11), " ")
    sCalcCode = Replace(sCalcCode, Chr(12), " ")
    sCalcCode = Replace(sCalcCode, Chr(13), " ")
  
    Do While InStr(sCalcCode, "  ") > 0
      sCalcCode = Replace(sCalcCode, "  ", " ")
    Loop
  End If
  
          If objRelation.RequiredExprID = 0 Then
            strOutputMain = strOutputMain & _
                  " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & _
                  "(" & mstrTable2RealSource & ".ID = " & pobjTableView.RealSource & ".ID_" & CStr(mlngTable2ID) & vbCrLf & _
                  IIf(sCalcCode <> vbNullString, " AND " & sCalcCode, "") & ")" & vbCrLf
          End If
  
          If sCalcCode <> vbNullString Then
            If objRelation.Table1ID <> mlngTable1ID Then
  'MH20030909
              strOutputChildBreakdown = "FULL OUTER JOIN " & objRelation.Table2RealSource & " ON " & sCalcCode
              mcolSQLJoin.Add strOutputChildBreakdown, "T" & CStr(objRelation.Table1ID)
              'mcolSQLJoin.Add sCalcCode, "T" & CStr(objRelation.Table1ID)
            End If
          End If
  
        End If
  
      End If
    End If
  
  Next
  
  
  mcolSQLJoin.Add strOutputBaseBreakdown & strOutputMain & strOutputGrade, "T0"
  mcolSQLJoin.Add strOutputBaseBreakdown & strOutputGrade, "T" & CStr(mlngTable1ID)
  
  
  GenerateSQLJoin = True
  Exit Function

GenerateSQLJoin_ERROR:
  GenerateSQLJoin = False
  mstrErrorMessage = "Error in GenerateSQLJoin." & vbCrLf & Err.Description

End Function


Private Function GenerateSQLWhere(plngTableID As Long, plngRecordID As Long) As Boolean

  Dim objRelation As clsMatchRelation
  Dim objCalcExpr As clsExprExpression
  Dim strPicklistFilterSelect As String
  Dim sCalcCode As String
  Dim pintLoop1 As Long
  Dim strReportingStructure As String
  
  Dim lngTable1RecordID As Long
  Dim lngTable2RecordID As Long
  
  Set mcolSQLWhere = Nothing
  Set mcolSQLWhere = New Collection

  
  mstrTable1Where = vbNullString
  mstrTable2Where = vbNullString
  mstrSQLWhere = vbNullString
  
  
  'Single Record
  If plngRecordID > 0 Then
    If mlngMatchReportType = mrtSucession Then
      If mlngTable1ID = glngPostTableID Then
        lngTable1RecordID = GetJobTableID(plngRecordID)
      ElseIf mlngTable2ID = glngPostTableID Then
        lngTable2RecordID = GetJobTableID(plngRecordID)
      End If
    Else
      If mlngTable1ID = plngTableID Then
        lngTable1RecordID = plngRecordID
      Else
        lngTable2RecordID = plngRecordID
      End If
    End If
  End If


  strPicklistFilterSelect = GetPicklistFilterSelect(lngTable1RecordID, mlngTable1PickListID, mlngTable1FilterID)
  If fOK = False Then
    Exit Function
  End If
  If strPicklistFilterSelect <> vbNullString Then
    mstrTable1Where = mstrTable1Where & _
      IIf(mstrTable1Where <> vbNullString, " AND ", vbNullString) & _
      mstrTable1RealSource & ".ID IN (" & strPicklistFilterSelect & ")"
  End If


  strPicklistFilterSelect = GetPicklistFilterSelect(lngTable2RecordID, mlngTable2PickListID, mlngTable2FilterID)
  If fOK = False Then
    Exit Function
  End If
  If strPicklistFilterSelect <> vbNullString Then
    mstrTable2Where = mstrTable2Where & _
      IIf(mstrTable2Where <> vbNullString, " AND ", vbNullString) & _
      mstrTable2RealSource & ".ID IN (" & strPicklistFilterSelect & ")"
  End If


  For Each objRelation In mcolRelations
      
    If objRelation.RequiredExprID > 0 Then
      Set objCalcExpr = New clsExprExpression
'MH20030918 Fault 7005
'      If objRelation.Table2ID > 0 Then
'        fOK = objCalcExpr.Initialise(objRelation.Table2ID, objRelation.RequiredExprID, giEXPR_MATCHWHEREEXPRESSION, giEXPRVALUE_LOGIC, objRelation.Table1ID)
'      Else
'        fOK = objCalcExpr.Initialise(objRelation.Table1ID, objRelation.RequiredExprID, giEXPR_MATCHWHEREEXPRESSION, giEXPRVALUE_LOGIC, 0)
'      End If
      fOK = objCalcExpr.Initialise(objRelation.Table1ID, objRelation.RequiredExprID, giEXPR_MATCHWHEREEXPRESSION, giEXPRVALUE_LOGIC, objRelation.Table2ID)

      If fOK Then
        fOK = objCalcExpr.RuntimeFilterCode(sCalcCode, True, False)
        If fOK And gbEnableUDFFunctions Then
          fOK = objCalcExpr.UDFFilterCode(mastrUDFsRequired(), True, False)
        End If
      End If

      If fOK Then
        For pintLoop1 = 1 To UBound(alngSourceTables, 2)
          AddToJoinArray Val(alngSourceTables(1, pintLoop1)), Val(alngSourceTables(2, pintLoop1))
        Next
      Else
        'mstrErrorMessage = objCalcExpr.ErrorMessage
        mstrErrorMessage = "You do not have permission to use the required match expression."
        Exit Function
      End If

      Set objCalcExpr = Nothing
      
If ASRDEVELOPMENT Then
  sCalcCode = Replace(sCalcCode, Chr(10), " ")
  sCalcCode = Replace(sCalcCode, Chr(11), " ")
  sCalcCode = Replace(sCalcCode, Chr(12), " ")
  sCalcCode = Replace(sCalcCode, Chr(13), " ")
End If
      
      If objRelation.Table1ID <> mlngTable1ID And objRelation.Table2ID > 0 Then
        'If mlngMatchReportType = mrtNormal Then
        '  mstrTable2Where = mstrTable2Where & _
        '    IIf(mstrTable2Where <> vbNullString, " AND ", vbNullString) & _
        '    "count(distinct " & objRelation.Table1RealSource & ".ID) = " & _
        '    "count(distinct " & objRelation.Table2RealSource & ".ID)"
        'Else
        '  mstrTable1Where = mstrTable1Where & _
        '    IIf(mstrTable1Where <> vbNullString, " AND ", vbNullString) & _
        '    "count(distinct " & objRelation.Table1RealSource & ".ID) = " & _
        '    "count(distinct " & objRelation.Table2RealSource & ".ID)"
        'End If
        If mlngMatchReportType = mrtNormal Then
          mstrTable2Where = mstrTable2Where & _
            IIf(mstrTable2Where <> vbNullString, " AND ", vbNullString) & _
            "count(" & objRelation.Table1RealSource & ".ID) = " & _
            "count(" & objRelation.Table2RealSource & ".ID)"
        Else
          mstrTable1Where = mstrTable1Where & _
            IIf(mstrTable1Where <> vbNullString, " AND ", vbNullString) & _
            "count(" & objRelation.Table1RealSource & ".ID) = " & _
            "count(" & objRelation.Table2RealSource & ".ID)"
        End If
      Else
        mstrSQLWhere = mstrSQLWhere & _
          IIf(mstrSQLWhere <> vbNullString, " AND ", vbNullString) & _
          "(" & sCalcCode & ") = 1 "
      End If

      If objRelation.Table2ID > 0 Then
        mcolSQLWhere.Add sCalcCode, "T" & CStr(objRelation.Table2ID)
      Else
        mcolSQLWhere.Add sCalcCode, "T" & CStr(objRelation.Table1ID)
      End If

    End If
      
  Next
  
  GenerateSQLWhere = True

End Function


Private Function GetMatchReportDefinition() As Boolean
  
  On Error GoTo GetMatchReportDefinition_ERROR

  Dim rsTemp_Definition As Recordset
  Dim strSQL As String
  Dim lblnReportPackMode As Boolean
  
  lblnReportPackMode = gblnReportPackMode
  
  strSQL = "SELECT ASRSYSMatchReportName.*, " & _
           "a.TableName as 'Table1Name', " & _
           "a.RecordDescExprID as 'Table1RecDescExprID', " & _
           "b.TableName as 'Table2Name', " & _
           "b.RecordDescExprID as 'Table2RecDescExprID' " & _
           "FROM ASRSYSMatchReportName " & _
           "JOIN ASRSysTables a ON ASRSysMatchReportName.Table1ID = a.TableID " & _
           "LEFT OUTER JOIN ASRSysTables b ON ASRSysMatchReportName.Table2ID = b.TableID " & _
           "WHERE MatchReportID = " & CStr(mlngMatchReportID)

  Set rsTemp_Definition = datGeneral.GetReadOnlyRecords(strSQL)

  With rsTemp_Definition
  
    If .BOF And .EOF Then
      GetMatchReportDefinition = False
      mstrErrorMessage = "Could not find specified definition !"
      Exit Function
    End If
    
    mstrName = !Name
    mstrDescription = !Description
    mlngNumRecords = !NumRecords
    
    mlngScoreMode = IIf(IsNull(!ScoreMode), 0, !ScoreMode)
    mblnScoreCheck = IIf(IsNull(!ScoreCheck), False, !ScoreCheck)
    mlngScoreLimit = IIf(IsNull(!ScoreLimit), 0, !ScoreLimit)
    mblnEqualGrade = IIf(IsNull(!EqualGrade), False, !EqualGrade)
    mblnReportingStructure = IIf(IsNull(!ReportingStructure), 0, !ReportingStructure)
    
    mlngTable1ID = !Table1ID
    mstrTable1Name = !Table1Name
    mlngTable1RecDescExprID = !Table1RecDescExprID
    mlngTable1AllRecords = !Table1AllRecords
    mlngTable1PickListID = !Table1Picklist
    mlngTable1FilterID = !Table1Filter
    ' Override filter if in Report pack mode
    If mlngTable1ID = glngPersonnelTableID And gblnReportPackMode Then
      mlngTable1FilterID = mlngOverrideFilterID
    End If
    
    If Not TablePermission(!Table1ID) Then
      mstrErrorMessage = "You do not have permission to read the '" & !Table1Name & "' table either directly or through any views."
      GetMatchReportDefinition = False
      Exit Function
    End If
    
    If Not IsNull(!PrintFilterHeader) Then
      If !PrintFilterHeader Then
        If mlngTable1PickListID > 0 Then
          mstrRecordSelectionName = " (Base Table picklist: " & datGeneral.GetPicklistName(mlngTable1PickListID) & ")"
        ElseIf mlngTable1FilterID > 0 Then
          mstrRecordSelectionName = " (Base Table filter: " & datGeneral.GetFilterName(mlngTable1FilterID) & ")"
        End If
      End If
    End If
    
    mlngTable2ID = !Table2ID
    If mlngTable2ID > 0 Then
      mstrTable2Name = !Table2Name
      mlngTable2RecDescExprID = !Table2RecDescExprID
      mlngTable2AllRecords = !Table2AllRecords
      mlngTable2PickListID = !Table2Picklist
      mlngTable2FilterID = !Table2Filter
    
      If mlngTable2ID = glngPersonnelTableID And gblnReportPackMode Then
        mlngTable2FilterID = mlngOverrideFilterID
      End If

      If Not TablePermission(!Table2ID) Then
        mstrErrorMessage = "You do not have permission to read the '" & !Table2Name & "' table either directly or through any views."
        GetMatchReportDefinition = False
        Exit Function
      End If
    End If

    mbDefinitionOwner = (LCase(Trim(gsUserName)) = LCase(Trim(!UserName)))
    
    'Change Output Options to Report Pack owning these Jobs if in Report Pack mode
    mblnPreviewOnScreen = IIf(lblnReportPackMode, mblnPreviewOnScreen, !OutputPreview)
    mblnOutputScreen = IIf(lblnReportPackMode, mblnOutputScreen, !OutputScreen)
    mlngOutputFormat = IIf(lblnReportPackMode, mlngOutputFormat, !OutputFormat)
    mblnOutputPrinter = IIf(lblnReportPackMode, mblnOutputPrinter, !OutputPrinter)
    mstrOutputPrinterName = IIf(lblnReportPackMode, mstrOutputPrinterName, !OutputPrinterName)
    mblnOutputSave = IIf(lblnReportPackMode, mblnOutputSave, !OutputSave)
    mlngOutputSaveExisting = IIf(lblnReportPackMode, mlngOutputSaveExisting, !OutputSaveExisting)
    mblnOutputEmail = IIf(lblnReportPackMode, mblnOutputEmail, !OutputEmail)
    mlngOutputEmailAddr = IIf(lblnReportPackMode, mlngOutputEmailAddr, !OutputEmailAddr)
    mstrOutputEmailSubject = IIf(lblnReportPackMode, mstrOutputEmailSubject, !OutputEmailSubject)
    mstrOutputEmailAttachAs = IIf(lblnReportPackMode, mstrOutputEmailAttachAs, !OutputEmailAttachAs)
    mstrOutputFileName = IIf(lblnReportPackMode, mstrOutputFileName, !OutputFilename)
    mlngOverrideFilterID = IIf(lblnReportPackMode, mlngOverrideFilterID, 0)
    
  End With

  If Not gblnBatchMode Then
    If frmBreakDown Is Nothing Then
      Set frmBreakDown = New frmMatchRunBreakDown
    End If
    frmBreakDown.lblTable1Name.Caption = mstrTable1Name
    frmBreakDown.lblTable2Name.Caption = mstrTable2Name
  End If

  GetMatchReportDefinition = IsRecordSelectionValid
  
TidyAndExit:
  Set rsTemp_Definition = Nothing

Exit Function

GetMatchReportDefinition_ERROR:
  GetMatchReportDefinition = False
  mstrErrorMessage = "Error whilst reading the definition !" & vbCrLf & Err.Description
  Resume TidyAndExit

End Function


Private Function IsRecordSelectionValid() As Boolean
  Dim sSQL As String
  Dim lCount As Long
  Dim rsTemp As Recordset
  Dim iResult As RecordSelectionValidityCodes

' Base Table First
  If mlngTable1FilterID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngTable1FilterID)
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
  ElseIf mlngTable1PickListID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, mlngTable1PickListID)
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

  If Len(mstrErrorMessage) = 0 Then
    ' Criteria Table
    If mlngTable2FilterID > 0 Then
      iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngTable2FilterID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorMessage = "The match table filter used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorMessage = "The match table filter used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorMessage = "The match table filter used in this definition has been made hidden by another user."
          End If
      End Select
    ElseIf mlngTable2PickListID > 0 Then
      iResult = ValidateRecordSelection(REC_SEL_PICKLIST, mlngTable2PickListID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorMessage = "The match table picklist used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorMessage = "The match table picklist used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorMessage = "The match table picklist used in this definition has been made hidden by another user."
          End If
      End Select
    End If
  End If
  
  IsRecordSelectionValid = (Len(mstrErrorMessage) = 0)
  
End Function


Private Sub AddToJoinArray(lngType As Long, lngTableID As Long)

  Dim lngIndex As Integer

  If lngType = 0 Then   'Table
    If lngTableID = mlngTable1ID Or _
       lngTableID = mlngTable2ID Then
          Exit Sub
    End If
  End If

  For lngIndex = 1 To UBound(mlngTableViews, 2)
    If mlngTableViews(1, lngIndex) = lngType And _
      mlngTableViews(2, lngIndex) = lngTableID Then
      Exit Sub
    End If
  Next

  If lngTableID = 0 Then
    Stop
  End If

  'Only get here if not already in array
  lngIndex = UBound(mlngTableViews, 2) + 1
  ReDim Preserve mlngTableViews(2, lngIndex)
  mlngTableViews(1, lngIndex) = lngType
  mlngTableViews(2, lngIndex) = lngTableID

End Sub
      

Private Function GenerateSQLOrderBy() As Boolean

  Dim objColumn As clsColumn
  Dim lngIndex As Long
  
  mstrSQLOrderBy = " ORDER BY "
  For lngIndex = 1 To mcolColDetails.Count
  
    For Each objColumn In mcolColDetails
      If objColumn.SortSeq = lngIndex Then
        mstrSQLOrderBy = mstrSQLOrderBy & _
          IIf(lngIndex > 1, ", ", "") & _
          "[" & objColumn.Heading & "]" & _
          IIf(objColumn.SortDir = "D", " DESC", "")
      End If
    Next
  
  Next
  
  GenerateSQLOrderBy = True

End Function

Public Function FormatGrid(grdTemp As SSDBGrid, colColumns As Collection) As Boolean
  
  Dim objColumn As clsColumn
  Dim lngIndex As Long
  
  With grdTemp
    
    .Columns.RemoveAll
    For Each objColumn In colColumns

      lngIndex = .Columns.Count
      .Columns.Add lngIndex
      '.Columns(lngIndex).Caption = objColumn.Heading
      'NHRD25082004 Fault 7930
      .Columns(lngIndex).Caption = Replace(objColumn.Heading, "_", " ")
      .Columns(lngIndex).Visible = (objColumn.Hidden = False)
      .Columns(lngIndex).CaptionAlignment = ssCaptionAlignmentCenter
      '.ColumnHeaders(lngIndex).Alignment = ssCaptionAlignmentCenter
      
      Select Case objColumn.DataType
      Case sqlNumeric, sqlInteger
        .Columns(lngIndex).Alignment = ssCaptionAlignmentRight
      Case sqlBoolean
        .Columns(lngIndex).Alignment = ssCaptionAlignmentCenter
      End Select
    
    Next
  
  End With
  
  FormatGrid = True

End Function


Public Function PopulateGridBreakdown(lngTableID As Long) As Boolean

  Dim objRelation As clsMatchRelation
  Dim lngWidth As Long
  
  Set objRelation = mcolRelations("T" & CStr(lngTableID))
  
  If frmBreakDown Is Nothing Then
    Set frmBreakDown = New frmMatchRunBreakDown
  End If
  frmBreakDown.grdBreakdown.Redraw = False

  PopulateGridBreakdown = False
  If FormatGrid(frmBreakDown.grdBreakdown, objRelation.BreakdownColumns) Then
    If PopulateGrid(objRelation.BreakdownColumns, True, False) Then
      PopulateGridBreakdown = True
    End If
  End If
  
  'If its the first time we are looking at the breakdown
  'then size the breakdown form by the columns...
  If frmBreakDown.Visible = False Then
    With frmBreakDown.grdBreakdown
      lngWidth = (frmBreakDown.Width - frmBreakDown.ScaleWidth) + _
                 .Columns(.Cols - 1).Left + .Columns(.Cols - 1).Width + 270
      If lngWidth > Screen.Width Then
        lngWidth = Screen.Width
      End If
      frmBreakDown.Width = lngWidth
    End With
  End If

  frmBreakDown.chkAllRecords.Enabled = (objRelation.Table1ID <> mlngTable1ID)
  If objRelation.Table1ID = mlngTable1ID Then
    frmBreakDown.chkAllRecords.Value = False
  End If
  
  frmBreakDown.grdBreakdown.Redraw = True

  Set objRelation = Nothing

End Function


Private Function PopulateGridMain() As Boolean

  PopulateGridMain = False
  If FormatGrid(grdOutput, mcolColDetails) Then
    
'    If mlngMatchReportType <> mrtNormal Then
'      With grdOutput
'        .Columns.Add .Columns.Count
'        With .Columns(.Columns.Count - 1)
'          .Caption = vbNullString
'          .Locked = True
'          .Style = ssStyleButton
'          .ButtonsAlways = True
'          .BackColor = vbButtonFace   'required for printing !
'        End With
'      End With
'    End If

    mstrSQL = "SELECT DISTINCT * FROM [" & gsUserName & "].[" & mstrTempTableName & "]" & vbCrLf & _
      "WHERE not (ID1 is null) " & mstrSQLOrderBy

    If PopulateGrid(mcolColDetails, False, True) Then

      grdOutput.Caption = Replace(mstrName & mstrRecordSelectionName, "&", "&&")
      PopulateGridMain = True

    End If
  End If

End Function


Private Function PopulateGrid(colColumns As Collection, blnBreakdownOutput As Boolean, blnProgress As Boolean) As Boolean

  Dim datData As clsDataAccess
  Dim objColumn As clsColumn
  Dim rsMatchReportsData As Recordset
  Dim strOutput As String
  Dim lngColumnWidth() As Long
  Dim vData As Variant
  Dim vDataTemp As Variant
  Dim lngIndex As Long
  Dim lngTextWidth As Long
  Dim iCount As Integer
  Dim iCount2 As Integer
    
  Dim grdTemp As SSDBGrid
  Dim rsTemp As Recordset

  On Error GoTo 0


  Set datData = New clsDataAccess
  Set rsMatchReportsData = datData.OpenRecordsetAsync(mstrSQL, adOpenStatic, adLockReadOnly)
  Set datData = Nothing

  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    fOK = False
    Exit Function
  End If

  If rsMatchReportsData.BOF And rsMatchReportsData.EOF Then
    mstrErrorMessage = "No records meet selection criteria."
    mblnNoRecords = True
    fOK = False
    Exit Function
  End If
  
  
  If blnBreakdownOutput Then
    Set grdTemp = frmBreakDown.grdBreakdown
  Else
    Set grdTemp = grdOutput
  End If


  With rsMatchReportsData
    
    If Not blnBreakdownOutput Then
      If Not gblnBatchMode Then
        gobjProgress.Bar1MaxValue = .RecordCount
      Else
        gobjProgress.Bar2MaxValue = .RecordCount
      End If
    End If
    
    grdTemp.RemoveAll
    
    'Initialise column widths by size of heading text...
    ReDim lngColumnWidth(grdTemp.Columns.Count - 1)
    For lngIndex = 0 To grdTemp.Columns.Count - 1
      lngColumnWidth(lngIndex) = Me.TextWidth(grdTemp.Columns(lngIndex).Caption) + 195
    Next
    
    
    If Not .EOF Then
      .MoveFirst
      Do While Not .EOF

        strOutput = vbNullString
        'strOutput = .Fields(0).Value & vbTab & .Fields(1).Value
        For lngIndex = 0 To .Fields.Count - 1
          
          Set objColumn = colColumns(lngIndex + 1)
          vData = IIf(IsNull(.Fields(lngIndex).Value), vbNullString, .Fields(lngIndex).Value)
          vData = Replace(vData, vbCr, "")
          vData = Replace(vData, vbTab, "")
          
          If objColumn.IsNumeric Then
          
            If objColumn.DecPlaces > 0 Then
              vData = Format(vData, "0." & String(objColumn.DecPlaces, "0"))
            Else
              If objColumn.Size > 0 Then
                If vData = "0" Then
                  vData = Format(vData, "0")
                Else
                  vData = Format(vData, "#")
                End If
              End If
            End If

            If objColumn.ThousandSeparator Then
              vDataTemp = vData
              vData = ""
              iCount2 = 1
              If InStr(1, vDataTemp, ".") > 0 Then
                For iCount = InStr(1, vDataTemp, ".") - 1 To 1 Step -1
                  vData = IIf(iCount2 Mod 3 = 0 And iCount > 1, ",", "") & Mid(vDataTemp, iCount, 1) & vData
                  iCount2 = iCount2 + 1
                Next iCount
                vData = vData & "." & Right(vDataTemp, Len(vDataTemp) - InStr(1, vDataTemp, "."))
              Else
                For iCount = Len(vDataTemp) To 1 Step -1
                  vData = IIf(iCount2 Mod 3 = 0 And iCount > 1, ",", "") & Mid(vDataTemp, iCount, 1) & vData
                  iCount2 = iCount2 + 1
                Next iCount
              End If
            End If
            
          End If
          
          ' If its a date column, format it as dateformat
          If objColumn.DataType = sqlDate Then
            vData = Format(vData, DateFormat)
          End If
          
          If objColumn.Size > 0 Then   'Size restriction
            If objColumn.DecPlaces > 0 Then
              If InStr(vData, ".") > objColumn.Size Then
                vData = Left(vData, objColumn.Size) & _
                        Mid(vData, InStr(vData, "."))
              End If

            Else
              If Len(vData) > objColumn.Size Then
                vData = Left(vData, objColumn.Size)
              End If

            End If
          End If

          strOutput = strOutput & _
              IIf(lngIndex > 0, vbTab, "") & _
              vData

          lngTextWidth = BigTextWidth(vData, 0) + 195
          If lngTextWidth > lngColumnWidth(lngIndex) Then
            lngColumnWidth(lngIndex) = lngTextWidth
          End If

        Next

'          If mlngMatchReportType <> mrtNormal Then
'            strOutput = strOutput & _
'                IIf(lngIndex > 0, vbTab, "") & "..."
'          End If

        If Not gblnBatchMode And Not blnBreakdownOutput Then
          If Not IsNull(.Fields(0).Value) And Not IsNull(.Fields(1).Value) Then
            frmBreakDown.AddToCrossRef Val(.Fields(0).Value), Val(.Fields(1).Value)
          End If
        End If
        
        grdTemp.AddItem strOutput
        
        'If blnProgress Then
        '  gobjProgress.UpdateProgress gblnBatchMode
        'End If
        
        .MoveNext
      Loop
    End If
  
  End With
  
  rsMatchReportsData.Close
  Set rsMatchReportsData = Nothing
  
  For lngIndex = 0 To UBound(lngColumnWidth)
    If grdTemp.Columns(lngIndex).Visible Then
      grdTemp.Columns(lngIndex).Width = lngColumnWidth(lngIndex)
    End If
  Next

'  If mlngMatchReportType <> mrtNormal Then
'    grdTemp.Columns(grdTemp.Columns.Count - 1).Width = Me.TextWidth("...")
'  End If

  PopulateGrid = True

Exit Function

LocalErr:
  If ASRDEVELOPMENT Then
    Clipboard.Clear
    Clipboard.SetText mstrSQL
    COAMsgBox Err.Description
  End If
  PopulateGrid = False

End Function

Private Sub cmdClose_Click()
  Unload Me
End Sub


Private Function GetPicklistFilterSelect(lngSingleID As Long, lngPicklistID As Long, lngFilterID As Long) As String

  Dim rsTemp As Recordset

  On Error GoTo LocalErr
  
  
  If lngSingleID > 0 Then
    GetPicklistFilterSelect = CStr(lngSingleID)
  
  ElseIf lngPicklistID > 0 Then
    
    mstrErrorMessage = IsPicklistValid(lngPicklistID)
    If mstrErrorMessage <> vbNullString Then
      fOK = False
      Exit Function
    End If
    
    
    'Get List of IDs from Picklist
    Set rsTemp = datGeneral.GetReadOnlyRecords("EXEC sp_ASRGetPickListRecords " & CStr(lngPicklistID))
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
      fOK = False
      Exit Function
    End If
    
    'Get list of IDs from Filter
    fOK = datGeneral.FilteredIDs(lngFilterID, GetPicklistFilterSelect)

    ' Generate any UDFs that are used in this filter
    If fOK And gbEnableUDFFunctions Then
      datGeneral.FilterUDFs lngFilterID, mastrUDFsRequired()
    End If

    If Not fOK Then
      ' Permission denied on something in the filter.
      mstrErrorMessage = "You do not have permission to use the '" & datGeneral.GetFilterName(lngFilterID) & "' filter."
    End If

  End If

Exit Function

LocalErr:
  mstrErrorMessage = "Error processing record selection"
  fOK = False

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
End Sub

Private Sub Form_Load()
  
  Hook Me.hWnd, 9405, 6990
  
  If frmBreakDown Is Nothing Then
    Set frmBreakDown = New frmMatchRunBreakDown
  End If
  
  Set frmOutput = New frmOutputOptions
  frmOutput.PageRange = False
  ReDim mastrUDFsRequired(0)
End Sub

Private Sub Form_Resize()

  Const lngGap As Long = 120

  'JPD 20030908 Fault 5756
  DisplayApplication

'  If Me.WindowState = vbNormal Then
'    If Me.Height > Screen.Height Then Me.Height = 6990
'    If Me.Width > Screen.Width Then Me.Width = 9405
'    If Me.Height < 6990 Then Me.Height = 6990
'    If Me.Width < 9405 Then Me.Width = 9405
'  End If

  If Me.WindowState <> vbMinimized Then
    cmdClose.Top = Me.ScaleHeight - (cmdClose.Height + lngGap)
    cmdClose.Left = Me.ScaleWidth - (cmdClose.Width + lngGap)
    cmdOutput.Top = cmdClose.Top
    cmdOutput.Left = cmdClose.Left - (cmdOutput.Width + lngGap)
    Me.grdOutput.Move lngGap, lngGap, Me.ScaleWidth - (lngGap * 2), cmdClose.Top - (lngGap * 2)
  End If

  ' Get rid of the icon off the form
  RemoveIcon Me

End Sub


Private Function InitialiseFormBreakdown()

  Dim objRelation As clsMatchRelation

  If frmBreakDown Is Nothing Then
    Set frmBreakDown = New frmMatchRunBreakDown
  End If
  frmBreakDown.Caption = Me.Caption & " Breakdown"
  frmBreakDown.HelpContextID = Me.HelpContextID

  With frmBreakDown
    .Loading = True
  
    .ParentForm = Me
    .lblTable1Name.Caption = mstrTable1Name & " :"
    .Table1RecDescExprID = mlngTable1RecDescExprID
  
    If mlngTable2ID = 0 Then
      .lblTable2Name.Visible = False
      .cboTable2.Visible = False
      .Frame1.Top = .Frame1.Top - 400
      Me.Height = Me.Height - 400
    Else
      .lblTable2Name.Caption = mstrTable2Name & " :"
      .Table2RecDescExprID = mlngTable2RecDescExprID
    End If
    
    
    With .cboRelation
      .Clear
      For Each objRelation In mcolRelations
        .AddItem objRelation.Table1Name
        'Itemdata needs to be Table2ID so that it can reference the relation collection
        .ItemData(.NewIndex) = objRelation.Table1ID
      Next
      If .ListCount > 0 Then
        .ListIndex = 0
      End If
    End With
  
    .Loading = False
  
  End With
  
  InitialiseFormBreakdown = True

End Function


Private Sub Form_Unload(Cancel As Integer)
  Unload frmBreakDown
  Set frmBreakDown = Nothing
  
  Unload frmOutput
  Set frmOutput = Nothing
  
  Unhook Me.hWnd
End Sub

Private Sub grdOutput_DblClick()
  
  Dim lngRecord1ID As Long
  Dim lngRecord2ID As Long
  
  If mlngTable1ID > 0 And mlngTable1RecDescExprID = 0 Then
    COAMsgBox "Unable to show cell breakdown details as no record description " & _
           "has been set up for the '" & mstrTable1Name & "' table.", vbInformation, Me.Caption
    Exit Sub
  End If
  
  If mlngTable2ID > 0 And mlngTable2RecDescExprID = 0 Then
    COAMsgBox "Unable to show cell breakdown details as no record description " & _
           "has been set up for the '" & mstrTable2Name & "' table.", vbInformation, Me.Caption
    Exit Sub
  End If


  With grdOutput
    lngRecord1ID = Val(.Columns(0).CellValue(.Bookmark))
    lngRecord2ID = Val(.Columns(1).CellValue(.Bookmark))
  End With

  UDFFunctions mastrUDFsRequired, True
  frmBreakDown.ShowBreakdown lngRecord1ID, lngRecord2ID, mlngMatchReportType
  UDFFunctions mastrUDFsRequired, False

End Sub


Private Sub cmdOutput_Click()
  OutputReport True
End Sub


Private Function OutputReport(blnPrompt As Boolean) As Boolean

  Dim objOutput As clsOutputRun
  Dim objColumn As clsColumn

  Set objOutput = New clsOutputRun

'  If objOutput.SetOptions(blnPrompt, mlngOutputFormat, mblnOutputScreen, _
'      mblnOutputPrinter, mstrOutputPrinterName, _
'      mblnOutputSave, mlngOutputSaveExisting, _
'      mblnOutputEmail, mlngOutputEmailAddr, mstrOutputEmailSubject, _
'      mstrOutputEmailAttachAs, mstrOutputFileName) Then

  If objOutput.SetOptions _
      (blnPrompt, _
      mlngOutputFormat, _
      mblnOutputScreen, _
      mblnOutputPrinter, _
      mstrOutputPrinterName, _
      mblnOutputSave, _
      mlngOutputSaveExisting, _
      mblnOutputEmail, _
      mlngOutputEmailAddr, _
      mstrOutputEmailSubject, _
      mstrOutputEmailAttachAs, _
      mstrOutputFileName, _
      False, _
      mblnPreviewOnScreen, _
      mstrOutputTitlePage, _
      mstrOutputReportPackTitle, _
      mstrOutputOverrideFilter, _
      mblnOutputTOC, _
      mblnOutputCoverSheet, _
      mlngOverrideFilterID) Then
      
    objOutput.PageTitles = False

    If Not gblnBatchMode Then
      objOutput.OpenProgress Me.Caption, mstrName, 1
    End If

    objOutput.SizeColumnsIndependently = True
    If objOutput.GetFile Then
      objOutput.AddPage mstrName & mstrRecordSelectionName, mstrTable1Name
  
      For Each objColumn In mcolColDetails
        'Ignore hidden columns
        If objColumn.Heading <> vbNullString And objColumn.Hidden = False Then
          objOutput.AddColumn objColumn.Heading, objColumn.DataType, objColumn.DecPlaces, objColumn.ThousandSeparator
        End If
      Next
  
      objOutput.DataGrid grdOutput
  
      If Not gblnBatchMode Then
        gobjProgress.CloseProgress
      End If
  
      objOutput.Complete

    End If

    mblnUserCancelled = objOutput.UserCancelled
    mstrErrorMessage = objOutput.ErrorMessage
    fOK = (mstrErrorMessage = vbNullString)

  Else
    blnPrompt = (blnPrompt And Not objOutput.UserCancelled)
    mstrErrorMessage = objOutput.ErrorMessage
    fOK = (mstrErrorMessage = vbNullString)

  End If


  If blnPrompt Then
    gobjProgress.CloseProgress
    
    'MH20040302 Fault 8143
    'Not ideal but the only way to prevent a runtime error was a doevents
    DoEvents
    
    If fOK Then
      COAMsgBox Me.Caption & ": '" & mstrName & "' output complete.", _
          vbInformation, Me.Caption
    Else
      COAMsgBox Me.Caption & ": '" & mstrName & "' output failed." & vbCrLf & vbCrLf & mstrErrorMessage, _
          vbExclamation, Me.Caption
    End If
  End If

  Set objOutput = Nothing

  OutputReport = fOK

End Function


Public Property Get PreviewOnScreen() As Boolean
  PreviewOnScreen = ((fOK And mblnPreviewOnScreen) And Not gblnBatchMode And Not mblnNoRecords)
End Property


Private Function GetReportingStructure(lngSingleRecord As Long) As String

  Dim objColumnPrivileges As CColumnPrivileges
  Dim blnColumnOK As Boolean
  
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strSQL1 As String
  Dim strSQL2 As String
  Dim strResult As String
  Dim strLastResult As String
  
  Dim strESelect As String
  Dim strEJoin As String
  Dim strMSelect As String
  Dim strMJoin As String
  Dim strViewIDs As String
  
  If mblnReportingStructure Then
    
    strViewIDs = vbNullString
    GetSelectAndJoinForColumn glngPersonnelTableID, gsPersonnelTableName, gsPersonnelEmployeeNumberColumnName, strESelect, strEJoin, strViewIDs
    GetSelectAndJoinForColumn glngPersonnelTableID, gsPersonnelTableName, gsPersonnelManagerStaffNoColumnName, strMSelect, strMJoin, strViewIDs

  
  
    strSQL = "SELECT " & gsPersonnelTableName & ".ID, " & strESelect & _
             " FROM " & gsPersonnelTableName & strEJoin & _
             " WHERE " & gsPersonnelTableName & ".ID = " & CStr(lngSingleRecord)
  
    If mlngMatchReportType = mrtSucession Then
      strSQL1 = "SELECT " & gsPersonnelTableName & ".ID, " & strESelect & _
                " FROM " & gsPersonnelTableName & strEJoin & _
                " WHERE " & strMSelect & " IN ("
      strSQL2 = ")"
    Else
      strSQL1 = "SELECT " & gsPersonnelTableName & ".ID, " & strESelect & _
                " FROM " & gsPersonnelTableName & strEJoin & _
                " WHERE " & strESelect & " IN (" & _
                "SELECT " & strMSelect & _
                " FROM " & gsPersonnelTableName & _
                strEJoin & strMJoin & _
                " WHERE " & strESelect & " IN ("
      strSQL2 = "))"
    End If
  
  
    strResult = "0"
    Do
      Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)

      strLastResult = vbNullString
      If Not rsTemp.BOF Or Not rsTemp.EOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
          If Not IsNull(rsTemp.Fields(1).Value) Then
            If Trim(rsTemp.Fields(1).Value) <> vbNullString Then
  
              If mlngMatchReportType = mrtSucession Then
                strResult = strResult & _
                  IIf(strResult <> vbNullString, ", ", "") & _
                  rsTemp.Fields(0).Value
              Else
                strResult = strResult & _
                  IIf(strResult <> vbNullString, ", ", "") & _
                  CStr(GetJobTableID(rsTemp.Fields(0).Value))
              End If
  
              strLastResult = strLastResult & _
                IIf(strLastResult <> vbNullString, ", ", "") & _
                "'" & CStr(rsTemp.Fields(1).Value) & "'"
  
            End If
          End If
          rsTemp.MoveNext
        Loop
      End If
  
      rsTemp.Close
      Set rsTemp = Nothing
  
      strSQL = strSQL1 & strLastResult & strSQL2
  
    Loop While strLastResult <> vbNullString


    If strResult <> vbNullString Then
      If mlngMatchReportType = mrtSucession Then
        strResult = _
          IIf(mlngTable1ID = glngPersonnelTableID, mstrTable1RealSource, mstrTable2RealSource) & _
          ".ID IN (" & strResult & ")"
      Else
        strResult = _
          IIf(mlngTable1ID = glngPostTableID, mstrTable1RealSource, mstrTable2RealSource) & _
          ".ID IN (" & strResult & ")"
      End If
    End If

  End If
    
  strResult = strResult & _
    IIf(strResult <> vbNullString, " AND ", vbNullString) & _
    "ASRSys" & gsPersonnelTableName & gstrGradeTableName & "." & gstrNumLevelColumnName & _
    " <" & IIf(mblnEqualGrade, "=", "") & " " & _
    "ASRSys" & gstrPostTableName & gstrGradeTableName & "." & gstrNumLevelColumnName
  
  GetReportingStructure = strResult

End Function


Private Function GetJobTableID(lngRecordID As Long) As Long

  Dim mobjColumnPrivileges As CColumnPrivileges

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strJoin As String

  'If Not gcoTablePrivileges.Item(gsPersonnelTableName).AllowSelect Then
  '  mstrErrorMessage = "Unable to run this report as you do not have access to the " & gsPersonnelTableName & " Table"
  '  Exit Function
  'End If

  'If Not gcoTablePrivileges.Item(gstrPostTableName).AllowSelect Then
  '  mstrErrorMessage = "Unable to run this report as you do not have access to the " & gstrPostTableName & " Table"
  '  Exit Function
  'End If


  strSQL = GetSQLForColumn(glngPostTableID, gstrPostTableName, gstrJobTitleColumnName, 1) & _
           " = (" & _
           GetSQLForColumn(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelJobTitleColumnName, 2) & _
           " = " & CStr(lngRecordID) & ")"
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)

  If Not (rsTemp.BOF And rsTemp.EOF) Then
    GetJobTableID = rsTemp.Fields("ID").Value
  Else
    GetJobTableID = 0
  End If

  rsTemp.Close
  Set rsTemp = Nothing

End Function


Private Function GetSQLForColumn(lngTableID As Long, strTable As String, strColumn As String, intMode As Integer) As String

  Dim strSelect As String
  Dim strJoin As String

  GetSelectAndJoinForColumn lngTableID, strTable, strColumn, strSelect, strJoin, vbNullString

  If strSelect = vbNullString Then
    mstrErrorMessage = vbCrLf & vbCrLf & "You do not have permission to see the column '" & strColumn & "'" & vbCrLf & "either directly or through any views."
  Else
    If intMode = 1 Then
      GetSQLForColumn = _
        "SELECT " & strTable & ".ID FROM " & strTable & _
        strJoin & " WHERE " & strSelect
    Else
      GetSQLForColumn = _
        "SELECT " & strSelect & " FROM " & strTable & _
        strJoin & " WHERE " & strTable & ".ID"
    End If
  End If

End Function


Private Sub GetSelectAndJoinForColumn(lngTableID As Long, strTable As String, strColumn As String, ByRef strSelect As String, ByRef strJoin As String, ByRef strViewIDs As String)

  Dim mobjColumnPrivileges As CColumnPrivileges
  Dim mobjTableView As CTablePrivilege
  Dim pblnColumnOK As Boolean

  Set mobjColumnPrivileges = GetColumnPrivileges(strTable)

  pblnColumnOK = mobjColumnPrivileges.IsValid(strColumn)
  If pblnColumnOK Then
    pblnColumnOK = mobjColumnPrivileges.Item(strColumn).AllowSelect
  End If

  If pblnColumnOK Then
    strSelect = strTable & "." & strColumn
    strJoin = vbNullString
  Else
    strSelect = vbNullString
    strJoin = vbNullString
    For Each mobjTableView In gcoTablePrivileges.Collection
      If (Not mobjTableView.IsTable) And _
        (mobjTableView.TableID = lngTableID) And _
        (mobjTableView.AllowSelect) Then

        Set mobjColumnPrivileges = GetColumnPrivileges(mobjTableView.ViewName)
        If mobjColumnPrivileges.IsValid(strColumn) Then
          If mobjColumnPrivileges.Item(strColumn).AllowSelect Then

            strSelect = strSelect & _
              " WHEN NOT " & mobjTableView.ViewName & "." & strColumn & _
              " IS NULL THEN " & mobjTableView.ViewName & "." & strColumn

            If InStr(strViewIDs, CStr(mobjTableView.ViewID)) = 0 Then
              strJoin = strJoin & _
                " LEFT OUTER JOIN " & mobjTableView.ViewName & _
                " ON " & mobjTableView.TableName & ".ID = " & mobjTableView.ViewName & ".ID" & vbCrLf
              strViewIDs = strViewIDs & " " & CStr(mobjTableView.ViewID)
            End If

          End If
        End If
      End If

    Next mobjTableView

    If strSelect <> vbNullString Then
      strSelect = "CASE" & strSelect & " ELSE NULL END"
    End If

  End If

End Sub


Private Function GetCalcCode(lngTable1ID As Long, lngTable2ID As Long, lngExprID As Long, lngExprType As Integer, lngReturnType As Integer) As Boolean

  Dim objCalcExpr As clsExprExpression
  Dim sCalcCode As String
  Dim blnOK As Boolean
  Dim iLoop1 As Long
  
  Set objCalcExpr = New clsExprExpression

  blnOK = objCalcExpr.Initialise(lngTable1ID, lngExprID, lngExprType, lngReturnType, lngTable2ID)
  
  If blnOK Then
    blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sCalcCode, True)
    If blnOK And gbEnableUDFFunctions Then
      blnOK = objCalcExpr.UDFCalculationCode(alngSourceTables, mastrUDFsRequired(), True)
    End If
  End If
  Set objCalcExpr = Nothing

  If blnOK Then
    ' Add the required views to the JOIN code.
    For iLoop1 = 1 To UBound(alngSourceTables, 2)
      AddToJoinArray alngSourceTables(1, iLoop1), alngSourceTables(2, iLoop1)
    Next iLoop1
  Else
    ' Permission denied on something in the calculation.
    mstrErrorMessage = "You do not have permission to use a match score calculation."
  End If

End Function



Private Function GetTempTable() As String

  Dim objColumn As clsColumn
  Dim strTempTable As String
  Dim strError As String
  Dim strSQL As String
  Dim lngIndex As Long
  Dim lngSize As Long
  
  For Each objColumn In mcolColDetails
    strSQL = strSQL & IIf(strSQL <> vbNullString, ", ", vbNullString) & vbCrLf

    'If objColumn.ColType = "C" Then
    '  strSQL = strSQL & "[" & objColumn.TableName & "-" & objColumn.ColumnName & "] "
    'Else
    '  strSQL = strSQL & "[Match Score] "
    'End If
    strSQL = strSQL & "[" & objColumn.Heading & "]"

    Select Case objColumn.DataType
    Case sqlVarChar, sqlLongVarChar   'sqlLongVarChar = Working Pattern
      lngSize = datGeneral.GetDataSize(objColumn.ID)
      strSQL = strSQL & "[varchar] (" & IIf(lngSize = VARCHAR_MAX_Size, "MAX", lngSize) & ")"
    Case sqlBoolean
      strSQL = strSQL & "[varchar] (1)"
    Case sqlDate
      strSQL = strSQL & "[datetime]"
    Case sqlNumeric, sqlInteger
      strSQL = strSQL & "[float]"
    Case Else
      strSQL = strSQL & "[int]"
    End Select
  
    strSQL = strSQL & " NULL"
  Next

  strTempTable = datGeneral.UniqueSQLObjectName("ASRSysTempMatchReport", 3)
  strSQL = "CREATE TABLE [" & gsUserName & "].[" & strTempTable & "]" & _
           " (" & strSQL & ")"
  
  datGeneral.ExecuteSql strSQL, strError
  mstrErrorMessage = strError
  fOK = (mstrErrorMessage = vbNullString)
  
  GetTempTable = strTempTable

End Function


Private Function RemoveTemporarySQLObjects()

  UDFFunctions mastrUDFsRequired, False
  datGeneral.DropUniqueSQLObject mstrTempTableName, 3

End Function

Private Function TablePermission(lngTableID As Long) As Boolean

  Dim objTableView As CTablePrivilege
  Dim blnFound As Boolean

  blnFound = False
  For Each objTableView In gcoTablePrivileges.Collection
    If (objTableView.TableID = lngTableID) And _
      (objTableView.AllowSelect) Then
      blnFound = True
      Exit For
    End If
  Next objTableView
  Set objTableView = Nothing

  TablePermission = blnFound

End Function


Private Function HasColumnPermission(lngTableID As Long, strTable As String, strColumn As String) As Boolean

  Dim mobjColumnPrivileges As CColumnPrivileges
  Dim mobjTableView As CTablePrivilege
  Dim pblnColumnOK As Boolean

  Set mobjColumnPrivileges = GetColumnPrivileges(strTable)

  pblnColumnOK = mobjColumnPrivileges.IsValid(strColumn)
  If pblnColumnOK Then
    pblnColumnOK = mobjColumnPrivileges.Item(strColumn).AllowSelect
  End If


  If Not pblnColumnOK Then
    
    For Each mobjTableView In gcoTablePrivileges.Collection
      If (Not mobjTableView.IsTable) And _
        (mobjTableView.TableID = lngTableID) And _
        (mobjTableView.AllowSelect) Then

        Set mobjColumnPrivileges = GetColumnPrivileges(mobjTableView.ViewName)
        If mobjColumnPrivileges.IsValid(strColumn) Then
          If mobjColumnPrivileges.Item(strColumn).AllowSelect Then
            pblnColumnOK = True
            Exit For
          End If
        End If
      
      End If
    Next
  
  End If

  HasColumnPermission = pblnColumnOK

End Function


Private Function CheckModuleSetupPermissions() As Boolean

  If mlngMatchReportType <> mrtNormal Then

    CheckModuleSetupPermissions = False

    If Not HasColumnPermission(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelGradeColumnName) Then
      mstrErrorMessage = "You do not have permission to see the column '" & gsPersonnelTableName & "." & gsPersonnelGradeColumnName & "' either directly or through any views."
      Exit Function
    End If
    If Not HasColumnPermission(glngPostTableID, gstrPostTableName, gstrPostGradeColumnName) Then
      mstrErrorMessage = "You do not have permission to see the column '" & gstrPostTableName & "." & gstrPostGradeColumnName & "' either directly or through any views."
      Exit Function
    End If
    If Not HasColumnPermission(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelJobTitleColumnName) Then
      mstrErrorMessage = "You do not have permission to see the column '" & gsPersonnelTableName & "." & gsPersonnelJobTitleColumnName & "' either directly or through any views."
      Exit Function
    End If
    If Not HasColumnPermission(glngPostTableID, gstrPostTableName, gstrJobTitleColumnName) Then
      mstrErrorMessage = "You do not have permission to see the column '" & gstrPostTableName & "." & gstrJobTitleColumnName & "' either directly or through any views."
      Exit Function
    End If

  End If

  CheckModuleSetupPermissions = True

End Function
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
          blnChkPicklistFilter As Boolean, _
          Optional strOutputTitlePage As String, _
          Optional strOutputReportPackTitle As String, _
          Optional strOutputOverrideFilter As String, _
          Optional blnOutputTOC As Boolean, _
          Optional blnOutputCoverSheet As Boolean, _
          Optional lngOverrideFilterID As Long)

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
  mstrOutputFileName = strOutputFilename
  mblnChkPicklistFilter = blnChkPicklistFilter
  mblnPreviewOnScreen = (blnPreviewOnScreen Or (mlngOutputFormat = fmtDataOnly And mblnOutputScreen))
  mstrOutputTitlePage = IIf(IsMissing(strOutputTitlePage), giEXPRVALUE_CHARACTER, strOutputTitlePage)
  mstrOutputReportPackTitle = IIf(IsMissing(strOutputReportPackTitle), giEXPRVALUE_CHARACTER, strOutputReportPackTitle)
  mstrOutputOverrideFilter = IIf(IsMissing(strOutputOverrideFilter), giEXPRVALUE_CHARACTER, strOutputOverrideFilter)
  mblnOutputTOC = IIf(IsMissing(blnOutputTOC), giEXPRVALUE_CHARACTER, blnOutputTOC)
  mblnOutputCoverSheet = IIf(IsMissing(blnOutputCoverSheet), giEXPRVALUE_CHARACTER, blnOutputCoverSheet)
  mlngOverrideFilterID = IIf(IsMissing(lngOverrideFilterID), giEXPRVALUE_CHARACTER, lngOverrideFilterID)
End Sub


