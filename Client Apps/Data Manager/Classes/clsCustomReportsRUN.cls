VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomReportsRUN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' To hold Properties
Private mlngCustomReportID As Long
Private mstrErrorString As String
' String to hold the temp table name
Private mstrTempTableName As String
' Variables to store definition
Private mstrCustomReportsName As String
Private mstrCustomReportsDescription As String
Private mlngCustomReportsBaseTable As Long
Private mstrCustomReportsBaseTableName As String
Private mlngCustomReportsAllRecords As Long
Private mlngCustomReportsPickListID As Long
Private mlngCustomReportsFilterID As Long
Private mlngCustomReportsParent1Table As Long
Private mstrCustomReportsParent1TableName As String
Private mlngCustomReportsParent1FilterID As Long
Private mlngCustomReportsParent2Table As Long
Private mstrCustomReportsParent2TableName As String
Private mlngCustomReportsParent2FilterID As Long
Private mblnCustomReportsSummaryReport As Boolean
Private mblnIgnoreZerosInAggregates As Boolean
Private mblnCustomReportsPrintFilterHeader As Boolean
Private mintReportJobNumberOrder As Integer
'New Default Output Variables
Private mlngOriginalOutputFormat As Long
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
Private mstrOutputFileName As String
Private mblnPreviewOnScreen As Boolean
Private mstrOutputTitlePage As String
Private mstrOutputReportPackTitle As String
Private mstrOutputOverrideFilter As String
Private mblnOutputTOC As Boolean
Private mblnOutputCoverSheet As Boolean
Private mlngOverrideFilterID As Long
Private mblnOutputRetainPivotOrChart As Boolean
'Private mblnOutputRetainCharts As Boolean
          
Private mlngPicklistFilterID As Long
Private mstrPicklistFilterType As String

Private mvarChildTables() As Variant
Private miChildTablesCount As Integer
Private miUsedChildCount As Integer

Private miColumnsInReport As Integer

Private mlngCustomReportsParent1AllRecords As Long
Private mlngCustomReportsParent1PickListID As Long
Private mlngCustomReportsParent2AllRecords As Long
Private mlngCustomReportsParent2PickListID As Long

'' OLD Default Output Variables
'Private mintCustomReportsDefaultOutput As Integer
'Private mintCustomReportsDefaultExportTo As Integer
'Private mblnCustomReportsDefaultSave As Boolean
'Private mstrCustomReportsDefaultSaveAs As String
'Private mblnCustomReportsDefaultCloseApp As Boolean

' Recordsets to store the definition and column information
Private mrstCustomReportsDetails As New Recordset

' Recordset to store the final data from the temp table
Private mrstCustomReportsOutput As New Recordset

' Strings to hold the SQL statement
Private mstrSQLSelect As String
Private mstrSQLFrom As String
Private mstrSQLJoin As String
Private mstrSQLWhere As String
Private mstrSQLOrderBy As String
Private mstrSQL As String

' Data access class
Private mclsData As clsDataAccess
Private mclsGeneral As clsGeneral

' Array holding the columns to sort the report by
Private mvarSortOrder() As Variant

' Array to hold the columns used in the report
Dim mvarColDetails() As Variant
'Dim mstrExcelFormats() As String
'Dim mlngExcelDecimals() As Long

' Array to hold the column widths used when formatting the grid
Dim mlngColWidth() As Long

'Array used to store the 'GroupWithNextColumn' option strings.
Private mvarGroupWith() As Variant

'Array used to store the 'POC' values when outputting.
Private mvarPageBreak() As String
Private mblnPageBreak As Boolean
Private mintPageBreakRowIndex As Long

' Flags used when populating the grid
Private mblnReportHasSummaryInfo As Boolean
Private mblnReportHasPageBreak As Boolean
Private mblnDoesHaveGrandSummary As Boolean

' Instance of the previewform
Private mfrmOutput As frmCustomReportsPreview

' TableViewsGuff
Private mstrRealSource As String
Private mstrBaseTableRealSource As String
Private mlngTableViews() As Long
Private mstrViews() As String
Private mobjTableView As CTablePrivilege
Private mobjColumnPrivileges As CColumnPrivileges

'Batch Job Mode ?
'Private gblnBatchMode As Boolean

'Has the user cancelled the report ?
Private mblnUserCancelled As Boolean

'Does the report generate no records ?
Private mblnNoRecords As Boolean

' Is this a Bradford Index Report
Private mbIsBradfordIndexReport As Boolean
Private mbBradfordSRV As Boolean
Private mbBradfordTotals As Boolean
Private mbBradfordCount As Boolean
Private mbBradfordWorkings As Boolean
Private mstrPicklistFilterIDs As String
Private mstrOrderByColumn As String
Private mlngOrderByColumnID As Long
Private mstrGroupByColumn As String
Private mlngGroupByColumnID As Long
Private mbOrderBy1Asc As Boolean
Private mbOrderBy2Asc As Boolean
Private mblnOmitBeforeStart As Boolean
Private mblnOmitAfterEnd  As Boolean
Private msAbsenceTypes As String
Private mdtAbsenceTo As Date
Private mdtAbsenceFrom As Date
Private mbMinBradford As Boolean
Private mlngMinBradfordAmount As Long
Private mbDisplayBradfordDetail As Boolean
Private mbDefinitionOwner As Boolean
Private mblnInvalidPicklistFilter As Boolean
Private mstrAbsenceRealSource As String


Private Const lng_SEQUENCECOLUMNNAME = "?ID_SEQUENCE_COLUMN"

'Private msBaseRecordIDColumn As String

Private mbUseSequence As Boolean

' Array holding the User Defined functions that are needed for this report
Private mastrUDFsRequired() As String

Private Function AddPageBreakValue(pintRowIndex As Long, pvarValue As String)
  ReDim Preserve mvarPageBreak(pintRowIndex)
  mvarPageBreak(pintRowIndex) = pvarValue
End Function

Private Function GetMostChildsForParent(avChildRecs() As Recordset, iParentCount As Integer) As Integer

  Dim i As Integer
  Dim iMostChildRecords As Integer
  Dim iChildRecordCount As Integer

  On Error GoTo Error_Trap
  
  iMostChildRecords = 0
  iChildRecordCount = 0
  
  For i = 0 To UBound(avChildRecs, 2) Step 1
    If (avChildRecs(iParentCount, i).BOF) And (avChildRecs(iParentCount, i).EOF) Then
      iChildRecordCount = 0
    Else
      iChildRecordCount = avChildRecs(iParentCount, i).RecordCount
    End If
    If iChildRecordCount > iMostChildRecords Then
      iMostChildRecords = iChildRecordCount
    End If
  Next i

  GetMostChildsForParent = iMostChildRecords
  
  Exit Function
  
Error_Trap:
  GetMostChildsForParent = 0

End Function



Public Property Get ReportJobNumberOrder() As Integer
  ReportJobNumberOrder = mintReportJobNumberOrder
End Property

Public Property Let ReportJobNumberOrder(intJob As Integer)
  mintReportJobNumberOrder = intJob
End Property




Public Property Get InvalidPicklistFilter() As Boolean
  InvalidPicklistFilter = mblnInvalidPicklistFilter
End Property
Private Property Let InvalidPicklistFilter(bValid As Boolean)
  mblnInvalidPicklistFilter = bValid
End Property

Private Function IsChildTableUsed(iChildTableID As Long) As Boolean

  Dim i As Integer
  
  IsChildTableUsed = False
  
  For i = 1 To UBound(mvarColDetails, 2) Step 1
    If mvarColDetails(14, i) = iChildTableID Then
      IsChildTableUsed = True
      Exit Function
    End If
  Next i

End Function

Private Function IsReportChildTable(lngTableID As Long) As Boolean

  Dim i As Integer
  
  IsReportChildTable = False 'cowardly programming!!!
  
  If miChildTablesCount > 0 Then
    For i = 0 To UBound(mvarChildTables, 2) Step 1
      If lngTableID = mvarChildTables(0, i) Then
        IsReportChildTable = True
        Exit Function
      End If
    Next i
  End If

End Function

Private Function IsReportParentTable(lngTableID As Long) As Boolean

  IsReportParentTable = False 'cowardly programming!!!
  
  If lngTableID = mlngCustomReportsParent1Table _
      Or lngTableID = mlngCustomReportsParent2Table Then
    IsReportParentTable = True
  End If
  
End Function

Private Function IsReportBaseTable(lngTableID As Long) As Boolean

  IsReportBaseTable = False 'cowardly programming!!!
  
  If lngTableID = mlngCustomReportsBaseTable Then
    IsReportBaseTable = True
  End If
  
End Function

Private Function CreateMutipleChildTempTable() As Boolean
 
  Dim sMCTempTable As String
  Dim sSQL As String
  Dim iColCount As Integer
  Dim sParentSelectSQL As String
  Dim rsParent As ADODB.Recordset
  Dim lngColumnID As Long
  Dim lngTableID As Long
  Dim iChildCount As Integer
  Dim rsChild As ADODB.Recordset
  Dim iParentCount As Integer
  Dim avChildRecordsets() As ADODB.Recordset
  Dim sChildSelectSQL As String
  Dim sChildWhereSQL As String
  Dim iFields As Integer
'  Dim rsTemp As ADODB.Recordset
  Dim i As Integer
  Dim iChildUsed As Integer
  Dim iMostChilds As Integer
  'Dim sTempFieldName As String
  Dim lngCurrentTableID As Long
  Dim lngSequenceCount As Long
  
  Dim sFIELDS As String
  Dim sVALUES As String
  Dim SQLSTRING As String
  
  On Error GoTo Error_Trap
  
'******************* Create multiple child temp table ***************************
  sMCTempTable = datGeneral.UniqueSQLObjectName("ASRSysTempCustomReport", 3)

  sSQL = "SELECT * INTO [" & sMCTempTable & "] FROM [" & mstrTempTableName & "]"
  
  mclsData.ExecuteSql sSQL

  sSQL = "DELETE FROM [" & sMCTempTable & "]"
  
  mclsData.ExecuteSql sSQL
  
'************** Get the Parent SELECT SQL statment ******************************
  For iColCount = 1 To UBound(mvarColDetails, 2) Step 1
    lngTableID = mvarColDetails(14, iColCount)
    If IsReportParentTable(lngTableID) Or IsReportBaseTable(lngTableID) Then
      sParentSelectSQL = sParentSelectSQL & "[" & mvarColDetails(0, iColCount) & "]"
      sParentSelectSQL = sParentSelectSQL & ", "
    End If
  Next iColCount
  
  sParentSelectSQL = Left(sParentSelectSQL, Len(sParentSelectSQL) - 2) & " "

  sSQL = "SELECT DISTINCT " & sParentSelectSQL
  sSQL = sSQL & " FROM [" & mstrTempTableName & "] "

'Order the Parent recorset
  sSQL = sSQL & OrderBy(mlngCustomReportsBaseTable)

  Set rsParent = datGeneral.GetRecords(sSQL)

  lngColumnID = 0
  lngTableID = 0
  iChildUsed = 0
  
'**** Create the temporary recordset that is built up in the required way.     **
'  Set rsTemp = New ADODB.Recordset
'  With rsTemp
'    .CursorType = adOpenKeyset
'    .LockType = adLockOptimistic
'    .Open sMCTempTable, gADOCon, , , adCmdTable
'  End With


'*************** Circle through the distinct list of parent records *************
  With rsParent

    'TM20020802 Fault 4272
    If (.BOF And .EOF) Then
      CreateMutipleChildTempTable = False
      mstrErrorString = "No records meet selection criteria."
      mblnNoRecords = True
      
      sMCTempTable = vbNullString
'      Set rsTemp = Nothing
      Set rsParent = Nothing
      Set rsChild = Nothing
      Exit Function
    End If

    .MoveFirst
    iParentCount = 0
    lngSequenceCount = 1
 
    mbUseSequence = True
    
    Do Until .EOF
      
      If gobjProgress.Cancelled Then
        datGeneral.DropUniqueSQLObject sMCTempTable, 3
        sMCTempTable = vbNullString
        CreateMutipleChildTempTable = False
        mblnUserCancelled = True
        Exit Function
      End If
      iParentCount = iParentCount + 1
      
      ReDim avChildRecordsets(0, miUsedChildCount - 1)
      For iChildCount = 0 To UBound(mvarChildTables, 2) Step 1
        lngCurrentTableID = mvarChildTables(0, iChildCount)
        If mvarChildTables(4, iChildCount) Then   'is the child table used???
          Set rsChild = Nothing
          For iColCount = 1 To UBound(mvarColDetails, 2) Step 1
            lngTableID = mvarColDetails(14, iColCount)
            If (mvarColDetails(20, iColCount)) And (lngTableID = mvarChildTables(0, iChildCount)) _
                And (mvarColDetails(16, iColCount) <> ("?ID_" & CStr(mlngCustomReportsBaseTable))) Then
              sChildSelectSQL = sChildSelectSQL & "[" & mvarColDetails(0, iColCount) & "]"
              sChildSelectSQL = sChildSelectSQL & ", "
            End If
          Next iColCount
          sChildSelectSQL = Left(sChildSelectSQL, Len(sChildSelectSQL) - 2) & " "
          
          sChildWhereSQL = sChildWhereSQL & "[?ID_" & mvarChildTables(0, iChildCount) & "] = "
          sChildWhereSQL = sChildWhereSQL & .Fields("?ID").Value
         
          sSQL = "SELECT DISTINCT " & sChildSelectSQL
          sSQL = sSQL & " FROM [" & mstrTempTableName & "]"
          sSQL = sSQL & " WHERE " & sChildWhereSQL

'Order the child recordset.
          sSQL = sSQL & OrderBy(lngCurrentTableID)
          
          sChildSelectSQL = vbNullString
          sChildWhereSQL = vbNullString
          
          Set rsChild = datGeneral.GetRecords(sSQL)
          
          'Add the child tables recordset to the array of child tables.
          Set avChildRecordsets(0, iChildUsed) = rsChild
          iChildUsed = iChildUsed + 1
        End If
      Next iChildCount

'      With rsTemp
        iMostChilds = GetMostChildsForParent(avChildRecordsets, 0)
        If iMostChilds > 0 Then
          For i = 0 To iMostChilds - 1 Step 1
'            .AddNew
            
            sFIELDS = vbNullString
            sVALUES = vbNullString
            SQLSTRING = vbNullString
            
            '<<<<<<<<<<<<<<<<<<< Add Values To Parent Fields >>>>>>>>>>>>>>>>>>>>>>>
            For iFields = 0 To rsParent.Fields.Count - 1 Step 1
'              .Fields(rsParent.Fields(iFields).Name) = rsParent.Fields(iFields).Value
              
              sFIELDS = sFIELDS & "[" & rsParent.Fields(iFields).Name & "],"

              Select Case rsParent.Fields(iFields).Type
                Case adNumeric, adInteger, adSingle, adDouble
                'NPG20071011 Fault 12272
                  ' sVALUES = sVALUES & IIf(IsNull(rsParent.Fields(iFields).Value), 0, rsParent.Fields(iFields).Value) & ","
                  ' NPG20081201 Fault 13363
                  ' sVALUES = sVALUES & IIf(IsNull(rsParent.Fields(iFields).Value), 0, datGeneral.ConvertNumberForSQL(rsParent.Fields(iFields).Value)) & ","
                  If IsNull(rsParent.Fields(iFields).Value) Then
                    sVALUES = sVALUES & 0 & ","
                  Else
                    sVALUES = sVALUES & datGeneral.ConvertNumberForSQL(rsParent.Fields(iFields).Value) & ","
                  End If
                Case adDBTimeStamp, adDate, adDBDate, adDBTime
                  'TM20030124 Fault 4973
                  If Not IsNull(rsParent.Fields(iFields).Value) Then
                    sVALUES = sVALUES & "'" & Replace(Format(rsParent.Fields(iFields).Value, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "',"
                  Else
                    sVALUES = sVALUES & "NULL,"
                  End If
                Case adBoolean
                  sVALUES = sVALUES & IIf(rsParent.Fields(iFields).Value, 1, 0) & ","
                Case Else
                  'MH20021119 Fault 4315
                  If Not IsNull(rsParent.Fields(iFields).Value) Then
                    sVALUES = sVALUES & "'" & Replace(rsParent.Fields(iFields).Value, "'", "''") & "',"
                  Else
                    sVALUES = sVALUES & "'',"
                  End If
              End Select
              
            Next iFields

            For iChildCount = 0 To UBound(avChildRecordsets, 2) Step 1
              If Not avChildRecordsets(0, iChildCount).EOF Then
                '<<<<<<<<<<<<<<<<<<< Add Values To Child Fields >>>>>>>>>>>>>>>>>>>>>>>
                For iFields = 0 To avChildRecordsets(0, iChildCount).Fields.Count - 1 Step 1
'                  .Fields(avChildRecordsets(0, iChildCount).Fields(iFields).Name) = avChildRecordsets(0, iChildCount).Fields(iFields).Value
                  
                  sFIELDS = sFIELDS & "[" & avChildRecordsets(0, iChildCount).Fields(iFields).Name & "],"

                  Select Case avChildRecordsets(0, iChildCount).Fields(iFields).Type
                    Case adNumeric, adInteger, adSingle, adDouble
                    'NPG20071011 Fault 12272
                      ' sVALUES = sVALUES & IIf(IsNull(avChildRecordsets(0, iChildCount).Fields(iFields).Value), 0, avChildRecordsets(0, iChildCount).Fields(iFields).Value) & ","
                      ' NPG20081201 Fault 13363
                      ' sVALUES = sVALUES & IIf(IsNull(avChildRecordsets(0, iChildCount).Fields(iFields).Value), 0, datGeneral.ConvertNumberForSQL(avChildRecordsets(0, iChildCount).Fields(iFields).Value)) & ","
                      If IsNull(avChildRecordsets(0, iChildCount).Fields(iFields).Value) Then
                        sVALUES = sVALUES & 0 & ","
                      Else
                        sVALUES = sVALUES & datGeneral.ConvertNumberForSQL(avChildRecordsets(0, iChildCount).Fields(iFields).Value) & ","
                      End If
                    Case adDBTimeStamp, adDate, adDBDate, adDBTime
                      'TM20030124 Fault 4973
                      If Not IsNull(avChildRecordsets(0, iChildCount).Fields(iFields).Value) Then
                        sVALUES = sVALUES & "'" & Replace(Format(avChildRecordsets(0, iChildCount).Fields(iFields).Value, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "',"
                      Else
                        sVALUES = sVALUES & "NULL,"
                      End If
                    Case adBoolean
                      sVALUES = sVALUES & IIf(avChildRecordsets(0, iChildCount).Fields(iFields).Value, 1, 0) & ","
                    Case Else
                      'MH20021119 Fault 4315
                      If Not IsNull(avChildRecordsets(0, iChildCount).Fields(iFields).Value) Then
                        sVALUES = sVALUES & "'" & Replace(avChildRecordsets(0, iChildCount).Fields(iFields).Value, "'", "''") & "',"
                      Else
                        sVALUES = sVALUES & "'',"
                      End If
                  End Select
                  
                Next iFields
                avChildRecordsets(0, iChildCount).MoveNext
              End If
            Next iChildCount
            
            'Add the Sequence number to the sequence column for ordering the data later.
'            .Fields(lng_SEQUENCECOLUMNNAME) = lngSequenceCount
            
            sFIELDS = sFIELDS & "[" & lng_SEQUENCECOLUMNNAME & "]"
            sVALUES = sVALUES & lngSequenceCount
            
            lngSequenceCount = lngSequenceCount + 1
            
            SQLSTRING = "INSERT INTO " & sMCTempTable & " (" & sFIELDS & ") "
            SQLSTRING = SQLSTRING & " VALUES (" & sVALUES & ") "

            gADOCon.Execute SQLSTRING
            
'            .Update
          Next i
        Else
'          .AddNew
          
          sFIELDS = vbNullString
          sVALUES = vbNullString
          SQLSTRING = vbNullString

          '<<<<<<<<<<<<<<<<<<< Add Values To Parent Fields >>>>>>>>>>>>>>>>>>>>>>>
          For iFields = 0 To rsParent.Fields.Count - 1 Step 1
'            .Fields(rsParent.Fields(iFields).Name) = rsParent.Fields(iFields).Value
            
            sFIELDS = sFIELDS & "[" & rsParent.Fields(iFields).Name & "],"

            Select Case rsParent.Fields(iFields).Type
              Case adNumeric, adInteger, adSingle, adDouble
              'NPG20071011 Fault 12272
                ' sVALUES = sVALUES & IIf(IsNull(rsParent.Fields(iFields).Value), 0, rsParent.Fields(iFields).Value) & ","
                ' NPG20081201 Fault 13363
                ' sVALUES = sVALUES & IIf(IsNull(rsParent.Fields(iFields).Value), 0, datGeneral.ConvertNumberForSQL(rsParent.Fields(iFields).Value)) & ","
                If IsNull(rsParent.Fields(iFields).Value) Then
                  sVALUES = sVALUES & 0 & ","
                Else
                  sVALUES = sVALUES & datGeneral.ConvertNumberForSQL(rsParent.Fields(iFields).Value) & ","
                End If
              Case adDBTimeStamp, adDate, adDBDate, adDBTime
                'TM20030124 Fault 4973
                If Not IsNull(rsParent.Fields(iFields).Value) Then
                  sVALUES = sVALUES & "'" & Replace(Format(rsParent.Fields(iFields).Value, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "',"
                Else
                  sVALUES = sVALUES & "NULL,"
                End If
              Case adBoolean
                sVALUES = sVALUES & IIf(rsParent.Fields(iFields).Value, 1, 0) & ","
              Case Else
                'MH20020819 Fault 4315
                'sVALUES = sVALUES & "'" & Replace(rsParent.Fields(iFields).Value, "'", "''") & "',"
                If Not IsNull(rsParent.Fields(iFields).Value) Then
                  sVALUES = sVALUES & "'" & Replace(CStr(rsParent.Fields(iFields).Value), "'", "''") & "',"
                Else
                  sVALUES = sVALUES & "'',"
                End If
            End Select
            
          Next iFields
  
          'Add the Sequence number to the sequence column for ordering the data later.
'          .Fields(lng_SEQUENCECOLUMNNAME) = lngSequenceCount
          
          sFIELDS = sFIELDS & "[" & lng_SEQUENCECOLUMNNAME & "]"
          sVALUES = sVALUES & lngSequenceCount
          
          lngSequenceCount = lngSequenceCount + 1

          SQLSTRING = "INSERT INTO " & sMCTempTable & " (" & sFIELDS & ") "
          SQLSTRING = SQLSTRING & " VALUES (" & sVALUES & ") "

          gADOCon.Execute SQLSTRING

'          .Update
        End If

'      End With

      .MoveNext
      iChildUsed = 0
    Loop
  End With
  
 
'************ Re-Order the data using the defined sort orders. ******************
  sSQL = "DELETE FROM [" & mstrTempTableName & "]"

  mclsData.ExecuteSql sSQL
 
  sSQL = "INSERT INTO [" & mstrTempTableName & "] SELECT * FROM [" & sMCTempTable & "]"
  mclsData.ExecuteSql sSQL


'***************** Drop the multiple child temp table. **************************
  ' Delete the temptable if exists, and then clear the object.
'  If Len(sMCTempTable) > 0 Then
'    mclsData.ExecuteSql ("IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & sMCTempTable & "') " & _
'                      "DROP TABLE [" & sMCTempTable & "]")
'  End If
  
  datGeneral.DropUniqueSQLObject sMCTempTable, 3
  sMCTempTable = vbNullString

'************ Drop the ID columns from the temp table. ******************
'  With rsTemp
'    'Remove the ".ID" & "ID" columns from the report.
'    For iColCount = 1 To UBound(mvarColDetails, 2) Step 1
'      If (mvarColDetails(16, iColCount) = "ID") Or (mvarColDetails(16, iColCount) = ("ID_" & CStr(mlngCustomReportsBaseTable))) Then
'        sSQL = "ALTER TABLE [" & mstrTempTableName & "] DROP COLUMN [" & mvarColDetails(0, iColCount) & "]"
'        mclsData.ExecuteSql sSQL
'      End If
'    Next iColCount
'    .Close
'  End With
'  'remove the id columns from column details array.
'  ReDim Preserve mvarColDetails(20, miColumnsInReport)


'********************************************************************************

  CreateMutipleChildTempTable = True
  
TidyUpAndExit:
  sMCTempTable = vbNullString
'  Set rsTemp = Nothing
  Set rsParent = Nothing
  Set rsChild = Nothing
  
  Exit Function

Error_Trap:
  
  mstrErrorString = "Error creating temporary table for multiple childs." & vbNewLine & "(" & Err.Description & ")"
  CreateMutipleChildTempTable = False
  GoTo TidyUpAndExit
  
End Function

Private Function OrderBy(plngTableID As Long) As String

  ' This function creates an ORDER BY statement by searching
  ' through the columns defined as the reports sort order, then
  ' uses the relevant alias name

  Dim iColCount As Integer
  Dim iSortCount  As Integer
  Dim bHasOrder As Boolean
  'Dim rsSort As Recordset
  'Dim lngDefaultOrderID As Long
  
  bHasOrder = False
  
  For iSortCount = 1 To UBound(mvarSortOrder, 2)
    For iColCount = 1 To UBound(mvarColDetails, 2)

      If mvarSortOrder(1, iSortCount) = mvarColDetails(12, iColCount) _
          And mvarSortOrder(0, iSortCount) = plngTableID Then
        OrderBy = OrderBy & "[" & mvarColDetails(0, iColCount) & "] " & mvarSortOrder(2, iSortCount)
        OrderBy = OrderBy & ", "
        bHasOrder = True
        Exit For
      End If
    Next iColCount
  Next iSortCount

  If bHasOrder Then
    OrderBy = " ORDER BY " & Left(OrderBy, Len(OrderBy) - 2) & " "
  Else
    OrderBy = vbNullString
  End If
  
End Function

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property

Public Property Get CustomReportID() As Long
  CustomReportID = mlngCustomReportID
End Property

Public Property Let CustomReportID(ByVal lNewID As Long)
  mlngCustomReportID = lNewID
End Property

Public Property Get ErrorString() As String
  ErrorString = mstrErrorString
End Property

Public Property Let ErrorString(ByVal sErrorString As String)
  mstrErrorString = sErrorString
End Property

Public Property Get NoRecords() As Boolean
  NoRecords = mblnNoRecords
End Property

Private Sub Class_Initialize()
  
  ' Purpose : Sets references to other classes and redimensions arrays
  '           used for table usage information
  
  Set mclsData = New DataMgr.clsDataAccess
  Set mclsGeneral = New DataMgr.clsGeneral
  ReDim mvarSortOrder(2, 0)
  ReDim mvarColDetails(25, 0)
  ReDim mlngTableViews(2, 0)
  ReDim mstrViews(0)
  ReDim mvarGroupWith(1, 0)
  ReDim mvarPageBreak(0)
  
  ' By default this is not a Bradford Index Report
  mbIsBradfordIndexReport = False
  ' Setup this value which holds the job number.
  'mintReportJobNumberOrder = 0
End Sub

Private Sub Class_Terminate()
  
  ' Purpose : Clears references to other classes.
  
  Set mclsData = Nothing
  Set mclsGeneral = Nothing
  
End Sub

Public Function RunCustomReport(pstrPicklistFilterIDs As String) As Boolean
  ' Purpose : This function is called from frmDefsel and Batch Jobs.
  '           It is the main function which runs the report.
  On Error GoTo RunCustomReport_ERROR
  
  Dim fOK As Boolean
  
  fOK = True
  
  Screen.MousePointer = vbHourglass
  mstrPicklistFilterIDs = pstrPicklistFilterIDs
  
  If fOK Then fOK = GetCustomReportDefinition
  
  gobjEventLog.AddHeader eltCustomReport, mstrCustomReportsName
  
  If fOK Then
    With gobjProgress
      .AVI = dbText
      .MainCaption = "Custom Report"
      If Not gblnBatchMode Then
        .NumberOfBars = 1
        .Caption = "Custom Reports"
        .Time = False
        .Cancel = True
        .Bar1Caption = "Custom Report : " & mstrCustomReportsName
        .OpenProgress
      Else
        .Bar2Caption = "Custom Report : " & mstrCustomReportsName
      End If
    End With
  End If
  
  If fOK Then fOK = GetDetailsRecordsets
  If fOK Then fOK = GenerateSQL
  If fOK Then fOK = AddTempTableToSQL
  If fOK Then fOK = MergeSQLStrings
  If fOK Then fOK = UDFFunctions(mastrUDFsRequired, True)
  If fOK Then fOK = ExecuteSql
  If fOK Then fOK = UDFFunctions(mastrUDFsRequired, False)
  
  ' check that at least one child table is used.
  If fOK And miChildTablesCount > 1 And miUsedChildCount > 1 Then
    fOK = CreateMutipleChildTempTable
  End If
  
  If fOK Then fOK = CheckRecordSet
  If fOK Then fOK = PopulateGrid

  If fOK Then
    If gblnBatchMode Or gblnReportPackMode Or Not mblnPreviewOnScreen Then
      fOK = OutputReport(False)
    ElseIf mblnPreviewOnScreen Then
      Screen.MousePointer = vbDefault
      mfrmOutput.BaseTable = mstrCustomReportsBaseTableName
      mfrmOutput.Parent = Me
      mfrmOutput.HelpContextID = 1017
      mfrmOutput.Show vbModal
    End If
  End If
  
  Set mfrmOutput = Nothing
    
  Unload frmOutputOptions
  Set frmOutputOptions = Nothing
  
  Call UtilUpdateLastRun(utlCustomReport, mlngCustomReportID)
  mblnUserCancelled = _
    (mblnUserCancelled Or gobjProgress.Cancelled Or _
    (InStr(LCase(mstrErrorString), "cancelled by user") > 0))

  If mblnNoRecords Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful
    gobjEventLog.AddDetailEntry mstrErrorString
    mstrErrorString = "Completed successfully." & vbNewLine & mstrErrorString
    fOK = True
  ElseIf fOK Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful
    mstrErrorString = "Completed successfully."
  ElseIf mblnUserCancelled Then
    gobjEventLog.ChangeHeaderStatus elsCancelled
    mstrErrorString = "Cancelled by user."
  Else
    'Only details records for failures !
    gobjEventLog.AddDetailEntry mstrErrorString
    gobjEventLog.ChangeHeaderStatus elsFailed
    mstrErrorString = "Failed." & vbNewLine & vbNewLine & mstrErrorString
  End If

  mstrErrorString = "Custom Report : '" & mstrCustomReportsName & "' " & mstrErrorString

  If Not gblnBatchMode Then
    If gobjProgress.Visible Then gobjProgress.CloseProgress
    'If (fOK = False) Or (mblnNoRecords = True) Then
    If (fOK = False) Or (mblnNoRecords = True) Or Not mblnPreviewOnScreen Then
      COAMsgBox mstrErrorString, IIf(fOK, vbInformation, vbExclamation) + vbOKOnly, "Custom Reports"
    End If
  End If
  
  Screen.MousePointer = vbDefault
  
  RunCustomReport = fOK
  If fOK = True Then fOK = ClearUp Else ClearUp
  
  Exit Function
  
RunCustomReport_ERROR:

  fOK = False
  RunCustomReport = False
  mstrErrorString = "Error Whilst Running Custom Report." & vbNewLine & Err.Description
  Resume Next
  
End Function

Private Function AddTempTableToSQL() As Boolean

  ' Purpose : This function retrieves a unique temp table name and
  '           inserts it into the SQL Select statement

  On Error GoTo AddTempTableToSQL_ERROR
  
  mstrTempTableName = datGeneral.UniqueSQLObjectName("ASRSysTempCustomReport", 3)
  
  mstrSQLSelect = mstrSQLSelect & " INTO [" & mstrTempTableName & "]"
  
  AddTempTableToSQL = True
  Exit Function
  
AddTempTableToSQL_ERROR:
  
  mstrErrorString = "Error whilst retrieving unique temp table name." & vbNewLine & Err.Description
  AddTempTableToSQL = False
  
End Function

Private Function MergeSQLStrings() As Boolean
    
  ' Purpose : This function merges all the SQL string variables
  '           into one long string

  On Error GoTo MergeSQLStrings_ERROR
  
  mstrSQL = mstrSQLSelect & _
            " FROM " & _
            mstrSQLFrom & _
            IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin) & _
            IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere)

  MergeSQLStrings = True
  Exit Function
  
MergeSQLStrings_ERROR:

  mstrErrorString = "Error whilst merging SQL string components." & vbNewLine & Err.Description
  MergeSQLStrings = False
  
End Function

Private Function ExecuteSql() As Boolean

  ' Purpose : This function executes the SQL string

  On Error GoTo ExecuteSQL_ERROR

  'COAMsgBox "Recordset will be generated from : " & vbNewLine & mstrSQL
'  mclsData.ExecuteSql mstrSQL
' Execute the given SQL statement.

'  Open App.Path & "\customreportSQL.sql" For Output As #1
'
'  Print #1, mstrSQL
'
'  Close #1
  
  gADOCon.Execute mstrSQL, , adExecuteNoRecords
  
  ExecuteSql = True
  
  Exit Function
  
ExecuteSQL_ERROR:
  mstrErrorString = "Error whilst executing SQL statement." & vbNewLine & vbNewLine & Err.Description
  ExecuteSql = False
  
End Function

Private Function GetCustomReportDefinition() As Boolean
  
  ' Purpose : This function retrieves the basic definition details
  '           and stores it in module level variables
  
  On Error GoTo GetCustomReportDefinition_ERROR

  Dim rsTemp_Definition As Recordset
  Dim rsBatchOutputOptions As Recordset
  Dim strSQL As String
  Dim i As Integer
  Dim lblnReportPackMode As Boolean
  
  lblnReportPackMode = gblnReportPackMode

  strSQL = "SELECT * FROM ASRSYSCustomReportsName " & _
           "WHERE ID = " & mlngCustomReportID & " "
  
  Set rsTemp_Definition = mclsData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  
  With rsTemp_Definition
  
    If .BOF And .EOF Then
      GetCustomReportDefinition = False
      mstrErrorString = "Could not find specified Custom Report definition !"
      Exit Function
    End If
    
    mstrCustomReportsName = !Name
    mstrCustomReportsDescription = !Description
    mlngCustomReportsBaseTable = !BaseTable
    mstrCustomReportsBaseTableName = datGeneral.GetTableName(mlngCustomReportsBaseTable)
    mlngCustomReportsAllRecords = !AllRecords
    'if basetable is personnel and ReportPack Override filter exists use that swap
    If mlngCustomReportsBaseTable = glngPersonnelTableID And gblnReportPackMode Then
      mlngCustomReportsFilterID = IIf(mlngOverrideFilterID <> 0, mlngOverrideFilterID, IIf(IsNull(!Filter), 0, !Filter))
      mlngCustomReportsPickListID = IIf(mlngOverrideFilterID <> 0, 0, IIf(IsNull(!picklist), 0, !picklist))
    Else
      mlngCustomReportsFilterID = !Filter
      mlngCustomReportsPickListID = !picklist
    End If       
    mlngCustomReportsParent1Table = !parent1table
    mstrCustomReportsParent1TableName = datGeneral.GetTableName(mlngCustomReportsParent1Table)
    mlngCustomReportsParent1FilterID = !parent1filter
    mlngCustomReportsParent2Table = !parent2table
    mstrCustomReportsParent2TableName = datGeneral.GetTableName(mlngCustomReportsParent2Table)
    mlngCustomReportsParent2FilterID = !parent2filter
    mblnCustomReportsSummaryReport = !Summary
    mblnIgnoreZerosInAggregates = !IgnoreZeros
    mblnCustomReportsPrintFilterHeader = !PrintFilterHeader
    mlngCustomReportsParent1AllRecords = IIf(IsNull(!parent1AllRecords), 0, !parent1AllRecords)
    mlngCustomReportsParent1PickListID = IIf(IsNull(!parent1picklist), 0, !parent1picklist)
    mlngCustomReportsParent2AllRecords = IIf(IsNull(!parent2AllRecords), 0, !parent2AllRecords)
    mlngCustomReportsParent2PickListID = IIf(IsNull(!parent2picklist), 0, !parent2picklist)
    mbDefinitionOwner = (LCase(Trim(gsUserName)) = LCase(Trim(!UserName)))
    mblnPreviewOnScreen = IIf(lblnReportPackMode, mblnPreviewOnScreen, !OutputPreview)
    mblnOutputScreen = IIf(lblnReportPackMode, mblnOutputScreen, !OutputScreen)
    'Output Option vars of the Report Pack owning these Jobs
    mlngOriginalOutputFormat = !OutputFormat
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
    mblnOutputRetainPivotOrChart = IIf(lblnReportPackMode, mblnOutputRetainPivotOrChart, 0)
    
    mblnPreviewOnScreen = (mblnPreviewOnScreen Or (mlngOutputFormat = fmtDataOnly And mblnOutputScreen))
  End With
  
  strSQL = "SELECT C.ChildTable, C.ChildFilter, C.ChildMaxRecords, T.TableName, C.ChildOrder " & _
           "FROM ASRSYSCustomReportsChildDetails C " & _
           "      INNER JOIN ASRSysTables T " & _
           "      ON T.TableID = C.ChildTable " & _
            "WHERE C.CustomReportID = " & mlngCustomReportID & " "
  
  Set rsTemp_Definition = mclsData.OpenRecordset(strSQL, adOpenStatic, adLockReadOnly)
  
  i = 0
  With rsTemp_Definition
    If Not (.BOF And .EOF) Then
      .MoveLast
      .MoveFirst
      miChildTablesCount = .RecordCount
      .MoveFirst
      Do Until .EOF
        ReDim Preserve mvarChildTables(5, i)
        mvarChildTables(0, i) = !ChildTable         'Childs Table ID
        mvarChildTables(1, i) = !childFilter        'Childs Filter ID (if any)
        mvarChildTables(2, i) = !ChildMaxRecords    'Number of records to take from child
        mvarChildTables(3, i) = !TableName          'Child Table Name
        mvarChildTables(4, i) = False               'Boolean - True if table is used, False if not
        mvarChildTables(5, i) = !childorder         'Childs Order ID (if any)
        i = i + 1
        .MoveNext
      Loop
    End If
  End With
  
  If Not IsRecordSelectionValid Then
    GetCustomReportDefinition = False
    Exit Function
  End If

  GetCustomReportDefinition = True
  
TidyAndExit:
  
  Set rsTemp_Definition = Nothing

Exit Function

GetCustomReportDefinition_ERROR:

  GetCustomReportDefinition = False
  mstrErrorString = "Error whilst reading the Custom Report definition !" & vbNewLine & Err.Description
  Resume TidyAndExit

End Function

Private Function GetDetailsRecordsets() As Boolean

  ' Purpose : This function loads report details and sort details into
  '           arrays and leaves the details recordset reference there
  '           (dont remove it...used for summary info !)

  On Error GoTo GetDetailsRecordsets_ERROR
  
  Dim strTempSQL As String
  Dim intTemp As Integer
  Dim prstCustomReportsSortOrder As Recordset
  Dim lngTableID As Long
  
  ' Get the column information from the Details table, in order
  strTempSQL = "SELECT t.TableID, t.TableName, c.ColumnName, c.Use1000Separator, c.DataType, d.* " & _
               "FROM AsrSysCustomReportsDetails d " & _
               "LEFT JOIN asrSysColumns c ON d.ColExprID = c.ColumnID " & _
               "LEFT JOIN asrSysTables t ON c.TableID = t.TableID " & _
               "WHERE d.CustomReportID = " & mlngCustomReportID & " " & _
               "ORDER BY d.[Sequence]"
  Set mrstCustomReportsDetails = mclsData.OpenRecordset(strTempSQL, adOpenForwardOnly, adLockReadOnly)

  With mrstCustomReportsDetails
    If .BOF And .EOF Then
      GetDetailsRecordsets = False
      mstrErrorString = "No columns found in the specified Custom Report definition." & vbNewLine & "Please remove this definition and create a new one."
      Exit Function
    End If
    
    Do Until .EOF
      intTemp = UBound(mvarColDetails, 2) + 1
      ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), intTemp)
      
      '*************************************************************************
      'Now we need to decide on what the heading needs to be because QA want to
      'be able to have similar headings for hidden columns...I warned them, but
      'NO...they thought that the best move was to spend ages fixing faults in
      'v2 and put OpenHR .NET on the back-burner so that we can release a
      'Limited Edition of OpenHR called OpenHR .NET 2012 Olympic Edition.
      'What twats!!!...Fault 10211.
      
      If IIf((IsNull(!Hidden) Or (!Hidden)), True, False) Then
        mvarColDetails(0, intTemp) = "?ID_HD_" & !Type & "_" & !ColExprID
      Else
        mvarColDetails(0, intTemp) = !Heading
      End If
      
      '*************************************************************************
      
      mvarColDetails(1, intTemp) = !Size
      mvarColDetails(2, intTemp) = !dp
      mvarColDetails(3, intTemp) = !IsNumeric
      mvarColDetails(4, intTemp) = !Avge
      mvarColDetails(5, intTemp) = !cnt
      mvarColDetails(6, intTemp) = !tot
      mvarColDetails(7, intTemp) = !boc
      mvarColDetails(8, intTemp) = !poc
      mvarColDetails(9, intTemp) = !voc
      mvarColDetails(10, intTemp) = !srv
      mvarColDetails(11, intTemp) = ""
      mvarColDetails(12, intTemp) = !ColExprID
      mvarColDetails(13, intTemp) = !Type
      
      'TM20030514 Fault 5647 - need to get the table ID information for CALCS as well as columns.
      If !Type = "C" Then
        lngTableID = .Fields("TableID").Value
        mvarColDetails(14, intTemp) = lngTableID
        mvarColDetails(15, intTemp) = .Fields("TableName").Value
        mvarColDetails(16, intTemp) = .Fields("ColumnName").Value
        mvarColDetails(17, intTemp) = (.Fields("DataType").Value = sqlDate)
        mvarColDetails(18, intTemp) = (.Fields("DataType").Value = sqlBoolean)
        mvarColDetails(22, intTemp) = CBool(.Fields("Use1000Separator").Value)
        
      Else
        Dim objExpr As DataMgr.clsExprExpression
        Set objExpr = New clsExprExpression

        objExpr.ExpressionID = CLng(!ColExprID)
        objExpr.ConstructExpression
        
        'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
        ' when all we need is the expression return type (time saving measure).
        objExpr.ValidateExpression True, True
        
        lngTableID = objExpr.BaseTableID
        mvarColDetails(14, intTemp) = lngTableID
        mvarColDetails(15, intTemp) = objExpr.BaseTableName
        mvarColDetails(16, intTemp) = ""
        mvarColDetails(17, intTemp) = (objExpr.ReturnType = giEXPRVALUE_DATE)
        mvarColDetails(18, intTemp) = (objExpr.ReturnType = giEXPRVALUE_LOGIC)
        mvarColDetails(22, intTemp) = False
        Set objExpr = Nothing
     
      End If
      
      ' JDM - 09/04/01 - Give capability to hide a column on the report
      mvarColDetails(19, intTemp) = IIf((IsNull(!Hidden) Or (!Hidden)), True, False)

      mvarColDetails(20, intTemp) = IsReportChildTable(lngTableID)    'Indicates if column is a report child table.
      mvarColDetails(21, intTemp) = IIf(!Repetition = 1, True, False)
      
      'Adjust the size of the field if digit separator is used
      If mvarColDetails(22, intTemp) Then
        mvarColDetails(1, intTemp) = !Size + Int((!Size - !dp) / 3)
      End If
      
      
      'UNUSED - mvarColDetails(23, intTemp)
      mvarColDetails(24, intTemp) = IIf((IsNull(!GroupWithNextColumn) Or (Not !GroupWithNextColumn)), False, True)
      
     .MoveNext
    Loop
  .MoveFirst
  End With
  
  '******************************************************************************
  ' Add the ID columns for the tables so that we can re-select the child records
  ' when we create the multiple child temp table.
  ' NB. Is called only when there is more than one child in the report.
  '******************************************************************************
  
'  If miChildTablesCount > 1 Then
    'remember how many columns were in the report.
    miColumnsInReport = UBound(mvarColDetails, 2)
 
    intTemp = UBound(mvarColDetails, 2) + 1
    ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), intTemp)
    
    mvarColDetails(1, intTemp) = 99
    mvarColDetails(2, intTemp) = 0
    mvarColDetails(3, intTemp) = False
    mvarColDetails(4, intTemp) = False
    mvarColDetails(5, intTemp) = False
    mvarColDetails(6, intTemp) = False
    mvarColDetails(7, intTemp) = False
    mvarColDetails(8, intTemp) = False
    mvarColDetails(9, intTemp) = False
    mvarColDetails(10, intTemp) = False
    mvarColDetails(11, intTemp) = ""
    mvarColDetails(12, intTemp) = -1
    mvarColDetails(13, intTemp) = "C"
    
    mvarColDetails(14, intTemp) = mlngCustomReportsBaseTable
    mvarColDetails(15, intTemp) = datGeneral.GetTableName(CLng(mvarColDetails(14, intTemp)))
    
    mvarColDetails(0, intTemp) = "?ID"

    mvarColDetails(16, intTemp) = "ID"

    mvarColDetails(17, intTemp) = False
    mvarColDetails(18, intTemp) = False

    ' JDM - 09/04/01 - Give capability to hide a column on the report
    mvarColDetails(19, intTemp) = True

    mvarColDetails(20, intTemp) = IsReportChildTable(lngTableID)    'Indicates if column is a report child table.
  
    mvarColDetails(21, intTemp) = True
    
    mvarColDetails(24, intTemp) = False    'Group With Next Column.
    
    Dim iChildCount As Integer
    Dim lngChildTableID As Long
    If miChildTablesCount > 0 Then
      For iChildCount = 0 To UBound(mvarChildTables, 2) Step 1
        'TM20020409 Fault 3745 - only add the ID columns for tables that are actually used.
        lngChildTableID = mvarChildTables(0, iChildCount)
        If IsChildTableUsed(lngChildTableID) Then
          intTemp = UBound(mvarColDetails, 2) + 2
          ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), intTemp)
          
          mvarColDetails(1, intTemp - 1) = 99
          mvarColDetails(2, intTemp - 1) = 0
          mvarColDetails(3, intTemp - 1) = False
          mvarColDetails(4, intTemp - 1) = False
          mvarColDetails(5, intTemp - 1) = False
          mvarColDetails(6, intTemp - 1) = False
          mvarColDetails(7, intTemp - 1) = False
          mvarColDetails(8, intTemp - 1) = False
          mvarColDetails(9, intTemp - 1) = False
          mvarColDetails(10, intTemp - 1) = False
          mvarColDetails(11, intTemp - 1) = ""
          mvarColDetails(12, intTemp - 1) = -1
          mvarColDetails(13, intTemp - 1) = "C"
          mvarColDetails(14, intTemp - 1) = mvarChildTables(0, iChildCount)
          mvarColDetails(15, intTemp - 1) = mvarChildTables(3, iChildCount)
          mvarColDetails(0, intTemp - 1) = "?ID_" & mvarColDetails(14, intTemp - 1)
          mvarColDetails(16, intTemp - 1) = "ID_" & mlngCustomReportsBaseTable
          mvarColDetails(17, intTemp - 1) = False
          mvarColDetails(18, intTemp - 1) = False
          mvarColDetails(19, intTemp - 1) = True
          mvarColDetails(20, intTemp - 1) = True    'Indicates if column is a report child table.
          mvarColDetails(21, intTemp - 1) = True
          
          mvarColDetails(24, intTemp - 1) = False    'Group With Next Column.
          
          '*********************************************
          
          mvarColDetails(1, intTemp) = 99
          mvarColDetails(2, intTemp) = 0
          mvarColDetails(3, intTemp) = False
          mvarColDetails(4, intTemp) = False
          mvarColDetails(5, intTemp) = False
          mvarColDetails(6, intTemp) = False
          mvarColDetails(7, intTemp) = False
          mvarColDetails(8, intTemp) = False
          mvarColDetails(9, intTemp) = False
          mvarColDetails(10, intTemp) = False
          mvarColDetails(11, intTemp) = ""
          mvarColDetails(12, intTemp) = -1
          mvarColDetails(13, intTemp) = "C"
          mvarColDetails(14, intTemp) = mvarChildTables(0, iChildCount)
          mvarColDetails(15, intTemp) = mvarChildTables(3, iChildCount)
          mvarColDetails(0, intTemp) = "?ID_" & mvarColDetails(15, intTemp)
          mvarColDetails(16, intTemp) = "ID"
          mvarColDetails(17, intTemp) = False
          mvarColDetails(18, intTemp) = False
          mvarColDetails(19, intTemp) = True
          mvarColDetails(20, intTemp) = True     'Indicates if column is a report child table.
          mvarColDetails(21, intTemp) = True
          
          mvarColDetails(24, intTemp) = False    'Group With Next Column.
        End If
      Next iChildCount
    End If
    
    If miChildTablesCount > 1 Then
      intTemp = UBound(mvarColDetails, 2) + 1
      ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), intTemp)

      mvarColDetails(1, intTemp) = 99
      mvarColDetails(2, intTemp) = 0
      mvarColDetails(3, intTemp) = True
      mvarColDetails(4, intTemp) = False
      mvarColDetails(5, intTemp) = False
      mvarColDetails(6, intTemp) = False
      mvarColDetails(7, intTemp) = False
      mvarColDetails(8, intTemp) = False
      mvarColDetails(9, intTemp) = False
      mvarColDetails(10, intTemp) = False
      mvarColDetails(11, intTemp) = ""
      mvarColDetails(12, intTemp) = -1
      mvarColDetails(13, intTemp) = "C"
      mvarColDetails(14, intTemp) = -1
      mvarColDetails(15, intTemp) = ""
      mvarColDetails(0, intTemp) = lng_SEQUENCECOLUMNNAME
      mvarColDetails(16, intTemp) = ""
      mvarColDetails(17, intTemp) = False
      mvarColDetails(18, intTemp) = False
      mvarColDetails(19, intTemp) = True
      mvarColDetails(20, intTemp) = True     'Indicates if column is a report child table.
      mvarColDetails(21, intTemp) = True
      
      mvarColDetails(24, intTemp) = False    'Group With Next Column.
    End If
    
    intTemp = 0
  
  '******************************************************************************
  ' Get those columns defined as a SortOrder and load into array
  
  strTempSQL = "SELECT * FROM ASRSysCustomReportsDetails WHERE " & _
               "CustomReportID = " & mlngCustomReportID & " " & _
               "AND SortOrderSequence > 0 " & _
               "ORDER BY [SortOrderSequence]"
               
  Set prstCustomReportsSortOrder = datGeneral.GetReadOnlyRecords(strTempSQL)
  
  With prstCustomReportsSortOrder
    If .BOF And .EOF Then
      GetDetailsRecordsets = False
      mstrErrorString = "No columns have been defined as a sort order for the specified Custom Report definition." & vbNewLine & "Please remove this definition and create a new one."
      Exit Function
    End If
    Do Until .EOF
      intTemp = UBound(mvarSortOrder, 2) + 1
      ReDim Preserve mvarSortOrder(2, intTemp)
      mvarSortOrder(0, intTemp) = GetTableIDFromColumn(!ColExprID)
      mvarSortOrder(1, intTemp) = !ColExprID
      mvarSortOrder(2, intTemp) = !SortOrder
      .MoveNext
    Loop
  End With
  
  Set prstCustomReportsSortOrder = Nothing
  
  GetDetailsRecordsets = True
  Exit Function
  
GetDetailsRecordsets_ERROR:
  
  GetDetailsRecordsets = False
  mstrErrorString = "Error whilst retrieving the details recordsets'." & vbNewLine & Err.Description
  
End Function

Private Function GenerateSQL() As Boolean

  ' Purpose : This function calls the individual functions that
  '           general the components of the main SQL string.
  
  Dim fOK As Boolean
  
  fOK = True
  
  If fOK Then fOK = GenerateSQLSelect
  If fOK Then fOK = GenerateSQLFrom
  
  'Need to create the order string before the join as the order may require
  'another view. (for Bradford)
  If fOK Then fOK = GenerateSQLOrderBy

  If fOK Then fOK = GenerateSQLJoin
  If fOK Then fOK = GenerateSQLWhere
  
  If fOK Then GenerateSQL = True Else GenerateSQL = False

End Function


Private Function GenerateSQLSelect() As Boolean
  
  ' Purpose : This function compiles the SQLSelect string looping
  '           thru the column details recordset.
  
  On Error GoTo GenerateSQLSelect_ERROR
  
  Dim plngTempTableID As Long
  Dim pstrTempTableName As String
  Dim pstrTempColumnName As String
  
  Dim pblnOK As Boolean
  Dim pblnColumnOK As Boolean
  Dim iLoop1 As Integer
  Dim pblnNoSelect As Boolean
  Dim pblnFound As Boolean
  
  Dim pintLoop As Integer
  Dim pstrColumnList As String
  Dim pstrColumnCode As String
  Dim pstrSource As String
  Dim pintNextIndex As Integer
  
  Dim blnOK As Boolean
  Dim sCalcCode As String
  Dim alngSourceTables() As Long
  Dim objCalcExpr As clsExprExpression
  Dim objTableView As CTablePrivilege
  
  ' Set flags with their starting values
  pblnOK = True
  pblnNoSelect = False
  
  ReDim mastrUDFsRequired(0)
  
  ' JPD20030219 Fault 5067
  ' Check the user has permission to read the base table.
  pblnOK = False
  For Each objTableView In gcoTablePrivileges.Collection
    If (objTableView.TableID = mlngCustomReportsBaseTable) And _
      (objTableView.AllowSelect) Then
      pblnOK = True
      Exit For
    End If
  Next objTableView
  Set objTableView = Nothing
    
  If Not pblnOK Then
    GenerateSQLSelect = False
    mstrErrorString = "You do not have permission to read the base table" & vbNewLine & "either directly or through any views."
    Exit Function
  End If
  
  ' COWBOY ALERT !!!!!!! (Forgive me)
  ' JDM - 05/09/2005 - Fault 10302 - SQL 2005 ORDER BY clause does not work if used in conjunction with the INTO clause.
  '                                  This looks like its by design and not just a beta fault, but we can re-invetiagte when
  '                                  Micr*s*ft release the full product.
  ' NOTE : Don't try using TOP 100 PERCENT - It don't workaround the above problem!!!
  
  ' COWBOY ALERT REVISITED !!!!!!!
  ' JDM - 20/7/2012 - JIRA xxxx - Congratulations to Microsoft - you have now reintroduced this with SQL 2012. Exactly the same issue
  '                               as with SQL 2005, for some reason when using the SELECT... INTO it ignores the sort.
  '                               Run the query analyser with display execution plan to find out for yourself.
  '                               This doesn't affect 2008. I wonder what the next version of SQL will do?
  
  ' COWBOY ALERT RE-REVISTITED !!!!!!
  ' JDM - 17/09/2014 - TFS-9973 - Well, well, well, here we go again.
  
  ' COWBOY ALERT RE-RE-REVISITED !!!!!!
  ' JDM - 9/11/2016 - TFS-25065 - Commented out as hack no longer worked on large datasets that retrieved data from multiple tables
  '                               Moved the order by clause to where the data is retrieved from the temp table.
  'If glngSQLVersion = 9 Or glngSQLVersion >= 10.5 Then
  '  mstrSQLSelect = "SELECT TOP 1000000000000 "
  'Else
  mstrSQLSelect = "SELECT "
  'End If
  
  ' Dimension an array of tables/views joined to the base table/view
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view
  ' Column 2 = table/view ID
  ' (should contain everything which needs to be joined to the base tbl/view)
  ReDim mlngTableViews(2, 0)
  
  ' Loop thru the columns collection creating the SELECT and JOIN code
  For pintLoop = 1 To UBound(mvarColDetails, 2)
    
    ' Clear temp vars
    plngTempTableID = 0
    pstrTempTableName = vbNullString
    pstrTempColumnName = vbNullString
      
    ' If its a COLUMN then...
    If mvarColDetails(13, pintLoop) = "C" Then
      If mvarColDetails(0, pintLoop) <> lng_SEQUENCECOLUMNNAME Then
        ' Load the temp variables
        plngTempTableID = mvarColDetails(14, pintLoop)
        pstrTempTableName = mvarColDetails(15, pintLoop)
        pstrTempColumnName = mvarColDetails(16, pintLoop)
          
        ' Check permission on that column
        Set mobjColumnPrivileges = GetColumnPrivileges(pstrTempTableName)
        mstrRealSource = gcoTablePrivileges.Item(pstrTempTableName).RealSource
        
        If mbIsBradfordIndexReport Then
          If plngTempTableID <> mlngCustomReportsBaseTable Then
            mstrAbsenceRealSource = mstrRealSource
          End If
        End If
               
        pblnColumnOK = mobjColumnPrivileges.IsValid(pstrTempColumnName)
          
        If pblnColumnOK Then
          pblnColumnOK = mobjColumnPrivileges.Item(pstrTempColumnName).AllowSelect
        End If
        
        If pblnColumnOK Then
          
          ' JDM - 22/04/2005 - Fault 10017 - Pad out the duration field because it may not be long enough
          If mbIsBradfordIndexReport And (pintLoop = 12 Or pintLoop = 13) Then
            pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & _
            "convert(numeric(10,2)," & mstrRealSource & "." & Trim(pstrTempColumnName) & ")" & _
            " AS [" & mvarColDetails(0, pintLoop) & "]"
          Else
            pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & _
            mstrRealSource & "." & Trim(pstrTempColumnName) & _
            " AS [" & mvarColDetails(0, pintLoop) & "]"
          End If
  
          ' If the table isnt the base table (or its realsource) then
          ' Check if it has already been added to the array. If not, add it.
          If plngTempTableID <> mlngCustomReportsBaseTable Then
            pblnFound = False
            For pintNextIndex = 1 To UBound(mlngTableViews, 2)
              If mlngTableViews(1, pintNextIndex) = 0 And _
              mlngTableViews(2, pintNextIndex) = plngTempTableID Then
                pblnFound = True
                Exit For
              End If
            Next pintNextIndex
            
            If Not pblnFound Then
              pintNextIndex = UBound(mlngTableViews, 2) + 1
              ReDim Preserve mlngTableViews(2, pintNextIndex)
              mlngTableViews(1, pintNextIndex) = 0
              mlngTableViews(2, pintNextIndex) = plngTempTableID
            End If
          End If
        Else
        
          ' this column cannot be read direct. If its from a parent, try parent views
          ' Loop thru the views on the table, seeing if any have read permis for the column
          
          ReDim mstrViews(0)
          For Each mobjTableView In gcoTablePrivileges.Collection
            If (Not mobjTableView.IsTable) And _
            (mobjTableView.TableID = plngTempTableID) And _
            (mobjTableView.AllowSelect) Then
            
              pstrSource = mobjTableView.ViewName
              mstrRealSource = gcoTablePrivileges.Item(pstrSource).RealSource
              
              ' Get the column permission for the view
              Set mobjColumnPrivileges = GetColumnPrivileges(pstrSource)
              
              ' If we can see the column from this view
              If mobjColumnPrivileges.IsValid(pstrTempColumnName) Then
                If mobjColumnPrivileges.Item(pstrTempColumnName).AllowSelect Then
                  
                  ReDim Preserve mstrViews(UBound(mstrViews) + 1)
                  mstrViews(UBound(mstrViews)) = mobjTableView.ViewName
                  
                  ' Check if view has already been added to the array
                  pblnFound = False
                  For pintNextIndex = 1 To UBound(mlngTableViews, 2)
                    If mlngTableViews(1, pintNextIndex) = 1 And _
                    mlngTableViews(2, pintNextIndex) = mobjTableView.ViewID Then
                      pblnFound = True
                      Exit For
                    End If
                  Next pintNextIndex
                  
                  If Not pblnFound Then
                  
                    ' View hasnt yet been added, so add it !
                    pintNextIndex = UBound(mlngTableViews, 2) + 1
                    ReDim Preserve mlngTableViews(2, pintNextIndex)
                    mlngTableViews(1, pintNextIndex) = 1
                    mlngTableViews(2, pintNextIndex) = mobjTableView.ViewID
                    
                  End If
                End If
              End If
            End If
          
          Next mobjTableView
        
          Set mobjTableView = Nothing
          
          ' Does the user have select permission thru ANY views ?
          If UBound(mstrViews()) = 0 Then
            pblnNoSelect = True
          Else
          
            ' Add the column to the column list
            pstrColumnCode = ""
            For pintNextIndex = 1 To UBound(mstrViews)
              If pintNextIndex = 1 Then
                pstrColumnCode = "CASE"
              End If
              
              pstrColumnCode = pstrColumnCode & _
              " WHEN NOT " & mstrViews(pintNextIndex) & "." & pstrTempColumnName & " IS NULL THEN " & mstrViews(pintNextIndex) & "." & pstrTempColumnName
              
            Next pintNextIndex
            
            If Len(pstrColumnCode) > 0 Then
              pstrColumnCode = pstrColumnCode & _
              " ELSE NULL" & _
              " END AS [" & mvarColDetails(0, pintLoop) & "]"
              
              pstrColumnList = pstrColumnList & _
              IIf(Len(pstrColumnList) > 0, ",", "") & _
              pstrColumnCode
            End If
          
          End If
        
          ' If we cant see a column, then get outta here
          If pblnNoSelect Then
            GenerateSQLSelect = False
            mstrErrorString = "You do not have permission to see the column '" & mvarColDetails(16, pintLoop) & "'" & vbNewLine & "either directly or through any views."
            Exit Function
          End If
          
          If Not pblnOK Then
            GenerateSQLSelect = False
            Exit Function
          End If
          
        End If
      Else
        'Add the column which can store the sequence the records are added to the Temp table
        'when more than one Child table is selected.
         pstrColumnList = pstrColumnList & _
              IIf(Len(pstrColumnList) > 0, ",", "") & _
              0 & " AS [" & mvarColDetails(0, pintLoop) & "] "
              
      End If
    Else
        
      ' UH OH ! Its an expression rather than a column
      
      ' Get the calculation SQL, and the array of tables/views that are used to create it.
      ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
      ' Column 2 = table/view ID.
      ReDim alngSourceTables(2, 0)
      Set objCalcExpr = New clsExprExpression
      blnOK = objCalcExpr.Initialise(mlngCustomReportsBaseTable, CLng(mvarColDetails(12, pintLoop)), giEXPR_RUNTIMECALCULATION, giEXPRVALUE_UNDEFINED)
      If blnOK Then
        blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sCalcCode, True)
        
        If blnOK And gbEnableUDFFunctions Then
          blnOK = objCalcExpr.UDFCalculationCode(alngSourceTables, mastrUDFsRequired(), True)
        End If
      End If
      
      'TM20030422 Fault 5243 - The "SELECT ... INTO..." statement errors when it trys to create a column for
      'and empty string. Therefore wrap this empty sting in a CONVERT(varchar... clause if an sql empty string
      'is returned.
      'TM20030521 Fault 5702 - Compare the empty string with the calc code value converted to varchar
      sCalcCode = "CASE WHEN CONVERT(varchar," & sCalcCode & ") = '' " & _
                  "THEN CONVERT(varchar," & sCalcCode & ") " & _
                  "ELSE " & sCalcCode & " END"

      '**************************************************************************
      'TM20020730 Fault 4252
      '
      'If there are no Table/View IDs returned in the alngSourceTables array and
      'the RuntimeCalculation code returned successfully (i.e. True) then the
      'current user can see all columns required by the calc on the CALC's basetable,
      'therefore must add the CALC'S BaseTableID to the mlngTableViews array so it
      'can be added to the SQLs Join code.
      '
      'NOTE: The above only applies to the REPORT'S parent tables 1 & 2 as the
      'expression code does not return the calc's BaseTableID in the alngSourceTables
      'array.
      '**************************************************************************
      
      If mlngCustomReportsParent1Table > 0 Or mlngCustomReportsParent2Table > 0 Then
        If blnOK Then
          If objCalcExpr.BaseTableID = mlngCustomReportsParent1Table Or _
              objCalcExpr.BaseTableID = mlngCustomReportsParent2Table Then
            ' Check if table has already been added to the array
            pblnFound = False
            For pintNextIndex = 1 To UBound(mlngTableViews, 2)
              If mlngTableViews(1, pintNextIndex) = 0 And _
                mlngTableViews(2, pintNextIndex) = objCalcExpr.BaseTableID Then
                pblnFound = True
                Exit For
              End If
            Next pintNextIndex
                    
            If Not pblnFound Then
              ' View hasnt yet been added, so add it !
              pintNextIndex = UBound(mlngTableViews, 2) + 1
              ReDim Preserve mlngTableViews(2, pintNextIndex)
              mlngTableViews(1, pintNextIndex) = 0
              mlngTableViews(2, pintNextIndex) = objCalcExpr.BaseTableID
            End If
          End If
        End If
      End If
      
      '**************************************************************************
      
      Set objCalcExpr = Nothing
      
      If blnOK Then
        pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & _
              sCalcCode & " AS [" & mvarColDetails(0, pintLoop) & "]"
              'sCalcCode & " AS '" & mvarColDetails(0, pintLoop) & "'"

        ' Add the required views to the JOIN code.
         For iLoop1 = 1 To UBound(alngSourceTables, 2)
           If alngSourceTables(1, iLoop1) = 1 Then
             ' Check if view has already been added to the array
             pblnFound = False
             For pintNextIndex = 1 To UBound(mlngTableViews, 2)
               If mlngTableViews(1, pintNextIndex) = 1 And _
                 mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1) Then
                 pblnFound = True
                 Exit For
               End If
             Next pintNextIndex
                     
             If Not pblnFound Then
                     
               ' View hasnt yet been added, so add it !
               pintNextIndex = UBound(mlngTableViews, 2) + 1
               ReDim Preserve mlngTableViews(2, pintNextIndex)
               mlngTableViews(1, pintNextIndex) = 1
               mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1)
                       
             End If
'********************************************************************************
          ElseIf alngSourceTables(1, iLoop1) = 0 Then
             ' Check if table has already been added to the array
             pblnFound = False
             For pintNextIndex = 1 To UBound(mlngTableViews, 2)
               If mlngTableViews(1, pintNextIndex) = 0 And _
                 mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1) Then
                 pblnFound = True
                 Exit For
               End If
             Next pintNextIndex
            
            ' JPD20020514 Fault 3883 - Only want to check if the source table is the base table
            ' if we have NOT just found the source table in the array of joined tables.
            If Not pblnFound Then
              pblnFound = (alngSourceTables(2, iLoop1) = mlngCustomReportsBaseTable)
            End If
            
             If Not pblnFound Then
               ' table hasnt yet been added, so add it !
               pintNextIndex = UBound(mlngTableViews, 2) + 1
               ReDim Preserve mlngTableViews(2, pintNextIndex)
               mlngTableViews(1, pintNextIndex) = 0
               mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1)
             End If
'********************************************************************************
          End If
        Next iLoop1
      Else
        ' Permission denied on something in the calculation.
        mstrErrorString = "You do not have permission to use the '" & mvarColDetails(0, pintLoop) & "' calculation."
        GenerateSQLSelect = False
        Exit Function
      End If
    
    End If
      
  Next pintLoop
 
  mstrSQLSelect = mstrSQLSelect & pstrColumnList
  
  GenerateSQLSelect = True
  
  Exit Function
  
GenerateSQLSelect_ERROR:
  
  GenerateSQLSelect = False
  mstrErrorString = "Error whilst generating SQL Select statement." & vbNewLine & Err.Description
  
End Function

Private Function GenerateSQLFrom() As Boolean

  ' Purpose : It doesnt take Einstein to work out that this function
  '           adds the base table name to the from clause of the SQL string.
 
  'Dim iLoop As Integer
  Dim pobjTableView As CTablePrivilege
  
  Set pobjTableView = New CTablePrivilege
  
  mstrSQLFrom = gcoTablePrivileges.Item(mstrCustomReportsBaseTableName).RealSource
    
  Set pobjTableView = Nothing
  
  GenerateSQLFrom = True
  Exit Function
  
GenerateSQLFrom_ERROR:
  
  GenerateSQLFrom = False
  mstrErrorString = "Error in GenerateSQLFrom." & vbNewLine & Err.Description
  
End Function

Private Function GenerateSQLJoin() As Boolean

  ' Purpose : Add the join strings for parent/child/views.
  '           Also adds filter clauses to the joins if used
  
  On Error GoTo GenerateSQLJoin_ERROR

  Dim pobjTableView As CTablePrivilege
  Dim objChildTable As CTablePrivilege
  Dim pintLoop As Integer
  Dim sChildJoinCode As String
  'Dim sReuseJoinCode As String
  Dim sChildOrderString As String
  Dim rsTemp As Recordset
  Dim strFilterIDs As String
  Dim blnOK As Boolean
  Dim pblnChildUsed As Boolean
  Dim sChildJoin As String
  Dim lngTempChildID As Long
  Dim lngTempMaxRecords As Long
  Dim lngTempFilterID As Long
  Dim lngTempOrderID As Long
  Dim i As Integer
  Dim sOtherParentJoinCode As String
  Dim iLoop2 As Integer
  
  ' Get the base table real source
  mstrBaseTableRealSource = mstrSQLFrom
  
  sOtherParentJoinCode = ""
  
  ' First, do the join for all the views etc...
  
  For pintLoop = 1 To UBound(mlngTableViews, 2)
  
    ' Get the table/view object from the id stored in the array
    If mlngTableViews(1, pintLoop) = 0 Then
      Set pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
    Else
      Set pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
    End If
  
  
    ' Dont add a join here if its the child table...do that later
    'If pobjTableView.TableID <> mlngCustomReportsChildTable Then
    If Not IsReportChildTable(pobjTableView.TableID) Then
      If pobjTableView.TableID <> mlngCustomReportsParent1Table Then
        If pobjTableView.TableID <> mlngCustomReportsParent2Table Then
  
          If (pobjTableView.TableID = mlngCustomReportsBaseTable) Then
            If (pobjTableView.ViewName <> mstrBaseTableRealSource) Then
              mstrSQLJoin = mstrSQLJoin & _
              " LEFT OUTER JOIN " & pobjTableView.RealSource & _
              " ON " & mstrBaseTableRealSource & ".ID = " & pobjTableView.RealSource & ".ID"
            End If
          Else
            'JPD 20031119 Fault 7659
            ' This is a parent of a child of the report base table, not explicitly
            ' included in the report, but referred to by a child table calculation.
            For iLoop2 = 1 To UBound(mlngTableViews, 2)
              If mlngTableViews(1, iLoop2) = 0 Then
                If mclsGeneral.IsAChildOf(mlngTableViews(2, iLoop2), pobjTableView.TableID) Then
                  Set objChildTable = gcoTablePrivileges.FindTableID(mlngTableViews(2, iLoop2))
                  
                  sOtherParentJoinCode = sOtherParentJoinCode & _
                    " LEFT OUTER JOIN " & pobjTableView.RealSource & _
                    " ON " & objChildTable.RealSource & ".ID_" & CStr(pobjTableView.TableID) & " = " & pobjTableView.RealSource & ".ID"
                  Exit For
                End If
              End If
            Next iLoop2
          End If
        End If
      End If
    End If
    
    If (pobjTableView.TableID = mlngCustomReportsParent1Table) Or _
    (pobjTableView.TableID = mlngCustomReportsParent2Table) Then
      mstrSQLJoin = mstrSQLJoin & _
           " LEFT OUTER JOIN " & pobjTableView.RealSource & _
           " ON " & mstrBaseTableRealSource & ".ID_" & pobjTableView.TableID & " = " & pobjTableView.RealSource & ".ID"
    End If
  Next pintLoop
  
  'Now do the childview(s) bit, if required

  lngTempChildID = 0
  lngTempMaxRecords = 0
  lngTempFilterID = 0

'  If mlngCustomReportsChildTable > 0 Then
  If miChildTablesCount > 0 Then
    For i = 0 To UBound(mvarChildTables, 2) Step 1
      lngTempChildID = mvarChildTables(0, i)
      lngTempFilterID = mvarChildTables(1, i)
      lngTempOrderID = mvarChildTables(5, i)
      lngTempMaxRecords = mvarChildTables(2, i)
      
      pblnChildUsed = False
      
'      ' are any child fields in the report ? # 12/06/00 RH - FAULT 419
'      For pintLoop = 1 To UBound(mvarColDetails, 2)
'        If GetTableIDFromColumn(CLng(mvarColDetails(12, pintLoop))) = lngTempChildID Then
'          pblnChildUsed = True
'          Exit For
'        End If
'      Next pintLoop
     
      'TM20020409 Fault 3745 - Only do the join if columns from the table are used.
      pblnChildUsed = IsChildTableUsed(lngTempChildID)
     
      mvarChildTables(4, i) = pblnChildUsed
      If pblnChildUsed Then miUsedChildCount = miUsedChildCount + 1
      
      If pblnChildUsed = True Then
        
  '      Set objChildTable = gcoTablePrivileges.FindTableID(mlngCustomReportsChildTable)
        Set objChildTable = gcoTablePrivileges.FindTableID(lngTempChildID)
  
        If objChildTable.AllowSelect Then
          sChildJoinCode = sChildJoinCode & " LEFT OUTER JOIN " & objChildTable.RealSource & _
                           " ON " & mstrBaseTableRealSource & ".ID = " & _
                           objChildTable.RealSource & ".ID_" & mlngCustomReportsBaseTable
            
          sChildJoinCode = sChildJoinCode & " AND " & objChildTable.RealSource & ".ID IN"
          
'          sChildJoinCode = sChildJoinCode & _
'          " (SELECT TOP" & IIf(mlngCustomReportsChildMaxRecords = 0, " 100 PERCENT", " " & mlngCustomReportsChildMaxRecords) & _
'          " " & objChildTable.RealSource & ".ID FROM " & objChildTable.RealSource
        
          'TM20020328 Fault 3714 - ensure the maxrecords is >= zero.
          sChildJoinCode = sChildJoinCode & _
          " (SELECT TOP" & IIf(lngTempMaxRecords < 1, " 100 PERCENT", " " & lngTempMaxRecords) & _
          " " & objChildTable.RealSource & ".ID FROM " & objChildTable.RealSource
        
          ' Now the child order by bit - done here in case tables need to be joined.
'          Set rsTemp = datGeneral.GetOrderDefinition(datGeneral.GetDefaultOrder(mlngCustomReportsChildTable))
          If lngTempOrderID > 0 Then
            Set rsTemp = datGeneral.GetOrderDefinition(lngTempOrderID)
          Else
            Set rsTemp = datGeneral.GetOrderDefinition(datGeneral.GetDefaultOrder(lngTempChildID))
          End If
          
          sChildOrderString = DoChildOrderString(rsTemp, sChildJoin, lngTempChildID)
          Set rsTemp = Nothing
                  
          sChildJoinCode = sChildJoinCode & sChildJoin
          
          sChildJoin = vbNullString
          
          sChildJoinCode = sChildJoinCode & _
            " WHERE (" & objChildTable.RealSource & ".ID_" & mlngCustomReportsBaseTable & _
            " = " & mstrBaseTableRealSource & ".ID)"
        
          ' is the child filtered ?
          
'          If mlngCustomReportsChildFilterID > 0 Then
          If lngTempFilterID > 0 Then
'            blnOK = datGeneral.FilteredIDs(mlngCustomReportsChildFilterID, strFilterIDs)
            blnOK = datGeneral.FilteredIDs(lngTempFilterID, strFilterIDs)
      
            ' Generate any UDFs that are used in this filter
            If blnOK Then
              datGeneral.FilterUDFs lngTempFilterID, mastrUDFsRequired()
            End If
      
            If blnOK Then
              sChildJoinCode = sChildJoinCode & " AND " & _
                objChildTable.RealSource & ".ID IN (" & strFilterIDs & ")"
            Else
              ' Permission denied on something in the filter.
'              mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngCustomReportsChildFilterID) & "' filter."
              mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(lngTempFilterID) & "' filter."
              GenerateSQLJoin = False
              Exit Function
            End If
          End If
      
        End If
      
        sChildJoinCode = sChildJoinCode & IIf(Len(sChildOrderString) > 0, " ORDER BY " & sChildOrderString & ")", "")

      End If
    Next i
  End If
  
'  mstrSQLJoin = mstrSQLJoin & sChildJoinCode & IIf(Len(sChildOrderString) > 0, " ORDER BY " & sChildOrderString & ")", "")
  mstrSQLJoin = mstrSQLJoin & sChildJoinCode
  mstrSQLJoin = mstrSQLJoin & sOtherParentJoinCode
  
  GenerateSQLJoin = True
  Exit Function
  
GenerateSQLJoin_ERROR:
  
  GenerateSQLJoin = False
  mstrErrorString = "Error in GenerateSQLJoin." & vbNewLine & Err.Description
  
End Function


Private Function DoChildOrderString(rsTemp As Recordset, ByRef psJoinCode As String, plngChildID As Long) As String

  ' This function loops through the child tables default order
  ' checking if the user has privileges. If they do, add to the order string
  ' if not, leave it out.
  
  On Error GoTo DoChildOrderString_ERROR
  
  Dim fColumnOK As Boolean
  Dim fFound As Boolean
  Dim iNextIndex As Integer
  Dim sSource As String
  Dim sRealSource As String
  Dim sColumnCode As String
  Dim sCurrentTableViewName As String
  Dim objColumnPrivileges As CColumnPrivileges
  Dim pobjOrderCol As CTablePrivilege
  Dim objTableView As CTablePrivilege
  Dim alngTableViews() As Long
  Dim asViews() As String
  Dim iTempCounter As Integer
  
  ' Dimension an array of tables/views joined to the base table/view.
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
  ' Column 2 = table/view ID.
  ReDim alngTableViews(2, 0)

'  Set pobjOrderCol = gcoTablePrivileges.FindTableID(mlngCustomReportsChildTable)
  Set pobjOrderCol = gcoTablePrivileges.FindTableID(plngChildID)
  sCurrentTableViewName = pobjOrderCol.RealSource
  Set pobjOrderCol = Nothing
  
  Do Until rsTemp.EOF
    If rsTemp!Type = "O" Then
      ' Check if the user can read the column.
      Set pobjOrderCol = gcoTablePrivileges.FindTableID(rsTemp!TableID)
      Set objColumnPrivileges = GetColumnPrivileges(pobjOrderCol.TableName)
      fColumnOK = objColumnPrivileges.Item(rsTemp!ColumnName).AllowSelect
      Set objColumnPrivileges = Nothing

      If fColumnOK Then
'        If rsTemp!TableID = mlngCustomReportsChildTable Then
        If rsTemp!TableID = plngChildID Then
          DoChildOrderString = DoChildOrderString & _
            IIf(Len(DoChildOrderString) > 0, ",", "") & _
            pobjOrderCol.RealSource & "." & rsTemp!ColumnName & _
            IIf(rsTemp!Ascending, "", " DESC")
        Else
          ' If the column comes from a parent table, then add the table to the Join code.
          ' Check if the table has already been added to the join code.
          fFound = False
          iTempCounter = 0
          For iNextIndex = 1 To UBound(alngTableViews, 2)
            If alngTableViews(1, iNextIndex) = 0 And _
              alngTableViews(2, iNextIndex) = rsTemp!TableID Then
              iTempCounter = iNextIndex
              fFound = True
              Exit For
            End If
          Next iNextIndex

          If Not fFound Then
            ' The table has not yet been added to the join code, so add it to the array and the join code.
            iNextIndex = UBound(alngTableViews, 2) + 1
            ReDim Preserve alngTableViews(2, iNextIndex)
            alngTableViews(1, iNextIndex) = 0
            alngTableViews(2, iNextIndex) = rsTemp!TableID

            iTempCounter = iNextIndex
            
            psJoinCode = psJoinCode & _
              " LEFT OUTER JOIN " & pobjOrderCol.RealSource & " ASRSysTemp_" & Trim(Str(iTempCounter)) & _
              " ON " & sCurrentTableViewName & ".ID_" & Trim(Str(rsTemp!TableID)) & _
              " = ASRSysTemp_" & Trim(Str(iTempCounter)) & ".ID"
          End If
        
          DoChildOrderString = DoChildOrderString & _
            IIf(Len(DoChildOrderString) > 0, ",", "") & _
            "ASRSysTemp_" & Trim(Str(iTempCounter)) & "." & rsTemp!ColumnName & _
            IIf(rsTemp!Ascending, "", " DESC")
        End If
      Else
        ' The column cannot be read from the base table/view, or directly from a parent table.
        ' If it is a column from a prent table, then try to read it from the views on the parent table.
'        If rsTemp!TableID <> mlngCustomReportsChildTable Then
        If rsTemp!TableID <> plngChildID Then
          ' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
          ReDim asViews(0)
          For Each objTableView In gcoTablePrivileges.Collection
            If (Not objTableView.IsTable) And _
              (objTableView.TableID = rsTemp!TableID) And _
              (objTableView.AllowSelect) Then

              sSource = objTableView.ViewName
              sRealSource = gcoTablePrivileges.Item(sSource).RealSource

              ' Get the column permission for the view.
              Set objColumnPrivileges = GetColumnPrivileges(sSource)

              If objColumnPrivileges.IsValid(rsTemp!ColumnName) Then
                If objColumnPrivileges.Item(rsTemp!ColumnName).AllowSelect Then
                  ' Add the view info to an array to be put into the column list or order code below.
                  iNextIndex = UBound(asViews) + 1
                  ReDim Preserve asViews(iNextIndex)
                  asViews(iNextIndex) = objTableView.ViewName

                  ' Add the view to the Join code.
                  ' Check if the view has already been added to the join code.
                  fFound = False
                  iTempCounter = 0
                  For iNextIndex = 1 To UBound(alngTableViews, 2)
                    If alngTableViews(1, iNextIndex) = 1 And _
                      alngTableViews(2, iNextIndex) = objTableView.ViewID Then
                      fFound = True
                      iTempCounter = iNextIndex
                      Exit For
                    End If
                  Next iNextIndex

                  If Not fFound Then
                    ' The view has not yet been added to the join code, so add it to the array and the join code.
                    iNextIndex = UBound(alngTableViews, 2) + 1
                    ReDim Preserve alngTableViews(2, iNextIndex)
                    alngTableViews(1, iNextIndex) = 1
                    alngTableViews(2, iNextIndex) = objTableView.ViewID

                    iTempCounter = iNextIndex

                    psJoinCode = psJoinCode & _
                      " LEFT OUTER JOIN " & sRealSource & " ASRSysTemp_" & Trim(Str(iTempCounter)) & _
                      " ON " & sCurrentTableViewName & ".ID_" & Trim(Str(objTableView.TableID)) & _
                      " = ASRSysTemp_" & Trim(Str(iTempCounter)) & ".ID"
                  End If
                End If
              End If
              Set objColumnPrivileges = Nothing
            End If
          Next objTableView
          Set objTableView = Nothing

          ' The current user does have permission to 'read' the column through a/some view(s) on the
          ' table.
          If UBound(asViews) > 0 Then
            ' Add the column to the column list.
            sColumnCode = ""
            For iNextIndex = 1 To UBound(asViews)
              If iNextIndex = 1 Then
                sColumnCode = "CASE "
              End If

              sColumnCode = sColumnCode & _
                " WHEN NOT ASRSysTemp_" & Trim(Str(iNextIndex)) & "." & rsTemp!ColumnName & " IS NULL THEN ASRSysTemp_" & Trim(Str(iNextIndex)) & "." & rsTemp!ColumnName
            Next iNextIndex

            If Len(sColumnCode) > 0 Then
              sColumnCode = sColumnCode & _
                " ELSE NULL" & _
                " END"

              ' Add the column to the order string.
              DoChildOrderString = DoChildOrderString & _
                IIf(Len(DoChildOrderString) > 0, ", ", "") & _
                sColumnCode & _
                IIf(rsTemp!Ascending, "", " DESC")
            End If
          End If
        End If
      End If
    
      Set pobjOrderCol = Nothing
    End If
    
    rsTemp.MoveNext
  Loop

  ' JDM - Force the ID to be part of the sort order because the UDFs sort by ID too
  DoChildOrderString = DoChildOrderString & "," & sCurrentTableViewName & ".ID"

  Exit Function
  
DoChildOrderString_ERROR:
  
  Set pobjOrderCol = Nothing
  mstrErrorString = "Error while generating child order string" & vbNewLine & Err.Description
  DoChildOrderString = ""

End Function

Private Function GenerateSQLWhere() As Boolean

  ' Purpose : Generate the where clauses that cope with the joins
  '           NB Need to add the where clauses for filters/picklists etc

  On Error GoTo GenerateSQLWhere_ERROR
  
  Dim pintLoop As Integer
  Dim pobjTableView As CTablePrivilege
  Dim prstTemp As New Recordset
  Dim pstrPickListIDs As String
  Dim blnOK As Boolean
  Dim strFilterIDs As String
  'Dim objExpr As clsExprExpression
  Dim pstrParent1PickListIDs As String
  Dim pstrParent2PickListIDs As String
  
  
  If mstrPicklistFilterIDs <> "" And mstrPicklistFilterIDs <> "0" Then
    mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
      mstrSQLFrom & ".ID IN (" & mstrPicklistFilterIDs & ")"
    GenerateSQLWhere = True
    Exit Function
  End If
  
  
  Set pobjTableView = gcoTablePrivileges.FindTableID(mlngCustomReportsBaseTable)
  If pobjTableView.AllowSelect = False Then
  
  ' First put the where clauses in for the joins...only if base table is a top level table
    If UCase(Left(mstrBaseTableRealSource, 6)) <> "ASRSYS" Then
      
      For pintLoop = 1 To UBound(mlngTableViews, 2)
        ' Get the table/view object from the id stored in the array
        If mlngTableViews(1, pintLoop) = 0 Then
          Set pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
        Else
          Set pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
        End If
  
        ' dont add where clause for the base/chil/p1/p2 TABLES...only add views here
'        If (mlngTableViews(1, pintLoop) = 1) And _
'          (mlngTableViews(2, pintLoop) <> mlngCustomReportsChildTable) _
'          And (mlngTableViews(2, pintLoop) <> mlngCustomReportsParent1Table) _
'          And (mlngTableViews(2, pintLoop) <> mlngCustomReportsParent2Table) _
'          And (mlngTableViews(2, pintLoop) <> mlngCustomReportsBaseTable) Then
        ' JPD20030207 Fault 5033
        If (mlngTableViews(1, pintLoop) = 1) Then
            mstrSQLWhere = mstrSQLWhere & _
            IIf(Len(mstrSQLWhere) > 0, " OR ", " WHERE (") & _
            mstrBaseTableRealSource & ".ID IN (SELECT ID FROM " & pobjTableView.RealSource & ")"
        End If
        
      Next pintLoop
    
      If Len(mstrSQLWhere) > 0 Then mstrSQLWhere = mstrSQLWhere & ")"
    
    End If
        
  End If
    
  ' Parent 1 filter and picklist
  If mlngCustomReportsParent1PickListID > 0 Then
    pstrParent1PickListIDs = ""
    Set prstTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngCustomReportsParent1PickListID, adOpenForwardOnly, adLockReadOnly)
    
    If prstTemp.BOF And prstTemp.EOF Then
      mstrErrorString = "The first parent table picklist contains no records."
      GenerateSQLWhere = False
      Exit Function
    End If
      
    Do While Not prstTemp.EOF
      pstrParent1PickListIDs = pstrParent1PickListIDs & IIf(Len(pstrParent1PickListIDs) > 0, ", ", "") & prstTemp.Fields(0)
      prstTemp.MoveNext
    Loop
    
    prstTemp.Close
    Set prstTemp = Nothing

    mstrSQLWhere = mstrSQLWhere & _
      IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
      mstrBaseTableRealSource & ".ID_" & mlngCustomReportsParent1Table & " IN (" & pstrParent1PickListIDs & ") "
  ElseIf mlngCustomReportsParent1FilterID > 0 Then
    blnOK = True
    blnOK = datGeneral.FilteredIDs(mlngCustomReportsParent1FilterID, strFilterIDs)
    
    ' Generate any UDFs that are used in this filter
    If blnOK Then
      datGeneral.FilterUDFs mlngCustomReportsParent1FilterID, mastrUDFsRequired()
    End If
    
    ' RH 26/01/01 - Check filter is valid - unlikely not to be but just incase...
    'Set objExpr = New clsExprExpression
    'objExpr.ExpressionID = mlngCustomReportsParent1FilterID
    'If objExpr.ValidateExpression(True) <> giEXPRVALIDATION_NOERRORS Then
    '  mstrErrorString = "Error in filter definition used on the first parent table." & vbNewLine & _
    '                    "Please correct the filter definition."
    '  GenerateSQLWhere = False
    '  Exit Function
    'End If
    'Set objExpr = Nothing
      
    If blnOK Then
      mstrSQLWhere = mstrSQLWhere & _
        IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
        mstrBaseTableRealSource & ".ID_" & mlngCustomReportsParent1Table & " IN (" & strFilterIDs & ") "
    Else
      mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngCustomReportsParent1FilterID) & "' filter."
      GenerateSQLWhere = False
      Exit Function
    End If
  End If

  ' Parent 2 filter and picklist
  If mlngCustomReportsParent2PickListID > 0 Then
    pstrParent2PickListIDs = ""
    Set prstTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngCustomReportsParent2PickListID, adOpenForwardOnly, adLockReadOnly)
    
    If prstTemp.BOF And prstTemp.EOF Then
      mstrErrorString = "The second parent table picklist contains no records."
      GenerateSQLWhere = False
      Exit Function
    End If
      
    Do While Not prstTemp.EOF
      pstrParent2PickListIDs = pstrParent2PickListIDs & IIf(Len(pstrParent2PickListIDs) > 0, ", ", "") & prstTemp.Fields(0)
      prstTemp.MoveNext
    Loop
    
    prstTemp.Close
    Set prstTemp = Nothing

    mstrSQLWhere = mstrSQLWhere & _
      IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
      mstrBaseTableRealSource & ".ID_" & mlngCustomReportsParent2Table & " IN (" & pstrParent2PickListIDs & ") "
  ElseIf mlngCustomReportsParent2FilterID > 0 Then
    blnOK = True
    blnOK = datGeneral.FilteredIDs(mlngCustomReportsParent2FilterID, strFilterIDs)
    
    ' Generate any UDFs that are used in this filter
    If blnOK Then
      datGeneral.FilterUDFs mlngCustomReportsParent2FilterID, mastrUDFsRequired()
    End If
    
    ' RH 26/01/01 - Check filter is valid - unlikely not to be but just incase...
    'Set objExpr = New clsExprExpression
    'objExpr.ExpressionID = mlngCustomReportsParent2FilterID
    'If objExpr.ValidateExpression(True) <> giEXPRVALIDATION_NOERRORS Then
    '  mstrErrorString = "Error in filter definition used on the second parent table." & vbNewLine & _
    '                    "Please correct the filter definition."
    '  GenerateSQLWhere = False
    '  Exit Function
    'End If
    'Set objExpr = Nothing
    
    If blnOK Then
      mstrSQLWhere = mstrSQLWhere & _
        IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
        mstrBaseTableRealSource & ".ID_" & mlngCustomReportsParent2Table & " IN (" & strFilterIDs & ") "
    Else
      mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngCustomReportsParent2FilterID) & "' filter."
      GenerateSQLWhere = False
      Exit Function
    End If
  End If

  ' Now if we are using a picklist, add a where clause for that
  'Get List of IDs from Picklist
  If mlngCustomReportsPickListID > 0 Then
  
    Set prstTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngCustomReportsPickListID, adOpenForwardOnly, adLockReadOnly)
    
    If prstTemp.BOF And prstTemp.EOF Then
      mstrErrorString = "The selected picklist contains no records."
      GenerateSQLWhere = False
      Exit Function
    End If
      
    Do While Not prstTemp.EOF
      pstrPickListIDs = pstrPickListIDs & IIf(Len(pstrPickListIDs) > 0, ", ", "") & prstTemp.Fields(0)
      prstTemp.MoveNext
    Loop
    
    prstTemp.Close
    Set prstTemp = Nothing

    mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
    mstrSQLFrom & ".ID IN (" & pstrPickListIDs & ")"
  
  ElseIf mlngCustomReportsFilterID > 0 Then
    
    blnOK = datGeneral.FilteredIDs(mlngCustomReportsFilterID, strFilterIDs)

    ' Generate any UDFs that are used in this filter
    If blnOK Then
      datGeneral.FilterUDFs mlngCustomReportsFilterID, mastrUDFsRequired()
    End If
  
    ' RH 26/01/01 - Check filter is valid - unlikely not to be but just incase...
    'Set objExpr = New clsExprExpression
    'objExpr.ExpressionID = mlngCustomReportsFilterID
    'If objExpr.ValidateExpression(False) <> giEXPRVALIDATION_NOERRORS Then
    '  mstrErrorString = "Error in filter definition used on the base table." & vbNewLine & _
    '                    "Please correct the filter definition."
    '  GenerateSQLWhere = False
    '  Exit Function
    'End If
    'Set objExpr = Nothing
  
    If blnOK Then
      mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
        mstrSQLFrom & ".ID IN (" & strFilterIDs & ")"
    Else
      ' Permission denied on something in the filter.
      mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngCustomReportsFilterID) & "' filter."
      GenerateSQLWhere = False
      Exit Function
    End If
  End If
    
  Set prstTemp = Nothing

  GenerateSQLWhere = True
  Exit Function
  
GenerateSQLWhere_ERROR:
  
  GenerateSQLWhere = False
  mstrErrorString = "Error in GenerateSQLWhere." & vbNewLine & Err.Description
  
End Function

Private Function GenerateSQLOrderBy() As Boolean

  ' Purpose : Returns order by string from the sort order array

  Dim strOrder As String
  Dim pblnColumnOK As Boolean
  Dim pblnNoSelect As Boolean
  Dim pblnFound As Boolean
  Dim pstrSource As String
  Dim pstrOrderFrom1 As String
  Dim pstrOrderFrom2 As String
  Dim pintNextIndex As Integer

  On Error GoTo GenerateSQLOrderBy_ERROR
    
  pblnNoSelect = False
  
  ' Bradford Factor has it own sort order code
  If mbIsBradfordIndexReport Then
    '*********************************************************************************
    'TM20020605 Fault 3912 - check that the current user has permission to
    ' see and therefore order by the selected order columns on the table.
    
    'First Order Column - Check the user has select access through a table or view.
    If mlngOrderByColumnID > 0 Then
      Set mobjColumnPrivileges = GetColumnPrivileges(mstrCustomReportsBaseTableName)
      pblnColumnOK = mobjColumnPrivileges.IsValid(mstrOrderByColumn)
      If pblnColumnOK Then
        pblnColumnOK = mobjColumnPrivileges.Item(mstrOrderByColumn).AllowSelect
      End If
      
      If Not pblnColumnOK Then
        ' this column cannot be read direct. If its from a parent, try parent views
        ' Loop thru the views on the table, seeing if any have read permis for the column
        ReDim mstrViews(0)
        For Each mobjTableView In gcoTablePrivileges.Collection
          If (Not mobjTableView.IsTable) And _
          (mobjTableView.TableID = mlngCustomReportsBaseTable) And _
          (mobjTableView.AllowSelect) Then
          
            pstrSource = mobjTableView.ViewName
            
            ' Get the column permission for the view
            Set mobjColumnPrivileges = GetColumnPrivileges(pstrSource)
            
            ' If we can see the column from this view
            If mobjColumnPrivileges.IsValid(mstrOrderByColumn) Then
              If mobjColumnPrivileges.Item(mstrOrderByColumn).AllowSelect Then
                
                ReDim Preserve mstrViews(UBound(mstrViews) + 1)
                mstrViews(UBound(mstrViews)) = mobjTableView.ViewName
                
                pstrOrderFrom1 = mobjTableView.ViewName
                
                ' Check if view has already been added to the array
                pblnFound = False
                For pintNextIndex = 1 To UBound(mlngTableViews, 2)
                  If mlngTableViews(1, pintNextIndex) = 1 And _
                    mlngTableViews(2, pintNextIndex) = mobjTableView.ViewID Then
                    pblnFound = True
                    Exit For
                  End If
                Next pintNextIndex
                
                If Not pblnFound Then
                
                  ' View hasnt yet been added, so add it !
                  pintNextIndex = UBound(mlngTableViews, 2) + 1
                  ReDim Preserve mlngTableViews(2, pintNextIndex)
                  mlngTableViews(1, pintNextIndex) = 1
                  mlngTableViews(2, pintNextIndex) = mobjTableView.ViewID
                  Exit For
                End If
              End If
            End If
          End If
        
        Next mobjTableView
      
        Set mobjTableView = Nothing

        ' Does the user have select permission thru ANY views ?
        If UBound(mstrViews()) = 0 Then
          pblnNoSelect = True
        End If
      
      Else
        pstrOrderFrom1 = mstrCustomReportsBaseTableName
      End If
      
      If pblnNoSelect Then
        GenerateSQLOrderBy = False
        mstrErrorString = "You do not have permission to see the column '" & mstrOrderByColumn & "' " & vbNewLine & "either directly or through any views."
        Exit Function
      End If
    End If
    
    'Second Order Column - Check the user has select access through a table or view.
    If mlngGroupByColumnID > 0 Then
      pblnNoSelect = False
      Set mobjColumnPrivileges = GetColumnPrivileges(mstrCustomReportsBaseTableName)
      pblnColumnOK = mobjColumnPrivileges.IsValid(mstrGroupByColumn)
      If pblnColumnOK Then
        pblnColumnOK = mobjColumnPrivileges.Item(mstrGroupByColumn).AllowSelect
      End If
      
      If Not pblnColumnOK Then
        ' this column cannot be read direct. If its from a parent, try parent views
        ' Loop thru the views on the table, seeing if any have read permis for the column
        ReDim mstrViews(0)
        For Each mobjTableView In gcoTablePrivileges.Collection
          If (Not mobjTableView.IsTable) And _
          (mobjTableView.TableID = mlngCustomReportsBaseTable) And _
          (mobjTableView.AllowSelect) Then
          
            pstrSource = mobjTableView.ViewName
            
            ' Get the column permission for the view
            Set mobjColumnPrivileges = GetColumnPrivileges(pstrSource)
            
            ' If we can see the column from this view
            If mobjColumnPrivileges.IsValid(mstrOrderByColumn) Then
              If mobjColumnPrivileges.Item(mstrOrderByColumn).AllowSelect Then
                
                ReDim Preserve mstrViews(UBound(mstrViews) + 1)
                mstrViews(UBound(mstrViews)) = mobjTableView.ViewName
                
                pstrOrderFrom2 = mobjTableView.ViewName
                
                ' Check if view has already been added to the array
                pblnFound = False
                For pintNextIndex = 1 To UBound(mlngTableViews, 2)
                  If mlngTableViews(1, pintNextIndex) = 1 And _
                  mlngTableViews(2, pintNextIndex) = mobjTableView.ViewID Then
                    pblnFound = True
                    Exit For
                  End If
                Next pintNextIndex
                
                If Not pblnFound Then
                
                  ' View hasnt yet been added, so add it !
                  pintNextIndex = UBound(mlngTableViews, 2) + 1
                  ReDim Preserve mlngTableViews(2, pintNextIndex)
                  mlngTableViews(1, pintNextIndex) = 1
                  mlngTableViews(2, pintNextIndex) = mobjTableView.ViewID
                  Exit For
                End If
              End If
            End If
          End If
        
        Next mobjTableView
      
        Set mobjTableView = Nothing

        ' Does the user have select permission thru ANY views ?
        If UBound(mstrViews()) = 0 Then
          pblnNoSelect = True
        End If
      
      Else
        pstrOrderFrom2 = mstrCustomReportsBaseTableName
      End If
    
      If pblnNoSelect Then
        GenerateSQLOrderBy = False
        mstrErrorString = "You do not have permission to see the column '" & mstrGroupByColumn & "' " & vbNewLine & "either directly or through any views."
        Exit Function
      End If
    End If
      '*********************************************************************************
'TM24032004
'      'MH20020521 Fault 3820
'      'strOrder = "[" & mstrOrderByColumn & "] " & IIf(mbOrderBy1Asc = True, "Asc", "Desc")
'      'If Not mstrGroupByColumn = "<None>" And Not mstrGroupByColumn = mstrOrderByColumn Then
'      '  strOrder = strOrder & ",[" & mstrGroupByColumn & "] " & IIf(mbOrderBy2Asc = True, "Asc", "Desc")
'      'End If
'      'mstrSQLOrderBy = " ORDER BY " & strOrder & ", [Personnel_ID] Asc, [Start_Date] Asc"
'      strOrder = "[Order_1] " & IIf(mbOrderBy1Asc = True, "Asc", "Desc")
'      If Not mstrGroupByColumn = "<None>" And Not mstrGroupByColumn = mstrOrderByColumn Then
'        strOrder = strOrder & ", [Order_2] " & IIf(mbOrderBy2Asc = True, "Asc", "Desc")
'      End If
'      mstrSQLOrderBy = " ORDER BY " & strOrder & ", [Personnel_ID] Asc"
''      If InStr(strOrder, "[Start_Date]") = 0 Then
''        mstrSQLOrderBy = mstrSQLOrderBy & ", [Start_Date] Asc"
''      End If
    If mlngOrderByColumnID > 0 Then
      strOrder = "[Order_1] " & IIf(mbOrderBy1Asc = True, "Asc", "Desc")
    End If
    If mlngGroupByColumnID > 0 And (mlngOrderByColumnID <> mlngGroupByColumnID) Then
      If mlngOrderByColumnID > 0 Then
        strOrder = strOrder & ", "
        strOrder = strOrder & "[Order_2] " & IIf(mbOrderBy2Asc = True, "Asc", "Desc")
      Else
        strOrder = strOrder & "[Order_1] " & IIf(mbOrderBy2Asc = True, "Asc", "Desc")
      End If
    End If
    If (mlngOrderByColumnID = 0) And (mlngGroupByColumnID = 0) Then
      mstrSQLOrderBy = " ORDER BY [Personnel_ID] Asc"
    Else
      mstrSQLOrderBy = " ORDER BY " & strOrder & ", [Personnel_ID] Asc"
    End If
    
  Else
  
    If UBound(mvarSortOrder, 2) > 0 Then
      ' Columns have been defined, so use these for the base table/view
      mstrSQLOrderBy = DoDefinedOrderBy
    End If
  
    If Len(mstrSQLOrderBy) > 0 Then mstrSQLOrderBy = " ORDER BY " & mstrSQLOrderBy
  End If
   
  GenerateSQLOrderBy = True
  Exit Function
  
GenerateSQLOrderBy_ERROR:
  
  GenerateSQLOrderBy = False
  mstrErrorString = "Error in GenerateSQLOrderBy." & vbNewLine & Err.Description
  
End Function

Private Function DoDefinedOrderBy() As String

  ' This function creates the base ORDER BY statement by searching
  ' through the columns defined as the reports sort order, then
  ' uses the relevant alias name
  
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  
  For iLoop = 1 To UBound(mvarSortOrder, 2)
      
    For iLoop2 = 1 To UBound(mvarColDetails, 2)
      
      If mvarSortOrder(1, iLoop) = mvarColDetails(12, iLoop2) Then
        
        DoDefinedOrderBy = DoDefinedOrderBy & _
        IIf(Len(DoDefinedOrderBy) > 0, ",", "") & _
        "[" & mvarColDetails(0, iLoop2) & "] " & _
        mvarSortOrder(2, iLoop)
      
        Exit For
        
      End If
    
    Next iLoop2
    
  Next iLoop
 
End Function

Public Function GetTableIDFromColumn(lngColumnID As Long) As Long
  
  ' Purpose : To return the table id for which the given column belongs
  
  Dim rsInfo As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT ASRSysTables.TableID " & _
           "FROM ASRSysColumns JOIN ASRSysTables " & _
           "ON (ASRSysTables.TableID = ASRSysColumns.TableID) " & _
           "WHERE ColumnID = " & CStr(lngColumnID)

  Set rsInfo = mclsData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
          
  If rsInfo.BOF And rsInfo.EOF Then
    GetTableIDFromColumn = 0
  Else
    GetTableIDFromColumn = rsInfo!TableID
  End If
  
  Set rsInfo = Nothing

End Function

Private Function CheckRecordSet() As Boolean

  ' Purpose : To get recordset from temptable and show recordcount
  
  Dim sSQL As String
  
  On Error GoTo CheckRecordSet_ERROR
  
'  Set mrstCustomReportsOutput = mclsData.OpenRecordset("SELECT * FROM " & mstrTempTableName, adOpenStatic, adLockReadOnly)
  
  
  ' Order the entire recordset.
  'TM20020409 Fault 3746 - Order the code as it is iserted back into the original temp table.
'  sSQL = sSQL & " ORDER BY [" & lng_SEQUENCECOLUMNNAME & "] ASC"

  
  'TM20020429 Fault 3764
  If mbUseSequence Then
    sSQL = "SELECT * FROM [" & mstrTempTableName & "]"
    sSQL = sSQL & " ORDER BY [" & lng_SEQUENCECOLUMNNAME & "] ASC"
  Else
    sSQL = "SELECT * FROM " & mstrTempTableName & " " & mstrSQLOrderBy
  End If
  
  Set mrstCustomReportsOutput = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If mrstCustomReportsOutput.BOF And mrstCustomReportsOutput.EOF Then
    CheckRecordSet = False
    mstrErrorString = "No records meet selection criteria."
    mblnNoRecords = True
    Exit Function
  End If
  
  CheckRecordSet = True
  Exit Function
  
CheckRecordSet_ERROR:
  
  mstrErrorString = "Error while checking returned recordset." & vbNewLine & "(" & Err.Description & ")"
  CheckRecordSet = False
  
End Function

Private Function PopulateGrid() As Boolean

  On Error GoTo PopulateGrid_ERROR
  
  Dim fOK As Boolean
  
  fOK = True
  fOK = PopulateGrid_InitForm
  ' Turn the grid redraw off to speed it up a little (shaves 4 mins off 80k records)
  mfrmOutput.grdOutput.Redraw = False

  If fOK Then
    fOK = PopulateGrid_LoadRecords
  End If
  
  'JDM - 09/04/2001 - Loop through and hide all the hidden columns
  If fOK Then
    fOK = PopulateGrid_HideColumns
  End If
  
  ' Turn the grid redraw back on
  mfrmOutput.grdOutput.Redraw = True
  
  ' Set the tag of grid to indicate page breaks or not
  mfrmOutput.grdOutput.Tag = mblnReportHasPageBreak
  
  'MH20010307
  mfrmOutput.HasSummaryInfo = mblnReportHasSummaryInfo
  
  PopulateGrid = fOK
  
  Exit Function
   
PopulateGrid_ERROR:
  
  PopulateGrid = False
  mstrErrorString = "Error populating grid." & vbNewLine & "(" & Err.Description & ")"

End Function

Private Function PopulateGrid_InitForm() As Boolean

  ' Purpose : To prepare the output form, load grid with columns etc)
  ' Input   : None
  ' Output  : None

  On Error GoTo InitForm_ERROR
  
  ReDim mlngColWidth(0)
  Dim iLoop As Integer
  Dim intColCounter As Integer
  
  Set mfrmOutput = New frmCustomReportsPreview
  
  With mfrmOutput

    .Tag = mstrCustomReportsName
    
    ' JPD20011109 Fault 2969
    If mblnCustomReportsPrintFilterHeader Then
      If (mlngCustomReportsFilterID > 0) Then
        .Tag = .Tag & " (Base Table filter : " & datGeneral.GetFilterName(mlngCustomReportsFilterID) & ")"
      ElseIf (mlngCustomReportsPickListID > 0) Then
        .Tag = .Tag & " (Base Table picklist : " & datGeneral.GetPicklistName(mlngCustomReportsPickListID) & ")"
      End If
    End If
    
    ' JPD20011109 Fault 2968
    ' Put a title bar on the grid.
    .grdOutput.Caption = Left(Replace(Replace(.Tag, "_", " "), "&", "&&"), 250)
    
    .grdOutput.Columns.RemoveAll
    
    ' Add first column (for page break info only)
    intColCounter = 0
    .grdOutput.Columns.Add intColCounter
    .grdOutput.Columns(intColCounter).Name = "PageBreak"
    .grdOutput.Columns(intColCounter).Caption = "PageBreak"
    ' store the length of the column heading in order to set the column width in PopulateGrid_LoadColumns
    ReDim Preserve mlngColWidth(intColCounter)
    'mlngColWidth(intColCounter) = mfrmOutput.TextWidth(.grdOutput.Columns(intColCounter).Caption) + 200
    mlngColWidth(intColCounter) = mfrmOutput.TextWidth(.grdOutput.Columns(intColCounter).Caption)

    ' Add first column (for summary info headings)
    intColCounter = intColCounter + 1
    .grdOutput.Columns.Add intColCounter
    .grdOutput.Columns(intColCounter).Name = "AsrSysSummary"
    .grdOutput.Columns(intColCounter).Caption = "Summary Info" '"AsrSysSummary"
    ' store the length of the column heading in order to set the column width in PopulateGrid_LoadColumns
    ReDim Preserve mlngColWidth(intColCounter)
    'mlngColWidth(intColCounter) = mfrmOutput.TextWidth(.grdOutput.Columns(intColCounter).Caption) + 200
    mlngColWidth(intColCounter) = mfrmOutput.TextWidth(.grdOutput.Columns(intColCounter).Caption)
    
    ' Add a column to the grid for each column in the temp table
    For iLoop = 0 To (mrstCustomReportsOutput.Fields.Count - 1)
      
      intColCounter = intColCounter + 1
      
      .grdOutput.Columns.Add intColCounter
      .grdOutput.Columns(intColCounter).Name = mrstCustomReportsOutput.Fields(iLoop).Name
      .grdOutput.Columns(intColCounter).Caption = Replace(mrstCustomReportsOutput.Fields(iLoop).Name, "_", " ")

      If (mrstCustomReportsOutput.Fields(iLoop).Type = adInteger) Or _
         (mrstCustomReportsOutput.Fields(iLoop).Type = adNumeric) Or _
         (mrstCustomReportsOutput.Fields(iLoop).Type = adDouble) Or _
         (mrstCustomReportsOutput.Fields(iLoop).Type = adSingle) Then
        
        .grdOutput.Columns(intColCounter).Alignment = ssCaptionAlignmentRight
        
      End If
      
      If mrstCustomReportsOutput.Fields(iLoop).Type = adBoolean Then
        .grdOutput.Columns(intColCounter).Alignment = ssCaptionAlignmentCenter
      End If
             
      ' store the length of the column heading in order to set the column width later on
      ReDim Preserve mlngColWidth(intColCounter)
      mlngColWidth(intColCounter) = mfrmOutput.TextWidth(mrstCustomReportsOutput.Fields(iLoop).Name) + 350

    Next iLoop
  
  End With
  
  PopulateGrid_InitForm = True
  Exit Function
  
InitForm_ERROR:

  PopulateGrid_InitForm = False
  mstrErrorString = "Error whilst initialising output form." & vbNewLine & "(" & Err.Description & ")"

End Function

Private Function PopulateGrid_AddToGroupWith(pstrValue As String, _
                                             pintRowIndex As Integer, _
                                             pintGridColIndex As Integer) As Boolean

  Dim intCount As Integer
  Dim strAddString As String
  Dim blnNewGroup As Boolean
  
  blnNewGroup = False
  
  If pintRowIndex > UBound(mvarGroupWith, 2) Then
    ReDim Preserve mvarGroupWith(1, pintRowIndex)
    mvarGroupWith(1, pintRowIndex) = 0
    strAddString = vbNullString
  ElseIf UBound(mvarGroupWith, 2) = 0 Then
    blnNewGroup = True
    strAddString = mvarGroupWith(0, pintRowIndex)
  Else
    blnNewGroup = True
    strAddString = mvarGroupWith(0, pintRowIndex)
  End If
  
  If blnNewGroup Then
    strAddString = strAddString & String(pintGridColIndex - mvarGroupWith(1, pintRowIndex), vbTab) & pstrValue
  Else
    strAddString = strAddString & String(pintGridColIndex, vbTab) & pstrValue
  End If
  
  mvarGroupWith(0, pintRowIndex) = strAddString
  mvarGroupWith(1, pintRowIndex) = pintGridColIndex
  
  'If len of data is greater than the previous length of data, store len in the array.
  If BigTextWidth(pstrValue, 0) > mlngColWidth(pintGridColIndex) Then
    mlngColWidth(pintGridColIndex) = BigTextWidth(pstrValue, 0)
  End If

End Function

Private Function PopulateGrid_HideColumns() As Boolean

  ' Purpose : This function hides any columns we don't want the user to see.
  Dim iCount As Integer
  Dim intColCounter As Integer
  
  On Error GoTo HideColumns_ERROR

  ' If report contains no summary info, hide the column
  intColCounter = 0
  ' Hide the pagebreak column regardless
  mfrmOutput.grdOutput.Columns(intColCounter).Visible = False
'  mfrmOutput.grdOutput.Columns(intColCounter).Visible = True
  
  intColCounter = intColCounter + 1
  If Not mblnReportHasSummaryInfo Then
    mfrmOutput.grdOutput.Columns(intColCounter).Visible = False
'    mfrmOutput.grdOutput.Columns(intColCounter).Visible = True
  Else
    mfrmOutput.grdOutput.Columns(intColCounter).Visible = True
  End If

  For iCount = 1 To UBound(mvarColDetails, 2)
    intColCounter = intColCounter + 1
    
    If (mvarColDetails(24, iCount - 1) = True) Then
      mfrmOutput.grdOutput.Columns(intColCounter).Visible = False
'      mfrmOutput.grdOutput.Columns(intColCounter).Visible = True
    Else
      mfrmOutput.grdOutput.Columns(intColCounter).Visible = (Not mvarColDetails(19, iCount))
'      mfrmOutput.grdOutput.Columns(intColCounter).Visible = True
    End If
  Next iCount

  PopulateGrid_HideColumns = True
  Exit Function

HideColumns_ERROR:

  PopulateGrid_HideColumns = False
  If Not gblnBatchMode Then gobjProgress.CloseProgress
  mstrErrorString = "Error in PopulateGrid_HideColumns." & vbNewLine & Err.Description

End Function



Private Function PopulateGrid_FormatData(sfieldname As String, vData As Variant _
                                        , mbSuppressRepeated As Boolean, pbNewBaseRecord As Boolean _
                                        ) As Variant

  ' Purpose : Format the data to the form the user has specified to see it
  '           in the grid
  ' Input   : None
  ' Output  : True/False
  ' Notes   : Fault 2514 - Fault 2514 - Make sure SRV works (added suppress parameter)

  Dim pintLoop As Integer
  Dim strTemp As String
  Dim iDigitCount As Integer
  Dim iCount As Integer
  Dim vOriginalData As Variant
  
  vOriginalData = vData
  
  For pintLoop = 1 To UBound(mvarColDetails, 2)
    If mvarColDetails(0, pintLoop) = sfieldname Then

      'Do the DP thing
      If mvarColDetails(3, pintLoop) Then 'Is Numeric
        If mvarColDetails(2, pintLoop) <> 0 Then
          vData = Format(vData, "0." & String(mvarColDetails(2, pintLoop), "0"))
        Else

          If mvarColDetails(1, pintLoop) > 0 Then   'Size restriction
            'RH 02/02/01 - Bug 1706 - If we are displaying with zero decimal places
            '                         and the data is 0, format as 0 not #
            If vData = "0" Then
              vData = Format(vData, "0")
            Else
              vData = Format(vData, "#")
            End If
          End If

        End If
      End If

      'Is it a boolean calculation ? If so, change to Y or N
      If mvarColDetails(18, pintLoop) Then
        If (vData = "True") Or (vData = 1) Then vData = "Y"
        If (vData = "False") Or (vData = 0) Then vData = "N"
      End If

      'If its a date column, format it as dateformat
      If mvarColDetails(17, pintLoop) Then
        vData = Format(vData, DateFormat)
      End If

      'Numeric digit separators
      If mvarColDetails(22, pintLoop) Then
        If mvarColDetails(2, pintLoop) <> 0 Then
          vData = Format(vData, "#,0." & String(mvarColDetails(2, pintLoop), "0"))
        Else
          vData = Format(vData, "#,0")
        End If
      End If

      'Check if has decimal places
      If mvarColDetails(1, pintLoop) > 0 Then   'Size restriction
        If mvarColDetails(2, pintLoop) > 0 Then
          If InStr(vData, ".") > mvarColDetails(1, pintLoop) Then
            vData = Left(vData, mvarColDetails(1, pintLoop)) & _
                    Mid(vData, InStr(vData, "."))
          End If

        Else
          If Len(vData) > mvarColDetails(1, pintLoop) Then
            vData = Left(vData, mvarColDetails(1, pintLoop))
          End If

        End If
      End If

      ' SRV ?
      If Not mbIsBradfordIndexReport Then
        If mbSuppressRepeated = True Then
          'check if column value should be repeated or not.
          If Not mvarColDetails(21, pintLoop) And Not pbNewBaseRecord _
              And Not mvarColDetails(10, pintLoop) And Not mvarColDetails(20, pintLoop) Then
            'TM20020430 Fault 3812 - compare string to string
            'TM20020503 Fault 3833 - check for nulls before comparison is made.
            If CStr(RTrim(IIf(IsNull(mvarColDetails(11, pintLoop)), vbNullString, mvarColDetails(11, pintLoop)))) = CStr(RTrim(IIf(IsNull(vOriginalData), vbNullString, vOriginalData))) Then
              vData = ""
            End If
          ElseIf mvarColDetails(10, pintLoop) Then
            'TM20020430 Fault 3812 - compare string to string
            'TM20020503 Fault 3833 - check for nulls before comparison is made.
            If CStr(RTrim(IIf(IsNull(mvarColDetails(11, pintLoop)), vbNullString, mvarColDetails(11, pintLoop)))) = CStr(RTrim(IIf(IsNull(vOriginalData), vbNullString, vOriginalData))) Then
              vData = ""
            End If
          End If
        End If
        Exit For

      Else
        'Bradford Factor does not use the repetition functionality.
        If mbSuppressRepeated = True Then
          If mvarColDetails(10, pintLoop) Then  'Suppress Repeated Values
            If CStr(RTrim(IIf(IsNull(mvarColDetails(11, pintLoop)), vbNullString, mvarColDetails(11, pintLoop)))) = CStr(RTrim(IIf(IsNull(vOriginalData), vbNullString, vOriginalData))) Then
              vData = ""
            End If
          End If
        End If
      End If
    End If

  Next pintLoop

  'JPD20010807 Fault 2215 - Moved Roy's line around so that they actually
  ' alter the data returned.
  ' RH 09/10/00 - BUG - vbNewLine messess up exporting to word/excel, so
  '               remove the vbNewLine from the text
  If Not IsNull(vData) Then vData = Replace(vData, Chr(0), "")
  If Not IsNull(vData) Then vData = Replace(vData, vbNewLine, " ")
  ' lose tabs too...
  If Not IsNull(vData) Then vData = Replace(vData, vbTab, " ")

  PopulateGrid_FormatData = vData
  
End Function



Private Function DoesHaveGrandSummary() As Boolean

  On Error Resume Next
  
  Dim pintLoop As Integer
  
  ' Does it have any grand totals to produce ?
  mrstCustomReportsDetails.MoveFirst
  
  For pintLoop = 1 To mrstCustomReportsDetails.RecordCount
    
    If mrstCustomReportsDetails!Avge Or _
       mrstCustomReportsDetails!cnt Or _
       mrstCustomReportsDetails!tot Then
          DoesHaveGrandSummary = True
          Exit Function
    End If
    
    mrstCustomReportsDetails.MoveNext
  
  Next pintLoop
  
  DoesHaveGrandSummary = False
  
End Function

Private Function PopulateGrid_DoSummaryInfo(pavColumns As Variant, _
  piColumnIndex As Integer, _
  piSortIndex As Integer) As Boolean

  Dim fDoValue As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iColumnIndex As Integer
  Dim sSQL As String
  Dim rsTemp As Recordset
  Dim fHasAverage As Boolean
  Dim fHasCount As Boolean
  Dim fHasTotal As Boolean
  Dim sWhereCode As String
  Dim sFromCode As String
  Dim sCountAddString As String
  Dim sAverageAddString As String
  Dim sTotalAddString As String
  Dim iLogicValue As Integer
  
  Dim miAmountOfRecords As Single
  Dim miAmountofAbsence As Single
  Dim sBradfordSummary As String
  Dim asBradfordSummaryLine() As String

  Dim intColCounter As Integer
  
  Dim strAggrValue As String
  
  intColCounter = 1
  strAggrValue = vbNullString
  
  ' Construct the summary where clause.
  sWhereCode = ""
  For iLoop = 1 To piSortIndex
    iColumnIndex = 0
    For iLoop2 = 1 To UBound(mvarColDetails, 2)
      If (mvarColDetails(12, iLoop2) = mvarSortOrder(1, iLoop)) And _
        (mvarColDetails(13, iLoop2) = "C") Then
              
        iColumnIndex = iLoop2
        Exit For
      End If
    Next iLoop2
      
    If iColumnIndex > 0 Then
      If mvarColDetails(7, iColumnIndex) Or mvarColDetails(8, iColumnIndex) Then
        ' The column is a break/page on change column so put it in the Where clause.
        sWhereCode = sWhereCode & _
        IIf(Len(sWhereCode) = 0, " WHERE ", " AND ")
        
        If (Not mvarColDetails(3, iColumnIndex)) And _
          (Not mvarColDetails(17, iColumnIndex)) And _
          (Not mvarColDetails(18, iColumnIndex)) Then
          ' Character column. Treat empty strings along with nulls.
          If Len(pavColumns(3, iColumnIndex)) = 0 Then
            sWhereCode = sWhereCode & "(([" & CStr(mvarColDetails(0, iColumnIndex)) & "] = '') OR ([" & CStr(mvarColDetails(0, iColumnIndex)) & "] IS NULL))"
          Else
            sWhereCode = sWhereCode & "([" & CStr(mvarColDetails(0, iColumnIndex)) & "] = '" & Replace(pavColumns(3, iColumnIndex), "'", "''") & "')"
          End If
        Else
          If IsNull(pavColumns(3, iColumnIndex)) Or pavColumns(3, iColumnIndex) = "" Then
            sWhereCode = sWhereCode & "([" & CStr(mvarColDetails(0, iColumnIndex)) & "] IS NULL)"
          Else
            If mvarColDetails(17, iColumnIndex) Then
              ' Date column.
              sWhereCode = sWhereCode & "([" & CStr(mvarColDetails(0, iColumnIndex)) & "] = '" & Replace(Format(pavColumns(3, iColumnIndex), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')"
            Else
              If mvarColDetails(18, iColumnIndex) Then
                ' Logic Column.
                'TM20020523 Fault 3910 - if logic column then convert the stored 'Y' or 'N' to 1 or 0.
                iLogicValue = IIf(pavColumns(3, iColumnIndex) = "Y", 1, 0)
                sWhereCode = sWhereCode & "([" & CStr(mvarColDetails(0, iColumnIndex)) & "] = " & iLogicValue & ")"
              Else
                ' Numeric column.
                'NPG20071205 Fault 12552
                sWhereCode = sWhereCode & "([" & CStr(mvarColDetails(0, iColumnIndex)) & "] = " & datGeneral.ConvertNumberForSQL(pavColumns(3, iColumnIndex)) & ")"
              End If
            End If
          End If
        End If
        
        
      End If
    End If
  Next iLoop
 
  ' Construct the required select statement.
  sSQL = ""
  sFromCode = ""
  For iLoop = 1 To UBound(mvarColDetails, 2)
    
    If mvarColDetails(4, iLoop) Then
      ' Average.
      mblnReportHasSummaryInfo = True
      'TM20020718 Fault 4169 - indicate in the hidden column that the row is an average row.
      sAverageAddString = "*average*" & vbTab & "Sub Average" & vbTab
      
      If Not mbIsBradfordIndexReport Then
        If mvarColDetails(20, iLoop) Then
          ' JPD20020712 Fault 4155
          sSQL = sSQL & ",(SELECT AVG(convert(float,[" & mvarColDetails(0, iLoop) & "])) " & _
                        "FROM (SELECT DISTINCT [?ID_" & mvarColDetails(15, iLoop) & "], [" & mvarColDetails(0, iLoop) & "] " & _
                                "FROM " & mstrTempTableName & " " & _
                                " " & sWhereCode & " "
                              
         If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
           sSQL = sSQL & " AND ([" & mvarColDetails(0, iLoop) & "] <> 0) "
         End If
        
         sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'avg_" & Trim(Str(iLoop)) & "'"
        Else
          ' JPD20020712 Fault 4155
          sSQL = sSQL & ",(SELECT AVG(convert(float, [" & mvarColDetails(0, iLoop) & "])) " & _
                        "FROM (SELECT DISTINCT [?ID], [" & mvarColDetails(0, iLoop) & "] " & _
                                "FROM " & mstrTempTableName & " " & _
                                " " & sWhereCode & " "
                              
         If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
           sSQL = sSQL & " AND ([" & mvarColDetails(0, iLoop) & "] <> 0) "
         End If
        
         sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'avg_" & Trim(Str(iLoop)) & "'"
        End If
      Else
        'Bradford Index
        ' JPD20020712 Fault 4155
        sSQL = sSQL & IIf(Len(sSQL) = 0, "SELECT", ",") & _
        " avg(convert(float,[" & mvarColDetails(0, iLoop) & "])) AS avg_" & Trim(Str(iLoop))
      End If
    End If
    
    If mvarColDetails(5, iLoop) Then
      ' Count.
      mblnReportHasSummaryInfo = True
      
      'TM20010928 Fault 2291
      'Add a hidden key '*count*' so that when outputting to excel it does not format the
      'count to a date.
      sCountAddString = "*count*" & vbTab & "Sub Count" & vbTab

      If Not mbIsBradfordIndexReport Then
        If mvarColDetails(20, iLoop) Then
          sSQL = sSQL & ",(SELECT COUNT([?ID_" & mvarColDetails(15, iLoop) & "]) " & _
                        "FROM (SELECT DISTINCT [?ID_" & mvarColDetails(15, iLoop) & "], [" & mvarColDetails(0, iLoop) & "] " & _
                                "FROM " & mstrTempTableName & " " & _
                                " " & sWhereCode & " "
                              
         If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
           sSQL = sSQL & " AND ([" & mvarColDetails(0, iLoop) & "] <> 0) "
         End If
        
         sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'cnt_" & Trim(Str(iLoop)) & "'"
        Else
          sSQL = sSQL & ",(SELECT COUNT([?ID]) " & _
                        "FROM (SELECT DISTINCT [?ID], [" & mvarColDetails(0, iLoop) & "] " & _
                                "FROM " & mstrTempTableName & " " & _
                                " " & sWhereCode & " "
                              
         If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
           sSQL = sSQL & " AND ([" & mvarColDetails(0, iLoop) & "] <> 0) "
         End If
        
         sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'cnt_" & Trim(Str(iLoop)) & "'"
        End If
      Else
        'Bradford Index
        sSQL = sSQL & IIf(Len(sSQL) = 0, "SELECT", ",") & _
        " count(*) AS cnt_" & Trim(Str(iLoop))
      End If
    End If
    
    If mvarColDetails(6, iLoop) Then
      ' Total.
      mblnReportHasSummaryInfo = True
      
      'TM20020718 Fault 4169 - indicate in the hidden column that the row is a total row.
      sTotalAddString = "*total*" & vbTab & "Sub Total" & vbTab

      If Not mbIsBradfordIndexReport Then
        If mvarColDetails(20, iLoop) Then
          sSQL = sSQL & ",(SELECT SUM([" & mvarColDetails(0, iLoop) & "]) " & _
                        "FROM (SELECT DISTINCT [?ID_" & mvarColDetails(15, iLoop) & "], [" & mvarColDetails(0, iLoop) & "] " & _
                                "FROM " & mstrTempTableName & " " & _
                                " " & sWhereCode & " "
                              
          If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
            sSQL = sSQL & " AND ([" & mvarColDetails(0, iLoop) & "] <> 0) "
          End If
        
          sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'ttl_" & Trim(Str(iLoop)) & "'"
        Else
          sSQL = sSQL & ",(SELECT SUM([" & mvarColDetails(0, iLoop) & "]) " & _
                        "FROM (SELECT DISTINCT [?ID], [" & mvarColDetails(0, iLoop) & "] " & _
                                "FROM " & mstrTempTableName & " " & _
                                " " & sWhereCode & " "
                              
          If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
            sSQL = sSQL & " AND ([" & mvarColDetails(0, iLoop) & "] <> 0) "
          End If
        
          sSQL = sSQL & ") AS [vt." & Str(iLoop) & "]) AS 'ttl_" & Trim(Str(iLoop)) & "'"
        End If
      Else
        'Bradford Index
        sSQL = sSQL & IIf(Len(sSQL) = 0, "SELECT", ",") & _
        " sum([" & mvarColDetails(0, iLoop) & "])  AS ttl_" & Trim(Str(iLoop))
      End If
    End If
  Next iLoop

  If Len(sSQL) > 0 Then
  
    If Not mbIsBradfordIndexReport Then
      sSQL = "SELECT " & Right(sSQL, Len(sSQL) - 1)
    Else
      sSQL = sSQL & _
        " FROM " & mstrTempTableName & IIf(Len(sFromCode) > 0, sFromCode, "") & _
        IIf(Len(sWhereCode) > 0, sWhereCode, "")
    End If
    
    Set rsTemp = datGeneral.GetRecords(sSQL)
    
    For iLoop = 1 To UBound(mvarColDetails, 2)
      intColCounter = intColCounter + 1
    
      If mvarColDetails(4, iLoop) Then
        
        If Not mvarColDetails(19, iLoop) _
          And (Not mvarColDetails(24, iLoop)) _
          And (Not mvarColDetails(24, iLoop - 1)) Then
          fHasAverage = True
        End If
        
        ' Average.
        If IsNull(rsTemp.Fields("avg_" & Trim(Str(iLoop)))) Then
          strAggrValue = "0"
        'TM20020430 Fault 3810 - if the size and decimals of the report column are zero then
        'do not format the data, show it as it is.
        ElseIf mvarColDetails(2, iLoop) = 0 And mvarColDetails(1, iLoop) = 0 Then
          strAggrValue = rsTemp.Fields("avg_" & Trim(Str(iLoop)))
        ElseIf mvarColDetails(1, iLoop) > 0 And mvarColDetails(2, iLoop) = 0 Then
          strAggrValue = Format(rsTemp.Fields("avg_" & Trim(Str(iLoop))), "#0")
        ElseIf mvarColDetails(2, iLoop) = 0 Then
          strAggrValue = rsTemp.Fields("avg_" & Trim(Str(iLoop)))
        Else
          strAggrValue = Format(rsTemp.Fields("avg_" & Trim(Str(iLoop))), "0." & String(mvarColDetails(2, iLoop), "0"))
        End If
        
        If mvarColDetails(22, iLoop) Then
          If mvarColDetails(2, iLoop) = 0 And (InStr(1, strAggrValue, ".") <= 0) Then
            strAggrValue = Format(strAggrValue, "#,0" & String(mvarColDetails(2, iLoop), "0"))
          ElseIf (mvarColDetails(1, iLoop) > 0) And (mvarColDetails(2, iLoop) = 0) Then
            strAggrValue = Format(strAggrValue, "#,0")
          ElseIf mvarColDetails(2, iLoop) = 0 Then
            strAggrValue = Format(strAggrValue, "#,0.#")
          Else
            strAggrValue = Format(strAggrValue, "#,0." & String(mvarColDetails(2, iLoop), "0"))
          End If
        End If
        
        sAverageAddString = sAverageAddString & strAggrValue & vbTab
        
'        If Not IsNull(rsTemp.Fields("avg_" & Trim(Str(iLoop)))) Then
'          If mfrmOutput.TextWidth(rsTemp.Fields("avg_" & Trim(Str(iLoop)))) > mlngColWidth(intColCounter) Then
'            mlngColWidth(intColCounter) = mfrmOutput.TextWidth(rsTemp.Fields("avg_" & Trim(Str(iLoop)))) + 200
'          End If
'        End If
        If Not IsNull(strAggrValue) Then
          If mfrmOutput.TextWidth(strAggrValue) > mlngColWidth(intColCounter) Then
            mlngColWidth(intColCounter) = mfrmOutput.TextWidth(strAggrValue)
          End If
        End If
        
        strAggrValue = vbNullString
        
      Else
'        If (mvarColDetails(24, iLoop) = False) Then
          ' Display the value ?
          fDoValue = False
          If (mvarColDetails(9, iLoop)) Then
            For iLoop2 = 1 To UBound(mvarSortOrder, 2)
              If mvarSortOrder(1, iLoop2) = mvarColDetails(12, iLoop) Then
                fDoValue = (iLoop2 <= piSortIndex)
                Exit For
              End If
            Next iLoop2
          End If
          
          If fDoValue Then
            sAverageAddString = sAverageAddString & PopulateGrid_FormatData(CStr(mvarColDetails(0, iLoop)), pavColumns(3, iLoop), False, True) & vbTab
          Else
            sAverageAddString = sAverageAddString & vbTab
          End If
'        End If
      End If
    
      If mvarColDetails(5, iLoop) Then
        ' Count.
        
        If Not mvarColDetails(19, iLoop) _
          And (Not mvarColDetails(24, iLoop)) _
          And (Not mvarColDetails(24, iLoop - 1)) Then
          fHasCount = True
        End If
        
        'JDM - Make a note of count the Bradford Index Report
        If mbIsBradfordIndexReport Then
          miAmountOfRecords = IIf(Not IsNull(rsTemp.Fields("cnt_" & Trim(Str(iLoop)))), rsTemp.Fields("cnt_" & Trim(Str(iLoop))), 0)
        End If
        
        sCountAddString = sCountAddString & IIf(IsNull(rsTemp.Fields("cnt_" & Trim(Str(iLoop)))), "0", Format(rsTemp.Fields("cnt_" & Trim(Str(iLoop))), "0")) & vbTab
        
        If Not IsNull(rsTemp.Fields("cnt_" & Trim(Str(iLoop)))) Then
          If mfrmOutput.TextWidth(rsTemp.Fields("cnt_" & Trim(Str(iLoop)))) > mlngColWidth(intColCounter) Then
            'mlngColWidth(intColCounter) = mfrmOutput.TextWidth(rsTemp.Fields("cnt_" & Trim(Str(iLoop)))) + 200
            mlngColWidth(intColCounter) = mfrmOutput.TextWidth(rsTemp.Fields("cnt_" & Trim(Str(iLoop))))
          End If
        End If
      Else
'        If (mvarColDetails(24, iLoop) = False) Then
          ' Display the value ?
          fDoValue = False
          If (mvarColDetails(9, iLoop)) Then
            For iLoop2 = 1 To UBound(mvarSortOrder, 2)
              If mvarSortOrder(1, iLoop2) = mvarColDetails(12, iLoop) Then
                fDoValue = (iLoop2 <= piSortIndex)
                Exit For
              End If
            Next iLoop2
          End If
            
          If (mbIsBradfordIndexReport And mblnCustomReportsSummaryReport) And (mbBradfordCount) Then
            fDoValue = True
          End If
          
         If fDoValue Then
            sCountAddString = sCountAddString & PopulateGrid_FormatData(CStr(mvarColDetails(0, iLoop)), pavColumns(3, iLoop), False, True) & vbTab
          Else
            sCountAddString = sCountAddString & vbTab
          End If
'        End If
      End If
    
      If mvarColDetails(6, iLoop) Then
        ' Total.
                
        If Not mvarColDetails(19, iLoop) _
          And (Not mvarColDetails(24, iLoop)) _
          And (Not mvarColDetails(24, iLoop - 1)) Then
          fHasTotal = True
        End If
        
        If IsNull(rsTemp.Fields("ttl_" & Trim(Str(iLoop)))) Then
'          sTotalAddString = sTotalAddString & "0" & vbTab
          strAggrValue = "0"
        'TM20020430 Fault 3810 - if the size and decimals of the report column are zero then
        'do not format the data, show it as it is.
        ElseIf mvarColDetails(2, iLoop) = 0 And mvarColDetails(1, iLoop) = 0 Then
          strAggrValue = rsTemp.Fields("ttl_" & Trim(Str(iLoop)))
        ElseIf mvarColDetails(2, iLoop) = 0 Then
          strAggrValue = Format(rsTemp.Fields("ttl_" & Trim(Str(iLoop))), "0")
        Else
          strAggrValue = Format(rsTemp.Fields("ttl_" & Trim(Str(iLoop))), "0." & String(mvarColDetails(2, iLoop), "0"))
        End If
        
        If mvarColDetails(22, iLoop) Then
          If mvarColDetails(2, iLoop) = 0 Then
            strAggrValue = Format(strAggrValue, "#,0" & String(mvarColDetails(2, iLoop), "0"))
          Else
            strAggrValue = Format(strAggrValue, "#,0." & String(mvarColDetails(2, iLoop), "0"))
          End If
        End If
        
        sTotalAddString = sTotalAddString & strAggrValue & vbTab

        If Not IsNull(strAggrValue) Then
          If mfrmOutput.TextWidth(strAggrValue) > mlngColWidth(intColCounter) Then
            mlngColWidth(intColCounter) = mfrmOutput.TextWidth(strAggrValue)
          End If
        End If
        
        strAggrValue = vbNullString
      Else
'        If (mvarColDetails(24, iLoop) = False) Then
          ' Display the value ?
          fDoValue = False
          If (mvarColDetails(9, iLoop)) Then
            For iLoop2 = 1 To UBound(mvarSortOrder, 2)
              If mvarSortOrder(1, iLoop2) = mvarColDetails(12, iLoop) Then
                fDoValue = (iLoop2 <= piSortIndex)
                Exit For
              End If
            Next iLoop2
          End If
            
          If (mbIsBradfordIndexReport And mblnCustomReportsSummaryReport) Then
            If Not mbBradfordCount Then
              fDoValue = True
            End If
          End If
            
          If fDoValue Then
            sTotalAddString = sTotalAddString & PopulateGrid_FormatData(CStr(mvarColDetails(0, iLoop)), pavColumns(3, iLoop), False, True) & vbTab
          Else
            sTotalAddString = sTotalAddString & vbTab
          End If
'        End If
      End If

    Next iLoop
  
    rsTemp.Close
    Set rsTemp = Nothing
  End If

  ' Do a different summary if we are a Bradford Index Report
  If Not mbIsBradfordIndexReport Then

      ' Put a blank line in here if its not a page break as well
'MH20030916
      
      If ((Not mblnCustomReportsSummaryReport) And (fHasAverage Or fHasCount Or fHasTotal)) _
          Or ((Not mblnCustomReportsSummaryReport) And Not (fHasAverage Or fHasCount Or fHasTotal Or (mvarColDetails(8, iColumnIndex)))) Then
'      If (fHasAverage Or fHasCount Or fHasTotal) Then
'        If Not mvarColDetails(8, piColumnIndex) Then
'        If Not mblnPageBreak Then
          mfrmOutput.grdOutput.AddItem "*indicator*"
          mintPageBreakRowIndex = mintPageBreakRowIndex + 1
'        End If
      End If
      
      
      If fHasAverage Then
        mfrmOutput.grdOutput.AddItem sAverageAddString
        mintPageBreakRowIndex = mintPageBreakRowIndex + 1
      End If
      
      If fHasCount Then
        mfrmOutput.grdOutput.AddItem sCountAddString
        mintPageBreakRowIndex = mintPageBreakRowIndex + 1
      End If

      If fHasTotal Then
        mfrmOutput.grdOutput.AddItem sTotalAddString
        mintPageBreakRowIndex = mintPageBreakRowIndex + 1
      End If

      If Not mblnCustomReportsSummaryReport Then
        If (Not mvarColDetails(8, iColumnIndex)) Then
          If fHasAverage Or fHasCount Or fHasTotal Then
            mfrmOutput.grdOutput.AddItem "*indicator*"
            mintPageBreakRowIndex = mintPageBreakRowIndex + 1
          End If
        End If
      End If

  Else
      mblnReportHasSummaryInfo = False

      asBradfordSummaryLine = Split(sTotalAddString, vbTab)

      ' Build Bradford Total Summary
      asBradfordSummaryLine(11) = "Total"
      asBradfordSummaryLine(13) = Val(Str(asBradfordSummaryLine(13)))
      asBradfordSummaryLine(14) = Val(Str(asBradfordSummaryLine(14)))
      sTotalAddString = Join(asBradfordSummaryLine, vbTab)

      ' Calculate Bradford index line
      asBradfordSummaryLine(11) = "Bradford Factor"

      If mbBradfordWorkings = True Then
        asBradfordSummaryLine(13) = CStr(Val(asBradfordSummaryLine(13)) * (miAmountOfRecords * miAmountOfRecords)) + " (" + Str(miAmountOfRecords) + Chr(178) + " * " + asBradfordSummaryLine(13) + ")"
        asBradfordSummaryLine(14) = CStr(Val(asBradfordSummaryLine(14)) * (miAmountOfRecords * miAmountOfRecords)) + " (" + Str(miAmountOfRecords) + Chr(178) + " * " + asBradfordSummaryLine(14) + ")"
      Else
        asBradfordSummaryLine(13) = asBradfordSummaryLine(13) * (miAmountOfRecords * miAmountOfRecords)
        asBradfordSummaryLine(14) = asBradfordSummaryLine(14) * (miAmountOfRecords * miAmountOfRecords)
      End If
      
      Dim iWidthTemp As Integer
      For iWidthTemp = 11 To 14 Step 1
        'If (mfrmOutput.TextWidth(asBradfordSummaryLine(iWidthTemp)) + 200) > mlngColWidth(iWidthTemp) Then
        '  mlngColWidth(iWidthTemp) = mfrmOutput.TextWidth(asBradfordSummaryLine(iWidthTemp)) + 200
        'End If
        If (mfrmOutput.TextWidth(asBradfordSummaryLine(iWidthTemp))) > mlngColWidth(iWidthTemp) Then
          mlngColWidth(iWidthTemp) = mfrmOutput.TextWidth(asBradfordSummaryLine(iWidthTemp))
        End If
      Next iWidthTemp
      
      If (mblnCustomReportsSummaryReport) And (mbBradfordCount Or mbBradfordTotals) Then
        asBradfordSummaryLine(2) = vbNullString
        asBradfordSummaryLine(3) = vbNullString
        asBradfordSummaryLine(4) = vbNullString
        asBradfordSummaryLine(5) = vbNullString
      End If
      
      sBradfordSummary = Join(asBradfordSummaryLine, vbTab)

      ' Build Bradford Count Summary
      asBradfordSummaryLine = Split(sCountAddString, vbTab)
      asBradfordSummaryLine(11) = "Instances"
      sCountAddString = Join(asBradfordSummaryLine, vbTab)

      ' Add the summary lines
      If mbBradfordCount = True Then
        mfrmOutput.grdOutput.AddItem sCountAddString
        mintPageBreakRowIndex = mintPageBreakRowIndex + 1
      End If

      If mbBradfordTotals = True Then
        mfrmOutput.grdOutput.AddItem sTotalAddString
        mintPageBreakRowIndex = mintPageBreakRowIndex + 1
      End If

      mfrmOutput.grdOutput.AddItem sBradfordSummary
      mintPageBreakRowIndex = mintPageBreakRowIndex + 1
      mfrmOutput.grdOutput.AddItem "*indicator*"
      mintPageBreakRowIndex = mintPageBreakRowIndex + 1

  End If

End Function


Private Function PopulateGrid_DoGrandSummary() As Boolean

  ' Purpose : To calculate the final grand summaries
  ' Input   : None
  ' Output  : True/False

  On Error GoTo PopulateGrid_DoGrandSummary_ERROR

  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim rsTemp As Recordset
  
  Dim sAverageAddString As String
  Dim sCountAddString As String
  Dim sTotalAddString As String
  
  Dim fHasAverage As Boolean
  Dim fHasCount As Boolean
  Dim fHasTotal As Boolean
  
  Dim sSQL As String
  
  Dim intColCounter As Integer
  
  Dim strAggrValue As String
  
  intColCounter = 1
  strAggrValue = vbNullString
  
  sAverageAddString = vbNullString
  sCountAddString = vbNullString
  sTotalAddString = vbNullString
  
  ' Construct the required select statement.
  sSQL = vbNullString
  
  For iLoop = 1 To mrstCustomReportsDetails.RecordCount
    If mvarColDetails(4, iLoop) Then
      ' Average.
      
      'TM20020718 Fault 4169 - indicate in the hidden column that the row is an average row.
      sAverageAddString = "*average*" & vbTab & "Average" & vbTab
      
      If mvarColDetails(20, iLoop) Then
        ' JPD20020712 Fault 4155
        sSQL = sSQL & ",(SELECT AVG(convert(float, [" & mvarColDetails(0, iLoop) & "])) " & _
                      "FROM (SELECT DISTINCT [?ID_" & mvarColDetails(15, iLoop) & "], [" & mvarColDetails(0, iLoop) & "] " & _
                              "FROM " & mstrTempTableName & " "
                              
        If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
          sSQL = sSQL & "WHERE ([" & mvarColDetails(0, iLoop) & "] <> 0) "
        End If
       
        sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'avg_" & Trim(Str(iLoop)) & "'"
                              
      Else
        ' JPD20020712 Fault 4155
        sSQL = sSQL & ",(SELECT AVG(convert(float, [" & mvarColDetails(0, iLoop) & "])) " & _
                      "FROM (SELECT DISTINCT [?ID], [" & mvarColDetails(0, iLoop) & "] " & _
                              "FROM " & mstrTempTableName & " "
                              
        If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
          sSQL = sSQL & "WHERE ([" & mvarColDetails(0, iLoop) & "] <> 0) "
        End If
       
        sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'avg_" & Trim(Str(iLoop)) & "'"
      End If
     
    End If
    
    If mvarColDetails(5, iLoop) Then
      ' Count.
      
      'TM20010928 Fault 2291
      'Add a hidden key '*count*' so that when outputting to excel it does not format the
      'count to a date.
      sCountAddString = "*count*" & vbTab & "Count" & vbTab

      If mvarColDetails(20, iLoop) Then
        sSQL = sSQL & ",(SELECT COUNT([?ID_" & mvarColDetails(15, iLoop) & "]) " & _
                      "FROM (SELECT DISTINCT [?ID_" & mvarColDetails(15, iLoop) & "], [" & mvarColDetails(0, iLoop) & "] " & _
                              "FROM " & mstrTempTableName & " "
                              
        If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
          sSQL = sSQL & "WHERE ([" & mvarColDetails(0, iLoop) & "] <> 0) "
        End If
       
        sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'cnt_" & Trim(Str(iLoop)) & "'"
      Else
        sSQL = sSQL & ",(SELECT COUNT([?ID]) " & _
                      "FROM (SELECT DISTINCT [?ID], [" & mvarColDetails(0, iLoop) & "] " & _
                              "FROM " & mstrTempTableName & " "
                              
        If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
          sSQL = sSQL & "WHERE ([" & mvarColDetails(0, iLoop) & "] <> 0) "
        End If
       
        sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'cnt_" & Trim(Str(iLoop)) & "'"
      End If

    End If
    
    If mvarColDetails(6, iLoop) Then
      ' Total.
      
      'TM20020718 Fault 4169 - indicate in the hidden column that the row is a total row.
      sTotalAddString = "*total*" & vbTab & "Total" & vbTab
     
      If mvarColDetails(20, iLoop) Then
        sSQL = sSQL & ",(SELECT SUM([" & mvarColDetails(0, iLoop) & "]) " & _
                      "FROM (SELECT DISTINCT [?ID_" & mvarColDetails(15, iLoop) & "], [" & mvarColDetails(0, iLoop) & "] " & _
                              "FROM " & mstrTempTableName & " "
                              
        If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
          sSQL = sSQL & "WHERE ([" & mvarColDetails(0, iLoop) & "] <> 0) "
        End If
       
        sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'ttl_" & Trim(Str(iLoop)) & "'"
      Else
        sSQL = sSQL & ",(SELECT SUM([" & mvarColDetails(0, iLoop) & "]) " & _
                      "FROM (SELECT DISTINCT [?ID], [" & mvarColDetails(0, iLoop) & "] " & _
                              "FROM " & mstrTempTableName & " "
                              
        If mblnIgnoreZerosInAggregates And mvarColDetails(3, iLoop) Then
          sSQL = sSQL & "WHERE ([" & mvarColDetails(0, iLoop) & "] <> 0) "
        End If
       
        sSQL = sSQL & " ) AS [vt." & Str(iLoop) & "]) AS 'ttl_" & Trim(Str(iLoop)) & "'"
      End If

    End If
  Next iLoop

  If Len(sSQL) > 0 Then
    sSQL = "SELECT " & Right(sSQL, Len(sSQL) - 1)
    
    Set rsTemp = datGeneral.GetRecords(sSQL)
    
    For iLoop = 1 To UBound(mvarColDetails, 2)
      intColCounter = intColCounter + 1
      
      If mvarColDetails(4, iLoop) Then
        ' Average.
     
        If Not mvarColDetails(19, iLoop) _
          And (Not mvarColDetails(24, iLoop)) _
          And (Not mvarColDetails(24, iLoop - 1)) Then
          fHasAverage = True
        End If

        If IsNull(rsTemp.Fields("avg_" & Trim(Str(iLoop)))) Then
'          sAverageAddString = sAverageAddString & "0" & vbTab
          strAggrValue = "0"
        'TM20020430 Fault 3810 - if the size and decimals of the report column are zero then
        'do not format the data, show it as it is.
        ElseIf mvarColDetails(2, iLoop) = 0 And mvarColDetails(1, iLoop) = 0 Then
          strAggrValue = rsTemp.Fields("avg_" & Trim(Str(iLoop)))
        ElseIf mvarColDetails(1, iLoop) > 0 And mvarColDetails(2, iLoop) = 0 Then
          strAggrValue = Format(rsTemp.Fields("avg_" & Trim(Str(iLoop))), "#0")
        ElseIf mvarColDetails(2, iLoop) = 0 Then
          strAggrValue = rsTemp.Fields("avg_" & Trim(Str(iLoop)))
        Else
          strAggrValue = Format(rsTemp.Fields("avg_" & Trim(Str(iLoop))), "0." & String(mvarColDetails(2, iLoop), "0"))
        End If
      
        If mvarColDetails(22, iLoop) Then
          If mvarColDetails(2, iLoop) = 0 And (InStr(1, strAggrValue, ".") <= 0) Then
            strAggrValue = Format(strAggrValue, "#,0" & String(mvarColDetails(2, iLoop), "0"))
          ElseIf (mvarColDetails(1, iLoop) > 0) And (mvarColDetails(2, iLoop) = 0) Then
            strAggrValue = Format(strAggrValue, "#,0")
          ElseIf mvarColDetails(2, iLoop) = 0 Then
            strAggrValue = Format(strAggrValue, "#,0.#")
          Else
            strAggrValue = Format(strAggrValue, "#,0." & String(mvarColDetails(2, iLoop), "0"))
          End If
        End If
        
        sAverageAddString = sAverageAddString & strAggrValue & vbTab
        
'        If Not IsNull(rsTemp.Fields("avg_" & Trim(Str(iLoop)))) Then
'          If Len(rsTemp.Fields("avg_" & Trim(Str(iLoop)))) > mlngColWidth(intColCounter) Then
'            mlngColWidth(intColCounter) = Len(rsTemp.Fields("avg_" & Trim(Str(iLoop))))
'          End If
'        End If
        If Not IsNull(strAggrValue) Then
          If mfrmOutput.TextWidth(strAggrValue) > mlngColWidth(intColCounter) Then
            mlngColWidth(intColCounter) = mfrmOutput.TextWidth(strAggrValue)
          End If
        End If

        strAggrValue = vbNullString
      Else
'        ' Display the value ?
'        fDoValue = False
'        If (mvarColDetails(9, iLoop)) Then
'          For iLoop2 = 1 To UBound(mvarSortOrder, 2)
'            If mvarSortOrder(1, iLoop2) = mvarColDetails(12, iLoop) Then
'              fDoValue = (iLoop2 <= piSortIndex)
'              Exit For
'            End If
'          Next iLoop2
'        End If
'
'        If fDoValue Then
'          sAverageAddString = sAverageAddString & pavColumns(3, iLoop) & vbTab
'        Else
'        If (mvarColDetails(24, iLoop) = False) Then
          sAverageAddString = sAverageAddString & vbTab
'        End If
'        End If
      End If
    
      If mvarColDetails(5, iLoop) Then
        ' Count.
        
        If Not mvarColDetails(19, iLoop) _
          And (Not mvarColDetails(24, iLoop)) _
          And (Not mvarColDetails(24, iLoop - 1)) Then
          fHasCount = True
        End If

        sCountAddString = sCountAddString & IIf(IsNull(rsTemp.Fields("cnt_" & Trim(Str(iLoop)))), "0", Format(rsTemp.Fields("cnt_" & Trim(Str(iLoop))), "0")) & vbTab
      
        If Not IsNull(rsTemp.Fields("cnt_" & Trim(Str(iLoop)))) Then
          If mfrmOutput.TextWidth(rsTemp.Fields("cnt_" & Trim(Str(iLoop)))) > mlngColWidth(intColCounter) Then
            mlngColWidth(intColCounter) = mfrmOutput.TextWidth(rsTemp.Fields("cnt_" & Trim(Str(iLoop))))
          End If
        End If
      Else
'        ' Display the value ?
'        fDoValue = False
'        If (mvarColDetails(9, iLoop)) Then
'          For iLoop2 = 1 To UBound(mvarSortOrder, 2)
'            If mvarSortOrder(1, iLoop2) = mvarColDetails(12, iLoop) Then
'              fDoValue = (iLoop2 <= piSortIndex)
'              Exit For
'            End If
'          Next iLoop2
'        End If
'
'        If fDoValue Then
'          sCountAddString = sCountAddString & pavColumns(3, iLoop) & vbTab
'        Else
'        If (mvarColDetails(24, iLoop) = False) Then
          sCountAddString = sCountAddString & vbTab
'        End If
'        End If
      End If
    
      If mvarColDetails(6, iLoop) Then
        ' Total.

        If Not mvarColDetails(19, iLoop) _
          And (Not mvarColDetails(24, iLoop)) _
          And (Not mvarColDetails(24, iLoop - 1)) Then
          fHasTotal = True
        End If

        If IsNull(rsTemp.Fields("ttl_" & Trim(Str(iLoop)))) Then
'          sTotalAddString = sTotalAddString & "0" & vbTab
          strAggrValue = "0"
        'TM20020430 Fault 3810 - if the size and decimals of the report column are zero then
        'do not format the data, show it as it is.
        ElseIf mvarColDetails(2, iLoop) = 0 And mvarColDetails(1, iLoop) = 0 Then
'          sTotalAddString = sTotalAddString & rsTemp.Fields("ttl_" & Trim(Str(iLoop))) & vbTab
          strAggrValue = rsTemp.Fields("ttl_" & Trim(Str(iLoop)))
        ElseIf mvarColDetails(2, iLoop) = 0 Then
'          sTotalAddString = sTotalAddString & Format(rsTemp.Fields("ttl_" & Trim(Str(iLoop))), "0") & vbTab
          strAggrValue = Format(rsTemp.Fields("ttl_" & Trim(Str(iLoop))), "0")
        Else
'          sTotalAddString = sTotalAddString & Format(rsTemp.Fields("ttl_" & Trim(Str(iLoop))), "0." & String(mvarColDetails(2, iLoop), "0")) & vbTab
          strAggrValue = Format(rsTemp.Fields("ttl_" & Trim(Str(iLoop))), "0." & String(mvarColDetails(2, iLoop), "0"))
        End If
        
        If mvarColDetails(22, iLoop) Then
          If (mvarColDetails(2, iLoop) = 0) Then
            strAggrValue = Format(strAggrValue, "#,0" & String(mvarColDetails(2, iLoop), "0"))
          Else
            strAggrValue = Format(strAggrValue, "#,0." & String(mvarColDetails(2, iLoop), "0"))
          End If
        End If
        
        sTotalAddString = sTotalAddString & strAggrValue & vbTab
        
        If Not IsNull(strAggrValue) Then
          If mfrmOutput.TextWidth(strAggrValue) > mlngColWidth(intColCounter) Then
            mlngColWidth(intColCounter) = mfrmOutput.TextWidth(strAggrValue)
          End If
        End If
        
        strAggrValue = vbNullString
      Else
'        ' Display the value ?
'        fDoValue = False
'        If (mvarColDetails(9, iLoop)) Then
'          For iLoop2 = 1 To UBound(mvarSortOrder, 2)
'            If mvarSortOrder(1, iLoop2) = mvarColDetails(12, iLoop) Then
'              fDoValue = (iLoop2 <= piSortIndex)
'              Exit For
'            End If
'          Next iLoop2
'        End If
'
'        If fDoValue Then
'          sTotalAddString = sTotalAddString & pavColumns(3, iLoop) & vbTab
'        Else
'        If (mvarColDetails(24, iLoop) = False) Then
          sTotalAddString = sTotalAddString & vbTab
'        End If
'        End If
      End If

    Next iLoop
  
    rsTemp.Close
    Set rsTemp = Nothing
  End If

  mblnDoesHaveGrandSummary = (fHasAverage Or fHasCount Or fHasTotal)

  'Output the 4 lines of grand aggregates (blank,AVG,CNT,TTL)
  If mblnDoesHaveGrandSummary Then
    mfrmOutput.grdOutput.AddItem IIf(mblnPageBreak, "*", "*indicator*")
  End If
  
  If fHasAverage Then
    mblnReportHasSummaryInfo = True
    mfrmOutput.grdOutput.AddItem sAverageAddString
    mintPageBreakRowIndex = mintPageBreakRowIndex + 1
  End If

  If fHasCount Then
    mblnReportHasSummaryInfo = True
    mfrmOutput.grdOutput.AddItem sCountAddString
    mintPageBreakRowIndex = mintPageBreakRowIndex + 1
  End If

  If fHasTotal Then
    mblnReportHasSummaryInfo = True
    mfrmOutput.grdOutput.AddItem sTotalAddString
    mintPageBreakRowIndex = mintPageBreakRowIndex + 1
  End If

'MH20030916 Don't need this at all?
'  'MH20030707 Fault 5792 blank line at bottom of grid was throwing extra page.
'  If Not mblnCustomReportsSummaryReport Then
'    If fHasAverage Or fHasCount Or fHasTotal Then
'      mfrmOutput.grdOutput.AddItem "*indicator*"
'      mintPageBreakRowIndex = mintPageBreakRowIndex + 1
'    End If
'  End If

  Exit Function
  
PopulateGrid_DoGrandSummary_ERROR:
  
  Set rsTemp = Nothing
  mstrErrorString = "Error while calculating grand summary." & vbNewLine & "(" & Err.Description & ")"
  PopulateGrid_DoGrandSummary = False
  
End Function

Private Function ClearUp() As Boolean

  ' Purpose : To clear all variables/recordsets/references and drops temptable
  ' Input   : None
  ' Output  : True/False success

  ' Definition variables
  
  On Error GoTo ClearUp_ERROR
  
  mlngCustomReportID = 0
  mstrCustomReportsName = vbNullString
  mstrCustomReportsDescription = vbNullString
  mlngCustomReportsBaseTable = 0
  mstrCustomReportsBaseTableName = vbNullString
  mlngCustomReportsAllRecords = 1
  mlngCustomReportsPickListID = 0
  mlngCustomReportsFilterID = 0
  mlngCustomReportsParent1Table = 0
  mstrCustomReportsParent1TableName = vbNullString
  mlngCustomReportsParent1FilterID = 0
  mlngCustomReportsParent2Table = 0
  mstrCustomReportsParent2TableName = vbNullString
  mlngCustomReportsParent2FilterID = 0
'  mlngCustomReportsChildTable = 0
'  mstrCustomReportsChildTableName = vbNullString
'  mlngCustomReportsChildFilterID = 0
'  mlngCustomReportsChildMaxRecords = 0
  mblnCustomReportsSummaryReport = False
  mblnCustomReportsPrintFilterHeader = False
  
  mlngCustomReportsParent1AllRecords = 1
  mlngCustomReportsParent1PickListID = 0
  mlngCustomReportsParent2AllRecords = 1
  mlngCustomReportsParent2PickListID = 0
  
  miChildTablesCount = 0
  
  ' Recordsets
  
  Set mrstCustomReportsDetails = Nothing
  Set mrstCustomReportsOutput = Nothing
  
  'TM20020531 Fault 3756
'  ' Delete the temptable if exists, and then clear the variable
'  If Len(mstrTempTableName) > 0 Then
'    mclsData.ExecuteSql ("IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & _
'                      "DROP TABLE " & mstrTempTableName)
'  End If
  datGeneral.DropUniqueSQLObject mstrTempTableName, 3
  mstrTempTableName = vbNullString

  ' SQL strings
  
  mstrSQLSelect = vbNullString
  mstrSQLFrom = vbNullString
  mstrSQLWhere = vbNullString
  mstrSQLJoin = vbNullString
  mstrSQLOrderBy = vbNullString
  mstrSQL = vbNullString
  
  ' Class references
  
  Set mclsData = Nothing
  Set mclsGeneral = Nothing
  
  ' Arrays
  
'  ReDim mvarColDetails(24, 0)
  ReDim mvarSortOrder(2, 0)
  ReDim mlngColWidth(1, 0)
  ReDim mvarChildTables(5, 0)
  
  ' Flags
  
  mblnReportHasSummaryInfo = False
  mblnReportHasPageBreak = False
  
  ' Column Privilege arrays / collections / variables
  
  mstrBaseTableRealSource = vbNullString
  mstrRealSource = vbNullString
  Set mobjTableView = Nothing
  Set mobjColumnPrivileges = Nothing
  ReDim mlngTableViews(2, 0)
  ReDim mstrViews(0)
  ReDim mvarGroupWith(1, 0)
  ReDim mvarPageBreak(0)
  
  gobjProgress.ResetBar2
  
  ClearUp = True
  Exit Function
  
ClearUp_ERROR:
  
  mstrErrorString = "Error whilst clearing data." & vbNewLine & "(" & Err.Description & ")"
  ClearUp = False
  
End Function

'Private Function InitialiseOutPutFormWithDefaults() As Boolean
'
'  ' Purpose : To instantiate the output options form and load with
'  '           the defaults for the current custom report
'
'  On Error GoTo InitialiseOutPutFormWithDefaults_ERROR
'
''MH20031203
''  ' Load the output form and hide the page/percentage frame used by MH
''  frmOutputOptions.Initialise mfrmOutput, False
''
''  ' Set the relevant option group button and other options to the defaults
''  If mintCustomReportsDefaultOutput = 0 Then
''    frmOutputOptions.optOutput(0).Value = True
''  Else
''    frmOutputOptions.optOutput(1).Value = True
''    Select Case mintCustomReportsDefaultExportTo
''      Case 0: SetComboText frmOutputOptions.cboExportTo, "Html"
''      Case 1: SetComboText frmOutputOptions.cboExportTo, "Microsoft Excel"
''      Case 2: SetComboText frmOutputOptions.cboExportTo, "Microsoft Word"
''    End Select
''    If mblnCustomReportsDefaultSave Then frmOutputOptions.chkSave.Value = vbChecked Else frmOutputOptions.chkSave.Value = vbUnchecked
''    If mblnCustomReportsDefaultCloseApp Then frmOutputOptions.chkCloseApplication.Value = vbChecked Else frmOutputOptions.chkCloseApplication.Value = vbUnchecked
''    frmOutputOptions.txtFilename.Text = mstrCustomReportsDefaultSaveAs
''  End If
'
'  InitialiseOutPutFormWithDefaults = True
'  Exit Function
'
'InitialiseOutPutFormWithDefaults_ERROR:
'
'  mstrErrorString = "Error preparing the output option form with defaults." & vbNewLine & "(clsCustomReportsRUN.InitialiseOutPutFormWithDefaults : " & Err.Description & ")"
'  InitialiseOutPutFormWithDefaults = False
'
'End Function

Private Function IsRecordSelectionValid() As Boolean

  Dim sSQL As String
  'Dim lCount As Long
  Dim rsTemp As Recordset
  Dim iResult As RecordSelectionValidityCodes
  
' Base Table First
  If mlngCustomReportsFilterID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngCustomReportsFilterID)
    Select Case iResult
      Case REC_SEL_VALID_DELETED
        mstrErrorString = "The base table filter used in this definition has been deleted by another user."
      Case REC_SEL_VALID_INVALID
        mstrErrorString = "The base table filter used in this definition is invalid."
      Case REC_SEL_VALID_HIDDENBYOTHER
        If Not gfCurrentUserIsSysSecMgr Then
          mstrErrorString = "The base table filter used in this definition has been made hidden by another user."
        End If
    End Select
  ElseIf mlngCustomReportsPickListID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, mlngCustomReportsPickListID)
    Select Case iResult
      Case REC_SEL_VALID_DELETED
        mstrErrorString = "The base table picklist used in this definition has been deleted by another user."
      Case REC_SEL_VALID_INVALID
        mstrErrorString = "The base table picklist used in this definition is invalid."
      Case REC_SEL_VALID_HIDDENBYOTHER
        If Not gfCurrentUserIsSysSecMgr Then
          mstrErrorString = "The base table picklist used in this definition has been made hidden by another user."
        End If
      End Select
  End If
  
  If Len(mstrErrorString) = 0 Then
    ' Parent 1 Table
    If mlngCustomReportsParent1FilterID > 0 Then
      iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngCustomReportsParent1FilterID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "The first parent table filter used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "The first parent table filter used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "The first parent table filter used in this definition has been made hidden by another user."
          End If
      End Select
    ElseIf mlngCustomReportsParent1PickListID > 0 Then
      iResult = ValidateRecordSelection(REC_SEL_PICKLIST, mlngCustomReportsParent1PickListID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "The first parent table picklist used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "The first parent table picklist used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "The first parent table picklist used in this definition has been made hidden by another user."
          End If
        End Select
    End If
  End If
  
  ' Parent 2 Table
  If Len(mstrErrorString) = 0 Then
    If mlngCustomReportsParent2FilterID > 0 Then
      iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngCustomReportsParent2FilterID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "The second parent table filter used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "The second parent table filter used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "The second parent table filter used in this definition has been made hidden by another user."
          End If
      End Select
    ElseIf mlngCustomReportsParent2PickListID > 0 Then
      iResult = ValidateRecordSelection(REC_SEL_PICKLIST, mlngCustomReportsParent2PickListID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "The second parent table picklist used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "The second parent table picklist used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "The second parent table picklist used in this definition has been made hidden by another user."
          End If
      End Select
    End If
  End If
  
  Dim i As Integer
  Dim lngFilterID As Long
  
  ' Child Table
  If Len(mstrErrorString) = 0 Then
    If miChildTablesCount > 0 Then
      For i = 0 To UBound(mvarChildTables, 2) Step 1
        lngFilterID = mvarChildTables(1, i)
        If lngFilterID > 0 Then
          iResult = ValidateRecordSelection(REC_SEL_FILTER, lngFilterID)
          Select Case iResult
            Case REC_SEL_VALID_DELETED
              mstrErrorString = "The child table filter used in this definition has been deleted by another user."
            Case REC_SEL_VALID_INVALID
              mstrErrorString = "The child table filter used in this definition is invalid."
            Case REC_SEL_VALID_HIDDENBYOTHER
              If Not gfCurrentUserIsSysSecMgr Then
                mstrErrorString = "The child table filter used in this definition has been made hidden by another user."
              End If
          End Select
        End If
      
        If Len(mstrErrorString) > 0 Then
          Exit For
        End If
      Next i
    End If
  End If
  
  ' JDM - 13/10/03 - Fault 7228 - Problems if somehow a customreportid of 0 gets in the database.
  If Not mbIsBradfordIndexReport Then
  
    '******* Check calculations for hidden/deleted elements *******
    If Len(mstrErrorString) = 0 Then
      sSQL = "SELECT * FROM ASRSYSCustomReportsDetails " & _
               "WHERE CustomReportID = " & mlngCustomReportID & _
               " AND LOWER(Type) = 'e' "
    
      Set rsTemp = datGeneral.GetRecords(sSQL)
      With rsTemp
        If Not (.EOF And .BOF) Then
          .MoveFirst
          Do Until .EOF
            iResult = ValidateCalculation(!ColExprID)
            Select Case iResult
              Case REC_SEL_VALID_DELETED
                mstrErrorString = "A calculation used in this definition has been deleted by another user."
              Case REC_SEL_VALID_INVALID
                mstrErrorString = "A calculation used in this definition is invalid."
              Case REC_SEL_VALID_HIDDENBYOTHER
                If Not gfCurrentUserIsSysSecMgr Then
                  mstrErrorString = "A calculation used in this definition has been made hidden by another user."
                End If
            End Select
            
            If Len(mstrErrorString) > 0 Then
              Exit Do
            End If
            
            .MoveNext
          Loop
        End If
      End With
      
      Set rsTemp = Nothing
    End If
  End If
  
  IsRecordSelectionValid = (Len(mstrErrorString) = 0)
  
End Function


Private Function GetDataType(lColumnID As Long) As Long

  'Needed to be created as the one in datgeneral needs on tableid
  
  Dim sSQL As String
  Dim rsTemp As Recordset
  
  sSQL = "Select DataType From ASRSysColumns Where ColumnID = " & lColumnID
  Set rsTemp = New Recordset
  rsTemp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  GetDataType = rsTemp(0)
  
  rsTemp.Close
  Set rsTemp = Nothing

End Function
Private Function CalculateBradfordFactors() As Boolean

  ' Purpose : To calculate any bradford factors, and place into the created temporary table
  Dim sSQL As String
  'Dim cmADO As ADODB.Command
  'Dim pmADO As ADODB.Parameter
  'Dim iResult As Single
  Dim strReportStartDate As String
  Dim strReportEndDate As String

  On Error GoTo CalculateBradfordFactors_ERROR
 
  'TM24062004 Fault 8800 - we don't need to pass the full date AND time through.
'  strReportStartDate = Format(mdtAbsenceFrom, "YYYY/MM/DD 00:00:00.000")
'  strReportEndDate = Format(mdtAbsenceTo, "YYYY/MM/DD 00:00:00.000")
  strReportStartDate = Format(mdtAbsenceFrom, "YYYY/MM/DD")
  strReportEndDate = Format(mdtAbsenceTo, "YYYY/MM/DD")

  ' Merge the absence records if the continuous field is defined.
  If Not Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECONTINUOUS)) = "0" Then
    sSQL = "EXECUTE sp_ASR_Bradford_MergeAbsences '" + strReportStartDate + "','" + strReportEndDate + "','" + mstrTempTableName + "'"
    mclsData.ExecuteSql sSQL
  End If

  ' Delete unwanted absences from the table.
  sSQL = "EXECUTE sp_ASR_Bradford_DeleteAbsences '" + strReportStartDate + "','" + strReportEndDate + "'," + IIf(mblnOmitBeforeStart, "1,", "0,") + IIf(mblnOmitAfterEnd, "1,'", "0,'") + mstrTempTableName + "'"
  mclsData.ExecuteSql sSQL

  ' Calculate the included durations for the absences.
  sSQL = "EXECUTE sp_ASR_Bradford_CalculateDurations '" + strReportStartDate + "','" + strReportEndDate + "','" + mstrTempTableName + "'"
  mclsData.ExecuteSql sSQL

  ' Remove absences that are below the required Bradford Factor
  If mbMinBradford Then
    sSQL = "DELETE FROM " & mstrTempTableName & " WHERE personnel_id IN (SELECT personnel_id FROM " & mstrTempTableName _
      & " GROUP BY personnel_id HAVING((count(duration)*count(duration))*sum(duration)) < " + Str(mlngMinBradfordAmount) + ")"
    mclsData.ExecuteSql sSQL
  End If

  CalculateBradfordFactors = True
  Exit Function
  
CalculateBradfordFactors_ERROR:
  
  CalculateBradfordFactors = "Error while checking calculating Bradford factors." & vbNewLine & "(" & Err.Description & ")"
  CalculateBradfordFactors = False
  
End Function

Public Function GenerateSQLBradford() As Boolean

    ' NOTE: Checks are made elsewhere to ensure that from and to dates are not blank
    ' NOTE: Put in some code to handle blank end dates (do we include as an option on the main screen ?)

    On Error GoTo GenerateSQLBradford_ERROR

    Dim strAbsenceStartField As String
    Dim strAbsenceEndField As String
    Dim strAbsenceType As String
    Dim strReportStartDate As String
    Dim strReportEndDate As String
    Dim iCount As Integer

    ' Get the absence start/end field details
'    strAbsenceStartField = mstrRealSource + "." + datGeneral.GetColumnName(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE)))
'    strAbsenceEndField = mstrRealSource + "." + datGeneral.GetColumnName(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE)))
    strAbsenceType = mstrAbsenceRealSource + "." + datGeneral.GetColumnName(Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE)))

    ' Convert start and end dates of reporting period
'    strReportStartDate = "CONVERT(datetime, '" & Format(mdtAbsenceFrom, "MM/dd/yyyy") + "')"
'    strReportEndDate = "CONVERT(datetime, '" & Format(mdtAbsenceTo, "MM/dd/yyyy") + "')"

    ' Add in the absnece records to include
'    Select Case pcRangeType

        ' Start Inside reporting period, and end inside it.
'        Case "StartInsideEndInside"
'            mstrSQLWhere = IIf(mstrSQLWhere = vbNullString, "WHERE ", mstrSQLWhere + " AND ")
'            mstrSQLWhere = mstrSQLWhere + strAbsenceStartField + " >= " + strReportStartDate
'            mstrSQLWhere = mstrSQLWhere + " AND " + strAbsenceEndField + " <= " + strReportEndDate

        ' Starting before reporting period, but ending inside it.
'        Case "StartBeforeEndInside"
'            mstrSQLWhere = IIf(mstrSQLWhere = vbNullString, "WHERE ", mstrSQLWhere + " AND ")
'            mstrSQLWhere = mstrSQLWhere + strAbsenceEndField + " >= " + strReportStartDate
'            mstrSQLWhere = mstrSQLWhere + " AND " + strAbsenceEndField + " <= " + strReportEndDate

        ' Starting before reporting period and ending outside it.
'        Case "StartBeforeEndOutside"
'            mstrSQLWhere = IIf(mstrSQLWhere = vbNullString, "WHERE ", mstrSQLWhere + " AND ")
'            mstrSQLWhere = mstrSQLWhere + strAbsenceStartField + " <= " + strReportEndDate
'            mstrSQLWhere = mstrSQLWhere + " AND " + strAbsenceEndField + " >= " + strReportStartDate

        ' Starting inside the reporting period, but ending outside it.
'        Case "StartInsideEndOutside"
'            mstrSQLWhere = IIf(mstrSQLWhere = vbNullString, "WHERE ", mstrSQLWhere + " AND ")
'            mstrSQLWhere = mstrSQLWhere + strAbsenceStartField + " >= " + strReportStartDate
'            mstrSQLWhere = mstrSQLWhere + " AND " + strAbsenceStartField + " <= " + strReportEndDate

'    End Select

    ' Add the different reason types
    If msAbsenceTypes <> vbNullString Then
      mstrSQLWhere = _
        IIf(mstrSQLWhere = vbNullString, "WHERE (", mstrSQLWhere & " AND (") & "UPPER(" & strAbsenceType & ") IN (" & _
        msAbsenceTypes & "))"
    End If
    
'    If Not pcIncludeTypes = vbNullString Then
'      mstrSQLWhere = IIf(mstrSQLWhere = vbNullString, "WHERE ", mstrSQLWhere + " AND ")
'      mstrSQLWhere = mstrSQLWhere + "(CharIndex(Upper(" + strAbsenceType + ")," + pcIncludeTypes + ") > 0)"
'    End If

    ' Add the ID to the select string
    ' This is needed to re-calculate the duration amounts
    mstrSQLSelect = mstrSQLSelect + "," + mstrSQLFrom + ".ID AS 'Personnel_ID'," + mstrAbsenceRealSource + ".ID as 'Absence_ID'"

    ' Redimension arrays (to handle the ID fields (Personel/absnece)
    ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), UBound(mvarColDetails, 2) + 2)

    'Personel ID
    mvarColDetails(0, UBound(mvarColDetails, 2) - 1) = "Personnel_ID"
    mvarColDetails(1, UBound(mvarColDetails, 2) - 1) = 99
    mvarColDetails(2, UBound(mvarColDetails, 2) - 1) = 0
    mvarColDetails(3, UBound(mvarColDetails, 2) - 1) = False
    mvarColDetails(4, UBound(mvarColDetails, 2) - 1) = False
    mvarColDetails(5, UBound(mvarColDetails, 2) - 1) = False
    mvarColDetails(6, UBound(mvarColDetails, 2) - 1) = False
    mvarColDetails(7, UBound(mvarColDetails, 2) - 1) = True
    mvarColDetails(8, UBound(mvarColDetails, 2) - 1) = False
    mvarColDetails(9, UBound(mvarColDetails, 2) - 1) = True
    mvarColDetails(10, UBound(mvarColDetails, 2) - 1) = False
    mvarColDetails(11, UBound(mvarColDetails, 2) - 1) = ""
    mvarColDetails(12, UBound(mvarColDetails, 2) - 1) = -1
    mvarColDetails(13, UBound(mvarColDetails, 2) - 1) = "C"
    mvarColDetails(14, UBound(mvarColDetails, 2) - 1) = 0
    mvarColDetails(15, UBound(mvarColDetails, 2) - 1) = ""
    mvarColDetails(16, UBound(mvarColDetails, 2) - 1) = ""
    mvarColDetails(17, UBound(mvarColDetails, 2) - 1) = False
    mvarColDetails(18, UBound(mvarColDetails, 2) - 1) = False
    mvarColDetails(19, UBound(mvarColDetails, 2) - 1) = True        ' Is column hidden
    mvarColDetails(20, UBound(mvarColDetails, 2) - 1) = False
    mvarColDetails(21, UBound(mvarColDetails, 2) - 1) = False

    'Absence ID
    mvarColDetails(0, UBound(mvarColDetails, 2)) = "Absence_ID"
    mvarColDetails(1, UBound(mvarColDetails, 2)) = 99
    mvarColDetails(2, UBound(mvarColDetails, 2)) = 0
    mvarColDetails(3, UBound(mvarColDetails, 2)) = False
    mvarColDetails(4, UBound(mvarColDetails, 2)) = False
    mvarColDetails(5, UBound(mvarColDetails, 2)) = False
    mvarColDetails(6, UBound(mvarColDetails, 2)) = False
    mvarColDetails(7, UBound(mvarColDetails, 2)) = True
    mvarColDetails(8, UBound(mvarColDetails, 2)) = False
    mvarColDetails(9, UBound(mvarColDetails, 2)) = True
    mvarColDetails(10, UBound(mvarColDetails, 2)) = False
    mvarColDetails(11, UBound(mvarColDetails, 2)) = ""
    mvarColDetails(12, UBound(mvarColDetails, 2)) = -1
    mvarColDetails(13, UBound(mvarColDetails, 2)) = "C"
    mvarColDetails(14, UBound(mvarColDetails, 2)) = 0
    mvarColDetails(15, UBound(mvarColDetails, 2)) = ""
    mvarColDetails(16, UBound(mvarColDetails, 2)) = ""
    mvarColDetails(17, UBound(mvarColDetails, 2)) = False
    mvarColDetails(18, UBound(mvarColDetails, 2)) = False
    mvarColDetails(19, UBound(mvarColDetails, 2)) = True        ' Is column hidden
    mvarColDetails(20, UBound(mvarColDetails, 2)) = False
    mvarColDetails(21, UBound(mvarColDetails, 2)) = False

    ReDim Preserve mvarSortOrder(2, UBound(mvarSortOrder, 2) + 1)
    mvarSortOrder(1, UBound(mvarSortOrder, 2)) = -1
    mvarSortOrder(2, UBound(mvarSortOrder, 2)) = "Asc"

    ' All done correctly
    GenerateSQLBradford = True
    Exit Function


GenerateSQLBradford_ERROR:
  
    GenerateSQLBradford = False
    mstrErrorString = "Error in GenerateSQLBradford." & vbNewLine & Err.Description

End Function

Private Function GetBradfordReportDefinition() As Boolean

  '(mdtAbsenceFrom As Date, mdtAbsenceTo As Date) As Boolean
  
  ' Purpose : This function retrieves the basic definition details
  '           and stores it in module level variables
  
  On Error GoTo GetBradfordReportDefinition_ERROR

  Const strReportType As String = "BradfordFactor"

  'Dim rsTemp_Definition As Recordset
  'Dim strSQL As String
  'Dim strTable As String
  'Dim strColumn As String
  Dim pblnChkPrintFilterHeader As Boolean
  Dim strExprId As Variant
  
    
    
  'Dim objBradford As clsAbsenceBreakdown
  'Dim rsType As Recordset
  'Dim strSQL As String
  'Dim strType As String
  'Dim lngID As Long
  
  
'  strSQL = "SELECT * " & _
'           "FROM " & gsAbsenceTypeTableName & " " & _
'           "ORDER BY " & gsAbsenceTypeTypeColumnName
'  Set rsType = datGeneral.GetReadOnlyRecords(strSQL)
'
'  Set objBradford = New clsAbsenceBreakdown
'  objBradford.ReportType = strReportType
'
'  msAbsenceTypes = vbNullString
'  Do Until rsType.EOF
'
'    strType = rsType.Fields(gsAbsenceTypeTypeColumnName).Value
'    If objBradford.CheckIfAbsenceTypeSelected(strType) = True Then
'      msAbsenceTypes = _
'        IIf(msAbsenceTypes <> vbNullString, msAbsenceTypes & ", ", "") & _
'        "'" & Replace(strType, "'", "''") & "'"
'    End If
'
'    rsType.MoveNext
'
'  Loop
'  rsType.Close
'  Set rsType = Nothing
  
  If msAbsenceTypes <> vbNullString Then
    msAbsenceTypes = UCase(msAbsenceTypes)
  End If
  
  
  'If mblnCustomDates Then
  '  mdtAbsenceTo = datGeneral.GetValueForRecordIndependantCalc(mlngStartDateExprID)
  '  mdtAbsenceFrom = datGeneral.GetValueForRecordIndependantCalc(mlngEndDateExprID)
  'Else
  '  mdtAbsenceFrom = DateAdd("d", Day(Date) * -1, Date)
  '  mdtAbsenceTo = DateAdd("d", 1, DateAdd("yyyy", -1, mdtReportEndDate))
  'End If
  
  
  'Set objBradford = Nothing
  
  
  If pblnChkPrintFilterHeader = True Then
    If mlngCustomReportsFilterID > 0 Then
        'Add a filter name
        strExprId = GetExprField(mlngCustomReportsFilterID, "NAME")
        strExprId = " (Base Table Filter: " & strExprId & ")"
    ElseIf mlngCustomReportsPickListID > 0 Then
        'Add a picklist name
        strExprId = GetPickListField(mlngCustomReportsPickListID, "NAME")
        strExprId = " (Base Table Picklist: " & strExprId & ")"
        Else
            'Indicate no picklist or filter was selected
            strExprId = " (No Picklist or Filter Selected)"
    End If
    
    'Append the resulting string expression to the grid header
    mstrCustomReportsName = "Bradford Factor Report (" + Format(mdtAbsenceFrom, DateFormat) + _
        " - " + Format(mdtAbsenceTo, DateFormat) & ")" & strExprId
  Else
    'Set the grid header with no picklist/filter information
    mstrCustomReportsName = "Bradford Factor Report (" + Format(mdtAbsenceFrom, DateFormat) + _
        " - " + Format(mdtAbsenceTo, DateFormat) & ")"
  End If
              
  mstrCustomReportsDescription = mstrCustomReportsName
  mlngCustomReportsBaseTable = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE))
  mstrCustomReportsBaseTableName = datGeneral.GetTableName(mlngCustomReportsBaseTable)
  '    mlngCustomReportsAllRecords = !AllRecords
  '    mlngCustomReportsPickListID = !PickListID
  '    mlngCustomReportsFilterID = !Filter
  mlngCustomReportsParent1Table = 0
  mstrCustomReportsParent1TableName = ""
  mlngCustomReportsParent1FilterID = 0
  mlngCustomReportsParent2Table = 0
  mstrCustomReportsParent2TableName = ""
  mlngCustomReportsParent2FilterID = 0
  
  '****************************************
  '    mlngCustomReportsChildTable = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE))
  '    mstrCustomReportsChildTableName = datGeneral.GetTableName(mlngCustomReportsChildTable)
  '    mlngCustomReportsChildFilterID = 0
  '    mlngCustomReportsChildMaxRecords = 0
  
  ReDim Preserve mvarChildTables(5, 0)
  
  mvarChildTables(0, 0) = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETABLE))  'Childs Table ID
  mvarChildTables(1, 0) = 0  'Childs Filter ID (if any)
  mvarChildTables(2, 0) = 0  'Number of records to take from child
  mvarChildTables(3, 0) = datGeneral.GetTableName(mvarChildTables(0, 0))  'Child Table Name
  mvarChildTables(4, 0) = True  'Boolean - True if table is used, False if not
  mvarChildTables(5, 0) = 0

  miChildTablesCount = 1
  '****************************************
  
'  mblnCustomReportsSummaryReport = False
'  mblnCustomReportsPrintFilterHeader = False
'  mintCustomReportsDefaultOutput = 0
'  mintCustomReportsDefaultExportTo = 0
'  mblnCustomReportsDefaultSave = False
'  mstrCustomReportsDefaultSaveAs = ""
'  mblnCustomReportsDefaultCloseApp = False
  
  mlngCustomReportsParent1AllRecords = True
  mlngCustomReportsParent1PickListID = 0
  mlngCustomReportsParent2AllRecords = True
  mlngCustomReportsParent2PickListID = 0
  
  'TM20020503 Fault 3837 - Automatically the definition owner as this is a bradford adhoc report.
  mbDefinitionOwner = True
  
  mlngCustomReportsAllRecords = 0
  mlngCustomReportsPickListID = 0
  mlngCustomReportsFilterID = 0
  
  Select Case mstrPicklistFilterType
  Case "A"
    mlngCustomReportsAllRecords = 1
  Case "P"
    mlngCustomReportsPickListID = mlngPicklistFilterID
  Case "F"
    mlngCustomReportsFilterID = mlngPicklistFilterID
  End Select

  If Not IsRecordSelectionValid Then
    GetBradfordReportDefinition = False
    Exit Function
  End If


  GetBradfordReportDefinition = True
  
TidyAndExit:

    Exit Function

GetBradfordReportDefinition_ERROR:

  GetBradfordReportDefinition = False
  mstrErrorString = "Error whilst reading the Custom Report definition !" & vbNewLine & Err.Description
  Resume TidyAndExit

End Function

Private Function GetBradfordRecordSet() As Boolean

  ' Purpose : This function loads report details and sort details into
  '           arrays and leaves the details recordset reference there
  '           (dont remove it...used for summary info !)

  On Error GoTo GetBradfordRecordSet_ERROR
  
  'Dim strTempSQL As String
  Dim intTemp As Integer
  'Dim prstCustomReportsSortOrder As Recordset
  Dim lngTableID As Long
  Dim iCount As Integer
  Dim lngColumnID As Long
  'Dim lcDataType As Integer

  Dim lbHideStaffNumber As Boolean

  ReDim aStrRequiredFields(15, 1) As String

  aStrRequiredFields(1, 1) = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_EMPLOYEENUMBER))
  aStrRequiredFields(2, 1) = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME))
  aStrRequiredFields(3, 1) = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FORENAME))
  aStrRequiredFields(4, 1) = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_DEPARTMENT))

  aStrRequiredFields(5, 1) = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCETYPE))
  aStrRequiredFields(6, 1) = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE))
  aStrRequiredFields(7, 1) = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTSESSION))
  aStrRequiredFields(8, 1) = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDDATE))
  aStrRequiredFields(9, 1) = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEENDSESSION))
  aStrRequiredFields(10, 1) = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEREASON))
  aStrRequiredFields(11, 1) = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCECONTINUOUS))
  aStrRequiredFields(12, 1) = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION))

  'This field is later recalculated for the included days
  aStrRequiredFields(13, 1) = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION))

  '****************************************************************************
  If mlngOrderByColumnID > 0 Then
    aStrRequiredFields(14, 1) = mlngOrderByColumnID
  Else
    aStrRequiredFields(14, 1) = -1
  End If
  
  If mlngGroupByColumnID > 0 Then
    aStrRequiredFields(15, 1) = mlngGroupByColumnID
  Else
    aStrRequiredFields(15, 1) = -1
  End If
  '****************************************************************************
  
  ' Allow the staff number to be undefined (Let system read the surname field)
  lbHideStaffNumber = False
  If aStrRequiredFields(1, 1) = "0" Then
    aStrRequiredFields(1, 1) = Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME))
    lbHideStaffNumber = True
  End If

  ' Allow the continuous field to be undefined (Let system read the absence reason)
  If aStrRequiredFields(11, 1) = "0" Then
    aStrRequiredFields(11, 1) = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCEDURATION))
  End If

  ' Ensure that module setup has been run
  For iCount = 1 To UBound(aStrRequiredFields, 1)
    If aStrRequiredFields(iCount, 1) = "0" Then
      GetBradfordRecordSet = False
      mstrErrorString = "Module setup has not been completed."
      Exit Function
    End If
  Next iCount

  mblnCustomReportsSummaryReport = (Not mbDisplayBradfordDetail)
  
  ' Load the field list
  For iCount = 1 To UBound(aStrRequiredFields, 1)
  
    If aStrRequiredFields(iCount, 1) <> -1 Then
      intTemp = UBound(mvarColDetails, 2) + 1
      ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), intTemp)
     
      lngColumnID = CLng(aStrRequiredFields(iCount, 1))
      lngTableID = GetTableIDFromColumn(lngColumnID)
  
      ' Specify the column names and whether they are visible or not
      Select Case intTemp
        Case 1
          mvarColDetails(0, intTemp) = "Staff_No"
          mvarColDetails(19, intTemp) = lbHideStaffNumber
        Case 2
          mvarColDetails(0, intTemp) = "Surname"
          mvarColDetails(19, intTemp) = False
        Case 3
          mvarColDetails(0, intTemp) = "Forenames"
          mvarColDetails(19, intTemp) = False
        Case 4
          mvarColDetails(0, intTemp) = "Department"
          mvarColDetails(19, intTemp) = False
        Case 5
          mvarColDetails(0, intTemp) = "Type"
          mvarColDetails(19, intTemp) = Not mbDisplayBradfordDetail
        Case 6
          mvarColDetails(0, intTemp) = "Start_Date"
          mvarColDetails(19, intTemp) = Not mbDisplayBradfordDetail
        Case 7
          mvarColDetails(0, intTemp) = "Start_Session"
          mvarColDetails(19, intTemp) = Not mbDisplayBradfordDetail
        Case 8
          mvarColDetails(0, intTemp) = "End_Date"
          mvarColDetails(19, intTemp) = Not mbDisplayBradfordDetail
        Case 9
          mvarColDetails(0, intTemp) = "End_Session"
          mvarColDetails(19, intTemp) = Not mbDisplayBradfordDetail
        Case 10
          If mbDisplayBradfordDetail Then
            mvarColDetails(0, intTemp) = "Reason"
            mvarColDetails(19, intTemp) = False
          Else
            mvarColDetails(0, intTemp) = "Summary Info"
            mvarColDetails(19, intTemp) = False
          End If
        Case 11
          mvarColDetails(0, intTemp) = "Continuous"
          mvarColDetails(19, intTemp) = Not mbDisplayBradfordDetail
        Case 12
          mvarColDetails(0, intTemp) = "Duration"
          mvarColDetails(19, intTemp) = False
        Case 13
          mvarColDetails(0, intTemp) = "Included_Days"
          mvarColDetails(19, intTemp) = False
        
        '**********************************************************************
        Case 14
          mvarColDetails(0, intTemp) = "Order_1"
          mvarColDetails(19, intTemp) = True
            
        Case 15
          mvarColDetails(0, intTemp) = "Order_2"
          mvarColDetails(19, intTemp) = True
        '**********************************************************************
        
        Case Else
          'MH20020521 Fault 3820
          'mvarColDetails(0, intTemp) = datGeneral.GetColumnName(lngColumnID)
          If lngTableID = mlngCustomReportsBaseTable Then
            'Personnel
            mvarColDetails(0, intTemp) = mstrSQLFrom & "." & datGeneral.GetColumnName(lngColumnID)
          Else
            'Absence
            mvarColDetails(0, intTemp) = mstrRealSource & "." & datGeneral.GetColumnName(lngColumnID)
          End If
          mvarColDetails(19, intTemp) = True
            
      End Select
  
      mvarColDetails(1, intTemp) = 99
      mvarColDetails(2, intTemp) = IIf(intTemp = 12 Or intTemp = 13, 1, 0)        'Decimals
      'JDM - 02/07/01 - Fault 2144 - Needs to know if we're numeric or not.
      mvarColDetails(3, intTemp) = IIf(datGeneral.GetDataType(lngTableID, lngColumnID) = 2, True, False) 'Is Numeric
      mvarColDetails(4, intTemp) = False                                          'Average
      mvarColDetails(5, intTemp) = IIf(intTemp = 12 Or intTemp = 13, True, False)  'Count
      mvarColDetails(6, intTemp) = IIf(intTemp = 12 Or intTemp = 13, True, False) 'Total
      mvarColDetails(7, intTemp) = False                  'Break on change
      mvarColDetails(8, intTemp) = False                                          'Page break on change
      If mblnCustomReportsSummaryReport Then
        mvarColDetails(9, intTemp) = True                                           'Value on change
      Else
        mvarColDetails(9, intTemp) = False                                           'Value on change
      End If
      mvarColDetails(10, intTemp) = IIf(intTemp < 5 And mbBradfordSRV, True, False)                 'Suppress repeated values
      mvarColDetails(11, intTemp) = ""
      mvarColDetails(12, intTemp) = lngColumnID
      
      ' Set the expression/column type of this column
      mvarColDetails(13, intTemp) = "C"
  
      mvarColDetails(14, intTemp) = lngTableID
      mvarColDetails(15, intTemp) = datGeneral.GetTableName(CLng(lngTableID))
      
      If mvarColDetails(13, intTemp) = "C" Then
        mvarColDetails(16, intTemp) = datGeneral.GetColumnName(CLng(mvarColDetails(12, intTemp)))
      Else
        mvarColDetails(16, intTemp) = ""
      
        'MH20010307
        Dim objExpr As DataMgr.clsExprExpression
        Set objExpr = New clsExprExpression
        objExpr.ExpressionID = CLng(mvarColDetails(12, intTemp))
        objExpr.ConstructExpression
        Set objExpr = Nothing
      
      End If
      
      mvarColDetails(17, intTemp) = datGeneral.DateColumn("C", lngTableID, lngColumnID)   '??? - check these out 22/03/01
      mvarColDetails(18, intTemp) = datGeneral.BitColumn("C", lngTableID, lngColumnID)
  
      mvarColDetails(22, intTemp) = datGeneral.DoesColumnUseSeparators(lngColumnID)       'Does this column use 1000 separators?
      
      'Adjust the size of the field if digit separator is used
      If mvarColDetails(22, intTemp) Then
        mvarColDetails(1, intTemp) = mvarColDetails(1, intTemp) + Int((mvarColDetails(1, intTemp) - mvarColDetails(2, intTemp)) / 3)
      End If
    
    End If
    
  Next iCount

  ' Get those columns defined as a SortOrder and load into array
  ReDim mvarSortOrder(2, 3)

  'Employee surname
  mvarSortOrder(1, 1) = mstrOrderByColumn  'Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_SURNAME))
  mvarSortOrder(2, 1) = "Asc"

  'Employee forename
  mvarSortOrder(1, 2) = mstrGroupByColumn  'Val(GetModuleParameter(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_FORENAME))
  mvarSortOrder(2, 2) = "Asc"

'    ' Absence start date
'    mvarSortOrder(1, 3) = "Start_Date" 'Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_ABSENCESTARTDATE))"
'    mvarSortOrder(2, 3) = "Asc"

  '' Force duration and included days to be numeric format in Excel
  'iCount = 11 - IIf(lbHideStaffNumber = True, 1, 0)
  'mstrExcelFormats(iCount) = "0.0"
  'mstrExcelFormats(iCount + 1) = "0.0"

  GetBradfordRecordSet = True
  Exit Function
  
GetBradfordRecordSet_ERROR:
  
  GetBradfordRecordSet = False
  mstrErrorString = "Error whilst retrieving the details recordsets'." & vbNewLine & Err.Description
  
End Function

Public Function RunBradfordReport(pstrPicklistFilterIDs As String) As Boolean

'(mdtAbsenceFrom As Date, mdtAbsenceTo As Date, pastrIncludeTypes() As String _
  , mblnOmitBeforeStart As Boolean, mblnOmitAfterEnd As Boolean, pbSRV As Boolean, pbShowTotals As Boolean _
  , pbShowCount As Boolean, pbShowWorkings As Boolean _
  , pgblnBatchMode As Boolean, plngPersonnelID As Long _
  , pstrOrderBy As String, pstrGroupBy As String _
  , pbOrder1Asc As Boolean, pbOrder2Asc As Boolean) As Boolean

  ' Purpose : This function is called from frmDefsel and Batch Jobs.
  '           It is the main function which runs the report.

  On Error GoTo RunBradfordReport_ERROR
  
  Dim fOK As Boolean
  
  fOK = True
  
  Screen.MousePointer = vbHourglass
  
  'Set this as a Bradford Index report
  mbIsBradfordIndexReport = True
  mstrPicklistFilterIDs = pstrPicklistFilterIDs
  
  If fOK And gblnBatchMode = True Then
    fOK = GetReportConfig
  End If
    
  If fOK Then fOK = GetBradfordReportDefinition '(mdtAbsenceFrom, mdtAbsenceTo)
  
  gobjEventLog.AddHeader eltStandardReport, "Bradford Factor Report"
  
  With gobjProgress
    '.AviFile = App.Path & "\videos\report.avi"
    .AVI = dbText
    .MainCaption = "Bradford Report"
    If Not gblnBatchMode Then
      .NumberOfBars = 1
      .Caption = "Bradford Factor Report"
      .Time = False
      .Cancel = True
      .Bar1Caption = "Bradford Factor Report"
      .OpenProgress
    Else
      .Bar2Caption = "Bradford Factor Report"
    End If
  End With

  'Check that from and to dates have been entered
  If IsEmpty(mdtAbsenceFrom) Or IsEmpty(mdtAbsenceTo) Then
    fOK = False
    mstrErrorString = "Error Whilst Running Bradford Factor Report." & vbNewLine & "Report range dates not specified"
  End If

  If fOK Then fOK = GetBradfordRecordSet
  If fOK Then fOK = GenerateSQL
  If fOK Then fOK = GenerateSQLBradford
  If fOK Then fOK = AddTempTableToSQL
  If fOK Then fOK = MergeSQLStrings
  If fOK Then fOK = UDFFunctions(mastrUDFsRequired, True)
  If fOK Then fOK = ExecuteSql
  If fOK Then fOK = UDFFunctions(mastrUDFsRequired, True)
  If fOK Then fOK = CalculateBradfordFactors
  If fOK Then fOK = CheckRecordSet
  If fOK Then fOK = PopulateGrid
  
  If fOK Then
    If gblnBatchMode Or Not mblnPreviewOnScreen Then
      fOK = OutputReport(False)
  
    ElseIf mblnPreviewOnScreen Then
      Screen.MousePointer = vbDefault
      mfrmOutput.Caption = "Bradford Factor Report"
      mfrmOutput.BaseTable = mstrCustomReportsBaseTableName
      mfrmOutput.Parent = Me
      mfrmOutput.HelpContextID = 1011
      mfrmOutput.Show vbModal
    End If
  End If
  
  Set mfrmOutput = Nothing
    
  Unload frmOutputOptions
  Set frmOutputOptions = Nothing
  
  mblnUserCancelled = _
    (mblnUserCancelled Or gobjProgress.Cancelled Or _
    (InStr(LCase(mstrErrorString), "cancelled by user") > 0))
  
  If mblnNoRecords Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful
    gobjEventLog.AddDetailEntry mstrErrorString
    mstrErrorString = "Completed successfully." & vbNewLine & mstrErrorString
    fOK = True
  ElseIf fOK Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful
    mstrErrorString = "Completed successfully."
  ElseIf mblnUserCancelled Then
    gobjEventLog.ChangeHeaderStatus elsCancelled
    mstrErrorString = "Cancelled by user."
  Else
    'Only details records for failures !
    gobjEventLog.AddDetailEntry mstrErrorString
    gobjEventLog.ChangeHeaderStatus elsFailed
    mstrErrorString = "Failed." & vbNewLine & vbNewLine & mstrErrorString
  End If
  
  mstrErrorString = "Bradford Factor Report " & mstrErrorString
  
  If Not gblnBatchMode Then
    If gobjProgress.Visible Then gobjProgress.CloseProgress
    If (fOK = False) Or (mblnNoRecords = True) Then COAMsgBox mstrErrorString, IIf(fOK, vbInformation, vbExclamation) + vbOKOnly, "Bradford Factor Report"
  End If
  
  Screen.MousePointer = vbDefault
  
  If fOK = False Then RunBradfordReport = False Else RunBradfordReport = True
  If fOK = True Then fOK = ClearUp Else ClearUp
  
  Exit Function
  
RunBradfordReport_ERROR:

  fOK = False
  RunBradfordReport = False
  mstrErrorString = "Error Whilst Running Bradford Factor Report." & vbNewLine & Err.Description
  Resume Next
  
End Function

Public Property Let FilterID(ByVal lNewValue As Long)
  mlngCustomReportsFilterID = lNewValue
End Property
Public Property Let PicklistID(ByVal lNewValue As Long)
  mlngCustomReportsPickListID = lNewValue
End Property


Public Function OutputReport(blnPrompt As Boolean) As Boolean

  Dim objOutput As clsOutputRun
  Dim lngIndex As Long
  
  Dim iRow As Long
  Dim icol As Integer
  Dim bm As Variant
  Dim sAddString As String
  Dim sBreakValue As String
  Dim iVisCount As Integer
  Dim strReportBreakName As String
  Dim blnBreakCheck As Boolean
  Dim iPageBreakCount As Integer
  
  Dim strReportType As String
  Dim strMBTitle As String
  
  Dim fOK As Boolean
  
  fOK = True
  strReportType = IIf(mbIsBradfordIndexReport, "Bradford Factor", "Custom Report: '" & mstrCustomReportsName & "'")
  strMBTitle = IIf(mbIsBradfordIndexReport, "Bradford Factor", "Custom Report")
  
  Set objOutput = New clsOutputRun

  If mbIsBradfordIndexReport Then
    objOutput.ShowFormats True, False, True, True, True, True, False
  End If

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
      mblnCustomReportsPrintFilterHeader, _
      mblnPreviewOnScreen, _
      mstrOutputTitlePage, _
      mstrOutputReportPackTitle, _
      mstrOutputOverrideFilter, _
      mblnOutputTOC, _
      mblnOutputCoverSheet, _
      mlngOverrideFilterID, _
      mblnOutputRetainPivotOrChart, _
      mlngOriginalOutputFormat) Then

    If Not gblnBatchMode Then
      If mbIsBradfordIndexReport Then
        objOutput.OpenProgress "Bradford Factor", vbNullString, 1
      Else
        objOutput.OpenProgress "Custom Report", mstrCustomReportsName, 1
      End If
    End If

    If gblnReportPackMode Then
      objOutput.IndicatorColumn = (mlngOutputFormat = fmtExcelWorksheet) And (mlngOriginalOutputFormat = fmtExcelchart Or mlngOriginalOutputFormat = fmtExcelPivotTable) And Not mblnCustomReportsSummaryReport
    Else
      objOutput.IndicatorColumn = (mlngOutputFormat = fmtExcelchart Or mlngOutputFormat = fmtExcelPivotTable) And Not mblnCustomReportsSummaryReport
    End If
    
    objOutput.SizeColumnsIndependently = True
    If objOutput.GetFile Then

      objOutput.PageTitles = mblnReportHasPageBreak
      
      If Not gblnBatchMode Then
        gobjProgress.Bar1MaxValue = mfrmOutput.grdOutput.Rows + 1
      End If

      If mblnReportHasPageBreak Then
        iVisCount = 0
        mfrmOutput.grdBreak.Columns.RemoveAll
        
        With mfrmOutput.grdOutput
      
          If mblnReportHasSummaryInfo Then
            objOutput.AddColumn "Summary Info", sqlVarChar, 0
            mfrmOutput.grdBreak.Columns.Add iVisCount
            mfrmOutput.grdBreak.Columns(iVisCount).Name = "Summary Info"
            mfrmOutput.grdBreak.Columns(iVisCount).Caption = mfrmOutput.grdBreak.Columns(iVisCount).Name
            iVisCount = iVisCount + 1
          End If
      
          For lngIndex = 1 To UBound(mvarColDetails, 2)

            If mvarColDetails(19, lngIndex) = False Then
              If (mvarColDetails(24, lngIndex - 1) <> True) Then

                If mvarColDetails(17, lngIndex) And mvarColDetails(24, lngIndex) = False Then
                  objOutput.AddColumn _
                    CStr(mvarColDetails(0, lngIndex)), sqlDate, _
                    CLng(mvarColDetails(2, lngIndex)), (mvarColDetails(22, lngIndex) = True)
                ElseIf mvarColDetails(3, lngIndex) And mvarColDetails(24, lngIndex) = False Then
                  objOutput.AddColumn _
                    CStr(mvarColDetails(0, lngIndex)), sqlNumeric, _
                    CLng(mvarColDetails(2, lngIndex)), (mvarColDetails(22, lngIndex) = True)
                Else
                  objOutput.AddColumn _
                    CStr(mvarColDetails(0, lngIndex)), sqlVarChar, _
                    CLng(mvarColDetails(2, lngIndex)), (mvarColDetails(22, lngIndex) = True)
                End If
                
                
                mfrmOutput.grdBreak.Columns.Add iVisCount
                mfrmOutput.grdBreak.Columns(iVisCount).Name = CStr(mvarColDetails(0, lngIndex))
                mfrmOutput.grdBreak.Columns(iVisCount).Caption = mfrmOutput.grdBreak.Columns(iVisCount).Name
          
                If (mvarColDetails(3, lngIndex)) Then
                  mfrmOutput.grdBreak.Columns(iVisCount).Alignment = ssCaptionAlignmentRight
                ElseIf (mvarColDetails(18, lngIndex)) Then
                  mfrmOutput.grdBreak.Columns(iVisCount).Alignment = ssCaptionAlignmentCenter
                Else
                  mfrmOutput.grdBreak.Columns(iVisCount).Alignment = ssCaptionAlignmentLeft
                End If
                
                mfrmOutput.grdBreak.Columns(iVisCount).Visible = True
                iVisCount = iVisCount + 1
              End If
            End If
          Next
          
          'clear grid
          mfrmOutput.grdBreak.RemoveAll
      
'          For iRow = 0 To .Rows - 1 Step 1
          For iRow = 0 To .Rows '- 1 Step 1
            
            mblnUserCancelled = gobjProgress.Cancelled
            If mblnUserCancelled Then
              If blnPrompt Then
                gobjProgress.CloseProgress
                COAMsgBox strReportType & " Output cancelled by user.", vbExclamation + vbOKOnly, strMBTitle
              End If
              
              OutputReport = False
              If Not objOutput Is Nothing Then objOutput.ClearUp
              Set objOutput = Nothing
              Exit Function
            Else
              If Not gblnBatchMode Then
                gobjProgress.UpdateProgress gblnBatchMode
              End If
            End If
            
            bm = .AddItemBookmark(iRow)
            
'            If iRow = (.Rows - 1) Then
            If iRow = .Rows Then
              'iPageBreakCount = iPageBreakCount + 1
              If mblnReportHasSummaryInfo Then
                sBreakValue = "Grand Totals"
              Else
                'sBreakValue = mvarPageBreak(iRow - 1)
                sBreakValue = mvarPageBreak(iRow)
              End If
              
              'add a new page to the output class pass it the section of the grid.
              If mfrmOutput.grdBreak.Rows > 0 Then
                objOutput.AddPage Replace(mfrmOutput.grdOutput.Caption, "&&", "&"), sBreakValue, mblnOutputCoverSheet
                objOutput.DataGrid mfrmOutput.grdBreak
              End If
              
              sAddString = vbNullString
              sBreakValue = vbNullString
              mfrmOutput.grdBreak.RemoveAll
      
            ElseIf (.Columns(0).CellText(bm) = "*") And ((Not blnBreakCheck)) Then
              'iPageBreakCount = iPageBreakCount + 1
              sBreakValue = mvarPageBreak(iRow)

              'add a new page to the output class pass it the section of the grid.
              objOutput.AddPage Replace(mfrmOutput.grdOutput.Caption, "&&", "&"), sBreakValue, mblnOutputCoverSheet
              objOutput.DataGrid mfrmOutput.grdBreak
              
              blnBreakCheck = True
              sAddString = vbNullString
              sBreakValue = vbNullString
              mfrmOutput.grdBreak.RemoveAll
              
            ElseIf (.Columns(0).CellText(bm) <> "*") Then
              blnBreakCheck = False
              If objOutput.Format = fmtExcelPivotTable And _
                    ((.Columns(0).CellText(bm) = "*average*") Or _
                    (.Columns(0).CellText(bm) = "*count*") Or _
                    (.Columns(0).CellText(bm) = "*total*")) Then
              
              Else
                'loop through the columns build up a string to add to the break grid.
                For icol = 0 To (.Columns.Count - 1) Step 1
                  If .Columns(icol).Visible Then
                    sAddString = sAddString & .Columns(icol).CellText(bm) & vbTab
                  End If
                Next icol
                
                mfrmOutput.grdBreak.AddItem sAddString
                sAddString = vbNullString
              End If
              
            End If
          
            If gobjProgress.Cancelled Then
              If Not gblnBatchMode Then
                gobjProgress.CloseProgress
              End If
              mstrErrorString = "Cancelled by user"
              If Not objOutput Is Nothing Then objOutput.ClearUp
              Set objOutput = Nothing
              OutputReport = False
              Exit Function
            End If
          
          Next iRow
          
        End With
        
      Else
      
        If mblnReportHasSummaryInfo Then
          objOutput.AddColumn "Summary Info", sqlVarChar, 0, False
        End If

        'No page breaks in the report therefore just pass the entire grid to the output run class.
        For lngIndex = 1 To UBound(mvarColDetails, 2)
          If mvarColDetails(19, lngIndex) = False Then
            If mvarColDetails(17, lngIndex) And mvarColDetails(24, lngIndex) = False Then
              objOutput.AddColumn _
                CStr(mvarColDetails(0, lngIndex)), sqlDate, _
                CLng(mvarColDetails(2, lngIndex)), (mvarColDetails(22, lngIndex) = True)
            ElseIf mvarColDetails(3, lngIndex) And mvarColDetails(24, lngIndex) = False Then
              objOutput.AddColumn _
                CStr(mvarColDetails(0, lngIndex)), sqlNumeric, _
                CLng(mvarColDetails(2, lngIndex)), (mvarColDetails(22, lngIndex) = True)
            Else
              objOutput.AddColumn _
                CStr(mvarColDetails(0, lngIndex)), sqlVarChar, _
                CLng(mvarColDetails(2, lngIndex)), (mvarColDetails(22, lngIndex) = True)
            End If
          End If
        Next
      
        mblnUserCancelled = gobjProgress.Cancelled
        If mblnUserCancelled Then
          If blnPrompt Then
            gobjProgress.CloseProgress
            COAMsgBox strReportType & " Output cancelled by user.", vbExclamation + vbOKOnly, strMBTitle
          End If
          
          OutputReport = False
          If Not objOutput Is Nothing Then objOutput.ClearUp
          Set objOutput = Nothing
          Exit Function
        End If
        objOutput.AddPage Replace(mfrmOutput.grdOutput.Caption, "&&", "&"), IIf(mbIsBradfordIndexReport, "Bradford Factor", mstrCustomReportsBaseTableName), mblnOutputCoverSheet
        objOutput.UpdateProgressPerRow = Not gblnBatchMode
        objOutput.DataGrid mfrmOutput.grdOutput
      End If
    End If

    objOutput.Complete
    mblnUserCancelled = objOutput.UserCancelled
    mstrErrorString = objOutput.ErrorMessage
    fOK = (mstrErrorString = vbNullString)
  Else
    blnPrompt = (blnPrompt And Not objOutput.UserCancelled)
    mstrErrorString = objOutput.ErrorMessage
    fOK = (mstrErrorString = vbNullString)
  End If
    
  If blnPrompt Then
    gobjProgress.CloseProgress
    'Not ideal but the only way to prevent a runtime error was a doevents
    DoEvents
    
    If fOK Then
      COAMsgBox strReportType & " Output complete.", _
          vbInformation, strMBTitle
    Else
      COAMsgBox strReportType & " Output failed." & vbNewLine & vbNewLine & mstrErrorString, _
          vbExclamation, strMBTitle
    End If
  End If

  Set objOutput = Nothing
  OutputReport = fOK
  
End Function

Public Property Get PreviewOnScreen() As Boolean
  PreviewOnScreen = ((mstrErrorString = vbNullString And mblnPreviewOnScreen) And Not gblnBatchMode)
End Property

Private Function PopulateGrid_LoadRecords() As Boolean

  ' Purpose : Blimey ! This function does the actual work of populating the
  '           grid, calculating summary info, breaking, page breaking etc.
  '           Its a bit of a 'mare but it works ok. TRUE...TRUE

  On Error GoTo LoadRecords_ERROR
  
  Dim sAddString As String
  Dim vDisplayData As Variant
  'Dim vPreviousData As Variant
  Dim avColumns() As Variant
  Dim vValue As Variant
  Dim fBreak As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iColumnIndex As Integer
  Dim iOtherColumnIndex As Integer
  Dim bSuppress As Boolean
  Dim intColCounter As Integer
  Dim sBreakValue As String
  'Group With Next Column variables
  Dim intRowIndex_GW As Integer
  Dim intColIndex_GW As Integer
  Dim strGroupedValue As String
  Dim intGroupCount As Integer
  Dim blnHasGroupWithNext As Boolean
  Dim blnSkipped As Boolean
  Dim intSkippedIndex As Integer
  Dim lngRecordNo As Long
  Dim sDateFormat As String
  
  intRowIndex_GW = 0
  intColIndex_GW = 0
  strGroupedValue = vbNullString
  intGroupCount = 0
  blnHasGroupWithNext = False
  blnSkipped = False
  intSkippedIndex = 0
  
  
  'Variables for Suppress Repeated Values within Table.
  'Dim lngCurrentTableID As Long
  Dim lngCurrentRecordID As Long
  Dim bBaseRecordChanged As Boolean
  
  Dim tmpLogicValue As Variant

  ' Construct an array of the columns in the report. Basically this is an extension of the mvarColDetails array.
  ' Col 1 = TRUE if the column is used for breaking/paging on change.
  ' Col 2 = TRUE if the column is to be aggregated (average/count/total), else FALSE.
  ' Col 3 = last column value.
  ReDim avColumns(3, UBound(mvarColDetails, 2))
  For iLoop = 1 To UBound(mvarColDetails, 2)
    avColumns(1, iLoop) = mvarColDetails(7, iLoop) Or mvarColDetails(8, iLoop)
    avColumns(2, iLoop) = mvarColDetails(4, iLoop) Or mvarColDetails(5, iLoop) Or mvarColDetails(6, iLoop)
    avColumns(3, iLoop) = Null
  Next iLoop
 
  'Initialise the progress bar
  With gobjProgress
    If Not gblnBatchMode Then
      .Bar1MaxValue = mrstCustomReportsOutput.RecordCount / 100
    Else
      .ResetBar2
      .Bar2MaxValue = (mrstCustomReportsOutput.RecordCount + 1)
      .Bar2Caption = IIf(mbIsBradfordIndexReport, "Bradford Factor : ", "Custom Report : ") & mstrCustomReportsName
    End If
  End With
  
  With mfrmOutput
    'Ensure we are at the beginning of the output recordset
    mrstCustomReportsOutput.MoveFirst
    lngRecordNo = 0
    
    sDateFormat = DateFormat
    
    Do Until mrstCustomReportsOutput.EOF
      
      lngRecordNo = lngRecordNo + 1
      
      'Update the progress bar
      If lngRecordNo Mod 100 = 0 Then
        gobjProgress.UpdateProgress gblnBatchMode
        If gobjProgress.Cancelled Then
          mblnUserCancelled = True
          PopulateGrid_LoadRecords = False
          Exit Function
        End If
      End If
      
      'bRecordChanged used for repetition funcionality.
      If Not mbIsBradfordIndexReport Then
        If mrstCustomReportsOutput.Fields("?ID").Value <> lngCurrentRecordID Then
          bBaseRecordChanged = True
          lngCurrentRecordID = mrstCustomReportsOutput("?ID").Value
        Else
          bBaseRecordChanged = False
        End If
      End If
      
      'Offset the addstring by 2 columns (for pagebreak and summary info)
      sAddString = vbTab & "" & vbTab
      
      'Dont do summary info for first record (otherwise blank!),
      ' becuase we have nothing to summarise.
      If mrstCustomReportsOutput.AbsolutePosition > 1 Then

        'Put the values from the previous record in the column array.
        For iLoop = 1 To UBound(mvarColDetails, 2)
          avColumns(3, iLoop) = mvarColDetails(11, iLoop)
        Next iLoop

        'From last column in the order to first, check changes.
        For iLoop = UBound(mvarSortOrder, 2) To 1 Step -1
          'Find the column in the details array.
          iColumnIndex = 0
          For iLoop2 = 1 To UBound(mvarColDetails, 2)
            If (mvarColDetails(12, iLoop2) = mvarSortOrder(1, iLoop)) And _
              (mvarColDetails(13, iLoop2) = "C") Then

              iColumnIndex = iLoop2
              Exit For
            End If
          Next iLoop2

          If iColumnIndex > 0 Then
            If avColumns(1, iColumnIndex) Then  'has Break or Page on Change
              fBreak = False

              'The column breaks. Check if its changed.
              'if value is null and not numeric and not date and not bit then set to nullstring.
              If IsNull(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value) And _
                (Not mvarColDetails(3, iColumnIndex)) And _
                (Not mvarColDetails(17, iColumnIndex)) And _
                (Not mvarColDetails(18, iColumnIndex)) Then
                'Field value is null but a character data type, so set it to be "".
                vValue = ""

              Else
                'RH 01/02/01
                'Dates need to be formatted with yyyy for boc to work ok
                If mvarColDetails(17, iColumnIndex) Then 'Date
                  vValue = Left(Format(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value, sDateFormat), mvarColDetails(1, iColumnIndex))
                  
                ElseIf mvarColDetails(3, iColumnIndex) Then 'Numeric
                  vValue = Left(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value, mvarColDetails(1, iColumnIndex))
                  
                ElseIf mvarColDetails(18, iColumnIndex) Then 'Bit
                  If (mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value = True) Or (mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value = 1) Then vValue = "Y"
                  If (mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value = False) Or (mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value = 0) Then vValue = "N"
                  If IsNull(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value) Then vValue = ""

                Else 'Varchar
                  vValue = Left(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value, mvarColDetails(1, iColumnIndex))
                  
                End If

              End If
              
              'TM20020320 Fault 3675
              'Now that we store the formatted value in position (11) of the mcolDetails
              'Comparison made after adjusting the size of the field.
              If IsNull(vValue) Or IsNull(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value) Then
                fBreak = ("" <> mvarColDetails(11, iColumnIndex))
              Else
                If mvarColDetails(18, iColumnIndex) Then 'Bit
                  fBreak = (RTrim(LCase(vValue)) <> RTrim(LCase(mvarColDetails(11, iColumnIndex))))
                Else
                  'TM23112004 Fault 9072
                  fBreak = (RTrim(LCase(mrstCustomReportsOutput.Fields(iColumnIndex - 1).Value)) <> RTrim(LCase(mvarColDetails(11, iColumnIndex))))
                End If
              End If
              
              If mvarColDetails(8, iColumnIndex) Then
                sBreakValue = IIf(Len(mvarColDetails(11, iColumnIndex)) < 1, "<Empty>", mvarColDetails(11, iColumnIndex)) & IIf(Len(sBreakValue) > 0, " - ", "") & sBreakValue
              End If

              If Not fBreak Then
                'The value has not changed, but check if we need to do the summary due to another column changing.
                For iLoop2 = (iLoop - 1) To 1 Step -1
                  iOtherColumnIndex = 0
                  For iLoop3 = 1 To UBound(mvarColDetails, 2)
                    If (mvarColDetails(12, iLoop3) = mvarSortOrder(1, iLoop2)) And _
                      (mvarColDetails(13, iLoop3) = "C") Then

                      iOtherColumnIndex = iLoop3
                      Exit For
                    End If
                  Next iLoop3

                  If iOtherColumnIndex > 0 Then
                    If avColumns(1, iOtherColumnIndex) Then
                      ' The column breaks. Check if its changed.
                      If IsNull(mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value) And _
                        (Not mvarColDetails(3, iOtherColumnIndex)) And _
                        (Not mvarColDetails(17, iOtherColumnIndex)) And _
                        (Not mvarColDetails(18, iOtherColumnIndex)) Then
                        ' Field value is null but a character data type, so set it to be "".
                        vValue = ""
                        
                      ElseIf mvarColDetails(3, iOtherColumnIndex) Then 'Numeric
                        vValue = Left(mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value, mvarColDetails(1, iOtherColumnIndex))
                  
                      ElseIf mvarColDetails(18, iOtherColumnIndex) Then 'Bit
                        If (mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value = True) Or (mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value = 1) Then vValue = "Y"
                        If (mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value = False) Or (mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value = 0) Then vValue = "N"
                        If IsNull(mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value) Then vValue = ""
      
                      Else
                        vValue = Left(mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value, mvarColDetails(1, iOtherColumnIndex))
                  
                      End If
                      
                      If IsNull(vValue) Or IsNull(mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value) Then
                        fBreak = ("" <> mvarColDetails(11, iOtherColumnIndex))
                      Else
                        If mvarColDetails(18, iOtherColumnIndex) Then 'Bit
                          fBreak = (RTrim(LCase(vValue)) <> RTrim(LCase(mvarColDetails(11, iOtherColumnIndex))))
                        Else
                          'TM23112004 Fault 9072
                          fBreak = (RTrim(LCase(mrstCustomReportsOutput.Fields(iOtherColumnIndex - 1).Value)) <> RTrim(LCase(mvarColDetails(11, iOtherColumnIndex))))
                        End If
                      End If
                      
                      If (fBreak = True) Then
                        Exit For
                      End If

                    End If
                  End If
                Next iLoop2
              End If

              'Check for aggregate too, and only do summary info if its true.
              If fBreak Then
                PopulateGrid_DoSummaryInfo avColumns, iColumnIndex, iLoop
                
                'Do the page break ?
                If mvarColDetails(8, iColumnIndex) Then
                  mblnPageBreak = True
                  mblnReportHasPageBreak = True
                End If
              End If

            End If
          End If
        Next iLoop
      End If

      If mblnPageBreak Then
        mfrmOutput.grdOutput.AddItem "*"
        mintPageBreakRowIndex = mintPageBreakRowIndex + 1
        AddPageBreakValue mintPageBreakRowIndex - 1, sBreakValue
      End If
      mblnPageBreak = False
      sBreakValue = vbNullString

      intColCounter = 1
      'Loop thru each field, adding to the string to add to the grid
      For iLoop = 0 To (mrstCustomReportsOutput.Fields.Count - 1)
        
        intColCounter = intColCounter + 1
        
        'Only suppress values for new records in the Bradford Factor report
        bSuppress = IIf(mbIsBradfordIndexReport And fBreak, False, True)
        
        If (mvarColDetails(18, iLoop + 1)) Then 'Bit
          If (mrstCustomReportsOutput.Fields(iLoop).Value = "True") Or (mrstCustomReportsOutput.Fields(iLoop).Value = 1) Then tmpLogicValue = "Y"
          If (mrstCustomReportsOutput.Fields(iLoop).Value = "False") Or (mrstCustomReportsOutput.Fields(iLoop).Value = 0) Then tmpLogicValue = "N"
          If IsNull(mrstCustomReportsOutput.Fields(iLoop).Value) Then tmpLogicValue = ""
          
          ' Get the formatted data to display in the grid
          vDisplayData = PopulateGrid_FormatData(mrstCustomReportsOutput.Fields(iLoop).Name, tmpLogicValue, bSuppress, bBaseRecordChanged)
          
        Else
          ' Get the formatted data to display in the grid
          vDisplayData = PopulateGrid_FormatData(mrstCustomReportsOutput.Fields(iLoop).Name, mrstCustomReportsOutput.Fields(iLoop).Value, bSuppress, bBaseRecordChanged)
        End If
        
        '************************************************************************
        'Check if the current value is Grouped OR Grouped with the previous column
        'if the column value is Empty or Hidden then need to get the next column value
        If (mvarColDetails(24, iLoop + 1)) Or (mvarColDetails(24, iLoop)) Then
          If (intRowIndex_GW < 1) And ((IsNull(vDisplayData)) Or (vDisplayData = vbNullString)) Then
            blnSkipped = True
            intSkippedIndex = iLoop + 1
            'Get the formatted data of the next column to display in the grid
            vDisplayData = PopulateGrid_FormatData(mrstCustomReportsOutput.Fields(intSkippedIndex).Name, mrstCustomReportsOutput.Fields(intSkippedIndex).Value, bSuppress, bBaseRecordChanged)
          End If
          
        End If
        '************************************************************************
        
        If blnSkipped Then
          ' Store the ACTUAL data in the array (previous value dimension)
          If IsNull(mrstCustomReportsOutput.Fields(intSkippedIndex).Value) And _
            (Not mvarColDetails(3, intSkippedIndex + 1)) And _
            (Not mvarColDetails(17, intSkippedIndex + 1)) And _
            (Not mvarColDetails(18, intSkippedIndex + 1)) Then
            ' Field value is null but a character data type, so set it to be "".
            mvarColDetails(11, intSkippedIndex + 1) = ""
          Else
  
            'TM17052005 Fault 10086 - Need to store diffent values depending on the type.
            If mvarColDetails(17, intSkippedIndex + 1) Then 'Date
              mvarColDetails(11, intSkippedIndex + 1) = Format(mrstCustomReportsOutput.Fields(intSkippedIndex).Value, sDateFormat)
              
            ElseIf (mvarColDetails(3, intSkippedIndex + 1)) Then 'Numeric
              mvarColDetails(11, intSkippedIndex + 1) = IIf(IsNull(mrstCustomReportsOutput.Fields(intSkippedIndex).Value), "", mrstCustomReportsOutput.Fields(intSkippedIndex).Value)
              
            ElseIf (mvarColDetails(18, intSkippedIndex + 1)) Then 'Bit
              If (mrstCustomReportsOutput.Fields(intSkippedIndex).Value = "True") Or (mrstCustomReportsOutput.Fields(intSkippedIndex).Value = 1) Then mvarColDetails(11, intSkippedIndex + 1) = "Y"
              If (mrstCustomReportsOutput.Fields(intSkippedIndex).Value = "False") Or (mrstCustomReportsOutput.Fields(intSkippedIndex).Value = 0) Then mvarColDetails(11, intSkippedIndex + 1) = "N"
              If IsNull(mrstCustomReportsOutput.Fields(intSkippedIndex).Value) Then mvarColDetails(11, intSkippedIndex + 1) = ""
              
            Else 'Varchar
              mvarColDetails(11, intSkippedIndex + 1) = mrstCustomReportsOutput.Fields(intSkippedIndex).Value
            End If
          
          End If
        
'          If Not IsNull(mvarColDetails(11, intSkippedIndex + 1)) Then
'            'If len of data is greater than the previous length of data, store len in the array.
'            If (mfrmOutput.TextWidth(mvarColDetails(11, intSkippedIndex + 1)) + 150) > mlngColWidth(intColCounter) Then
'              mlngColWidth(intColCounter) = mfrmOutput.TextWidth(mvarColDetails(11, intSkippedIndex + 1)) + 150
'            End If
'          End If
          
        Else
          ' Store the ACTUAL data in the array (previous value dimension)
          If IsNull(mrstCustomReportsOutput.Fields(iLoop).Value) And _
            (Not mvarColDetails(3, iLoop + 1)) And _
            (Not mvarColDetails(17, iLoop + 1)) And _
            (Not mvarColDetails(18, iLoop + 1)) Then
            ' Field value is null but a character data type, so set it to be "".
            mvarColDetails(11, iLoop + 1) = ""
          Else
         
            'TM17052005 Fault 10086 - Need to store diffent values depending on the type.
            If mvarColDetails(17, iLoop + 1) Then 'Date
              mvarColDetails(11, iLoop + 1) = Format(mrstCustomReportsOutput.Fields(iLoop).Value, sDateFormat)
              
            ElseIf (mvarColDetails(3, iLoop + 1)) Then 'Numeric
              mvarColDetails(11, iLoop + 1) = IIf(IsNull(mrstCustomReportsOutput.Fields(iLoop).Value), "", mrstCustomReportsOutput.Fields(iLoop).Value)
              
            ElseIf (mvarColDetails(18, iLoop + 1)) Then 'Bit
              If (mrstCustomReportsOutput.Fields(iLoop).Value = "True") Or (mrstCustomReportsOutput.Fields(iLoop).Value = 1) Then mvarColDetails(11, iLoop + 1) = "Y"
              If (mrstCustomReportsOutput.Fields(iLoop).Value = "False") Or (mrstCustomReportsOutput.Fields(iLoop).Value = 0) Then mvarColDetails(11, iLoop + 1) = "N"
              If IsNull(mrstCustomReportsOutput.Fields(iLoop).Value) Then mvarColDetails(11, iLoop + 1) = ""
            Else 'Varchar
              mvarColDetails(11, iLoop + 1) = mrstCustomReportsOutput.Fields(iLoop).Value
            End If
          End If
        End If
        
        If Not IsNull(vDisplayData) Then
          If BigTextWidth(vDisplayData, 0) > mlngColWidth(intColCounter) Then
            mlngColWidth(intColCounter) = BigTextWidth(vDisplayData, 0)
          End If
        End If
        
        '************************************************************************
        
        'Add the displaydata to the main addstring.
        If IsNull(vDisplayData) Then
          sAddString = sAddString & vDisplayData
        Else
          sAddString = sAddString & IIf(vDisplayData = Space(Len(vDisplayData)), vbNullString, vDisplayData)
        End If
      
        'Add <tab> unless this is the last field
        If iLoop <> (mrstCustomReportsOutput.Fields.Count - 1) Then
          sAddString = sAddString & vbTab
        End If
        
        If (mvarColDetails(24, iLoop) And (Not mvarColDetails(19, iLoop))) Then
          
          If intColIndex_GW < 1 Then
            intColIndex_GW = intColCounter - 1
          End If

          strGroupedValue = IIf(IsNull(vDisplayData), "", vDisplayData)
        
          
          If (strGroupedValue <> vbNullString) And (Not blnSkipped) Then
            'add the grouped value to the string according to the row index
            PopulateGrid_AddToGroupWith strGroupedValue, intRowIndex_GW, intColIndex_GW
            blnHasGroupWithNext = True
            intRowIndex_GW = intRowIndex_GW + 1
          End If
          
          If blnSkipped Then
            blnSkipped = False
            intSkippedIndex = 0
          End If

        Else
          intRowIndex_GW = 0
          intColIndex_GW = 0
          
        End If
        
        '************************************************************************
          
      Next iLoop
      
      'Only Add the addstring to the grid if its not a summary report
      If mblnCustomReportsSummaryReport = False Then
        .grdOutput.AddItem sAddString
        mintPageBreakRowIndex = mintPageBreakRowIndex + 1
        
        If blnHasGroupWithNext Then
          'Add the Grouped Strings to to the grid
          For intGroupCount = 0 To UBound(mvarGroupWith, 2) Step 1
            .grdOutput.AddItem mvarGroupWith(0, intGroupCount)
            mintPageBreakRowIndex = mintPageBreakRowIndex + 1
          Next intGroupCount
        End If
      End If
        
      'Clear the Group Arrays/Variables
      ReDim mvarGroupWith(1, 0)
      intRowIndex_GW = 0
      intColIndex_GW = 0
      strGroupedValue = vbNullString
      intGroupCount = 0
      blnHasGroupWithNext = False
      
      intColCounter = 0
      
      ' Move to next row in the grid
      mrstCustomReportsOutput.MoveNext
    Loop
    
    mblnPageBreak = False
    
    ' Now do the final summary for the last bit (before the grand summary)
    ' Put the values from the previous record in the column array.
    For iLoop = 1 To UBound(mvarColDetails, 2)
      avColumns(3, iLoop) = mvarColDetails(11, iLoop)
    Next iLoop
    ' From last column in the order to first, check changes.
    For iLoop = UBound(mvarSortOrder, 2) To 1 Step -1
      ' Find the column in the details array.
      iColumnIndex = 0
      For iLoop2 = 1 To UBound(mvarColDetails, 2)
        If (mvarColDetails(12, iLoop2) = mvarSortOrder(1, iLoop)) And _
          (mvarColDetails(13, iLoop2) = "C") Then
              
          iColumnIndex = iLoop2
          Exit For
        End If
      Next iLoop2
      
      If iColumnIndex > 0 Then
        If (avColumns(1, iColumnIndex)) Then

          If mvarColDetails(8, iColumnIndex) Then
            mblnPageBreak = True
            sBreakValue = IIf(Len(mvarColDetails(11, iColumnIndex)) < 1, "<Empty>", mvarColDetails(11, iColumnIndex)) & IIf(Len(sBreakValue) > 0, " - ", "") & sBreakValue
          End If

          PopulateGrid_DoSummaryInfo avColumns, iColumnIndex, iLoop
        End If
      End If
    Next iLoop

    If mblnPageBreak Then
      mintPageBreakRowIndex = mintPageBreakRowIndex + 1
      AddPageBreakValue mintPageBreakRowIndex - 1, sBreakValue
    End If
    
    ' Now do the grand summary information
    If Not mbIsBradfordIndexReport Then
      PopulateGrid_DoGrandSummary

      If mblnPageBreak Then
        AddPageBreakValue mintPageBreakRowIndex - 1, sBreakValue
        mintPageBreakRowIndex = mintPageBreakRowIndex + 1
      End If

    End If
    
    ' Set the column widths to those stored in the array
    For iLoop = 0 To UBound(mlngColWidth) - 1
      .grdOutput.Columns(iLoop).Width = mlngColWidth(iLoop) + 200
    Next iLoop
  
  End With
  
  PopulateGrid_LoadRecords = True
  
  If Not gblnBatchMode Then
    gobjProgress.CloseProgress
  End If
  
  Exit Function

LoadRecords_ERROR:

  PopulateGrid_LoadRecords = False
  If Not gblnBatchMode Then gobjProgress.CloseProgress
  mstrErrorString = "Error in PopulateGrid_LoadRecords." & vbNewLine & Err.Description

End Function
Public Sub SetBradfordParameters( _
          dtStartDate As Date, _
          dtEndDate As Date, _
          lngPicklistFilterID As Long, _
          strPicklistFilterType As String, _
          sAbsenceTypes As String, _
          strOrderByColumn As String, _
          bOrderBy1Asc As Boolean, _
          strGroupByColumn As String, _
          bOrderBy2Asc As Boolean, _
          bBradfordSRV As Boolean, _
          bBradfordCount As Boolean, _
          bBradfordTotals As Boolean, _
          bBradfordWorkings As Boolean, _
          blnOmitBeforeStart As Boolean, _
          blnOmitAfterEnd As Boolean, _
          bMinBradford As Boolean, _
          lngMinBradfordAmount As Long, _
          bDisplayDetail As Boolean, _
          lngOrderByColumnID As Long, _
          lngGroupByColumnID As Long)

  'Dim blnCustomDates As Boolean

  'blnCustomDates = blnCustomDates
  'mlngStartDateExprID = lngStartDateExprID
  'mlngEndDateExprID = lngEndDateExprID
  mdtAbsenceFrom = dtStartDate
  mdtAbsenceTo = dtEndDate
  mlngPicklistFilterID = lngPicklistFilterID
  mstrPicklistFilterType = strPicklistFilterType
  msAbsenceTypes = sAbsenceTypes
  mbIsBradfordIndexReport = True
  mbBradfordSRV = bBradfordSRV
  mbBradfordTotals = bBradfordTotals
  mbBradfordCount = bBradfordCount
  mbBradfordWorkings = bBradfordWorkings
  mlngOrderByColumnID = lngOrderByColumnID
  mstrOrderByColumn = strOrderByColumn
  mlngGroupByColumnID = lngGroupByColumnID
  mstrGroupByColumn = strGroupByColumn
  mbOrderBy1Asc = bOrderBy1Asc
  mbOrderBy2Asc = bOrderBy2Asc
  mblnOmitBeforeStart = blnOmitBeforeStart
  mblnOmitAfterEnd = blnOmitAfterEnd
  mbMinBradford = bMinBradford
  mlngMinBradfordAmount = lngMinBradfordAmount
  mbDisplayBradfordDetail = bDisplayDetail

  'If blnCustomDates Then
  '  mdtAbsenceFrom = datGeneral.GetValueForRecordIndependantCalc(lngStartDateExprID)
  '  mdtAbsenceTo = datGeneral.GetValueForRecordIndependantCalc(lngEndDateExprID)
  'Else
  '  mdtAbsenceTo = DateAdd("d", Day(Date) * -1, Date)
  '  mdtAbsenceFrom = DateAdd("d", 1, DateAdd("yyyy", -1, mdtAbsenceTo))
  'End If

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
          strOutputFileName As String, _
          blnPreviewOnScreen As Boolean, _
          blnPrintFilterHeader As Boolean, _
          Optional strOutputTitlePage As String, _
          Optional strOutputReportPackTitle As String, _
          Optional strOutputOverrideFilter As String, _
          Optional blnOutputTOC As Boolean, _
          Optional blnOutputCoverSheet As Boolean, _
          Optional lngOverrideFilterID As Long, _
          Optional blnOutputRetainPivotOrChart As Boolean, _
          Optional lngOriginalOutputFormat As Long)
          
  mlngOriginalOutputFormat = lngOriginalOutputFormat
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
  mstrOutputFileName = strOutputFileName
  mblnCustomReportsPrintFilterHeader = blnPrintFilterHeader
  mblnPreviewOnScreen = (blnPreviewOnScreen Or (mlngOutputFormat = fmtDataOnly And mblnOutputScreen))
  mstrOutputTitlePage = IIf(IsMissing(strOutputTitlePage), giEXPRVALUE_CHARACTER, strOutputTitlePage)
  mstrOutputReportPackTitle = IIf(IsMissing(strOutputReportPackTitle), giEXPRVALUE_CHARACTER, strOutputReportPackTitle)
  mstrOutputOverrideFilter = IIf(IsMissing(strOutputOverrideFilter), giEXPRVALUE_CHARACTER, strOutputOverrideFilter)
  mblnOutputTOC = IIf(IsMissing(blnOutputTOC), giEXPRVALUE_CHARACTER, blnOutputTOC)
  mblnOutputCoverSheet = IIf(IsMissing(blnOutputCoverSheet), giEXPRVALUE_CHARACTER, blnOutputCoverSheet)
  mlngOverrideFilterID = IIf(IsMissing(lngOverrideFilterID), giEXPRVALUE_NUMERIC, lngOverrideFilterID)
  mblnOutputRetainPivotOrChart = IIf(IsMissing(blnOutputRetainPivotOrChart), giEXPRVALUE_CHARACTER, blnOutputRetainPivotOrChart)
  'mblnOutputRetainCharts = IIf(IsMissing(blnOutputRetainCharts), giEXPRVALUE_CHARACTER, blnOutputRetainCharts)
End Sub


Public Function GetReportConfig() As Boolean

  Const strReportType As String = "BradfordFactor"
  Dim lngStartDateExprID As Long
  Dim lngEndDateExprID As Long
  Dim lngOrderByColumnID As Long
  Dim lngGroupByColumnID As Long

  If GetSystemSetting(strReportType, "Custom Dates", "0") = "1" Then
    lngStartDateExprID = GetSystemSetting(strReportType, "Start Date", 0)
    mstrErrorString = IsCalcValid(lngStartDateExprID)
    If mstrErrorString <> vbNullString Then
      GetReportConfig = False
      Exit Function
    End If
    mdtAbsenceFrom = datGeneral.GetValueForRecordIndependantCalc(lngStartDateExprID)

    lngEndDateExprID = GetSystemSetting(strReportType, "End Date", 0)
    mstrErrorString = IsCalcValid(lngEndDateExprID)
    If mstrErrorString <> vbNullString Then
      GetReportConfig = False
      Exit Function
    End If
    mdtAbsenceTo = datGeneral.GetValueForRecordIndependantCalc(lngEndDateExprID)

    'MH20030911 Fault 5991
    If DateDiff("d", mdtAbsenceFrom, mdtAbsenceTo) < 0 Then
      mstrErrorString = "The report end date is before the report start date."
      GetReportConfig = False
      Exit Function
    End If

  Else
    mdtAbsenceTo = DateAdd("d", Day(Date) * -1, Date)
    mdtAbsenceFrom = DateAdd("d", 1, DateAdd("yyyy", -1, mdtAbsenceTo))
  End If
  
  
  
'  If GetSystemSetting(strReportType, "Order By", 0) = 0 Then
'    mstrErrorString = "The Bradford Factor Report has not been setup." & vbNewLine & _
'                      "Please save the report configuration."
'    GetReportConfig = False
'    Exit Function
'  End If


  mblnCustomReportsPrintFilterHeader = (GetSystemSetting(strReportType, "PrintFilterHeader", "0") = "1")
  mlngPicklistFilterID = GetSystemSetting(strReportType, "ID", "0")
  mstrPicklistFilterType = GetSystemSetting(strReportType, "Type", "A")

  ' Set Report Display Options
  mbBradfordSRV = (GetSystemSetting(strReportType, "SRV", "0") = "1")
  mbBradfordTotals = (GetSystemSetting(strReportType, "Show Totals", "1") = "1")
  mbBradfordCount = (GetSystemSetting(strReportType, "Show Count", "0") = "1")
  mbBradfordWorkings = (GetSystemSetting(strReportType, "Show Workings", "0") = "1")
  
  lngOrderByColumnID = GetSystemSetting(strReportType, "Order By", 0)
  If lngOrderByColumnID > 0 Then
    mstrOrderByColumn = datGeneral.GetColumnName(lngOrderByColumnID)
  End If
  
  lngGroupByColumnID = GetSystemSetting(strReportType, "Group By", 0)
  If lngGroupByColumnID > 0 Then
    mstrGroupByColumn = datGeneral.GetColumnName(lngGroupByColumnID)
  End If
  
  mbOrderBy1Asc = (GetSystemSetting(strReportType, "Order By Asc", "1") = "1")
  mbOrderBy2Asc = (GetSystemSetting(strReportType, "Group By Asc", "1") = "1")
  mblnOmitBeforeStart = (GetSystemSetting(strReportType, "Omit Before", "0") = "1")
  mblnOmitAfterEnd = (GetSystemSetting(strReportType, "Omit After", "0") = "1")
  
  If strReportType = "BradfordFactor" Then
    msAbsenceTypes = GetAbsenceTypesConfig(strReportType)
  End If
  
  mlngOutputFormat = GetSystemSetting(strReportType, "Format", 0)
  mblnOutputScreen = GetSystemSetting(strReportType, "Screen", 1)
  mblnOutputPrinter = GetSystemSetting(strReportType, "Printer", 0)
  mstrOutputPrinterName = GetSystemSetting(strReportType, "PrinterName", vbNullString)
  mblnOutputSave = GetSystemSetting(strReportType, "Save", 0)
  mlngOutputSaveExisting = GetSystemSetting(strReportType, "SaveExisting", -1)
  mblnOutputEmail = GetSystemSetting(strReportType, "Email", 0)
  mlngOutputEmailAddr = GetSystemSetting(strReportType, "EmailAddr", 0)
  mstrOutputEmailSubject = GetSystemSetting(strReportType, "EmailSubject", vbNullString)
  mstrOutputEmailAttachAs = GetSystemSetting(strReportType, "EmailAttachAs", vbNullString)
  mstrOutputFileName = GetSystemSetting(strReportType, "FileName", vbNullString)

  'mblnPreviewOnScreen = (GetSystemSetting(strReportType, "Preview", vbNullString) Or (mlngOutputFormat = fmtDataOnly And mblnOutputScreen))
  mblnPreviewOnScreen = (GetSystemSetting(strReportType, "Preview", True) Or (mlngOutputFormat = fmtDataOnly And mblnOutputScreen))
  
  ' Minimum Bradford Factor
  mbMinBradford = (GetSystemSetting(strReportType, "Minimum Bradford Factor", "0") = "1")
  mlngMinBradfordAmount = GetSystemSetting(strReportType, "Minimum Bradford Factor Amount", 0)
  mbDisplayBradfordDetail = GetSystemSetting(strReportType, "Display Absence Details", 0)

  GetReportConfig = True

End Function

Private Function GetAbsenceTypesConfig(pstrReportType As String) As String

  Dim rsType As Recordset
  Dim strSQL As String
  Dim strSettingKey As String
  Dim strOutput As String

  strSQL = "SELECT * " & _
           "FROM " & gsAbsenceTypeTableName & " " & _
           "ORDER BY " & gsAbsenceTypeTypeColumnName
  Set rsType = datGeneral.GetReadOnlyRecords(strSQL)

  strOutput = vbNullString
  Do Until rsType.EOF

    strSettingKey = "Absence Type " & rsType.Fields(gsAbsenceTypeTypeColumnName).Value
    If GetSystemSetting(pstrReportType, strSettingKey, vbNullString) = "1" Then
      strOutput = strOutput & _
        IIf(strOutput <> vbNullString, ",", vbNullString) & _
        "'" & rsType.Fields(gsAbsenceTypeTypeColumnName).Value & "'"
    End If

    rsType.MoveNext
  Loop

  GetAbsenceTypesConfig = strOutput

  Set rsType = Nothing

End Function


