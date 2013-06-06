VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalendarReportsRUN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
  
Private mlngCalendarReportID As Long
Private mstrErrorString As String
Private mstrTempTableName As String
Private mblnUDFsCreated As Boolean
Private mblnTempTableCreated As Boolean
Private mblnOrderByCreated As Boolean
Private mlngBaseDescriptionType As Long

'Variables to store definition (report level variables)
Private mstrCalendarReportsName As String
Private mlngCalendarReportsBaseTable As Long
Private mstrCalendarReportsBaseTableName As String
Private mlngCalendarReportsAllRecords As Long
Private mlngCalendarReportsPickListID As Long
Private mlngCalendarReportsFilterID As Long

Private mlngDescription1 As Long
Private mstrDescription1 As String
Private mblnDesc1IsDate As Boolean
Private mlngDescription2 As Long
Private mstrDescription2 As String
Private mblnDesc2IsDate As Boolean
Private mlngDescriptionExpr As Long
Private mstrDescriptionExpr As String
Private mblnDescExprIsDate As Boolean

Private mlngRegion As Long
Private mstrRegion As String
Private mstrRegionColumnRealSource As String
Private mblnGroupByDescription As Boolean
Private mlngDescription1ID, mlngDescription2ID As Long
Private mstrDescriptionSeparator As String

Private mlngStartDateExpr As Long
Private mstrStartDate As String
Private mdtStartDate As Date
Private mlngEndDateExpr As Long
Private mstrEndDate As String
Private mdtEndDate As String

Private mblnShowBankHolidays As Boolean
Private mblnShowCaptions As Boolean
Private mblnShowWeekends As Boolean
Private mblnIncludeWorkingDaysOnly As Boolean
Private mblnIncludeBankHolidays As Boolean
Private mbStartOnCurrentMonth As Boolean
Private mblnCustomReportsPrintFilterHeader As Boolean
Private mstrFilteredIDs As String


  'New Default Output Variables
  Private mlngOutputFormat As Long
  Private mblnOutputScreen As Boolean
  Private mblnOutputPrinter As Boolean
  Private mstrOutputPrinterName As String
  Private mblnOutputSave As Boolean
  Private mlngOutputSaveExisting As Long
  Private mlngOutputSaveFormat As Long
  Private mblnOutputEmail As Boolean
  Private mlngOutputEmailAddr As Long
  Private mstrOutputEmailSubject As String
  Private mstrOutputEmailAttachAs As String
  Private mlngOutputEmailFileFormat As Long
  Private mstrOutputFilename As String
  Private mblnPreviewOnScreen As Boolean
 
'Recordset to store the final data from the temp table
Private mrsCalendarReportsOutput As ADODB.Recordset
Private mrsCalendarBaseInfo As ADODB.Recordset

'Array to store data for each session label object,
'array also holds information on working days and bank holidays etc.
Private mvarDateLabelInfo() As Variant

'Strings to hold the SQL statement
Private mstrSQLEvent As String
Private mstrSQLSelect As String
Private mstrSQLFrom As String
Private mstrSQLJoin As String
Private mstrSQLWhere As String
Private mstrSQLOrderBy As String
Private mstrSQL As String
Private mstrSQLBaseData As String
Private mstrSQLBaseDateClause As String
Private mstrSQLOrderList As String
Private mstrSQLIDs As String
Private mstrSQLDynamicLegendWhere As String
Private mintDynamicEventCount As Integer
Private mstrSQLCreateTable As HRProDataMgr.clsStringBuilder

Private mblnHasEventFilterIDs As Boolean
Private mstrEventFilterIDs As String

Private mblnDescCalcCode As Boolean
Private mstrDescCalcCode As String

'used to temporarily store the Base table Start & End date table.columnname for the
'current event. Then used when creating the mstrSQLBaseDateClause.
Private mstrSQLBaseStartDateColumn As String
Private mstrSQLBaseStartSessionColumn As String
Private mstrSQLBaseEndDateColumn As String
Private mstrSQLBaseEndSessionColumn As String
Private mstrSQLBaseDurationColumn As String

'Data access class
Private mclsData As clsDataAccess
Private mclsGeneral As clsGeneral

'Array holding the columns to sort the report by
Private mvarSortOrder() As Variant

Private mcolEvents As clsCalendarEvents

'Instance of the previewform
Private mfrmOutput As frmCalendarReportPreview

'TableViews
Private mstrRealSource As String
Private mstrBaseTableRealSource As String
Private mlngTableViews() As Long
Private mstrViews() As String
Private mobjTableView As CTablePrivilege
Private mobjColumnPrivileges As CColumnPrivileges

Private mvarEventColumnViews() As Variant
Private mlngEventViewColumn As Long

'Batch Job Mode ?
'Private gblnBatchMode As Boolean

'Runnning report for single record only!
Private mlngSingleRecordID As Long

'Has the user cancelled the report ?
Private mblnUserCancelled As Boolean

'Does the report generate no records ?
Private mblnNoRecords As Boolean

'Is the current user the definition owner ?
Private mblnDefinitionOwner As Boolean

' Array holding the User Defined functions that are needed for this report
Private mastrUDFsRequired() As String

Private mblnHasMultipleEvents As Boolean

'TM01042004 Fault 8428
Private mblnCheckingRegionColumn As Boolean
Private mblnCheckingDescColumn As Boolean

Private Function CheckColumnPermissions(plngTableID As Long, pstrTableName As String, _
                                        pstrColumnName As String, strSQLRef As String) As Boolean

  'This function checks if the current user has read(select) permissions
  'on this column. If the user only has access through views then the
  'relevent views are added to the mlngTableViews() array which in turn
  'are used to create the join part of the query.

  Dim lngTempTableID As Long
  Dim strTempTableName As String
  Dim strTempColumnName As String
  Dim blnColumnOK As Boolean
  Dim blnFound As Boolean
  Dim blnNoSelect As Boolean
  Dim iLoop1 As Integer
  Dim intLoop As Integer
  Dim strColumnCode As String
  Dim strSource As String
  Dim intNextIndex As Integer
  Dim blnOK As Boolean
  Dim strTable As String
  Dim strColumn As String
  
  Dim pintNextIndex  As Integer
  
  Dim bDateColumn As Boolean
  
  ' Set flags with their starting values
  blnOK = True
  blnNoSelect = False
  bDateColumn = False
  
  strTable = vbNullString
  strColumn = vbNullString
 
  ' Load the temp variables
  lngTempTableID = plngTableID
  strTempTableName = pstrTableName
  strTempColumnName = pstrColumnName

  ' Check permission on that column
  Set mobjColumnPrivileges = GetColumnPrivileges(strTempTableName)
  mstrRealSource = gcoTablePrivileges.Item(strTempTableName).RealSource

  blnColumnOK = mobjColumnPrivileges.IsValid(strTempColumnName)

  If blnColumnOK Then
    blnColumnOK = mobjColumnPrivileges.Item(strTempColumnName).AllowSelect
  End If
  
  If mobjColumnPrivileges.Item(strTempColumnName).DataType = SQLDataType.sqlDate Then
    bDateColumn = True
  End If
  
  If blnColumnOK Then
    ' this column can be read direct from the tbl/view or from a parent table
    strTable = mstrRealSource
    strColumn = strTempColumnName
    
    ' If the table isnt the base table (or its realsource) then
    ' Check if it has already been added to the array. If not, add it.
    If lngTempTableID <> mlngCalendarReportsBaseTable Then
      blnFound = False
      For intNextIndex = 1 To UBound(mlngTableViews, 2)
        If mlngTableViews(1, intNextIndex) = 0 And _
        mlngTableViews(2, intNextIndex) = lngTempTableID Then
        blnFound = True
          Exit For
        End If
      Next intNextIndex

      If Not blnFound Then
        intNextIndex = UBound(mlngTableViews, 2) + 1
        ReDim Preserve mlngTableViews(2, intNextIndex)
        mlngTableViews(1, intNextIndex) = 0
        mlngTableViews(2, intNextIndex) = lngTempTableID
      End If
    End If
  
    If bDateColumn And mblnCheckingDescColumn Then
      strSQLRef = SQLDateConvertToLocale(strTable & "." & strColumn)
    Else
      strSQLRef = strTable & "." & strColumn
    End If
  Else

    ' this column cannot be read direct. If its from a parent, try parent views
    ' Loop thru the views on the table, seeing if any have read permis for the column

    ReDim mstrViews(0)
    For Each mobjTableView In gcoTablePrivileges.Collection
      If (Not mobjTableView.IsTable) And _
          (mobjTableView.TableID = lngTempTableID) And _
          (mobjTableView.AllowSelect) Then

        strSource = mobjTableView.ViewName
        mstrRealSource = gcoTablePrivileges.Item(strSource).RealSource

        ' Get the column permission for the view
        Set mobjColumnPrivileges = GetColumnPrivileges(strSource)

        ' If we can see the column from this view
        If mobjColumnPrivileges.IsValid(strTempColumnName) Then
          If mobjColumnPrivileges.Item(strTempColumnName).AllowSelect Then
            
            ReDim Preserve mstrViews(UBound(mstrViews) + 1)
            mstrViews(UBound(mstrViews)) = mobjTableView.ViewName
            
            If mlngEventViewColumn > 0 Then
              ReDim Preserve mvarEventColumnViews(1, UBound(mvarEventColumnViews, 2) + 1)
              mvarEventColumnViews(0, UBound(mvarEventColumnViews, 2)) = mobjTableView.ViewID
              mvarEventColumnViews(1, UBound(mvarEventColumnViews, 2)) = mlngEventViewColumn
            End If
            
            ' Check if view has already been added to the array
            blnFound = False
            For intNextIndex = 0 To UBound(mlngTableViews, 2)
              If mlngTableViews(1, intNextIndex) = 1 And _
              mlngTableViews(2, intNextIndex) = mobjTableView.ViewID Then
                blnFound = True
                Exit For
              End If
            Next intNextIndex

            If Not blnFound Then
              ' View hasnt yet been added, so add it !
              intNextIndex = UBound(mlngTableViews, 2) + 1
              ReDim Preserve mlngTableViews(2, intNextIndex)
              mlngTableViews(1, intNextIndex) = 1
              mlngTableViews(2, intNextIndex) = mobjTableView.ViewID
'              Exit For
            Else
'              Exit For
            End If
            
          End If
        End If
      End If

    Next mobjTableView
    Set mobjTableView = Nothing

    ' Does the user have select permission thru ANY views ?
    If UBound(mstrViews()) = 0 Then
      blnNoSelect = True
    Else
      strSQLRef = ""
      For pintNextIndex = 1 To UBound(mstrViews)
        If pintNextIndex = 1 Then
          strSQLRef = "CASE"
        End If
        
        strSQLRef = strSQLRef & _
        " WHEN NOT " & mstrViews(pintNextIndex) & "." & strTempColumnName & " IS NULL THEN "
        
        If bDateColumn And mblnCheckingDescColumn Then
          strSQLRef = strSQLRef & SQLDateConvertToLocale(mstrViews(pintNextIndex) & "." & strTempColumnName)
        Else
          strSQLRef = strSQLRef & mstrViews(pintNextIndex) & "." & strTempColumnName
        End If
        
      Next pintNextIndex
      
      If Len(strSQLRef) > 0 Then
        strSQLRef = strSQLRef & _
        " ELSE NULL" & _
        " END "
      End If

'      strTable = mstrRealSource
'      strColumn = strTempColumnName
    End If

    ' If we cant see a column, then get outta here
    If blnNoSelect Then
      strSQLRef = vbNullString
      CheckColumnPermissions = False
      mstrErrorString = "You do not have permission to see the column '" & pstrColumnName & "'" & vbNewLine & "either directly or through any views."
      Exit Function
    End If

    If Not blnOK Then
      strSQLRef = vbNullString
      CheckColumnPermissions = False
      Exit Function
    End If

  End If

  'TM01042004 Fault 8428
  If mblnCheckingRegionColumn = True Then
    mstrRegionColumnRealSource = mstrRealSource
  End If

  mlngEventViewColumn = 0

'  strSQLRef = strTable & "." & strColumn
  CheckColumnPermissions = True
  
End Function

Private Function CheckCalculationPermissions(plngBaseTableID As Long, plngExprID As Long, strSQLRef As String) As Boolean

  'This function checks if the current user has read(select) permissions
  'on this calculation. If the user only has access through views then the
  'relevent views are added to the mlngTableViews() array which in turn
  'are used to create the join part of the query.

  Dim lngTempTableID As Long
  Dim strTempTableName As String
  Dim strTempColumnName As String
  Dim blnColumnOK As Boolean
  Dim blnFound As Boolean
  Dim blnNoSelect As Boolean
  Dim iLoop1 As Integer
  Dim intLoop As Integer
  Dim strColumnCode As String
  Dim strSource As String
  Dim intNextIndex As Integer
  Dim blnOK As Boolean
  Dim strTable As String
  Dim strColumn As String
  Dim sCalcCode As String
  Dim alngSourceTables() As Long
  Dim objCalcExpr As clsExprExpression

  ' Set flags with their starting values
  blnOK = True
 
  ' Get the calculation SQL, and the array of tables/views that are used to create it.
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
  ' Column 2 = table/view ID.
  ReDim alngSourceTables(2, 0)
  Set objCalcExpr = New clsExprExpression
  blnOK = objCalcExpr.Initialise(plngBaseTableID, plngExprID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_UNDEFINED)
  If blnOK Then
    blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sCalcCode, True)
    
    If blnOK And gbEnableUDFFunctions Then
      blnOK = objCalcExpr.UDFCalculationCode(alngSourceTables, mastrUDFsRequired(), True)
    End If
  End If
  
  'The "SELECT ... INTO..." statement errors when it trys to create a column for
  'and empty string. Therefore wrap this empty sting in a CONVERT(varchar... clause if an sql empty string
  'is returned.
  'TM20030521 Fault 5702 - Compare the empty string with the calc code value converted to varchar
  sCalcCode = "CASE WHEN CONVERT(varchar," & sCalcCode & ") = '' " & _
              "THEN CONVERT(varchar," & sCalcCode & ") " & _
              "ELSE " & sCalcCode & " END"
              
  If blnOK Then
    strSQLRef = sCalcCode

    ' Add the required views to the JOIN code.
     For iLoop1 = 1 To UBound(alngSourceTables, 2)
       If alngSourceTables(1, iLoop1) = 1 Then
         ' Check if view has already been added to the array
         blnFound = False
         For intNextIndex = 1 To UBound(mlngTableViews, 2)
           If mlngTableViews(1, intNextIndex) = 1 And _
             mlngTableViews(2, intNextIndex) = alngSourceTables(2, iLoop1) Then
             blnFound = True
             Exit For
           End If
         Next intNextIndex
                 
         If Not blnFound Then
                 
           ' View hasnt yet been added, so add it !
           intNextIndex = UBound(mlngTableViews, 2) + 1
           ReDim Preserve mlngTableViews(2, intNextIndex)
           mlngTableViews(1, intNextIndex) = 1
           mlngTableViews(2, intNextIndex) = alngSourceTables(2, iLoop1)
                   
         End If
'********************************************************************************
      ElseIf alngSourceTables(1, iLoop1) = 0 Then
         ' Check if table has already been added to the array
         blnFound = False
         For intNextIndex = 1 To UBound(mlngTableViews, 2)
           If mlngTableViews(1, intNextIndex) = 0 And _
             mlngTableViews(2, intNextIndex) = alngSourceTables(2, iLoop1) Then
             blnFound = True
             Exit For
           End If
         Next intNextIndex
        
        'Only want to check if the source table is the base table
        ' if we have NOT just found the source table in the array of joined tables.
        If Not blnFound Then
          blnFound = (alngSourceTables(2, iLoop1) = mlngCalendarReportsBaseTable)
        End If
        
        If Not blnFound Then
          ' table hasnt yet been added, so add it !
          intNextIndex = UBound(mlngTableViews, 2) + 1
          ReDim Preserve mlngTableViews(2, intNextIndex)
          mlngTableViews(1, intNextIndex) = 0
          mlngTableViews(2, intNextIndex) = alngSourceTables(2, iLoop1)
        End If
'********************************************************************************
      End If
    Next iLoop1
  Else
    ' Permission denied on something in the calculation.
    mstrErrorString = "You do not have permission to use the '" & objCalcExpr.Name & "' calculation."
    CheckCalculationPermissions = False
    Exit Function
  End If

  Set objCalcExpr = Nothing
  
  CheckCalculationPermissions = True
  
End Function

Private Function GenerateSQLEvent(pstrEventKey As String, _
                                    pstrDynamicKey As String, _
                                    pstrDynamicName As String) As Boolean

  Dim fOK As Boolean
  
  fOK = True
  
  If fOK Then fOK = GenerateSQLSelect(pstrEventKey, pstrDynamicKey, pstrDynamicName)
  If fOK Then fOK = GenerateSQLFrom()
  If fOK Then fOK = GenerateSQLJoin(pstrEventKey, pstrDynamicKey)
  
  If fOK Then fOK = GenerateSQLWhere(pstrEventKey, pstrDynamicKey, pstrDynamicName)
  
  If fOK Then
    mstrSQLEvent = mstrSQLSelect & vbNewLine & _
                   mstrSQLFrom & vbNewLine & _
                   mstrSQLJoin & vbNewLine & _
                   mstrSQLWhere & vbNewLine
  End If
  
  ' reset strings to hold the SQL statement
  mstrSQLSelect = vbNullString
  mstrSQLFrom = vbNullString
  mstrSQLJoin = vbNullString
  mstrSQLWhere = vbNullString

  GenerateSQLEvent = fOK
  
End Function


Private Function GetOrderArray() As Boolean
  
  On Error GoTo Error_Trap
  
  Dim rsTemp As ADODB.Recordset
  
  Dim sSQL As String
  
  Dim intTemp As Integer
  
  ' Get columns defined as a SortOrder and load into array
  sSQL = "SELECT * FROM ASRSysCalendarReportOrder WHERE " & _
               "CalendarReportID = " & mlngCalendarReportID & " " & _
               "ORDER BY [OrderSequence]"
               
  Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
  
  With rsTemp
    If .BOF And .EOF Then
      GetOrderArray = False
      mstrErrorString = "No columns have been defined as a sort order for the specified Calendar Report definition." & vbNewLine & "Please remove this definition and create a new one."
      Exit Function
    End If
    Do Until .EOF
      intTemp = UBound(mvarSortOrder, 2) + 1
      ReDim Preserve mvarSortOrder(2, intTemp)
      
      mvarSortOrder(0, intTemp) = !ColumnID
      mvarSortOrder(1, intTemp) = datGeneral.GetColumnName(!ColumnID)
      mvarSortOrder(2, intTemp) = !OrderType
      
      .MoveNext
    Loop
  End With

  GetOrderArray = True

TidyUpAndExit:
  Set rsTemp = Nothing
  Exit Function

Error_Trap:
  GetOrderArray = False
  mstrErrorString = "Error whilst retrieving the event details recordsets'." & vbNewLine & Err.Description

End Function
Private Function InsertIntoTempTable(pstrSelectString As String) As Boolean

  Dim strSQL As String
  Dim fOK As Boolean
  
  fOK = True
  
  If Not mblnTempTableCreated Then
    If Not CreateTempTable Then
      InsertIntoTempTable = False
      mstrErrorString = "Error creating the temporary table"
      GoTo TidyUpAndExit
    End If
  End If

  If (Not mblnUDFsCreated) Then
    fOK = UDFFunctions(mastrUDFsRequired, True)
    mblnUDFsCreated = fOK
    If Not fOK Then
      InsertIntoTempTable = False
      mstrErrorString = "Error creating SQL User Defined Functions"
      GoTo TidyUpAndExit
    End If
  End If
  
  strSQL = "INSERT INTO [" & gsUserName & "].[" & mstrTempTableName & "] " _
    & pstrSelectString
 
  mclsData.ExecuteSql strSQL
 
  InsertIntoTempTable = True
  
TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  InsertIntoTempTable = False
  mstrErrorString = "Error inserting into the temporary table"
  GoTo TidyUpAndExit
  
End Function

Private Function SQLDateConvertToLocale(pstrTableColumn As String) As String

  'Takes the Column value and Returns a string with the SQL Code to format the
  'SQL date value into the known locale.
  
  Dim strSQL As String
  Dim strDateFormat As String
  
  Dim blnDateComplete As Boolean
  Dim blnMonthDone As Boolean
  Dim blnDayDone As Boolean
  Dim blnYearDone As Boolean
  
  Dim strShortDate As String
  
  Dim strDateSeparator As String
  
  Dim i As Integer
  
  ' eg. DateFormat = "mm/dd/yyyy"
  '     Calendar   = "dd/mm/yyyy"
  '     DateString = "06/02/2000"
  '     Compare to = 02/06/2000
  
  strDateFormat = DateFormat

  strDateSeparator = UI.GetSystemDateSeparator
  
  blnDateComplete = False
  blnMonthDone = False
  blnDayDone = False
  blnYearDone = False

  strShortDate = vbNullString
  
  For i = 1 To Len(strDateFormat) Step 1
    
    If (LCase(Mid(strDateFormat, i, 1)) = "d") And (Not blnDayDone) Then
      strShortDate = strShortDate + LCase(Mid(strDateFormat, i, 1))
      blnDayDone = True
    End If
    
    If (LCase(Mid(strDateFormat, i, 1)) = "m") And (Not blnMonthDone) Then
      strShortDate = strShortDate + LCase(Mid(strDateFormat, i, 1))
      blnMonthDone = True
    End If
    
    If (LCase(Mid(strDateFormat, i, 1)) = "y") And (Not blnYearDone) Then
      strShortDate = strShortDate + LCase(Mid(strDateFormat, i, 1))
      blnYearDone = True
    End If

    If blnDayDone And blnMonthDone And blnYearDone Then
      blnDateComplete = True
      Exit For
    End If
  
  Next i
  
  Select Case strShortDate
    Case "dmy"
      SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & _
                                "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & _
                                "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) "
    Case "mdy"
      SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & _
                                "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & _
                                "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) "
    Case "ydm"
      SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & _
                                "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & _
                                "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) "
    Case "myd"
      SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & _
                                "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & _
                                "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) "
    Case "ymd"
      SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & _
                                "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & _
                                "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) "
  End Select
  

End Function

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property
Public Property Get CalendarReportID() As Long
  CalendarReportID = mlngCalendarReportID
End Property
Public Property Let CalendarReportID(ByVal lNewID As Long)
  mlngCalendarReportID = lNewID
End Property
Public Property Get ErrorString() As String
  ErrorString = mstrErrorString
End Property
Public Property Let ErrorString(ByVal sErrorString As String)
  mstrErrorString = sErrorString
End Property
Private Sub Class_Initialize()
  
  ' Purpose : Sets references to other classes and redimensions arrays
  '           used for table usage information
  
  Set mclsData = New HRProDataMgr.clsDataAccess
  Set mclsGeneral = New HRProDataMgr.clsGeneral
  ReDim mvarSortOrder(2, 0)
  
  ' Dimension an array of tables/views joined to the base table/view
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view
  ' Column 2 = table/view ID
  ' (should contain everything which needs to be joined to the base tbl/view)
  ReDim mlngTableViews(2, 0)
  
  ReDim mstrViews(0)
  ReDim mastrUDFsRequired(0)
  
  ReDim mvarEventColumnViews(1, 0)
  
  Set mstrSQLCreateTable = New HRProDataMgr.clsStringBuilder
  mstrSQLCreateTable.TheString = vbNullString
  
  
End Sub

Private Sub Class_Terminate()
  
  ' Purpose : Clears references to other classes.
  Set mclsData = Nothing
  Set mclsGeneral = Nothing
  Set mfrmOutput = Nothing
  Set mcolEvents = Nothing
  Set mobjTableView = Nothing
  Set mobjColumnPrivileges = Nothing
  
End Sub

Public Function RunCalendarReport(Optional plngRecordID As Long) As Boolean

  ' Purpose : This function is called from frmDefsel and Batch Jobs.
  '           It is the main function which runs the report.

  On Error GoTo Error_Trap
  
  Dim fOK As Boolean
  
  fOK = True
  
  Screen.MousePointer = vbHourglass
  
  mlngSingleRecordID = plngRecordID
  
  If fOK Then fOK = GetCalendarReportDefinition
  
  gobjEventLog.AddHeader eltCalandarReport, mstrCalendarReportsName
  
  If fOK Then
    With gobjProgress
      '.AviFile = App.Path & "\videos\report.avi"
      .AVI = dbText
      .MainCaption = "Calendar Report"
      If Not gblnBatchMode Then
        .Bar1Value = 0    'MH20040106 Fault 7887
        .NumberOfBars = 1
        .Caption = "Calendar Reports"
        .Time = False
        .Cancel = True
        .Bar1Caption = "Calendar Report : " & mstrCalendarReportsName
        .OpenProgress
      Else
        .Bar2Value = 0    'MH20040106 Fault 7887
        .Bar2Caption = "Calendar Report : " & mstrCalendarReportsName
      End If
    End With
  End If
  
  If fOK Then fOK = GetEventsCollection
  If fOK Then fOK = GetOrderArray
  If fOK Then fOK = GenerateSQL
  If fOK Then fOK = ExecuteSql
  If fOK Then fOK = UDFFunctions(mastrUDFsRequired, False)
  If fOK Then fOK = InitialisePreview
  
  If fOK Then
    fOK = mfrmOutput.Initialise
    If gblnBatchMode Then
      mblnUserCancelled = gobjProgress.Cancelled
    Else
      mblnUserCancelled = mfrmOutput.UserCancelled
    End If
    If Not fOK Then mstrErrorString = mfrmOutput.ErrorMessage
  End If
  
  If fOK And Not mblnUserCancelled Then
    If gblnBatchMode Then
      fOK = mfrmOutput.OutputReport(False)
      If Not fOK Then mstrErrorString = mfrmOutput.ErrorMessage
      
    Else
      If Not mblnPreviewOnScreen Then
        fOK = mfrmOutput.OutputReport(False)
      Else
        fOK = mfrmOutput.ShowCalendar
      End If
      mblnUserCancelled = mfrmOutput.UserCancelled
      If Not fOK Then mstrErrorString = mfrmOutput.ErrorMessage
      
      If fOK And Not mblnUserCancelled Then
        gobjProgress.CloseProgress
        Screen.MousePointer = vbNormal
        If fOK And mblnPreviewOnScreen Then
          mfrmOutput.Show vbModal
        End If
      End If
      
    End If
  End If
  
  Set mfrmOutput = Nothing
      
  Call UtilUpdateLastRun(utlCalendarReport, mlngCalendarReportID)
  
  If mblnNoRecords Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful
    gobjEventLog.AddDetailEntry mstrErrorString
    mstrErrorString = "Completed successfully." & vbNewLine & mstrErrorString
    fOK = True
  ElseIf fOK And Not mblnUserCancelled Then
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

  mstrErrorString = "Calendar Report : '" & mstrCalendarReportsName & "' " & mstrErrorString

  If Not gblnBatchMode Then
    If gobjProgress.Visible Then gobjProgress.CloseProgress
    If (fOK = False) Or (mblnNoRecords) Or (mblnUserCancelled) Or Not mblnPreviewOnScreen Then
      MsgBox mstrErrorString, IIf(fOK, vbInformation, vbExclamation) + vbOKOnly, "Calendar Reports"
    End If
  End If
  
  Screen.MousePointer = vbNormal
  
  RunCalendarReport = fOK
  
  If fOK = True Then
    fOK = ClearUp
  Else
    ClearUp
  End If

  Exit Function
  
Error_Trap:
  fOK = False
  RunCalendarReport = False
  mstrErrorString = "Error Whilst Running Calendar Report." & vbNewLine & Err.Description
  Resume Next
  
End Function

Private Function ExecuteSql() As Boolean

  ' Purpose : This function executes the SQL string 'into' a recordset.

  On Error GoTo ExecuteSQL_ERROR
  
'  'get all the base & event data into a recordset
  mstrSQL = vbNullString
  'MH20060301 Fault 10846
  'mstrSQL = mstrSQL & "SELECT * FROM [" & datGeneral.UserNameForSQL & "].[" & mstrTempTableName & "] "
  mstrSQL = mstrSQL & "SELECT * FROM [" & gsUserName & "].[" & mstrTempTableName & "] "
 
  'get the ORDER BY statement which applies to the entire UNIONed query.
  GenerateSQLOrderBy
  mstrSQL = mstrSQL & mstrSQLOrderBy
  
  Set mrsCalendarReportsOutput = mclsData.OpenRecordset(mstrSQL, adOpenForwardOnly, adLockReadOnly)
  
  If mrsCalendarReportsOutput.BOF And mrsCalendarReportsOutput.EOF Then
    ExecuteSql = False
    mstrErrorString = "No records meet selection criteria."
    mblnNoRecords = True
    Exit Function
  End If
  
  MultipleCheck
  
  'get only the base table info into a recordset
  Set mrsCalendarBaseInfo = mclsData.OpenRecordset(mstrSQLBaseData, adOpenForwardOnly, adLockReadOnly)

  If mrsCalendarBaseInfo.BOF And mrsCalendarBaseInfo.EOF Then
    ExecuteSql = False
    mstrErrorString = "No records meet selection criteria."
    mblnNoRecords = True
    Exit Function
  End If

  ExecuteSql = True
  Exit Function
  
ExecuteSQL_ERROR:
  
  mstrErrorString = "Error whilst executing SQL statement." & vbNewLine & Err.Description
  ExecuteSql = False
  
End Function

Private Function IsColumnInView(plngViewID As Long, plngColumnID As Long) As Boolean

  Dim lngCount As Long
  
  For lngCount = 1 To UBound(mvarEventColumnViews, 2) Step 1
    If (mvarEventColumnViews(0, lngCount) = plngViewID) And (mvarEventColumnViews(1, lngCount) = plngColumnID) Then
      IsColumnInView = True
      Exit Function
    End If
  Next lngCount
  
End Function

Private Function MultipleCheck() As Boolean

  Dim rsMultiple As ADODB.Recordset
  Dim sSQL As String
  Dim dtSD As Date
  Dim dtED As Date
  Dim strStartSession As String
  Dim strEndSession As String
  Dim lngBaseID As Long
  Dim strDescription1 As String
  Dim strDescription2 As String
  Dim strDescriptionExpr As String
  Dim lngCurrentBaseID As Long
  Dim avDateRanges() As Variant
  Dim i As Integer
  Dim blnHasOverlap As Boolean
  Dim intNewIndex As Integer
  Dim strFullDesc As String
  Dim strCurrentDesc As String
  Dim blnFirstCalendarRecord As Boolean
  
  blnFirstCalendarRecord = True
  
  ReDim avDateRanges(6, 0)
  
  sSQL = vbNullString
  sSQL = sSQL & "SELECT [BaseID], [Description1], [Description2], [DescriptionExpr], [StartDate], [StartSession], [EndDate], [EndSession] " & vbNewLine
  'MH20060301 Fault 10846
  'sSQL = sSQL & "FROM [" & datGeneral.UserNameForSQL & "].[" & mstrTempTableName & "]" & vbNewLine
  sSQL = sSQL & "FROM [" & gsUserName & "].[" & mstrTempTableName & "]" & vbNewLine
  sSQL = sSQL & mstrSQLOrderBy
  
  Set rsMultiple = datGeneral.GetReadOnlyRecords(sSQL)
  
  With rsMultiple
    If Not (.BOF And .EOF) Then
      Do Until .EOF
        dtSD = .Fields("StartDate")
        dtED = IIf(IsNull(.Fields("EndDate")), .Fields("StartDate"), .Fields("EndDate"))
        strStartSession = IIf(IsNull(.Fields("StartSession")), "AM", .Fields("StartSession"))
        If ((IsNull(.Fields("EndDate"))) And (IsNull(.Fields("EndSession"))) And (IsNull(.Fields("StartSession")))) Then
          strEndSession = "PM"
        ElseIf ((IsNull(.Fields("EndDate"))) And (IsNull(.Fields("EndSession")))) Then
          strEndSession = strStartSession
        ElseIf IsNull(.Fields("EndSession")) Then
          strEndSession = "PM"
        Else
          strEndSession = .Fields("EndSession")
        End If
        
        If mblnGroupByDescription Then
          strDescription1 = IIf(IsNull(.Fields("Description1")), "", .Fields("Description1"))
          strDescription2 = IIf(IsNull(.Fields("Description2")), "", .Fields("Description2"))
          strDescriptionExpr = IIf(IsNull(.Fields("DescriptionExpr")), "", .Fields("DescriptionExpr"))
          strFullDesc = strDescription1 & mstrDescriptionSeparator & strDescription2 & mstrDescriptionSeparator & strDescriptionExpr
        
          If (strFullDesc <> strCurrentDesc) Or blnFirstCalendarRecord Then
            strCurrentDesc = strFullDesc
            blnFirstCalendarRecord = False
            
            ReDim avDateRanges(6, 0)
            avDateRanges(0, 0) = strFullDesc
            avDateRanges(1, 0) = dtSD
            avDateRanges(2, 0) = dtED
            avDateRanges(3, 0) = strStartSession
            avDateRanges(4, 0) = strEndSession
            
          Else
            'Loop through the array for the current calendar row, checking if any dates overlap.
            For i = 0 To UBound(avDateRanges, 2) Step 1
              
              'if the start or end dates 'equal' any other start orend dates then check if the sessions are also equal.
              If ((dtSD = avDateRanges(1, i)) And (strStartSession = avDateRanges(3, i))) _
                Or ((dtSD = avDateRanges(2, i)) And (strStartSession = avDateRanges(4, i))) _
                Or ((dtED = avDateRanges(1, i)) And (strEndSession = avDateRanges(3, i))) _
                Or ((dtED = avDateRanges(2, i)) And (strEndSession = avDateRanges(4, i))) Then
                mblnHasMultipleEvents = True
                MultipleCheck = True
                GoTo TidyUpAndExit
              End If
              
              If ((dtSD > avDateRanges(1, i)) And (dtED < avDateRanges(2, i))) _
                Or ((dtSD > avDateRanges(1, i)) And (dtSD < avDateRanges(2, i)) And (dtED > avDateRanges(2, i))) _
                Or ((dtED > avDateRanges(1, i)) And (dtED < avDateRanges(2, i)) And (dtSD < avDateRanges(1, i))) _
                Or ((dtSD < avDateRanges(1, i)) And (dtED > avDateRanges(2, i))) Then
                mblnHasMultipleEvents = True
                MultipleCheck = True
                GoTo TidyUpAndExit
              End If
            Next i
            
            intNewIndex = UBound(avDateRanges, 2) + 1
            ReDim Preserve avDateRanges(6, intNewIndex)
            avDateRanges(0, intNewIndex) = lngBaseID
            avDateRanges(1, intNewIndex) = dtSD
            avDateRanges(2, intNewIndex) = dtED
            avDateRanges(3, intNewIndex) = strStartSession
            avDateRanges(4, intNewIndex) = strEndSession
            
          End If
        Else
          lngBaseID = .Fields("BaseID")
          
          If (lngBaseID <> lngCurrentBaseID) Or blnFirstCalendarRecord Then
            lngCurrentBaseID = lngBaseID
            blnFirstCalendarRecord = False
            
            ReDim avDateRanges(6, 0)
            avDateRanges(0, 0) = lngBaseID
            avDateRanges(1, 0) = dtSD
            avDateRanges(2, 0) = dtED
            avDateRanges(3, 0) = strStartSession
            avDateRanges(4, 0) = strEndSession
            
          Else
            'Loop through the array for the current calendar row, checking if any dates overlap.
            For i = 0 To UBound(avDateRanges, 2) Step 1
              
              'if the start or end dates 'equal' any other start orend dates then check if the sessions are also equal.
              If ((dtSD = avDateRanges(1, i)) And (strStartSession = avDateRanges(3, i))) _
                Or ((dtSD = avDateRanges(2, i)) And (strStartSession = avDateRanges(4, i))) _
                Or ((dtED = avDateRanges(1, i)) And (strEndSession = avDateRanges(3, i))) _
                Or ((dtED = avDateRanges(2, i)) And (strEndSession = avDateRanges(4, i))) Then
                mblnHasMultipleEvents = True
                MultipleCheck = True
                GoTo TidyUpAndExit
              End If
              
              If ((dtSD > avDateRanges(1, i)) And (dtED < avDateRanges(2, i))) _
                Or ((dtSD > avDateRanges(1, i)) And (dtSD < avDateRanges(2, i)) And (dtED > avDateRanges(2, i))) _
                Or ((dtED > avDateRanges(1, i)) And (dtED < avDateRanges(2, i)) And (dtSD < avDateRanges(1, i))) _
                Or ((dtSD < avDateRanges(1, i)) And (dtED > avDateRanges(2, i))) Then
                mblnHasMultipleEvents = True
                MultipleCheck = True
                GoTo TidyUpAndExit
              End If
            Next i
            
            intNewIndex = UBound(avDateRanges, 2) + 1
            ReDim Preserve avDateRanges(6, intNewIndex)
            avDateRanges(0, intNewIndex) = lngBaseID
            avDateRanges(1, intNewIndex) = dtSD
            avDateRanges(2, intNewIndex) = dtED
            avDateRanges(3, intNewIndex) = strStartSession
            avDateRanges(4, intNewIndex) = strEndSession
            
          End If
        End If
      
        .MoveNext
      Loop
    End If
    .Close
  End With

  mblnHasMultipleEvents = False

  MultipleCheck = True
  
TidyUpAndExit:
  Set rsMultiple = Nothing
  Exit Function
  
ErrorTrap:
  MultipleCheck = False
  GoTo TidyUpAndExit
  
End Function

Private Function GetCalendarReportDefinition() As Boolean
  
  ' Purpose : This function retrieves the basic definition details
  '           and stores it in module level variables
  
  On Error GoTo Error_Trap

  Dim rsTemp As Recordset
  
  Dim sSQL As String
  Dim sTable As String
  Dim sColumn As String
  Dim sDateInterval As String
  
  Dim i As Integer
  
  Dim rsIDs As ADODB.Recordset
  Dim blnOK As Boolean
  
  Dim intExprReturnType As Integer
  Dim objExpression As clsExprExpression
  
  mstrSQLIDs = vbNullString
  
  sSQL = "SELECT * FROM ASRSYSCalendarReports " & _
           "WHERE ID = " & CStr(mlngCalendarReportID) & " "
  
  Set rsTemp = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  With rsTemp
  
    If .BOF And .EOF Then
      GetCalendarReportDefinition = False
      mstrErrorString = "Could not find specified Calendar Report definition."
      Exit Function
    End If
    
    mstrCalendarReportsName = !Name
    mlngCalendarReportsBaseTable = !BaseTable
    mstrCalendarReportsBaseTableName = datGeneral.GetTableName(mlngCalendarReportsBaseTable)
    
    Dim pblnOK
    Dim objTableView As CTablePrivilege
    ' Check the user has permission to read the base table.
    pblnOK = False
    For Each objTableView In gcoTablePrivileges.Collection
      If (objTableView.TableID = mlngCalendarReportsBaseTable) And _
        (objTableView.AllowSelect) Then
        pblnOK = True
        Exit For
      End If
    Next objTableView
    Set objTableView = Nothing
    If Not pblnOK Then
      GetCalendarReportDefinition = False
      mstrErrorString = "You do not have permission to read the base table" & vbNewLine & "either directly or through any views."
      GoTo TidyAndExit
    End If
    
    
    mlngCalendarReportsAllRecords = !AllRecords
    mlngCalendarReportsPickListID = !picklist
    mlngCalendarReportsFilterID = !Filter
      
    mlngDescription1 = IIf(IsNull(!Description1), 0, !Description1)
    If mlngDescription1 > 0 Then
      mstrDescription1 = datGeneral.GetColumnName(!Description1)
      mblnDesc1IsDate = (datGeneral.GetDataType(mlngCalendarReportsBaseTable, mlngDescription1) = SQLDataType.sqlDate)
    Else
      mstrDescription1 = vbNullString
      mblnDesc1IsDate = False
    End If
    
    mlngDescription2 = IIf(IsNull(!Description2), 0, !Description2)
    If mlngDescription2 > 0 Then
      mstrDescription2 = datGeneral.GetColumnName(!Description2)
      mblnDesc2IsDate = (datGeneral.GetDataType(mlngCalendarReportsBaseTable, mlngDescription2) = SQLDataType.sqlDate)
    Else
      mstrDescription2 = vbNullString
      mblnDesc2IsDate = False
    End If
    
    mlngDescriptionExpr = IIf(IsNull(!DescriptionExpr), 0, !DescriptionExpr)
    If mlngDescriptionExpr > 0 Then
    
      Set objExpression = New clsExprExpression
      objExpression.ExpressionID = mlngDescriptionExpr
      objExpression.ConstructExpression
      
      'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
      ' when all we need is the expression return type (time saving measure).
      objExpression.ValidateExpression True, True
      
      If objExpression.ReturnType = 4 Then ' its date
        mblnDescExprIsDate = True
      Else
        mblnDescExprIsDate = False
      End If
      mlngBaseDescriptionType = objExpression.ReturnType
      Set objExpression = Nothing
      
      mstrDescriptionExpr = datGeneral.GetExpression(!DescriptionExpr)
    Else
      mlngBaseDescriptionType = -1
      mstrDescriptionExpr = vbNullString
      mblnDescExprIsDate = False
    End If
    
    mlngRegion = !Region
    If mlngRegion > 0 Then
      mstrRegion = datGeneral.GetColumnName(!Region)
      
    ElseIf (mlngCalendarReportsBaseTable = glngPersonnelTableID) _
        And (grtRegionType = rtStaticRegion) Then
        
        mlngRegion = glngBHolRegionID
        mstrRegion = gsBHolRegionColumnName
        
    Else
      mstrRegion = vbNullString
      
    End If
    
    mlngDescription1ID = IIf(IsNull(!Description1), 0, !Description1)
    mlngDescription2ID = IIf(IsNull(!Description2), 0, !Description2)
    mstrDescriptionSeparator = IIf(IsNull(!DescriptionSeparator), ", ", IIf(!DescriptionSeparator = "<None>", " ", !DescriptionSeparator))
    mblnGroupByDescription = IIf(!GroupByDesc, True, False)
    
    'create the events collection here so that the event filters can bee checked
    If Not GetEventsCollection Then
      GetCalendarReportDefinition = False
      Exit Function
    End If
    
    mlngStartDateExpr = IIf(IsNull(!StartDateExpr), 0, !StartDateExpr)
    mlngEndDateExpr = IIf(IsNull(!EndDateExpr), 0, !EndDateExpr)

    'validate all filters/picklist & calculationshere before, actually using them
    mblnDefinitionOwner = (LCase(Trim(gsUserName)) = LCase(Trim(!UserName)))
    If Not IsRecordSelectionValid Then
      GetCalendarReportDefinition = False
      Exit Function
    End If

    '************** Must do the dates stuff here *****************
    'calculate and store the start and end dates
    
    'START DATE
    Select Case !StartType
      Case 0
        mdtStartDate = !FixedStart
      Case 1
        mdtStartDate = CDate(Format(Now, DateFormat))
      Case 3
        mdtStartDate = datGeneral.GetValueForRecordIndependantCalc(mlngStartDateExpr)
    End Select
    
    'END DATE
    Select Case !EndType
      Case 0
        mdtEndDate = !FixedEnd
      Case 1
        mdtEndDate = CDate(Format(Now, DateFormat))
      Case 3
        mdtEndDate = datGeneral.GetValueForRecordIndependantCalc(mlngEndDateExpr)
    End Select
    
    
    If (!StartType = 2) And (!EndType = 2) Then
      'START DATE
      Select Case !StartPeriod
        Case 0: sDateInterval = "d"
        Case 1: sDateInterval = "ww"
        Case 2: sDateInterval = "m"
        Case 3: sDateInterval = "yyyy"
      End Select
      mdtStartDate = DateAdd(sDateInterval, CDbl(!StartFrequency), CDate(Format(Now, DateFormat)))
      
      'END DATE
      Select Case !EndPeriod
        Case 0: sDateInterval = "d"
        Case 1: sDateInterval = "ww"
        Case 2: sDateInterval = "m"
        Case 3: sDateInterval = "yyyy"
      End Select
      mdtEndDate = DateAdd(sDateInterval, CDbl(!EndFrequency), CDate(Format(Now, DateFormat)))
    
    ElseIf !StartType = 2 And !EndType <> 2 Then
      'START DATE
      Select Case !StartPeriod
        Case 0: sDateInterval = "d"
        Case 1: sDateInterval = "ww"
        Case 2: sDateInterval = "m"
        Case 3: sDateInterval = "yyyy"
      End Select
      mdtStartDate = DateAdd(sDateInterval, CDbl(!StartFrequency), mdtEndDate)
      
    ElseIf !EndType = 2 And !StartType <> 2 Then
      'END DATE
      Select Case !EndPeriod
        Case 0: sDateInterval = "d"
        Case 1: sDateInterval = "ww"
        Case 2: sDateInterval = "m"
        Case 3: sDateInterval = "yyyy"
      End Select
      mdtEndDate = DateAdd(sDateInterval, CDbl(!EndFrequency), mdtStartDate)
      
    End If
    
    If mdtStartDate > mdtEndDate Then
      mstrErrorString = "The report end date is before the report start date."
      GetCalendarReportDefinition = False
      Exit Function
    End If
    
    '************************************************
    
    mblnShowBankHolidays = !ShowBankHolidays
    mblnShowCaptions = !ShowCaptions
    mblnShowWeekends = !ShowWeekends
    mblnIncludeWorkingDaysOnly = !IncludeWorkingDaysOnly
    mblnIncludeBankHolidays = !IncludeBankHolidays
    mbStartOnCurrentMonth = !StartOnCurrentMonth
    mblnCustomReportsPrintFilterHeader = !PrintFilterHeader
    
    'New Default Output Variables
    mlngOutputFormat = !OutputFormat
    mblnOutputScreen = !OutputScreen
    mblnOutputPrinter = !OutputPrinter
    mstrOutputPrinterName = !OutputPrinterName
    mblnOutputSave = !OutputSave
    mlngOutputSaveExisting = !OutputSaveExisting
    mlngOutputSaveFormat = !OutputSaveFormat
    mblnOutputEmail = !OutputEmail
    mlngOutputEmailAddr = !OutputEmailAddr
    mstrOutputEmailSubject = !OutputEmailSubject
    mstrOutputEmailAttachAs = IIf(IsNull(!OutputEmailAttachAs), vbNullString, !OutputEmailAttachAs)
    mlngOutputEmailFileFormat = !OutputEmailFileFormat
    mstrOutputFilename = !OutputFilename

    mblnPreviewOnScreen = (!OutputPreview Or (mlngOutputFormat = fmtDataOnly And mblnOutputScreen))

    If mlngSingleRecordID > 0 Then
      mstrSQLIDs = CStr(mlngSingleRecordID)
    
    ElseIf mlngCalendarReportsPickListID > 0 Then
      Set rsIDs = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngCalendarReportsPickListID, adOpenForwardOnly, adLockReadOnly)
      
      If rsIDs.BOF And rsIDs.EOF Then
        mstrErrorString = "The selected picklist contains no records."
        GetCalendarReportDefinition = False
        Exit Function
      End If
  
      Do While Not rsIDs.EOF
        mstrSQLIDs = mstrSQLIDs & IIf(Len(mstrSQLIDs) > 0, ", ", "") & rsIDs.Fields(0)
        rsIDs.MoveNext
      Loop
      rsIDs.Close
      Set rsIDs = Nothing
      
    ElseIf mlngCalendarReportsFilterID > 0 Then
      blnOK = datGeneral.FilteredIDs(mlngCalendarReportsFilterID, mstrFilteredIDs)
     
      ' Generate any UDFs that are used in this filter
      If blnOK And gbEnableUDFFunctions Then
        datGeneral.FilterUDFs mlngCalendarReportsFilterID, mastrUDFsRequired()
      End If
      
      If blnOK Then
      
        UDFFunctions mastrUDFsRequired, True
        Set rsIDs = mclsData.OpenRecordset(mstrFilteredIDs, adOpenForwardOnly, adLockReadOnly)
        UDFFunctions mastrUDFsRequired, False
        
        If rsIDs.BOF And rsIDs.EOF Then
          GetCalendarReportDefinition = False
          mstrErrorString = "The base table filter returned no records."
          mblnNoRecords = True
          Exit Function
        End If
        
        Do While Not rsIDs.EOF
          mstrSQLIDs = mstrSQLIDs & IIf(Len(mstrSQLIDs) > 0, ", ", "") & rsIDs.Fields(0)
          rsIDs.MoveNext
        Loop
        rsIDs.Close
        Set rsIDs = Nothing

      Else
        ' Permission denied on something in the filter.
        mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngCalendarReportsFilterID) & "' filter."
        GetCalendarReportDefinition = False
        Exit Function
      End If
      
    End If
  
  End With
  rsTemp.Close

  mstrBaseTableRealSource = gcoTablePrivileges.Item(mstrCalendarReportsBaseTableName).RealSource

  GetCalendarReportDefinition = True
  
TidyAndExit:
  Set rsTemp = Nothing
  Exit Function

Error_Trap:
  GetCalendarReportDefinition = False
  mstrErrorString = "Error whilst reading the Calendar Report definition." & vbNewLine & Err.Description
  Resume TidyAndExit

End Function
Private Function GetEventsCollection() As Boolean

  On Error GoTo Error_Trap
  
  Dim sSQL As String
  Dim intTemp As Integer
  Dim rsTemp As ADODB.Recordset
  Dim lngTableID As Long
  
  Dim sTempTableName As String
  Dim sTempStartDateName As String
  Dim sTempStartSessionName As String
  Dim sTempEndDateName As String
  Dim sTempEndSessionName As String
  Dim sTempDurationName As String
  Dim sTempLegendTableName As String
  Dim sTempLegendColumnName As String
  Dim sTempLegendCodeName As String
  Dim sTempLegendEventTypeName As String
  Dim sTempDesc1Name As String
  Dim sTempDesc2Name As String
  
  ' Get the column information from the Details table, in order
  sSQL = "SELECT * FROM AsrSysCalendarReportEvents WHERE " & _
               "CalendarReportID = " & CStr(mlngCalendarReportID) & _
               " ORDER BY Name ASC "
  
  Set rsTemp = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  With rsTemp
    If .BOF And .EOF Then
      mstrErrorString = "No events found in the specified Calendar Report definition." & vbNewLine & "Please remove this definition and create a new one."
      GetEventsCollection = False
      GoTo TidyUpAndExit
    End If
  
    Set mcolEvents = New clsCalendarEvents
    
    Do Until .EOF
    
      sTempTableName = datGeneral.GetTableName(!TableID)

      If !EventStartDateID > 0 Then
        sTempStartDateName = datGeneral.GetColumnName(!EventStartDateID)
      Else
        MsgBox "Cannot load the event information for this Calendar Report", vbExclamation + vbOKOnly, "Calendar Reports"
        GetEventsCollection = False
        GoTo TidyUpAndExit
      End If
      
      If !EventStartSessionID > 0 Then
        sTempStartSessionName = datGeneral.GetColumnName(!EventStartSessionID)
      Else
        sTempStartSessionName = vbNullString
      End If
      
      If !EventEndDateID > 0 Then
        sTempEndDateName = datGeneral.GetColumnName(!EventEndDateID)
      Else
        sTempEndDateName = vbNullString
      End If
      
      If !EventEndSessionID > 0 Then
        sTempEndSessionName = datGeneral.GetColumnName(!EventEndSessionID)
      Else
        sTempEndSessionName = vbNullString
      End If
      
      If !EventDurationID > 0 Then
        sTempDurationName = datGeneral.GetColumnName(!EventDurationID)
      Else
        sTempDurationName = vbNullString
      End If
      
      If !LegendLookupTableID > 0 Then
        sTempLegendTableName = datGeneral.GetTableName(!LegendLookupTableID)
      Else
        sTempLegendTableName = vbNullString
      End If
      
      If !LegendLookupColumnID > 0 Then
        sTempLegendColumnName = datGeneral.GetColumnName(!LegendLookupColumnID)
      Else
        sTempLegendColumnName = vbNullString
      End If
      
      If !LegendLookupCodeID > 0 Then
        sTempLegendCodeName = datGeneral.GetColumnName(!LegendLookupCodeID)
      Else
        sTempLegendCodeName = vbNullString
      End If
      
      If !LegendEventColumnID > 0 Then
        sTempLegendEventTypeName = datGeneral.GetColumnName(!LegendEventColumnID)
      Else
        sTempLegendEventTypeName = vbNullString
      End If
      
      If !EventDesc1ColumnID > 0 Then
        sTempDesc1Name = datGeneral.GetColumnName(!EventDesc1ColumnID)
      Else
        sTempDesc1Name = vbNullString
      End If
      
      If !EventDesc2ColumnID > 0 Then
        sTempDesc2Name = datGeneral.GetColumnName(!EventDesc2ColumnID)
      Else
        sTempDesc2Name = vbNullString
      End If

      mcolEvents.Add !EventKey, !Name, _
                      !TableID, sTempTableName, _
                      !FilterID, _
                      !EventStartDateID, sTempStartDateName, _
                      !EventStartSessionID, sTempStartSessionName, _
                      !EventEndDateID, sTempEndDateName, _
                      !EventEndSessionID, sTempEndSessionName, _
                      !EventDurationID, sTempDurationName, _
                      !LegendType, !LegendCharacter, 0, _
                      !LegendLookupTableID, sTempLegendTableName, _
                      !LegendLookupColumnID, sTempLegendColumnName, _
                      !LegendLookupCodeID, sTempLegendCodeName, _
                      !LegendEventColumnID, sTempLegendEventTypeName, _
                      !EventDesc1ColumnID, sTempDesc1Name, _
                      !EventDesc2ColumnID, sTempDesc2Name
      
      .MoveNext
    Loop
    .Close
  End With
 
  GetEventsCollection = True

TidyUpAndExit:
  Set rsTemp = Nothing
  Exit Function
  
Error_Trap:
  GetEventsCollection = False
  mstrErrorString = "Error whilst retrieving the event details recordsets'." & vbNewLine & Err.Description
  GoTo TidyUpAndExit
  
End Function
Private Function GenerateSQL() As Boolean
  
  Dim fOK As Boolean
  Dim objEvent As clsCalendarEvent
  Dim rsLegendBreakdown As ADODB.Recordset
  
  Dim strSQL As String
  Dim strDynamicKey As String
  Dim strDynamicName As String
  
  fOK = True
  mintDynamicEventCount = 0
  
  'loop through the events col and UNION the Event queries together
  For Each objEvent In mcolEvents.Collection
    
    mblnHasEventFilterIDs = False
    mstrEventFilterIDs = vbNullString
    
    With objEvent
      If (.LegendType = 1) And (.LegendTableID > 0) Then
        'Event is using a lookup table to find the calendar code for the event.
        'Therefore use the unique types from the legend information.
        
        strSQL = "SELECT DISTINCT " & .LegendTableName & "." & .LegendColumnName & vbNewLine & _
                 " FROM " & .LegendTableName & vbNewLine
        
        Set rsLegendBreakdown = New ADODB.Recordset
        rsLegendBreakdown.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
        
        If rsLegendBreakdown.BOF And rsLegendBreakdown.EOF Then
          mstrErrorString = "The '" & .LegendTableName & "' event lookup table contains no records."
          GenerateSQL = False
          Exit Function
        End If
        
        rsLegendBreakdown.MoveFirst
        Do While Not rsLegendBreakdown.EOF
          mintDynamicEventCount = mintDynamicEventCount + 1
          
          strDynamicKey = "DYNAMICEVENT" & CStr(mintDynamicEventCount)
          strDynamicName = Replace(IIf(IsNull(rsLegendBreakdown.Fields(CStr(.LegendColumnName)).Value), "", rsLegendBreakdown.Fields(CStr(.LegendColumnName)).Value), "'", "''")
          
          mstrSQLDynamicLegendWhere = vbNullString
          
          If fOK Then fOK = GenerateSQLEvent(objEvent.Key, strDynamicKey, strDynamicName)
          
          If Not fOK Then
            GenerateSQL = False
            Exit Function
          End If
          
          'mstrSQL = mstrSQL & mstrSQLEvent & " UNION "
          fOK = InsertIntoTempTable(mstrSQLEvent)
          
          mstrSQLEvent = vbNullString

          rsLegendBreakdown.MoveNext
        Loop
        rsLegendBreakdown.Close
        Set rsLegendBreakdown = Nothing
        
      Else
        If fOK Then fOK = GenerateSQLEvent(objEvent.Key, vbNullString, vbNullString)
        
        If Not fOK Then
          GenerateSQL = False
          Exit Function
        End If
        
'        Open App.Path & "\calendarreporteventSQL.sql" For Output As #1
'
'        Print #1, mstrSQLEvent
'
'        Close #1
        
        If fOK Then fOK = InsertIntoTempTable(mstrSQLEvent)
        
        mstrSQLEvent = vbNullString
      
      End If
    End With
    
  Next objEvent
  
  'remove the last UNION command from the SQL string
'  mstrSQL = Left(mstrSQL, Len(mstrSQL) - 7)
  
  GenerateSQL = fOK

TidyUpAndExit:
  Set objEvent = Nothing

End Function

Private Function GenerateSQLSelect(pstrEventKey As String, _
                                    pstrDynamicKey As String, _
                                    pstrDynamicName As String) As Boolean
  
  ' Purpose : This function compiles the SQLSelect string looping
  '           thru the column details recordset.
  
  On Error GoTo GenerateSQLSelect_ERROR
  
  Dim objEvent As clsCalendarEvent

  Dim strColList As HRProDataMgr.clsStringBuilder
  Dim strBaseColList As String
  
  Dim strLookupTableName As String
  Dim strLookupColumnName As String
  Dim strLookupCodeName As String
  Dim strEventType As String
  Dim strLegendSQL As String
  Dim strTableColumn As String
  Dim strRegionSQL As String
  Dim lngTempTableID As Long
  Dim strTempTableName As String
  Dim strTempColumnName As String
  
  Dim strTempStartSession As String
  Dim strTempEndSession As String
  
  Set strColList = New HRProDataMgr.clsStringBuilder

  'Get the Base ID column values so that these can be used in the group by clause when checking
  'for multiple events in MultipleCheck().
  mstrSQLCreateTable.Append "[BaseID] [Integer] NOT NULL, "
  strColList.Append "[" & mstrBaseTableRealSource & "].[ID] AS 'BaseID', " & vbNewLine
  
  mblnCheckingDescColumn = True
  
  If mlngDescription1 > 0 Then
    If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, mstrDescription1, strTableColumn) Then
      strColList.Append " CONVERT(varchar," & strTableColumn & ") AS 'Description1', " & vbNewLine
      strBaseColList = strBaseColList & " CONVERT(varchar," & strTableColumn & ") AS 'Description1', " & vbNewLine
      strTableColumn = vbNullString
    Else
      GenerateSQLSelect = False
      GoTo TidyUpAndExit
    End If
  Else
    strBaseColList = strBaseColList & "NULL AS 'Description1', " & vbNewLine
    strColList.Append "NULL AS 'Description1', " & vbNewLine
  End If
  mstrSQLCreateTable.Append "[Description1] [varchar] (MAX) NULL, "

  If mlngDescription2 > 0 Then
    If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, mstrDescription2, strTableColumn) Then
      strColList.Append "CONVERT(varchar," & strTableColumn & ") AS 'Description2', " & vbNewLine
      strBaseColList = strBaseColList & "CONVERT(varchar," & strTableColumn & ") AS 'Description2', " & vbNewLine
      strTableColumn = vbNullString
    Else
      GenerateSQLSelect = False
      GoTo TidyUpAndExit
    End If
  Else
    strBaseColList = strBaseColList & "NULL AS 'Description2', " & vbNewLine
    strColList.Append "NULL AS 'Description2', " & vbNewLine
  End If
  mstrSQLCreateTable.Append "[Description2] [varchar] (MAX) NULL, "
  
  If mlngDescriptionExpr > 0 Then
    If mblnDescCalcCode Then
      strColList.Append " " & mstrDescCalcCode & " AS 'DescriptionExpr', " & vbNewLine
      strBaseColList = strBaseColList & " " & mstrDescCalcCode & " AS 'DescriptionExpr', " & vbNewLine

    Else
      If CheckCalculationPermissions(mlngCalendarReportsBaseTable, mlngDescriptionExpr, strTableColumn) Then
        mstrDescCalcCode = strTableColumn
        mblnDescCalcCode = True
        strColList.Append " " & strTableColumn & " AS 'DescriptionExpr', " & vbNewLine
        strBaseColList = strBaseColList & " " & strTableColumn & " AS 'DescriptionExpr', " & vbNewLine
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        GoTo TidyUpAndExit
      End If
      
    End If
  Else
    strBaseColList = strBaseColList & "NULL AS 'DescriptionExpr', " & vbNewLine
    strColList.Append "NULL AS 'DescriptionExpr', " & vbNewLine
  End If
  
  mblnCheckingDescColumn = False

  'need to set the type of the expression column for the CREAT TABLE...statement.
  Select Case mlngBaseDescriptionType
    Case ExpressionValueTypes.giEXPRVALUE_NUMERIC, ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC
      mstrSQLCreateTable.Append "[DescriptionExpr] [float] NULL, "
    Case Else
      mstrSQLCreateTable.Append "[DescriptionExpr] [varchar] (MAX) NULL, "
  End Select
  
  
  Set objEvent = mcolEvents.Item(pstrEventKey)
  With objEvent
    If pstrDynamicKey <> vbNullString Then
      strColList.Append "'" & Replace(pstrDynamicKey, "'", "''") & "' AS '?ID_EventID', " & vbNewLine
    Else
      strColList.Append "'" & Replace(.Key, "'", "''") & "' AS '?ID_EventID', " & vbNewLine
    End If
    mstrSQLCreateTable.Append "[?ID_EventID] [varchar] (255) NULL, "
    
    If pstrDynamicName <> vbNullString Then
      strColList.Append "'" & Replace(pstrDynamicName, "'", "''") & "' AS 'Name', " & vbNewLine
    Else
      strColList.Append "'" & Replace(.Name, "'", "''") & "' AS 'Name', " & vbNewLine
    End If
    mstrSQLCreateTable.Append "[Name] [varchar] (255) NULL, "

    
    '****************************************************************************
    mlngEventViewColumn = .StartDateID
    If CheckColumnPermissions(.TableID, .TableName, .StartDateName, strTableColumn) Then
      strColList.Append strTableColumn & " AS 'StartDate', " & vbNewLine
      mstrSQLBaseStartDateColumn = strTableColumn
      strTableColumn = vbNullString
    Else
      GenerateSQLSelect = False
      Exit Function
    End If
    
    mstrSQLCreateTable.Append "[StartDate] [datetime] NULL, "
    
    If .StartSessionID > 0 Then
      If CheckColumnPermissions(.TableID, .TableName, .StartSessionName, strTableColumn) Then
        mstrSQLBaseStartSessionColumn = strTableColumn
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        Exit Function
      End If
    End If
    
    If .EndDateID > 0 Then
      mlngEventViewColumn = .EndDateID
      If CheckColumnPermissions(.TableID, .TableName, .EndDateName, strTableColumn) Then
        mstrSQLBaseEndDateColumn = strTableColumn
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        Exit Function
      End If
    End If
   
    If .EndSessionID > 0 Then
      If CheckColumnPermissions(.TableID, .TableName, .EndSessionName, strTableColumn) Then
        mstrSQLBaseEndSessionColumn = strTableColumn
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        Exit Function
      End If
    End If
    
    If .DurationID > 0 Then
      mlngEventViewColumn = .DurationID
      If CheckColumnPermissions(.TableID, .TableName, .DurationName, strTableColumn) Then
        mstrSQLBaseDurationColumn = strTableColumn
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        Exit Function
      End If

      If .StartSessionID > 0 Then
        strColList.Append mstrSQLBaseStartSessionColumn & " AS 'StartSession', " & vbNewLine
      Else
        strColList.Append "'AM' AS 'StartSession', " & vbNewLine
        mstrSQLBaseStartSessionColumn = "'AM'"
      End If
     
'      strColList.Append " DATEADD(dd, " & mstrSQLBaseDurationColumn & " - 1 , " & mstrSQLBaseStartDateColumn & ") AS 'EndDate', " & vbNewLine
      strColList.Append "CASE " & vbNewLine _
       & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '1' " & vbNewLine _
       & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '2' " & vbNewLine _
       & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '3' " & vbNewLine _
       & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '4' " & vbNewLine _
       & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '5') THEN " & vbNewLine _
       & "         DATEADD(dd " & vbNewLine _
       & "                 , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine _
       & "                           , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 )) " & vbNewLine _
       & "                 , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine _
       & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '6' " & vbNewLine _
       & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '7' " & vbNewLine _
       & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '8' " & vbNewLine _
       & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '9') " & vbNewLine _
       & "          AND (" & mstrSQLBaseStartSessionColumn & " = 'AM') THEN " & vbNewLine _
       & "         DATEADD(dd " & vbNewLine _
       & "                 , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine _
       & "                         , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 )) " & vbNewLine _
       & "                 , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine _
       & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '6' " & vbNewLine _
       & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '7' " & vbNewLine _
       & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '8' " & vbNewLine _
       & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '9') " & vbNewLine _
       & "          AND (" & mstrSQLBaseStartSessionColumn & " = 'PM') THEN " & vbNewLine
            
      strColList.Append "           DATEADD(dd " & vbNewLine _
       & "                 , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine _
       & "                         , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 ) + 1) " & vbNewLine _
       & "                 , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine _
       & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '0') " & vbNewLine _
       & "          AND (" & mstrSQLBaseStartSessionColumn & " = 'AM') THEN " & vbNewLine _
       & "           DATEADD(dd " & vbNewLine _
       & "                   , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine _
       & "                         , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 ) - 1) " & vbNewLine _
       & "                   , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine _
       & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '0') " & vbNewLine _
       & "          AND (" & mstrSQLBaseStartSessionColumn & " = 'PM') THEN " & vbNewLine _
       & "           DATEADD(dd " & vbNewLine _
       & "                 , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine _
       & "                               , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 )) " & vbNewLine _
       & "                 , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine _
       & "END AS 'EndDate', " & vbNewLine

      If .EndSessionID > 0 Then
        strColList.Append mstrSQLBaseEndSessionColumn & " AS 'EndSession', " & vbNewLine
      Else
'        strColList.Append "'PM' AS 'EndSession'," & vbNewLine

        strColList.Append "CASE" _
         & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '1' " & vbNewLine _
         & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '2' " & vbNewLine _
         & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '3' " & vbNewLine _
         & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '4' " & vbNewLine _
         & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '5') THEN " & vbNewLine _
         & "           " & mstrSQLBaseStartSessionColumn & " " & vbNewLine
        
        strColList.Append "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '6' " & vbNewLine _
         & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '7' " & vbNewLine _
         & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '8' " & vbNewLine _
         & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '9' " & vbNewLine _
         & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '0') " & vbNewLine _
         & "          AND (" & mstrSQLBaseStartSessionColumn & " = 'AM') THEN " & vbNewLine _
         & "           'PM'  " & vbNewLine
        
        strColList.Append "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '6' " & vbNewLine _
         & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '7' " & vbNewLine _
         & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '8' " & vbNewLine _
         & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '9' " & vbNewLine _
         & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '0') " & vbNewLine _
         & "          AND (" & mstrSQLBaseStartSessionColumn & " = 'PM') THEN " & vbNewLine _
         & "           'AM' " & vbNewLine _
         & "END AS 'EndSession', " & vbNewLine
      End If
      
      strColList.Append mstrSQLBaseDurationColumn & " AS 'Duration', " & vbNewLine

    ElseIf .EndDateID > 0 Then
    
      If .StartSessionID > 0 Then
        strColList.Append mstrSQLBaseStartSessionColumn & " AS 'StartSession', " & vbNewLine
        strTempStartSession = mstrSQLBaseStartSessionColumn
      Else
        strColList.Append "'AM' AS 'StartSession'," & vbNewLine
        strTempStartSession = "'AM'"
      End If
    
      strColList.Append mstrSQLBaseEndDateColumn & " AS 'EndDate', " & vbNewLine
       
      If .EndSessionID > 0 Then
        strColList.Append mstrSQLBaseEndSessionColumn & " AS 'EndSession', " & vbNewLine
        strTempEndSession = mstrSQLBaseEndSessionColumn
      Else
        strColList.Append "'PM' AS 'EndSession'," & vbNewLine
        strTempEndSession = "'PM'"
      End If

      strColList.Append " CASE " & vbNewLine _
       & " WHEN " & strTempStartSession & " = " & strTempEndSession & vbNewLine _
       & "   THEN CONVERT(float,(DATEDIFF(dd, " & mstrSQLBaseStartDateColumn & ", " & mstrSQLBaseEndDateColumn & ") + 0.5)) " & vbNewLine _
       & " ELSE " & vbNewLine _
       & "   CONVERT(float,(DATEDIFF(dd, " & mstrSQLBaseStartDateColumn & ", " & mstrSQLBaseEndDateColumn & ") + 1)) " & vbNewLine _
       & " END AS 'Duration'," & vbNewLine
    Else
    
      If .StartSessionID > 0 Then
        strColList.Append mstrSQLBaseStartSessionColumn & " AS 'StartSession', " & vbNewLine
      Else
        strColList.Append "'AM' AS 'StartSession'," & vbNewLine
      End If

      strColList.Append mstrSQLBaseStartDateColumn & " AS 'EndDate', " & vbNewLine
      
      If .StartSessionID > 0 Then
        strColList.Append mstrSQLBaseStartSessionColumn & " AS 'EndSession'," & vbNewLine _
          & " 0.5 AS 'Duration', " & vbNewLine
      Else
        strColList.Append "'PM' AS 'EndSession'," & vbNewLine _
         & " 1 AS 'Duration', " & vbNewLine
      End If
    
    End If
    mstrSQLCreateTable.Append "[StartSession] [varchar] (255) NULL, " _
      & "[EndDate] [datetime] NULL, " _
      & "[EndSession] [varchar] (255) NULL, " _
      & "[Duration] [float] NULL, " _
    '****************************************************************************
    
    mblnCheckingDescColumn = True

    If .Description1ID > 0 Then
    
      lngTempTableID = datGeneral.GetColumnTable(.Description1ID)
      strTempTableName = datGeneral.GetColumnTableName(.Description1ID)
      strTempColumnName = .Description1Name
      If CheckColumnPermissions(lngTempTableID, strTempTableName, strTempColumnName, strTableColumn) Then
        strColList.Append .Description1ID & " AS 'EventDescription1ColumnID', " & vbNewLine _
          & "'" & .Description1Name & "' AS 'EventDescription1Column', " & vbNewLine
        
        'TM20030407 Fault 5259 - if logic field...convert to 'Y' or 'N' accordingly.
        If datGeneral.GetDataType(lngTempTableID, .Description1ID) = SQLDataType.sqlBoolean Then
          strColList.Append "CASE " & strTableColumn & " WHEN 1 THEN 'Y' ELSE 'N' END AS 'EventDescription1', " & vbNewLine
        Else
          strColList.Append "CONVERT(varchar(MAX)," & strTableColumn & ") AS 'EventDescription1', " & vbNewLine
        End If
        
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        Exit Function
      End If
    Else
      strColList.Append "NULL AS 'EventDescription1ColumnID', " & vbNewLine _
        & "NULL AS 'EventDescription1Column', " & vbNewLine _
        & "NULL AS 'EventDescription1', " & vbNewLine
    End If
    mstrSQLCreateTable.Append "[EventDescription1ColumnID] [int] NULL, " _
      & "[EventDescription1Column] [varchar] (MAX) NULL, " _
      & "[EventDescription1] [varchar] (MAX) NULL, "

    If .Description2ID > 0 Then
      lngTempTableID = datGeneral.GetColumnTable(.Description2ID)
      strTempTableName = datGeneral.GetColumnTableName(.Description2ID)
      strTempColumnName = .Description2Name
      If CheckColumnPermissions(lngTempTableID, strTempTableName, strTempColumnName, strTableColumn) Then
        strColList.Append .Description2ID & " AS 'EventDescription2ColumnID', " & vbNewLine _
          & "'" & .Description2Name & "' AS 'EventDescription2Column', " & vbNewLine
        
        'TM20030407 Fault 5259 - if logic field...convert to 'Y' or 'N' accordingly.
        If datGeneral.GetDataType(lngTempTableID, .Description2ID) = SQLDataType.sqlBoolean Then
          strColList.Append "CASE " & strTableColumn & " WHEN 1 THEN 'Y' ELSE 'N' END AS 'EventDescription2', " & vbNewLine
        Else
          strColList.Append "CONVERT(varchar(MAX)," & strTableColumn & ") AS 'EventDescription2', " & vbNewLine
        End If
        
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        Exit Function
      End If
    Else
      strColList.Append "NULL AS 'EventDescription2ColumnID', " & vbNewLine _
        & "NULL AS 'EventDescription2Column', " & vbNewLine _
        & "NULL AS 'EventDescription2', " & vbNewLine
    End If
    mstrSQLCreateTable.Append "[EventDescription2ColumnID] [int] NULL, " _
      & "[EventDescription2Column] [varchar] (MAX) NULL, " _
      & "[EventDescription2] [varchar] (MAX) NULL, "
   
    mblnCheckingDescColumn = False
  
    If .LegendType = 1 Then
      If CheckColumnPermissions(.LegendTableID, .LegendTableName, .LegendCodeName, strTableColumn) Then
        strLegendSQL = "LEFT((SELECT TOP 1 " & strTableColumn
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        Exit Function
      End If
      
      strLegendSQL = strLegendSQL & " FROM " & .LegendTableName
      
      If CheckColumnPermissions(.LegendTableID, .LegendTableName, .LegendColumnName, strTableColumn) Then
        strLegendSQL = strLegendSQL & " WHERE " & strTableColumn
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        Exit Function
      End If
      
      If CheckColumnPermissions(.TableID, .TableName, .LegendEventTypeName, strTableColumn) Then
        strLegendSQL = strLegendSQL & " = " & strTableColumn & "), 2) AS 'Legend', "
        mstrSQLDynamicLegendWhere = strTableColumn & " = '" & pstrDynamicName & "' "
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        Exit Function
      End If
      
    Else
      strLegendSQL = "'" & Replace(Left(.LegendCharacter, 2), "'", "''") & "' AS 'Legend', "
    
    End If
     
    strColList.Append strLegendSQL & vbNewLine
    mstrSQLCreateTable.Append "[Legend] [varchar] (255) NULL, "
  End With
  
  'Add the static region column if required.
  If mlngRegion > 0 Then
    'TM01042004 Fault 8428
    mblnCheckingRegionColumn = True
    If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, mstrRegion, strTableColumn) Then
      strColList.Append "CONVERT(varchar," & strTableColumn & ") AS 'Region', " & vbNewLine
      strBaseColList = strBaseColList & "CONVERT(varchar," & strTableColumn & ") AS 'Region', " & vbNewLine
      strTableColumn = vbNullString
    Else
      'TM19112004 Fault 8942
      strColList.Append "NULL AS 'Region', "
      strBaseColList = strBaseColList & "NULL AS 'Region', "
'      GenerateSQLSelect = False
'      GoTo TidyUpAndExit
    End If
    'TM01042004 Fault 8428
    mblnCheckingRegionColumn = False
  Else
    strColList.Append "NULL AS 'Region', "
    strBaseColList = strBaseColList & "NULL AS 'Region', "
  End If
  mstrSQLCreateTable.Append "[Region] [varchar] (255) NULL, "
  
  If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, "ID", strTableColumn) Then
    strColList.Append strTableColumn & " AS '?ID_" & mstrCalendarReportsBaseTableName & "', " & vbNewLine
    strBaseColList = strBaseColList & strTableColumn & " AS '?ID_" & mstrCalendarReportsBaseTableName & "', " & vbNewLine
    strTableColumn = vbNullString
  Else
    GenerateSQLSelect = False
    GoTo TidyUpAndExit
  End If
  mstrSQLCreateTable.Append "[?ID_" & mstrCalendarReportsBaseTableName & "] [varchar] (255) NULL, "

  'Add the static Working Pattern column if required.
  If (mlngCalendarReportsBaseTable = glngPersonnelTableID) _
      And (gwptWorkingPatternType = wptStaticWPattern) _
      And (Not mblnGroupByDescription) Then
    If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, gsPersonnelWorkingPatternColumnName, strTableColumn) Then
      strColList.Append "CONVERT(varchar," & strTableColumn & ") AS 'Working_Pattern', " & vbNewLine
      strBaseColList = strBaseColList & "CONVERT(varchar," & strTableColumn & ") AS 'Working_Pattern', " & vbNewLine
      strTableColumn = vbNullString
    Else
      'TM06122004 Fault 9578
      strColList.Append "NULL AS 'Working_Pattern', "
      strBaseColList = strBaseColList & "NULL AS 'Working_Pattern', "
'      GenerateSQLSelect = False
'      GoTo TidyUpAndExit
    End If
  Else
    strColList.Append "NULL AS 'Working_Pattern', "
    strBaseColList = strBaseColList & "NULL AS 'Working_Pattern', "
  End If
  mstrSQLCreateTable.Append "[Working_Pattern] [varchar] (255) NULL, "

  Dim intOrderCount As Integer
  Dim strOrderColumn As String
  
  mstrSQLOrderList = vbNullString
  
  For intOrderCount = 1 To UBound(mvarSortOrder, 2) Step 1
    strOrderColumn = mvarSortOrder(1, intOrderCount)
    strTableColumn = vbNullString
    If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, strOrderColumn, strTableColumn) Then
      strColList.Append "CONVERT(varchar, " & strTableColumn & ") AS 'ORDER_" & CStr(intOrderCount) & "',"
      strBaseColList = strBaseColList & strTableColumn & " AS 'ORDER_" & CStr(intOrderCount) & "',"
      mstrSQLOrderList = mstrSQLOrderList & strTableColumn & " " & mvarSortOrder(2, intOrderCount) & ","
      strTableColumn = vbNullString
      If Not mblnOrderByCreated Then
        mstrSQLOrderBy = mstrSQLOrderBy & " [ORDER_" & CStr(intOrderCount) & "] " & mvarSortOrder(2, intOrderCount) & ","
      End If
    Else
      GenerateSQLSelect = False
      GoTo TidyUpAndExit
    End If
    mstrSQLCreateTable.Append "[ORDER_" & CStr(intOrderCount) & "] [varchar] (255) NULL,"
  Next intOrderCount
  
  If Not mblnOrderByCreated Then
    mstrSQLOrderBy = Left(mstrSQLOrderBy, (Len(mstrSQLOrderBy) - 1))
    mblnOrderByCreated = True
  End If
  mstrSQLCreateTable.TheString = Left(mstrSQLCreateTable.ToString, (mstrSQLCreateTable.Length - 1))
  
  strColList.TheString = Left(strColList.ToString, strColList.Length - 1)
  strBaseColList = Left(strBaseColList, Len(strBaseColList) - 1)
  mstrSQLOrderList = Left(mstrSQLOrderList, Len(mstrSQLOrderList) - 1)
  
  ' Start off the select statement
  mstrSQLSelect = "SELECT "
  mstrSQLSelect = mstrSQLSelect & strColList.ToString
  
  mstrSQLBaseData = "SELECT "
  mstrSQLBaseData = mstrSQLBaseData & strBaseColList
  
  GenerateSQLSelect = True

TidyUpAndExit:
  Set objEvent = Nothing
  Exit Function
  
GenerateSQLSelect_ERROR:
  GenerateSQLSelect = False
  mstrErrorString = "Error whilst generating SQL Select statement." & vbNewLine & Err.Description
  GoTo TidyUpAndExit
  
End Function

Private Function CreateTempTable() As Boolean
  
  Dim strSQL As String
  
  mstrTempTableName = datGeneral.UniqueSQLObjectName("ASRSysTempCalendarReport", 3)
  
  strSQL = "CREATE TABLE [" & gsUserName & "].[" & mstrTempTableName & "] (" _
    & mstrSQLCreateTable.ToString & ")"
  
  mclsData.ExecuteSql strSQL
  
  mblnTempTableCreated = True
  CreateTempTable = True
  
TidyUpAndExit:
  Exit Function
  
CreateTempTable_ERROR:
  CreateTempTable = False
  mstrErrorString = "Error whilst creating Temporary Table." & vbNewLine & Err.Description
  GoTo TidyUpAndExit
  
End Function

Private Function GenerateSQLFrom() As Boolean
  Dim iLoop As Integer
  Dim pobjTableView As CTablePrivilege
  
  Set pobjTableView = New CTablePrivilege
  
  mstrSQLFrom = "FROM " & mstrBaseTableRealSource & vbNewLine
  mstrSQLBaseData = mstrSQLBaseData & " FROM " & mstrBaseTableRealSource & vbNewLine
  
  GenerateSQLFrom = True

TidyUpAndExit:
  Set pobjTableView = Nothing
  Exit Function
  
GenerateSQLFrom_ERROR:
  GenerateSQLFrom = False
  mstrErrorString = "Error in GenerateSQLFrom." & vbNewLine & Err.Description
  GoTo TidyUpAndExit
  
End Function
Private Function GenerateSQLJoin(pstrEventKey As String, _
                                    pstrDynamicKey As String) As Boolean

  On Error GoTo GenerateSQLJoin_ERROR

  Dim objTableView As CTablePrivilege
  Dim objChildTable As CTablePrivilege
  Dim rsTemp As Recordset
  Dim objEvent As clsCalendarEvent
  
  Dim sChildJoinCode As String
  Dim strFilterIDs As String
  Dim sChildJoin As String
  
  Dim blnOK As Boolean
  
  Dim i As Integer
  Dim intLoop As Integer

  Dim bViewContains_StartColumn As Boolean
  Dim bViewContains_EndColumn As Boolean
  Dim bViewContains_DurationColumn As Boolean
  
  ' First, do the join for all the views etc...

  Set objEvent = mcolEvents.Item(pstrEventKey)

  For intLoop = 1 To UBound(mlngTableViews, 2)
    
    If mlngTableViews(1, intLoop) = 0 Then
      Set objTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, intLoop))
      bViewContains_StartColumn = True
      bViewContains_EndColumn = True
      bViewContains_DurationColumn = True
    Else
      Set objTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, intLoop))
      bViewContains_StartColumn = IsColumnInView(objTableView.ViewID, objEvent.StartDateID)
      bViewContains_EndColumn = IsColumnInView(objTableView.ViewID, objEvent.EndDateID)
      bViewContains_DurationColumn = IsColumnInView(objTableView.ViewID, objEvent.DurationID)
    End If
   
    If (objTableView.TableID = mlngCalendarReportsBaseTable) And (objTableView.ViewID > 0) Or (datGeneral.IsAParentOf(objTableView.TableID, mlngCalendarReportsBaseTable)) Then
      ' Get the table/view object from the id stored in the array
      
      If (datGeneral.IsAParentOf(objTableView.TableID, mlngCalendarReportsBaseTable)) Then
        mstrSQLJoin = mstrSQLJoin & _
        " LEFT OUTER JOIN " & objTableView.RealSource & _
        " ON " & mstrBaseTableRealSource & ".ID_" & objTableView.TableID & " = " & objTableView.RealSource & ".ID"
        
        mstrSQLBaseData = mstrSQLBaseData & _
        " LEFT OUTER JOIN " & objTableView.RealSource & _
        " ON " & mstrBaseTableRealSource & ".ID_" & objTableView.TableID & " = " & objTableView.RealSource & ".ID"

      Else
        mstrSQLJoin = mstrSQLJoin & _
        " LEFT OUTER JOIN " & objTableView.RealSource & _
        " ON " & mstrBaseTableRealSource & ".ID = " & objTableView.RealSource & ".ID"
        
        mstrSQLBaseData = mstrSQLBaseData & _
        " LEFT OUTER JOIN " & objTableView.RealSource & _
        " ON " & mstrBaseTableRealSource & ".ID = " & objTableView.RealSource & ".ID"
      End If
      
      If (objTableView.TableID = objEvent.TableID) Then
        'add clause to SQL, so that only dates within the specified range are retrieved.
        If (objEvent.StartDateID > 0) And (objEvent.EndDateID > 0) _
            And bViewContains_StartColumn And bViewContains_EndColumn Then
          'event is defined by start date and end date
          mstrSQLJoin = mstrSQLJoin & _
                " AND ((" & objTableView.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') AND " & objTableView.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine & _
                      " OR (" & objTableView.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') AND " & objTableView.RealSource & "." & objEvent.EndDateName & " <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine & _
                      " OR (((" & objTableView.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')) AND (" & objTableView.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))) AND " & objTableView.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine & _
                      " OR (" & objTableView.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') AND " & objTableView.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
          mstrSQLJoin = mstrSQLJoin & ")" & vbNewLine
          mstrSQLJoin = mstrSQLJoin & " AND (" & objTableView.RealSource & "." & objEvent.EndDateName & " >= " & objTableView.RealSource & "." & objEvent.StartDateName & ")" & vbNewLine
                   
        ElseIf (objEvent.StartDateID > 0) And (objEvent.DurationID > 0) _
                And bViewContains_StartColumn And bViewContains_DurationColumn Then
          'event is defined by start date and duration
          mstrSQLJoin = mstrSQLJoin & " OR (" & objTableView.RealSource & "." & objEvent.StartDateName & " IS NOT NULL AND " & objTableView.RealSource & "." & objEvent.DurationName & " > 0)" & vbNewLine
        
        ElseIf (objEvent.StartDateID) > 0 And (objEvent.EndDateID < 1) And (objEvent.DurationID < 1) _
                And bViewContains_StartColumn Then
          'event is defined by just the start date - one off event with a range of one
          mstrSQLJoin = mstrSQLJoin & _
                " AND ((" & objTableView.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') AND " & objTableView.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
          mstrSQLJoin = mstrSQLJoin & ")" & vbNewLine
          
        End If
      End If
      
    ElseIf (datGeneral.IsAChildOf(mlngTableViews(2, intLoop), mlngCalendarReportsBaseTable)) And (objEvent.TableID = objTableView.TableID) Then
      Set objChildTable = gcoTablePrivileges.FindTableID(mlngTableViews(2, intLoop))

      If objChildTable.AllowSelect Then
        sChildJoinCode = sChildJoinCode & " INNER JOIN " & objChildTable.RealSource & _
                         " ON " & mstrBaseTableRealSource & ".ID = " & _
                         objChildTable.RealSource & ".ID_" & mlngCalendarReportsBaseTable
  
        If (objEvent.FilterID > 0) Then
          
          'TM20030407 Fault 5257 - only get the filter string once for each event to avoid being prompted
          'more tahn once for the save event if the event is split into dynamic events.
          If mblnHasEventFilterIDs Then
            blnOK = True
          Else
            blnOK = datGeneral.FilteredIDs(objEvent.FilterID, strFilterIDs)
            mblnHasEventFilterIDs = blnOK
            mstrEventFilterIDs = strFilterIDs
          End If
          
          ' Generate any UDFs that are used in this filter
          If blnOK And gbEnableUDFFunctions Then
            datGeneral.FilterUDFs objEvent.FilterID, mastrUDFsRequired()
          End If

          If blnOK Then
            sChildJoinCode = sChildJoinCode & " AND " & _
              objChildTable.RealSource & ".ID IN (" & mstrEventFilterIDs & ")"
          Else
            ' Permission denied on something in the filter.
            mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(objEvent.FilterID) & "' filter."
            GenerateSQLJoin = False
            GoTo TidyUpAndExit
          End If
        End If
        
        'add clause to SQL, so that only dates within the specified range are retrieved.
        If (objEvent.StartDateID > 0 And objEvent.EndDateID > 0) Then
          'event is defined by start date and end date
          sChildJoinCode = sChildJoinCode & _
                " AND ((" & objChildTable.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') AND " & objChildTable.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine & _
                      " OR (" & objChildTable.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') AND " & objChildTable.RealSource & "." & objEvent.EndDateName & " <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine & _
                      " OR (((" & objChildTable.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')) AND (" & objChildTable.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))) AND " & objChildTable.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine & _
                      " OR (" & objChildTable.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') AND " & objChildTable.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
          sChildJoinCode = sChildJoinCode & ")" & vbNewLine
          sChildJoinCode = sChildJoinCode & " AND (" & objChildTable.RealSource & "." & objEvent.EndDateName & " >= " & objChildTable.RealSource & "." & objEvent.StartDateName & ") "
        
        ElseIf (objEvent.StartDateID > 0) And (objEvent.DurationID > 0) Then
          'event is defined by start date and duration
          sChildJoinCode = sChildJoinCode & " AND (" & objChildTable.RealSource & "." & objEvent.StartDateName & " IS NOT NULL AND " & objChildTable.RealSource & "." & objEvent.DurationName & " > 0)" & vbNewLine
            
        ElseIf objEvent.StartDateID > 0 And (objEvent.EndDateID < 1) And (objEvent.DurationID < 1) Then
          'event is defined by just the start date - one off event with a range of one
          sChildJoinCode = sChildJoinCode & _
                " AND ((" & objChildTable.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') AND " & objChildTable.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
          sChildJoinCode = sChildJoinCode & ")" & vbNewLine
          
        End If
      End If
    End If
    
  Next intLoop

  mstrSQLJoin = mstrSQLJoin & sChildJoinCode
'  mstrSQLBaseData = mstrSQLBaseData & mstrSQLJoin
  
  GenerateSQLJoin = True

TidyUpAndExit:
  strFilterIDs = vbNullString
  Set objTableView = Nothing
  Set objChildTable = Nothing
  Set rsTemp = Nothing
  Set objEvent = Nothing
  Exit Function
  
GenerateSQLJoin_ERROR:
  GenerateSQLJoin = False
  mstrErrorString = "Error in GenerateSQLJoin." & vbNewLine & Err.Description
  GoTo TidyUpAndExit
  
End Function

Private Function GenerateSQLWhere(pstrEventKey As String, _
                                    pstrDynamicKey As String, _
                                    pstrDynamicName As String) As Boolean

  ' Purpose : Generate the where clauses that cope with the joins
  '           NB Need to add the where clauses for filters/picklists etc

  On Error GoTo GenerateSQLWhere_ERROR
  
  Dim objExpr As clsExprExpression
  Dim rsTemp As New Recordset
  Dim objEvent As clsCalendarEvent
  
  Dim strPickListIDs As String
  Dim strFilterIDs As String
  
  Dim blnOK As Boolean
  
  Set objEvent = mcolEvents.Item(pstrEventKey)

'*******************************************************************************
  Dim pintLoop As Integer
  Dim pobjTableView As CTablePrivilege
  
  Set pobjTableView = gcoTablePrivileges.FindTableID(mlngCalendarReportsBaseTable)
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
        If (mlngTableViews(1, pintLoop) = 1) Then
            mstrSQLWhere = mstrSQLWhere & _
            IIf(Len(mstrSQLWhere) > 0, " OR ", " WHERE (") & _
            mstrBaseTableRealSource & ".ID IN (SELECT ID FROM " & pobjTableView.RealSource & ")"
        End If
        
      Next pintLoop
    
      If Len(mstrSQLWhere) > 0 Then mstrSQLWhere = mstrSQLWhere & ")"
    End If
        
  End If
'*******************************************************************************
  If mlngSingleRecordID > 0 Then
    mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
    mstrBaseTableRealSource & ".ID IN (" & mstrSQLIDs & ") "
    
  ' Now if we are using a picklist, add a where clause for that
  'Get List of IDs from Picklist
  ElseIf mlngCalendarReportsPickListID > 0 Then
    mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
    mstrBaseTableRealSource & ".ID IN (" & mstrSQLIDs & ")"
    
    
  ElseIf mlngCalendarReportsFilterID > 0 Then
      mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
        mstrBaseTableRealSource & ".ID IN (" & mstrSQLIDs & ")"
      
  End If
  
  'TM03032004 Fault 8183
  mstrSQLBaseData = mstrSQLBaseData & mstrSQLWhere

  'add clause to SQL, so that only dates within the specified range are retieved.
  If objEvent.TableID = mlngCalendarReportsBaseTable Then
    mstrSQLBaseDateClause = vbNullString
    'mstrSQLBaseDateClause = mstrSQLBaseDateClause & " AND ("
    
    'add clause to SQL, so that only dates within the specified range are retrieved.
    If objEvent.StartDateID > 0 And objEvent.EndDateID > 0 Then
      'event is defined by start date and end date
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "((" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') AND " & mstrSQLBaseEndDateColumn & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & " OR (" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') AND " & mstrSQLBaseEndDateColumn & " <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & " OR (((" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')) AND (" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))) AND " & mstrSQLBaseEndDateColumn & " >= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & " OR (" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') AND " & mstrSQLBaseEndDateColumn & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')))" & vbNewLine
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & " AND (" & mstrSQLBaseEndDateColumn & ">=" & mstrSQLBaseStartDateColumn & ")"

    ElseIf (objEvent.StartDateID > 0) And (objEvent.DurationID > 0) Then
      'TM 25/04/2005 - Fault 10038 - Check if the Start Date + Duration puts event in the report range.
      'event is defined by start date and duration
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "    (" & mstrSQLBaseDurationColumn & " > 0)" & vbNewLine
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "    AND (" & vbNewLine & vbNewLine
      
      ' 1 Event Start Date before Report Start Date, Duration carrys event into, but not beyond the Report Range.
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        (" & mstrSQLBaseStartDateColumn & " < convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')" & vbNewLine
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "      AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')))" & vbNewLine & vbNewLine
      
      ' 2 Event Start Date within Report Range, Duration carrys event beyond Report End Date.
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "     OR ((" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        AND (" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "      AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") > convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')))" & vbNewLine & vbNewLine
      
      ' 3 Event Start Date within Report Range and Duration keeps event within Report Range.
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "     OR ((" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        AND (" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "      AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')))" & vbNewLine & vbNewLine
      
      ' 4 Event Start Date before Report Start Date and Duration carrys event beyond Report End Date.
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "     OR ((" & mstrSQLBaseStartDateColumn & " < convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'))" & vbNewLine
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "      AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") > convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')))" & vbNewLine & vbNewLine
              
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        )" & vbNewLine
      
    ElseIf objEvent.StartDateID > 0 And (objEvent.EndDateID < 1) And (objEvent.DurationID < 1) Then
      'event is defined by just the start date - one off event with a range of one
      mstrSQLBaseDateClause = mstrSQLBaseDateClause & "(" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & Replace(Format(mdtStartDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "') AND " & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & Replace(Format(mdtEndDate, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "')) " & vbNewLine

    End If
    
    mstrSQLBaseDateClause = mstrSQLBaseDateClause & " AND (" & mstrSQLBaseStartDateColumn & " IS NOT NULL)" & vbNewLine & vbNewLine
    
    If objEvent.FilterID > 0 Then
      blnOK = datGeneral.FilteredIDs(objEvent.FilterID, strFilterIDs)
      
      ' Generate any UDFs that are used in this filter
      If blnOK And gbEnableUDFFunctions Then
        datGeneral.FilterUDFs objEvent.FilterID, mastrUDFsRequired()
      End If

      If blnOK Then
        mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & _
        mstrBaseTableRealSource & ".ID IN (" & strFilterIDs & ")"
      Else
        ' Permission denied on something in the filter.
        mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(objEvent.FilterID) & "' filter."
        GenerateSQLWhere = False
        GoTo TidyUpAndExit
      End If
    End If

    If Len(mstrSQLWhere) > 0 Then
      mstrSQLWhere = mstrSQLWhere & " AND (" & mstrSQLBaseDateClause & ") "
    Else
      mstrSQLWhere = mstrSQLWhere & " WHERE " & mstrSQLBaseDateClause
    End If
  End If

  If (Len(mstrSQLDynamicLegendWhere) > 0) Then
    If (Len(mstrSQLWhere) > 0) Then
      mstrSQLWhere = mstrSQLWhere & " AND (" & mstrSQLDynamicLegendWhere & ") "
    Else
      mstrSQLWhere = mstrSQLWhere & " WHERE " & mstrSQLDynamicLegendWhere
    End If
  End If
  
  mstrSQLDynamicLegendWhere = vbNullString
  mstrSQLBaseStartDateColumn = vbNullString
  mstrSQLBaseEndDateColumn = vbNullString
  mstrSQLBaseDurationColumn = vbNullString
  
  GenerateSQLWhere = True
  
TidyUpAndExit:
  Set rsTemp = Nothing
  Exit Function
  
GenerateSQLWhere_ERROR:
  GenerateSQLWhere = False
  mstrErrorString = "Error in GenerateSQLWhere." & vbNewLine & Err.Description
  GoTo TidyUpAndExit
  
End Function
Private Function GenerateSQLOrderBy() As Boolean

  ' Purpose : Returns order by string from the sort order array
  
  On Error GoTo GenerateSQLOrderBy_ERROR
  
  mstrSQLOrderBy = " ORDER BY " & mstrSQLOrderBy
  
  mstrSQLBaseData = mstrSQLBaseData & mstrSQLOrderBy
  
  GenerateSQLOrderBy = True
  Exit Function
  
GenerateSQLOrderBy_ERROR:
  
  GenerateSQLOrderBy = False
  mstrErrorString = "Error in GenerateSQLOrderBy." & vbNewLine & Err.Description
  
End Function
Private Function InitialisePreview() As Boolean

  ' Purpose : To prepare the output form
  ' Input   : None
  ' Output  : None

  On Error GoTo InitForm_ERROR
  
  Set mfrmOutput = New frmCalendarReportPreview
  
  ' Initialise the progress bar
  With gobjProgress
    If Not gblnBatchMode Then
      .Bar1MaxValue = (mrsCalendarBaseInfo.RecordCount + mrsCalendarReportsOutput.RecordCount)
    Else
      .ResetBar2
      .Bar2MaxValue = (mrsCalendarBaseInfo.RecordCount + mrsCalendarReportsOutput.RecordCount + 1)
      .Bar2Caption = "Calendar Report : " & mstrCalendarReportsName
    End If
  End With
  
  With mfrmOutput
    .BaseIDColumn = "?ID_" & mstrCalendarReportsBaseTableName
    .EventIDColumn = "?ID_EventID"
    
    .PersonnelBase = (mlngCalendarReportsBaseTable = glngPersonnelTableID)
    
    .BaseTableRealSource = mstrBaseTableRealSource
    .BaseTableID = mlngCalendarReportsBaseTable
    .BaseTableName = mstrCalendarReportsBaseTableName
    .BaseDesc1IsDate = mblnDesc1IsDate
    .BaseDesc2IsDate = mblnDesc2IsDate
    .BaseDescExprIsDate = mblnDescExprIsDate
    
    .SQLIDs = mstrSQLIDs
    
    .StaticRegionColumn = mstrRegion
    .StaticRegionColumnID = mlngRegion
    .StaticRegionRealSource = mstrRegionColumnRealSource
        
    .ReportStartDate = mdtStartDate
    .ReportEndDate = mdtEndDate
  
    If mblnCustomReportsPrintFilterHeader And (mlngSingleRecordID < 1) Then
      If (mlngCalendarReportsFilterID > 0) Then
        .CalendarReportName = mstrCalendarReportsName & " (Base Table filter : " & datGeneral.GetFilterName(mlngCalendarReportsFilterID) & ")"
      ElseIf (mlngCalendarReportsPickListID > 0) Then
        .CalendarReportName = mstrCalendarReportsName & " (Base Table picklist : " & datGeneral.GetPicklistName(mlngCalendarReportsPickListID) & ")"
      End If
    Else
      .CalendarReportName = mstrCalendarReportsName
    End If

    .Tag = mstrCalendarReportsName
    
    .EventsRecordset = mrsCalendarReportsOutput
    .BaseRecordset = mrsCalendarBaseInfo
    
    .DescriptionSeparator = mstrDescriptionSeparator
    .Description1ID = mlngDescription1ID
    .Description2ID = mlngDescription2ID
    .DescriptionExprID = mlngDescriptionExpr
    .GroupByDescription = mblnGroupByDescription
    
    .ReportStartDate = mdtStartDate
    .ReportEndDate = mdtEndDate
    
    .ShowBankHolidays = mblnShowBankHolidays
    .ShowCaptions = mblnShowCaptions
    .ShowWeekends = mblnShowWeekends
    .IncludeWorkingDaysOnly = mblnIncludeWorkingDaysOnly
    .IncludeBankHolidays = mblnIncludeBankHolidays
    .StartOnCurrentMonth = mbStartOnCurrentMonth
    
    'pass the default output options to the calendar preview screen
    .OutputFormat = mlngOutputFormat
    .OutputScreen = mblnOutputScreen
    .OutputPrinter = mblnOutputPrinter
    .OutputPrinterName = mstrOutputPrinterName
    .OutputSave = mblnOutputSave
    .OutputSaveExisting = mlngOutputSaveExisting
    .OutputSaveFormat = mlngOutputSaveFormat
    .OutputEmail = mblnOutputEmail
    .OutputEmailAddr = mlngOutputEmailAddr
    .OutputEmailSubject = mstrOutputEmailSubject
    .OutputEmailAttachAs = mstrOutputEmailAttachAs
    .OutputEmailFileFormat = mlngOutputEmailFileFormat
    .OutputFilename = mstrOutputFilename
    
    .HasMultipleEvents = mblnHasMultipleEvents
  End With
  
  InitialisePreview = True

TidyUpAndExit:
  Exit Function
  
InitForm_ERROR:
  InitialisePreview = False
  mstrErrorString = "Error whilst initialising Calendar Report Preview screen." & vbNewLine & Err.Description
  GoTo TidyUpAndExit
  
End Function
Private Function ClearUp() As Boolean

  ' Purpose : To clear all variables/recordsets/references and drops temptable
  ' Input   : None
  ' Output  : True/False success

  ' Definition variables
  
  On Error GoTo ClearUp_ERROR
  
  mstrCalendarReportsName = vbNullString
  mlngCalendarReportsBaseTable = 0
  mstrCalendarReportsBaseTableName = vbNullString
  mlngCalendarReportsAllRecords = 0
  mlngCalendarReportsPickListID = 0
  mlngCalendarReportsFilterID = 0
  
  mlngDescription1 = 0
  mstrDescription1 = vbNullString
  mlngDescription2 = 0
  mstrDescription2 = vbNullString
  mlngDescriptionExpr = 0
  mstrDescriptionExpr = vbNullString
  mblnGroupByDescription = False
  
  mlngRegion = 0
  mstrRegion = vbNullString
  
  mlngStartDateExpr = 0
  mstrStartDate = vbNullString
  'mdtStartDate = Null
  mlngEndDateExpr = 0
  mstrEndDate = vbNullString
  'mdtEndDate = Null
  
  mblnShowBankHolidays = False
  mblnShowCaptions = False
  mblnShowWeekends = False
  mblnIncludeWorkingDaysOnly = False
  mblnIncludeBankHolidays = False
  
  'New Default Output Variables
  mlngOutputFormat = 0
  mblnOutputScreen = True
  mblnOutputPrinter = False
  mstrOutputPrinterName = vbNullString
  mblnOutputSave = False
  mlngOutputSaveExisting = 0
  mlngOutputSaveFormat = 0
  mblnOutputEmail = False
  mlngOutputEmailAddr = 0
  mstrOutputEmailSubject = vbNullString
  mstrOutputEmailAttachAs = vbNullString
  mlngOutputEmailFileFormat = 0
  mstrOutputFilename = vbNullString
  
  mblnDefinitionOwner = False
   
  ' Recordsets
  Set mrsCalendarReportsOutput = Nothing
  Set mrsCalendarBaseInfo = Nothing

  'kill the temp table
  datGeneral.DropUniqueSQLObject mstrTempTableName, 3
  mstrTempTableName = vbNullString

  ' SQL strings
  mstrSQLEvent = vbNullString
  mstrSQLSelect = vbNullString
  mstrSQLFrom = vbNullString
  mstrSQLWhere = vbNullString
  mstrSQLJoin = vbNullString
  mstrSQLOrderBy = vbNullString
  mstrSQL = vbNullString
  mstrSQLBaseData = vbNullString
  mstrSQLBaseDateClause = vbNullString
  mstrSQLOrderList = vbNullString
  mstrSQLIDs = vbNullString
  mstrSQLDynamicLegendWhere = vbNullString
  mstrSQLCreateTable.TheString = vbNullString
  
  mstrDescCalcCode = vbNullString
  mblnDescCalcCode = False
  mintDynamicEventCount = 0
  
  ' Class references
  Set mclsData = Nothing
  Set mclsGeneral = Nothing
  
  ' Arrays
  ReDim mavEventDetails(21, 0)
  ReDim mavSortOrder(2, 0)

  ' Column Privilege arrays / collections / variables
  mstrBaseTableRealSource = vbNullString
  mstrRealSource = vbNullString
  Set mobjTableView = Nothing
  Set mobjColumnPrivileges = Nothing
  ReDim mlngTableViews(2, 0)
  ReDim mstrViews(0)
  
  Set mcolEvents = Nothing
  Set mfrmOutput = Nothing
  
  gobjProgress.ResetBar2
  
  ClearUp = True
  Exit Function
  
ClearUp_ERROR:
  
  mstrErrorString = "Error whilst clearing data." & vbNewLine & "(" & Err.Description & ")"
  ClearUp = False
  
End Function
Private Function IsRecordSelectionValid() As Boolean

  Dim i As Integer
  Dim lngFilterID As Long
  Dim objEvent As clsCalendarEvent
  Dim iResult As RecordSelectionValidityCodes
  
  ' Base Table First
  If mlngCalendarReportsFilterID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngCalendarReportsFilterID)
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
  ElseIf mlngCalendarReportsPickListID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, mlngCalendarReportsPickListID)
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
    
  'Description Calculation
  If Len(mstrErrorString) = 0 Then
    If mlngDescriptionExpr > 0 Then
      iResult = ValidateCalculation(mlngDescriptionExpr)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "The base description calculation used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "The base description calculation used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "The base description calculation used in this definition has been made hidden by another user."
          End If
      End Select
    End If
  End If
  
  'Events Filters
  For Each objEvent In mcolEvents.Collection
    lngFilterID = objEvent.FilterID
    If lngFilterID > 0 Then
      iResult = ValidateRecordSelection(REC_SEL_FILTER, lngFilterID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "An event table filter used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "An event table filter used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "An event table filter used in this definition has been made hidden by another user."
          End If
      End Select
    End If
  Next objEvent
  Set objEvent = Nothing

  'Report Start Date Calculation
  If Len(mstrErrorString) = 0 Then
    If mlngStartDateExpr > 0 Then
      iResult = ValidateCalculation(mlngStartDateExpr)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "The report start date calculation used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "The report start date calculation used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "The report start date calculation used in this definition has been made hidden by another user."
          End If
      End Select
    End If
  End If

  'Report End Date Calculation
  If Len(mstrErrorString) = 0 Then
    If mlngEndDateExpr > 0 Then
      iResult = ValidateCalculation(mlngEndDateExpr)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          mstrErrorString = "The report end date calculation used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          mstrErrorString = "The report end date calculation used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            mstrErrorString = "The report end date calculation used in this definition has been made hidden by another user."
          End If
      End Select
    End If
  End If

  IsRecordSelectionValid = (Len(mstrErrorString) = 0)

End Function
