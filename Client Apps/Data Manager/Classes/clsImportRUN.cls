VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportRUN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Data access classes
Private mclsData As clsDataAccess
Private mclsGeneral As clsGeneral
Private mclsDeadlock As clsDeadlock

' To hold Properties
'Private gblnBatchMode As Boolean
Private mstrErrorString As String
Private mlngImportID As Long
Private mblnUserCancelled As Boolean
Private mblnExcelFileHasBeenConverted As Boolean
Private mblnCanInsert As Boolean
Private mblnCanUpdate As Boolean

' Variables to store definition
Private mstrImportName As String
Private mstrImportDescription As String
Private mstrImportBaseTableName As String
Private mlngImportBaseTable As Long
Private mstrBaseTableRealSource As String
Private mstrBaseTableRealSources() As String
Private mstrBaseTableViewJoins As String
Private mlngImportParent1Table As Long
Private mstrImportParent1TableName As String
Private mstrParent1TableRealSource As String
Private mstrParent1TableRealSources() As String
Private mstrParent1ViewJoins As String
Private mlngImportParent2Table As Long
Private mstrImportParent2TableName As String
Private mstrParent2TableRealSource As String
Private mstrParent2TableRealSources() As String
Private mstrParent2ViewJoins As String
Private mintImportFileType As ImportType
Private mstrImportFileName As String
Private mstrImportDelimiter As String
Private mstrImportOtherDelimiter As String
Private mstrDateFormat As String
Private mstrImportEncapsulator As String
Private mblnImportMultipleRecordAction As Boolean

Private mintImportHeaderLines As Integer
Private mintImportFooterLines As Integer

Private mstrLinkedServer As String
Private mstrLinkedCatalog As String
Private mstrLinkedTable As String
Private maryRecordSuccessIDs() As String
Private mbUseUpdateBlob As Boolean

Private mblnImportCreateNewOnly As Boolean
Private mlngFilterID As Long
Private mstrDateSeparator As String

Private mstrParent1Where As String
Private mstrParent2Where As String

Private mlngImportType As Long
    
' Store progress/status of imported records
Private mlngTotalRecords As Long
Private mlngSuccessRecords As Long
Private mlngFailRecords As Long

' Integer to hold the file reference
Private mintFileNo As Integer

' Strings to hold the SQL components
Private mstrSQL As String
Private mstrSQLWhere As String
Private mstrIDsToUpdate As String

' Hold working string
Private mstrBufferString As String

' Array to hold the columns used in the import
Private mvarColDetails() As Variant
Private mavColumns() As Variant

Private mstrLookupArray() As String

' Excel Object reference
Private mobjExcelApp As Excel.Application

' String to try and speed things up - contains the column list for the SQL statement
' (to save having to go through the mvarcoldetails array for each row)
Private mstrColumnList As String

Private mbLoggingImportSuccess As Boolean
Private mlngRecordDescExprID As Long
Private mbBypassTrigger As Boolean

Private mstrEOLChars As String


Private Function CheckFilter(pobjFilter As clsExprExpression) As Boolean
  ' Check if the current file record satisfies the defined filter criteria.
  Dim fOK As Boolean
  Dim iCounter As Integer
  Dim iLoop As Integer
  Dim sTemp As String
  Dim vEvaluatedValue As Variant
  Dim sTempDate As String

  On Local Error GoTo LocalErr

  fOK = True
  iLoop = 0
  
  If mlngFilterID > 0 Then
    ' Perform the filter check.
    
    ' Get the values from the file string.
    For iCounter = 1 To UBound(mvarColDetails, 2)
      If mvarColDetails(3, iCounter) > 0 Then
        iLoop = iLoop + 1
        
        ' NB. The 'GetSpecificKeyData' function returns the import value as it should be
        ' for a SQL insert/update string. We need to convert string values into a real value.
        mavColumns(2, iLoop) = GetSpecificKeyData(iCounter)
        If mstrErrorString <> vbNullString Then
          mlngFailRecords = mlngFailRecords + 1
          CheckFilter = False
          Exit Function
        End If

        If mvarColDetails(7, iCounter) Then
          ' Date column.
          If mavColumns(2, iLoop) = "NULL" Then
            ' No date value from the import file.
            mavColumns(2, iLoop) = vbNullString
          Else
            ' Date value from the import file. Convert from a string in SQL format ('mmddyyyy') to vbDate type.
            sTemp = Mid(mavColumns(2, iLoop), 2)
            sTemp = Left(sTemp, Len(sTemp) - 1)
            
            ' JPD20011012 Fault 2949 Check the converted string is a valid date before using CDate.
            sTempDate = ConvertSQLDateToLocale(sTemp)
            If IsDate(sTempDate) Then
              mavColumns(2, iLoop) = CDate(sTempDate)
            Else
              mavColumns(2, iLoop) = Null
            End If
          End If
        ElseIf mvarColDetails(8, iCounter) Then
          ' Numeric column.
          mavColumns(2, iLoop) = CDbl(datGeneral.ConvertNumberForDisplay(mavColumns(2, iLoop)))
        
        ElseIf mvarColDetails(9, iCounter) Then
          ' Logic column.
          mavColumns(2, iLoop) = (UCase(mavColumns(2, iLoop)) = "1")
        Else
          ' String column.
            sTemp = Mid(mavColumns(2, iLoop), 2)
            sTemp = Left(sTemp, Len(sTemp) - 1)
            ' JPD20021024 Fault 4652
            sTemp = RTrim(UCase(sTemp))
            mavColumns(2, iLoop) = Replace(sTemp, "''", "'")
        End If
      End If
    Next iCounter
  
    ' We've got an array of the import values for each column.
    ' Now get the filter object to evaluate the filter expression with these
    ' values to see if the record is to be imported.
    fOK = pobjFilter.EvaluateExpression(mavColumns, vEvaluatedValue)
    If fOK Then
      ' The filter was evaluated okay, so check the evaluated value.
      ' JPD20011012 Fault 2949 Check the result is valid.
      If IsNull(vEvaluatedValue) Then
        fOK = False
      Else
        fOK = vEvaluatedValue
      End If
    End If
  End If

  CheckFilter = fOK
  
Exit Function

LocalErr:
  RecordError "Error evaluating filter for this record."
  CheckFilter = False

End Function

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property

Public Property Get ImportID() As Long
  ImportID = mlngImportID
End Property

Public Property Let ImportID(ByVal lNewID As Long)
  mlngImportID = lNewID
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
  ' Input   : None
  ' Output  : None
  
  Set mclsData = New DataMgr.clsDataAccess
  Set mclsGeneral = New DataMgr.clsGeneral
  Set mclsDeadlock = New DataMgr.clsDeadlock
  'MH20020402 Moved to ConvertExcelFile
  'Set mobjExcelApp = New Excel.Application
  ReDim mvarColDetails(13, 0)

End Sub

Private Sub Class_Terminate()
  
  ' Purpose : Clears references to other classes.
  ' Input   : None
  ' Output  : None
  
  Set mclsData = Nothing
  Set mclsGeneral = Nothing
  Set mclsDeadlock = Nothing
  Set mobjExcelApp = Nothing
  
End Sub

Public Function RunImport() As Boolean

  ' Purpose : This function is called from frmDefsel and Batch Jobs.
  '           It is the main function which runs the import.
  ' Input   : Is the util being run in batch mode or not
  ' Output  : True/False.
  
  On Error GoTo RunImport_ERROR
  
  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim iShowMeAgain As Integer
  Dim strSQL As String
  
  Dim msb As frmMessageBox

  Dim fOK As Boolean
  fOK = True
  Screen.MousePointer = vbHourglass
  
  If fOK Then fOK = GetImportDefinition
  
  mbLoggingImportSuccess = CBool(GetUserSetting("LogEvents", "Import_Success", False))
 
  If Not gblnBatchMode Then
    'TM20010726 Fault 1607 (Suggestion)
    'NHRD04062003 Fault 5695 I see your SUG and raise it to a BUG.
    'Added the "It is Recommended...." line.  Ok it's a SUG really
    strMBText = "WARNING: You are about to Import " & _
                "into the '" & mstrImportBaseTableName & "' table." & vbNewLine & vbNewLine & _
                "It is recommended that a backup of the system be made before any Imports are run!" & vbNewLine & vbNewLine & _
                "Do you wish to continue?"
    
    intMBButtons = vbExclamation + vbYesNo + vbDefaultButton2
    
    iShowMeAgain = GetUserSetting("warningmsg", "warning import", 1)
    If iShowMeAgain = 1 Then
      Set msb = New frmMessageBox
      iShowMeAgain = IIf(iShowMeAgain = 1, 0, 1)
      mblnUserCancelled = (msb.MessageBox(strMBText, intMBButtons, "Import", iShowMeAgain, "&Don't show me this warning again.") <> vbYes)
      SaveUserSetting "warningmsg", "warning import", IIf(iShowMeAgain = 1, 0, 1)
      Set msb = Nothing
    End If
  End If
  
  fOK = Not mblnUserCancelled
  
  'TM20010814 Fault 2685
  'Initialise the progress bar after prompted the user.
  If fOK Then fOK = DoProgressBar
 
  gobjEventLog.AddHeader eltImport, mstrImportName
  
  If fOK Then fOK = UpdateProgressBarWithName
  If fOK Then fOK = GetImportDetails
  If fOK Then fOK = CheckIfImportIsValid
  If fOK Then fOK = GetRealSources
  
  If mbBypassTrigger Then
    strSQL = "INSERT ASRSysSystemSettings ([Section], [SettingKey], [SettingValue]) VALUES ('database', 'ByPassTrigger_" & mlngImportBaseTable & "_SPID' ,@@SPID)"
    gADOCon.Execute strSQL, , adExecuteNoRecords
  End If
  
  ' Because the whole import is tied to read row by row we need to fudge a file from the datasource
  If mintImportFileType = ImportType.SQLTable Then
    fOK = ReadSQLTableInToFile
  End If
  
  If fOK Then fOK = ProcessFile
  
  If mbBypassTrigger Then
    strSQL = "DELETE ASRSysSystemSettings WHERE [Section] = 'database' AND [SettingKey] = 'ByPassTrigger_" & mlngImportBaseTable & "_SPID'"
    gADOCon.Execute strSQL, , adExecuteNoRecords
  End If
 
  If fOK And mbUseUpdateBlob Then
    UpdateImportedRecords
  End If
 
  If fOK And mintImportFileType = ImportType.SQLTable Then
    Kill mstrImportFileName
  End If
  
  If Not gblnBatchMode Then gobjProgress.CloseProgress
  
  Call UtilUpdateLastRun(utlImport, mlngImportID)
  Call PopulateEventLog(fOK)
  Call ClearUp

  'JDM - 23/09/03 - Only send emails at end of the job
  objEmail.SendImmediateEmails

  RunImport = fOK
  Exit Function
  
RunImport_ERROR:

  fOK = False
  mstrErrorString = "Error whilst running Import." & vbCrLf & "(" & Err.Description & ")"
  Resume Next
  
End Function
  
Private Function Records(lngRec As Long) As String
  Records = CStr(lngRec) & IIf(lngRec <> 1, " records", " record")
End Function
  
Private Sub PopulateEventLog(fOK As Boolean)

  'MH20000704 Fault 527
  'Moved log header creation as detail records need header id !
  'gobjEventLog.AddHeader eltImport, mstrImportName
  If fOK Then
    If mlngFailRecords = 0 Then
      gobjEventLog.ChangeHeaderStatus elsSuccessful, mlngSuccessRecords, mlngFailRecords
      mstrErrorString = "Completed successfully."
    
      If mlngSuccessRecords = 0 And mlngFailRecords = 0 Then
        mstrErrorString = mstrErrorString & vbCrLf & _
                            "No records meet selection criteria."
        gobjEventLog.AddDetailEntry mstrErrorString
      End If
    
    Else
      gobjEventLog.ChangeHeaderStatus elsFailed, mlngSuccessRecords, mlngFailRecords
      mstrErrorString = "Failed."
      fOK = False   'MH20021209 Faults 4743 & 4745
    End If
  
  ElseIf mblnUserCancelled Then
    gobjEventLog.ChangeHeaderStatus elsCancelled, mlngSuccessRecords, mlngFailRecords
    mstrErrorString = "Cancelled by user."
  Else
    If mlngFailRecords = 0 Then
      gobjEventLog.AddDetailEntry mstrErrorString
      gobjEventLog.ChangeHeaderStatus elsFailed, mlngSuccessRecords, mlngFailRecords
      mstrErrorString = "Failed." & vbCrLf & mstrErrorString
    Else
      gobjEventLog.ChangeHeaderStatus elsFailed, mlngSuccessRecords, mlngFailRecords
      mstrErrorString = "Failed. One or more rows could not be imported."
    End If
  End If
    
  'Add record count to output
  mstrErrorString = mstrErrorString & vbCrLf & vbCrLf & _
    Records(mlngSuccessRecords) & " added/updated"
    
  If mlngFailRecords > 0 Then
    mstrErrorString = mstrErrorString & vbCrLf & _
      Records(mlngFailRecords) & " failed"
  End If
  
  If Not gblnBatchMode Then
    COAMsgBox "Import : " & mstrImportName & " " & mstrErrorString, IIf(fOK, vbInformation, vbExclamation) + vbOKOnly, "Import"
  Else
    mstrErrorString = "Import : " & mstrImportName & " " & mstrErrorString
  End If

End Sub

Private Function DoProgressBar() As Boolean
  
  ' Purpose : Initialises the progress bar, takes into consideration whether
  '           or not the import is being run  in batch mode.
  ' Input   : None.
  ' Output  : True/False.
  
  On Error GoTo ErrTrap
  
  With gobjProgress
    '.AviFile = App.Path & "\videos\import.avi"
    .AVI = dbText
    .MainCaption = "Import"
    If Not gblnBatchMode Then
      .NumberOfBars = 1
      .Caption = "Import"
      .Time = False
      .Cancel = True
      .Bar1RecordsCaption = "Records Processed :"
      .OpenProgress
    Else
      .ResetBar2
      .Bar2RecordsCaption = "Records Processed :"
    End If
  End With

  DoProgressBar = True
  Exit Function
  
ErrTrap:
  
  mstrErrorString = "Error initialising progress bar." & vbCrLf & vbCrLf & "(" & Err.Description & ")"
  DoProgressBar = False
  
End Function

Private Function UpdateProgressBarWithName() As Boolean

  ' Purpose : Updates the bar caption of the progress bar with the
  '           definition name. (Not available until Definition has been read)
  ' Input   : None.
  ' Output  : True/False.

  On Error GoTo ErrTrap
  
  If gblnBatchMode Then
    gobjProgress.Bar2Caption = "Import : " & mstrImportName
  Else
    gobjProgress.Bar1Caption = "Import : " & mstrImportName
    'gobjEventLog.AddHeader eltImport, mstrImportName
  End If

  UpdateProgressBarWithName = True
  Exit Function

ErrTrap:

  mstrErrorString = "Error updating progress bar with definition name." & vbCrLf & vbCrLf & "(" & Err.Description & ")"
  UpdateProgressBarWithName = False

End Function

Private Function GetImportDefinition() As Boolean
  
  ' Purpose : This function retrieves the basic definition details
  '           and stores them in module level variables
  ' Input   : None
  ' Output  : True/False.
  
  On Error GoTo GetImportDefinition_ERROR

  Dim prstTemp_Definition As Recordset
  Dim pstrSQL As String
  
  ' If user cancels the import, abort
  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    GetImportDefinition = False
    Exit Function
  End If
  
  pstrSQL = "SELECT * FROM AsrSysImportName " & _
           "WHERE ID = " & mlngImportID & " "
  
  Set prstTemp_Definition = mclsData.OpenRecordset(pstrSQL, adOpenForwardOnly, adLockReadOnly)
  
  With prstTemp_Definition
  
    If .BOF And .EOF Then
      GetImportDefinition = False
      mstrErrorString = "Could not find specified Import definition !" & vbCrLf & vbCrLf & "Another user may have deleted this definition."
      Exit Function
    End If
    
    mstrImportName = !Name
    mstrImportDescription = !Description
    mlngImportBaseTable = !BaseTable
    mstrImportBaseTableName = datGeneral.GetTableName(mlngImportBaseTable)
    mlngImportType = !ImportType
    mintImportFileType = !filetype
    mstrImportFileName = !FileName
    
    If mintImportFileType = ImportType.SQLTable Then
      mstrImportDelimiter = ","
      mstrImportOtherDelimiter = ""
      mstrDateFormat = "dmy"
      mstrImportEncapsulator = Chr(34)
      mintImportHeaderLines = 0
      mintImportFooterLines = 0
    Else
      mstrImportDelimiter = !delimiter
      mstrImportOtherDelimiter = !otherdelimiter
      mstrDateFormat = IIf(IsNull(!DateFormat), "", !DateFormat)
      mstrImportEncapsulator = IIf(IsNull(!encapsulator), "", !encapsulator)
      mintImportHeaderLines = IIf(IsNull(!HeaderLines), 0, !HeaderLines)
      mintImportFooterLines = IIf(IsNull(!FooterLines), 0, !FooterLines)
    End If
    
    mblnImportMultipleRecordAction = !MultipleRecordAction

    mlngRecordDescExprID = datGeneral.GetRecDescExprID(mlngImportBaseTable)
    mbBypassTrigger = IIf(IsNull(!BypassTrigger), False, !BypassTrigger)

    mblnImportCreateNewOnly = (mlngImportType = 0)

    mlngFilterID = IIf(IsNull(!FilterID), 0, !FilterID)
    
    mstrDateSeparator = IIf(IsNull(!dateseparator) Or (!dateseparator = "<None>"), "", !dateseparator)
    
    mstrLinkedServer = IIf(IsNull(!LinkedServer), "", !LinkedServer)
    mstrLinkedCatalog = IIf(IsNull(!LinkedCatalog), "", !LinkedCatalog)
    mstrLinkedTable = IIf(IsNull(!LinkedTable), "", !LinkedTable)
    mbUseUpdateBlob = IIf(IsNull(!UseUpdateBlob), 0, !UseUpdateBlob)
    
  End With
  
  If mintImportFileType = ImportType.ExcelWorksheet Then
    gobjProgress.AVI = dbText
  End If
  
  
  GetImportDefinition = True
  
TidyAndExit:
  
  Set prstTemp_Definition = Nothing

Exit Function

GetImportDefinition_ERROR:

  GetImportDefinition = False
  mstrErrorString = "Error whilst reading the Import definition !" & vbCrLf & "(" & Err.Description & ")"
  Resume TidyAndExit

End Function


Private Function GetImportDetails() As Boolean

  ' Purpose : This function loads import details into an array.
  ' Input   : None.
  ' Output  : True/False.

  On Error GoTo GetImportDetails_ERROR

  ReDim mvarColDetails(13, 0)

  Dim pstrTempSQL As String
  Dim prstDetails As Recordset
  Dim pintTemp As Integer
  
  ' If user cancels the import, abort
  If gobjProgress.Cancelled Then
    mblnUserCancelled = True
    GetImportDetails = False
    Exit Function
  End If
      
  ' Get the column information from the Details table, in order
  pstrTempSQL = "SELECT * FROM AsrSysImportDetails WHERE " & _
                "ImportID = " & mlngImportID & " " & _
                "ORDER BY [ID]"
  Set prstDetails = mclsData.OpenRecordset(pstrTempSQL, adOpenKeyset, adLockReadOnly)

  With prstDetails
    
    If .BOF And .EOF Then
      GetImportDetails = False
      mstrErrorString = "No columns found in the specified Import definition." & vbCrLf & "Please remove this definition and create a new one."
      Exit Function
    End If
    
    Do Until .EOF
      pintTemp = UBound(mvarColDetails, 2) + 1
      DoEvents
      ReDim Preserve mvarColDetails(13, pintTemp)
      mvarColDetails(0, pintTemp) = !Type
      mvarColDetails(1, pintTemp) = Trim(!TableID)
      mvarColDetails(2, pintTemp) = IIf(!TableID = 0, "", datGeneral.GetTableName(CLng(!TableID)))
      mvarColDetails(3, pintTemp) = !ColExprID
      
      If !ColExprID = 0 Then
        mvarColDetails(4, pintTemp) = ""
      Else
        mvarColDetails(4, pintTemp) = datGeneral.GetColumnName(CLng(!ColExprID))
      End If
      
      mvarColDetails(5, pintTemp) = !KeyField
      mvarColDetails(6, pintTemp) = !Size
      
      If !Type = "C" Then
        mvarColDetails(7, pintTemp) = datGeneral.DateColumn(!Type, Trim(!TableID), !ColExprID)
        mvarColDetails(8, pintTemp) = datGeneral.NumericColumn(!Type, Trim(!TableID), !ColExprID)
        mvarColDetails(9, pintTemp) = datGeneral.BitColumn(!Type, Trim(!TableID), !ColExprID)
      Else
        mvarColDetails(7, pintTemp) = False
        mvarColDetails(8, pintTemp) = False
        mvarColDetails(9, pintTemp) = False
      End If
      
      ' Trimming format
      mvarColDetails(10, pintTemp) = datGeneral.GetColumnTrimmingType(!ColExprID)

      mvarColDetails(13, pintTemp) = datGeneral.IsColumnReadOnly(!ColExprID)

      'MH20030520 TEMPORARY!
      If !LookupEntries Then
        GetLookupInsert pintTemp, !ColExprID
      End If
      
      .MoveNext
    Loop
  .MoveFirst
  End With

  If mbUseUpdateBlob Then
    AddLMSImportBlob
  End If

  ' Now populate the Table used in the import variables
  Set prstDetails = mclsData.OpenRecordset("SELECT DISTINCT TableID FROM ASRSysImportDetails WHERE TableID <> 0 AND TableID <> " & mlngImportBaseTable & " AND ImportID = " & mlngImportID, adOpenForwardOnly, adLockReadOnly)
  
  Do Until prstDetails.EOF
    If mlngImportParent1Table = 0 Then
      mlngImportParent1Table = prstDetails.Fields(0)
      mstrImportParent1TableName = datGeneral.GetTableName(mlngImportParent1Table)
    Else
      mlngImportParent2Table = prstDetails.Fields(0)
      mstrImportParent2TableName = datGeneral.GetTableName(mlngImportParent2Table)
    End If
    prstDetails.MoveNext
  Loop
 

  Set prstDetails = Nothing

  GetImportDetails = True
  Exit Function
  
GetImportDetails_ERROR:
  
  GetImportDetails = False
  mstrErrorString = "Error whilst retrieving the import details" & vbCrLf & vbCrLf & "(" & Err.Description & ")"
  
End Function

Private Function GetRealSources() As Boolean

  ' Purpose : This function retrieves the real source of the (upto) 3 tables
  '           that can be used in the import.
  ' Input   : None
  ' Output  : True/False.

  Dim pobjTableView As CTablePrivilege
  
  On Error GoTo ErrTrap

  'ReDim mstrBaseTableRealSources(-1)
  'ReDim mstrParent1TableRealSources(-1)
  'ReDim mstrParent2TableRealSources(-1)


  ' Get real source for the base table (necessary)
  If Not GetRealSource(mlngImportBaseTable) Then
    GetRealSources = False
    Exit Function
  End If
  
  ' Get real source for the parent 1 table (optional)
  If mlngImportParent1Table > 0 Then
    Set pobjTableView = gcoTablePrivileges.FindTableID(mlngImportParent1Table)
    mstrParent1TableRealSource = pobjTableView.RealSource
    Set pobjTableView = Nothing

    If Not GetRealSource(mlngImportParent1Table) Then
      GetRealSources = False
      Exit Function
    End If
  End If
  
  ' Get real source for the parent 2 table (optional)
  If mlngImportParent2Table > 0 Then
    Set pobjTableView = gcoTablePrivileges.FindTableID(mlngImportParent2Table)
    mstrParent2TableRealSource = pobjTableView.RealSource
    Set pobjTableView = Nothing

    If Not GetRealSource(mlngImportParent2Table) Then
      GetRealSources = False
      Exit Function
    End If
  End If
  
  GetRealSources = True
  Exit Function
  
ErrTrap:
  
  mstrErrorString = "Error retrieving real sources." & vbCrLf & vbCrLf & "(" & Err.Description & ")"
  GetRealSources = False
  
End Function

Private Function GetRealSource(plngTableID As Long) As Boolean

  ' Purpose : This function checks users privileges on tables and columns
  '           used in the import.
  ' Input   : None
  ' Output  : True/False.

  Dim pintTemp As Integer
  Dim pobjTableView As CTablePrivilege
  Dim pobjColumnPrivileges As CColumnPrivileges
  Dim pblnAllowed As Boolean
  Dim pblnInsertOnly As Boolean
  Dim strTableName As String
  Dim intBaseTableColumnCount As Integer

  GetRealSource = False
  pblnInsertOnly = (mlngImportType = 2)
  If plngTableID = mlngImportBaseTable Then
    mblnCanInsert = False
    mblnCanUpdate = False
  End If

  ' Loop through each table/view in the collection
  For Each pobjTableView In gcoTablePrivileges.Collection
  
    ' If the tableID matches and we have insert/update privilege then continue
    If pobjTableView.TableID = plngTableID Then 'pobjTableView.AllowUpdate) Then

      If plngTableID = mlngImportBaseTable Then
        mblnCanInsert = mblnCanInsert Or pobjTableView.AllowInsert
        mblnCanUpdate = mblnCanUpdate Or pobjTableView.AllowUpdate
      End If


      'If Insert only then must have insert permission (or let it through if its a parent table).
      If (mblnCanInsert Or Not pblnInsertOnly) Or (plngTableID <> mlngImportBaseTable) Then

        ' Get a list of column privileges for the objtableview
        If pobjTableView.IsTable Then
          Set pobjColumnPrivileges = GetColumnPrivileges(pobjTableView.TableName)
        Else
          Set pobjColumnPrivileges = GetColumnPrivileges(pobjTableView.ViewName)
        End If
    
        intBaseTableColumnCount = 0
        'Loop through those columns checking privileges
        For pintTemp = 1 To UBound(mvarColDetails, 2)
          
          ' Is it a column (and not a filler)
          If mvarColDetails(0, pintTemp) = "C" And mvarColDetails(1, pintTemp) = plngTableID Then
            
            
            'MH20090114 Fault 13496
            If plngTableID = mlngImportBaseTable Then
              intBaseTableColumnCount = intBaseTableColumnCount + 1
            End If
            
            
            'MH20090114 Fault 13494
            ''' Is its not valid (ie, not in the objtableview) then exit for
            'If Not pobjColumnPrivileges.IsValid(mvarColDetails(4, pintTemp)) Then Exit For
            pblnAllowed = (pobjColumnPrivileges.IsValid(mvarColDetails(4, pintTemp)))
            If Not pblnAllowed Then
              'Do not have all required columns for this view
              mstrErrorString = mvarColDetails(2, pintTemp) & "." & mvarColDetails(4, pintTemp) & " : No read permission"
              Exit For
            End If
            
            
            ' If its the base table, check INSERT privilege
            If plngTableID = mlngImportBaseTable Then
              pblnAllowed = pobjColumnPrivileges.item(mvarColDetails(4, pintTemp)).AllowUpdate
              If Not pblnAllowed Then mstrErrorString = mvarColDetails(2, pintTemp) & "." & mvarColDetails(4, pintTemp) & " : No update permission"
            Else ' Otherwise just check the select privilege
              pblnAllowed = pobjColumnPrivileges.item(mvarColDetails(4, pintTemp)).AllowSelect
              If Not pblnAllowed Then mstrErrorString = mvarColDetails(2, pintTemp) & "." & mvarColDetails(4, pintTemp) & " : No read permission"
            End If
            
            
            If Not pblnAllowed Then Exit For
          
          End If
    
        Next pintTemp


        If plngTableID = mlngImportBaseTable And intBaseTableColumnCount = 0 Then
          mstrErrorString = "Invalid definition - no columns selected from the base table."
          Exit Function
        End If


        If pblnAllowed Then
          GetRealSource = True
          
          Select Case plngTableID
          Case mlngImportBaseTable
            AddToArray mstrBaseTableRealSources(), pobjTableView.RealSource
            If pobjTableView.IsTable Then
              'Direct Table Access so exit function
              Exit Function
            Else
              mstrBaseTableViewJoins = mstrBaseTableViewJoins & _
                  " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & _
                  pobjTableView.RealSource & ".ID = " & mstrImportBaseTableName & ".ID"
            End If
          Case mlngImportParent1Table
            AddToArray mstrParent1TableRealSources(), pobjTableView.RealSource
            If pobjTableView.IsTable Then
              'Direct Table Access so exit function
              Exit For
            Else
              mstrParent1ViewJoins = mstrParent1ViewJoins & _
                  " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & _
                  pobjTableView.RealSource & ".ID = " & mstrImportParent1TableName & ".ID" & vbCrLf
            End If
          Case mlngImportParent2Table
            AddToArray mstrParent2TableRealSources(), pobjTableView.RealSource
            If pobjTableView.IsTable Then
              'Direct Table Access so exit function
              Exit For
            Else
              mstrParent2ViewJoins = mstrParent2ViewJoins & _
                  " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & _
                  pobjTableView.RealSource & ".ID = " & mstrImportParent2TableName & ".ID"
            End If
          End Select

        End If

      End If
    End If

  Next pobjTableView

  ' Check that there is a realsource for each table used in the import. If not, then
  ' the user does not have permission to run the import

  If Not GetRealSource Then
  
    If mstrErrorString <> "" Then
      mstrErrorString = vbCrLf & IIf(mstrErrorString = "", "", mstrErrorString)
    Else
      mstrErrorString = "No " & IIf(pblnInsertOnly, "insert", "update") & " permission."
    End If
    
    Select Case plngTableID
    Case mlngImportBaseTable
      strTableName = mstrImportBaseTableName
    Case mlngImportParent1Table
      strTableName = mstrImportParent1TableName
    Case mlngImportParent2Table
      strTableName = mstrImportParent2TableName
    End Select
    mstrErrorString = "Insufficient privileges on the '" & strTableName & "' table." & vbCrLf & mstrErrorString

  End If
  
End Function

Private Function CheckIfImportIsValid() As Boolean
  
  ' Purpose : This function checks that any read only or mandatory columns
  '           are included in the import
  ' Input   : None
  ' Output  : True/False.

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strSQL1 As String
  Dim strColumnsInDef As String
  Dim strErrorColumns As String
  
  strColumnsInDef = "SELECT ASRSysImportDetails.ColExprID " & _
                   "FROM ASRSysImportDetails " & _
                   "WHERE ASRSysImportDetails.ImportID = " & CStr(mlngImportID)

  'This will retrieve all of the read only destination columns in the import
  'which are from the base table only
  
  '15/08/2001 MH Fault 2274
  'strSQL1 = "SELECT ASRSysTables.TableName + '.' + ASRSysColumns.ColumnName as 'TableColumn', " & _
            "'Read Only' as Reason " & _
           "FROM ASRSysImportDetails " & _
           "JOIN ASRSysColumns ON ASRSysColumns.ColumnID = ASRSysImportDetails.ColExprID " & _
           "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysImportDetails.TableID " & _
           "WHERE ASRSysImportDetails.ImportID = " & CStr(mlngImportID) & _
           " AND ASRSysColumns.ReadOnly <> 0" & _
           " AND ASRSysTables.TableName = '" & mstrImportBaseTableName & "'"
  strSQL1 = "SELECT ASRSysTables.TableName + '.' + ASRSysColumns.ColumnName as 'TableColumn', " & _
            "'Read Only' as Reason " & _
           "FROM ASRSysImportDetails " & _
           "JOIN ASRSysColumns ON ASRSysColumns.ColumnID = ASRSysImportDetails.ColExprID " & _
           "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysImportDetails.TableID " & _
           "WHERE ASRSysImportDetails.ImportID = " & CStr(mlngImportID) & _
           " AND ASRSysImportDetails.KeyField = 0" & _
           " AND ASRSysColumns.ReadOnly <> 0" & _
           " AND ASRSysTables.TableName = '" & mstrImportBaseTableName & "'"

  'MH20000904
  'Allow save if mandatory ommitted and it is a calculated column

  'This will get all of the mandatory columns which have not been included
  'in the definition which are from the base table only
'  strSQL2 = "SELECT ASRSysTables.TableName+'.'+ASRSysColumns.ColumnName as 'TableColumn', " & _
'            "'Mandatory' as Reason " & _
'           "FROM ASRSysColumns " & _
'           "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID " & _
'           "WHERE ASRSysColumns.TableID = " & mlngImportBaseTable & _
'           " AND ASRSysColumns.ColumnID NOT IN (" & strColumnsInDef & ") " & _
'           " AND Mandatory = '1'" & _
'           " AND RTrim(defaultvalue) = '' " & _
'           " AND CONVERT(int,dfltValueExprID) = 0 " & _
'           " AND CalcExprID = 0 " & _
'           " AND ASRSysTables.TableName = '" & mstrImportBaseTableName & "'"
'
'  strSQL = strSQL1 & " UNION " & strSQL2 & " ORDER BY 'TableColumn'"

  strSQL = strSQL1
  
  Set rsTemp = mclsData.OpenRecordset(strSQL, adOpenDynamic, adLockReadOnly)

  CheckIfImportIsValid = (rsTemp.EOF)

  If Not CheckIfImportIsValid Then

    strErrorColumns = vbNullString
    While Not rsTemp.EOF
      strErrorColumns = strErrorColumns & _
          rsTemp!Reason & " :" & vbTab & _
          rsTemp!TableColumn & vbCrLf
      rsTemp.MoveNext
    Wend

    mstrErrorString = "Unable to run this Import for the following " & _
                      "reasons:" & vbCrLf & vbCrLf & _
                       strErrorColumns
  End If

  rsTemp.Close
  Set rsTemp = Nothing

End Function

Private Function ProcessFile() As Boolean

  ' Purpose : This function contains the main loop which reads in a line
  '           of the file and runs the insert/update functions
  ' Input   : None
  ' Output  : True/False.

  Dim fOK As Boolean
  Dim iCounter As Long
  Dim lngLineCount As Long
  Dim blnRecordOK As Boolean
  
  Dim objFilter As clsExprExpression
  Dim sChar As String
  Dim lngCurrentPos As Long
  
  Dim strSearchChar As String
  
  On Error GoTo ErrTrap
  
  ' Set flag to true to start with
  fOK = True
  
  Set objFilter = New clsExprExpression
  ReDim mavColumns(2, 0)
  ReDim maryRecordSuccessIDs(0)
  
  ' Create an array of the column IDs, and values to pass to the expression.
  ' Column 1= column id.
  ' Coulmn 2 = import value.
  For iCounter = 1 To UBound(mvarColDetails, 2)
    If mvarColDetails(3, iCounter) > 0 Then
      ReDim Preserve mavColumns(2, UBound(mavColumns, 2) + 1)
      mavColumns(1, UBound(mavColumns, 2)) = mvarColDetails(3, iCounter)
    End If
  Next iCounter
  
  If mlngFilterID > 0 Then
    ' Create the expression object fo rth file filter.
    objFilter.Initialise 0, mlngFilterID, giEXPR_UTILRUNTIMEFILTER, giEXPRVALUE_LOGIC
    objFilter.ConstructExpression
  End If
  
  ' If we are importing an excel spreadsheet, open it in excel, save as csv,
  ' then treat as a comma delimited file from here on (!)
  If mintImportFileType = ImportType.ExcelWorksheet Then
    If Not ConvertExcelFile Then
      ProcessFile = False
      Exit Function
    End If
  End If

  mstrErrorString = "File not found."
  If Dir(mstrImportFileName) = vbNullString Then
    ProcessFile = False
    Exit Function
  End If
  mstrErrorString = vbNullString
  
  ' Get the total number of records in the file, and set progress bar maximum
  GetFileLineCount mintImportFileType
  If mlngTotalRecords < 1 Then
    'MH20040325 Fault 8331
    'If no records then successful!
    ProcessFile = True
    Exit Function
  End If

  ' Open the file for input
  mintFileNo = FreeFile
  ' JPD20011012 Fault 2948 Open as binary file to avoid errors when the file contains dodgy characters.
  'Open mstrImportFileName For Input As #mintFileNo
  Open mstrImportFileName For Binary Access Read As #mintFileNo


  'MH20071003
  '' If we are to ignore the first line, then read it in, but dont process it
  ''If mblnImportIgnoreFirstLine Then Line Input #mintFileNo, mstrBufferString
  
  'NPG20090713 Fault HRPRO-39
  ' Moved the header line skip to within the file read - line input no good for multi line data...
  '  For lngLineCount = 1 To mintImportHeaderLines
  '    Line Input #mintFileNo, mstrBufferString
  '  Next

  lngLineCount = 0
  Do
    ' Read the next line of the file
    ' JPD20011019 Fault 2990
    'Line Input #mintFileNo, mstrBufferString
    
    mstrBufferString = ""
    Do While lngCurrentPos < LOF(mintFileNo)   ' Loop until end of file.
      sChar = Input(1, #mintFileNo)   ' Get one character.
      lngCurrentPos = Loc(mintFileNo)
          
      'NPG20090713 Fault HRPRO-39
      If mstrEOLChars = vbCr Or mstrEOLChars = vbCrLf Then strSearchChar = vbCr
      If mstrEOLChars = vbLf Then strSearchChar = vbLf
      
      If sChar = strSearchChar And (UBound(Split(mstrBufferString, mstrImportEncapsulator)) Mod 2) = 0 Then
        Exit Do
      End If
      If sChar <> vbLf Or mstrBufferString <> "" Then
        mstrBufferString = mstrBufferString & sChar
      End If
    Loop
    
    lngLineCount = lngLineCount + 1
    
    'NPG20090713 Fault HRPRO-39 '  Here's the new header line skip...
    If lngLineCount > mintImportHeaderLines Then
      mstrErrorString = vbNullString
      If Len(mstrBufferString) > 0 Then
      ' JPD20010907 No Fault - Only insert/update the import record if it satisfies the
        ' defined filter criteria.
        If CheckFilter(objFilter) Then
  
          ReDim mstrLookupArray(0)  'MH20030520
  
          If mblnImportCreateNewOnly Then
            ' If we are inserting new records only, then run an INSERT
            blnRecordOK = DoInsert
          Else
            ' Create the Where code and then run the appropriate function
            blnRecordOK = FindMatchingRecord
          End If
      
          If mbUseUpdateBlob And blnRecordOK Then
            maryRecordSuccessIDs(UBound(maryRecordSuccessIDs)) = GetSpecificKeyData(UBound(mvarColDetails, 2))
            ReDim Preserve maryRecordSuccessIDs(UBound(maryRecordSuccessIDs) + 1)
          End If
  
          If blnRecordOK Then
            CheckIfInsertIntoLookups  'MH20030520
          End If
  
        End If
        
        ' Update progress bar
        If gblnBatchMode Then
          gobjProgress.UpdateProgress2 False
          gobjProgress.Bar2RecordsCaption = _
            "Records Processed : " & mlngSuccessRecords & " successful" & _
            IIf(mlngFailRecords > 0, ", " & mlngFailRecords & " failed", vbNullString)
        Else
          gobjProgress.UpdateProgress False
          gobjProgress.Bar1RecordsCaption = _
            "Records Processed : " & mlngSuccessRecords & " successful" & _
            IIf(mlngFailRecords > 0, ", " & mlngFailRecords & " failed", vbNullString)
        End If
      End If
      
      Select Case mintImportFileType
        Case ImportType.DelimitedFile, ImportType.FixedLengthFile, ImportType.SQLTable
          If lngLineCount = (mlngTotalRecords + mintImportHeaderLines) Then Exit Do
        Case ImportType.ExcelWorksheet
          'See code line page-up with 'ConvertExcelFile' most import files are
          'converted to CSVs so this part of code is rarely run
          If lngLineCount >= mlngTotalRecords Then Exit Do
      End Select
      
      ' If a DoInsert or DoUpdate have failed (codewise) then stop here
      If Not fOK Then Exit Do
      
      ' If user cancels the export, abort
      If gobjProgress.Cancelled Then
        mblnUserCancelled = True
        Close #mintFileNo
        ProcessFile = False
        Exit Function
      End If
    End If
  Loop
  
  Set objFilter = Nothing
  
  ' Tidy Up
  If fOK Then
    Close #mintFileNo
    ProcessFile = True
    DoEvents ' to show the complete progress bar before it disappears
    Exit Function
  End If

ErrTrap:
  
  If mstrErrorString = "" Then mstrErrorString = "Error processing import file."
  Close #mintFileNo
  ProcessFile = False
  
End Function
    
Private Function DoUpdate() As Boolean

  ' Purpose : This function attempts to update a record with the data from
  '           the import file
  ' Input   : None
  ' Output  : True/False.

  On Local Error GoTo UpdateErrTrap

  Dim fColFound As Boolean
  Dim fSavedOK As Boolean
  Dim pintCounter As Integer
  Dim strSplit() As String
  Dim lngIndex As Long
  Dim strJoinArray() As String

  Dim lngCount As Long
  Dim blnFound As Boolean
  
  ' Set flag to be ok to start with
  fSavedOK = True

  ' Begin a transaction here.
  gADOCon.BeginTrans

  mstrSQL = vbNullString
  'mstrSQL = mstrSQL & "UPDATE " & mstrBaseTableRealSource & " SET "

  fColFound = False
  For pintCounter = 1 To UBound(mvarColDetails, 2)
    'If mvarColDetails(0, pintCounter) = "C" And mvarColDetails(5, pintCounter) = False Then
    If mvarColDetails(0, pintCounter) = "C" And mvarColDetails(13, pintCounter) = False Then
      If mvarColDetails(1, pintCounter) = mlngImportBaseTable Then
        mstrSQL = mstrSQL & _
              IIf(mstrSQL <> vbNullString, ", ", "") & _
              mvarColDetails(4, pintCounter) & " = " & _
              GetSpecificKeyData(pintCounter)
        If mstrErrorString <> vbNullString Then
          gADOCon.RollbackTrans   'MH20040422 Fault 8277
          RecordError mstrErrorString, True 'NHRD21122004 Fault 9605
          DoUpdate = False
          Exit Function
        End If
        
        AddToLookupArray pintCounter
        fColFound = True
      End If
    End If
  Next pintCounter


Dim pstrParentWhereSQL As String
Dim sTemp As String
Dim strJoin As String

''###############
'  ' Do the second parent table if required
  If (mlngImportParent2Table > 0) Then

    pstrParentWhereSQL = ""
    strJoin = vbNullString
    ReDim strJoinArray(0) As String
    strJoinArray(0) = mstrParent2TableRealSource  'MH20050414 Fault 10004

    For pintCounter = 1 To UBound(mvarColDetails, 2)
      If mvarColDetails(1, pintCounter) = mlngImportParent2Table Then
        'If mvarColDetails(5, pintCounter) = False Then
          pstrParentWhereSQL = pstrParentWhereSQL & IIf(Len(pstrParentWhereSQL) > 0, " AND ", " WHERE ")

          'MH20050414 Fault 10004
          'If UBound(mstrParent2TableRealSources) = 0 Then
          '  pstrParentWhereSQL = pstrParentWhereSQL & _
          '        mstrParent2TableRealSources(0) & "." & mvarColDetails(4, pintCounter)
          'Else
            pstrParentWhereSQL = pstrParentWhereSQL & _
                  "CASE"
            For lngIndex = 0 To UBound(mstrParent2TableRealSources)
              pstrParentWhereSQL = pstrParentWhereSQL & _
                  " WHEN NOT " & mstrParent2TableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter) & _
                  " IS NULL THEN " & mstrParent2TableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter)

              'If lngIndex > 0 Then
                blnFound = False
                For lngCount = 0 To UBound(strJoinArray)
                  If strJoinArray(lngCount) = mstrParent2TableRealSources(lngIndex) Then
                    blnFound = True
                    Exit For
                  End If
                Next

                If Not blnFound Then
                  lngCount = UBound(strJoinArray) + 1
                  ReDim Preserve strJoinArray(lngCount) As String
                  strJoinArray(lngCount) = mstrParent2TableRealSources(lngIndex)

                  strJoin = strJoin & _
                      " LEFT OUTER JOIN " & mstrParent2TableRealSources(lngIndex) & _
                      " ON " & mstrParent2TableRealSources(lngIndex) & ".ID = " & mstrImportParent2TableName & ".ID "
                End If
              'End If

            Next
            pstrParentWhereSQL = pstrParentWhereSQL & _
                  " ELSE null END"
          'End If


          'pstrParentWhereSQL = pstrParentWhereSQL & " = " & GetSpecificKeyData(pintCounter)
          sTemp = GetSpecificKeyData(pintCounter)
          If mstrErrorString <> vbNullString Then
            Exit Function
          End If

          'MH20050124 Fault 9592 - If null date then use "IS NULL" and not "= NULL"
          If mvarColDetails(7, pintCounter) = True And sTemp = "null" Then
            pstrParentWhereSQL = pstrParentWhereSQL & " is null"
          Else
            pstrParentWhereSQL = pstrParentWhereSQL & " = " & sTemp
          End If

        'End If
      End If
    Next pintCounter

    mstrSQL = mstrSQL & _
      IIf(mstrSQL <> vbNullString, ", ", "") & _
      "ID_" & CStr(mlngImportParent2Table) & " = (SELECT " & mstrParent2TableRealSource & ".ID FROM " & _
      mstrParent2TableRealSource & strJoin & pstrParentWhereSQL & ")"

  End If


'###############
  If (mlngImportParent1Table > 0) Then

    pstrParentWhereSQL = ""
    strJoin = vbNullString
    ReDim strJoinArray(0) As String
    strJoinArray(0) = mstrParent1TableRealSource  'MH20050414 Fault 10004

    For pintCounter = 1 To UBound(mvarColDetails, 2)
      If mvarColDetails(1, pintCounter) = mlngImportParent1Table Then
        'If mvarColDetails(5, pintCounter) = False Then
          pstrParentWhereSQL = pstrParentWhereSQL & IIf(Len(pstrParentWhereSQL) > 0, " AND ", " WHERE ")

          'MH20050414 Fault 10004
          'If UBound(mstrParent1TableRealSources) = 0 Then
          '  pstrParentWhereSQL = pstrParentWhereSQL & _
          '        mstrParent1TableRealSources(0) & "." & mvarColDetails(4, pintCounter)
          'Else
            pstrParentWhereSQL = pstrParentWhereSQL & _
                  "CASE"
            For lngIndex = 0 To UBound(mstrParent1TableRealSources)
              pstrParentWhereSQL = pstrParentWhereSQL & _
                  " WHEN NOT " & mstrParent1TableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter) & _
                  " IS NULL THEN " & mstrParent1TableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter)

              'If lngIndex > 0 Then
                blnFound = False
                For lngCount = 0 To UBound(strJoinArray)
                  If strJoinArray(lngCount) = mstrParent1TableRealSources(lngIndex) Then
                    blnFound = True
                    Exit For
                  End If
                Next

                If Not blnFound Then
                  lngCount = UBound(strJoinArray) + 1
                  ReDim Preserve strJoinArray(lngCount) As String
                  strJoinArray(lngCount) = mstrParent1TableRealSources(lngIndex)

                  strJoin = strJoin & _
                      " LEFT OUTER JOIN " & mstrParent1TableRealSources(lngIndex) & _
                      " ON " & mstrParent1TableRealSources(lngIndex) & ".ID = " & mstrImportParent1TableName & ".ID "
                End If
              'End If

            Next
            pstrParentWhereSQL = pstrParentWhereSQL & _
                  " ELSE null END"
          'End If


          'pstrParentWhereSQL = pstrParentWhereSQL & " = " & GetSpecificKeyData(pintCounter)
          sTemp = GetSpecificKeyData(pintCounter)
          If mstrErrorString <> vbNullString Then
            Exit Function
          End If

          'MH20050124 Fault 9592 - If null date then use "IS NULL" and not "= NULL"
          If mvarColDetails(7, pintCounter) = True And sTemp = "null" Then
            pstrParentWhereSQL = pstrParentWhereSQL & " is null"
          Else
            pstrParentWhereSQL = pstrParentWhereSQL & " = " & sTemp
          End If

        'End If
      End If
    Next pintCounter

    mstrSQL = mstrSQL & _
      IIf(mstrSQL <> vbNullString, ", ", "") & _
      "ID_" & CStr(mlngImportParent1Table) & " = (SELECT " & mstrParent1TableRealSource & ".ID FROM " & _
      mstrParent1TableRealSource & strJoin & pstrParentWhereSQL & ")"

  End If

'##############

  If fColFound Then
    'mstrSQL = Left(mstrSQL, (Len(mstrSQL) - 1))
    For lngIndex = 0 To UBound(mstrBaseTableRealSources)
      'MH20040615 Fault 8806
      'mstrSQL = "UPDATE " & mstrBaseTableRealSources(lngIndex) & " SET " & mstrSQL & _
                " WHERE ID IN (" & mstrIDsToUpdate & ")"
                '" WHERE " & mstrSQLWhere & " AND " & mstrBaseTableRealSources(lngIndex) & ".ID IN (" & mstrIDsToUpdate & ")"
      'fSavedOK = mclsDeadlock.UpdateTableRecordJustDoIt(mstrSQL)
            
      fSavedOK = mclsDeadlock.UpdateTableRecordJustDoIt( _
                "UPDATE " & mstrBaseTableRealSources(lngIndex) & _
                " SET " & mstrSQL & _
                " WHERE ID IN (" & mstrIDsToUpdate & ")")
                                
      If fSavedOK Then
        Exit For
      End If
    Next

    If fSavedOK Then
      
      gADOCon.CommitTrans
      
      'MH20030522 Fault 3838
      strSplit = Split(mstrIDsToUpdate, ",")
      For lngIndex = 0 To UBound(strSplit)
      
        'JDM - 13/12/01 - Fault 3280 - Log successful records
        If mbLoggingImportSuccess Then
          gobjEventLog.AddDetailEntry GetRecordDesc(CLng(strSplit(lngIndex))) & " Updated successfully"
        End If
          
        mlngSuccessRecords = mlngSuccessRecords + 1
      Next

    
    Else
      gADOCon.RollbackTrans
      RecordError mclsDeadlock.ErrorString
      If mclsDeadlock.FatalError Then
        DoUpdate = False
        Exit Function
      End If
    
    End If
  
  Else
    gADOCon.RollbackTrans
    DoUpdate = False
    Exit Function

  End If
  
  'MH20021014
  'objEmail.SendImmediateEmails
  
  DoUpdate = True
  Exit Function
  
UpdateErrTrap:
  
  If Err.Number < 0 Then Resume Next
  gADOCon.RollbackTrans
  DoUpdate = False
  Resume Next

End Function
    
'Private Function DoUpdate() As Boolean
'
'  ' Purpose : This function attempts to update a record with the data from
'  '           the import file
'  ' Input   : None
'  ' Output  : True/False.
'
'  On Local Error GoTo UpdateErrTrap
'
'  Dim cmADO As ADODB.Command
'  Dim pmADO As ADODB.Parameter
'  Dim ADOErr As ADODB.Error
'  Dim fSavedOK As Boolean
'  Dim pintCounter As Integer
'  Dim iOldCursorLocation As Integer
'
'  ' Set flag to be ok to start with
'  fSavedOK = True
'
'  ' Begin a transaction here.
'  gADOCon.BeginTrans
'
'  mstrSQL = vbNullString
'  mstrSQL = mstrSQL & "UPDATE " & mstrBaseTableRealSource & " SET "
'
'  For pintCounter = 1 To UBound(mvarColDetails, 2)
'    If mvarColDetails(0, pintCounter) = "C" And mvarColDetails(5, pintCounter) = False Then
'      mstrSQL = mstrSQL & mvarColDetails(4, pintCounter) & " = " & _
'                          GetSpecificKeyData(pintCounter) & ","
'    End If
'  Next pintCounter
'
'  mstrSQL = Left(mstrSQL, (Len(mstrSQL) - 1))
'  mstrSQL = mstrSQL & mstrSQLWhere
'  mstrSQL = mstrSQL & " AND " & mstrBaseTableRealSource & ".ID IN (" & mstrIDsToUpdate & ")"
'
'  ' Change the cursor location to 'client' as the errors that might be raised
'  ' during the update cannot be read for 'server' cursors.
'  iOldCursorLocation = gADOCon.CursorLocation
'  gADOCon.Errors.Clear
'  gADOCon.CursorLocation = adUseClient
'
'  gADOCon.Execute mstrSQL
'
'  ' Restore the original cursor location to the ADO connection object.
'  gADOCon.CursorLocation = iOldCursorLocation
'
'  ' Check if the update produced any errors.
'  If gADOCon.Errors.Count > 0 Then
'
'    fSavedOK = False
'    mstrErrorString = vbNullString
'
'    For Each ADOErr In gADOCon.Errors
'      mstrErrorString = mstrErrorString & IIf(Len(mstrErrorString) > 0, vbCrLf, "") & _
'        ADOErr.Description
'    Next ADOErr
'
'    gADOCon.Errors.Clear
'
'  End If
'
'  Set cmADO = Nothing
'
'  If fSavedOK Then
'    mlngSuccessRecords = mlngSuccessRecords + 1
'    gADOCon.CommitTrans
'  Else
'    mlngFailRecords = mlngFailRecords + 1
'    gADOCon.RollbackTrans
'    gobjEventLog.AddDetailEntry "The following row failed to import (update): " & _
'                                 vbCrLf & vbCrLf & mstrBufferString & _
'                                 vbCrLf & vbCrLf & mstrErrorString
'  End If
'
'  DoUpdate = True
'  Exit Function
'
'UpdateErrTrap:
'
'  If Err.Number < 0 Then Resume Next
'  gADOCon.RollbackTrans
'  DoUpdate = False
'  Resume Next
'
'End Function
'
'Private Function FindMatchingRecord() As Boolean
'
'  ' Purpose : This function compiles the WHERE clause which is used to
'  '           determine if we will be doing an insert or update
'  ' Input   : None
'  ' Output  : True/False.
'
'  On Error GoTo ErrTrap
'
'
'  Dim pintCounter As Integer
'  Dim pstrBaseWhereSQL As String
'  Dim mstrParent1Where As String
'  Dim mstrParent2Where As String
'  Dim pstrFindRecordSQL As String
'  Dim prstFindRecord As Recordset
'  Dim plngRecordsReturned As Long
'  Dim blnParentExists As Boolean
'  Dim blnParentRequired As Boolean
'  Dim strTemp As String
'  Dim lngIndex As Long
'
'
''-------------------------------------------------------------------
'
'  ' Do the base table key fields first. We are to return the ID of the existing
'  ' row in the base table that we are updating
'  For pintCounter = 1 To UBound(mvarColDetails, 2)
'    If mvarColDetails(1, pintCounter) = mlngImportBaseTable Then
'      If mvarColDetails(5, pintCounter) = True Then
'        pstrBaseWhereSQL = pstrBaseWhereSQL & IIf(Len(pstrBaseWhereSQL) > 0, " AND ", " WHERE ")
'
'        'pstrBaseWhereSQL = pstrBaseWhereSQL & mstrBaseTableRealSource & "." & mvarColDetails(4, pintCounter)
'        If UBound(mstrBaseTableRealSources) = 0 Then
'          pstrBaseWhereSQL = pstrBaseWhereSQL & _
'                mstrBaseTableRealSources(0) & "." & mvarColDetails(4, pintCounter)
'        Else
'          pstrBaseWhereSQL = pstrBaseWhereSQL & _
'                "CASE"
'          For lngIndex = 0 To UBound(mstrBaseTableRealSources)
'            pstrBaseWhereSQL = pstrBaseWhereSQL & vbCrLf & _
'                " WHEN NOT " & mstrBaseTableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter) & _
'                " IS NULL THEN " & mstrBaseTableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter)
'          Next
'          pstrBaseWhereSQL = pstrBaseWhereSQL & _
'                " ELSE null END"
'        End If
'
'        pstrBaseWhereSQL = pstrBaseWhereSQL & " = " & GetSpecificKeyData(pintCounter)
'        If mstrErrorString <> vbNullString Then
'          FindMatchingRecord = False
'          Exit Function
'        End If
'      End If
'    End If
'  Next pintCounter
'
'  blnParentExists = False
'
'  ' Do the first parent table if required
'  If mlngImportParent1Table > 0 Then
'    For pintCounter = 1 To UBound(mvarColDetails, 2)
'      If mvarColDetails(1, pintCounter) = mlngImportParent1Table Then
'        If mvarColDetails(5, pintCounter) = True Then
'
'          strTemp = GetSpecificKeyData(pintCounter)
'          If mstrErrorString <> vbNullString Then
'            FindMatchingRecord = False
'            Exit Function
'          End If
'
'          ' JPD20011022 Faults 2996 & 2997 Disable Lofty's error handling.
'          gobjErrorStack.Disable
'          'Set prstFindRecord = datGeneral.GetReadOnlyRecords("SELECT ID FROM " & mstrImportParent1TableName & " WHERE " & mvarColDetails(4, pintCounter) & " = " & strTemp)
'          Set prstFindRecord = datGeneral.GetReadOnlyRecords("SELECT " & mstrImportParent1TableName & ".ID FROM " & mstrImportParent1TableName & mstrParent1ViewJoins & " " & mstrParent1Where)
'          gobjErrorStack.Enable
'
'          ' If the key field data is missing...
'          If Not (prstFindRecord.BOF And prstFindRecord.EOF) Then
'            mstrParent1Where = mstrParent1Where & IIf(Len(mstrParent1Where) > 0 Or Len(pstrBaseWhereSQL) > 0, " AND ", " WHERE ")
'
'            'mstrParent1Where = mstrParent1Where & mstrBaseTableRealSource & ".ID_" & mlngImportParent1Table
'            If UBound(mstrParent1TableRealSources) = 0 Then
'              mstrParent1Where = mstrParent1Where & _
'                    mstrParent1TableRealSources(0) & "." & mvarColDetails(4, pintCounter)
'            Else
'              mstrParent1Where = mstrParent1Where & _
'                    "CASE"
'              For lngIndex = 0 To UBound(mstrParent1TableRealSources)
'                mstrParent1Where = mstrParent1Where & vbCrLf & _
'                    " WHEN NOT " & mstrParent1TableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter) & _
'                    " IS NULL THEN " & mstrParent1TableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter)
'              Next
'              mstrParent1Where = mstrParent1Where & _
'                    " ELSE null END"
'            End If
'
'            mstrParent1Where = mstrParent1Where & " IN (" & GetSpecificKeyData(pintCounter) & ")"
'            If mstrErrorString <> vbNullString Then
'              FindMatchingRecord = False
'              Exit Function
'            End If
'
'            ' JPD20011022 Faults 2996 & 2997 Disable Lofty's error handling.
'            gobjErrorStack.Disable
'            'Set prstFindRecord = datGeneral.GetReadOnlyRecords("SELECT ID FROM " & mstrParent1TableRealSource & " WHERE " & mvarColDetails(4, pintCounter) & " = " & strTemp)
'            Set prstFindRecord = datGeneral.GetReadOnlyRecords("SELECT " & mstrImportParent1TableName & ".ID FROM " & mstrImportParent1TableName & mstrParent1ViewJoins & " " & mstrParent1Where)
'            gobjErrorStack.Enable
'
'            Do Until prstFindRecord.EOF
'              mstrParent1Where = mstrParent1Where & prstFindRecord.Fields(0) & ","
'              prstFindRecord.MoveNext
'            Loop
'            Set prstFindRecord = Nothing
'            mstrParent1Where = Left(mstrParent1Where, Len(mstrParent1Where) - 1) & ")"
'
'            blnParentExists = True
'
'          End If
'
'        End If
'      End If
'    Next pintCounter
'  End If
'
'  ' Do the second parent table if required
'  If mstrErrorString = vbNullString Then
'    For pintCounter = 1 To UBound(mvarColDetails, 2)
'      If mvarColDetails(1, pintCounter) = mlngImportParent2Table Then
'        If mvarColDetails(5, pintCounter) = True Then
'
'          strTemp = GetSpecificKeyData(pintCounter)
'          If mstrErrorString <> vbNullString Then
'            FindMatchingRecord = False
'            Exit Function
'          End If
'
'          ' JPD20011022 Faults 2996 & 2997 Disable Lofty's error handling.
'          gobjErrorStack.Disable
'          'Set prstFindRecord = datGeneral.GetReadOnlyRecords("SELECT ID FROM " & mstrImportParent2TableName & " WHERE " & mvarColDetails(4, pintCounter) & " = " & strTemp)
'          Set prstFindRecord = datGeneral.GetReadOnlyRecords("SELECT ID FROM " & mstrImportParent2TableName & mstrParent2ViewJoins & " " & mstrParent2Where)
'          gobjErrorStack.Enable
'
'          ' If the key field data is missing...
'          If Not (prstFindRecord.BOF And prstFindRecord.EOF) Then
'            mstrParent2Where = mstrParent2Where & IIf(Len(mstrParent2Where) > 0 Or Len(pstrBaseWhereSQL) > 0, " AND ", " WHERE ")
'
'            'mstrParent2Where = mstrParent2Where & mstrBaseTableRealSource & ".ID_" & mlngImportParent2Table
'            If UBound(mstrParent2TableRealSources) = 0 Then
'              mstrParent2Where = mstrParent2Where & _
'                    mstrParent2TableRealSources(0) & "." & mvarColDetails(4, pintCounter)
'            Else
'              mstrParent2Where = mstrParent2Where & _
'                    "CASE"
'              For lngIndex = 0 To UBound(mstrParent2TableRealSources)
'                mstrParent2Where = mstrParent2Where & vbCrLf & _
'                    " WHEN NOT " & mstrParent2TableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter) & _
'                    " IS NULL THEN " & mstrParent2TableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter)
'              Next
'              mstrParent2Where = mstrParent2Where & _
'                    " ELSE null END"
'            End If
'
'            mstrParent2Where = mstrParent2Where & " IN (" & GetSpecificKeyData(pintCounter) & ")"
'            If mstrErrorString <> vbNullString Then
'              FindMatchingRecord = False
'              Exit Function
'            End If
'
'            ' JPD20011022 Faults 2996 & 2997 Disable Lofty's error handling.
'            gobjErrorStack.Disable
'            'Set prstFindRecord = datGeneral.GetReadOnlyRecords("SELECT ID FROM " & mstrParent2TableRealSource & " WHERE " & mvarColDetails(4, pintCounter) & " = " & strTemp)
'            Set prstFindRecord = datGeneral.GetReadOnlyRecords("SELECT " & mstrImportParent2TableName & ".ID FROM " & mstrImportParent2TableName & mstrParent2ViewJoins & " " & mstrParent2Where)
'            gobjErrorStack.Enable
'
'            Do Until prstFindRecord.EOF
'              mstrParent2Where = mstrParent2Where & prstFindRecord.Fields(0) & ","
'              prstFindRecord.MoveNext
'            Loop
'            Set prstFindRecord = Nothing
'            mstrParent2Where = Left(mstrParent2Where, Len(mstrParent2Where) - 1) & ")"
'
'            blnParentExists = True
'
'          End If
'
'        End If
'      End If
'    Next pintCounter
'  End If
'
'  mstrIDsToUpdate = ""
'  mstrSQLWhere = vbNullString
'  plngRecordsReturned = 0
'
'  If blnParentExists Or (mlngImportParent1Table = 0 And mlngImportParent2Table = 0) Then
'    mstrSQLWhere = mstrSQLWhere & pstrBaseWhereSQL
'    If Len(mstrParent1Where) > 0 Then mstrSQLWhere = mstrSQLWhere & mstrParent1Where
'    If Len(mstrParent2Where) > 0 Then mstrSQLWhere = mstrSQLWhere & mstrParent2Where
'
'    ' JPD20011012 Fault 2950
'    If Len(mstrSQLWhere) > 0 Then
'      ' Now try and retrieve the ID of the parent / base record
'      pstrFindRecordSQL = "SELECT " & mstrImportBaseTableName & ".ID FROM " & mstrImportBaseTableName & mstrBaseTableViewJoins & " " & mstrSQLWhere
'
'      ' JPD20011022 Faults 2996 & 2997 Disable Lofty's error handling.
'      gobjErrorStack.Disable
'      Set prstFindRecord = datGeneral.GetReadOnlyRecords(pstrFindRecordSQL)
'      gobjErrorStack.Enable
'
'      Do Until prstFindRecord.EOF
'        mstrIDsToUpdate = mstrIDsToUpdate & IIf(Len(mstrIDsToUpdate) > 0, ",", "") & prstFindRecord.Fields(0)
'        plngRecordsReturned = plngRecordsReturned + 1
'        prstFindRecord.MoveNext
'      Loop
'
'      Set prstFindRecord = Nothing
'    End If
'  End If
'
'
'  blnParentRequired = (mlngImportParent1Table <> 0 Or mlngImportParent2Table <> 0)
'
'  If blnParentRequired = True And blnParentExists = False Then
'    RecordError "Matching parent record not found"
'  Else
'    Select Case plngRecordsReturned
'    Case 0
'      If mlngImportType = 1 Then  'Update Only
'        RecordError "Matching record not found"
'      Else
'        FindMatchingRecord = DoInsert
'      End If
'
'    Case 1
'      FindMatchingRecord = DoUpdate
'
'    Case Else
'      If mblnImportMultipleRecordAction Then
'        FindMatchingRecord = DoUpdate
'      Else
'        RecordError "Key fields returned multiple records, and definition states not to update."
'      End If
'
'    End Select
'  End If
'
'  Exit Function
'
'
'ErrTrap:
'  RecordError "Error finding matching record (creating 'where' statement)."
'  FindMatchingRecord = False
'
'End Function


Private Function GetSpecificKeyData(pintToken As Integer) As Variant

  ' Purpose : This function returns the specified 'column' of data from
  '           the string that has been imported from the import file
  ' Input   : None
  ' Output  : True/False.

  Dim pintCounter As Integer
  Dim pstrDelimiter As String
  Dim pstrToken As String
  Dim pintCurrentPos As Integer
  Dim pstrDateFormat As String
  Dim sTemp As String
  Dim dblTemp As Double
  Dim strDataType As String
  
  Dim pstrType As String
  Dim pintTableID As Long
  Dim pintColumnID As Long
  
  
  On Error GoTo ErrTrap
  'TM - Set variables to pass to generic NumericColumn function.
  pstrType = mvarColDetails(0, pintToken)
  pintTableID = mvarColDetails(1, pintToken)
  pintColumnID = mvarColDetails(3, pintToken)
  
  ' Read the data depending on the file type
  Select Case mintImportFileType
  
    Case ImportType.DelimitedFile, ImportType.ExcelWorksheet, ImportType.SQLTable
        
      ' Check delimiter. Ie, if Tab selected, replace "<Tab>" with a vbtab
      Select Case LCase(mstrImportDelimiter)
        Case ",": pstrDelimiter = mstrImportDelimiter
        Case "<tab>": pstrDelimiter = vbTab
        Case Else: pstrDelimiter = mstrImportOtherDelimiter
      End Select
        
      pstrToken = ParseStr(mstrBufferString, pintToken, pstrDelimiter, Trim(mstrImportEncapsulator))

      If mvarColDetails(7, pintToken) Then
        strDataType = "date"
        pstrToken = GetDateInSQLFormat(pstrToken)

      ElseIf mvarColDetails(8, pintToken) Then
        strDataType = "numeric"
        If Trim(pstrToken) = "" Then
          pstrToken = "0"
        Else
          dblTemp = CDbl(datGeneral.ConvertNumberForSQL(pstrToken))
          pstrToken = CStr(dblTemp)

          'MH20041111 Fault 9070
          'Check if this is a numeric value and do not allow dates.
          'If CStr(Val(pstrToken)) <> Replace(Trim(Replace(pstrToken, "0", " ")), " ", "0") Then
          If Not CheckIfNumeric(pstrToken) Then
            RecordError "Error converting '" & pstrToken & "' to numeric type value.", False
            Exit Function
          End If

        End If
      
      ElseIf mvarColDetails(9, pintToken) Then
        strDataType = "logic"
        sTemp = Trim(UCase(Left(pstrToken, 1)))
        pstrToken = IIf((sTemp = "F") Or (sTemp = "N") Or (sTemp = "0") Or (sTemp = ""), "0", "1")
      Else
      
      
        pstrToken = "'" & Replace(pstrToken, "'", "''") & "'"
      
      End If

      GetSpecificKeyData = pstrToken


    Case ImportType.FixedLengthFile:
    
      pintCurrentPos = 1
      For pintCounter = 1 To UBound(mvarColDetails, 2)

        If pintCounter = pintToken Then
          ' If column, get the data, else its a filler so get a blank string
          If mvarColDetails(0, pintCounter) = "C" Then
            pstrToken = Trim(Mid(mstrBufferString, pintCurrentPos, mvarColDetails(6, pintCounter)))
          Else
            pstrToken = Space(mvarColDetails(6, pintCounter))
          End If
  
          If mvarColDetails(7, pintCounter) Then
            strDataType = "date"
            pstrToken = GetDateInSQLFormat(pstrToken)

          ElseIf mvarColDetails(8, pintCounter) Then
            strDataType = "numeric"
            If Trim(pstrToken) = "" Then
              pstrToken = "0"
            Else
              dblTemp = CDbl(datGeneral.ConvertNumberForSQL(pstrToken))
              pstrToken = CStr(dblTemp)

              'MH20041111 Fault 9070
              'Check if this is a numeric value and do not allow dates.
              'If CStr(Val(pstrToken)) <> Replace(Trim(Replace(pstrToken, "0", " ")), " ", "0") Then
              If Not CheckIfNumeric(pstrToken) Then
                RecordError "Error converting '" & pstrToken & "' to numeric type value.", False
                Exit Function
              End If

            End If
          
          ElseIf mvarColDetails(9, pintCounter) Then
            strDataType = "logic"
            sTemp = Trim(UCase(Left(pstrToken, 1)))
            pstrToken = IIf((sTemp = "F") Or (sTemp = "N") Or (sTemp = "0") Or (sTemp = ""), "0", "1")
          Else
            
'            ' Trimming (if any)
'            Select Case mvarColDetails(10, pintToken)
'              Case giTRIMMING_LEFTRIGHT
'                pstrToken = Trim(pstrToken)
'              Case giTRIMMING_LEFTONLY
'                pstrToken = LTrim(pstrToken)
'              Case giTRIMMING_RIGHTONLY
'                pstrToken = RTrim(pstrToken)
'            End Select
              
            pstrToken = "'" & Replace(pstrToken, "'", "''") & "'"
          End If

          GetSpecificKeyData = pstrToken
          Exit Function
        End If

        ' Add the last columns length to the current position marker
        pintCurrentPos = pintCurrentPos + mvarColDetails(6, pintCounter)

      Next pintCounter
       
  End Select

  Exit Function

ErrTrap:
  If Err.Number = 13 And strDataType <> vbNullString Then  'Type Mismatch
    RecordError "Error converting '" & pstrToken & "' to " & strDataType & " type value.", False
  End If
  GetSpecificKeyData = ""

End Function

Private Function ConvertExcelFile() As Boolean

  ' Purpose : This function converts an Excel file into a CSV file and sets a flag
  '           which is used to delete the temp file on import completion
  ' Input   : None
  ' Output  : True/False.

  Dim pxlsSheet As Excel.Worksheet
  Dim pstrNewTempName As String
  Dim lngColumn As Long
  Dim lngMaxCol As Long
  Dim lngMaxRow As Long
  
  
  On Error GoTo ErrTrap

  'MH20020402 Moved from Class_Initialise
  Set mobjExcelApp = New Excel.Application
  
  With mobjExcelApp

    ' Open the Excel spreadsheet
    .Workbooks.Open mstrImportFileName, , True

    ' Set reference to the active sheet
    Set pxlsSheet = .ActiveSheet
    
    
    
    '16/08/2001 MH Fault 2263
    'Loop through columns forcing date columns to correct format
    With pxlsSheet
      With .Cells.SpecialCells(xlCellTypeLastCell)
        lngMaxCol = .Column
        lngMaxRow = .Row
      End With

      'For lngColumn = 1 To lngMaxCol
      For lngColumn = 1 To UBound(mvarColDetails, 2)
        If mvarColDetails(7, lngColumn) Then  'Date Column
          .Range(.Cells(1, lngColumn), .Cells(lngMaxRow, lngColumn)).Select
          mobjExcelApp.Selection.NumberFormat = DateFormat
        End If
      Next
    End With
    
    ' Obtain unique temp filename
    pstrNewTempName = GetTmpFName
    If Dir(pstrNewTempName) <> vbNullString Then
      Kill pstrNewTempName
    End If
    
    ' Save the sheet as a comma delimited txt file
    'AE20071009 Fault #10131 - Local:=True only seems to work on the Workbook, not WorkSheet??
    pxlsSheet.SaveAs pstrNewTempName, xlCSV
    '.ActiveWorkbook.SaveAs pstrNewTempName, xlCSV, Local:=True

    ' Close the workbook, not saving any changes
    .ActiveWorkbook.Close False
  
    ' Quit the excel app object
    mobjExcelApp.Quit
    
  End With

  ' Alter the import definition variables
  mstrImportFileName = pstrNewTempName
  mstrImportDelimiter = ","
  mstrImportEncapsulator = Chr(34)
  mintImportFileType = 0
  
  mblnExcelFileHasBeenConverted = True

  ConvertExcelFile = True
  Exit Function
  
ErrTrap:
  
  ConvertExcelFile = False
  If Err.Number = 429 Then
    mstrErrorString = "Microsoft Excel not installed on this machine"
  Else
    mstrErrorString = "Error whilst converting Excel file to csv." & vbCrLf & Err.Description
  End If

End Function

'
'
' RH 08/09/00 - Note the original DOINSERT function is commented out below this
'               amended one which uses new deadlocking code.
'
Private Function DoInsert() As Boolean

  ' Purpose : This function attempts to insert a row into the database
  '           using the data from the import file
  ' Input   : None
  ' Output  : True/False.

  On Error GoTo ErrTrap
  
  Dim fSavedOK As Boolean
  Dim mstrParent1Where As String
  Dim pintCounter As Integer
  Dim pblnAlreadyRolledBack As Boolean
  Dim plngRecID As Long
  Dim lngIndex As Long
  Dim strValues As String

  fSavedOK = False

  If mblnCanInsert Then

    ' Begin transaction
    gADOCon.BeginTrans
    pblnAlreadyRolledBack = False

    strValues = GetDataToImport(mintImportFileType)
    If strValues <> vbNullString Then
      For lngIndex = 0 To UBound(mstrBaseTableRealSources)
        mstrSQL = "INSERT INTO " & mstrBaseTableRealSources(lngIndex) & " " & _
                            "(" & GetColumnList(mlngImportBaseTable) & ") " & _
                            "VALUES (" & strValues & ")"
        fSavedOK = mclsDeadlock.InsertTableRecord(mstrSQL, mlngImportBaseTable, plngRecID)
        If fSavedOK Then
          Exit For
        End If
      Next
    End If

    If fSavedOK Then
      mlngSuccessRecords = mlngSuccessRecords + 1
      gADOCon.CommitTrans
       
      'JDM - 13/12/01 - Fault 3280 - Log successful records
      If mbLoggingImportSuccess Then
        gobjEventLog.AddDetailEntry GetRecordDesc(plngRecID) & " Inserted successfully"
      End If

    Else

      If Not pblnAlreadyRolledBack Then gADOCon.RollbackTrans
      RecordError mclsDeadlock.ErrorString
      If mclsDeadlock.FatalError Then
        DoInsert = fSavedOK
        Exit Function
      End If

    End If

  Else
    RecordError "You do not have insert privileges on the '" & mstrImportBaseTableName & "' table."

  End If

  DoInsert = fSavedOK
  Exit Function
  
ErrTrap:
  gADOCon.RollbackTrans
  pblnAlreadyRolledBack = True
  mstrErrorString = "Error inserting the following row : " & vbCrLf & vbCrLf & mstrBufferString & vbCrLf & vbCrLf & Err.Description
  DoInsert = False
  Resume Next
  
End Function

Private Sub GetFileLineCount(pintFileType As ImportType)

  Select Case pintFileType

    Case ImportType.DelimitedFile, ImportType.FixedLengthFile, ImportType.SQLTable

      'MH20071003
      'mlngTotalRecords = FileLineCount(mstrImportFileName) - IIf(mblnImportIgnoreFirstLine = True, 1, 0)
      'mlngTotalRecords = mlngTotalRecords - IIf(mblnImportIgnoreLastLine = True, 1, 0)
      
      'NPG20090713 Fault HRPRO-39
      'mlngTotalRecords = FileLineCount(mstrImportFileName)
      mlngTotalRecords = ProcessUploadFile(mstrImportFileName)
      mlngTotalRecords = mlngTotalRecords - (mintImportHeaderLines + mintImportFooterLines)

    Case ImportType.ExcelWorksheet

      With mobjExcelApp
        .Workbooks.Open mstrImportFileName, , True
        .ActiveCell.SpecialCells(xlLastCell).Select

        'MH20071003
        'mlngTotalRecords = .ActiveCell.Row - IIf(mblnImportIgnoreFirstLine = True, 1, 0)
        'mlngTotalRecords = mlngTotalRecords - IIf(mblnImportIgnoreLastLine = True, 1, 0)
        mlngTotalRecords = .ActiveCell.Row
        mlngTotalRecords = mlngTotalRecords - (mintImportHeaderLines + mintImportFooterLines)

        .ActiveWorkbook.Close
      End With
      
  End Select
  
  If mlngTotalRecords >= 0 Then
    If Not gblnBatchMode Then
      gobjProgress.Bar1MaxValue = mlngTotalRecords
    Else
      gobjProgress.Bar2MaxValue = mlngTotalRecords
    End If
  End If

End Sub

Private Function GetColumnList(plngTableID As Long) As String

  ' Purpose : This function returns a string containing the columns from the
  '           specified table
  ' Input   : Table ID
  ' Output  : String containing column names.

  On Error GoTo ErrTrap
  
  If Len(mstrColumnList) > 0 Then
    GetColumnList = mstrColumnList
    Exit Function
  End If
    
  Dim pintCounter As Integer
  
  For pintCounter = 1 To UBound(mvarColDetails, 2)
    If mvarColDetails(1, pintCounter) = plngTableID Then
      GetColumnList = GetColumnList & IIf(Len(GetColumnList) > 0, ",", "") & mvarColDetails(4, pintCounter)
    End If
  Next pintCounter
  
  '#############
  ' Get the parent table id (relationship) columns if required
  If mlngImportParent2Table > 0 Then GetColumnList = "ID_" & mlngImportParent2Table & "," & GetColumnList
  If mlngImportParent1Table > 0 Then GetColumnList = "ID_" & mlngImportParent1Table & "," & GetColumnList
  '#############
  
  mstrColumnList = GetColumnList
  
  Exit Function
  
ErrTrap:
  
  GetColumnList = ""
  mstrErrorString = "Error whilst getting the column list from the table." & vbCrLf & vbCrLf & "(" & Err.Description & ")"
  
End Function

Private Function GetDataToImport(pintFileType As ImportType) As String

  ' Purpose : This function retrieves data from the import file
  ' Input   : File Type (Delimited or Fixed Length).
  ' Output  : String containing the data from the import file.

  Dim pintCounter As Integer
  Dim pstrDelimiter As String
  Dim pstrToken As String
  Dim pstrParentWhereSQL As String
  Dim prstParentFind As Recordset
  Dim sTemp As String
  Dim strSQL As String
  Dim lngParentID As Long
  
  Dim pstrType As String
  Dim pintTableID As Long
  Dim pintColumnID As Long
  Dim lngIndex As Long
  Dim objTableView As CTablePrivilege

  On Error GoTo ErrTrap
  
  ' Read the data depending on the file type
  Select Case pintFileType
  
    Case ImportType.DelimitedFile, ImportType.SQLTable:
        
      ' Check delimiter. Ie, if Tab selected, replace "<Tab>" with a vbtab
      Select Case LCase(mstrImportDelimiter)
        Case ",": pstrDelimiter = mstrImportDelimiter
        Case "<tab>": pstrDelimiter = vbTab
        Case Else: pstrDelimiter = mstrImportOtherDelimiter
      End Select
        
      For pintCounter = 1 To UBound(mvarColDetails, 2)
        'TM - Set variables to pass to the NumberColumn generic function.
        pstrType = mvarColDetails(0, pintCounter)
        pintTableID = mvarColDetails(1, pintCounter)
        pintColumnID = mvarColDetails(3, pintCounter)
  

        ' RH 11/07/00
        ' Mike - prev fix causes certain imports to fail.  Removed fix.
        
        'If mvarColDetails(0, pintCounter) = "C" And mvarColDetails(5, pintCounter) = False Then
        'If mvarColDetails(0, pintCounter) = "C" And ((mvarColDetails(5, pintCounter) = False) Or _
                                                     (mvarColDetails(1, pintCounter) = mlngImportBaseTable)) Then
        If mvarColDetails(0, pintCounter) = "C" And mvarColDetails(1, pintCounter) = mlngImportBaseTable Then
          
          pstrToken = ParseStr(mstrBufferString, pintCounter, pstrDelimiter, mstrImportEncapsulator)
           
          If mvarColDetails(7, pintCounter) Then
            pstrToken = GetDateInSQLFormat(pstrToken)

          ElseIf mvarColDetails(8, pintCounter) Then
            If Trim(pstrToken) = "" Then
              pstrToken = "0"
            Else
              pstrToken = datGeneral.ConvertNumberForSQL(pstrToken)
          
              'MH20041111 Fault 9070
              'Check if this is a numeric value and do not allow dates.
              'If CStr(Val(pstrToken)) <> Replace(Trim(Replace(pstrToken, "0", " ")), " ", "0") Then
              If Not CheckIfNumeric(pstrToken) Then
                RecordError "Error converting '" & pstrToken & "' to numeric type value.", False
                Exit Function
              End If
            
            End If
          
          ElseIf mvarColDetails(9, pintCounter) Then
            sTemp = Trim(UCase(Left(pstrToken, 1)))
            pstrToken = IIf((sTemp = "F") Or (sTemp = "N") Or (sTemp = "0") Or (sTemp = ""), "0", "1")
          Else
            
            ' Trimming (if any)
            Select Case mvarColDetails(10, pintCounter)
              Case giTRIMMING_LEFTRIGHT
                pstrToken = Trim(pstrToken)
              Case giTRIMMING_LEFTONLY
                pstrToken = LTrim(pstrToken)
              Case giTRIMMING_RIGHTONLY
                pstrToken = RTrim(pstrToken)
            End Select
          
            pstrToken = "'" & Replace(pstrToken, "'", "''") & "'"
          End If

          AddToLookupArray pintCounter

          GetDataToImport = GetDataToImport & IIf(Len(GetDataToImport) > 0, ",", "") & pstrToken
        End If
      Next pintCounter
    
    Case ImportType.FixedLengthFile
    
      Dim pintCurrentPos As Integer
      pintCurrentPos = 1
      For pintCounter = 1 To UBound(mvarColDetails, 2)
      
        'TM - Set variables to pass to the NumberColumn generic function.
        pstrType = mvarColDetails(0, pintCounter)
        pintTableID = mvarColDetails(1, pintCounter)
        pintColumnID = mvarColDetails(3, pintCounter)

        ' RH 27/03/01 - Fixed length included key fields, doh !
        If mvarColDetails(0, pintCounter) = "C" And ((mvarColDetails(5, pintCounter) = False) Or _
                                                     (mvarColDetails(1, pintCounter) = mlngImportBaseTable)) Then
      
          ' If column, get the data, else its a filler so get a blank string
          If mvarColDetails(0, pintCounter) = "C" Then
            'pstrToken = Trim(Mid(mstrBufferString, pintCurrentPos, mvarColDetails(6, pintCounter)))
            'NHRD15122004 Fault 9598
            pstrToken = RTrim(Mid(mstrBufferString, pintCurrentPos, mvarColDetails(6, pintCounter)))
          Else
            pstrToken = Space(mvarColDetails(6, pintCounter))
          End If
            
          If mvarColDetails(7, pintCounter) Then
            pstrToken = GetDateInSQLFormat(pstrToken)

          ElseIf mvarColDetails(8, pintCounter) Then
            If Trim(pstrToken) = "" Then
              pstrToken = "0"
            Else
              pstrToken = datGeneral.ConvertNumberForSQL(pstrToken)

              'MH20041111 Fault 9070
              'Check if this is a numeric value and do not allow dates.
              'If CStr(Val(pstrToken)) <> Replace(Trim(Replace(pstrToken, "0", " ")), " ", "0") Then
              If Not CheckIfNumeric(pstrToken) Then
                RecordError "Error converting '" & pstrToken & "' to numeric type value.", False
                Exit Function
              End If

            End If
          
          ElseIf mvarColDetails(9, pintCounter) Then
            sTemp = Trim(UCase(Left(pstrToken, 1)))
            pstrToken = IIf((sTemp = "F") Or (sTemp = "N") Or (sTemp = "0") Or (sTemp = ""), "0", "1")
          Else
            
            ' Trimming (if any)
            Select Case mvarColDetails(10, pintCounter)
              Case giTRIMMING_LEFTRIGHT
                pstrToken = Trim(pstrToken)
              Case giTRIMMING_LEFTONLY
                pstrToken = LTrim(pstrToken)
              Case giTRIMMING_RIGHTONLY
                pstrToken = RTrim(pstrToken)
            End Select
            
            pstrToken = "'" & Replace(pstrToken, "'", "''") & "'"
          End If
            
            
          AddToLookupArray pintCounter



          ' Add token to the whole data string
          GetDataToImport = GetDataToImport & IIf(Len(GetDataToImport) > 0, ",", "") & pstrToken
        End If
        
        ' Add the last columns length to the current position marker
        pintCurrentPos = pintCurrentPos + mvarColDetails(6, pintCounter)
       
      Next pintCounter
       
  End Select

'###############
  ' Do the second parent table if required
  If mlngImportParent2Table > 0 Then
  
    ' JPD 9/3/2001
    pstrParentWhereSQL = ""
    
    For pintCounter = 1 To UBound(mvarColDetails, 2)
      If mvarColDetails(1, pintCounter) = mlngImportParent2Table Then
        'If mvarColDetails(5, pintCounter) = True Then
          pstrParentWhereSQL = pstrParentWhereSQL & _
                IIf(Len(pstrParentWhereSQL) > 0, " AND ", " WHERE ")
          
          If UBound(mstrParent2TableRealSources) = 0 Then
            pstrParentWhereSQL = pstrParentWhereSQL & _
                  mstrParent2TableRealSources(0) & "." & mvarColDetails(4, pintCounter)
          Else
            pstrParentWhereSQL = pstrParentWhereSQL & _
                  "CASE"
            For lngIndex = 0 To UBound(mstrParent2TableRealSources)
              pstrParentWhereSQL = pstrParentWhereSQL & vbCrLf & _
                  " WHEN NOT " & mstrParent2TableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter) & _
                  " IS NULL THEN " & mstrParent2TableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter)
            Next
            pstrParentWhereSQL = pstrParentWhereSQL & _
                  " ELSE null END"
          End If


          'pstrParentWhereSQL = pstrParentWhereSQL & " = " & GetSpecificKeyData(pintCounter)
          sTemp = GetSpecificKeyData(pintCounter)
          If mstrErrorString <> vbNullString Then
            GetDataToImport = vbNullString
            Exit Function
          End If

          'MH20050124 Fault 9592 - If null date then use "IS NULL" and not "= NULL"
          If mvarColDetails(7, pintCounter) = True And sTemp = "null" Then
            pstrParentWhereSQL = pstrParentWhereSQL & " is null"
          Else
            pstrParentWhereSQL = pstrParentWhereSQL & " = " & sTemp
          End If

        'End If
      End If
    Next pintCounter

    ' JPD20011022 Faults 2996 & 2997 Disable Lofty's error handling.
    gobjErrorStack.Disable
    
    'JPD 20040319 Fault 8306
    Set objTableView = gcoTablePrivileges.item(mstrImportParent2TableName)
    If objTableView.TableType = tabChild Then
      strSQL = "SELECT " & mstrParent2TableRealSources(0) & ".ID FROM " & mstrParent2TableRealSources(0) & _
        mstrParent2ViewJoins & " " & pstrParentWhereSQL
    Else
      strSQL = "SELECT " & mstrImportParent2TableName & ".ID FROM " & mstrImportParent2TableName & _
        mstrParent2ViewJoins & " " & pstrParentWhereSQL
    End If
    Set objTableView = Nothing
    
    Set prstParentFind = datGeneral.GetReadOnlyRecords(strSQL)
    
''{'MH20040309 Fault 8238
'    If Not (prstParentFind.BOF Or prstParentFind.EOF) Then
'      If InStr(prstParentFind.Fields(0).Value, GetDataToImport) = 0 Then
'        GetDataToImport = prstParentFind.Fields(0).Value & _
'            IIf(GetDataToImport <> vbNullString, ",", "") & _
'            GetDataToImport
'      End If
'    End If
'    prstParentFind.Close
'    Set prstParentFind = Nothing
'    gobjErrorStack.Enable
'
'    ' JPD 9/3/2001
'    If Not (prstParentFind.BOF Or prstParentFind.EOF) Then
'      GetDataToImport = prstParentFind.Fields(0) & "," & GetDataToImport
'    Else
'      GetDataToImport = "null, " & GetDataToImport
'    End If
    lngParentID = 0
    If Not (prstParentFind.BOF Or prstParentFind.EOF) Then
      lngParentID = prstParentFind.Fields(0).Value
    End If
    
    GetDataToImport = CStr(lngParentID) & _
        IIf(GetDataToImport <> vbNullString, ",", "") & _
        GetDataToImport

    prstParentFind.Close
    Set prstParentFind = Nothing
    gobjErrorStack.enable
'}
    
    
  End If
'###############
  ' Do the first parent table if required
  If (mlngImportParent1Table > 0) Then
    
    ' JPD 9/3/2001
    pstrParentWhereSQL = ""
    
    For pintCounter = 1 To UBound(mvarColDetails, 2)
      If mvarColDetails(1, pintCounter) = mlngImportParent1Table Then
        'If mvarColDetails(5, pintCounter) = True Then
          pstrParentWhereSQL = pstrParentWhereSQL & IIf(Len(pstrParentWhereSQL) > 0, " AND ", " WHERE ")

          If UBound(mstrParent1TableRealSources) = 0 Then
            pstrParentWhereSQL = pstrParentWhereSQL & _
                  mstrParent1TableRealSources(0) & "." & mvarColDetails(4, pintCounter)
          Else
            pstrParentWhereSQL = pstrParentWhereSQL & _
                  "CASE"
            For lngIndex = 0 To UBound(mstrParent1TableRealSources)
              pstrParentWhereSQL = pstrParentWhereSQL & _
                  " WHEN NOT " & mstrParent1TableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter) & _
                  " IS NULL THEN " & mstrParent1TableRealSources(lngIndex) & "." & mvarColDetails(4, pintCounter)
            Next
            pstrParentWhereSQL = pstrParentWhereSQL & _
                  " ELSE null END"
          End If


          'pstrParentWhereSQL = pstrParentWhereSQL & " = " & GetSpecificKeyData(pintCounter)
          sTemp = GetSpecificKeyData(pintCounter)
          If mstrErrorString <> vbNullString Then
            GetDataToImport = vbNullString
            Exit Function
          End If

          'MH20050124 Fault 9592 - If null date then use "IS NULL" and not "= NULL"
          If mvarColDetails(7, pintCounter) = True And sTemp = "null" Then
            pstrParentWhereSQL = pstrParentWhereSQL & " is null"
          Else
            pstrParentWhereSQL = pstrParentWhereSQL & " = " & sTemp
          End If


        'End If
      End If
    Next pintCounter
    
    ' JPD20011022 Faults 2996 & 2997 Disable Lofty's error handling.
    gobjErrorStack.Disable
    
    'JPD 20040319 Fault 8306
    Set objTableView = gcoTablePrivileges.item(mstrImportParent1TableName)
    If objTableView.TableType = tabChild Then
      strSQL = "SELECT " & mstrParent1TableRealSources(0) & ".ID FROM " & mstrParent1TableRealSources(0) & _
        mstrParent1ViewJoins & " " & pstrParentWhereSQL
    Else
      strSQL = "SELECT " & mstrImportParent1TableName & ".ID FROM " & mstrImportParent1TableName & _
        mstrParent1ViewJoins & " " & pstrParentWhereSQL
    End If
    Set objTableView = Nothing
    
    Set prstParentFind = datGeneral.GetReadOnlyRecords(strSQL)
''{'MH20040309 Fault 8238
'    If Not (prstParentFind.BOF Or prstParentFind.EOF) Then
'      If InStr(prstParentFind.Fields(0).Value, GetDataToImport) = 0 Then
'        GetDataToImport = prstParentFind.Fields(0).Value & _
'            IIf(GetDataToImport <> vbNullString, ",", "") & _
'            GetDataToImport
'      End If
'
'    'MH20040108 Fault 7896
'    Else
'      GetDataToImport = "0" & _
'            IIf(GetDataToImport <> vbNullString, ",", "") & _
'            GetDataToImport
'    End If
    lngParentID = 0
    If Not (prstParentFind.BOF Or prstParentFind.EOF) Then
      lngParentID = prstParentFind.Fields(0).Value
    End If
    
    GetDataToImport = CStr(lngParentID) & _
        IIf(GetDataToImport <> vbNullString, ",", "") & _
        GetDataToImport
'}

    prstParentFind.Close
    Set prstParentFind = Nothing
    gobjErrorStack.enable

  End If


'##############
  Exit Function

ErrTrap:

  GetDataToImport = ""

End Function

Public Function ParseStr(ByVal strWork As String, _
                               intTokenNum As Integer, _
                               strDelimitChr As String, _
                               Optional ByVal strEncapChr As String) As String

  On Local Error Resume Next
  
  
  'Need to handle things like this:
  '
  '  Clipper,95,05/08/2006,10/08/2006,"3"" nail, through foot",SICK
  '
  ' e.g. Field 4 should return:  3" nail, through foot
  '
  
  
  Dim blnExitDo As Boolean
  Dim intDPos As Integer
  Dim intSPtr As Integer
  Dim intEPtr As Integer
  Dim intCurrentTokenNum As Integer
  Dim intWorkStrLen As Integer
  Dim intEncapStatus As Integer
  Dim strTemp As String
  Dim blnLoop As Boolean
  
  Static intSPos As Integer
  Static intDelimitLen As Integer
  
  intWorkStrLen = Len(strWork)
  
  If Len(strEncapChr) Then
    intEncapStatus = Len(strEncapChr)
  End If

  If intWorkStrLen = 0 Or (intSPos > intWorkStrLen And intTokenNum = 0) Then
    intSPos = 0
    Exit Function
  ElseIf intTokenNum > 0 Or intSPos = 0 Then
    intSPos = 1
    intDelimitLen = Len(strDelimitChr)
  End If
  
  Do
  
    intDPos = InStr(intSPos, strWork, strDelimitChr)
    
    If intEncapStatus Then
      intSPtr = InStr(intSPos, strWork, strEncapChr)
      
      
      'MH20060612 Fault 11227
      'intEPtr = InStr(intSPtr + 1, strWork, strEncapChr)
      intEPtr = intSPtr + 1
      Do
        intEPtr = InStr(intEPtr, strWork, strEncapChr)
        
        blnLoop = (intEPtr = InStr(intEPtr, strWork, strEncapChr & strEncapChr))
        If blnLoop Then
          intEPtr = intEPtr + Len(strEncapChr & strEncapChr)
        End If
      Loop While blnLoop
      
      
      If intDPos > intSPtr And intDPos < intEPtr Then
        intDPos = InStr(intEPtr, strWork, strDelimitChr)
      End If
    End If
    
    If intDPos < intSPos Then
      intDPos = intWorkStrLen + intDelimitLen
    End If
    
    If intDPos Then
      If intTokenNum Then
        intCurrentTokenNum = intCurrentTokenNum + 1
        If intCurrentTokenNum = intTokenNum Then
          strTemp = Mid(strWork, intSPos, intDPos - intSPos)
          blnExitDo = True
        Else
          blnExitDo = False
        End If
      Else
        strTemp = Mid(strWork, intSPos, intDPos - intSPos)
        blnExitDo = True
      End If
      intSPos = intDPos + intDelimitLen
    Else
      intSPos = 0
      blnExitDo = True
    End If
  
  Loop Until blnExitDo
  

'MH20060612 Fault 11227
'  If intEncapStatus Then
'    ParseStr = ReplaceCS(strTemp, strEncapChr, "", StringBinaryCompare)
'  Else
'    ParseStr = strTemp
'  End If

  ParseStr = strTemp
  If intEncapStatus Then
    If Left(strTemp, 1) = strEncapChr Then
      strTemp = Mid(strTemp, 2)
    End If
    If Right(strTemp, 1) = strEncapChr Then
      strTemp = Left(strTemp, Len(strTemp) - 1)
    End If
    ParseStr = Replace(strTemp, strEncapChr & strEncapChr, strEncapChr)
  End If
  
  
End Function

Function ProcessUploadFile(strFileName As String) As Long
  
  'NPG20090713 Fault HRPRO-39
  ' This function counts the number of lines in a fixed-length/delimited file and
  ' replaces the FileLineCount function, which doesn't cater for multi-line data sources.
  
  On Local Error GoTo FileLineCount_ERROR
  
  Dim fso, objFile, iLineCount
  Dim sInLine, sOutLine
  'Dim strEOLChars As String
  
  mstrEOLChars = FileEOLChars(strFileName)
  If mstrEOLChars = vbNullString Then
    mstrErrorString = "Error identifying end of line characters." & vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"
    ProcessUploadFile = -1
    Exit Function
  End If
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set objFile = fso.OpenTextFile(strFileName, 1)
    
  iLineCount = 0
  Do Until objFile.AtEndOfStream
    sOutLine = ""
    Do
      sInLine = objFile.ReadLine
      sOutLine = sOutLine & sInLine
    Loop While (UBound(Split(sOutLine, mstrImportEncapsulator)) Mod 2 <> 0) And (Not objFile.AtEndOfStream)
      iLineCount = iLineCount + 1
  Loop
  
  objFile.Close
    
  ProcessUploadFile = iLineCount
  
  Exit Function
  
FileLineCount_ERROR:
  mstrErrorString = "Error counting number of lines in import file." & vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"
  ProcessUploadFile -1
  
End Function


'NPG20090713 Fault HRPRO-39 - replaced by function ProcessUploadFile
'Private Function FileLineCount(strFileName As String) As Long
'
'  'On Local Error Resume Next
'  On Local Error GoTo FileLineCount_ERROR
'
'  Dim intN As Long
'  Dim intSPos As Long
'  Dim intCPos As Long
'  Dim fh As Integer
'  Dim strBuffer As String
'  Dim lngBlocks As Long
'  Dim lngBlock As Long
'  Dim lngLastBlockSize As Long
'  Dim strEOLChars As String
'  Dim intEOLCharsLen As Integer
'  Dim lngFilePtr As Long
'  Dim lngFileSize As Long
'
'  Const kintBlockSize As Integer = 31477
'
'  strEOLChars = FileEOLChars(strFileName)
'  If strEOLChars = vbNullString Then
'    mstrErrorString = "Error identifying end of line characters." & vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"
'    FileLineCount = -1
'    Exit Function
'  End If
'
'  intEOLCharsLen = Len(strEOLChars)
'  strBuffer = Space(kintBlockSize)
'
'  fh = FreeFile
'  ' JPD20011012 Fault 2948 Open file as binary to avoid erors when the file
'  ' contains dodgy characters.
'  'Open strFilename For Input As #fh
'  Open strFileName For Binary Access Read As #fh
'
'  lngFileSize = LOF(fh)
'  lngFilePtr = 1
'
'  ' Determine the number of blocks and size of last (short) block
'  lngBlocks = (lngFileSize \ kintBlockSize) + 1
'
'  ' Read each block
'
'  For lngBlock = 1 To lngBlocks
'    If lngBlock = lngBlocks Then
'      lngLastBlockSize = lngFileSize - Seek(fh) + 1
'      strBuffer = Input(lngLastBlockSize, fh)
'    Else
'      strBuffer = Input(kintBlockSize, fh)
'    End If
'
'    DoEvents
'    intSPos = 1
'    Do
'      intCPos = InStr(intSPos, strBuffer, strEOLChars)
'      If intCPos Then
'        intN = intN + 1
'        intSPos = intCPos + intEOLCharsLen
'      Else
'        lngFilePtr = lngFilePtr + intSPos - 1
'        Seek #fh, lngFilePtr
'      End If
'    Loop While intCPos
'  Next
'
'  Close #fh
'
'  'Add one to count if last line does not end with a crlf
'  If Right(strBuffer, intEOLCharsLen) <> strEOLChars Then intN = intN + 1
'
'  FileLineCount = intN
'
'  Exit Function
'
'FileLineCount_ERROR:
'  mstrErrorString = "Error counting number of lines in import file." & vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")"
'  FileLineCount = -1
'
'End Function

Private Function FileEOLChars(strFile As String) As String

  On Local Error Resume Next
  
  Dim fh As Integer
  Dim lngBufSize As Long
  Dim strLineBuf As String
  
  fh = FreeFile
  
  ' JPD20011012 Fault 2948 Open as a binary file to avoid errors when the file has dodgy
  ' characters in it.
  'Open strFile For Input As #fh
  Open strFile For Binary Access Read As #fh
  
  lngBufSize = LOF(fh)
  
  'MH20030226 Fault 5081
  'If lngBufSize > 1024 Then
  '  strLineBuf = Input(1024, #fh)
  If lngBufSize > 2048 Then
    strLineBuf = Input(2048, #fh)
  Else
    strLineBuf = Input(lngBufSize, #fh)
  End If
  
  If InStr(strLineBuf, vbCrLf) Then
    FileEOLChars = vbCrLf
  ElseIf InStr(strLineBuf, vbCr) Then
    FileEOLChars = vbCr
  ElseIf InStr(strLineBuf, vbLf) Then
    FileEOLChars = vbLf
  Else
    'FileEOLChars = vbNullString
    FileEOLChars = vbCrLf
  End If
  
  Close #fh

End Function

Private Sub ClearUp()

  On Error Resume Next
  
  If mblnExcelFileHasBeenConverted Then
    Kill mstrImportFileName
  End If

''  mlngImportID = 0
''  mblnUserCancelled = False
''  mstrImportName = vbNullString
''  mstrImportDescription = vbNullString
''  mlngImportBaseTable = 0
''  mstrImportBaseTableName = vbNullString
''  mintImportFileType = 0
''  mstrImportFileName = vbNullString
''  mstrImportDelimiter = vbNullString
''  mstrImportOtherDelimiter = vbNullString
''  mstrImportEncapsulator = vbNullString
''  mblnImportMultipleRecordAction = False
''  mblnImportIgnoreFirstLine = False
''  mblnImportCreateNewOnly = False
''
''  'MH20010816 Fault 2017
''  mlngImportType = 0
''
''  mlngTotalRecords = 0
''  mlngSuccessRecords = 0
''  mlngFailRecords = 0
''  mintFileNo = 0
''  mstrSQL = vbNullString
''  mstrSQLWhere = vbNullString
''  mstrBufferString = vbNullString
''  'mstrBaseTableRealSource = vbNullString
''  mlngImportParent1Table = 0
''  'mstrImportParent1TableName = vbNullString
''  mlngImportParent2Table = 0
''  'mstrImportParent2TableName = vbNullString
''  ReDim mvarColDetails(12, 0)
''  If gblnBatchMode Then gobjProgress.ResetBar2
''  If Not gblnBatchMode Then mstrErrorString = ""
''
''  'MH20021031 Fault 4744
''  'Eeeekkk... this caused a messagebox on the screen if Import failed.
''  'gblnBatchMode = False
''  mstrColumnList = ""
''  Screen.MousePointer = vbDefault

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

  
  If Trim$(GetRecordDesc) = vbNullString Then
    GetRecordDesc = "Record Description Undefined"
  End If

Exit Function

LocalErr:
  GetRecordDesc = ""

End Function


Private Function AddToLookupArray(intCount As Integer)

  Dim strTable As String
  Dim strColumn As String
  Dim strValue As String
  Dim lngIndex As Long
  
  strTable = mvarColDetails(11, intCount)
  strColumn = mvarColDetails(12, intCount)
  
  If strTable <> vbNullString And _
     strColumn <> vbNullString Then
    
    strValue = GetSpecificKeyData(intCount)
    If mstrErrorString = vbNullString Then
    
      lngIndex = UBound(mstrLookupArray) + 1
      ReDim Preserve mstrLookupArray(lngIndex)
      mstrLookupArray(lngIndex) = _
          "IF not exists(" & _
          "SELECT * FROM [" & strTable & "] " & _
          "WHERE [" & strColumn & "] = " & strValue & ") " & _
          "  INSERT [" & strTable & "] ([" & strColumn & "]) " & _
          "  VALUES (" & strValue & ")"
    End If
  End If

End Function


Private Function GetLookupInsert(intCount As Integer, lngColumnID As Long) As Boolean

  Dim rsTemp As ADODB.Recordset
  Dim strSQL As String

  'This assumes that the user has full access to the lookup table and that
  'the lookup column is a character column.  If these rules change then this
  '
  
  strSQL = "SELECT tab.tablename, col.columnname " & _
           "FROM ASRSysColumns " & _
           "JOIN ASRSysColumns col ON col.columnid = ASRSysColumns.LookupColumnID " & _
           "JOIN ASRSysTables tab ON tab.tableid = ASRSysColumns.LookupTableID " & _
           "WHERE ASRSysColumns.ColumnID = " & CStr(lngColumnID)
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
  
  GetLookupInsert = (Not rsTemp.BOF And Not rsTemp.EOF)
  If GetLookupInsert Then
    mvarColDetails(11, intCount) = IIf(IsNull(rsTemp!TableName), vbNullString, rsTemp!TableName)
    mvarColDetails(12, intCount) = IIf(IsNull(rsTemp!ColumnName), vbNullString, rsTemp!ColumnName)
  End If
    
  rsTemp.Close
  Set rsTemp = Nothing

End Function


Private Function CheckIfInsertIntoLookups()

  Dim strErrorMessage As String
  Dim lngCount As Long

  For lngCount = 1 To UBound(mstrLookupArray)
    datGeneral.ExecuteSql mstrLookupArray(lngCount), strErrorMessage
'MH25102004: Don't EXIT FUNCTION!  Need to try all lookup entries even if one fails.
'            Also, this error recording didn't work as strErrorMessage is overwritten.
'            Spoke to JED and he suggested leaving it out anyway.
'    If strErrorMessage <> vbNullString Then
'      mstrErrorString = "Error Populating Lookup table" & vbCrLf & _
'          strErrorMessage & vbCrLf & _
'          mstrLookupArray(lngCount)
'      'Exit Function
'    End If
  Next

End Function


Private Function GetDateInSQLFormat(strDate As String) As String

  Dim strDateParts() As String
  Dim strTempSeparator As String
  Dim strOutput As String
  Dim lngOffset As Long
  
  Dim strYear As String
  Dim strMonth As String
  Dim strDay As String
  
  
  On Error GoTo LocalErr

  'If there is no separator then guess!!!!
  If mstrDateSeparator = vbNullString Then
    If InStr(strDate, "/") > 0 Then
      strTempSeparator = "/"
    ElseIf InStr(strDate, ".") > 0 Then
      strTempSeparator = "."
    ElseIf InStr(strDate, "-") > 0 Then
      strTempSeparator = "-"
    End If
  Else
    strTempSeparator = mstrDateSeparator
  End If


  'Check for empty date...
  If Trim(Replace(strDate, strTempSeparator, "")) = vbNullString Then
    GetDateInSQLFormat = "null"
    Exit Function
  End If


  'Divide date into array...
  If strTempSeparator <> vbNullString Then
    strDateParts = Split(strDate, strTempSeparator)
    If Len(strDateParts(0)) < 2 Then strDateParts(0) = Right("00" & strDateParts(0), 2)
    If Len(strDateParts(1)) < 2 Then strDateParts(1) = Right("00" & strDateParts(1), 2)
    If Len(strDateParts(2)) < 2 Then strDateParts(2) = Right("00" & strDateParts(2), 2)
  Else
    If Len(strDate) < 6 Then
      GoTo LocalErr
    End If

    ReDim strDateParts(2)
    'Check if there is a four digit year at the begining...
    lngOffset = IIf(Left(mstrDateFormat, 1) = "y" And Len(strDate) > 6, 2, 0)
    strDateParts(0) = Mid(strDate, lngOffset + 1, 2)
    strDateParts(1) = Mid(strDate, lngOffset + 3, 2)
    strDateParts(2) = Mid(strDate, lngOffset + 5)
  End If
  
  
  'Change date into SQL format (i.e. mm/dd/yyyy)
  Select Case mstrDateFormat
  Case "mdy"
    strMonth = strDateParts(0)
    strDay = strDateParts(1)
    strYear = strDateParts(2)
  Case "ymd"
    strYear = strDateParts(0)
    strMonth = strDateParts(1)
    strDay = strDateParts(2)
  Case "ydm"
    strYear = strDateParts(0)
    strDay = strDateParts(1)
    strMonth = strDateParts(2)
  Case Else '"dmy"
    strDay = strDateParts(0)
    strMonth = strDateParts(1)
    strYear = strDateParts(2)
  End Select

  If Val(strDay) < 1 Or Val(strDay) > 31 Then
    RecordError "Error converting '" & strDate & "' to date type value.", False
    GetDateInSQLFormat = vbNullString
    Exit Function
  End If
  
  'MH20041202
  If Val(strMonth) < 1 Or Val(strMonth) > 12 Then
    RecordError "Error converting '" & strDate & "' to date type value.", False
    GetDateInSQLFormat = vbNullString
    Exit Function
  End If
  
  
  'MH20041116 Fault 9477
  'If Len(strYear) <> 2 And Len(strYear) <> 4 Then
  If Len(strYear) <> 2 And (Len(strYear) <> 4 Or Val(strYear) < 1753 Or Val(strYear) > 2999) Then
    RecordError "Error converting '" & strDate & "' to date type value.", False
    GetDateInSQLFormat = vbNullString
    Exit Function
  End If
  
  'Format as "mm/dd/yyyy" for SQL
  strOutput = strMonth & "/" & strDay & "/" & strYear



  'Make sure that we have something like a date
  If Not IsDate(strOutput) Then
    'Err.Raise 13
    RecordError "Error converting '" & strDate & "' to date type value.", False
    GetDateInSQLFormat = vbNullString
    Exit Function
  End If
  
  GetDateInSQLFormat = "'" & strOutput & "'"

Exit Function

LocalErr:
  RecordError "Error converting '" & strDate & "' to date type value.", False
  GetDateInSQLFormat = vbNullString

End Function


Private Function AddToArray(strArray() As String, strInput As String)

  Dim lngIndex As Long

  On Local Error Resume Next

  Err.Clear
  lngIndex = UBound(strArray)
  
  If Err.Number = 9 Then   'Subscript out of range
    Err.Clear
    ReDim strArray(0)
    strArray(0) = strInput
  Else
    ReDim Preserve strArray(lngIndex + 1)
    strArray(lngIndex + 1) = strInput
  End If

End Function


Private Function FindMatchingRecord() As Boolean

  ' Purpose : This function compiles the WHERE clause which is used to
  '           determine if we will be doing an insert or update
  ' Input   : None
  ' Output  : True/False.

  On Error GoTo ErrTrap

  Dim pintCounter As Integer
  Dim pstrBaseWhereSQL As String
  
  Dim plngRecordsReturned As Long
  'Dim blnParentExists As Boolean
  'Dim blnParentRequired As Boolean
  Dim strSQL As String
  Dim lngIndex As Long
  Dim objTableView As CTablePrivilege
  Dim blnParentNotFound As Boolean

  mstrParent1Where = vbNullString
  mstrParent2Where = vbNullString


  For pintCounter = 1 To UBound(mvarColDetails, 2)
    If mvarColDetails(5, pintCounter) = True Then   'Keyed field

      Select Case mvarColDetails(1, pintCounter)  'TableID
      Case mlngImportBaseTable
        AddToWhere pstrBaseWhereSQL, mstrBaseTableRealSources(), pintCounter
      Case mlngImportParent1Table
        AddToWhere mstrParent1Where, mstrParent1TableRealSources(), pintCounter
      Case mlngImportParent2Table
        AddToWhere mstrParent2Where, mstrParent2TableRealSources(), pintCounter
      End Select

      If mstrErrorString <> vbNullString Then
        mlngFailRecords = mlngFailRecords + 1
        FindMatchingRecord = False
        Exit Function
      End If

    End If
  Next
  
  
  If mstrParent1Where <> vbNullString Then
    'JPD 20041130 Fault 8306
    'strSQL = "SELECT " & mstrImportParent1TableName & ".ID " & _
             "FROM " & mstrImportParent1TableName & _
             mstrParent1ViewJoins & " WHERE " & mstrParent1Where
    Set objTableView = gcoTablePrivileges.item(mstrImportParent1TableName)
    If objTableView.TableType = tabChild Then
      strSQL = "SELECT " & mstrParent1TableRealSources(0) & ".ID" & _
        " FROM " & mstrParent1TableRealSources(0) & _
        mstrParent1ViewJoins & _
        " WHERE " & mstrParent1Where
    Else
      strSQL = "SELECT " & mstrImportParent1TableName & ".ID" & _
        " FROM " & mstrImportParent1TableName & _
        mstrParent1ViewJoins & _
        " WHERE " & mstrParent1Where
    End If
    Set objTableView = Nothing
    
    mstrParent1Where = GetRecordIDs(strSQL)
    blnParentNotFound = (mstrParent1Where = vbNullString)
  
  End If
  
  If mstrParent2Where <> vbNullString Then
    'JPD 20041130 Fault 8306
    'strSQL = "SELECT " & mstrImportParent2TableName & ".ID " & _
             "FROM " & mstrImportParent2TableName & _
             mstrParent2ViewJoins & " WHERE " & mstrParent2Where
    Set objTableView = gcoTablePrivileges.item(mstrImportParent2TableName)
    If objTableView.TableType = tabChild Then
      strSQL = "SELECT " & mstrParent2TableRealSources(0) & ".ID" & _
        " FROM " & mstrParent2TableRealSources(0) & _
        mstrParent2ViewJoins & _
        " WHERE " & mstrParent2Where
    Else
      strSQL = "SELECT " & mstrImportParent2TableName & ".ID" & _
        " FROM " & mstrImportParent2TableName & _
        mstrParent2ViewJoins & _
        " WHERE " & mstrParent2Where
    End If
    Set objTableView = Nothing
    
    mstrParent2Where = GetRecordIDs(strSQL)
    blnParentNotFound = blnParentNotFound Or (mstrParent2Where = vbNullString)
  
  End If

  'blnParentExists = (mstrParent1Where <> vbNullString Or mstrParent2Where <> vbNullString)
  'blnParentRequired = (mlngImportParent1Table <> 0 Or mlngImportParent2Table <> 0)


  'If (blnParentExists = False And blnParentRequired = True) Or blnParentNotFound Then
  If blnParentNotFound Then
    RecordError "Matching parent record not found"
  Else

    mstrSQLWhere = pstrBaseWhereSQL
    
    If mstrParent1Where <> vbNullString Then
      mstrSQLWhere = mstrSQLWhere & _
          IIf(mstrSQLWhere <> vbNullString, " AND ", "") & _
          "ID_" & mlngImportParent1Table & " IN (" & mstrParent1Where & ")"
    End If
    
    If mstrParent2Where <> vbNullString Then
      mstrSQLWhere = mstrSQLWhere & _
          IIf(mstrSQLWhere <> vbNullString, " AND ", "") & _
          "ID_" & mlngImportParent2Table & " IN (" & mstrParent2Where & ")"
    End If
    
    
    lngIndex = 0
    mstrIDsToUpdate = vbNullString
    If mstrBaseTableViewJoins <> vbNullString Then
      strSQL = "SELECT " & mstrImportBaseTableName & ".ID " & _
               "FROM " & mstrImportBaseTableName & _
               mstrBaseTableViewJoins & " WHERE " & mstrSQLWhere
      mstrIDsToUpdate = mstrIDsToUpdate & GetRecordIDs(strSQL)
    Else
      Do
        strSQL = "SELECT " & mstrBaseTableRealSources(lngIndex) & ".ID " & _
                 "FROM " & mstrBaseTableRealSources(lngIndex) & _
                 " WHERE " & mstrSQLWhere
        mstrIDsToUpdate = mstrIDsToUpdate & GetRecordIDs(strSQL)
      Loop While mstrIDsToUpdate = vbNullString And lngIndex < UBound(mstrBaseTableRealSources)
    End If


    If mstrIDsToUpdate = vbNullString Then    'No records
      If mlngImportType = 1 Then  'Update Only
        RecordError "Matching record not found"
      Else
        FindMatchingRecord = DoInsert
      End If
    
    ElseIf InStr(mstrIDsToUpdate, ",") Then   'Multiple records
      
      If Not mblnImportMultipleRecordAction Then
        RecordError "Key fields returned multiple records, and definition states not to update."
      Else
        FindMatchingRecord = DoUpdate
      End If
    
    Else   'Single record
      FindMatchingRecord = DoUpdate

    End If
  End If

  Exit Function
  
  
ErrTrap:
  RecordError "Error finding matching record (creating 'where' statement)."
  FindMatchingRecord = False
  
End Function


Private Function AddToWhere(ByRef strWhereSQL As String, strRealSources() As String, intCounter As Integer)

  Dim pobjColumnPrivileges As CColumnPrivileges
  Dim strColumnName As String
  Dim strValue As String
  Dim lngIndex As Long

  strValue = GetSpecificKeyData(intCounter)
  If mstrErrorString <> vbNullString Then
    Exit Function
  End If

  strWhereSQL = strWhereSQL & _
    IIf(strWhereSQL <> vbNullString, vbCrLf & " AND ", "")

  If UBound(strRealSources) = 0 Then
    strWhereSQL = strWhereSQL & _
          strRealSources(0) & "." & mvarColDetails(4, intCounter)
  Else
    strWhereSQL = strWhereSQL & _
          "CASE"
    For lngIndex = 0 To UBound(strRealSources)

      strColumnName = mvarColDetails(4, intCounter)
      Set pobjColumnPrivileges = GetColumnPrivileges(strRealSources(lngIndex))

      If pobjColumnPrivileges.IsValid(strColumnName) Then
        If pobjColumnPrivileges.item(strColumnName).AllowSelect Then
          strWhereSQL = strWhereSQL & vbCrLf & _
              " WHEN NOT " & strRealSources(lngIndex) & "." & mvarColDetails(4, intCounter) & _
              " IS NULL THEN " & strRealSources(lngIndex) & "." & mvarColDetails(4, intCounter)
        End If
      End If

      Set pobjColumnPrivileges = Nothing

    Next
    strWhereSQL = strWhereSQL & _
          " ELSE null END"
  End If

  'MH20050106 Fault 9592 - If null date then use "IS NULL" and not "= NULL"
  'strWhereSQL = strWhereSQL & " = " & strValue
  If mvarColDetails(7, intCounter) = True And strValue = "null" Then
    strWhereSQL = strWhereSQL & " is null"
  Else
    strWhereSQL = strWhereSQL & " = " & strValue
  End If

End Function


Private Function GetRecordIDs(strSQL As String) As String

  Dim rsTemp As Recordset
  Dim strOutput As String

  'gobjErrorStack.Disable
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
  'gobjErrorStack.Enable

  strOutput = vbNullString
  Do Until rsTemp.EOF
    strOutput = strOutput & _
      IIf(strOutput <> vbNullString, ",", "") & _
      rsTemp.Fields(0).Value
    rsTemp.MoveNext
  Loop
  Set rsTemp = Nothing

  GetRecordIDs = strOutput

End Function
            

Private Function RecordError(strErrorDesc As String, Optional blnAddToFailCount As Boolean = True)
  If mstrErrorString = vbNullString Then
    mstrErrorString = strErrorDesc
  End If
  gobjEventLog.AddDetailEntry "The following row failed to import: " & _
                               vbCrLf & vbCrLf & mstrBufferString & _
                               vbCrLf & vbCrLf & mstrErrorString
  If blnAddToFailCount Then
    mlngFailRecords = mlngFailRecords + 1
  End If
End Function


Private Function CheckIfNumeric(strInput As String) As Boolean

  Dim lngCount As Long

  For lngCount = 1 To Len(strInput)
    'JPD 20050628 Fault 10181
    If ((lngCount <> 1) Or (Mid(strInput, lngCount, 1) <> "-")) Then
      If InStr("0123456789.", Mid(strInput, lngCount, 1)) = 0 Then
        CheckIfNumeric = False
        Exit Function
      End If
    End If
  Next

  CheckIfNumeric = True

End Function

Private Function ReadSQLTableInToFile() As Boolean

  ' Read the defined data source and create a temporary file for import.
  ' I'm sorry this is not a good way to design an import, but time pressures and existing code constraints make this the quickest route
  Dim sSQL As String
  Dim sSelectString As String
  Dim sOutputLine As String
  Dim rsDataTable As Recordset
  Dim rsColumns As Recordset
  Dim bOK As Boolean
  Dim pintFileNo As Integer
  Dim iFieldCount As Integer
  Dim bFirstRow As Boolean

  bOK = True
    
  On Error GoTo ErrorTrap
    
  If Not CheckLinkedServerIsValid(mstrLinkedServer, mstrLinkedCatalog, mstrLinkedTable) Then
    mstrErrorString = "Error establishing connection to the linked server." & vbCrLf
    ReadSQLTableInToFile = False
    Exit Function
  End If
    
  If mbUseUpdateBlob Then
  
    sSQL = "SELECT COUNT(t.Column_Name) AS [ColumnCount]" & _
      " FROM [" & mstrLinkedServer & "].[" & mstrLinkedCatalog & "].INFORMATION_SCHEMA.COLUMNS t" & _
      " WHERE t.TABLE_NAME = '" & mstrLinkedTable & "' AND Column_Name IN ('PROCESS_DATE', 'PROCESSED', 'ERROR', 'ID')"
    Set rsColumns = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
    With rsColumns
      If Not (.BOF And .EOF) Then
        rsColumns.MoveFirst
        If rsColumns.Fields("ColumnCount").Value < 4 Then
          mstrErrorString = "The update blob has been incorrectly defined on the linked table." & vbCrLf
          ReadSQLTableInToFile = False
          Exit Function
        End If
      End If
    End With
    
  End If
    
  sSQL = "SELECT t.Column_Name" & _
    " FROM [" & mstrLinkedServer & "].[" & mstrLinkedCatalog & "].INFORMATION_SCHEMA.COLUMNS t" & _
    " WHERE t.TABLE_NAME = '" & mstrLinkedTable & "'  ORDER BY ORDINAL_POSITION"
  Set rsColumns = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  With rsColumns
    
    If .BOF And .EOF Then
      ReadSQLTableInToFile = False
      Return
    End If

    Do Until .EOF
      sSelectString = sSelectString & IIf(Len(sSelectString) > 0, ", ", "") & "[" & .Fields("Column_Name").Value & "]"
      .MoveNext
    Loop

  End With
  
  ' Now go grab the data
  sSQL = "SELECT " & sSelectString & " FROM [" & mstrLinkedServer & "].[" & mstrLinkedCatalog & "].dbo.[" & mstrLinkedTable & "]"
  
  If mbUseUpdateBlob Then
    sSQL = sSQL & " WHERE PROCESS_DATE IS NULL"
  End If
  
  Set rsDataTable = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  With rsDataTable

    pintFileNo = FreeFile
    mstrImportFileName = GetTmpFName
    bFirstRow = True
    Open mstrImportFileName For Output As pintFileNo

    Do Until .EOF
    
      sOutputLine = ""
      For iFieldCount = 0 To .Fields.Count - 1
        sOutputLine = sOutputLine & IIf(Len(sOutputLine) > 0, ",", "") & Chr(34) & .Fields(iFieldCount).Value & Chr(34)
      Next
       
      Print #pintFileNo, sOutputLine
      bFirstRow = False
      .MoveNext
    Loop

  End With

  Set rsColumns = Nothing
  Set rsDataTable = Nothing

  Close #pintFileNo
  ReadSQLTableInToFile = bOK

TidyUpAndExit:
  ReadSQLTableInToFile = bOK
  Exit Function

ErrorTrap:
  mstrErrorString = "Error whilst reading linked table." & vbCrLf & "(" & Err.Description & ")"
  bOK = False
  GoTo TidyUpAndExit

End Function

Private Sub AddLMSImportBlob()

  Dim iTemp As Integer
  iTemp = UBound(mvarColDetails, 2) + 1

  ReDim Preserve mvarColDetails(13, iTemp)
  mvarColDetails(0, iTemp) = "C"
  mvarColDetails(1, iTemp) = 0
  mvarColDetails(2, iTemp) = ""
  mvarColDetails(3, iTemp) = 0
  mvarColDetails(4, iTemp) = "PROCESS_DATE"
  mvarColDetails(5, iTemp) = False
  mvarColDetails(6, iTemp) = 10
  mvarColDetails(7, iTemp) = True
  mvarColDetails(8, iTemp) = False
  mvarColDetails(9, iTemp) = False
  mvarColDetails(10, iTemp) = TrimmingTypes.giTRIMMING_NONE
  mvarColDetails(13, iTemp) = True


  iTemp = iTemp + 1
  ReDim Preserve mvarColDetails(13, iTemp)
  mvarColDetails(0, iTemp) = "C"
  mvarColDetails(1, iTemp) = 0
  mvarColDetails(2, iTemp) = ""
  mvarColDetails(3, iTemp) = 0
  mvarColDetails(4, iTemp) = "PROCESSED"
  mvarColDetails(5, iTemp) = False
  mvarColDetails(6, iTemp) = 10
  mvarColDetails(7, iTemp) = False
  mvarColDetails(8, iTemp) = False
  mvarColDetails(9, iTemp) = True
  mvarColDetails(10, iTemp) = TrimmingTypes.giTRIMMING_NONE
  mvarColDetails(13, iTemp) = True

  iTemp = iTemp + 1
  ReDim Preserve mvarColDetails(13, iTemp)
  mvarColDetails(0, iTemp) = "C"
  mvarColDetails(1, iTemp) = 0
  mvarColDetails(2, iTemp) = ""
  mvarColDetails(3, iTemp) = 0
  mvarColDetails(4, iTemp) = "ERROR"
  mvarColDetails(5, iTemp) = False
  mvarColDetails(6, iTemp) = 10
  mvarColDetails(7, iTemp) = False
  mvarColDetails(8, iTemp) = False
  mvarColDetails(9, iTemp) = False
  mvarColDetails(10, iTemp) = TrimmingTypes.giTRIMMING_NONE
  mvarColDetails(13, iTemp) = True


  iTemp = iTemp + 1
  ReDim Preserve mvarColDetails(13, iTemp)
  mvarColDetails(0, iTemp) = "C"
  mvarColDetails(1, iTemp) = 0
  mvarColDetails(2, iTemp) = ""
  mvarColDetails(3, iTemp) = 0
  mvarColDetails(4, iTemp) = "ID"
  mvarColDetails(5, iTemp) = False
  mvarColDetails(6, iTemp) = 10
  mvarColDetails(7, iTemp) = False
  mvarColDetails(8, iTemp) = True
  mvarColDetails(9, iTemp) = False
  mvarColDetails(10, iTemp) = TrimmingTypes.giTRIMMING_NONE
  mvarColDetails(13, iTemp) = True

End Sub

Private Sub UpdateImportedRecords()

  Dim sSQL As String

  ReDim Preserve maryRecordSuccessIDs(UBound(maryRecordSuccessIDs) - 1)
  sSQL = "UPDATE [" & mstrLinkedServer & "].[" & mstrLinkedCatalog & "].dbo.[" & mstrLinkedTable & "] SET PROCESS_DATE = GETDATE() WHERE id IN (" & _
    "" & Join(maryRecordSuccessIDs, ",") & ")"
  mclsData.ExecuteSql (sSQL)

End Sub
