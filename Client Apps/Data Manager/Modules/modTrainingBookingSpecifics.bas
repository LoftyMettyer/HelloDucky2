Attribute VB_Name = "modTrainingBookingSpecifics"
Option Explicit

' Module parameters.
Public gfTrainingBookingEnabled As Boolean

' Module constants.
Public Const gsMODULEKEY_TRAININGBOOKING = "MODULE_TRAININGBOOKING"

' Course Records constants.
Private Const gsPARAMETERKEY_COURSETABLE = "Param_CourseTable"
Private Const gsPARAMETERKEY_COURSETITLE = "Param_CourseTitle"
Private Const gsPARAMETERKEY_COURSESTARTDATE = "Param_CourseStartDate"
Private Const gsPARAMETERKEY_COURSEENDDATE = "Param_CourseEndDate"
Private Const gsPARAMETERKEY_COURSENUMBERBOOKED = "Param_CourseNumberBooked"
Private Const gsPARAMETERKEY_COURSEMAXNUMBER = "Param_CourseMaxNumber"
Private Const gsPARAMETERKEY_COURSECANCELLATIONDATE = "Param_CourseCancelDate"
Private Const gsPARAMETERKEY_COURSECANCELLEDBY = "Param_CourseCancelledBy"
Private Const gsPARAMETERKEY_COURSETRANSFERPROVISIONALS = "Param_CourseTransferProvisionals"
Private Const gsPARAMETERKEY_COURSEINCLUDEPROVISIONALS = "Param_CourseIncludeProvisionals"
Private Const gsPARAMETERKEY_COURSEOVERBOOKINGNOTIFICATION = "Param_CourseOverbookingNotification"
Private Const gsPARAMETERKEY_COURSEORDER = "Param_CourseOrder"
' Pre-requisite constants.
Private Const gsPARAMETERKEY_PREREQTABLE = "Param_PreReqTable"
Private Const gsPARAMETERKEY_PREREQCOURSETITLE = "Param_PreReqCourseTitle"
Private Const gsPARAMETERKEY_PREREQGROUPING = "Param_PreReqGrouping"
Private Const gsPARAMETERKEY_PREREQFAILURE = "Param_PreReqFailure"
Private Const gsPARAMETERKEY_PREREQDFLTFAILURE = "Param_PreReqDfltFailure"
' Employee Records constants.
Private Const gsPARAMETERKEY_EMPLOYEETABLE = "Param_EmployeeTable"
Private Const gsPARAMETERKEY_EMPLOYEEORDER = "Param_EmployeeOrder"
Private Const gsPARAMETERKEY_BULKBOOKINGDEFAULTVIEW = "Param_BulkBookingDefaultView" 'NHRD01052003 Fault 4687
' Unavailability constants.
Private Const gsPARAMETERKEY_UNAVAILTABLE = "Param_UnavailTable"
Private Const gsPARAMETERKEY_UNAVAILFROMDATE = "Param_UnavailFromDate"
Private Const gsPARAMETERKEY_UNAVAILTODATE = "Param_UnavailToDate"
Private Const gsPARAMETERKEY_UNAVAILFAILURE = "Param_UnavailFailure"
Private Const gsPARAMETERKEY_UNAVAILDFLTFAILURE = "Param_UnavailDfltFailure"
' Waiting List constants.
Private Const gsPARAMETERKEY_WAITLISTTABLE = "Param_WaitListTable"
Private Const gsPARAMETERKEY_WAITLISTCOURSETITLE = "Param_WaitListCourseTitle"
Public Const gsPARAMETERKEY_WAITLISTOVERRIDECOLUMN = "Param_WaitListOverRideColumn"
' Training Booking constants.
Private Const gsPARAMETERKEY_TRAINBOOKTABLE = "Param_TrainBookTable"
Private Const gsPARAMETERKEY_TRAINBOOKCOURSETITLE = "Param_TrainBookCourseTitle"
Private Const gsPARAMETERKEY_TRAINBOOKCANCELDATE = "Param_TrainBookCancelDate"
Private Const gsPARAMETERKEY_TRAINBOOKSTATUS = "Param_TrainBookStatus"
Private Const gsPARAMETERKEY_TRAINBOOKOVERLAPNOTIFICATION = "Param_TrainBookOverlapNotification"
' Related Column constants.
Private Const gsPARAMETERKEY_TBWLRELATEDCOLUMNS = "Param_TBWLRelatedColumns"

Public glngCourseTableID As Long
Public gsCourseTableName As String
Public gsCourseTitleColumnName As String
Public gsCourseStartDateColumnName As String
Public gsCourseEndDateColumnName As String
Public gsCourseMaxNumberColumnName As String
Public gsCourseCancelDateColumnName As String
Public gsCourseCancelledByColumnName As String
Public gfCourseTransferProvisionals As Boolean
Public gfCourseIncludeProvisionals As Boolean
'Public giCourseOverbookingNotification As Integer
Public glngCourseOrderID As Long

Public gsPreReqTableName As String
'Private mvar_lngPreReqCourseTitleID As Long
'Public gsPreReqCourseTitleName As String
'Private mvar_lngPreReqGroupingID As Long
'Public gsPreReqGroupingColumnName As String
'Private mvar_lngPreReqFailureNotificationID As Long
'Public gsPreReqFailureNotificationColumnName As String
'Private mvar_iPreReqDfltFailureNotification As Integer

Public glngEmployeeTableID As Long
Public gsEmployeeTableName As String
Public glngEmployeeOrderID As Long
Public glngDefaultBulkBookingViewID As Long 'NHRD01052003 Fault 4687

Public gsUnavailTableName As String
'Private mvar_lngUnavailFromDateID As Long
'Public gsUnavailFromDateColumnName As String
'Private mvar_lngUnavailToDateID As Long
'Public gsUnavailToDateColumnName As String
'Private mvar_lngUnavailFailureNotificationID As Long
'Public gsUnavailFailureNotificationColumnName As String
'Private mvar_iUnavailDfltFailureNotification As Integer

Public glngWaitListTableID As Long
Public gsWaitListTableName As String
Public gsWaitListCourseTitleColumnName As String
Public gsWaitListOverrideColumnName As String

Public glngTrainBookTableID As Long
Public gsTrainBookTableName As String
'''Public gsTrainBookCourseTitleName As String
Public gsTrainBookStatusColumnName As String
Public gsTrainBookCancelDateColumnName As String
'Public giTrainBookOverlapNotification As Integer

Private mvar_alngRelatedColumns() As Long

Public gfTrainBookStatus_B As Boolean
Public gfTrainBookStatus_C As Boolean
Public gfTrainBookStatus_P As Boolean
Public gfTrainBookStatus_T As Boolean
Public gfTrainBookStatus_CC As Boolean

Public Function ReadTrainingBookingParameters() As Boolean
  ' Read the Training Booking module parameters from the database.
  Dim fOK As Boolean
  
  ' Read the Course record parameters.
  fOK = ReadCourseRecordParameters

  ' Read the Pre-requisite record parameters.
  If fOK Then
    fOK = ReadPreRequisiteParameters
  End If

  ' Read the Employee record parameters.
  If fOK Then
    fOK = ReadEmployeeRecordParameters
  End If

  ' Read the Unavailablility record parameters.
  If fOK Then
    fOK = ReadUnavailabilityParameters
  End If

  ' Read the Waiting List record parameters.
  If fOK Then
    fOK = ReadWaitingListParameters
  End If

  ' Read the Training Booking record parameters.
  If fOK Then
    fOK = ReadTrainingBookingRecordParameters
  End If

  ' Read the Related Columns.
  If fOK Then
    fOK = ReadRelatedColumns
  End If
  
  ReadTrainingBookingParameters = fOK
  
End Function





Private Function ReadRelatedColumns() As Boolean
  ' Read the Related Columns information into a local array.
  Dim iNextIndex As Integer
  Dim lngTrainBookColumnID As Long
  Dim lngWaitListColumnID As Long
  Dim rsColumns As Recordset
  
  ReDim mvar_alngRelatedColumns(2, 0)

  Set rsColumns = GetModuleArray(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TBWLRELATEDCOLUMNS)
  With rsColumns
    Do While Not .EOF
      ' Read the related column IDs from the database.
      lngTrainBookColumnID = IIf(IsNull(!sourcecolumnid), 0, !sourcecolumnid)
      lngWaitListColumnID = IIf(IsNull(!destcolumnid), 0, !destcolumnid)
      
      ' Add the column IDs to the array if they are both valid.
      If (lngTrainBookColumnID > 0) And _
        (lngWaitListColumnID > 0) Then

        iNextIndex = UBound(mvar_alngRelatedColumns, 2) + 1
        ReDim Preserve mvar_alngRelatedColumns(2, iNextIndex)
        mvar_alngRelatedColumns(1, iNextIndex) = lngTrainBookColumnID
        mvar_alngRelatedColumns(2, iNextIndex) = lngWaitListColumnID
      End If
      
      .MoveNext
    Loop
  
    .Close
  End With
  Set rsColumns = Nothing
  
  ReadRelatedColumns = True
  
End Function

Private Function ReadTrainingBookingRecordParameters() As Boolean
  ' Read the Training Booking parameter values from the database into local variables.
  Dim fOK As Boolean
'''  Dim lngTrainBookCourseTitleID As Long
  Dim lngTrainBookStatusID As Long
  Dim lngTrainBookCancelDateID As Long
  Dim sSQL As String
  Dim sErrMsg As String
  Dim objTable As CTablePrivilege
  Dim objColumns As CColumnPrivileges
  Dim objColumn As CColumnPrivilege
  Dim rsTemp As Recordset
  Dim datData As clsDataAccess

  fOK = True
  
  gfTrainBookStatus_B = False
  gfTrainBookStatus_C = False
  gfTrainBookStatus_P = False
  gfTrainBookStatus_T = False
  gfTrainBookStatus_CC = False
  
  glngTrainBookTableID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKTABLE))
  fOK = (glngTrainBookTableID > 0)
  If Not fOK Then
    sErrMsg = "'Training Booking' table not defined."
  Else
    Set objTable = gcoTablePrivileges.FindTableID(glngTrainBookTableID)
    fOK = Not objTable Is Nothing
    If Not fOK Then
      sErrMsg = "'Training Booking' table not found."
    Else
      gsTrainBookTableName = objTable.TableName
    End If
    Set objTable = Nothing
  End If

  ' Get the Course Title column information.
'''  If fOK Then
'''    lngTrainBookCourseTitleID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKCOURSETITLE))
'''    fOK = (lngTrainBookCourseTitleID > 0)
'''    If Not fOK Then
'''      sErrMsg = "'Course Title' column in the '" & gsTrainBookTableName & "' table not defined."
'''    Else
'''      Set objColumns = GetColumnPrivileges(gsTrainBookTableName)
'''      Set objColumn = objColumns.FindColumnID(lngTrainBookCourseTitleID)
'''      fOK = Not objColumn Is Nothing
'''      If Not fOK Then
'''        sErrMsg = "'Course Title' column in the '" & gsTrainBookTableName & "' table not found."
'''      Else
'''        gsTrainBookCourseTitleName = objColumn.ColumnName
'''      End If
'''      Set objColumn = Nothing
'''      Set objColumns = Nothing
'''    End If
'''  End If
  
  ' Get the Status column information.
  If fOK Then
    lngTrainBookStatusID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKSTATUS))
    fOK = (lngTrainBookStatusID > 0)
    If Not fOK Then
      sErrMsg = "'Status' column in the '" & gsTrainBookTableName & "' table not defined."
    Else
      Set objColumns = GetColumnPrivileges(gsTrainBookTableName)
      Set objColumn = objColumns.FindColumnID(lngTrainBookStatusID)
      fOK = Not objColumn Is Nothing
      If Not fOK Then
        sErrMsg = "'Status' column in the '" & gsTrainBookTableName & "' table not found."
      Else
        gsTrainBookStatusColumnName = objColumn.ColumnName
      End If
      Set objColumn = Nothing
      Set objColumns = Nothing
    End If
  End If
  
  ' Check what Status column values have been configured.
  If fOK Then
    Set datData = New clsDataAccess
    sSQL = "SELECT value" & _
      " FROM ASRSysColumnControlValues" & _
      " WHERE columnID = " & Trim(Str(lngTrainBookStatusID))
    
    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Do While Not rsTemp.EOF
      Select Case UCase(Trim(IIf(IsNull(rsTemp!Value), "", rsTemp!Value)))
        Case "B"
          gfTrainBookStatus_B = True
        Case "C"
          gfTrainBookStatus_C = True
        Case "P"
          gfTrainBookStatus_P = True
        Case "T"
          gfTrainBookStatus_T = True
        Case "CC"
          gfTrainBookStatus_CC = True
      End Select
      
      rsTemp.MoveNext
    Loop
    Set rsTemp = Nothing
    Set datData = Nothing
    
    ' B and C are mandatory
    If (Not gfTrainBookStatus_B) Or (Not gfTrainBookStatus_C) Then
      fOK = False
      sErrMsg = "'" & gsTrainBookStatusColumnName & "' column in the '" & gsTrainBookTableName & "' table missing values : " & _
        IIf(Not gfTrainBookStatus_B, "'B'" & IIf(Not gfTrainBookStatus_C, " & 'C'", ""), "'C'")
    End If
  End If

  ' Get the Training Booking Cancellation Date column information.
  ' NB. This column is optional.
  If fOK Then
    lngTrainBookCancelDateID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKCANCELDATE))
    If lngTrainBookCancelDateID > 0 Then
      Set objColumns = GetColumnPrivileges(gsTrainBookTableName)
      Set objColumn = objColumns.FindColumnID(lngTrainBookCancelDateID)
      fOK = Not objColumn Is Nothing
      If Not fOK Then
        sErrMsg = "'Training Booking Cancellation Date' column in the '" & gsTrainBookTableName & "' table not found."
      Else
        gsTrainBookCancelDateColumnName = objColumn.ColumnName
      End If
      Set objColumn = Nothing
      Set objColumns = Nothing
    Else
      gsTrainBookCancelDateColumnName = ""
    End If
  End If

'  giTrainBookOverlapNotification = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_TRAINBOOKOVERLAPNOTIFICATION))

  If Not fOK Then
'    COAMsgBox "Error reading Training Booking parameters." & vbCrLf & _
      sErrMsg & vbCrLf & vbCrLf & _
      "Training Booking functionality will be disabled.", vbCritical + vbYesNo, App.Title
    COAMsgBox "Error reading Training Booking parameters." & vbCrLf & _
      sErrMsg & vbCrLf & vbCrLf & _
      "Training Booking functionality will be disabled.", vbExclamation + vbOKOnly, App.Title
  End If
  
  ReadTrainingBookingRecordParameters = fOK
  
End Function





Private Function ReadWaitingListParameters() As Boolean
  ' Read the Waiting List parameter values from the database into local variables.
  Dim fOK As Boolean
  Dim lngWaitListCourseTitleID As Long
  Dim lngWaitListOverrideID As Long
  Dim sErrMsg As String
  Dim objTable As CTablePrivilege
  Dim objColumns As CColumnPrivileges
  Dim objColumn As CColumnPrivilege
  
  fOK = True
  
  glngWaitListTableID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_WAITLISTTABLE))
  fOK = (glngWaitListTableID > 0)
  If Not fOK Then
    sErrMsg = "'Waiting List' table not defined."
  Else
    Set objTable = gcoTablePrivileges.FindTableID(glngWaitListTableID)
    fOK = Not objTable Is Nothing
    If Not fOK Then
      sErrMsg = "'Waiting List' table not found."
    Else
      gsWaitListTableName = objTable.TableName
    End If
    Set objTable = Nothing
  End If
  
  ' Get the Course Title column information.
  If fOK Then
    lngWaitListCourseTitleID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_WAITLISTCOURSETITLE))
    fOK = (lngWaitListCourseTitleID > 0)
    If Not fOK Then
      sErrMsg = "'Course Title' column in the '" & gsWaitListTableName & "' table not defined."
    Else
      Set objColumns = GetColumnPrivileges(gsWaitListTableName)
      Set objColumn = objColumns.FindColumnID(lngWaitListCourseTitleID)
      fOK = Not objColumn Is Nothing
      If Not fOK Then
        sErrMsg = "'Course Title' column in the '" & gsWaitListTableName & "' table not found."
      Else
        gsWaitListCourseTitleColumnName = objColumn.ColumnName
      End If
      Set objColumn = Nothing
      Set objColumns = Nothing
    End If
  End If
  
  ' Get the Waiting List Override column information.
  If fOK Then
    lngWaitListOverrideID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_WAITLISTOVERRIDECOLUMN))
    If (lngWaitListOverrideID > 0) Then
      Set objColumns = GetColumnPrivileges(gsWaitListTableName)
      Set objColumn = objColumns.FindColumnID(lngWaitListOverrideID)
      If Not objColumn Is Nothing Then
        gsWaitListOverrideColumnName = objColumn.ColumnName
      End If
      Set objColumn = Nothing
      Set objColumns = Nothing
    End If
  End If
  
  If Not fOK Then
'    COAMsgBox "Error reading Training Booking parameters." & vbCrLf & _
      sErrMsg & vbCrLf & vbCrLf & _
      "Training Booking functionality will be disabled.", vbCritical + vbYesNo, App.Title
    COAMsgBox "Error reading Training Booking parameters." & vbCrLf & _
      sErrMsg & vbCrLf & vbCrLf & _
      "Training Booking functionality will be disabled.", vbExclamation + vbOKOnly, App.Title
  End If
  
  ReadWaitingListParameters = fOK
  
End Function




Private Function ReadUnavailabilityParameters() As Boolean
  ' Read the Unavailability parameter values from the database into local variables.
  Dim fOK As Boolean
  Dim lngUnavailTableID As Long
  Dim sErrMsg As String
  Dim objTable As CTablePrivilege
'  Dim objColumns As CColumnPrivileges
'  Dim objColumn As CColumnPrivilege
  
  fOK = True

  ' Get the table information.
  ' NB. This table is optional.
  lngUnavailTableID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILTABLE))
  If (lngUnavailTableID > 0) Then
    Set objTable = gcoTablePrivileges.FindTableID(lngUnavailTableID)
    fOK = Not objTable Is Nothing
    If Not fOK Then
      sErrMsg = "'Unavailability' table not found."
    Else
      gsUnavailTableName = objTable.TableName
    End If
    Set objTable = Nothing
  Else
    gsUnavailTableName = ""
  End If


'  mvar_lngUnavailFromDateID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILFROMDATE))
'  If mvar_lngUnavailFromDateID > 0 Then
'    gsUnavailFromDateColumnName = datGeneral.GetColumnName(mvar_lngUnavailFromDateID)
'  Else
'    gsUnavailFromDateColumnName = ""
'  End If
'
'  mvar_lngUnavailToDateID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILTODATE))
'  If mvar_lngUnavailToDateID > 0 Then
'    gsUnavailToDateColumnName = datGeneral.GetColumnName(mvar_lngUnavailToDateID)
'  Else
'    gsUnavailToDateColumnName = ""
'  End If
'
'  mvar_lngUnavailFailureNotificationID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILFAILURE))
'  If mvar_lngUnavailFailureNotificationID > 0 Then
'    gsUnavailFailureNotificationColumnName = datGeneral.GetColumnName(mvar_lngUnavailFailureNotificationID)
'  Else
'    gsUnavailFailureNotificationColumnName = ""
'  End If
'
'  mvar_iUnavailDfltFailureNotification = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_UNAVAILDFLTFAILURE))
    
  If Not fOK Then
'    COAMsgBox "Error reading Training Booking parameters." & vbCrLf & _
      sErrMsg & vbCrLf & vbCrLf & _
      "Training Booking functionality will be disabled.", vbCritical + vbYesNo, App.Title
    COAMsgBox "Error reading Training Booking parameters." & vbCrLf & _
      sErrMsg & vbCrLf & vbCrLf & _
      "Training Booking functionality will be disabled.", vbExclamation + vbOKOnly, App.Title
  End If
  
  ReadUnavailabilityParameters = fOK
  
End Function



Private Function ReadEmployeeRecordParameters() As Boolean
  ' Read the Employee Records parameter values from the database into local variables.
  Dim fOK As Boolean
  Dim sErrMsg As String
  Dim objTable As CTablePrivilege
  
  fOK = True
  
  ' Get the Employee table information.
  glngEmployeeTableID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_EMPLOYEETABLE))
  fOK = (glngEmployeeTableID > 0)
  If Not fOK Then
    sErrMsg = "'Employee' table not defined."
  Else
    Set objTable = gcoTablePrivileges.FindTableID(glngEmployeeTableID)
    fOK = Not objTable Is Nothing
    If Not fOK Then
      sErrMsg = "'Employee' table not found."
    Else
      gsEmployeeTableName = objTable.TableName
    End If
    Set objTable = Nothing
  End If
    
  If fOK Then
    glngEmployeeOrderID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_EMPLOYEEORDER))
    If glngEmployeeOrderID <= 0 Then
      Set objTable = gcoTablePrivileges.Item(gsEmployeeTableName)
      glngEmployeeOrderID = objTable.DefaultOrderID
      Set objTable = Nothing
    End If
  End If
  
  If Not fOK Then
    COAMsgBox "Error reading Training Booking parameters." & vbCrLf & _
      sErrMsg & vbCrLf & vbCrLf & _
      "Training Booking functionality will be disabled.", vbExclamation + vbOKOnly, App.Title
  End If
        
  ReadEmployeeRecordParameters = fOK
  
End Function

Public Function TrainingBooking_CheckAvailability(plngCourseID As Long, plngEmployeeID As Long) As Boolean
  ' Check that the given employee is available for the given course.
  ' Return TRUE if the booking can be made.
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "modTrainingBookingSpecifics.TrainingBooking_CheckAvailability(plngCourseID, plngEmployeeID)", Array(plngCourseID, plngEmployeeID)
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fDoneOK As Boolean
  Dim fDeadlock As Boolean
  Dim iRetryCount As Integer
  Dim iOldCursorLocation As Integer
  Dim sErrorMsg As String
  Dim ADOErr As ADODB.Error
  Dim ODBC As New ODBC
    
  Const iRETRIES = 5
  Const iPAUSE = 5000
    
  iOldCursorLocation = gADOCon.CursorLocation
  fDoneOK = True
  iRetryCount = 0

  fOK = True

  ' If no Unavailability table is defined then do nothing.
  If Len(gsUnavailTableName) > 0 Then
    ' Check for the existence of the sp_ASR_TBCheckUnavailability.
    sSQL = "SELECT COUNT(*) AS objectCount" & _
      "   FROM sysobjects" & _
      "   WHERE id = object_id('sp_ASR_TBCheckUnavailability')" & _
      "     AND sysstat & 0xf = 4"
    Set rsInfo = datGeneral.GetRecords(sSQL)
    
    If rsInfo!objectCount > 0 Then
      ' If it exists then run it to see if the delegate is available.
      Set cmADO = New ADODB.Command
      With cmADO
        .CommandText = "sp_ASR_TBCheckUnavailability"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set .ActiveConnection = gADOCon

        Set pmADO = .CreateParameter("courseRecordID", adInteger, adParamInput)
        .Parameters.Append pmADO
        pmADO.Value = plngCourseID

        Set pmADO = .CreateParameter("employeeRecordID", adInteger, adParamInput)
        .Parameters.Append pmADO
        pmADO.Value = plngEmployeeID

        Set pmADO = .CreateParameter("result", adInteger, adParamOutput)
        .Parameters.Append pmADO
    
        Set pmADO = Nothing

        fDeadlock = True
        Do While fDeadlock
          fDeadlock = False
          
          ' Change the cursor location to 'client' as the errors that might be raised
          ' during the update cannot be read for 'server' cursors.
          gADOCon.Errors.Clear
          gADOCon.CursorLocation = adUseClient
                
          On Error GoTo DeadlockErrorTrap
DeadlockRecoveryPoint:
          cmADO.Execute

          On Error GoTo ErrorTrap
  
          ' Restore the original cursor location to the ADO connection object.
          gADOCon.CursorLocation = iOldCursorLocation
                
          ' Check if the update prodcued any errors.
          If gADOCon.Errors.Count > 0 Then
            sErrorMsg = ""
          
            For Each ADOErr In gADOCon.Errors
              ' If any 'deadlocks' occur, try to save changes again.
              ' Do this a few times and if errors still occur then display a more friendly
              ' error message than the ' deadlock victim' one generated by ODBC.
              If (ADOErr.Number = DEADLOCK_ERRORNUMBER) And _
                (((UCase(Left(ADOErr.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                  (UCase(Right(ADOErr.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
                ((UCase(Left(ADOErr.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                  (InStr(UCase(ADOErr.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
  
                ' The error is for a deadlock.
                ' Sorry about having to use the err.description to trap the error but the err.number
                ' is not specific and MSDN suggests using the err.description.
                If (iRetryCount < iRETRIES) And (gADOCon.Errors.Count = 1) Then
                  iRetryCount = iRetryCount + 1
                  fDeadlock = True
                  ' Pause before resubmitting the SQL command.
                  Sleep iPAUSE
                Else
                  sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                    "Another user is deadlocking the database. Try saving again."
                  fDoneOK = False
                End If
    
              Else
                sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                  ADOErr.Description
                fDoneOK = False
              End If
            Next ADOErr
          
            gADOCon.Errors.Clear
        
            If Not fDoneOK Then
              COAMsgBox "ERROR." & vbCrLf & vbCrLf & _
                sErrorMsg, vbOKOnly + vbExclamation, App.ProductName
            End If
          End If
  
          If fDoneOK And (Not fDeadlock) Then
            Select Case .Parameters("result").Value
              Case 1    ' Employee unavailable (error).
                fOK = False
                COAMsgBox "The delegate is unavailable for the course." & vbCrLf & _
                  "Unable to make the booking.", vbOKOnly + vbInformation, App.ProductName
                  
              Case 2    ' Employee unavailable (over-rideable by the user).
                fOK = (COAMsgBox("The delegate is unavailable for the course." & vbCrLf & _
                  "Do you still want to make the booking ?", vbYesNo + vbQuestion, App.ProductName) = vbYes)
              
              Case Else ' Employee available.
                fOK = True
            End Select
          End If
        Loop
      
        Set cmADO = Nothing
      End With
    End If
    
    rsInfo.Close
    Set rsInfo = Nothing
  End If

TidyUpAndExit:
  If (iOldCursorLocation = adUseClient) Or _
    (iOldCursorLocation = adUseServer) Then
    gADOCon.CursorLocation = iOldCursorLocation
  Else
    gADOCon.CursorLocation = adUseServer
  End If
  Set ODBC = Nothing
  
  TrainingBooking_CheckAvailability = fOK
  gobjErrorStack.PopStack
  Exit Function
  
ErrorTrap:
  fDoneOK = False
  fOK = False
  COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name

  gobjErrorStack.HandleError
  Exit Function
  
DeadlockErrorTrap:
  ' If any 'deadlocks' occur, try to save changes again.
  ' Do this a few times and if errors still occur then display a more friendly
  ' error message than the 'deadlock victim' one generated by ODBC.
  If (Err.Number = DEADLOCK_ERRORNUMBER) And _
    (((UCase(Left(Err.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
      (UCase(Right(Err.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
    ((UCase(Left(Err.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
      (InStr(UCase(Err.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    ' The error is for a deadlock.
    ' Sorry about having to use the err.description to trap the error but the err.number
    ' is not specific and MSDN suggests using the err.description.
    If iRetryCount < iRETRIES Then
      iRetryCount = iRetryCount + 1
      ' Pause before resubmitting the SQL command.
      Sleep iPAUSE
      Resume DeadlockRecoveryPoint
    Else
      fDoneOK = False
      COAMsgBox "Another user is deadlocking the database. Try saving again.", _
        vbExclamation + vbOKOnly, Application.Name
      gobjErrorStack.HandleError
      Resume TidyUpAndExit
    End If
  Else
    fDoneOK = False
    fOK = False
    COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
    gobjErrorStack.HandleError
    Resume TidyUpAndExit
  End If
  
  
End Function

Public Function TrainingBooking_CheckOverbooking(plngCourseID As Long, plngBookingID As Long) As Boolean
  ' Check that the selected course is not already fully booked.
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "modTrainingBookingSpecifics.TrainingBooking_CheckOverbooking(plngCourseID, plngBookingID)", Array(plngCourseID, plngBookingID)
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fDoneOK As Boolean
  Dim fDeadlock As Boolean
  Dim iRetryCount As Integer
  Dim iOldCursorLocation As Integer
  Dim sErrorMsg As String
  Dim ADOErr As ADODB.Error
  Dim ODBC As New ODBC
    
  Const iRETRIES = 5
  Const iPAUSE = 5000
    
  iOldCursorLocation = gADOCon.CursorLocation
  fDoneOK = True
  iRetryCount = 0

  fOK = True

  ' Check for the existence of the sp_ASR_TBCheckOverbooking.
  sSQL = "SELECT COUNT(*) AS objectCount" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('sp_ASR_TBCheckOverbooking')" & _
    "     AND sysstat & 0xf = 4"
  Set rsInfo = datGeneral.GetRecords(sSQL)
    
  If rsInfo!objectCount > 0 Then
    ' If it exists then run it to see if the prerequisites have been met.
    Set cmADO = New ADODB.Command
    With cmADO
      .CommandText = "sp_ASR_TBCheckOverbooking"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon

      Set pmADO = .CreateParameter("courseRecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngCourseID

      Set pmADO = .CreateParameter("bookingRecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngBookingID

      Set pmADO = .CreateParameter("newBookings", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = 1

      Set pmADO = .CreateParameter("result", adInteger, adParamOutput)
      .Parameters.Append pmADO
  
      Set pmADO = Nothing

      fDeadlock = True
      Do While fDeadlock
        fDeadlock = False
        
        ' Change the cursor location to 'client' as the errors that might be raised
        ' during the update cannot be read for 'server' cursors.
        gADOCon.Errors.Clear
        gADOCon.CursorLocation = adUseClient
              
        On Error GoTo DeadlockErrorTrap
DeadlockRecoveryPoint:
        cmADO.Execute

        On Error GoTo ErrorTrap

        ' Restore the original cursor location to the ADO connection object.
        gADOCon.CursorLocation = iOldCursorLocation
              
        ' Check if the update prodcued any errors.
        If gADOCon.Errors.Count > 0 Then
          sErrorMsg = ""
        
          For Each ADOErr In gADOCon.Errors
            ' If any 'deadlocks' occur, try to save changes again.
            ' Do this a few times and if errors still occur then display a more friendly
            ' error message than the ' deadlock victim' one generated by ODBC.
            If (ADOErr.Number = DEADLOCK_ERRORNUMBER) And _
              (((UCase(Left(ADOErr.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                (UCase(Right(ADOErr.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
              ((UCase(Left(ADOErr.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                (InStr(UCase(ADOErr.Description), DEADLOCK2_MESSAGEEND) > 0))) Then

              ' The error is for a deadlock.
              ' Sorry about having to use the err.description to trap the error but the err.number
              ' is not specific and MSDN suggests using the err.description.
              If (iRetryCount < iRETRIES) And (gADOCon.Errors.Count = 1) Then
                iRetryCount = iRetryCount + 1
                fDeadlock = True
                ' Pause before resubmitting the SQL command.
                Sleep iPAUSE
              Else
                sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                  "Another user is deadlocking the database. Try saving again."
                fDoneOK = False
              End If
  
            Else
              sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                ADOErr.Description
              fDoneOK = False
            End If
          Next ADOErr
        
          gADOCon.Errors.Clear
      
          If Not fDoneOK Then
            COAMsgBox "ERROR." & vbCrLf & vbCrLf & _
              sErrorMsg, vbOKOnly + vbExclamation, App.ProductName
          End If
        End If

        If fDoneOK And (Not fDeadlock) Then
          Select Case .Parameters("result").Value
            Case 1    ' Course fully booked (error).
              fOK = False
              COAMsgBox "The course is already fully booked." & vbCrLf & _
                "Unable to make the booking.", vbOKOnly + vbInformation, App.ProductName
                
            Case 2    ' Course fully booked (over-rideable by the user).
              fOK = (COAMsgBox("The course is already fully booked." & vbCrLf & _
                "Do you still want to make the booking ?", vbYesNo + vbQuestion, App.ProductName) = vbYes)
                      
            Case Else ' Course NOT fully booked.
              fOK = True
          End Select
        End If
      Loop
    
      Set cmADO = Nothing
    End With
  End If
    
  rsInfo.Close
  Set rsInfo = Nothing

TidyUpAndExit:
  If (iOldCursorLocation = adUseClient) Or _
    (iOldCursorLocation = adUseServer) Then
    gADOCon.CursorLocation = iOldCursorLocation
  Else
    gADOCon.CursorLocation = adUseServer
  End If
  Set ODBC = Nothing
  
  TrainingBooking_CheckOverbooking = fOK
  gobjErrorStack.PopStack
  Exit Function
  
ErrorTrap:
  fDoneOK = False
  fOK = False
  COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name

  gobjErrorStack.HandleError
  Exit Function
  
DeadlockErrorTrap:
  ' If any 'deadlocks' occur, try to save changes again.
  ' Do this a few times and if errors still occur then display a more friendly
  ' error message than the 'deadlock victim' one generated by ODBC.
  If (Err.Number = DEADLOCK_ERRORNUMBER) And _
    (((UCase(Left(Err.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
      (UCase(Right(Err.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
    ((UCase(Left(Err.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
      (InStr(UCase(Err.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    ' The error is for a deadlock.
    ' Sorry about having to use the err.description to trap the error but the err.number
    ' is not specific and MSDN suggests using the err.description.
    If iRetryCount < iRETRIES Then
      iRetryCount = iRetryCount + 1
      ' Pause before resubmitting the SQL command.
      Sleep iPAUSE
      Resume DeadlockRecoveryPoint
    Else
      fDoneOK = False
      COAMsgBox "Another user is deadlocking the database. Try saving again.", _
        vbExclamation + vbOKOnly, Application.Name
      gobjErrorStack.HandleError
      Resume TidyUpAndExit
    End If
  Else
    fDoneOK = False
    fOK = False
    COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
    gobjErrorStack.HandleError
    Resume TidyUpAndExit
  End If
  
End Function


Public Function TrainingBooking_CheckOverlappedBooking(plngCourseID As Long, plngEmployeeID As Long, plngBookingID As Long) As Boolean
  ' Check that the given employee is not already booked on a course that overlaps
  ' with the given course.
  ' Return TRUE if the booking can be made.
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "modTrainingBookingSpecifics.TrainingBooking_CheckOverlappedBooking(plngCourseID, plngEmployeeID, plngBookingID)", Array(plngCourseID, plngEmployeeID, plngBookingID)
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  Dim fDoneOK As Boolean
  Dim fDeadlock As Boolean
  Dim iRetryCount As Integer
  Dim iOldCursorLocation As Integer
  Dim sErrorMsg As String
  Dim ADOErr As ADODB.Error
  Dim ODBC As New ODBC
    
  Const iRETRIES = 5
  Const iPAUSE = 5000
    
  iOldCursorLocation = gADOCon.CursorLocation
  fDoneOK = True
  iRetryCount = 0

  fOK = True

  ' Check for the existence of the sp_ASR_TBCheckOverlappedBooking.
  sSQL = "SELECT COUNT(*) AS objectCount" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('sp_ASR_TBCheckOverlappedBooking')" & _
    "     AND sysstat & 0xf = 4"
  Set rsInfo = datGeneral.GetRecords(sSQL)
  
  If rsInfo!objectCount > 0 Then
    ' If it exists then run it to see if the prerequisites have been met.
    Set cmADO = New ADODB.Command
    With cmADO
      .CommandText = "sp_ASR_TBCheckOverlappedBooking"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon

      Set pmADO = .CreateParameter("courseRecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngCourseID

      Set pmADO = .CreateParameter("employeeRecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngEmployeeID

      Set pmADO = .CreateParameter("bookingRecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngBookingID

      Set pmADO = .CreateParameter("result", adInteger, adParamOutput)
      .Parameters.Append pmADO
  
      Set pmADO = Nothing

      fDeadlock = True
      Do While fDeadlock
        fDeadlock = False
        
        ' Change the cursor location to 'client' as the errors that might be raised
        ' during the update cannot be read for 'server' cursors.
        gADOCon.Errors.Clear
        gADOCon.CursorLocation = adUseClient
              
        On Error GoTo DeadlockErrorTrap
DeadlockRecoveryPoint:
        cmADO.Execute

        On Error GoTo ErrorTrap

        ' Restore the original cursor location to the ADO connection object.
        gADOCon.CursorLocation = iOldCursorLocation
              
        ' Check if the update prodcued any errors.
        If gADOCon.Errors.Count > 0 Then
          sErrorMsg = ""
        
          For Each ADOErr In gADOCon.Errors
            ' If any 'deadlocks' occur, try to save changes again.
            ' Do this a few times and if errors still occur then display a more friendly
            ' error message than the ' deadlock victim' one generated by ODBC.
            If (ADOErr.Number = DEADLOCK_ERRORNUMBER) And _
              (((UCase(Left(ADOErr.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                (UCase(Right(ADOErr.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
              ((UCase(Left(ADOErr.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                (InStr(UCase(ADOErr.Description), DEADLOCK2_MESSAGEEND) > 0))) Then

              ' The error is for a deadlock.
              ' Sorry about having to use the err.description to trap the error but the err.number
              ' is not specific and MSDN suggests using the err.description.
              If (iRetryCount < iRETRIES) And (gADOCon.Errors.Count = 1) Then
                iRetryCount = iRetryCount + 1
                fDeadlock = True
                ' Pause before resubmitting the SQL command.
                Sleep iPAUSE
              Else
                sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                  "Another user is deadlocking the database. Try saving again."
                fDoneOK = False
              End If
  
            Else
              sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                ADOErr.Description
              fDoneOK = False
            End If
          Next ADOErr
        
          gADOCon.Errors.Clear
      
          If Not fDoneOK Then
            COAMsgBox "ERROR." & vbCrLf & vbCrLf & _
              sErrorMsg, vbOKOnly + vbExclamation, App.ProductName
          End If
        End If

        If fDoneOK And (Not fDeadlock) Then
          Select Case .Parameters("result").Value
            Case 1    ' Overlapped booking (error).
              fOK = False
              COAMsgBox "The delegate is already booked on a course that overlaps with this course." & vbCrLf & _
                "Unable to make the booking.", vbOKOnly + vbInformation, App.ProductName
                
            Case 2    ' Overlapped booking (over-rideable by the user).
              fOK = (COAMsgBox("The delegate is already booked on a course that overlaps with this course." & vbCrLf & _
                "Do you still want to make the booking ?", vbYesNo + vbQuestion, App.ProductName) = vbYes)
                      
            Case Else ' Course NOT fully booked.
              fOK = True
          End Select
        End If
      Loop
    
      Set cmADO = Nothing
    End With
  End If
  
  rsInfo.Close
  Set rsInfo = Nothing

TidyUpAndExit:
  If (iOldCursorLocation = adUseClient) Or _
    (iOldCursorLocation = adUseServer) Then
    gADOCon.CursorLocation = iOldCursorLocation
  Else
    gADOCon.CursorLocation = adUseServer
  End If
  Set ODBC = Nothing
  
  TrainingBooking_CheckOverlappedBooking = fOK
  gobjErrorStack.PopStack
  Exit Function
  
ErrorTrap:
  fDoneOK = False
  fOK = False
  COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name

  gobjErrorStack.HandleError
  Exit Function
  
DeadlockErrorTrap:
  ' If any 'deadlocks' occur, try to save changes again.
  ' Do this a few times and if errors still occur then display a more friendly
  ' error message than the 'deadlock victim' one generated by ODBC.
  If (Err.Number = DEADLOCK_ERRORNUMBER) And _
    (((UCase(Left(Err.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
      (UCase(Right(Err.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
    ((UCase(Left(Err.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
      (InStr(UCase(Err.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    ' The error is for a deadlock.
    ' Sorry about having to use the err.description to trap the error but the err.number
    ' is not specific and MSDN suggests using the err.description.
    If iRetryCount < iRETRIES Then
      iRetryCount = iRetryCount + 1
      ' Pause before resubmitting the SQL command.
      Sleep iPAUSE
      Resume DeadlockRecoveryPoint
    Else
      fDoneOK = False
      COAMsgBox "Another user is deadlocking the database. Try saving again.", _
        vbExclamation + vbOKOnly, Application.Name
      gobjErrorStack.HandleError
      Resume TidyUpAndExit
    End If
  Else
    fDoneOK = False
    fOK = False
    COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
    gobjErrorStack.HandleError
    Resume TidyUpAndExit
  End If
    
End Function




Private Function ReadPreRequisiteParameters() As Boolean
  ' Read the Pre-requisite parameter values from the database into local variables.
  Dim fOK As Boolean
  Dim lngPreReqTableID As Long
  Dim sErrMsg As String
  Dim objTable As CTablePrivilege
'  Dim objColumns As CColumnPrivileges
'  Dim objColumn As CColumnPrivilege
  
  fOK = True

  ' Get the table information.
  ' NB. This table is optional.
  lngPreReqTableID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQTABLE))
  If (lngPreReqTableID > 0) Then
    Set objTable = gcoTablePrivileges.FindTableID(lngPreReqTableID)
    fOK = Not objTable Is Nothing
    If Not fOK Then
      sErrMsg = "'Course Pre-requisites' table not found."
    Else
      gsPreReqTableName = objTable.TableName
    End If
    Set objTable = Nothing
  Else
    gsPreReqTableName = ""
  End If
  
  ' JPD - NO LONGER NEED TO READ THESE PARAMETERS.
'  mvar_lngPreReqCourseTitleID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQCOURSETITLE))
'  If mvar_lngPreReqCourseTitleID > 0 Then
'    gsPreReqCourseTitleName = datGeneral.GetColumnName(mvar_lngPreReqCourseTitleID)
'  Else
'    gsPreReqCourseTitleName = ""
'  End If
'
'  mvar_lngPreReqGroupingID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQGROUPING))
'  If mvar_lngPreReqGroupingID > 0 Then
'    gsPreReqGroupingColumnName = datGeneral.GetColumnName(mvar_lngPreReqGroupingID)
'  Else
'    gsPreReqGroupingColumnName = ""
'  End If
'
'  mvar_lngPreReqFailureNotificationID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQFAILURE))
'  If mvar_lngPreReqFailureNotificationID > 0 Then
'    gsPreReqFailureNotificationColumnName = datGeneral.GetColumnName(mvar_lngPreReqFailureNotificationID)
'  Else
'    gsPreReqFailureNotificationColumnName = ""
'  End If
'
'  mvar_iPreReqDfltFailureNotification = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_PREREQDFLTFAILURE))

  If Not fOK Then
'    COAMsgBox "Error reading Training Booking parameters." & vbCrLf & _
      sErrMsg & vbCrLf & vbCrLf & _
      "Training Booking functionality will be disabled.", vbCritical + vbYesNo, App.Title
    COAMsgBox "Error reading Training Booking parameters." & vbCrLf & _
      sErrMsg & vbCrLf & vbCrLf & _
      "Training Booking functionality will be disabled.", vbExclamation + vbOKOnly, App.Title
  End If
  
  ReadPreRequisiteParameters = fOK
  
End Function


Public Function TrainingBooking_CheckPreRequisites(plngCourseID As Long, plngEmployeeID As Long) As Boolean
  ' Check that given employee has (or will have) satisfied the pre-requisite criteria
  ' for the given course.
  ' Return TRUE if the bookings can be made.
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "modTrainingBookingSpecifics.TrainingBooking_CheckPreRequisites(plngCourseID, plngEmployeeID)", Array(plngCourseID, plngEmployeeID)
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fDoneOK As Boolean
  Dim fDeadlock As Boolean
  Dim iRetryCount As Integer
  Dim iOldCursorLocation As Integer
  Dim sErrorMsg As String
  Dim ADOErr As ADODB.Error
  Dim ODBC As New ODBC
    
  Const iRETRIES = 5
  Const iPAUSE = 5000
    
  iOldCursorLocation = gADOCon.CursorLocation
  fDoneOK = True
  iRetryCount = 0
  
  fOK = True
  
  ' If no prerequisite table is defined then do nothing.
  If Len(gsPreReqTableName) > 0 Then
  
    ' Check for the existence of the sp_ASR_TBCheckPreRequisites.
    sSQL = "SELECT COUNT(*) AS objectCount" & _
      "   FROM sysobjects" & _
      "   WHERE id = object_id('sp_ASR_TBCheckPreRequisites')" & _
      "     AND sysstat & 0xf = 4"
    Set rsInfo = datGeneral.GetRecords(sSQL)
    
    If rsInfo!objectCount > 0 Then
      ' If it exists then run it to see if the prerequisites have been met.
      Set cmADO = New ADODB.Command
      With cmADO
        .CommandText = "sp_ASR_TBCheckPreRequisites"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set .ActiveConnection = gADOCon

        Set pmADO = .CreateParameter("courseRecordID", adInteger, adParamInput)
        .Parameters.Append pmADO
        pmADO.Value = plngCourseID

        Set pmADO = .CreateParameter("employeeRecordID", adInteger, adParamInput)
        .Parameters.Append pmADO
        pmADO.Value = plngEmployeeID

        Set pmADO = .CreateParameter("preReqsMet", adInteger, adParamOutput)
        .Parameters.Append pmADO
    
        Set pmADO = Nothing

        fDeadlock = True
        Do While fDeadlock
          fDeadlock = False
          
          ' Change the cursor location to 'client' as the errors that might be raised
          ' during the update cannot be read for 'server' cursors.
          gADOCon.Errors.Clear
          gADOCon.CursorLocation = adUseClient
                
          On Error GoTo DeadlockErrorTrap
DeadlockRecoveryPoint:
          cmADO.Execute
  
          On Error GoTo ErrorTrap
  
          ' Restore the original cursor location to the ADO connection object.
          gADOCon.CursorLocation = iOldCursorLocation
                
          ' Check if the update prodcued any errors.
          If gADOCon.Errors.Count > 0 Then
            sErrorMsg = ""
          
            For Each ADOErr In gADOCon.Errors
              ' If any 'deadlocks' occur, try to save changes again.
              ' Do this a few times and if errors still occur then display a more friendly
              ' error message than the ' deadlock victim' one generated by ODBC.
              If (ADOErr.Number = DEADLOCK_ERRORNUMBER) And _
                (((UCase(Left(ADOErr.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                  (UCase(Right(ADOErr.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
                ((UCase(Left(ADOErr.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                  (InStr(UCase(ADOErr.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
  
                ' The error is for a deadlock.
                ' Sorry about having to use the err.description to trap the error but the err.number
                ' is not specific and MSDN suggests using the err.description.
                If (iRetryCount < iRETRIES) And (gADOCon.Errors.Count = 1) Then
                  iRetryCount = iRetryCount + 1
                  fDeadlock = True
                  ' Pause before resubmitting the SQL command.
                  Sleep iPAUSE
                Else
                  sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                    "Another user is deadlocking the database. Try saving again."
                  fDoneOK = False
                End If
    
              Else
                sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                  ADOErr.Description
                fDoneOK = False
              End If
            Next ADOErr
          
            gADOCon.Errors.Clear
        
            If Not fDoneOK Then
              COAMsgBox "ERROR." & vbCrLf & vbCrLf & _
                sErrorMsg, vbOKOnly + vbExclamation, App.ProductName
            End If
          End If
  
          If fDoneOK And (Not fDeadlock) Then
            Select Case .Parameters("preReqsMet").Value
              Case 1    ' Pre-requisites not satisfied (error).
                fOK = False
                COAMsgBox "The delegate has not met the pre-requisites for the course." & vbCrLf & _
                  "Unable to make the booking.", vbOKOnly + vbInformation, App.ProductName
                  
              Case 2    ' Pre-requisites not satisfied (over-rideable by the user).
                fOK = (COAMsgBox("The delegate has not met the pre-requisites for the course." & vbCrLf & _
                  "Do you still want to make the booking ?", vbYesNo + vbQuestion, App.ProductName) = vbYes)
              
              Case Else ' Pre-requisites satisfied.
                fOK = True
            End Select
          End If
        Loop
      
        Set cmADO = Nothing
      End With
    End If
    
    rsInfo.Close
    Set rsInfo = Nothing
  End If
  
TidyUpAndExit:
  If (iOldCursorLocation = adUseClient) Or _
    (iOldCursorLocation = adUseServer) Then
    gADOCon.CursorLocation = iOldCursorLocation
  Else
    gADOCon.CursorLocation = adUseServer
  End If
  Set ODBC = Nothing
  
  TrainingBooking_CheckPreRequisites = fOK
  gobjErrorStack.PopStack
  Exit Function
  
ErrorTrap:
  fDoneOK = False
  fOK = False
  COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name

  gobjErrorStack.HandleError
  Exit Function
  
DeadlockErrorTrap:
  ' If any 'deadlocks' occur, try to save changes again.
  ' Do this a few times and if errors still occur then display a more friendly
  ' error message than the 'deadlock victim' one generated by ODBC.
  If (Err.Number = DEADLOCK_ERRORNUMBER) And _
    (((UCase(Left(Err.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
      (UCase(Right(Err.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
    ((UCase(Left(Err.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
      (InStr(UCase(Err.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    ' The error is for a deadlock.
    ' Sorry about having to use the err.description to trap the error but the err.number
    ' is not specific and MSDN suggests using the err.description.
    If iRetryCount < iRETRIES Then
      iRetryCount = iRetryCount + 1
      ' Pause before resubmitting the SQL command.
      Sleep iPAUSE
      Resume DeadlockRecoveryPoint
    Else
      fDoneOK = False
      COAMsgBox "Another user is deadlocking the database. Try saving again.", _
        vbExclamation + vbOKOnly, Application.Name
      gobjErrorStack.HandleError
      Resume TidyUpAndExit
    End If
  Else
    fDoneOK = False
    fOK = False
    COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
    gobjErrorStack.HandleError
    Resume TidyUpAndExit
  End If
  
End Function


Private Function ReadCourseRecordParameters() As Boolean
  ' Read the Course parameters from the database.
  Dim fOK As Boolean
  Dim lngCourseTitleID As Long
  Dim lngCourseStartDateID As Long
  Dim lngCourseEndDateID As Long
'  dim lngCourseNumberBookedID As Long
  Dim lngCourseMaxNumberID As Long
  Dim lngCourseCancelDateID As Long
  Dim lngCourseCancelledByID As Long
  Dim sErrMsg As String
  Dim objTable As CTablePrivilege
  Dim objColumns As CColumnPrivileges
  Dim objColumn As CColumnPrivilege
  
  fOK = True
  
  ' Get the Training Course table information.
  glngCourseTableID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSETABLE))
  fOK = (glngCourseTableID > 0)
  If Not fOK Then
    sErrMsg = "'Training Course' table not defined."
  Else
    Set objTable = gcoTablePrivileges.FindTableID(glngCourseTableID)
    fOK = Not objTable Is Nothing
    If Not fOK Then
      sErrMsg = "'Training Course' table not found."
    Else
      gsCourseTableName = objTable.TableName
    End If
    Set objTable = Nothing
  End If
  
  ' Get the Course Title column information.
  If fOK Then
    lngCourseTitleID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSETITLE))
    fOK = (lngCourseTitleID > 0)
    If Not fOK Then
      sErrMsg = "'Course Title' column in the '" & gsCourseTableName & "' table not defined."
    Else
      Set objColumns = GetColumnPrivileges(gsCourseTableName)
      Set objColumn = objColumns.FindColumnID(lngCourseTitleID)
      fOK = Not objColumn Is Nothing
      If Not fOK Then
        sErrMsg = "'Course Title' column in the '" & gsCourseTableName & "' table not found."
      Else
        gsCourseTitleColumnName = objColumn.ColumnName
      End If
      Set objColumn = Nothing
      Set objColumns = Nothing
    End If
  End If
  
  ' Get the Course Start Date column information.
  If fOK Then
    lngCourseStartDateID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSESTARTDATE))
    fOK = (lngCourseStartDateID > 0)
    If Not fOK Then
      sErrMsg = "'Course Start Date' column in the '" & gsCourseTableName & "' table not defined."
    Else
      Set objColumns = GetColumnPrivileges(gsCourseTableName)
      Set objColumn = objColumns.FindColumnID(lngCourseStartDateID)
      fOK = Not objColumn Is Nothing
      If Not fOK Then
        sErrMsg = "'Course Start Date' column in the '" & gsCourseTableName & "' table not found."
      Else
        gsCourseStartDateColumnName = objColumn.ColumnName
      End If
      Set objColumn = Nothing
      Set objColumns = Nothing
    End If
  End If
  
  ' Get the Course End Date column information.
  If fOK Then
    lngCourseEndDateID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEENDDATE))
    fOK = (lngCourseEndDateID > 0)
    If Not fOK Then
      sErrMsg = "'Course End Date' column in the '" & gsCourseTableName & "' table not defined."
    Else
      Set objColumns = GetColumnPrivileges(gsCourseTableName)
      Set objColumn = objColumns.FindColumnID(lngCourseEndDateID)
      fOK = Not objColumn Is Nothing
      If Not fOK Then
        sErrMsg = "'Course End Date' column in the '" & gsCourseTableName & "' table not found."
      Else
        gsCourseEndDateColumnName = objColumn.ColumnName
      End If
      Set objColumn = Nothing
      Set objColumns = Nothing
    End If
  End If
  
  ' NOT USED ?
  ' lngCourseNumberBookedID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSENUMBERBOOKED))

  ' Get the Course Max. Number of Delegates column information.
  If fOK Then
    lngCourseMaxNumberID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEMAXNUMBER))
    fOK = (lngCourseMaxNumberID > 0)
    If Not fOK Then
      sErrMsg = "'Course Max. Number of Delegates' column in the '" & gsCourseTableName & "' table not defined."
    Else
      Set objColumns = GetColumnPrivileges(gsCourseTableName)
      Set objColumn = objColumns.FindColumnID(lngCourseMaxNumberID)
      fOK = Not objColumn Is Nothing
      If Not fOK Then
        sErrMsg = "'Course Max. Number of Delegates' column in the '" & gsCourseTableName & "' table not found."
      Else
        gsCourseMaxNumberColumnName = objColumn.ColumnName
      End If
      Set objColumn = Nothing
      Set objColumns = Nothing
    End If
  End If

  ' Get the Course Cancellation Date column information.
  ' NB. This column is optional.
  If fOK Then
    lngCourseCancelDateID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSECANCELLATIONDATE))
    If lngCourseCancelDateID > 0 Then
      Set objColumns = GetColumnPrivileges(gsCourseTableName)
      Set objColumn = objColumns.FindColumnID(lngCourseCancelDateID)
      fOK = Not objColumn Is Nothing
      If Not fOK Then
        sErrMsg = "'Course Cancellation Date' column in the '" & gsCourseTableName & "' table not found."
      Else
        gsCourseCancelDateColumnName = objColumn.ColumnName
      End If
      Set objColumn = Nothing
      Set objColumns = Nothing
    Else
      gsCourseCancelDateColumnName = ""
    End If
  End If

  lngCourseCancelledByID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSECANCELLEDBY))
  If lngCourseCancelledByID > 0 Then
    gsCourseCancelledByColumnName = datGeneral.GetColumnName(lngCourseCancelledByID)
  Else
    gsCourseCancelledByColumnName = ""
  End If

  ' JPD20020806 Fault 4274
  gfCourseTransferProvisionals = _
    IIf(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSETRANSFERPROVISIONALS) = "TRUE", True, False)

  ' JPD20020806 Fault 4274
  gfCourseIncludeProvisionals = _
    IIf(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEINCLUDEPROVISIONALS) = "TRUE", True, False)

  ' JPD - NO LONGER NEED TO READ THIS PARAMETER.
  'giCourseOverbookingNotification = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEOVERBOOKINGNOTIFICATION))

  If fOK Then
    glngCourseOrderID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_COURSEORDER))
    If glngCourseOrderID <= 0 Then
      Set objTable = gcoTablePrivileges.Item(gsCourseTableName)
      glngCourseOrderID = objTable.DefaultOrderID
      Set objTable = Nothing
    End If
  End If
  
  If Not fOK Then
'    COAMsgBox "Error reading Training Booking parameters." & vbCrLf & _
      sErrMsg & vbCrLf & vbCrLf & _
      "Training Booking functionality will be disabled.", vbCritical + vbYesNo, App.Title
    COAMsgBox "Error reading Training Booking parameters." & vbCrLf & _
      sErrMsg & vbCrLf & vbCrLf & _
      "Training Booking functionality will be disabled.", vbExclamation + vbOKOnly, App.Title
  End If
    
  'NHRD01052003 Fault 4687
  If fOK Then
    glngDefaultBulkBookingViewID = Val(GetModuleParameter(gsMODULEKEY_TRAININGBOOKING, gsPARAMETERKEY_BULKBOOKINGDEFAULTVIEW))
    If glngDefaultBulkBookingViewID <= 0 Then
      glngDefaultBulkBookingViewID = 0 ' needs work
      Set objTable = Nothing
    End If
  End If
  
  ReadCourseRecordParameters = fOK
  
End Function










Public Function RelatedColumns() As Variant
  ' Return the array of related columns.
  ' Column 1 = Training Booking table column ID.
  ' Column 2 = Waiting List table column ID.
  RelatedColumns = mvar_alngRelatedColumns

End Function
