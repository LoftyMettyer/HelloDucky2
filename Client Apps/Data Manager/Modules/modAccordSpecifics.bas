Attribute VB_Name = "modAccordSpecifics"
Option Explicit

' Payroll Module Settings
Public Const gsMODULEKEY_ACCORD = "MODULE_ACCORD"
Public Const gsPARAMETERKEY_ALLOWSDELETE = "Param_AllowDelete"
Public Const gsPARAMETERKEY_ALLOWSTATUSCHANGE = "Param_AllowStatusChange"
Public Const gsPARAMETERKEY_LOGINDETAILS = "Param_FieldsLoginDetails"
Public Const gsPARAMETERKEY_DEFAULTSTATUS = "Param_DefaultStatus"
Public Const gsPARAMETERKEY_PURGEOPTION = "Param_PurgeOption"
Public Const gsPARAMETERKEY_PURGEOPTIONPERIOD = "Param_PurgeOptionPeriod"
Public Const gsPARAMETERKEY_PURGEOPTIONPERIODTYPE = "Param_PurgeOptionPeriodType"

' Payroll Data Transfer connection variables
Public gADOAccordConnection As ADODB.Connection
Public gstrAccordLoginName As String
Public gstrAccordPassword As String
Public gstrAccordDatabase As String
Public gstrAccordServer As String

' Used to flag which set of Payroll data to look at (HRPro send, or Payroll read)
Public Enum AccordConnection
  ACCORD_REMOTE = 1
  ACCORD_LOCAL = 0
End Enum

Public Enum AccordTransactionStatus
  ACCORD_STATUS_UNKNOWN = 0
  ACCORD_STATUS_PENDING = 1
  ACCORD_STATUS_SUCCESS = 10
  ACCORD_STATUS_SUCCESS_WARNINGS = 11
  ACCORD_STATUS_FAILURE_UNKNOWN = 20
  ACCORD_STATUS_IGNORED = 21
  ACCORD_STATUS_ALREADY_EXISTS = 22
  ACCORD_STATUS_DOESNOT_EXIST = 23
  ACCORD_STATUS_MOREINFO_REQUIRED = 24
  ACCORD_STATUS_BLOCKED = 30
  ACCORD_STATUS_VOID = 31
End Enum


Public Sub GetAccordLogonDetails()

  Dim strInput As String
  Dim strEKey As String
  Dim strLens As String
  Dim lngStart As Long
  Dim lngFinish As Long

  ' Get logon details
  strInput = GetModuleParameter(gsMODULEKEY_ACCORD, gsPARAMETERKEY_LOGINDETAILS)
  If strInput = vbNullString Then
    Exit Sub
  End If

  lngStart = Len(strInput) - 14
  strEKey = Mid(strInput, lngStart + 1, 10)
  strLens = Right(strInput, 4)
  strInput = XOREncript(Left(strInput, lngStart), strEKey)

  lngStart = 1
  lngFinish = Asc(Mid(strLens, 1, 1)) - 127
  gstrAccordLoginName = Mid(strInput, lngStart, lngFinish)

  lngStart = lngStart + lngFinish
  lngFinish = Asc(Mid(strLens, 2, 1)) - 127
  gstrAccordPassword = Mid(strInput, lngStart, lngFinish)

  lngStart = lngStart + lngFinish
  lngFinish = Asc(Mid(strLens, 3, 1)) - 127
  gstrAccordDatabase = Mid(strInput, lngStart, lngFinish)

  lngStart = lngStart + lngFinish
  lngFinish = Asc(Mid(strLens, 4, 1)) - 127
  gstrAccordServer = Mid(strInput, lngStart, lngFinish)
  
End Sub

Public Function OpenAccordConnection() As Boolean

  On Error GoTo ErrorTrap

  Dim sConnect As String
  Dim bValidConnection As Boolean

  bValidConnection = True

  sConnect = "Driver=SQL Server;" & _
             "Server=" & gstrAccordServer & ";" & _
             "UID=" & gstrAccordLoginName & ";" & _
             "PWD=" & gstrAccordPassword & ";" & _
             "Database=" & gstrAccordDatabase & ";"

  Set gADOAccordConnection = New ADODB.Connection
  With gADOAccordConnection
    .ConnectionString = sConnect
    .Provider = "SQLOLEDB"
    .CommandTimeout = 0
    .ConnectionTimeout = 0
    .CursorLocation = adUseServer
    .Mode = adModeReadWrite
    .Properties("Packet Size") = 32767
    .Open
  End With

TidyUpAndExit:
  OpenAccordConnection = bValidConnection
  Exit Function

ErrorTrap:
  bValidConnection = False
  GoTo TidyUpAndExit


End Function

Public Function OpenAccordRecordset(ByVal piConnectionType As DataMgr.AccordConnection, _
  sSQL As String, CursorType As ADODB.CursorTypeEnum, _
  LockType As ADODB.LockTypeEnum, Optional varCursorLocation As Variant) As ADODB.Recordset
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "modAccordSpecifics.OpenAccordRecordset(sSQL,CursorType,LockType)", Array(sSQL, CursorType, LockType)
  
  ' Open a recordset from the given SQL query, with the given recordset properties.
  Dim tmpADOAccordConnection As ADODB.Connection
  Dim rsTemp As ADODB.Recordset
  Dim fDoneOK As Boolean
  Dim fDeadlock As Boolean
  Dim iRetryCount As Integer
  Dim iOldCursorLocation As Integer
  Dim sErrorMsg As String
  Dim ADOErr As ADODB.Error
  Dim ODBC As New ODBC
    
  Const iRETRIES = 5
  Const iPAUSE = 5000
    
  If piConnectionType = ACCORD_LOCAL Then
    Set tmpADOAccordConnection = gADOCon
  Else
    Set tmpADOAccordConnection = gADOAccordConnection
  End If
    
  iOldCursorLocation = tmpADOAccordConnection.CursorLocation
  fDoneOK = True
  iRetryCount = 0
    
  'JPD 20031120 Fault 7677
  If IsMissing(varCursorLocation) Then
    varCursorLocation = adUseClient
  End If
  
  Set rsTemp = New ADODB.Recordset
  
  fDeadlock = True
  Do While fDeadlock
    fDeadlock = False
    
    ' Change the cursor location to 'client' as the errors that might be raised
    ' during the update cannot be read for 'server' cursors.
    tmpADOAccordConnection.Errors.Clear
    
    'JPD 20031120 Fault 7677
    tmpADOAccordConnection.CursorLocation = adUseServer
          
    On Error GoTo DeadlockErrorTrap
DeadlockRecoveryPoint:
    rsTemp.Open sSQL, tmpADOAccordConnection, CursorType, LockType, adCmdText

    tmpADOAccordConnection.CursorLocation = iOldCursorLocation
          
    ' Check if the update prodcued any errors.
    If tmpADOAccordConnection.Errors.Count > 0 Then
      sErrorMsg = ""
    
      For Each ADOErr In tmpADOAccordConnection.Errors
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
          If (iRetryCount < iRETRIES) And (tmpADOAccordConnection.Errors.Count = 1) Then
            iRetryCount = iRetryCount + 1
            fDeadlock = True
            ' Pause before resubmitting the SQL command.
            Sleep iPAUSE
          Else
            sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
              "Another user is deadlocking the database."
            fDoneOK = False
          End If
        Else
          sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
            ADOErr.Description
          fDoneOK = False
        End If
      Next ADOErr
    
      tmpADOAccordConnection.Errors.Clear
  
      If Not fDoneOK Then
        If gobjProgress.Visible Then
          gobjProgress.CloseProgress
        End If
        COAMsgBox "ERROR." & vbCrLf & vbCrLf & _
          sErrorMsg, vbOKOnly + vbExclamation, App.ProductName
      End If
    End If

    If fDoneOK And (Not fDeadlock) Then
      Set OpenAccordRecordset = rsTemp
    End If
  Loop
        
TidyUpAndExit:
  If (iOldCursorLocation = adUseClient) Or _
    (iOldCursorLocation = adUseServer) Then
    tmpADOAccordConnection.CursorLocation = iOldCursorLocation
  Else
    tmpADOAccordConnection.CursorLocation = adUseServer
  End If
  Set ODBC = Nothing
  
  gobjErrorStack.PopStack
  Exit Function

ErrorTrap:
  fDoneOK = False
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
      If gobjProgress.Visible Then
        gobjProgress.CloseProgress
      End If
      COAMsgBox "Another user is deadlocking the database.", _
        vbExclamation + vbOKOnly, Application.Name
      gobjErrorStack.HandleError
      Resume TidyUpAndExit
    End If
  Else
    fDoneOK = False
    If gobjProgress.Visible Then
      gobjProgress.CloseProgress
    End If
    COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
    gobjErrorStack.HandleError
    Resume TidyUpAndExit
  End If

End Function

Public Sub ExecuteAccordSql(ByVal piConnectionType As DataMgr.AccordConnection, sSQL As String)
   
  Dim tmpADOAccordConnection As ADODB.Connection
    
  If piConnectionType = ACCORD_LOCAL Then
    Set tmpADOAccordConnection = gADOCon
  Else
    Set tmpADOAccordConnection = gADOAccordConnection
  End If
 
  ' Execute the given SQL statement.
  tmpADOAccordConnection.Execute sSQL, , adCmdText

End Sub

Public Function PopulateAccordTransferTypes(ByRef cboCombo As ComboBox, ByVal pbIncludeAll As Boolean) As Boolean

  Dim datData As New clsDataAccess
  Dim rstData As ADODB.Recordset
  Dim sSQL As String

  sSQL = "SELECT TransferType, TransferTypeID FROM ASRSysAccordTransferTypes WHERE IsVisible = 1" _
    & " AND ASRBaseTableID > 0 ORDER BY TransferTypeID"

  ' Get all of the transfer types
  With cboCombo
    
    If pbIncludeAll Then
      .AddItem "<All>"
      .ItemData(.NewIndex) = -1
    End If
    
    Set rstData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Do Until rstData.EOF
      .AddItem rstData.Fields("TransferType").Value
      .ItemData(.NewIndex) = rstData.Fields("TransferTypeID").Value
      rstData.MoveNext
    Loop
  
    If .ListCount > 0 Then
      .ListIndex = 0
    End If
  
    rstData.Close
    Set rstData = Nothing
  
  End With


End Function

Public Function IsTableMappedToAccord(ByVal lngTableID As Long) As Boolean
  IsTableMappedToAccord = True
  
  Dim datData As New clsDataAccess
  Dim rstData As ADODB.Recordset
  Dim sSQL As String

  IsTableMappedToAccord = False
  sSQL = "SELECT Count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE IsVisible = 1" _
    & " AND ASRBaseTableID = " & lngTableID
  Set rstData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
  If rstData.Fields(0).Value > 0 Then IsTableMappedToAccord = True
    
End Function

Public Function MappedAccordTransfers(ByVal lngTableID As Long) As Integer()

  Dim datData As New clsDataAccess
  Dim rstData As ADODB.Recordset
  Dim sSQL As String
  Dim iCount As Integer

  sSQL = "SELECT TransferTypeID FROM ASRSysAccordTransferTypes WHERE IsVisible = 1" _
    & " AND ASRBaseTableID = " & lngTableID
  Set rstData = datData.OpenRecordset(sSQL, adOpenKeyset, adLockReadOnly)
    
  ReDim iMappedAccordTransfers(rstData.RecordCount - 1) As Integer
  For iCount = 0 To rstData.RecordCount - 1
    iMappedAccordTransfers(iCount) = rstData.Fields("TransferTypeID").Value
    rstData.MoveNext
  Next
  
  MappedAccordTransfers = iMappedAccordTransfers
    
End Function

