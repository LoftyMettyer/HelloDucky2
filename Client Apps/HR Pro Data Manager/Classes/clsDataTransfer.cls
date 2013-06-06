VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataTransferRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'ExecuteDataTransfer
'  GetDataTransferDetails
'  CreateInsertTableArray
'    GetPicklistFilterSelect
'    BuildInsertForTable
'      GetInsertObject
'        ObjectCanUpdateAllColumns
'      CheckDirectTableAccess
'      AddToJoinCode
'      CheckViewAccess
'        AddToJoinCode
'      AddToInsertTableArray
'  CreateStoredProcedure
'
'
Option Explicit

Private fOK As Boolean
Private mblnUserCancelled As Boolean
Private datData As HRProDataMgr.clsDataAccess
Private rsDataTransferColumns As Recordset
Private mstrStatusMessage As String
Private mlngSelectedID As Long
Private mlngTableViews() As Long
Private mstrProcedureName As String
Private mstrGetMaxID As String

Private mlngSuccessCount As Long
Private mlngFailCount As Long

Private mstrSQLInsert As String
Private mstrSQLSelect As String
Private mstrSQLFrom As String
Private mstrSQLJoin As String
Private mstrSQLWhere As String
Private mstrSQLTable() As String

Private mstrTransferName As String
Private mlngFromTableID As Long
Private mstrFromTableName As String
Private mlngRecordDescExprID As Long
Private mlngToTableID As Long
Private mstrToTableName As String

Private mlngFilterID As Long
Private mlngPicklistID As Long
'Private mlngSingleRecordID As Long

Private mbDefinitionOwner As Boolean
Private mbLoggingDTSuccess As Boolean

' Array holding the User Defined functions that are needed for this report
Private mastrUDFsRequired() As String

Private mstrPicklistFilterIDs As String

Private Function IsRecordSelectionValid() As Boolean

  Dim sSQL As String
  Dim lCount As Long
  Dim rsTemp As Recordset
  Dim iResult As RecordSelectionValidityCodes
  
  ' Filter
  If mlngFilterID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_FILTER, mlngFilterID)
    Select Case iResult
      Case REC_SEL_VALID_DELETED
        mstrStatusMessage = "The base table filter used in this definition has been deleted by another user."
      Case REC_SEL_VALID_INVALID
        mstrStatusMessage = "The base table filter used in this definition is invalid."
      Case REC_SEL_VALID_HIDDENBYOTHER
        If Not gfCurrentUserIsSysSecMgr Then
          mstrStatusMessage = "The base table filter used in this definition has been made hidden by another user."
        End If
    End Select
  ElseIf mlngPicklistID > 0 Then
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, mlngPicklistID)
    Select Case iResult
      Case REC_SEL_VALID_DELETED
        mstrStatusMessage = "The base table picklist used in this definition has been deleted by another user."
      Case REC_SEL_VALID_INVALID
        mstrStatusMessage = "The base table picklist used in this definition is invalid."
      Case REC_SEL_VALID_HIDDENBYOTHER
        If Not gfCurrentUserIsSysSecMgr Then
          mstrStatusMessage = "The base table picklist used in this definition has been made hidden by another user."
        End If
      End Select
  End If
   
  IsRecordSelectionValid = (Len(mstrStatusMessage) = 0)
  
End Function



Private Sub Class_Initialize()
  Set datData = New clsDataAccess
End Sub

Private Sub Class_Terminate()
  Set datData = Nothing
End Sub

Public Property Get ErrorString()
  ErrorString = mstrStatusMessage
End Property

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property

Private Function Records(lngRec As Long) As String
  Records = CStr(lngRec) & IIf(lngRec <> 1, " records", " record")
End Function


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
  
  Progress = fOK
  
End Function

'Public Function ExecuteDataTransfer(lngSelectedID As Long, Optional lngRecordID As Long) As Boolean
Public Function ExecuteDataTransfer(lngSelectedID As Long, Optional strRecordIDs As String) As Boolean

  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim iShowMeAgain As Integer
  
  Dim msb As frmMessageBox

  ReDim mastrUDFsRequired(0)

  mlngSelectedID = lngSelectedID
  'mlngSingleRecordID = lngRecordID
  mstrPicklistFilterIDs = strRecordIDs
  
  fOK = True
  Screen.MousePointer = vbHourglass

  If fOK Then Call GetDataTransferDetails

'TM20020625 Fault 4056 - this is done in GetRecordIDs.
'  If fOK Then GetPicklistFilterSelect

  mbLoggingDTSuccess = CBool(GetUserSetting("LogEvents", "Data_Transfer_Success", False))

  'TM20010808 Fault 2659
  'TM20010814 Fault 2685
  'Moved message box code to prompt the user before progressbar is shown.
  If Not gblnBatchMode And fOK Then
    'TM20010726 Fault 1607 (Suggestion)
    strMBText = "WARNING: You are about to Transfer Data " & _
                "from the '" & mstrFromTableName & "' table into the '" & mstrToTableName & "' table." & vbNewLine & vbNewLine & _
                "Do you wish to continue?"
    
    intMBButtons = vbExclamation + vbYesNo + vbDefaultButton2
    
    iShowMeAgain = GetUserSetting("warningmsg", "warning datatransfer", 1)
    If iShowMeAgain = 1 Then
      Set msb = New frmMessageBox
      iShowMeAgain = IIf(iShowMeAgain = 1, 0, 1)
      mblnUserCancelled = (msb.MessageBox(strMBText, intMBButtons, "Data Transfer", iShowMeAgain, "&Don't show me this warning again.") <> vbYes)
      SaveUserSetting "warningmsg", "warning datatransfer", IIf(iShowMeAgain = 1, 0, 1)
      Set msb = Nothing
    End If
    fOK = Not mblnUserCancelled
  End If
  

  If fOK Then Call InitialiseProgressBar
  If Progress Then Call CheckIfTransferIsValid

  Call gobjEventLog.AddHeader(eltDataTransfer, mstrTransferName)

  If Progress Then Call CreateInsertTableArray
  If Progress Then Call GetUniqueName
  If Progress Then Call CreateStoredProcedure
  If Progress Then Call ProcessRecords
  If Progress Then Call UDFFunctions(mastrUDFsRequired, True)
  
  Call TidyUpAndExit
  Call OutputJobStatus
  
  ExecuteDataTransfer = fOK

End Function


Private Sub OutputJobStatus()

  Dim blnNoRecords As Boolean
  
  Call UtilUpdateLastRun(utlDataTransfer, mlngSelectedID)

  If mlngFailCount > 0 And Not mblnUserCancelled Then
    fOK = False
  End If

  'If Len(mstrStatusMessage) = 0 Then
  'If fOK Or blnNoRecords Then
  If fOK Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful, mlngSuccessCount, mlngFailCount
    mstrStatusMessage = "Completed successfully."

    blnNoRecords = (mlngSuccessCount = 0 And mlngFailCount = 0)

    If blnNoRecords Then
      mstrStatusMessage = mstrStatusMessage & vbCrLf & _
                        "No records meet selection criteria."
      gobjEventLog.AddDetailEntry mstrStatusMessage
      fOK = True
    End If

  ElseIf mblnUserCancelled Then
    gobjEventLog.ChangeHeaderStatus elsCancelled, mlngSuccessCount, mlngFailCount
    mstrStatusMessage = "Cancelled by user."
  
  Else
    gobjEventLog.ChangeHeaderStatus elsFailed, mlngSuccessCount, mlngFailCount
    If mstrStatusMessage <> vbNullString Then
      'Only details records for failures with description !
      gobjEventLog.AddDetailEntry mstrStatusMessage
      mstrStatusMessage = "Failed." & vbCrLf & vbCrLf & mstrStatusMessage
    Else
      mstrStatusMessage = "Failed." & vbCrLf
    End If

  End If


  If Not blnNoRecords Then
    mstrStatusMessage = mstrStatusMessage & vbCrLf & vbCrLf & _
      Records(mlngSuccessCount) & " successfully transferred."

    'If mlngFailCount > 0 And Not mlngSingleRecordID > 0 Then
    If mlngFailCount > 0 Then
      'If mstrPicklistFilterIDs = "" Or InStr(mstrPicklistFilterIDs, ",") > 0 Then
        mstrStatusMessage = mstrStatusMessage & vbCrLf & vbCrLf & _
          Records(mlngFailCount) & " failed during transfer."
      'End If
    End If
  End If


  mstrStatusMessage = "Data Transfer : '" & mstrTransferName & "' " & mstrStatusMessage

  If Not gblnBatchMode Then
    gobjProgress.CloseProgress
    Screen.MousePointer = vbNormal
    COAMsgBox mstrStatusMessage, IIf(fOK, vbInformation, vbExclamation), "Data Transfer"
  End If

End Sub

Private Sub InitialiseProgressBar()
  
  With gobjProgress
    '.AviFile = App.Path & "\videos\transfer.avi"
    .AVI = dbTransfer
    .MainCaption = "Data Transfer"
    If Not gblnBatchMode Then
      .NumberOfBars = 1
      .Caption = "Data Transfer"
      .Bar1Caption = "Data Transfer : " & mstrTransferName
      .Bar1RecordsCaption = "Records Processed :"
      .Time = False
      .Cancel = True
      .OpenProgress
    Else
      .ResetBar2
      .Bar2Caption = "Data Transfer : " & mstrTransferName
      .Bar2RecordsCaption = "Records Processed :"
    End If
  End With

End Sub
  

Private Sub GetDataTransferDetails()
  
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim lngTableCount As Long
  Dim strLastTable As String
  
  On Error GoTo LocalErr
  
  strSQL = "SELECT ASRSysDatatransferName.*, " & _
           "FromTable.TableName AS FromTableName, " & _
           "ToTable.TableName AS ToTableName, " & _
           "FromTable.RecordDescExprID AS RecDescExprID " & _
           "FROM ASRSysDatatransferName " & _
           "JOIN ASRSysTables FromTable " & _
           "  ON FromTable.TableID = ASRSysDataTransferName.FromTableID " & _
           "JOIN ASRSysTables ToTable " & _
           "  ON ToTable.TableID = ASRSysDataTransferName.ToTableID " & _
           "WHERE DataTransferID = " & CStr(mlngSelectedID)

  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  If rsTemp.BOF And rsTemp.EOF Then
    Set rsTemp = Nothing
    mstrStatusMessage = "This definition has been deleted by another user"
    fOK = False
    Exit Sub
  End If


  With rsTemp
    mstrTransferName = !Name
    mlngFromTableID = !fromTableID
    mstrFromTableName = !FromTableName
    mlngRecordDescExprID = !RecDescExprID
    mlngToTableID = !toTableID
    mstrToTableName = !ToTableName
    
    
    mlngFilterID = !FilterID
    mlngPicklistID = !PicklistID

    mbDefinitionOwner = (LCase(Trim(gsUserName)) = LCase(Trim(!UserName)))
  End With

  Set rsTemp = Nothing

  If Not IsRecordSelectionValid Then
    fOK = False
    Exit Sub
  End If

  strSQL = "SELECT " & _
           "FromTable.TableName AS FromTableName, " & _
           "FromColumn.ColumnName AS FromColumnName, " & _
           "ToTable.TableName AS ToTableName, " & _
           "ToColumn.ColumnName AS ToColumnName, " & _
           "ASRSysDataTransferColumns.* " & _
           "FROM ASRSysDataTransferColumns " & _
           "LEFT OUTER JOIN ASRSysTables FromTable " & _
           "  ON FromTable.TableID = ASRSysDataTransferColumns.FromTableID " & _
           "LEFT OUTER JOIN ASRSysColumns FromColumn " & _
           "  ON FromColumn.ColumnID = ASRSysDataTransferColumns.FromColumnID " & _
           "LEFT OUTER JOIN ASRSysTables ToTable " & _
           "  ON ToTable.TableID = ASRSysDataTransferColumns.ToTableID " & _
           "LEFT OUTER JOIN ASRSysColumns ToColumn " & _
           "  ON ToColumn.ColumnID = ASRSysDataTransferColumns.ToColumnID " & _
           "WHERE DataTransferID = " & CStr(mlngSelectedID) & _
           " ORDER BY ToTableID"

  
  Set rsDataTransferColumns = New Recordset
  rsDataTransferColumns.ActiveConnection = gADOCon
  rsDataTransferColumns.Properties("Preserve On Commit") = True
  rsDataTransferColumns.Properties("Preserve On Abort") = True
  rsDataTransferColumns.Open strSQL, , adOpenKeyset, adLockReadOnly, adCmdText

Exit Sub

LocalErr:
  mstrStatusMessage = "Error reading definition"
  fOK = False

End Sub


Private Sub GetUniqueName()

  On Error GoTo LocalErr
  
  mstrProcedureName = datGeneral.UniqueSQLObjectName("sp_ASRTempDataTransfer", 4)

Exit Sub

LocalErr:
  mstrStatusMessage = "Error Generating Unique Reference."
  fOK = False

End Sub


Private Sub CreateInsertTableArray()

  Dim strSQL As String
  Dim rsChildTables As Recordset
  
  ReDim mstrSQLTable(0) As String
  ReDim mlngTableViews(2, 0) As Long
  
  On Error GoTo LocalErr
  
  'mstrSQLFrom remains the same all insert statements !
  mstrSQLFrom = gcoTablePrivileges.Item(mstrFromTableName).RealSource

  mstrSQLInsert = vbNullString
  mstrSQLSelect = vbNullString
  mstrSQLJoin = vbNullString
  mstrSQLWhere = vbNullString
  
  
  Call CreateChildToChildRef
  Call BuildInsertForTable(mlngToTableID)
  If fOK = False Then
    Exit Sub
  End If

  
  'This will get all of the destination child table names
  'used in the definition
  strSQL = "SELECT DISTINCT " & _
           "ASRSysDataTransferColumns.ToTableID, " & _
           "ASRSysTables.TableName AS ToTableName " & _
           "FROM ASRSysDataTransferColumns " & _
           "LEFT OUTER JOIN ASRSysTables " & _
           "  ON ASRSysTables.TableID = ASRSysDataTransferColumns.ToTableID " & _
           "WHERE DataTransferID = " & CStr(mlngSelectedID) & _
           " AND ToTableID <> " & CStr(mlngToTableID)

  Set rsChildTables = datData.OpenRecordset(strSQL, adOpenDynamic, adLockReadOnly)
  
  With rsChildTables
  
    If Not .EOF Then .MoveFirst
    Do While Not .EOF
      
      'This is a child table so also need to insert into
      'ID column so that child references parent record
      mstrSQLInsert = "ID_" & CStr(mlngToTableID)
      mstrSQLSelect = "@MaxID"
      mstrSQLJoin = vbNullString
      mstrSQLWhere = vbNullString
      
      Call BuildInsertForTable(!toTableID)
      If fOK = False Then
        Exit Sub
      End If
      .MoveNext
    Loop

  End With

Exit Sub

LocalErr:
  mstrStatusMessage = "Error reading columns to transfer"
  fOK = False

End Sub


Private Function GetInsertObject(lngTableID As Long) As String

  'This sub will return the table/view which can access all
  'of the columns on the table which is in the data transfer
  'if no table/view can access all columns then vbnullstring
  'will be returned.
  
  Dim objTableView As CTablePrivilege
  Dim sTableName As String
  
  fOK = True
  sTableName = "<unknown>"
  
  For Each objTableView In gcoTablePrivileges.Collection
    If objTableView.TableID = lngTableID Then
      sTableName = objTableView.TableName
      
      If objTableView.AllowInsert Then
        If ObjectCanUpdateAllColumns(objTableView) Then
          GetInsertObject = objTableView.RealSource
          Exit Function
        End If
      
        If fOK = False Then
          Exit Function
        End If
      End If
    End If
  Next

  'If code reaches this point then we have looped through
  'all of the tables/views and not found any which can
  'access all of the columns on this table within the
  'definition.  Therefore we will not be able to proceed
  'with the insert command nor the data transfer
  GetInsertObject = vbNullString
  'JPD 20030411 Fault 5320
  mstrStatusMessage = "You do not have permission to insert all of the required " & _
                      "columns on the '" & sTableName & "' table."
  fOK = False

End Function


Private Function ObjectCanUpdateAllColumns(objTableView As CTablePrivilege) As Boolean

  'This sub will check the objTableView which is passed into this sub
  'and return true/false as to whether it it can update all columns
  
  Dim objColumnPrivileges As CColumnPrivileges
  Dim intCount As Integer
  Dim blnResult As Boolean

  On Error GoTo LocalErr
  
  If objTableView.IsTable Then
    Set objColumnPrivileges = GetColumnPrivileges(objTableView.TableName)
  Else
    Set objColumnPrivileges = GetColumnPrivileges(objTableView.ViewName)
  End If

  With rsDataTransferColumns

    .MoveFirst
    Do While Not .EOF

      If !toTableID = objTableView.TableID Then

        If objColumnPrivileges.IsValid(!ToColumnName) = False Then
          ObjectCanUpdateAllColumns = False
          Exit Function
                  
        ElseIf objColumnPrivileges.Item(!ToColumnName).AllowUpdate = False Then
          ObjectCanUpdateAllColumns = False
          Exit Function
        
        End If

      End If
    
      .MoveNext
    Loop

  End With

  ObjectCanUpdateAllColumns = True

Exit Function

LocalErr:
  'Err.Number -2147418113 = Catastrophic failure !!
  mstrStatusMessage = "Error reading destination table permissions"
  fOK = False

End Function


Private Sub BuildInsertForTable(lngTableID As Long)

  Dim sRealSource As String
  Dim strThisColumn As String
  Dim strInsertObjectName As String
  
  On Error GoTo LocalErr
  
  fOK = True
  
  'Check if we can do an insert either directly on
  'the table or through a view on the table
  strInsertObjectName = GetInsertObject(lngTableID)
  If strInsertObjectName = vbNullString Then
    Exit Sub
  End If
      
  ' JPD20030314 Fault 5159
  ReDim mlngTableViews(2, 0) As Long
    
  With rsDataTransferColumns

    .MoveFirst
    Do While Not .EOF And fOK

      If !toTableID = lngTableID Then
      
        mstrSQLInsert = mstrSQLInsert & _
          IIf(mstrSQLInsert <> vbNullString, ", ", vbNullString) & _
          !ToColumnName

        
        If !fromColumnID = 0 Then
          'Either System date or free text
          If !fromSysDate Then
            strThisColumn = Replace(Format(Now, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/")
          Else
            strThisColumn = Replace(!fromText, "'", "''")
          End If
          strThisColumn = "'" & strThisColumn & "'"


        ElseIf CheckDirectTableAccess(!FromTableName, !FromColumnName) Then
          'Get directly from table

          sRealSource = gcoTablePrivileges.Item(!FromTableName).RealSource
          strThisColumn = sRealSource & "." & !FromColumnName
        
          'Add Child table to join code
          If sRealSource <> mstrSQLFrom Then
            Call AddToJoinCode(0, !fromTableID, !fromTableID, sRealSource)
          End If
        
        
        Else
          'Build a case statement checking access from each view
          strThisColumn = CheckViewAccess(!fromTableID, !FromTableName, !FromColumnName)
          If fOK = False Then
            Exit Sub
          End If
        
        End If
      
        mstrSQLSelect = mstrSQLSelect & _
          IIf(mstrSQLSelect <> vbNullString, ", ", "") & _
          strThisColumn

      End If


      .MoveNext
    Loop

  End With

  
  Call AddToInsertTableArray(strInsertObjectName)

Exit Sub

LocalErr:
  mstrStatusMessage = "Error processing transfer columns"
  fOK = False

End Sub


Private Function CheckDirectTableAccess(strTableName As String, strColumnName As String) As Boolean

  'returns true/false as to whether the user can directly access columm

  Dim objColumnPrivileges As CColumnPrivileges
  Dim fColumnOK As Boolean
  
  On Error GoTo LocalErr
  
  CheckDirectTableAccess = False
  
  Set objColumnPrivileges = GetColumnPrivileges(strTableName)
  
  With objColumnPrivileges
    If .IsValid(strColumnName) Then
      CheckDirectTableAccess = .Item(strColumnName).AllowSelect
    End If
  End With
  
  Set objColumnPrivileges = Nothing

Exit Function

LocalErr:
  mstrStatusMessage = "Error reading source table permissions"
  fOK = False

End Function


Private Function CheckViewAccess(lngTableID As Long, strTableName As String, strColumnName As String) As String

  'This sub loop through all of the views on the given table and
  'builds the select case statement, adding the views to the join
  'statement where required.
  

  Dim objTableView As CTablePrivilege
  Dim objColumnPrivileges As CColumnPrivileges
  Dim iNextIndex As Integer
  Dim strWhereIDs As String

  On Error GoTo LocalErr

  CheckViewAccess = vbNullString
  strWhereIDs = vbNullString
        

  For Each objTableView In gcoTablePrivileges.Collection

    With objTableView

      'Loop thru all of the objects for this table
      'where the user has select access and the object is a view
      If (.TableID = lngTableID) And .AllowSelect And (Not .IsTable) Then

        ' Get the column permission for the view.
        Set objColumnPrivileges = GetColumnPrivileges(.ViewName)

        If objColumnPrivileges.IsValid(strColumnName) Then
          If objColumnPrivileges.Item(strColumnName).AllowSelect Then

            CheckViewAccess = CheckViewAccess & " WHEN NOT " & _
              .ViewName & "." & strColumnName & " IS NULL THEN " & _
              .ViewName & "." & strColumnName & vbCrLf

            'Add View to join Code
            ' JPD20030314 Fault 5159
            If AddToJoinCode(1, .TableID, .ViewID, .ViewName) Then

              'If this view is on the base table
              If .TableID = mlngFromTableID Then
                strWhereIDs = strWhereIDs & _
                  IIf(strWhereIDs <> vbNullString, " OR ", vbNullString) & _
                  .ViewName & ".ID IN (SELECT ID FROM " & mstrFromTableName & ")"
              End If
            End If
          End If
        End If

        Set objColumnPrivileges = Nothing

      End If

    End With

  Next objTableView
  Set objTableView = Nothing

      
  If CheckViewAccess <> vbNullString Then
    CheckViewAccess = _
      "CASE" & CheckViewAccess & " ELSE NULL END"
  Else
      mstrStatusMessage = "You do not have permission to see the '" & strTableName & "." & strColumnName & "'" & vbCrLf & _
                          "column either directly or through any views." & vbCrLf
      fOK = False
      Exit Function
  End If


  ' The current user does have permission to 'read'
  ' the column through a/some view(s) on the table.
  If strWhereIDs <> vbNullString Then
    mstrSQLWhere = mstrSQLWhere & _
      IIf(Len(mstrSQLWhere) > 0, " AND ", vbNullString) & _
      "(" & strWhereIDs & ")"
  End If

Exit Function

LocalErr:
  mstrStatusMessage = "Error checking available viewing on source table"
  fOK = False

End Function


Private Function AddToJoinCode(intType As Integer, lngTableID As Long, lngObjectID As Long, strSource As String) As Boolean

  Dim iNextIndex As Integer
  
  On Error GoTo LocalErr
  
  AddToJoinCode = False
  
  For iNextIndex = 1 To UBound(mlngTableViews, 2)
    If mlngTableViews(1, iNextIndex) = intType And _
      mlngTableViews(2, iNextIndex) = lngObjectID Then
      Exit Function
    End If
  Next iNextIndex
          
  AddToJoinCode = True
  
  ' if this column is not from the base table then it must be from a parent
  ' table, therefore include it in the join code
  iNextIndex = UBound(mlngTableViews, 2) + 1
  ReDim Preserve mlngTableViews(2, iNextIndex)
  mlngTableViews(1, iNextIndex) = intType
  mlngTableViews(2, iNextIndex) = lngObjectID
        
  ' The table has not yet been added to the join code, and it is
  ' not the base table so add it to the array and the join code.
  If lngTableID <> mlngFromTableID Then
    'If its a child table then make sure that
    'we reference the parent ID column and do an inner join
    mstrSQLJoin = mstrSQLJoin & _
      " INNER JOIN " & strSource & _
      " ON " & mstrFromTableName & ".ID" & _
      " = " & strSource & ".ID_" & CStr(mlngFromTableID)
  Else
    'Join view of parent and left outer join
    mstrSQLJoin = mstrSQLJoin & _
      " LEFT OUTER JOIN " & strSource & _
      " ON " & mstrFromTableName & ".ID" & _
      " = " & strSource & ".ID"

  End If

Exit Function

LocalErr:
  mstrStatusMessage = "Error joining source tables"
  fOK = False

End Function


Private Sub AddToInsertTableArray(strInsertObjectName As String)

  Dim intIndex As Integer

  On Error GoTo LocalErr
  
  
  If mstrSQLTable(0) = vbNullString Then
    intIndex = 0
  Else
    intIndex = UBound(mstrSQLTable) + 1
    ReDim Preserve mstrSQLTable(intIndex)
  End If


  mstrSQLTable(intIndex) = _
    "INSERT " & strInsertObjectName & vbCrLf & _
    "(" & mstrSQLInsert & ")" & vbCrLf

  mstrSQLTable(intIndex) = mstrSQLTable(intIndex) & _
    "SELECT " & mstrSQLSelect & vbCrLf & _
    "FROM " & mstrSQLFrom & vbCrLf
    
  If mstrSQLJoin <> vbNullString Then
    mstrSQLTable(intIndex) = mstrSQLTable(intIndex) & _
      mstrSQLJoin & vbCrLf
  End If

  mstrSQLTable(intIndex) = mstrSQLTable(intIndex) & _
    " WHERE " & mstrSQLWhere & _
    IIf(mstrSQLWhere <> vbNullString, " AND ", vbNullString) & _
    mstrSQLFrom & ".ID = @RecordID" & vbCrLf & vbCrLf


  If intIndex = 0 Then
    'If its the first table then select the max id from the table afterwards
    mstrGetMaxID = strInsertObjectName
  End If


Exit Sub

LocalErr:
  mstrStatusMessage = "Error building SQL insert statement"
  fOK = False

End Sub


Private Sub CreateStoredProcedure()

  Dim blnChildDestinationTables As Boolean
  Dim strSQL As String
  Dim intCount As Integer

  On Error GoTo LocalErr
  

  strSQL = "/* ------------------------------------- */" & vbCrLf & _
           "/* HR Pro Data Transfer stored procedure. */" & vbCrLf & _
           "/* ------------------------------------- */" & vbCrLf & _
           "CREATE PROCEDURE " & mstrProcedureName & "(@RecordID int)" & vbCrLf & _
           "AS" & vbCrLf & _
           "BEGIN" & vbCrLf & vbCrLf

         
         'MH20010308 Fault 1859
         '"IF NOT EXISTS(SELECT ID FROM " & mstrFromTableName & _

  strSQL = strSQL & _
         "IF NOT EXISTS(SELECT ID FROM " & mstrSQLFrom & _
         " WHERE ID = @RecordID)" & vbCrLf & _
         "BEGIN" & vbCrLf & _
         "  RAISERROR('Record not found', 16, 1)" & vbCrLf & _
         "  RETURN 1" & vbCrLf & _
         "END" & vbCrLf & vbCrLf

  If UBound(mstrSQLTable) > 0 Then
    'More then one table to insert to
    
    strSQL = strSQL & _
           "DECLARE @MaxID int" & vbCrLf & vbCrLf & _
           "BEGIN TRANSACTION" & vbCrLf & vbCrLf
    
    ' JPD20030314 Fault 5159
    strSQL = strSQL & _
      mstrSQLTable(0) & vbCrLf & vbCrLf & _
      "/* -------------------------------------------------- */" & vbCrLf & _
      "/* Get max id from primary table to ensure that child */" & vbCrLf & _
      "/* tables correctly reference the parent table        */" & vbCrLf & _
      "/* -------------------------------------------------- */" & vbCrLf & _
      "EXEC spASRMaxID " & Trim(Str(mlngToTableID)) & ", @MaxID OUTPUT" & vbCrLf
      '"SELECT @MaxID = MAX(ID) FROM " & mstrGetMaxID & vbCrLf
    
    For intCount = 1 To UBound(mstrSQLTable)
      strSQL = strSQL & _
           mstrSQLTable(intCount) & vbCrLf & vbCrLf
    Next
    
    strSQL = strSQL & _
           "COMMIT TRANSACTION" & vbCrLf & vbCrLf

  Else
    'Simple insert on a single table
    strSQL = strSQL & _
           mstrSQLTable(0) & vbCrLf & vbCrLf
  
  End If
  
  strSQL = strSQL & _
           "END"

  datData.ExecuteSql strSQL

Exit Sub

LocalErr:
  mstrStatusMessage = "Error creating stored procedure"
  fOK = False

End Sub


Private Sub ProcessRecords()

  Dim objDeadlock As clsDeadlock
  Dim rsRecords As Recordset
  Dim strSQL As String
  Dim strPicklistFilter As String
'  Dim strToTableRealSource As String
  
  Dim blnTransferOK As Boolean
  Dim strRecordError As String
  Dim strOutput As String

  On Error GoTo LocalErr

  Set objDeadlock = New clsDeadlock
'  strToTableRealSource = gcoTablePrivileges.Item(mstrToTableName).RealSource

'  strPicklistFilter = GetPicklistFilterSelect
'  If strPicklistFilter <> vbNullString Then
'    strPicklistFilter = " WHERE ID IN (" & strPicklistFilter & ")"
'  End If
'  strPicklistFilter = "SELECT ID FROM " & mstrSQLFrom & strPicklistFilter
'
'  If fOK = False Then
'    Exit Sub
'  End If
'
'  Set rsRecords = datData.OpenRecordset(strPicklistFilter, adOpenKeyset, adLockReadOnly)

  Set rsRecords = GetRecordIDs
  If fOK = False Then
    Exit Sub
  End If

  'Run the stored procedure for each record id
  With rsRecords

    If gblnBatchMode Then
      gobjProgress.Bar2MaxValue = rsRecords.RecordCount
    Else
      gobjProgress.Bar1MaxValue = rsRecords.RecordCount
    End If

    mlngSuccessCount = 0
    mlngFailCount = 0
    Do While Not .EOF

      strSQL = "EXEC " & mstrProcedureName & " " & CStr(!ID)
      'blnTransferOK = ValidSQL(strSQL, strRecordError)
      blnTransferOK = objDeadlock.InsertTableRecord(strSQL, mlngToTableID, 0)
      strRecordError = objDeadlock.ErrorString
            
      If objDeadlock.FatalError Then
        mstrStatusMessage = "Error transferring records (" & objDeadlock.ErrorString & ")"
        fOK = False
        Exit Sub
      End If
      
      
      If blnTransferOK Then
        mlngSuccessCount = mlngSuccessCount + 1
        
        'JDM - 06/02/02 - Fault 3280 - Log successful records
        If mbLoggingDTSuccess Then
          gobjEventLog.AddDetailEntry GetRecordDesc(!ID) & " transferred successfully"
        End If
      
      ElseIf strRecordError <> vbNullString Then
        
        'If mlngSingleRecordID > 0 Then
        If mstrPicklistFilterIDs <> "" And InStr(mstrPicklistFilterIDs, ",") = 0 Then
          mstrStatusMessage = strRecordError
        End If
        
        strRecordError = GetRecordDesc(!ID) & vbCrLf & vbCrLf & _
                         strRecordError
        Call gobjEventLog.AddDetailEntry(strRecordError)
        mlngFailCount = mlngFailCount + 1
        'fOK = False
      End If

      gobjProgress.UpdateProgress gblnBatchMode
      strOutput = "Records Processed : " & CStr(mlngSuccessCount) & " successful" & _
                  IIf(mlngFailCount > 0, ", " & CStr(mlngFailCount) & " failed", "")

      If gblnBatchMode Then
        gobjProgress.Bar2RecordsCaption = strOutput
      Else
        gobjProgress.Bar1RecordsCaption = strOutput
      End If
      
      If gobjProgress.Cancelled Then
        mblnUserCancelled = True
        fOK = False
        Exit Do
      End If
      
      .MoveNext
      
    Loop
  
  End With

  Set objDeadlock = Nothing
  
  'MH20021014
  objEmail.SendImmediateEmails
  
Exit Sub

LocalErr:
  mstrStatusMessage = "Error transferring records (" & Err.Description & ")"
  fOK = False

End Sub


Private Function GetRecordIDs() As Recordset

  Dim objTableView As CTablePrivilege
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim lngLoop As Long
  Dim strPicklistFilterIDs As String
  Dim blnAllowed As Boolean

  On Error GoTo LocalErr
  
  
  'Set up an array of record ids to be updated
  ReDim mlngRecordIDs(0) As Long
  
  strPicklistFilterIDs = GetPicklistFilterSelect
  If fOK = False Then
    Exit Function
  End If
  
    
  'We need to build a string which will select the ID from all of the views
  'that we have select permission on
  strSQL = vbNullString
  For Each objTableView In gcoTablePrivileges.Collection

    blnAllowed = (objTableView.TableID = mlngFromTableID And objTableView.AllowSelect)
    
    If blnAllowed Then
      strSQL = strSQL & _
             IIf(strSQL <> vbNullString, vbCrLf & "UNION" & vbCrLf, vbNullString) & _
             "SELECT ID FROM " & objTableView.RealSource & _
             IIf(strPicklistFilterIDs <> vbNullString, _
                          " WHERE ID IN (" & strPicklistFilterIDs & ")", _
                          vbNullString)
    End If
  
  Next


  'Now get distinct IDs from the base table (as there couuld be duplicates in the above sql!)
  strSQL = "SELECT ID FROM " & mstrSQLFrom & _
           IIf(strSQL <> vbNullString, " WHERE ID IN (" & strSQL & ")", vbNullString)
  'Set GetRecordIDs = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  
  ' Create any UDFs
  fOK = UDFFunctions(mastrUDFsRequired, True)
  
  Set GetRecordIDs = datData.OpenRecordset(strSQL, adOpenKeyset, adLockReadOnly)
  
Exit Function

LocalErr:
  mstrStatusMessage = "Error selecting records"
  fOK = False

End Function


Private Sub TidyUpAndExit()

  Dim strSQL As String
  
  On Error Resume Next

  'TM20020531 Fault 3756
  'Drop the temp procedure
'  If mstrProcedureName <> vbNullString Then
'    strSQL = "IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & mstrProcedureName & "') " & _
'             "DROP PROCEDURE " & mstrProcedureName
'    datData.ExecuteSql (strSQL)
'  End If
  gobjProgress.ResetBar2
  datGeneral.DropUniqueSQLObject mstrProcedureName, 4
  
End Sub


Private Sub CreateChildToChildRef()

  Dim rsTemp As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT ParentID FROM ASRSysRelations " & _
           "WHERE ParentID IN " & _
           "(SELECT ParentID FROM ASRSysRelations " & _
           " WHERE ChildID = " & CStr(mlngFromTableID) & _
           ") AND ChildID = " & CStr(mlngToTableID)

  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  mstrSQLInsert = vbNullString
  Do While Not rsTemp.EOF
    mstrSQLInsert = mstrSQLInsert & _
      IIf(mstrSQLInsert <> vbNullString, ", ", vbNullString) & _
        "ID_" & CStr(rsTemp!ParentID)
    rsTemp.MoveNext
  Loop

  mstrSQLSelect = mstrSQLInsert

  rsTemp.Close
  Set rsTemp = Nothing

End Sub

Private Sub CheckIfTransferIsValid()

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strSQL1 As String
  Dim strSQL2 As String
  Dim strSQL3 As String
  Dim strTablesInDef As String
  Dim strColumnsInDef As String
  Dim strErrorColumns As String
  
  
  strTablesInDef = "SELECT DISTINCT ASRSysDataTransferColumns.ToTableID " & _
                   "FROM ASRSysDataTransferColumns " & _
                   "WHERE ASRSysDataTransferColumns.DataTransferID = " & CStr(mlngSelectedID)
  
  strColumnsInDef = "SELECT ASRSysDataTransferColumns.ToColumnID " & _
                   "FROM ASRSysDataTransferColumns " & _
                   "WHERE ASRSysDataTransferColumns.DataTransferID = " & CStr(mlngSelectedID)

  
  'This will retreive all of the read only
  'destination columns in the data transfer
  strSQL1 = "SELECT ASRSysTables.TableName+'.'+ASRSysColumns.ColumnName as 'TableColumn', " & _
            "'Read Only' as Reason " & _
           "FROM ASRSysDataTransferColumns " & _
           "JOIN ASRSysColumns ON ASRSysColumns.ColumnID = ASRSysDataTransferColumns.ToColumnID " & _
           "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID " & _
           "WHERE ASRSysDataTransferColumns.DataTransferID = " & CStr(mlngSelectedID) & _
           " AND ASRSysColumns.ReadOnly <> 0"
  
  
  'This will get all of the mandatory columns
  'which have not been included in the definition
  
  'MH20000814
  'Allow save if mandatory ommitted if it has a default value
  'This is to get around the staff number on a applicants to personnel transfer
  
  'MH20000904
  'Allow save if mandatory ommitted and it is a calculated column
  
  '******************************************************************************
  ' TM20010719 Fault 2242 - ColumnType <> 4 clause added to ignore all linked   *
  ' columns. (It doesn't need to validate the linked columns because this is    *
  ' done using the Vaidate SP.                                                  *
  '******************************************************************************

  strSQL2 = "SELECT ASRSysTables.TableName+'.'+ASRSysColumns.ColumnName as 'TableColumn', " & _
            "'Mandatory' as Reason " & _
           "FROM ASRSysColumns " & _
           "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID " & _
           "WHERE ASRSysColumns.TableID IN (" & strTablesInDef & ") " & _
           " AND ASRSysColumns.ColumnID NOT IN (" & strColumnsInDef & ") " & _
           " AND " & SQLWhereMandatoryColumn
           '" AND Mandatory = '1'" & _
           " AND CalcExprID = 0 " & _
           " AND ColumnType <> 4 " & _
           " AND Rtrim(DefaultValue) = '' AND Convert(int,dfltValueExprID) = 0 "


  'This will ensure matching data types and
  'compatable column sizes
  strSQL3 = "SELECT FromTable.TableName+'.'+FromColumn.ColumnName+' / '+ToTable.TableName+'.'+ToColumn.ColumnName as 'TableColumn', " & _
            "'Type or Size' as Reason " & _
            "FROM ASRSysDataTransferColumns " & _
            "JOIN ASRSysColumns FromColumn ON FromColumn.ColumnID = ASRSysDataTransferColumns.FromColumnID " & _
            "JOIN ASRSysTables FromTable ON FromTable.TableID = FromColumn.TableID " & _
            "JOIN ASRSysColumns ToColumn ON ToColumn.ColumnID = ASRSysDataTransferColumns.ToColumnID " & _
            "JOIN ASRSysTables ToTable ON ToTable.TableID = ToColumn.TableID " & _
            "WHERE ASRSysDataTransferColumns.DataTransferID = " & CStr(mlngSelectedID) & _
            "  AND (FromColumn.DataType <> ToColumn.DataType OR " & _
                   "FromColumn.Size > ToColumn.Size OR " & _
                   "FromColumn.Decimals > ToColumn.Decimals)"
  
  
  strSQL = strSQL1 & " UNION " & strSQL2 & " UNION " & strSQL3 & _
           " ORDER BY 'TableColumn'"

  Set rsTemp = datData.OpenRecordset(strSQL, adOpenDynamic, adLockReadOnly)


  fOK = (rsTemp.EOF)
  If Not fOK Then

    strErrorColumns = vbNullString
    While Not rsTemp.EOF
      strErrorColumns = strErrorColumns & _
          rsTemp!Reason & " :" & vbTab & _
          rsTemp!TableColumn & vbCrLf
      rsTemp.MoveNext
    Wend
  
    mstrStatusMessage = "Unable to run this Data Transfer as the following " & _
                        "column definitions have changed:" & vbCrLf & vbCrLf & _
                        strErrorColumns
  End If
  
  rsTemp.Close
  Set rsTemp = Nothing

End Sub


Private Function GetPicklistFilterSelect() As String

  Dim rsTemp As Recordset

  GetPicklistFilterSelect = vbNullString

  'If mlngSingleRecordID > 0 Then
  '  GetPicklistFilterSelect = CStr(mlngSingleRecordID)
  
  If mstrPicklistFilterIDs <> "" Then
    GetPicklistFilterSelect = mstrPicklistFilterIDs


  ElseIf mlngPicklistID > 0 Then
    
    mstrStatusMessage = IsPicklistValid(mlngPicklistID)
    If mstrStatusMessage <> vbNullString Or _
       (GetPickListField(mlngPicklistID, "Access") = "HD" And Not mbDefinitionOwner) Then
      If mstrStatusMessage = vbNullString Then
        mstrStatusMessage = "The picklist used in this definition has been made " & _
                          "hidden by another user."
      End If
      fOK = False
      Exit Function
    End If
    
    
    'Get List of IDs from Picklist
    Set rsTemp = datData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngPicklistID, adOpenForwardOnly, adLockReadOnly)
    fOK = Not (rsTemp.BOF And rsTemp.EOF)

    If Not fOK Then
      mstrStatusMessage = "The base table picklist contains no records."
    Else
      Do While Not rsTemp.EOF
        GetPicklistFilterSelect = GetPicklistFilterSelect & _
            IIf(Len(GetPicklistFilterSelect) > 0, ", ", "") & rsTemp.Fields(0)
        rsTemp.MoveNext
      Loop
    End If

    rsTemp.Close
    Set rsTemp = Nothing

  ElseIf mlngFilterID > 0 Then
    
    mstrStatusMessage = IsFilterValid(mlngFilterID)
    If mstrStatusMessage <> vbNullString Or _
       (GetExprField(mlngFilterID, "Access") = "HD" And Not mbDefinitionOwner) Then
      If mstrStatusMessage = vbNullString Then
        mstrStatusMessage = "The filter used in this definition has been made " & _
                          "hidden by another user."
      End If
      fOK = False
      Exit Function
    End If
    
    'Get list of IDs from Filter
    fOK = datGeneral.FilteredIDs(mlngFilterID, GetPicklistFilterSelect)

    ' Generate any UDFs that are used in this filter
    If fOK Then
      datGeneral.FilterUDFs mlngFilterID, mastrUDFsRequired()
    End If

    If Not fOK Then
      ' Permission denied on something in the filter.
      mstrStatusMessage = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngFilterID) & "' filter."
    End If

  End If

End Function


Private Function ValidSQL(strSQL, strErrorMsg) As Boolean

  Dim ADOErr As ADODB.Error
  Dim lngAffected As Long
  Dim intOldCursorLocation As Integer
  Dim rsTemp As Recordset

  On Local Error GoTo LocalErr

  ' Change the cursor location to 'client' as the errors that might be raised
  ' during the update cannot be read for 'server' cursors.
  intOldCursorLocation = gADOCon.CursorLocation
  gADOCon.Errors.Clear
  gADOCon.CursorLocation = adUseClient

  Set rsTemp = gADOCon.Execute(strSQL, lngAffected, adCmdText)

  ' Restore the original cursor location to the ADO connection object.
  gADOCon.CursorLocation = intOldCursorLocation

  
  ' Check if the update produced any errors.
  strErrorMsg = ""
  If gADOCon.Errors.Count > 0 Then
    
    For Each ADOErr In gADOCon.Errors
      strErrorMsg = strErrorMsg & _
            IIf(Len(strErrorMsg) > 0, vbCrLf, vbNullString) & _
            ADOErr.Description
    Next ADOErr
          
    gADOCon.Errors.Clear
  
    ValidSQL = False
  
  Else
  
    'The stored procedure will produce an error if
    'selected record is not found therefore a record
    'must have been transferred !
    ValidSQL = True

  End If

Exit Function

LocalErr:
  'This will occur if the current ID
  'can not be found on the from table
  ValidSQL = False

End Function


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
  mstrStatusMessage = "Error reading record description" & vbCr & _
                      "(ID = " & CStr(lngRecordID) & ", Record Description = " & CStr(mlngRecordDescExprID)
  fOK = False

End Function
