VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobalRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mdatData As clsDataAccess
Private mstrStatusMessage As String
Private fOK As Boolean
Private mlngSelectedID As Long
'Private gblnBatchMode As Boolean
'Private mblnGlobalAdd As Boolean
Private mintType As Integer
Private mstrType As String
Private mblnUserCancelled As Boolean

Private mlngSuccessCount As Long
Private mlngFailCount As Long

Private mlngTableID As Long
Private mstrTableName As String
Private mlngChildID As Long
Private mstrChildName As String
Private mlngPicklistID As Long
Private mlngFilterID As Long
Private mstrFunctionName As String
Private mlngRecordDescExprID As Long
Private mstrBaseRealSource As String

Private mstrColumns() As String
Private mstrTableViews() As String
Private mlngRecordIDs() As Long
Private mbBypassTrigger As Boolean

Private mbDefinitionOwner As Boolean
Private mbLogSuccessfulRecords As Boolean

' Array holding the User Defined functions that are needed for this update
Private mvarUDFsRequired() As String

Private mlngTableViews() As Long
Private mstrSQLJoin As String

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property

Private Sub Class_Initialize()
  Set mdatData = New HRProDataMgr.clsDataAccess
End Sub

Private Sub Class_Terminate()
  Set mdatData = Nothing
End Sub


Public Property Get ErrorString() As String
  ErrorString = mstrStatusMessage
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


Public Function RunGlobal(lngSelectedID As Long, intType As GlobalType) As Boolean

  On Error GoTo LocalErr

  mintType = intType
  'mblnGlobalAdd = (intType = glAdd)
  'mstrType = IIf(mintType = glAdd, "Global Add", "Global Update")
  mstrType = "Global " & Choose(mintType, "Add", "Update", "Delete")

  fOK = True
  mlngSelectedID = lngSelectedID

  Screen.MousePointer = vbHourglass

  If fOK Then Call GetDefinition

  If mintType <> glDelete Then
    'TM20010807 Fault 2656
'    If fOK Then Call CheckHiddenElements
    If fOK Then Call CheckIfFunctionIsValid
    If fOK Then Call BuildColumnArray
  End If

  If fOK Then Call BuildTableViewArray
  If fOK Then Call BuildRecordIDArray
  If fOK Then Call InitialiseProgressBar
  
  If fOK Then Call UDFFunctions(mvarUDFsRequired, True)
  
  Select Case mintType
  Case glAdd
    mbLogSuccessfulRecords = CBool(GetUserSetting("LogEvents", "Global_Add_Success", False))
    gobjEventLog.AddHeader eltGlobalAdd, mstrFunctionName
    If Progress Then Call AddRecords
    Call UtilUpdateLastRun(UtlGlobalAdd, mlngSelectedID)
  
  Case glUpdate
    mbLogSuccessfulRecords = CBool(GetUserSetting("LogEvents", "Global_Update_Success", False))
    gobjEventLog.AddHeader eltGlobalUpdate, mstrFunctionName
    If Progress Then Call UpdateRecords
    Call UtilUpdateLastRun(utlGlobalUpdate, mlngSelectedID)
  
  Case glDelete
    mbLogSuccessfulRecords = CBool(GetUserSetting("LogEvents", "Global_Delete_Success", False))
    gobjEventLog.AddHeader eltGlobalDelete, mstrFunctionName
    If Progress Then Call DeleteRecords
    Call UtilUpdateLastRun(utlGlobalDelete, mlngSelectedID)
  
  End Select

  If fOK Then Call UDFFunctions(mvarUDFsRequired, False)

TidyAndExit:
  Call OutputJobStatus
  
  RunGlobal = fOK

Exit Function

LocalErr:
  mstrStatusMessage = Err.Description
  fOK = False
  Resume TidyAndExit

End Function

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
 
'******* Check calculations for hidden/deleted elements *******

  If Len(mstrStatusMessage) = 0 Then
    sSQL = "SELECT * FROM ASRSysGlobalItems " & _
             "WHERE FunctionID = " & mlngSelectedID & _
             " AND ValueType = 4 "
  
    Set rsTemp = datGeneral.GetRecords(sSQL)
    With rsTemp
      If Not (.EOF And .BOF) Then
        .MoveFirst
        Do Until .EOF
          'JPD 20030812 Fault 6670
          iResult = ValidateCalculation(!ExprID)
          Select Case iResult
            Case REC_SEL_VALID_DELETED
              mstrStatusMessage = "A calculation used in this definition has been deleted by another user."
            Case REC_SEL_VALID_INVALID
              mstrStatusMessage = "A calculation used in this definition is invalid."
            Case REC_SEL_VALID_HIDDENBYOTHER
              If Not gfCurrentUserIsSysSecMgr Then
                mstrStatusMessage = "A calculation used in this definition has been made hidden by another user."
              End If
          End Select
          
          If Len(mstrStatusMessage) > 0 Then
            Exit Do
          End If
          
          .MoveNext
        Loop
      End If
    End With
    
    Set rsTemp = Nothing
  End If
  
  IsRecordSelectionValid = (Len(mstrStatusMessage) = 0)
  
End Function

Private Sub OutputJobStatus()
  
  Dim blnNoRecords As Boolean
  
  If mlngFailCount > 0 And Not mblnUserCancelled Then
    fOK = False
  End If

  blnNoRecords = (mlngSuccessCount = 0 And mlngFailCount = 0)

  With gobjEventLog
  
    If fOK Then
      .ChangeHeaderStatus elsSuccessful, mlngSuccessCount, mlngFailCount
      mstrStatusMessage = "Completed successfully."
    
      If blnNoRecords Then
        mstrStatusMessage = mstrStatusMessage & vbNewLine & _
                            "No records meet selection criteria."
        .AddDetailEntry mstrStatusMessage
        fOK = True
      End If
  
    ElseIf mblnUserCancelled Then
      .ChangeHeaderStatus elsCancelled, mlngSuccessCount, mlngFailCount
      mstrStatusMessage = "Cancelled by user."
    Else
      .ChangeHeaderStatus elsFailed, mlngSuccessCount, mlngFailCount
      If mstrStatusMessage <> vbNullString Then
        'Only details records for failures !
        .AddDetailEntry mstrStatusMessage
        mstrStatusMessage = "Failed." & vbNewLine & vbNewLine & mstrStatusMessage
      Else
        mstrStatusMessage = "failed" & vbNewLine
      End If
    End If
  
  End With
  
  
  If Not blnNoRecords Then
    mstrStatusMessage = mstrStatusMessage & vbNewLine & _
      Records(mlngSuccessCount) & Choose(mintType, " added", " updated", " deleted")

    If mlngFailCount > 0 Then
      mstrStatusMessage = mstrStatusMessage & vbNewLine & _
        Records(mlngFailCount) & " failed during " & Choose(mintType, "add", "update", "delete")
    End If
  End If
  
  
  mstrStatusMessage = mstrType & " : '" & mstrFunctionName & "' " & mstrStatusMessage
  
  
  If Not gblnBatchMode Then
    gobjProgress.CloseProgress
    Screen.MousePointer = vbDefault
    MsgBox mstrStatusMessage, IIf(fOK, vbInformation, vbExclamation), mstrType
  Else
    gobjProgress.ResetBar2
  End If

End Sub


Private Sub GetDefinition()

  On Error GoTo LocalErr
  
  Dim rsTemp As Recordset
  Dim strSQL As String

  ' Get the details of the given Global Function
  strSQL = "SELECT ASRSysGlobalFunctions.*, " & _
           "       ASRSysTables.RecordDescExprID " & _
           "FROM ASRSysGlobalFunctions " & _
           "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysGlobalFunctions.TableID " & _
           "WHERE functionID = " & mlngSelectedID
  Set rsTemp = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  If rsTemp.BOF And rsTemp.EOF Then
    Set rsTemp = Nothing
    mstrStatusMessage = "This definition has been deleted by another user."
    fOK = False
    Exit Sub
  End If
  
  mstrFunctionName = rsTemp!Name
  
  mlngTableID = rsTemp!TableID
  mstrTableName = datGeneral.GetTableName(mlngTableID)
  mstrBaseRealSource = gcoTablePrivileges.Item(mstrTableName).RealSource
  mlngPicklistID = rsTemp!PicklistID
  mlngFilterID = rsTemp!FilterID
  mlngRecordDescExprID = rsTemp!RecordDescExprID
  mbBypassTrigger = IIf(IsNull(rsTemp!BypassTrigger), False, rsTemp!BypassTrigger)
  
  If mintType = glAdd Then
    mlngChildID = rsTemp!ChildTableID
    mstrChildName = datGeneral.GetTableName(mlngChildID)
  End If
    
  mbDefinitionOwner = (LCase(Trim(gsUserName)) = LCase(Trim(rsTemp!UserName)))

  fOK = IsRecordSelectionValid

  rsTemp.Close
  Set rsTemp = Nothing
  
Exit Sub

LocalErr:
  mstrStatusMessage = "Error retrieving definition"
  fOK = False

End Sub
Private Sub CheckHiddenElements()

'Sub created as part of fix for Fault 2656.
  
  Dim sSQL As String
  Dim bShowMSG As Boolean
  Dim rsFunction As ADODB.Recordset
  Dim rsTemp As ADODB.Recordset
  Dim sMessage As String
  Dim sText As String
  
  On Error GoTo ErrorTrap
  
  bShowMSG = False

  sSQL = "SELECT * FROM ASRSysGlobalFunctions WHERE FunctionID = " & mlngSelectedID
  Set rsTemp = datGeneral.GetRecords(sSQL)

  'Check for hidden picklists.
  If rsTemp!PicklistID Then
    sText = IsPicklistValid(rsTemp!PicklistID)
    If sText <> vbNullString Then
      MsgBox "You cannot run this Global definition as it contains a hidden picklist which has been deleted or made hidden by another user." & vbNewLine & _
             "Please re-visit your definition to remove the hidden picklist.", vbExclamation, App.Title
      bShowMSG = True
      fOK = False
      GoTo TidyUpAndExit:
    End If
  End If

  'Check if Primary filter is hidden.
  If rsTemp!FilterID Then
    sText = IsFilterValid(rsTemp!FilterID)
    If sText <> vbNullString Then
      MsgBox "You cannot run this Global definition as it contains a hidden filter which has been deleted or made hidden by another user." & vbNewLine & _
             "Please re-visit your definition to remove the hidden filter.", vbExclamation, App.Title
      bShowMSG = True
      fOK = False
      GoTo TidyUpAndExit:
    End If
  End If

  sSQL = "SELECT * FROM ASRSysGlobalItems WHERE FunctionID = " & mlngSelectedID
  Set rsFunction = datGeneral.GetRecords(sSQL)
  
  With rsFunction
    If .RecordCount > 0 Then
      .MoveFirst
      Do Until .EOF
        'If the the column type in the mail merge is an expression then check the expression
        'for hidden components / deleted components.
        If !ValueType = 4 Then
          'If the expression has hidden components and is owned by another user or has been deleted then notify the user.
          sMessage = IsCalcValid(rsFunction!ExprID)
          If sMessage <> vbNullString Then
            If Not bShowMSG Then
              mstrStatusMessage = "You cannot run this Global definition as it contains one or more calculation(s) which have been deleted or made hidden by another user." & vbNewLine & _
                     "Please re-visit your definition to remove the hidden calculation(s)." & vbNewLine
              bShowMSG = True
            End If
          End If
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With

  fOK = Not bShowMSG
  
TidyUpAndExit:
  Set rsFunction = Nothing
  Exit Sub

ErrorTrap:
  fOK = False
  MsgBox "Error validating the Global definition.", vbOKOnly + vbExclamation, App.Title
  Resume TidyUpAndExit
  
End Sub
Private Sub CheckIfFunctionIsValid()

  'This sub will check (where possible) that the definition is valid and
  'if any invalid columns are found then the destination column and reason
  'why it is not valid will be reported to the user and the function aborted.
  
  'This check must be done at run time in case any column definitions have
  'changed since the global function was last amended.

  'Criteria which is checked in this subroutine:
  'Mandatory Columns (Global Add Only)
  'Read Only Columns
  'Parent Table Columns Data Type and Size match destination column
  
  'NOTE: Straight values and lookup table values are validated in
  '      the subroutine CheckStraightValues
  
  
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strSQLReadOnly As String
  Dim strSQLMandatory As String
  Dim strSQLParentColumns As String
  Dim strTablesInDef As String
  Dim strColumnsInDef As String
  Dim strErrorColumns As String

  On Error GoTo LocalErr
  
  
  strColumnsInDef = "SELECT ASRSysGlobalItems.ColumnID " & _
                   "FROM ASRSysGlobalItems " & _
                   "WHERE ASRSysGlobalItems.FunctionID = " & CStr(mlngSelectedID)

  strTablesInDef = "SELECT DISTINCT ASRSysColumns.TableID " & _
                   "FROM ASRSysColumns " & _
                   "WHERE ASRSysColumns.ColumnID IN (" & strColumnsInDef & ")"


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

  strSQLMandatory = _
      "SELECT ASRSysTables.TableName+'.'+ASRSysColumns.ColumnName as 'TableColumn', " & _
      "'Mandatory' as Reason " & _
      "FROM ASRSysColumns " & _
      "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID " & _
      "WHERE ASRSysColumns.TableID IN (" & strTablesInDef & ") " & _
      " AND ASRSysColumns.ColumnID NOT IN (" & strColumnsInDef & ") " & _
      " AND " & SQLWhereMandatoryColumn
      '" AND Mandatory = '1'" & _
      " AND CalcExprID = 0 " & _
      " AND ColumnType <> 4 " & _
      " AND Rtrim(DefaultValue) = '' AND Convert(int,dfltValueExprID) = 0"
  
  'This will retreive all of the read only
  'destination columns in the definition
  strSQLReadOnly = _
     "SELECT ASRSysTables.TableName+'.'+ASRSysColumns.ColumnName as 'TableColumn', " & _
     "'Read Only' as Reason " & _
     "FROM ASRSysGlobalItems " & _
     "JOIN ASRSysColumns ON ASRSysColumns.ColumnID = ASRSysGlobalItems.ColumnID " & _
     "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID " & _
     "WHERE ASRSysGlobalItems.FunctionID = " & CStr(mlngSelectedID) & _
     " AND ASRSysColumns.ReadOnly <> 0"

  'Parent Table Columns
  strSQLParentColumns = _
     "SELECT ASRSysTables.TableName+'.'+DestinCol.ColumnName as 'TableColumn', " & _
     "'Data Type or Size' as Reason " & _
     "FROM ASRSysGlobalItems " & _
     "JOIN ASRSysColumns DestinCol ON DestinCol.ColumnID = ASRSysGlobalItems.ColumnID " & _
     "JOIN ASRSysColumns SourceCol ON SourceCol.ColumnID = ASRSysGlobalItems.RefColumnID " & _
     "JOIN ASRSysTables ON ASRSysTables.TableID = DestinCol.TableID " & _
     "WHERE ASRSysGlobalItems.FunctionID = " & CStr(mlngSelectedID) & _
     "  AND ASRSysGlobalItems.ValueType = 3 " & _
     "  AND (DestinCol.DataType <> SourceCol.DataType OR " & _
     "       DestinCol.Size < SourceCol.Size OR " & _
     "       DestinCol.Decimals < SourceCol.Decimals)"
  
  strSQL = strSQLReadOnly & " UNION " & _
           strSQLParentColumns & _
           " ORDER BY 'TableColumn'"

  If mintType = glAdd Then
    'Don't bother checking mandatory columns
    'for global update as all of the required
    'columns will already be populated
    strSQL = strSQLMandatory & " UNION " & strSQL
  End If

  Set rsTemp = mdatData.OpenRecordset(strSQL, adOpenDynamic, adLockReadOnly)
  
  While Not rsTemp.EOF
    strErrorColumns = strErrorColumns & _
        rsTemp!Reason & " :" & vbTab & _
        rsTemp!TableColumn & vbNewLine
    rsTemp.MoveNext
  Wend
  
  rsTemp.Close
  Set rsTemp = Nothing


  
  'Need to check straight values in VB as this was awkward in SQL!
  strErrorColumns = strErrorColumns & CheckStraightValues

  
  fOK = (strErrorColumns = vbNullString)
  
  If Not fOK Then
    mstrStatusMessage = _
        "Unable to run this " & mstrType & _
        " as the following column definitions are incorrect:" & vbNewLine & vbNewLine & _
        strErrorColumns
  End If
  
  
Exit Sub

LocalErr:
  mstrStatusMessage = "Error validating definition"
  fOK = False

End Sub


Private Function CheckStraightValues() As String

  'This sub will validate all straight values which are to be inserted/updated.
  '(this also includes values from lookup tables).
  
  
  Dim rsTemp As Recordset
  Dim rsLookup As Recordset
  Dim strSQL As String
  Dim strErrorColumns As String
  Dim dblTemp As Double
  Dim blnValid As Boolean

  Dim strValue As String
  Dim lngSize As Long
  Dim lngDecimals As Long


  'Straight values
  'strSQL = _
     "SELECT ASRSysGlobalItems.Value, " & _
     "       ASRSysColumns.DataType, " & _
     "       ASRSysColumns.Size, " & _
     "       ASRSysColumns.Decimals, " & _
     "       ASRSysColumns.ColumnName, " & _
     "       ASRSysTables.TableName, " & _
     "       LookUpTable.TableName as LookupTableName, " & _
     "       LookUpColumn.ColumnName as LookupColumnName, " & _
     "       ASRSysGlobalTableValues.RecordID as LookupRecordID, " & _
     "       ASRSysGlobalItems.ValueType " & _
     "FROM ASRSysGlobalItems " & _
     "JOIN ASRSysColumns ON ASRSysColumns.ColumnID = ASRSysGlobalItems.ColumnID " & _
     "JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID " & _
     "JOIN ASRSysGlobalTableValues ON ASRSysGlobalTableValues.TableValueID = ASRSysGlobalItems.TableValueID " & _
     "JOIN ASRSysTables LookUpTable ON LookUpTable.TableID = ASRSysGlobalTableValues.TableID " & _
     "JOIN ASRSysColumns LookUpColumn ON LookUpColumn.ColumnID = ASRSysGlobalTableValues.ColumnID " & _
     "WHERE ASRSysGlobalItems.FunctionID = " & CStr(mlngSelectedID) & _
     "  AND (ASRSysGlobalItems.ValueType = 1 OR ASRSysGlobalItems.ValueType = 2)"
  
  strSQL = _
     "SELECT ASRSysGlobalItems.Value, " & _
     "       ASRSysColumns.DataType, " & _
     "       ASRSysColumns.Size, " & _
     "       ASRSysColumns.Decimals, " & _
     "       ASRSysColumns.ColumnName, " & _
     "       ASRSysTables.TableName, " & _
     "       ASRSysGlobalItems.ValueType " & _
     "FROM ASRSysGlobalItems " & _
     "JOIN ASRSysColumns ON ASRSysColumns.ColumnID = ASRSysGlobalItems.ColumnID " & _
     "JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID " & _
     "WHERE ASRSysGlobalItems.FunctionID = " & CStr(mlngSelectedID) & _
     "  AND (ASRSysGlobalItems.ValueType = 1 OR ASRSysGlobalItems.ValueType = 2)"
  
  Set rsTemp = mdatData.OpenRecordset(strSQL, adOpenDynamic, adLockReadOnly)
  
  
  Do While Not rsTemp.EOF
  
    
'    'Get insert value
'    If rsTemp!ValueType = 2 Then
'      'Lookup table value
'      strSQL = "SELECT " & rsTemp!LookupColumnName & _
'               " FROM " & rsTemp!LookupTableName & _
'               " WHERE ID = " & CStr(rsTemp!LookupRecordID)
'      Set rsLookup = mdatData.OpenRecordset(strSQL, adOpenDynamic, adLockReadOnly)
'      strValue = Trim$(rsLookup.Fields(0))
'      rsLookup.Close
'      Set rsLookup = Nothing
'    Else
'      'Straight value
    'strValue = IIf(IsNull(rsTemp!Value), vbNullString, Trim(rsTemp!Value))
    strValue = IIf(IsNull(rsTemp!Value), vbNullString, RTrim(rsTemp!Value))
'    End If
    
    blnValid = True
    
    
    Select Case rsTemp!DataType
    Case sqlNumeric, sqlInteger
      'If not value and not zero
      strValue = datGeneral.ConvertNumberForSQL(strValue)
      If Val(strValue) = 0 And Left$(Val(strValue), 1) <> "0" Then
        blnValid = False
      Else
        'Check size and decimals...
        If rsTemp!DataType = sqlNumeric Then
          lngSize = Len(CStr(Int(Val(strValue))))
          'lngDecimals = Len(CStr(Abs(Val(strValue) - Int(Val(strValue)))))
          lngDecimals = Len(strValue) - (lngSize + 1)
          'take off two chars for the "0." at the beginning
          lngDecimals = IIf(lngDecimals <= 1, 0, lngDecimals)
          If lngSize > rsTemp!Size Or lngDecimals > rsTemp!Decimals Then
            blnValid = False
          End If
        End If
      End If
    
    Case sqlDate
      'If not null and not a date
      If Not IsDate(rsTemp!Value) And Not IsNull(rsTemp!Value) Then
        blnValid = False
      End If
    
    Case sqlBoolean
      If UCase(strValue) <> "TRUE" And UCase(strValue) <> "FALSE" Then
        blnValid = False
      End If
    
    Case Else
      'If Len(strValue) > rsTemp!Size Then
      If rsTemp!Size > 0 And Len(strValue) > rsTemp!Size Then
        blnValid = False
      End If

    End Select
    
    
    If blnValid = False Then
      strErrorColumns = strErrorColumns & _
        "Type or Size :" & vbTab & _
        rsTemp!TableName & "." & rsTemp!ColumnName & vbNewLine
    End If
    
    rsTemp.MoveNext
  Loop

  rsTemp.Close
  Set rsTemp = Nothing

  CheckStraightValues = strErrorColumns

End Function


Private Sub BuildColumnArray()

  Dim rsTemp As Recordset
  Dim rsLookup As Recordset
  Dim strSQL As String
  Dim strSource As String
  Dim iNextIndex As Integer
  Dim iTrimming As TrimmingTypes

  On Error GoTo LocalErr
  
  ' Dimension an array to hold the names of the columns, and the new values.
  ' Column 1 = column name
  ' Column 2 = new value
  ' Column 3 = value type
  ReDim mstrColumns(3, 0)
  ReDim mvarUDFsRequired(0)

  ReDim mlngTableViews(0)
  mstrSQLJoin = vbNullString

  ' Get the parameters of the given Global Update.
  'strSQL = "EXEC sp_ASRGetGlobalUpdateDetails " & mlngSelectedID

  'strSQL = "SELECT ASRSysGlobalItems.valueType, " & _
                  "ASRSysGlobalItems.ExprID, " & _
                  "ASRSysGlobalItems.Value, " & _
                  "ASRSysColumns.columnName, " & _
                  "ASRSysColumns1.columnName AS RefColName, " & _
                  "ASRSysTables.TableName AS TableValueName, " & _
                  "ASRSysColumns2.columnName AS TableColName, " & _
                  "ASRSysGlobalTableValues.RecordID, " & _
                  "ASRSysColumns.DataType " & _
           "FROM ASRSysTables " & _
           "INNER JOIN ASRSysGlobalTableValues " & _
           "   ON ASRSysTables.tableID = ASRSysGlobalTableValues.tableID " & _
           "INNER JOIN ASRSysColumns " & _
           "   ON ASRSysGlobalItems.ColumnID = ASRSysColumns.columnID " & _
           "   ON ASRSysGlobalTableValues.tableValueID = ASRSysGlobalItems.tableValueID " & _
           "LEFT OUTER JOIN ASRSysColumns ASRSysColumns1 " & _
           "   ON ASRSysGlobalItems.refColumnID = ASRSysColumns1.columnID " & _
           "INNER JOIN ASRSysColumns ASRSysColumns2 " & _
           "   ON ASRSysGlobalTableValues.columnID = ASRSysColumns2.columnID " & _
           "RIGHT OUTER JOIN ASRSysGlobalItems " & _
           "WHERE ASRSysGlobalItems.functionID = " & CStr(mlngSelectedID) & " " & _
           "ORDER BY ASRSysGlobalItems.valueType"

  strSQL = "SELECT ASRSysGlobalItems.*, " & _
           "       ASRSysColumns.ColumnName, " & _
           "       ASRSysColumns.DataType, " & _
           "       ASRSysColumns1.columnName AS RefColName, " & _
           "       ASRSysColumns.Trimming " & _
           "FROM ASRSysGlobalItems " & _
           "INNER JOIN ASRSysColumns " & _
           "   ON ASRSysGlobalItems.ColumnID = ASRSysColumns.columnID " & _
           "LEFT OUTER JOIN ASRSysColumns ASRSysColumns1 " & _
           "   ON ASRSysGlobalItems.refColumnID = ASRSysColumns1.columnID " & _
           "WHERE ASRSysGlobalItems.FunctionID = " & CStr(mlngSelectedID) & " " & _
           "ORDER BY ASRSysGlobalItems.valueType"

  
  Set rsTemp = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)


  Do While Not rsTemp.EOF
    
    ' Character trimming type
    iTrimming = IIf(Not IsNull(rsTemp!Trimming), rsTemp!Trimming, giTRIMMING_NONE)
    
    Select Case rsTemp!ValueType
    Case globfuncvaltyp_STRAIGHTVALUE, globfuncvaltyp_LOOKUPTABLE
      ' Update the current field with a Straight Value.
      If (rsTemp!DataType = sqlVarChar) Or _
        (rsTemp!DataType = sqlLongVarChar) Then
        'TM20020910 Fault 4395 - Don't RTrim the value.
'        strSource = "'" & Replace(RTrim(rsTemp!Value), "'", "''") & "'"
        If (rsTemp!DataType = sqlLongVarChar) Then
          strSource = RTrim(rsTemp!Value)
        Else
          strSource = rsTemp!Value
        End If
        
        ' Does this data need to be trimmed
'        Select Case iTrimming
'          Case giTRIMMING_LEFTRIGHT
'            strSource = LTrim(RTrim(strSource))
'          Case giTRIMMING_LEFTONLY
'            strSource = LTrim(strSource)
'          Case giTRIMMING_RIGHTONLY
'            strSource = RTrim(strSource)
'          Case Else
'            strSource = strSource
'        End Select
        
        strSource = "'" & Replace(strSource, "'", "''") & "'"

      ElseIf rsTemp!DataType = sqlDate Then
        If IsNull(rsTemp!Value) Then
          strSource = "null"
        Else
          strSource = "'" & Trim(rsTemp!Value) & "'"
        End If
      ElseIf rsTemp!DataType = sqlBoolean Then
        strSource = IIf(UCase(Trim(rsTemp!Value)) = "TRUE", "1", "0")
      Else
        strSource = Val(datGeneral.ConvertNumberForSQL(Trim(rsTemp!Value)))
      End If

'    Case globfuncvaltyp_LOOKUPTABLE
'      ' Update the current field with a LOOKUP table value.
'      strSQL = "SELECT " & rsTemp!TableColName & _
'        " FROM " & rsTemp!TableValueName & _
'        " WHERE id = " & rsTemp!RecordID
'      Set rsLookup = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
'
'      If rsLookup.BOF And rsLookup.EOF Then
'        mstrStatusMessage = mstrFunctionName & " failed." & vbnewline & _
'            "The table value from '" & rsLookup!TableValueName & "." & _
'                          rsLookup!TableColName & "' no longer exists."
'        fOK = False
'        Exit Sub
'      End If
'
'      If rsTemp!DataType = sqlVarChar Then
'        strSource = "'" & Replace(RTrim(rsLookup(0)), "'", "''") & "'"
'      ElseIf rsTemp!DataType = sqlDate Then
'        strSource = "'" & Format(rsLookup(0), "mm/dd/yyyy") & "'"
'      Else
'        strSource = Trim(Str(rsLookup(0)))
'      End If
'
    Case globfuncvaltyp_FIELD
      ' Update the current field with a parent table FIELD value.
      'strSource = mstrTableName & "." & rsTemp!RefColName
      'strSource = rsTemp!RefColName


  Dim pintTemp As Integer
  Dim pobjTableView As CTablePrivilege
  Dim pobjColumnPrivileges As CColumnPrivileges
  Dim pblnAllowed As Boolean
  Dim pblnInsertOnly As Boolean
  Dim strTableViewName As String
  Dim strColumnRef As String
  Dim blnViews As Boolean
  
  
  ' Loop through each table/view in the collection
  strSource = vbNullString
  blnViews = False
  For Each pobjTableView In gcoTablePrivileges.Collection

    'MH20050210 Fault 9799
    'If pobjTableView.TableID = mlngTableID And pobjTableView.AllowInsert Then
    If pobjTableView.TableID = mlngTableID Then
      If (mintType = glAdd And pobjTableView.AllowInsert) Or _
         (mintType = glUpdate And pobjTableView.AllowUpdate) Then

        ' Get a list of column privileges for the objtableview
        strTableViewName = IIf(pobjTableView.IsTable, pobjTableView.TableName, pobjTableView.ViewName)
        Set pobjColumnPrivileges = GetColumnPrivileges(strTableViewName)

        'JPD 20050225 Fault 9834
        If pobjColumnPrivileges.IsValid(rsTemp!RefColName) Then
        
          If pobjColumnPrivileges(rsTemp!RefColName).AllowUpdate Then
  
            strColumnRef = strTableViewName & "." & rsTemp!RefColName
            If pobjTableView.IsTable Then
              strSource = pobjTableView.RealSource & "." & rsTemp!RefColName
              Exit For
            Else
              strSource = strSource & " WHEN NOT " & strColumnRef & " IS NULL THEN " & strColumnRef
              AddToViewArray pobjTableView.ViewID, pobjTableView.ViewName
              blnViews = True
            End If
          End If
        End If
      End If
    
    End If

  Next

  If strSource <> vbNullString And blnViews Then
    strSource = "SELECT CASE" & strSource & " ELSE NULL END " & _
                "FROM " & mstrBaseRealSource & " " & _
                mstrSQLJoin & " " & _
                "WHERE " & mstrBaseRealSource & ".ID = "
  End If
    
    Case globfuncvaltyp_CALCULATION
      ' do here  and check permissions.
      strSource = GetCalculationCode(rsTemp!ExprID)

    
    End Select
    
    ' Put the column name and the new value into the array
    iNextIndex = UBound(mstrColumns, 2) + 1
    ReDim Preserve mstrColumns(3, iNextIndex)
    mstrColumns(1, iNextIndex) = rsTemp!ColumnName
    mstrColumns(2, iNextIndex) = strSource
    mstrColumns(3, iNextIndex) = rsTemp!ValueType
   
    
    rsTemp.MoveNext
  Loop


'  'MH20010711 If no columns specified then we are doing a global update
'  'just to update column calculations... SO PICK A COLUMN... ANY COLUMN...
'  If iNextIndex = 0 Then
'
'    strSQL = "SELECT TOP 1 ASRSysColumns.ColumnName " & _
'             "FROM ASRSysColumns WHERE readonly = 0 and tableid = " & CStr(mlngTableID)
'    Set rsTemp = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
'
'    If Not rsTemp.EOF Then
'      iNextIndex = UBound(mstrColumns, 2) + 1
'      ReDim Preserve mstrColumns(3, iNextIndex)
'      mstrColumns(1, iNextIndex) = rsTemp!ColumnName
'      mstrColumns(2, iNextIndex) = rsTemp!ColumnName
'      mstrColumns(3, iNextIndex) = globfuncvaltyp_FIELD
'    End If
'
'  End If

  rsTemp.Close
  Set rsTemp = Nothing


Exit Sub

LocalErr:
  mstrStatusMessage = "Error reading column details"
  fOK = False

End Sub


Private Sub BuildTableViewArray()

  Dim colColumnPrivileges As CColumnPrivileges
  Dim objTableView As CTablePrivilege

  Dim blnAllowed As Boolean
  Dim iNextIndex As Integer

  On Error GoTo LocalErr

  ReDim mstrTableViews(0) As String

  For Each objTableView In gcoTablePrivileges.Collection

    Select Case mintType
    Case glAdd
      blnAllowed = (objTableView.TableID = mlngChildID And objTableView.AllowInsert)
    Case glUpdate
      blnAllowed = (objTableView.TableID = mlngTableID And objTableView.AllowUpdate)
    Case glDelete
      blnAllowed = (objTableView.TableID = mlngTableID And objTableView.AllowDelete)
    End Select


    If blnAllowed And mintType <> glDelete Then
      
      If objTableView.IsTable Then
        Set colColumnPrivileges = GetColumnPrivileges(objTableView.TableName)
      Else
        Set colColumnPrivileges = GetColumnPrivileges(objTableView.ViewName)
      End If
      
      'Loop through columns checking update access to each
      For iNextIndex = 1 To UBound(mstrColumns, 2)
        ' Check if we can perform each column's update using the table.
        'If found a column that we can't update in this table/view
        'exit for so don't bother checking the rest of the columns
          
        blnAllowed = colColumnPrivileges.IsValid(mstrColumns(1, iNextIndex))
        If Not blnAllowed Then Exit For
          
        blnAllowed = colColumnPrivileges.Item(mstrColumns(1, iNextIndex)).AllowUpdate
        If Not blnAllowed Then Exit For
          
        If mstrColumns(3, iNextIndex) = globfuncvaltyp_FIELD And Not mintType = glAdd Then
          ' Updating from another field. Check if we have read permission on the field.
          blnAllowed = colColumnPrivileges.Item(mstrColumns(1, iNextIndex)).AllowSelect
          If Not blnAllowed Then Exit For
        End If
                  
      Next
        
    End If
      
    If blnAllowed Then
      iNextIndex = UBound(mstrTableViews) + 1
      ReDim Preserve mstrTableViews(iNextIndex)
      mstrTableViews(iNextIndex) = objTableView.RealSource
      If objTableView.IsTable Then
        Exit For
      End If
    End If

  Next

  Set objTableView = Nothing
  Set colColumnPrivileges = Nothing

  
  'check ubound to see if no access on any tables/views
  If UBound(mstrTableViews) = 0 Then
    ' JPD20030219 Faults 5071 & 5072
    mstrStatusMessage = "You do not have permission to " & _
      Choose(mintType, "insert all of the required columns on ", _
      "update all of the required columns on ", "delete from ") & _
      "the '" & IIf(mintType = glAdd, mstrChildName, mstrTableName) & "' table."
    fOK = False
  End If


Exit Sub

LocalErr:
  mstrStatusMessage = "Error checking column privileges"
  fOK = False

End Sub


Private Sub BuildRecordIDArray()

  Dim objTableView As CTablePrivilege
  Dim rsTemp As Recordset
  Dim lngLoop As Long
  Dim strPicklistFilterIDs As String
  Dim lngLastID As Long
  
  Dim strSQL As String
  Dim strSelect As String
  Dim strJoin As String
  
  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim iShowMeAgain As Integer
  
  Dim msb As frmMessageBox
  Dim sGlobalType As String
  
  On Error GoTo LocalErr
  
  
  'Set up an array of record ids to be updated
  ReDim mlngRecordIDs(0) As Long
  
  strPicklistFilterIDs = GetPicklistFilterSelect(mlngPicklistID, mlngFilterID)
  If Not fOK Then
    Exit Sub
  End If
    
  'We need to build a string which will select the ID from all of the views
  'that we have select permission on
  strSelect = vbNullString
  strJoin = vbNullString
  For Each objTableView In gcoTablePrivileges.Collection

    If objTableView.AllowSelect Then
      'strSQL = strSQL & _
             IIf(strSQL <> vbNullString, vbnewline & "UNION" & vbnewline, vbNullString) & _
             "SELECT ID FROM " & objTableView.RealSource & _
             IIf(strPicklistFilterIDs <> vbNullString, _
                          " WHERE ID IN (" & strPicklistFilterIDs & ")", _
                          vbNullString)
      If objTableView.TableID = mlngTableID Then
        strSelect = strSelect & _
          "WHEN NOT [" & objTableView.RealSource & "].ID IS NULL THEN [" & objTableView.RealSource & "].ID " & vbNewLine

        If Not objTableView.IsTable Then
          strJoin = strJoin & _
            "LEFT OUTER JOIN [" & objTableView.RealSource & "] ON [" & mstrBaseRealSource & "].ID = [" & objTableView.RealSource & "].ID " & vbNewLine
        End If

      ElseIf datGeneral.IsAParentOf(objTableView.TableID, mlngTableID) Then
          strJoin = strJoin & _
            "LEFT OUTER JOIN [" & objTableView.RealSource & "] ON [" & mstrBaseRealSource & "].ID_" & CStr(objTableView.TableID) & " = [" & objTableView.RealSource & "].ID " & vbNewLine

'      ElseIf datGeneral.IsAChildOf(objTableView.TableID, mlngTableID) Then
'          strJoin = strJoin & _
'            "LEFT OUTER JOIN [" & objTableView.RealSource & "] ON [" & mstrBaseRealSource & "].ID = [" & objTableView.RealSource & "].ID_" & CStr(mlngTableID) & " " & vbnewline

      End If
    End If

  Next

  If strSelect = vbNullString Then
    mstrStatusMessage = "Error selecting records"
    fOK = False
    Exit Sub
  End If

  'Now get distinct IDs from the base table (as there couuld be duplicates in the above sql!)
  'JPD 20030915 Fault 5038
  'strSQL = "SELECT ID FROM " & mstrBaseRealSource & _
           IIf(strSQL <> vbNullString, " WHERE ID IN (" & strSQL & ")", vbNullString)
  'strSQL = strSQL & " ORDER BY id"
  strSQL = "SELECT CASE " & strSelect & " ELSE NULL END " & vbNewLine & _
           "FROM [" & mstrBaseRealSource & "]" & vbNewLine & _
           strJoin

  If strPicklistFilterIDs <> vbNullString Then
    strSQL = strSQL & _
      "WHERE [" & mstrBaseRealSource & "].ID IN (" & strPicklistFilterIDs & ") " & vbNewLine
  End If

  strSQL = strSQL & "ORDER BY [" & mstrBaseRealSource & "].ID"

  If fOK Then Call UDFFunctions(mvarUDFsRequired, True)
  Set rsTemp = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  If fOK Then Call UDFFunctions(mvarUDFsRequired, False)


  lngLoop = 0
  lngLastID = 0
  With rsTemp
    Do While Not .EOF
      If lngLastID <> rsTemp.Fields(0).Value Then
        lngLoop = lngLoop + 1
        ReDim Preserve mlngRecordIDs(lngLoop) As Long
        mlngRecordIDs(lngLoop) = rsTemp.Fields(0).Value
        lngLastID = rsTemp.Fields(0).Value
      End If
      
      .MoveNext
    Loop
  End With

  rsTemp.Close
  Set rsTemp = Nothing

  'TM20010726 Fault 1607 (Suggestion)
  'Need also to show warning message if the Global is of type Add.
  'Removed clause - mintType <> glAdd And _

  If lngLoop > 0 And Not gblnBatchMode Then

    Screen.MousePointer = vbDefault

    'NHRD20022003 Fault 4270
    'Split the old IF statement into a CASE so that each mintType option
    'can have it's own code section. Before Update and Delete shared messages.
    Select Case mintType
      Case glAdd:
              strMBText = "WARNING: You are about to add " & _
                Records(UBound(mlngRecordIDs)) & _
                " to the '" & mstrChildName & "' table." & vbNewLine & vbNewLine & _
                "Do you wish to continue?"
      Case glUpdate:
              strMBText = "WARNING: You are about to update " & _
                Records(UBound(mlngRecordIDs)) & _
                " from the '" & mstrTableName & "' table and you will NOT be able to undo this process." & vbNewLine & vbNewLine & _
                "It is recommended that a backup of the system be made before any Global Updates are run!" & vbNewLine & vbNewLine & _
                "Do you wish to continue?"
      Case glDelete:
              strMBText = "WARNING: You are about to delete " & _
                Records(UBound(mlngRecordIDs)) & _
                " from the '" & mstrTableName & "' table and you will NOT be able to undo this process." & vbNewLine & vbNewLine & _
                "It is recommended that a backup of the system be made before any Global Deletes are run!" & vbNewLine & vbNewLine & _
                "Do you wish to continue?"
    End Select
    
    intMBButtons = vbExclamation + vbYesNo + vbDefaultButton2
    
    iShowMeAgain = GetUserSetting("warningmsg", "warning " & LCase(Replace(mstrType, " ", "")), 1)
    If iShowMeAgain = 1 Then
      Set msb = New frmMessageBox
      iShowMeAgain = IIf(iShowMeAgain = 1, 0, 1)
      mblnUserCancelled = (msb.MessageBox(strMBText, intMBButtons, mstrType, iShowMeAgain, "&Don't show me this warning again.") <> vbYes)
      SaveUserSetting "warningmsg", "warning " & LCase(Replace(mstrType, " ", "")), IIf(iShowMeAgain = 1, 0, 1)
      Set msb = Nothing
    End If
    
'    mblnUserCancelled = (MsgBox(strMBText, intMBButtons, mstrType) <> vbYes)
    If mblnUserCancelled Then
      fOK = False
      Exit Sub
    End If
    Screen.MousePointer = vbHourglass
  
  End If
  
  
  If Not gblnBatchMode Then
    gobjProgress.Bar1MaxValue = UBound(mlngRecordIDs)
  Else
    gobjProgress.Bar2MaxValue = UBound(mlngRecordIDs)
  End If

Exit Sub

LocalErr:
  mstrStatusMessage = "Error selecting records"
  fOK = False

End Sub


Private Sub InitialiseProgressBar()

  On Error GoTo LocalErr
  
  With gobjProgress

    If mintType = glDelete Then
      '.AviFile = App.Path & "\videos\filedel.avi"
      'NHRD12042002 Fault 3367
      'created new AVI that is a shortened version of the filedelr.avi
      '.AviFile = App.Path & "\videos\fileVaporise.avi"
      .AVI = dbVaporise
    Else
      '.AviFile = App.Path & "\videos\transfer.avi"
      .AVI = dbTransfer
    End If

    .MainCaption = mstrType
    
    If Not gblnBatchMode Then
      .NumberOfBars = 1
      .Caption = mstrType
      .Time = False
      .Cancel = True
      .Bar1Caption = mstrType & " : " & mstrFunctionName
      .Bar1RecordsCaption = "Records Processed "
      .OpenProgress
    Else
      .ResetBar2
      .Bar2Caption = mstrType & " : " & mstrFunctionName
      .Bar2RecordsCaption = "Records Processed :"
    End If
  End With

Exit Sub

LocalErr:
  mstrStatusMessage = "Error initialising progress bar"
  fOK = False

End Sub


Private Sub UpdateRecords()

  Dim objDeadlock As clsDeadlock
  Dim strSQL As String
  Dim lngRec As Long
  Dim iNextIndex As Integer
  Dim blnUpdated As Boolean
  Dim strColumnResult As String
  Dim strUpdateColumns As String
  Dim strRecordError As String
  Dim strOutput As String
  
  On Error GoTo LocalErr

  If mbBypassTrigger Then
    strSQL = "INSERT ASRSysSystemSettings ([Section], [SettingKey], [SettingValue]) VALUES ('database', 'ByPassTrigger_" & mlngTableID & "_SPID' ,@@SPID)"
    gADOCon.Execute strSQL, , adExecuteNoRecords
  End If

  Set objDeadlock = New clsDeadlock

  mlngSuccessCount = 0
  mlngFailCount = 0
  For lngRec = 1 To UBound(mlngRecordIDs)
      
    'Loop through columns building update string for this recordID
    '(Need to be done here to include the record id for each calc)
    strUpdateColumns = vbNullString
    For iNextIndex = 1 To UBound(mstrColumns, 2)

      strColumnResult = mstrColumns(2, iNextIndex)
      
      If mstrColumns(3, iNextIndex) = globfuncvaltyp_CALCULATION Then
        'strColumnResult = strColumnResult & _
            " WHERE " & mstrTableName & ".ID = " & _
            CStr(mlngRecordIDs(lngRec))
        strColumnResult = strColumnResult & _
            " WHERE " & mstrBaseRealSource & ".ID = " & _
            CStr(mlngRecordIDs(lngRec))
      End If

      strUpdateColumns = strUpdateColumns & _
                         IIf(LenB(strUpdateColumns) <> 0, ", ", "") & _
                         mstrColumns(1, iNextIndex) & " = " & _
                         "(" & strColumnResult & _
                         IIf(Right$(strColumnResult, 6) = ".ID = ", CStr(mlngRecordIDs(lngRec)), "") & ")"
    Next


    ' Loop through the array of tables/views, trying to performing the global update
    For iNextIndex = 1 To UBound(mstrTableViews)

      'strSQL = "UPDATE " & mstrTableViews(iNextIndex) & " " & _
               "SET " & strUpdateColumns & " " & _
               "WHERE ID = " & CStr(mlngRecordIDs(lngRec))
      strSQL = "IF EXISTS(SELECT ID FROM " & mstrTableViews(iNextIndex) & vbNewLine & _
               " WHERE ID = " & CStr(mlngRecordIDs(lngRec)) & ")" & vbNewLine & _
               "    UPDATE " & mstrTableViews(iNextIndex) & vbNewLine & _
               "    SET " & strUpdateColumns & vbNewLine & _
               "    WHERE ID = " & CStr(mlngRecordIDs(lngRec)) '& vbNewLine & _
               "ELSE" & vbNewLine & _
               "  RAISERROR('Record not found', 16, 1)"

      'JPD - 11/01/02 - Fault 3346
      blnUpdated = objDeadlock.UpdateTableRecordJustDoIt(strSQL)
      strRecordError = objDeadlock.ErrorString
      If objDeadlock.FatalError Then
        mstrStatusMessage = "Error updating records (" & objDeadlock.ErrorString & ")"
        fOK = False
        Exit Sub
      End If


'MH20070321 Fault 12061
'''      'blnUpdated = ValidSQL(strSQL, strRecordError)
'''      'If blnUpdated Or strRecordError <> vbNullString Then
'''      If blnUpdated Or _
'''        (strRecordError <> vbNullString And strRecordError <> "Record not found") Then
'''        'Updated this record sucessfully so don't bother
'''        'trying to update this record on any other views
'''        Exit For
'''      End If
      
      'Updated this record sucessfully so don't bother
      'trying to update this record on any other views
      If blnUpdated Then
        Exit For
      End If
    
    Next iNextIndex
    
    
    If blnUpdated Then
      
      'JDM - 12/12/01 - Fault 3280 - Log successful records
      If mbLogSuccessfulRecords Then
        strRecordError = GetRecordDesc(mlngRecordIDs(lngRec)) & " Updated successfully"
        Call gobjEventLog.AddDetailEntry(strRecordError)
      End If
      
      mlngSuccessCount = mlngSuccessCount + 1

    ElseIf strRecordError <> vbNullString Then

      strRecordError = GetRecordDesc(mlngRecordIDs(lngRec)) & vbNewLine & vbNewLine & _
                       strRecordError

      Call gobjEventLog.AddDetailEntry(strRecordError)
      mlngFailCount = mlngFailCount + 1
      'fOK = False
    
    'Else You can't see the record whose ID we just tried to update
    
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
      GoTo TidyUpAndExit
    End If
  
  Next

  Set objDeadlock = Nothing

  'MH20021014
  objEmail.SendImmediateEmails

TidyUpAndExit:

  If mbBypassTrigger Then
    strSQL = "DELETE ASRSysSystemSettings WHERE [Section] = 'database' AND [SettingKey] = 'ByPassTrigger_" & mlngTableID & "_SPID'"
    gADOCon.Execute strSQL, , adExecuteNoRecords
  End If

'  mstrStatusMessage = _
'      CStr(mlngSuccessCount) & IIf(mlngSuccessCount <> 1, " records", " record") & _
'      " successfully transferred."
'
'  If mlngFailCount > 0 Then
'    mstrStatusMessage = mstrStatusMessage & vbnewline & _
'      CStr(mlngFailCount) & IIf(mlngFailCount <> 1, " records", " record") & _
'      " failed during transfer."
'  End If

  Exit Sub

LocalErr:
  mstrStatusMessage = "Error updating records (" & Err.Description & ")"
  fOK = False
  GoTo TidyUpAndExit

End Sub


Private Sub AddRecords()

  Dim objDeadlock As clsDeadlock
  Dim strSQL As String
  Dim lngRec As Long
  Dim iNextIndex As Integer
  Dim strColumnResult As String
  Dim blnAdded As Boolean
  Dim blnFromColumn As Boolean
  Dim strRecordError As String
  Dim strOutput As String

  Dim strInsertCols As String
  Dim strInsertVals As String

  On Error GoTo LocalErr

  Set objDeadlock = New clsDeadlock

  mlngSuccessCount = 0
  mlngFailCount = 0
  blnFromColumn = False
  For lngRec = 1 To UBound(mlngRecordIDs)
      
    'Loop through columns building update string for this recordID
    '(Need to be done here to include the record id for each calc
    'and so that each child record references the correct parent)
    strInsertCols = "ID_" & CStr(mlngTableID)
    strInsertVals = CStr(mlngRecordIDs(lngRec))

    For iNextIndex = 1 To UBound(mstrColumns, 2)
      strColumnResult = mstrColumns(2, iNextIndex)

      Select Case mstrColumns(3, iNextIndex)
      Case globfuncvaltyp_CALCULATION
        'strColumnResult = "(" & strColumnResult & _
            " WHERE " & mstrTableName & ".ID = " & _
            CStr(mlngRecordIDs(lngRec)) & ")"
        strColumnResult = "(" & strColumnResult & _
            " WHERE " & mstrBaseRealSource & ".ID = " & _
            CStr(mlngRecordIDs(lngRec)) & ")"
      Case globfuncvaltyp_FIELD
        blnFromColumn = True
      End Select

      strColumnResult = strColumnResult & _
            IIf(Right$(strColumnResult, 6) = ".ID = ", CStr(mlngRecordIDs(lngRec)), "")
      
      strInsertCols = strInsertCols & ", " & mstrColumns(1, iNextIndex)
      strInsertVals = strInsertVals & ", (" & strColumnResult & ")"
    Next

    If blnFromColumn Then
      'strInsertVals = strInsertVals & _
            " FROM " & mstrTableName & _
            " WHERE " & mstrTableName & ".ID = " & _
            CStr(mlngRecordIDs(lngRec))
      strInsertVals = strInsertVals & _
            " FROM " & mstrBaseRealSource & vbNewLine & mstrSQLJoin & _
            " WHERE " & mstrBaseRealSource & ".ID = " & _
            CStr(mlngRecordIDs(lngRec))
    End If
    
    
    ' Loop through the array of tables/views, trying to performing the global update
    For iNextIndex = 1 To UBound(mstrTableViews)

      'NOTE: You can not put SELECT statements in the VALUES clause
      'Therefore I have ommitted the VALUES clause and replaced it
      'with another SELECT statement.  You can still specify which
      'columns to insert into using the column list.
      'strSQL = "INSERT " & mstrTableViews(iNextIndex) & _
               "(" & strInsertCols & ") " & _
               "SELECT " & strInsertVals


      'strSQL = "IF EXISTS(SELECT ID FROM " & mstrTableName & vbnewline & _
               " WHERE ID = " & CStr(mlngRecordIDs(lngRec)) & ")" & vbnewline & _
               "    INSERT " & mstrTableViews(iNextIndex) & _
               "    (" & strInsertCols & ") " & _
               "    SELECT " & strInsertVals & vbnewline & _
               "ELSE" & vbnewline & _
               "  RAISERROR('Record not found', 16, 1)"
      strSQL = "IF EXISTS(SELECT ID FROM " & mstrBaseRealSource & vbNewLine & _
               " WHERE ID = " & CStr(mlngRecordIDs(lngRec)) & ")" & vbNewLine & _
               "    INSERT " & mstrTableViews(iNextIndex) & _
               "    (" & strInsertCols & ") " & _
               "    SELECT " & strInsertVals & vbNewLine & _
               "ELSE" & vbNewLine & _
               "  RAISERROR('Record not found', 16, 1)"
      
      'JPD - 11/01/02 - Fault 3346
      blnAdded = objDeadlock.InsertTableRecordJustDoIt(strSQL)
      strRecordError = objDeadlock.ErrorString
      If objDeadlock.FatalError Then
        mstrStatusMessage = "Error adding records (" & objDeadlock.ErrorString & ")"
        fOK = False
        Exit Sub
      End If
      
      'blnAdded = ValidSQL(strSQL, strRecordError)
      If blnAdded Or strRecordError <> vbNullString Then
        'Updated this record sucessfully so don't bother
        'trying to update this record on any other views
        Exit For
      End If

    Next iNextIndex
    
    
    If blnAdded Then
      
      'JDM - 12/12/01 - Fault 3280 - Log successful records
      If mbLogSuccessfulRecords Then
        strRecordError = GetRecordDesc(mlngRecordIDs(lngRec)) & " Added successfully"
        Call gobjEventLog.AddDetailEntry(strRecordError)
      End If
      
      mlngSuccessCount = mlngSuccessCount + 1

    ElseIf strRecordError <> vbNullString Then

      strRecordError = GetRecordDesc(mlngRecordIDs(lngRec)) & vbNewLine & vbNewLine & _
                       strRecordError

      Call gobjEventLog.AddDetailEntry(strRecordError)
      mlngFailCount = mlngFailCount + 1
      'fOK = False
    
    'Else You can't see the record whose ID we just tried to update
    
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
      Exit Sub
    End If
  
  Next

  'If fOK Then
  '  mstrStatusMessage = _
  '      "Completed Successfully." & vbnewline & _
  '      CStr(lngAddCount) & " record" & _
  '      IIf(lngAddCount <> 1, "s", vbNullString) & _
  '      " added"
  'End If

  Set objDeadlock = Nothing

  'MH20021014
  objEmail.SendImmediateEmails

Exit Sub

LocalErr:
  'String or binary data would be truncated.
  If Err.Number = -2147217833 Then
    'Calculation must be too long !!!
    blnAdded = False
    Resume Next
  Else
    mstrStatusMessage = "Error adding records (" & Err.Description & ")"
  End If
  fOK = False

End Sub


Private Function GetPicklistFilterSelect(lngPicklistID As Long, lngFilterID As Long) As String

  Dim rsTemp As Recordset
  Dim sSQL As String
  
  GetPicklistFilterSelect = vbNullString
  
  If lngPicklistID > 0 Then
    
    mstrStatusMessage = IsPicklistValid(lngPicklistID)
    If mstrStatusMessage <> vbNullString Then
      fOK = False
      Exit Function
    End If
    
    'Get List of IDs from Picklist
    Set rsTemp = mdatData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & lngPicklistID, adOpenForwardOnly, adLockReadOnly)
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

  ElseIf lngFilterID > 0 Then
    
    mstrStatusMessage = IsFilterValid(lngFilterID)
    If mstrStatusMessage <> vbNullString Then
      fOK = False
      Exit Function
    End If
    
    'Get list of IDs from Filter
    fOK = datGeneral.FilteredIDs(lngFilterID, GetPicklistFilterSelect)

    ' Generate any UDFs that are used in this filter
    If fOK And gbEnableUDFFunctions Then
      datGeneral.FilterUDFs lngFilterID, mvarUDFsRequired()
    End If

    If Not fOK Then
      ' Permission denied on something in the filter.
      mstrStatusMessage = "You do not have permission to use the '" & datGeneral.GetFilterName(lngFilterID) & "' filter."
    End If

  End If

'******* Check calculations for hidden/deleted elements *******

  sSQL = "SELECT * FROM ASRSysGlobalItems " & _
           "WHERE FunctionID = " & mlngSelectedID & _
           " AND ValueType = '4' "

  Set rsTemp = datGeneral.GetRecords(sSQL)
  With rsTemp
    If Not (.EOF And .BOF) Then
      .MoveFirst
      Do Until .EOF
        mstrStatusMessage = IsCalcValid(!ExprID)
        If mstrStatusMessage <> vbNullString Then
          fOK = False
          .Close
          Set rsTemp = Nothing
          Exit Function
        End If
        .MoveNext
      Loop
    End If
  End With

End Function


Private Function GetCalculationCode(lngExpID As Long) As String

  Dim lngCalcViews() As Long
  Dim objCalcExpr As clsExprExpression
  Dim intCount As Integer
  Dim sCalcCode As String
  Dim sSource As String
  
  Dim strJoinCode As String

  ReDim lngCalcViews(2, 0)
  Set objCalcExpr = New clsExprExpression
  fOK = objCalcExpr.Initialise(mlngTableID, lngExpID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_UNDEFINED)
  If fOK Then
    fOK = objCalcExpr.RuntimeCalculationCode(lngCalcViews, sCalcCode, True)
  
    If fOK And gbEnableUDFFunctions Then
      fOK = objCalcExpr.UDFCalculationCode(lngCalcViews, mvarUDFsRequired(), True)
    End If
  
  End If

  If fOK = False Then
    If Trim(objCalcExpr.Name) = vbNullString Then
      mstrStatusMessage = "This definition contains one or more calculation(s) which" & vbNewLine & _
                          "have been deleted by another user."
    Else
      mstrStatusMessage = "You do not have permission to use the '" & Trim(objCalcExpr.Name) & "' calculation."
    End If
    
    Set objCalcExpr = Nothing
    Exit Function
  End If
  Set objCalcExpr = Nothing

  ' Add the required views to the JOIN code.
  strJoinCode = vbNullString
  For intCount = 1 To UBound(lngCalcViews, 2)
    If lngCalcViews(1, intCount) = 1 Then
      sSource = gcoTablePrivileges.FindViewID(lngCalcViews(2, intCount)).RealSource

      strJoinCode = strJoinCode & vbNewLine & _
        " LEFT OUTER JOIN " & sSource & _
        " ON " & mstrBaseRealSource & ".ID = " & sSource & ".ID"
    
    'TM20020827 Fault 4340 - Add the required tables to the JOIN code.
    ElseIf lngCalcViews(1, intCount) = 0 Then
      ' JPD20021107 Fault 4721
      If datGeneral.IsAParentOf(lngCalcViews(2, intCount), mlngTableID) Then
        sSource = gcoTablePrivileges.FindTableID(lngCalcViews(2, intCount)).RealSource
        
        strJoinCode = strJoinCode & vbNewLine & _
          " LEFT OUTER JOIN " & sSource & _
          " ON " & mstrBaseRealSource & ".ID_" & CStr(lngCalcViews(2, intCount)) & " = " & sSource & ".ID"
      End If
    End If
  Next

  GetCalculationCode = "SELECT " & sCalcCode & _
                       " FROM " & mstrBaseRealSource & _
                       strJoinCode

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


Private Sub DeleteRecords()

  Dim objDeadlock As clsDeadlock
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim iNextIndex As Integer
  Dim lngRec As Long
  Dim strRecordDesc As String
  Dim blnDeleted As Boolean
  Dim strRecordError As String
  Dim strOutput As String

  On Error GoTo LocalErr

  Set objDeadlock = New clsDeadlock

  mlngSuccessCount = 0
  mlngFailCount = 0
  For lngRec = 1 To UBound(mlngRecordIDs)
  
    If mbLogSuccessfulRecords Then
      strRecordDesc = GetRecordDesc(mlngRecordIDs(lngRec))
    End If
  
    ' Loop through the array of tables/views, performing the global delete on each one.
    For iNextIndex = 1 To UBound(mstrTableViews)
  
      'Now actually delete the records
      strSQL = "DELETE FROM " & mstrTableViews(iNextIndex) & _
               " WHERE ID = " & CStr(mlngRecordIDs(lngRec))
      blnDeleted = objDeadlock.UpdateTableRecordJustDoIt(strSQL)
      strRecordError = objDeadlock.ErrorString
      If objDeadlock.FatalError Then
        mstrStatusMessage = "Error updating records (" & objDeadlock.ErrorString & ")"
        fOK = False
        Exit Sub
      End If

      If blnDeleted Then
        Exit For
      End If
    
    Next
    
    
    If blnDeleted Then
      
      'JDM - 12/12/01 - Fault 3280 - Log successful records
      If mbLogSuccessfulRecords Then
        strRecordDesc = strRecordDesc & " Deleted successfully"
        Call gobjEventLog.AddDetailEntry(strRecordDesc)
      End If
      
      mlngSuccessCount = mlngSuccessCount + 1

    ElseIf strRecordError <> vbNullString Then

      strRecordError = GetRecordDesc(mlngRecordIDs(lngRec)) & vbNewLine & vbNewLine & _
                       strRecordError

      Call gobjEventLog.AddDetailEntry(strRecordError)
      mlngFailCount = mlngFailCount + 1
    
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
      Exit For
    End If

  Next

  Set objDeadlock = Nothing

  'JDM - Send immediate emails
  objEmail.SendImmediateEmails

Exit Sub

LocalErr:
  mstrStatusMessage = "Error deleting records"
  fOK = False

End Sub


Private Function AddToViewArray(lngViewID As Long, strViewName As String) As Boolean
  
  Dim lngIndex As Integer
  
  For lngIndex = 1 To UBound(mlngTableViews)
    If mlngTableViews(lngIndex) = lngViewID Then
      Exit Function
    End If
  Next

  'Only get here if not already in array
  lngIndex = UBound(mlngTableViews) + 1
  ReDim Preserve mlngTableViews(lngIndex)
  mlngTableViews(lngIndex) = lngViewID
  
  'JPD 20050315 Fault 9904
  mstrSQLJoin = mstrSQLJoin & _
      "LEFT OUTER JOIN " & strViewName & " ON " & _
      mstrBaseRealSource & ".ID = " & strViewName & ".ID " & vbNewLine

End Function
