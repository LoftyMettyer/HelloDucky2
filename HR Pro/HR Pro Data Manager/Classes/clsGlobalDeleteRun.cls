VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobalDeleteRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mdatData As clsDataAccess
Private mstrStatusMessage As String
Private fOK As Boolean
Private mblnBatchMode As Boolean
Private mlngSelectedID As Long
Private mblnUserCancelled As Boolean
Private mblnNoRecords As Boolean

Private mstrTableName As String
Private mlngTableID As Long
Private mstrPicklistFilter As String
Private mstrFunctionName As String

Private mstrTableViews() As String


Private Sub Class_Initialize()
  Set mdatData = New HrPro.clsDataAccess
End Sub

Private Sub Class_Terminate()
  Set mdatData = Nothing
End Sub

Public Property Get ErrorString() As String
  ErrorString = mstrStatusMessage
End Property

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property

Public Function RunGlobalDelete(lngSelectedID As Long, blnBatchMode As Boolean) As Boolean
  
  Dim plngEventLogID As Long
  
  On Error GoTo LocalErr
  
  fOK = True
  Screen.MousePointer = vbHourglass
  
  mblnBatchMode = blnBatchMode
  mlngSelectedID = lngSelectedID
  
  If fOK Then Call GetDefinition
  If fOK Then Call BuildTableViewArray
  If fOK Then Call InitialiseProgressBar
  If fOK Then Call ProcessRecords
  
TidyAndExit:
  Call OutputJobStatus
  RunGlobalDelete = fOK

Exit Function

LocalErr:
  mstrStatusMessage = Err.Description
  fOK = False
  Resume TidyAndExit

End Function


Private Sub OutputJobStatus()
  
  On Error Resume Next
  
  Call UtilUpdateLastRun(utlGlobalDelete, mlngSelectedID)
  
  
  gobjEventLog.AddHeader eltGlobalDelete, mstrFunctionName

  If fOK Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful
    mstrStatusMessage = "Completed Successfully"
  
    If mblnNoRecords Then
      mstrStatusMessage = mstrStatusMessage & vbCrLf & _
          "No records meet selection criteria"
      gobjEventLog.AddDetailEntry mstrStatusMessage
    End If
  
  ElseIf mblnUserCancelled Then
    gobjEventLog.ChangeHeaderStatus elsCancelled
    mstrStatusMessage = "Cancelled by user"
  
  Else
    gobjEventLog.ChangeHeaderStatus elsFailed
    gobjEventLog.AddDetailEntry mstrStatusMessage
    mstrStatusMessage = "Failed" & vbCrLf & vbCrLf & mstrStatusMessage
  
  End If
  
  mstrStatusMessage = "Global Delete : " & mstrFunctionName & " " & mstrStatusMessage
  
  If Not mblnBatchMode Then
    If gobjProgress.Visible Then
      gobjProgress.CloseProgress
    End If
    Screen.MousePointer = vbDefault
    MsgBox mstrStatusMessage, IIf(fOK, vbInformation, vbExclamation), "Global Delete"
  Else
    gobjProgress.ResetBar2
  End If

End Sub


Private Sub GetDefinition()

  Dim rsTemp As Recordset
  Dim strSQL As String
  
  On Error GoTo LocalErr
  
  ' Get the Global Delete definition.
  strSQL = "SELECT * FROM ASRSysGlobalFunctions " & _
           "WHERE functionID = " & mlngSelectedID
  Set rsTemp = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  If rsTemp.BOF And rsTemp.EOF Then
    Set rsTemp = Nothing
    mstrStatusMessage = "This definition has been deleted by another user"
    fOK = False
    Exit Sub
  End If
  
  mstrFunctionName = rsTemp!Name
  mlngTableID = rsTemp!TableID
  mstrTableName = datGeneral.GetTableName(mlngTableID)
  mstrPicklistFilter = GetPicklistFilterSelect(rsTemp!PickListID, rsTemp!FilterID)

  rsTemp.Close
  Set rsTemp = Nothing
  
Exit Sub

LocalErr:
  mstrStatusMessage = "Error retrieving definition"
  fOK = False

End Sub


Private Function GetPicklistFilterSelect(lngPicklistID As Long, lngFilterID As Long) As String

  Dim rsTemp As Recordset

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

    If Not fOK Then
      ' Permission denied on something in the filter.
      mstrStatusMessage = "You do not have permission to use the '" & datGeneral.GetFilterName(lngFilterID) & "' filter."
    End If

  End If

End Function


Private Sub BuildTableViewArray()

  Dim objTableView As CTablePrivilege
  Dim iNextIndex As Integer
  
  On Error GoTo LocalErr
  
  
  ReDim mstrTableViews(1, 0)
  
  ' Populate the array with the names of the table/views which the user has permission to delete from.
  If gcoTablePrivileges.Item(mstrTableName).AllowDelete Then
    ReDim Preserve mstrTableViews(1, 1)
    mstrTableViews(1, 1) = gcoTablePrivileges.Item(mstrTableName).RealSource
  Else
    ' The current user does not have permission to delete from the table,
    ' so find any views on this table which the user does have 'delete' permission on.
    For Each objTableView In gcoTablePrivileges.Collection
      If (Not objTableView.IsTable) And _
          (objTableView.TableID = mlngTableID) And _
          (objTableView.AllowDelete) Then

        iNextIndex = UBound(mstrTableViews, 2) + 1
        ReDim Preserve mstrTableViews(1, iNextIndex)
        mstrTableViews(1, iNextIndex) = objTableView.ViewName

      End If
    Next

    If UBound(mstrTableViews, 2) = 0 Then
      ' No 'delete' permission on the table or any views.
      mstrStatusMessage = "You do not have 'delete' permission on the '" & mstrTableName & "' table."
      fOK = False
    End If

  End If
  
  Set objTableView = Nothing

Exit Sub

LocalErr:
  mstrStatusMessage = "Error checking table privileges"
  fOK = False

End Sub


Private Sub InitialiseProgressBar()

  On Error GoTo LocalErr
  
  With gobjProgress
    .AviFile = App.Path & "\videos\filedel.avi"
    If mblnBatchMode = False Then
      .NumberOfBars = 1
      .Caption = "Global Delete"
      .Time = False
      .Cancel = True
      .Bar1MaxValue = UBound(mstrTableViews, 2)
      .Bar1Caption = "Global Delete : " & mstrFunctionName
      .OpenProgress
    Else
      .ResetBar2
      .Bar2MaxValue = UBound(mstrTableViews, 2)
      .Bar2Caption = "Global Delete : " & mstrFunctionName
    End If
  End With

Exit Sub

LocalErr:
  mstrStatusMessage = "Error initialising progress bar"
  fOK = False

End Sub


Private Sub ProcessRecords()

  Dim objDeadlock As clsDeadlock
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim iNextIndex As Integer

  On Error GoTo LocalErr

  Set objDeadlock = New clsDeadlock
  mblnNoRecords = True

  ' Loop through the array of tables/views, performing the global delete on each one.
  For iNextIndex = 1 To UBound(mstrTableViews, 2)

    'First check how many record match the selection criteria so that
    'we can check if no record match !!!
    strSQL = "SELECT COUNT(*) FROM " & mstrTableViews(1, iNextIndex)
    If mstrPicklistFilter <> vbNullString Then
      strSQL = strSQL & " WHERE ID IN (" & mstrPicklistFilter & ")"
    End If
    Set rsTemp = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

    If rsTemp.Fields(0).Value > 0 Then
      mblnNoRecords = False

      'Now actually delete the records
      strSQL = "DELETE FROM " & mstrTableViews(1, iNextIndex)
      If mstrPicklistFilter <> vbNullString Then
        strSQL = strSQL & " WHERE ID IN (" & mstrPicklistFilter & ")"
      End If
      'mdatData.ExecuteSql (strSQL)
      objDeadlock.UpdateTableRecordJustDoIt (strSQL)

    End If
    
    Set rsTemp = Nothing

    
    gobjProgress.UpdateProgress mblnBatchMode
  
    If gobjProgress.Cancelled Then
      mblnUserCancelled = True
      fOK = False
      Exit Sub
    End If
    
  Next iNextIndex

  Set objDeadlock = Nothing

Exit Sub

LocalErr:
  mstrStatusMessage = "Error deleting records"
  fOK = False

End Sub
