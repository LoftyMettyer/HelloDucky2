Attribute VB_Name = "modUtilAccessLog"
Option Explicit

Public Enum UtilityType
  utlBatchJob = 0
  utlCrossTab = 1
  utlCustomReport = 2
  utlDataTransfer = 3
  utlExport = 4
  UtlGlobalAdd = 5
  utlGlobalDelete = 6
  utlGlobalUpdate = 7
  utlImport = 8
  utlMailMerge = 9
  utlPicklist = 10
  utlFilter = 11
  utlCalculation = 12
  utlOrder = 13
  utlMatchReport = 14
  utlAbsenceBreakdown = 15
  utlBradfordFactor = 16
  utlCalendarReport = 17
  utlLabel = 18
  utlLabelType = 19
  utlRecordProfile = 20
  utlEmailAddress = 21
  utlEmailGroup = 22
  utlSuccession = 23
  utlCareer = 24
  utlWorkflow = 25
  utlWorkFlowPendingSteps = 26
  utlOrderDefinition = 27
  utlDocumentMapping = 28
  utlReportPack = 29
  utlTurnover = 30
  utlStability = 31
  utlScreen = 32
  utlTable = 33
  utlColumn = 34
  utlNineBoxGrid = 35
  utlTalent = 38
End Enum


Public Sub UtilCreated(utlType As UtilityType, lngID As Long)

  Dim strSQL As String

  strSQL = "INSERT ASRSysUtilAccessLog " & _
           "(Type, UtilID, CreatedBy, CreatedDate, CreatedHost, SavedBy, SavedDate, SavedHost) " & _
           "VALUES (" & _
           "'" & utlType & "', " & CStr(lngID) & ", " & _
           " system_user, getdate(), host_name(), system_user, getdate(), host_name())"

  gADOCon.Execute strSQL

End Sub


Public Sub UtilUpdateLastSaved(utlType As UtilityType, lngID As Long)
  Call UpdateUserAndDate("Saved", utlType, lngID)
End Sub

Public Sub UtilUpdateLastSavedMultiple(utlType As UtilityType, sIDs As String)
  
  Dim lngIDs As Variant
  Dim intCount As Integer
  
  If InStr(sIDs, ",") > 0 Then
  
    lngIDs = Split(sIDs, ",")
    
    For intCount = LBound(lngIDs) To UBound(lngIDs)
      
      If Trim(lngIDs(intCount)) <> vbNullString Then
        
        Call UpdateUserAndDate("Saved", utlType, CLng(lngIDs(intCount)))
        
      End If
    
    Next
  
  Else
  
    Call UpdateUserAndDate("Saved", utlType, CLng(sIDs))

  End If
  
End Sub


Public Sub UtilUpdateLastRun(utlType As UtilityType, lngID As Long)
  Call UpdateUserAndDate("Run", utlType, lngID)
End Sub


Private Sub UpdateUserAndDate(strMode As String, utlType As UtilityType, lngID As Long)

  Dim rsTemp As New ADODB.Recordset
  Dim strSQL As String
  Dim strHostName As String

  strSQL = "SELECT * FROM ASRSysUtilAccessLog " & _
           "WHERE UtilID = " & CStr(lngID) & _
           " AND Type = " & CStr(utlType)
  rsTemp.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  'Have to do this to catch existing utilities !
  If rsTemp.BOF And rsTemp.EOF Then
    strSQL = "INSERT ASRSysUtilAccessLog " & _
             "(Type, UtilID, " & _
             strMode & "By, " & strMode & "Date, " & strMode & "Host) " & _
             "VALUES (" & _
             "'" & utlType & "', " & CStr(lngID) & ", " & _
             "system_user, getdate(), host_name() )"
  Else
    strSQL = "UPDATE ASRSysUtilAccessLog SET " & _
             strMode & "By = system_user, " & _
             strMode & "Date = getdate(), " & _
             strMode & "Host = host_name() " & _
             "WHERE UtilID = " & CStr(lngID) & _
             " AND Type = " & CStr(utlType)
  End If
  gADOCon.Execute strSQL

  rsTemp.Close
  Set rsTemp = Nothing

End Sub


Public Sub DeleteUtilAccessLog(utlType As UtilityType, lngID As Long)

  Dim strSQL As String

  strSQL = "DELETE FROM ASRSYSUtilAccessLog " & _
           "WHERE UtilID = " & CStr(lngID) & _
           " AND Type = " & CStr(utlType)
  
  gADOCon.Execute strSQL

End Sub
