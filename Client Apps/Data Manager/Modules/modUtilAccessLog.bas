Attribute VB_Name = "modUtilAccessLog"
Option Explicit

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

  Dim datData As clsDataAccess
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strHostName As String

  Set datData = New clsDataAccess

  strSQL = "SELECT * FROM ASRSysUtilAccessLog " & _
           "WHERE UtilID = " & CStr(lngID) & _
           " AND Type = " & CStr(utlType)
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

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
  Set datData = Nothing

End Sub


Public Sub DeleteUtilAccessLog(utlType As UtilityType, lngID As Long)

  Dim strSQL As String

  strSQL = "DELETE FROM ASRSYSUtilAccessLog " & _
           "WHERE UtilID = " & CStr(lngID) & _
           " AND Type = " & CStr(utlType)
  
  gADOCon.Execute strSQL

End Sub
