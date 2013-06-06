Attribute VB_Name = "modDiary"
Option Explicit

Public Sub DiaryRebuild()

  Dim datData As clsDataAccess
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim fOK As Boolean
  Dim strErrorString As String

  Dim objTables As CTablePrivilege
  Dim objColumns As CColumnPrivileges
  
  Dim strMBText As String
  Dim intMBButtons As Integer
  Dim strMBTitle As String
  Dim intMBResponse As Integer
  
  On Error GoTo LocalErr
  fOK = True
  
  
  strMBText = "You are about to rebuild all system diary events." & vbCrLf & _
              "Do you wish to continue ?"
  intMBButtons = vbYesNo + vbQuestion + vbDefaultButton1
  strMBTitle = "Confirm Rebuild"
  intMBResponse = MsgBox(strMBText, intMBButtons, strMBTitle)

  If intMBResponse <> vbYes Then
    Exit Sub
  End If
  
  
  Set datData = New clsDataAccess
  
  With gobjProgress
    '.AviFile = App.Path & "\videos\diary.avi"
    .AVI = dbDiary
    .MainCaption = "Diary"
    .NumberOfBars = 1
    .Caption = "Rebuilding system diary events"
    .Time = False
    .Cancel = True
    .OpenProgress
  End With

  gobjEventLog.AddHeader eltDiaryRebuild, "Diary Rebuild"

  strSQL = "SELECT tableName, tableID From ASRSysTables"
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenKeyset, adLockReadOnly)
  gobjProgress.Bar1MaxValue = rsTemp.RecordCount

  
  strSQL = vbNullString
  Do While Not rsTemp.EOF And gobjProgress.Cancelled = False

    If TableHasDiaryLinks(rsTemp.Fields(1).Value) Then

      Set objTables = gcoTablePrivileges.Item(rsTemp!TableName)
      Set objColumns = GetColumnPrivileges(rsTemp.Fields(0).Value)
      If objColumns("ID").AllowSelect Then
        If strSQL <> vbNullString Then
          strSQL = strSQL & "UNION" & vbCrLf
        End If
        strSQL = strSQL & "SELECT ID, " & CStr(rsTemp!TableID) & " as 'TableID', '" & rsTemp!TableName & "' as 'TableName' FROM [" & objTables.RealSource & "]" & vbCrLf
      End If

    End If

    rsTemp.MoveNext
  Loop

  rsTemp.Close
  Set rsTemp = Nothing

  Set objColumns = Nothing
  Set objTables = Nothing


  If strSQL <> vbNullString Then
    strSQL = strSQL & " ORDER BY 'TableID'"
    Set rsTemp = datData.OpenRecordset(strSQL, adOpenKeyset, adLockReadOnly)
    gobjProgress.Bar1Value = 0
    gobjProgress.Bar1MaxValue = rsTemp.RecordCount
  
    Do While Not rsTemp.EOF And gobjProgress.Cancelled = False
      gobjProgress.Bar1Caption = "Rebuilding system diary events for " & rsTemp.Fields(2).Value
      strSQL = "EXEC spASRDiary_" & CStr(rsTemp.Fields(1).Value) & " " & CStr(rsTemp.Fields(0).Value)
      datData.ExecuteSql (strSQL)
      gobjProgress.UpdateProgress False
      rsTemp.MoveNext
    Loop
  
    rsTemp.Close
    Set rsTemp = Nothing
  End If

  Set datData = Nothing

  If gobjProgress.Cancelled Then
    strErrorString = "Diary Rebuild cancelled by user"
    gobjEventLog.ChangeHeaderStatus elsCancelled
    fOK = False
  Else
    gobjEventLog.ChangeHeaderStatus elsSuccessful
  End If
  

TidyAndExit:
  On Local Error Resume Next
  gobjProgress.CloseProgress

  If fOK Then
    MsgBox "Diary Rebuild Complete", vbInformation, "Diary Rebuild"
  Else
    MsgBox strErrorString, vbExclamation, "Diary Rebuild"
  End If

Exit Sub


LocalErr:
  strErrorString = "Diary Rebuild failed." & vbCrLf & Err.Description
  gobjEventLog.ChangeHeaderStatus elsFailed
  fOK = False
  gobjProgress.CloseProgress
  Resume TidyAndExit

End Sub


Private Function TableHasDiaryLinks(lngTableID As Long) As Boolean

  Dim rsInfo As ADODB.Recordset
  Dim sSQL As String
    
  sSQL = "SELECT COUNT(*) FROM ASRSysDiaryLinks " & _
         "WHERE ColumnID IN " & _
         "(SELECT ColumnID FROM ASRSysColumns WHERE TableID = " & CStr(lngTableID) & ")"
  
  Set rsInfo = datGeneral.GetRecords(sSQL)
  TableHasDiaryLinks = (rsInfo.Fields(0).Value > 0)
  rsInfo.Close
  Set rsInfo = Nothing

End Function
