Attribute VB_Name = "modVersionInfo"
Option Explicit

Public Function GetOldServerName() As String

  Dim sSQL As String
  Dim rsSQLInfo As ADODB.Recordset

  sSQL = "SELECT @@SERVERNAME"

  Set rsSQLInfo = New ADODB.Recordset
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      GetOldServerName = UCase$(.Fields(0).Value)
    End If
    .Close
  End With
  Set rsSQLInfo = Nothing

End Function

Public Function GetServerName() As String

  Dim sSQL As String
  Dim rsSQLInfo As ADODB.Recordset

  ' AE20090114 Fault #13490
  'sSQL = "SELECT @@SERVERNAME"
  sSQL = "SELECT SERVERPROPERTY('servername')"

  Set rsSQLInfo = New ADODB.Recordset
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      GetServerName = UCase$(.Fields(0).Value)
    End If
    .Close
  End With
  Set rsSQLInfo = Nothing

End Function

Public Function GetDBName()

  Dim sSQL As String
  Dim rsSQLInfo As ADODB.Recordset

  sSQL = "SELECT DB_NAME()"

  Set rsSQLInfo = New ADODB.Recordset
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      GetDBName = UCase$(.Fields(0).Value)
    End If
    .Close
  End With
  Set rsSQLInfo = Nothing

End Function
