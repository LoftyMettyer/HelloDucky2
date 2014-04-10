Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Structures

Public Class AccessLog

	Private ReadOnly DB As New clsDataAccess

	Public Sub New(LoginInfo As LoginInfo)
		DB = New clsDataAccess(LoginInfo)
	End Sub

	Public Sub UtilCreated(utlType As UtilityType, lngID As Integer)
		Dim strSQL As String
		strSQL = "INSERT ASRSysUtilAccessLog (Type, UtilID, CreatedBy, CreatedDate, CreatedHost, SavedBy, SavedDate, SavedHost) VALUES '" _
			& utlType & "', " & CStr(lngID) & ", " & " system_user, getdate(), host_name(), system_user, getdate(), host_name())"
		DB.ExecuteSql(strSQL)
	End Sub

	Public Sub UtilUpdateLastSaved(utlType As UtilityType, lngID As Integer)
		UpdateUserAndDate("Saved", utlType, lngID)
	End Sub
	
	Public Sub UtilUpdateLastRun(utlType As UtilityType, lngID As Integer)
		UpdateUserAndDate("Run", utlType, lngID)
	End Sub

	Private Sub UpdateUserAndDate(strMode As String, utlType As UtilityType, lngID As Integer)

		Dim rsTemp As DataTable
		Dim strSQL As String

		strSQL = "SELECT * FROM ASRSysUtilAccessLog WHERE UtilID = " & CStr(lngID) & " AND Type = " & CStr(utlType)
		rsTemp = DB.GetDataTable(strSQL)

		'Have to do this to catch existing utilities !
		If rsTemp.Rows.Count = 0 Then
			strSQL = "INSERT ASRSysUtilAccessLog (Type, UtilID, " & strMode & "By, " & strMode & "Date, " & strMode & "Host) VALUES (" & "'" & utlType & "', " & CStr(lngID) & ", " & "system_user, getdate(), host_name() )"
		Else
			strSQL = "UPDATE ASRSysUtilAccessLog SET " & strMode & "By = system_user, " & strMode & "Date = getdate(), " & strMode & "Host = host_name() " & "WHERE UtilID = " & CStr(lngID) & " AND Type = " & CStr(utlType)
		End If
		DB.ExecuteSql(strSQL)

	End Sub

End Class
