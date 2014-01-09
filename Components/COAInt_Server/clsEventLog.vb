Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Structures
Imports System.Data.SqlClient

Public Class clsEventLog

	Private ReadOnly DB As New clsDataAccess

	Public Sub New(ByVal LoginInfo As LoginInfo)
		DB = New clsDataAccess(LoginInfo)
	End Sub

	Private mlngEventLogID As Integer

	Public Property EventLogID() As Integer
		Get
			EventLogID = mlngEventLogID
		End Get
		Set(ByVal Value As Integer)
			mlngEventLogID = Value
		End Set
	End Property

	Public Function AddHeader(ByRef udtType As EventLog_Type, ByRef strName As String) As Boolean

		Try

			Dim prmNewID As New SqlParameter("piNewRecordID", SqlDbType.Int)
			prmNewID.Direction = ParameterDirection.Output

			Dim prmType As New SqlParameter("piType", SqlDbType.Int)
			prmType.Value = udtType

			Dim prmName As New SqlParameter("psName", SqlDbType.VarChar, 150)
			prmName.Value = strName

			Dim prmUserName = New SqlParameter("psUserName", SqlDbType.VarChar, 50)
			prmUserName.Value = StrConv(gsUsername, VbStrConv.ProperCase)

			Dim prmBatchName = New SqlParameter("psBatchName", SqlDbType.VarChar, 50)
			prmBatchName.Value = ""

			Dim prmBatchRunID = New SqlParameter("piBatchRunID", SqlDbType.Int)
			prmBatchRunID.Value = 0

			Dim prmBatchJobID = New SqlParameter("piBatchJobID", SqlDbType.Int)
			prmBatchJobID.Value = 0

			DB.ExecuteSP("sp_ASRIntAddEventLogHeader", prmNewID, prmType, prmName, prmUserName, prmBatchName, prmBatchRunID, prmBatchJobID)
			mlngEventLogID = CInt(prmNewID.Value)

		Catch ex As Exception
			mlngEventLogID = 0
			Return False

		End Try

		Return True

	End Function

	Public Function ChangeHeaderStatus(ByRef udtStatus As EventLog_Status, ByVal lngSuccess As Integer, ByVal lngFailed As Integer) As Boolean

		Dim strSQL As String

		Try

			strSQL = String.Format("UPDATE [AsrSysEventLog] SET [EndTime] = GETDATE(), [Duration] = DATEDIFF(second, [DateTime],  GETDATE()), Status = {1}, [SuccessCount] = {2},  [FailCount] = {3} WHERE [ID] = {0}" _
														 , mlngEventLogID, CInt(udtStatus), lngSuccess, lngFailed)
			DB.ExecuteSql(strSQL)
			Return True

		Catch ex As Exception
			Return False

		End Try

	End Function

	Public Function ChangeHeaderStatus(ByRef udtStatus As EventLog_Status) As Boolean

		Dim strSQL As String

		Try

			strSQL = String.Format("UPDATE [AsrSysEventLog] SET [EndTime] = GETDATE(), [Duration] = DATEDIFF(second, [DateTime],  GETDATE()), Status = {0} WHERE [ID] = {1}" _
														 , CInt(udtStatus), mlngEventLogID)
			DB.ExecuteSql(strSQL)
			Return True

		Catch ex As Exception
			Return False

		End Try

	End Function

	Public Function AddDetailEntry(ByRef pstrNotes As String) As Boolean

		Try
			If mlngEventLogID > 0 Then
				DB.ExecuteSql("INSERT INTO AsrSysEventLogDetails (EventLogID, Notes) VALUES (" & mlngEventLogID & ", '" & Replace(pstrNotes, "'", "''") & "')")
			End If

			Return True

		Catch ex As Exception
			Return False

		End Try


	End Function

End Class