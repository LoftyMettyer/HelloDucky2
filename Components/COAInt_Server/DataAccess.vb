Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports System.Security
Imports HR.Intranet.Server.Structures

Public Class clsDataAccess

	Const _CommandTimeOut = 600

	Private ReadOnly _objLogin As LoginInfo
	Private ReadOnly _connectionString As String
	Private ReadOnly _sqlCredential As SqlCredential

	Public Sub New()
		MyBase.New()
	End Sub

	Public Sub New(value As LoginInfo)
		_objLogin = value
		_connectionString = GetConnectionString(_objLogin)

		Dim objPassword As SecureString
		objPassword = value.Password.ToSecureString()
		objPassword.MakeReadOnly()

		If Not value.TrustedConnection Then
			_sqlCredential = New SqlCredential(value.Username, objPassword)
		End If

	End Sub

	Public Sub New(connectionString As String)
		_connectionString = connectionString
	End Sub

	Public ReadOnly Property Login As LoginInfo
		Get
			Return _objLogin
		End Get
	End Property

	Public Sub ChangePassword(objLogin As LoginInfo, sNewPassword As String)

		Dim strConn As String = GetConnectionString(Login)

		Try
			Dim newPassword = sNewPassword.ToSecureString
			newPassword.MakeReadOnly()

			Dim currentPassword = objLogin.Password.ToSecureString
			currentPassword.MakeReadOnly()

			Dim objCredential = New SqlCredential(objLogin.Username, currentPassword)

			SqlConnection.ChangePassword(strConn, objCredential, newPassword)

		Catch ex As Exception
			Throw

		End Try

	End Sub

	Public Sub ExecuteSql(sSQL As String)
		' Execute the given SQL statement.

		Try

			Using sqlConnection As New SqlConnection(_connectionString, _sqlCredential)
				Using objCommand = New SqlCommand(sSQL, sqlConnection)

					objCommand.CommandType = CommandType.Text
					objCommand.CommandTimeout = _CommandTimeOut

					objCommand.Parameters.Clear()
					sqlConnection.Open()
					objCommand.ExecuteNonQuery()
				End Using

			End Using

		Catch
			Throw

		End Try

	End Sub

	Private Shared Function GetConnectionString(LoginDetail As LoginInfo) As String

		Const _AppName As String = "OpenHR Web"
		Const _ConnectionTimeOut As String = "10"

		If LoginDetail.TrustedConnection Then
			Return String.Format("Data Source={0};Initial Catalog={1};Trusted_Connection=yes;Application Name={2};Connection Timeout={3}" _
													 , LoginDetail.Server, LoginDetail.Database, _AppName, _ConnectionTimeOut)

		Else
			Return String.Format("Data Source={0};Initial Catalog={1};Application Name={2};Connection Timeout={3}" _
													 , LoginDetail.Server, LoginDetail.Database, _AppName, _ConnectionTimeOut)

		End If

	End Function

	Public Function GetFromSP(ProcedureName As String, ParamArray args() As SqlParameter) As DataTable

		Try
			Dim dsData = GetDataSet(ProcedureName, CommandType.StoredProcedure, args)

			If dsData.Tables.Count > 0 Then
				Return dsData.Tables(0)
			Else
				Return Nothing
			End If

		Catch ex As Exception
			Throw

		End Try

	End Function

	Public Sub ExecuteSP(ProcedureName As String, ParamArray args() As SqlParameter)

		Dim retryCount = 5
		Dim success As Boolean = False

		Using sqlConnection As New SqlConnection(_connectionString, _sqlCredential)

			Using objCommand = New SqlCommand(ProcedureName, sqlConnection)

				objCommand.CommandType = CommandType.StoredProcedure
				objCommand.CommandTimeout = _CommandTimeOut

				objCommand.Parameters.Clear()
				For Each sqlParm As SqlParameter In args
					objCommand.Parameters.Add(sqlParm)
				Next

				sqlConnection.Open()

				While retryCount > 0 AndAlso Not success

					Try
						objCommand.ExecuteNonQuery()
						success = True

					Catch exception As SqlException

						' SQL Deadlock exception
						If exception.Number <> 1205 Then
							Throw
						End If

						' Add delay here if you wish. 
						retryCount -= 1
						If retryCount = 0 Then
							Throw
						End If

					Catch ex As Exception
						Throw

					End Try

				End While

			End Using

		End Using

	End Sub

	Public Function GetDataTable(sProcedureName As String) As DataTable

		Try
			Return GetDataTable(sProcedureName, CommandType.Text)

		Catch ex As Exception
			Throw

		End Try

	End Function

	Public Function GetDataTable(sProcedureName As String, CommandType As CommandType, ParamArray args() As SqlParameter) As DataTable

		Try
			Dim dtData = GetDataSet(sProcedureName, CommandType, args)

			If dtData.Tables.Count = 0 Then
				Return Nothing
			Else
				Return dtData.Tables(0)
			End If

		Catch ex As Exception
			Throw

		End Try

	End Function

	Public Function GetDataTable(procedureName As String, parameterName As String, dataList As DataTable) As DataTable

		Dim objDataSet As New DataSet
		Dim objAdaptor As New SqlDataAdapter

		Try

			Using sqlConnection As New SqlConnection(_connectionString, _sqlCredential)
				objAdaptor.SelectCommand = New SqlCommand(procedureName, sqlConnection)
				objAdaptor.SelectCommand.CommandType = CommandType.StoredProcedure
				objAdaptor.SelectCommand.CommandTimeout = _CommandTimeOut

				objAdaptor.SelectCommand.Parameters.Clear()

				Dim objParameter As SqlParameter = objAdaptor.SelectCommand.Parameters.AddWithValue(parameterName, dataList)
				objParameter.SqlDbType = SqlDbType.Structured

				objAdaptor.Fill(objDataSet)

			End Using


		Catch ex As Exception
			Throw

		End Try

		Return objDataSet.Tables(0)

	End Function

	Public Function CallToStoredProcedure(sProcedureName As String, ParamArray args() As SqlParameter) As String


		Using sqlConnection As New SqlConnection(_connectionString, _sqlCredential)
			Dim sqlCommand As New SqlCommand(sProcedureName, sqlConnection)
			For Each sqlParm In args
				sqlCommand.Parameters.Add(sqlParm)
			Next

			Dim DECLAREs As String = ""
			Dim EXEC As String = ""
			Dim SELECTs As String = ""

			'Declare output variables
			For Each p As SqlParameter In sqlCommand.Parameters
				If p.Direction = ParameterDirection.Output Then
					If DECLAREs = "" Then
						DECLAREs = "DECLARE " & Environment.NewLine
					End If
					If p.SqlDbType = SqlDbType.NVarChar Or p.SqlDbType = SqlDbType.VarChar Then
						DECLAREs &= "     " & p.ParameterName & " " & p.SqlDbType.ToString & "(MAX), " & Environment.NewLine
					Else
						DECLAREs &= "     " & p.ParameterName & " " & p.SqlDbType.ToString & ", " & Environment.NewLine
					End If
				End If
			Next
			If Not String.IsNullOrEmpty(DECLAREs) Then
				DECLAREs = DECLAREs.Substring(0, DECLAREs.LastIndexOf(", "))
			End If

			'EXEC
			EXEC = Environment.NewLine & "EXEC " & sqlCommand.CommandText & " " & Environment.NewLine
			For Each p As SqlParameter In sqlCommand.Parameters
				If p.SqlDbType = SqlDbType.Decimal Or p.SqlDbType = SqlDbType.Float Or p.SqlDbType = SqlDbType.Int Or p.SqlDbType = SqlDbType.Money Or p.SqlDbType = SqlDbType.Real Then
					If p.Direction = ParameterDirection.Output Then
						EXEC &= "     " & p.ParameterName & " = " & p.ParameterName & " OUTPUT, " & Environment.NewLine
					Else
						EXEC &= "     " & p.ParameterName & " = " & p.Value.ToString & ", " & Environment.NewLine
					End If
				Else
					If p.Direction = ParameterDirection.Output Then
						EXEC &= "     " & p.ParameterName & " = " & p.ParameterName & " OUTPUT, " & Environment.NewLine
					Else
						EXEC &= "     " & p.ParameterName & " = '" & p.Value.ToString & "', " & Environment.NewLine
					End If

				End If
			Next
			If EXEC <> "" Then
				EXEC = EXEC.Substring(0, EXEC.LastIndexOf(", "))
			End If

			'SELECTs
			For Each p As SqlParameter In sqlCommand.Parameters
				If p.Direction = ParameterDirection.Output Then
					If SELECTs = "" Then
						SELECTs = Environment.NewLine & "SELECT " & Environment.NewLine
					End If
					SELECTs &= "     " & p.ParameterName & " AS N'" & p.ParameterName & "', " & Environment.NewLine
				End If
			Next
			If Not String.IsNullOrEmpty(SELECTs) Then
				SELECTs = SELECTs.Substring(0, SELECTs.LastIndexOf(", "))
			End If

			Return DECLAREs & Environment.NewLine & EXEC & Environment.NewLine & SELECTs & Environment.NewLine

		End Using
	End Function

	Public Function GetDataSet(sProcedureName As String, ParamArray args() As SqlParameter) As DataSet
		Return GetDataSet(sProcedureName, CommandType.StoredProcedure, args)
	End Function

	Private Function GetDataSet(sProcedureName As String, CommandType As CommandType, ParamArray args() As SqlParameter) As DataSet

		Dim objDataSet As New DataSet
		Dim objAdaptor As New SqlDataAdapter

		Const RetryThreshold = 5
		Dim iRetryCount As Integer = 0
		Dim bRetry As Boolean = True

		Do While bRetry

			Try

				Using sqlConnection As New SqlConnection(_connectionString, _sqlCredential)

					objAdaptor.SelectCommand = New SqlCommand(sProcedureName, sqlConnection)
					objAdaptor.SelectCommand.CommandType = CommandType
					objAdaptor.SelectCommand.CommandTimeout = _CommandTimeOut

					objAdaptor.SelectCommand.Parameters.Clear()
					For Each sqlParm In args
						objAdaptor.SelectCommand.Parameters.Add(sqlParm)
					Next

					objAdaptor.Fill(objDataSet)

				End Using

				bRetry = False

			Catch ex As Exception

				' TODO Certain errors we should just try again, deadlocking for example?
				'			 others bomb immediately
				bRetry = False

				iRetryCount += 1
				If iRetryCount > RetryThreshold Or Not bRetry Then Throw

			End Try

		Loop

		Return objDataSet

	End Function

  public Function GetModuleSetting(ModuleKey As String, ParameterKey As String, ParameterType As String) As String

 		Dim prmValue = New SqlParameter("parameterValue", SqlDbType.NVarChar, -1) With {.Direction = ParameterDirection.Output}

    Try

      ExecuteSP("spsys_getmodulesetting" _
			  , New SqlParameter("moduleKey", SqlDbType.VarChar, 50) With {.Value = ModuleKey} _
			  , New SqlParameter("parameterKey", SqlDbType.VarChar, 50) With {.Value = ParameterKey} _
			  , New SqlParameter("paramterType", SqlDbType.VarChar, 50) With {.Value = ParameterType} _
			  , prmValue)

      Return prmValue.Value.ToString()

    Catch ex As Exception
      Return ""

    End Try

  End Function

End Class