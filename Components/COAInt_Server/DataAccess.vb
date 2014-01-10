Option Strict On
Option Explicit On

Imports ADODB
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports HR.Intranet.Server.Structures

Public Class clsDataAccess

	Private _objLogin As LoginInfo

	Public Sub New()
	End Sub

	Public Sub New(ByVal value As LoginInfo)
		_objLogin = value
	End Sub

	Friend Function OpenRecordset(ByRef sSQL As String, ByRef CursorType As CursorTypeEnum, ByRef LockType As LockTypeEnum _
		, Optional ByVal iCursorLocation As CursorLocationEnum = CursorLocationEnum.adUseServer) As Recordset

		' Open a recordset from the given SQL query, with the given recordset properties.
		Dim rsTemp As Recordset
		Dim iOldCursorLocation As CursorLocationEnum

		Try

			iOldCursorLocation = gADOCon.CursorLocation

			rsTemp = New Recordset
			gADOCon.CursorLocation = iCursorLocation

			rsTemp.Open(sSQL, gADOCon, CursorType, LockType, CommandTypeEnum.adCmdText)

		Catch ex As Exception
			Return Nothing

		Finally
			If (iOldCursorLocation = CursorLocationEnum.adUseClient) Or (iOldCursorLocation = CursorLocationEnum.adUseServer) Then
				gADOCon.CursorLocation = iOldCursorLocation
			Else
				gADOCon.CursorLocation = CursorLocationEnum.adUseServer
			End If

		End Try

		Return rsTemp

	End Function

	Friend Sub ExecuteSql(ByRef sSQL As String)
		' Execute the given SQL statement.
		Dim strConn As String = GetConnectionString(_objLogin)

		Try

			Using sqlConnection As New SqlConnection(strConn)
				Using objCommand = New SqlCommand(sSQL, sqlConnection)

					objCommand.CommandType = CommandType.Text

					objCommand.Parameters.Clear()
					sqlConnection.Open()
					objCommand.ExecuteNonQuery()
				End Using

			End Using

		Catch
			Throw

		End Try

	End Sub

	Friend Function ExecuteSqlReturnAffected(ByRef sSQL As String) As Object
		' Execute the given SQL statement, and return the number of rows affected.
		Dim lngAffected As Object

		gADOCon.Execute(sSQL, lngAffected, CommandTypeEnum.adCmdText)
		ExecuteSqlReturnAffected = lngAffected

	End Function


	Private Function GetConnectionString(ByVal LoginDetail As LoginInfo) As String

		Const _AppName As String = "OpenHR"

		If LoginDetail.TrustedConnection Then
			Return String.Format("Data Source={0};Initial Catalog={1};Trusted_Connection=yes;Application Name={2}" _
													 , LoginDetail.Server, LoginDetail.Database, _AppName)

		Else
			Return String.Format("Data Source={0};Initial Catalog={1};User Id={2};Password={3};Application Name={4}" _
													 , LoginDetail.Server, LoginDetail.Database, LoginDetail.Username, LoginDetail.Password, _AppName)

		End If

	End Function

	Public Function GetFromSP(ByVal ProcedureName As String, ParamArray args() As SqlParameter) As DataTable

		Try
			Return GetDataSet(ProcedureName, CommandType.StoredProcedure, args).Tables(0)

		Catch ex As Exception
			Throw

		End Try

	End Function


	Public Sub ExecuteSP(ByVal ProcedureName As String, ParamArray args() As SqlParameter)

		Dim strConn As String = GetConnectionString(_objLogin)

		Try

			Using sqlConnection As New SqlConnection(strConn)
				Using objCommand = New SqlCommand(ProcedureName, sqlConnection)

					objCommand.CommandType = CommandType.StoredProcedure

					objCommand.Parameters.Clear()
					For Each sqlParm As SqlParameter In args
						objCommand.Parameters.Add(sqlParm)
					Next

					sqlConnection.Open()
					objCommand.ExecuteNonQuery()
				End Using

			End Using

		Catch
			Throw

		End Try

	End Sub

	Public Function GetDataTable(ByVal sProcedureName As String) As DataTable

		Try
			Return GetDataTable(sProcedureName, CommandType.Text)

		Catch ex As Exception
			Throw

		End Try

		Return Nothing


	End Function


	Public Function GetDataTable(ByVal sProcedureName As String, ByVal CommandType As CommandType, ParamArray args() As SqlParameter) As DataTable

		Try
			Return GetDataSet(sProcedureName, CommandType, args).Tables(0)

		Catch ex As Exception
			Throw

		End Try

		Return Nothing

	End Function

	Public Function GetDataTable(ByVal procedureName As String, ByVal parameterName As String, dataList As DataTable) As DataTable

		Dim strConn As String = GetConnectionString(_objLogin)
		Dim objDataSet As New DataSet
		Dim objAdaptor As New SqlDataAdapter


		Try

			Using sqlConnection As New SqlConnection(strConn)
				objAdaptor.SelectCommand = New SqlCommand(procedureName, sqlConnection)
				objAdaptor.SelectCommand.CommandType = CommandType.StoredProcedure
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

	Public Function CallToStoredProcedure(sProcedureName As String, ByVal ParamArray args() As SqlParameter) As String
		Dim strConn As String = GetConnectionString(_objLogin)
		Using sqlConnection As New SqlConnection(strConn)
			Dim sqlCommand As New SqlCommand(sProcedureName, sqlConnection)
			For Each sqlParm In args
				sqlCommand.Parameters.Add(sqlParm)
			Next

			Dim DECLAREs As String = ""
			Dim EXEC As String = ""
			Dim SELECTs As String = ""

			'Declare output variables
			For Each p As System.Data.SqlClient.SqlParameter In sqlCommand.Parameters
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
			For Each p As System.Data.SqlClient.SqlParameter In sqlCommand.Parameters
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
			For Each p As System.Data.SqlClient.SqlParameter In sqlCommand.Parameters
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

	Public Function GetDataSet(ByVal sProcedureName As String, ParamArray args() As SqlParameter) As DataSet
		Return GetDataSet(sProcedureName, CommandType.StoredProcedure, args)
	End Function

	Private Function GetDataSet(ByVal sProcedureName As String, ByVal CommandType As CommandType, ParamArray args() As SqlParameter) As DataSet

		Dim strConn As String = GetConnectionString(_objLogin)
		Dim objDataSet As New DataSet
		Dim objAdaptor As New SqlDataAdapter

		Const RetryThreshold = 5
		Dim iRetryCount As Integer = 0
		Dim bRetry As Boolean = True

		Do While bRetry

			Try

				Using sqlConnection As New SqlConnection(strConn)

					objAdaptor.SelectCommand = New SqlCommand(sProcedureName, sqlConnection)
					objAdaptor.SelectCommand.CommandType = CommandType

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

	Friend Shared Function CreateTable(Of T)() As DataTable
		Dim entityType As Type = GetType(T)
		Dim table As New DataTable(entityType.Name)
		Dim properties As PropertyDescriptorCollection = TypeDescriptor.GetProperties(entityType)

		For Each prop As PropertyDescriptor In properties
			' HERE IS WHERE THE ERROR IS THROWN FOR NULLABLE TYPES
			table.Columns.Add(prop.Name, prop.PropertyType)
		Next

		Return table
	End Function


End Class