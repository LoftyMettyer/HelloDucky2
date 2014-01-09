Imports HR.Intranet.Server.BaseClasses
Imports System.Data.SqlClient

Public Class Database
	Inherits BaseForDMI

	Public Function GetEmailAddress(ByRef lngRecordID As Integer, lngEmailAddrCalc As Integer) As String

		' Check if the user can create New instances of the given category.

		Try

			Dim prmResult = New SqlParameter("Result", SqlDbType.VarChar)
			prmResult.Direction = ParameterDirection.Output

			Dim prmEmail = New SqlParameter("EmailID", SqlDbType.Int)
			prmEmail.Value = lngEmailAddrCalc

			Dim prmRecordID = New SqlParameter("RecordID", SqlDbType.Int)
			prmRecordID.Value = lngRecordID

			DB.ExecuteSP("spASRSysEmailAddr", prmResult, prmEmail, prmRecordID)

			Return prmResult.Value

		Catch ex As Exception
			Return ""

		End Try

	End Function


End Class
