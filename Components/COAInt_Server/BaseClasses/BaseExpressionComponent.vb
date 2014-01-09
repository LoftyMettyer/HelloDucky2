Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Structures
Imports System.Data.SqlClient

Namespace BaseClasses
	Public Class BaseExpressionComponent

		Protected ReadOnly Login As LoginInfo
		Protected General As New clsGeneral
		Protected DB As New clsDataAccess

		Public Sub New(ByVal Value As LoginInfo)
			Login = Value
			DB = New clsDataAccess(Login)
			General = New clsGeneral(Login)
		End Sub

		' keep a manual record of allocated IDs in case users in SYS MGR have created expressions but not yet saved changes
		Protected Function GetUniqueID(ByRef strSetting As String, ByRef strTable As String, ByRef strColumn As String) As Integer

			Dim prmSettingKey = New SqlParameter("settingkey", SqlDbType.VarChar, 50)
			prmSettingKey.Value = strSetting

			Dim prmSettingValue = New SqlParameter("settingvalue", SqlDbType.Int)
			prmSettingValue.Direction = ParameterDirection.Output

			DB.ExecuteSP("spASRIntGetUniqueExpressionID", prmSettingKey, prmSettingValue)

			Return CInt(prmSettingValue.Value)

		End Function

	End Class
End Namespace