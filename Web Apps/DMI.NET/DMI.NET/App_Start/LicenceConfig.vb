Option Strict On
Option Explicit On

Imports DMI.NET.Classes
Imports DMI.NET.Code.Hubs
Imports HR.Intranet.Server

Public Class LicenceConfig

	Public Shared Licence As New Licence

	Public Shared Sub RegisterLicence()

		Try

			Dim objDataAccess As New clsDataAccess(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)
			Dim sLicence As String = ""

			Dim dtSetting = objDataAccess.GetDataTable("SELECT SettingValue FROM ASRSysSystemSettings WHERE Section = 'licence' AND SettingKey = 'key'")
			If dtSetting.Rows.Count > 0 Then
				sLicence = dtSetting.Rows(0)("SettingValue").ToString()
			End If

			Licence.Populate(sLicence)

			If Not Licence.Type = LicenceType.Concurrency Then
				LicenceHub.ValidateHeadCount()
			End If

		Catch ex As Exception
			Throw

		End Try

	End Sub

End Class