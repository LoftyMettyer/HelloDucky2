Option Strict On
Option Explicit On

Imports DMI.NET.Classes
Imports HR.Intranet.Server
Imports System.Data.SqlClient

' Register the module parameters
Public Class SettingsConfig

	Public Shared Personnel_EmpTableID As Integer
	
	Public Shared Sub Register()

		If Licence.IsModuleLicenced(SoftwareModule.Personnel) Then
			PopulatePersonnelSessionVariables()
		End If

	End Sub

	Private Shared Sub PopulatePersonnelSessionVariables()

		Dim objDataAccess As New clsDataAccess(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)
		Dim prmEmpTableID = New SqlParameter("piEmployeeTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

		Try
			objDataAccess.ExecuteSP("sp_ASRIntGetPersonnelParameters", prmEmpTableID)
			Personnel_EmpTableID = CInt(prmEmpTableID.Value)

		Catch ex As Exception
			Throw

		End Try

	End Sub


End Class
