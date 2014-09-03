Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports DMI.NET.Code
Imports DMI.NET.Classes
Imports DMI.NET.Code.Hubs

Public Class DatabaseConfig

	'Public Shared Server As New DatabaseServer

	Public Shared Sub Connect()

		Dim sConnection = ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString

		Dim connection = New SqlConnection(sConnection)
		ApplicationSettings.LoginPage_Database = connection.Database
		ApplicationSettings.LoginPage_Server = connection.DataSource

		SqlDependency.Start(sConnection)
		Dim NotificationHub As New NotificationHub
		NotificationHub.GetMessages()

	End Sub

	Public Shared Sub Disconnect()

		SqlDependency.Stop(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

	End Sub

End Class
