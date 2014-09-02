Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports DMI.NET.Code.Hubs
Imports HR.Intranet.Server.Metadata

Public Class DatabaseConfig

	Public Shared Sub Connect()

		SqlDependency.Start(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)
		Dim NotificationHub As New NotificationHub
		NotificationHub.GetMessages()

	End Sub

	Public Shared Sub Disconnect()

		SqlDependency.Stop(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

	End Sub


End Class
