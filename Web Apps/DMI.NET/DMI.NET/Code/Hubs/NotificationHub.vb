Imports Microsoft.AspNet.SignalR.Hubs
Imports Microsoft.AspNet.SignalR
Imports System.Data.SqlClient

Namespace Code.Hubs

	<HubName("NotificationHub")>
	Public Class NotificationHub
		Inherits Hub

		Private Shared connection As SqlConnection

		Public Shared Sub Connect()

			connection = New SqlConnection(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)
			If connection.State = ConnectionState.Closed Then
				connection.Open()
			End If

		End Sub

		Public Shared Sub SendMessage(messageFrom As String, message As String, forceLogout As Boolean)
			Dim allContext = GlobalHost.ConnectionManager.GetHubContext(Of NotificationHub)()
			allContext.Clients.All.SystemAdminMessage(messageFrom, message, forceLogout)
		End Sub

		Private Sub OnMessageChange(sender As Object, e As SqlNotificationEventArgs)
			Dim Dependency As SqlDependency = CType(sender, SqlDependency)
			RemoveHandler Dependency.OnChange, AddressOf OnMessageChange
			GetMessages()
		End Sub

		Private Sub OnLockStatusChange(sender As Object, e As SqlNotificationEventArgs)
			Dim Dependency As SqlDependency = CType(sender, SqlDependency)
			RemoveHandler Dependency.OnChange, AddressOf OnLockStatusChange
			GetLockStatus()
		End Sub

		Public Sub GetLockStatus()

			Try

				Dim dt As New DataTable()

				Dim cmd As New SqlCommand("SELECT Priority, Username FROM dbo.ASRSysLock WHERE Priority IN (1,2) ", connection)
				cmd.CommandType = CommandType.Text
				cmd.Notification = Nothing

				Dim dependency As New SqlDependency(cmd)
				AddHandler dependency.OnChange, AddressOf OnLockStatusChange

				Dim iPriority As LockPriority
				Dim sMessage As String = ""
				Dim sMessageFrom As String = ""

				If connection.State = ConnectionState.Closed Then
					connection.Open()
				End If

				dt.Load(cmd.ExecuteReader())
				For Each objRow In dt.Rows
					iPriority = CType(objRow("Priority"), LockPriority)
					sMessageFrom = objRow("UserName").ToString()
					sMessage += "The system administrator has initiated a system save. You will need to logout."
				Next

				If iPriority = LockPriority.Manual OrElse iPriority = LockPriority.Saving Then
					SendMessage(sMessageFrom, sMessage, True)
				End If

			Catch ex As Exception
				Throw

			End Try

		End Sub

		Public Sub GetMessages()
			Dim dt As New DataTable()

			Try

				Dim sMessage As String = ""
				Dim sMessageFrom As String = ""

				Dim cmd As New SqlCommand("SELECT loginname, message, messageSource FROM dbo.ASRSysMessages WHERE LoginName = 'OpenHR Web Server'", connection)
				cmd.CommandType = CommandType.Text
				cmd.Notification = Nothing
				Dim dependency As New SqlDependency(cmd)
				AddHandler dependency.OnChange, AddressOf OnMessageChange

				If connection.State = ConnectionState.Closed Then
					connection.Open()
				End If

				dt.Load(cmd.ExecuteReader)
				For Each objRow In dt.Rows
					sMessageFrom = "System Administrator"
					sMessage += objRow("message")
				Next

				SendMessage(sMessageFrom, sMessage, False)

			Catch ex As Exception
				Throw
			End Try

		End Sub


	End Class

End Namespace