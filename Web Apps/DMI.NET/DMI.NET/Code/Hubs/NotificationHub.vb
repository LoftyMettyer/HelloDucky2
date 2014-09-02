Imports Microsoft.AspNet.SignalR.Hubs
Imports Microsoft.AspNet.SignalR
Imports System.Data.SqlClient

Namespace Code.Hubs

	<HubName("NotificationHub")>
	Public Class NotificationHub
		Inherits Hub

		Public Sub Send(messageFrom As String, message As String)
			Dim allContext = GlobalHost.ConnectionManager.GetHubContext(Of NotificationHub)()
			allContext.Clients.All.SystemAdminMessage(messageFrom, message)
		End Sub

		' Handler method
		Private Sub OnDependencyChange(sender As Object, e As SqlNotificationEventArgs)

			Dim Dependency As SqlDependency = CType(sender, SqlDependency)
			RemoveHandler Dependency.OnChange, AddressOf OnDependencyChange

			GetMessages()

		End Sub

		Public Function GetMessages() As DataTable
			Dim dt As New DataTable()

			Try

				Dim connection = New SqlConnection(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

				Dim cmd As New SqlCommand("SELECT loginname, message, messageSource FROM dbo.ASRSysMessages", connection)
				cmd.CommandType = CommandType.Text

				' Clear any existing notifications
				cmd.Notification = Nothing

				' Create the dependency for this command
				Dim dependency As New SqlDependency(cmd)

				' Add the event handler
				AddHandler dependency.OnChange, AddressOf OnDependencyChange

				' Open the connection if necessary
				If connection.State = ConnectionState.Closed Then
					connection.Open()
				End If

				' Get the messages
				dt.Load(cmd.ExecuteReader(CommandBehavior.CloseConnection))

				Dim sMessage As String = ""
				Dim sMessageFrom As String = ""

				'	Dim context = GlobalHost.ConnectionManager.GetHubContext(Of LogoutHub)()

				For Each objRow In dt.Rows
					sMessageFrom = "System Administrator"
					sMessage += objRow("message")
					'''''		'  objChatHub.Send("server", sMessage)
					''''				'        context.Clients.All.Send("Admin", sMessage)

					'context.Clients.All.broadcastMessage("Admin", sMessage)


				Next

				Send(sMessageFrom, sMessage)


			Catch ex As Exception
				Throw
			End Try

			Return dt
		End Function


	End Class

End Namespace