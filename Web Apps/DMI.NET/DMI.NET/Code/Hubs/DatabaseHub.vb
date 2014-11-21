Imports Microsoft.AspNet.SignalR.Hubs
Imports Microsoft.AspNet.SignalR
Imports System.Data.SqlClient

Namespace Code.Hubs

	<HubName("NotificationHub")>
	Public Class DatabaseHub
		Inherits Hub

		Private Shared Connection As SqlConnection

		Public Shared ServiceBrokerOK As Boolean
		Public Shared HeartbeatOK As Boolean

		Public Shared ReadOnly Property DatabaseOK As Boolean
			Get
				Return ServiceBrokerOK And HeartbeatOK
			End Get
		End Property

		Public Shared Sub RegisterDatabase()

			Dim sConnection = ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString

			' Initialise the heartbeat
			Try

				Connection = New SqlConnection(sConnection)
				If Connection.State = ConnectionState.Closed Then
					Connection.Open()
				End If

				ApplicationSettings.LoginPage_Database = Connection.Database
				ApplicationSettings.LoginPage_Server = Connection.DataSource
				HeartbeatOK = True

			Catch ex As Exception
				HeartbeatOK = False

			End Try


			' Initialise the service broker
			Try

				SqlDependency.Start(sConnection)
				GetMessages()
				GetLockStatus()
				ServiceBrokerOK = True

			Catch ex As Exception
				ServiceBrokerOK = False

			End Try

		End Sub

		Public Shared Sub UnRegister()
			SqlDependency.Stop(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)
		End Sub

		Public Shared Sub SendMessage(messageFrom As String, message As String, forceLogout As Boolean)
			Dim allContext = GlobalHost.ConnectionManager.GetHubContext(Of DatabaseHub)()
			allContext.Clients.All.SystemAdminMessage(messageFrom, message, forceLogout)
		End Sub

		Private Shared Sub OnMessageChange(sender As Object, e As SqlNotificationEventArgs)
			Dim Dependency As SqlDependency = CType(sender, SqlDependency)
			RemoveHandler Dependency.OnChange, AddressOf OnMessageChange
			GetMessages()
		End Sub

		Private Shared Sub OnLockStatusChange(sender As Object, e As SqlNotificationEventArgs)
			Dim Dependency As SqlDependency = CType(sender, SqlDependency)
			RemoveHandler Dependency.OnChange, AddressOf OnLockStatusChange
			GetLockStatus()
		End Sub

		Public Shared Sub GetLockStatus()

			Try

				Dim dt As New DataTable()

				Dim cmd As New SqlCommand("SELECT Priority, Username FROM dbo.ASRSysLock WHERE Priority IN (1,2) ", Connection)
				cmd.CommandType = CommandType.Text
				cmd.Notification = Nothing

				Dim dependency As New SqlDependency(cmd)
				AddHandler dependency.OnChange, AddressOf OnLockStatusChange

				Dim iPriority As LockPriority
				Dim sMessage As String = ""
				Dim sMessageFrom As String = ""

				If Connection.State = ConnectionState.Closed Then
					Connection.Open()
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

		Public Shared Sub GetMessages()
			Dim dt As New DataTable()

			Try

				Dim sMessage As String = ""
				Dim sMessageFrom As String = ""

				Dim cmd As New SqlCommand("SELECT loginname, message, messageSource FROM dbo.ASRSysMessages WHERE LoginName = 'OpenHR Web Server'", Connection)
				cmd.CommandType = CommandType.Text
				cmd.Notification = Nothing
				Dim dependency As New SqlDependency(cmd)
				AddHandler dependency.OnChange, AddressOf OnMessageChange

				If Connection.State = ConnectionState.Closed Then
					Connection.Open()
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