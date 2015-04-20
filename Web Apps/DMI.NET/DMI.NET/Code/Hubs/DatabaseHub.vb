﻿Option Explicit On
Option Strict On

Imports Microsoft.AspNet.SignalR.Hubs
Imports Microsoft.AspNet.SignalR
Imports System.Data.SqlClient
Imports System.Threading.Tasks
Imports HR.Intranet.Server
Imports DMI.NET.Models

Namespace Code.Hubs

	<HubName("NotificationHub")>
	Public Class DatabaseHub
		Inherits Hub

		Private Shared Connection As SqlConnection

		Public Shared ServiceBrokerOK As Boolean
		Public Shared HeartbeatOK As Boolean
		Public Shared SystemLockStatus As LockPriority = LockPriority.None
		Public Shared LockMessage As String

		Public Overrides Function OnConnected() As Task
			ToggleLoginButton(Not SystemLockStatus = LockPriority.None, LockMessage)
			Return MyBase.OnConnected()
		End Function

		Public Shared ReadOnly Property IISServerName As String
			Get
				Return Environment.MachineName
			End Get
		End Property

		Public Shared ReadOnly Property DatabaseOK As Boolean
			Get
				Return ServiceBrokerOK And HeartbeatOK
			End Get
		End Property

		Public Shared Sub RegisterDatabase()

			' Initialise the heartbeat
			Try

				Dim sConnection = ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString

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

				Dim sConnection = ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString

				SqlDependency.Start(sConnection)
				ClearTrackingTable()
				GetMessages()
				GetLockStatus()
				ServiceBrokerOK = True

			Catch ex As Exception
				ServiceBrokerOK = False

			End Try

		End Sub

		Public Shared Sub ClearTrackingTable()

			Try

				Dim objDataAccess As New clsDataAccess(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)
				objDataAccess.ExecuteSql(String.Format("DELETE FROM ASRSysCurrentSessions WHERE IISServer = '{0}'", IISServerName))

			Catch ex As Exception
				Throw

			End Try

		End Sub

		Public Shared Sub UnRegister()
			SqlDependency.Stop(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)
		End Sub

		Public Shared Sub SendMessage(messageFrom As String, message As String, forceLogout As Boolean, loggedInUsersOnly As Boolean)
			Dim allContext = GlobalHost.ConnectionManager.GetHubContext(Of DatabaseHub, IDatabaseHub)()
			allContext.Clients.All.SystemAdminMessage(messageFrom, message, forceLogout, loggedInUsersOnly)
		End Sub

		Public Shared Sub ToggleLoginButton(disable As Boolean, message As String)

			Dim allContext = GlobalHost.ConnectionManager.GetHubContext(Of DatabaseHub, IDatabaseHub)()
			allContext.Clients.All.ToggleLoginButton(disable, message)
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

				Dim objDataAccess As New clsDataAccess(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)
				Dim prmResult = New SqlParameter("psResult", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

				Dim dt As New DataTable()

				Dim cmd As New SqlCommand(String.Format("SELECT Priority, Username, Module, NotifyGroups FROM dbo.ASRSysLock WHERE Priority <> {0} " _
																								, CInt(LockPriority.ReadWrite)), Connection)
				cmd.CommandType = CommandType.Text
				cmd.Notification = Nothing

				Dim dependency As New SqlDependency(cmd)
				AddHandler dependency.OnChange, AddressOf OnLockStatusChange

				Dim iPriority As LockPriority
				Dim iModule As ApplicationModule
				Dim bForceLogout As Boolean = False

				If Connection.State = ConnectionState.Closed Then
					Connection.Open()
				End If

				LockMessage = ""
				SystemLockStatus = LockPriority.None

				dt.Load(cmd.ExecuteReader())
				For Each objRow As DataRow In dt.Rows
					iPriority = CType(objRow("Priority"), LockPriority)
					iModule = CType(objRow("Module"), ApplicationModule)

					If iModule = ApplicationModule.SecurityManager Then

						If iPriority = LockPriority.Saving Then

							' Notify the affected groups
							Dim notifyGroups As String() = Split(objRow("NotifyGroups").ToString(), ",")
							For Each group In notifyGroups
								SendMessageToGroup(group.Trim(), objRow("username").ToString(), "A security update is in progress.", True)
							Next

						End If

					ElseIf iModule = ApplicationModule.SystemManager Then

						SystemLockStatus = iPriority

						Select Case SystemLockStatus

							Case LockPriority.Manual
								objDataAccess.ExecuteSP("spASRIntGetSetting" _
										, New SqlParameter("psSection", SqlDbType.VarChar, -1) With {.Value = "messaging"} _
										, New SqlParameter("psKey", SqlDbType.VarChar, -1) With {.Value = "lockmessage"} _
										, New SqlParameter("psDefault", SqlDbType.VarChar, -1) With {.Value = "A system administrator has locked the database."} _
										, New SqlParameter("pfUserSetting", SqlDbType.Bit) With {.Value = False} _
										, prmResult)
								LockMessage = prmResult.Value.ToString

							Case LockPriority.SaveRequest
								LockMessage = "A database system save is in progress."
								bForceLogout = True

							Case LockPriority.Saving
								LockMessage = "A database system save is in progress."
								SendMessage("System Administrator", LockMessage, True, True)
								bForceLogout = True

							Case LockPriority.Manual
								LockMessage = "A system administrator has locked the database."
								bForceLogout = True

							Case Else
								bForceLogout = False

						End Select
					End If

				Next

				ToggleLoginButton(bForceLogout, LockMessage)


			Catch ex As Exception
				Throw

			End Try

		End Sub

		Public Shared Sub GetMessages()
			Dim dt As New DataTable()

			Try

				Dim sMessage As String = ""
				Dim sMessageFrom As String = ""

				Dim cmd As New SqlCommand("SELECT loginname, message, messageSource FROM dbo.ASRSysMessages WHERE LoginName = 'OpenHR Web Server' ORDER BY messageTime DESC", Connection)
				cmd.CommandType = CommandType.Text
				cmd.Notification = Nothing

				Dim dependency As New SqlDependency(cmd)
				AddHandler dependency.OnChange, AddressOf OnMessageChange

				If Connection.State = ConnectionState.Closed Then
					Connection.Open()
				End If

				dt.Load(cmd.ExecuteReader)
				For Each objRow As DataRow In dt.Rows
					sMessageFrom = "System Administrator"
					sMessage += objRow("message").ToString
				Next

				SendMessage(sMessageFrom, sMessage, False, True)

			Catch ex As Exception
				Throw
			End Try

		End Sub

		Friend Shared Sub TrackSession(objLogin As LoginViewModel, trackType As TrackType)

			Try

				Dim objDataAccess As New clsDataAccess(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

				objDataAccess.ExecuteSP("spASRIntTrackSession" _
						, New SqlParameter("@IISServer", SqlDbType.NVarChar, 255) With {.Value = IISServerName} _
						, New SqlParameter("@SessionID", SqlDbType.NVarChar, 255) With {.Value = objLogin.SessionId} _
						, New SqlParameter("@Username", SqlDbType.NVarChar, 255) With {.Value = objLogin.UserName} _
						, New SqlParameter("@SecurityGroup", SqlDbType.VarChar, 255) With {.Value = objLogin.SecurityGroup} _
						, New SqlParameter("@HostName", SqlDbType.VarChar, 255) With {.Value = objLogin.Device} _
						, New SqlParameter("@WebArea", SqlDbType.VarChar, 20) With {.Value = objLogin.WebAreaName} _
						, New SqlParameter("@TrackType", SqlDbType.TinyInt) With {.Value = CInt(trackType)})

			Catch ex As Exception
				Throw

			End Try

		End Sub

		<HubMethodName("joinGroup")>
		Public Sub JoinGroup(group As String)
			Groups.Add(Context.ConnectionId, group)
		End Sub

		Public Shared Sub SendMessageToGroup(group As String, messageFrom As String, message As String, forceLogout As Boolean)

			Dim allContext = GlobalHost.ConnectionManager.GetHubContext(Of DatabaseHub, IDatabaseHub)()
			allContext.Clients.Group(group).NotifyGroup(messageFrom, message, forceLogout)

		End Sub

	End Class

End Namespace