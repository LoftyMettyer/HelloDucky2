Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports DMI.NET.Classes
Imports Microsoft.AspNet.SignalR
Imports Microsoft.AspNet.SignalR.Hubs
Imports System.Threading.Tasks
Imports DMI.NET.Models
Imports System.Web.Script.Serialization
Imports System.Data.SqlClient

Namespace Code.Hubs

	<HubName("LicenceHub")> _
	Public Class LicenceHub
		Inherits Hub

		Private Shared ReadOnly Sessions As New List(Of String)
		Private Shared ReadOnly Logins As New List(Of LoginViewModel)

		Private Shared current_SSIUsers As Integer = 0
		Private Shared current_DMIUsers As Integer = 0
		Private Shared current_DMISingleUsers As Integer = 0
		Private Shared current_Headcount As Integer = 0

		'<HubMethodName("")
		Private Shared Sub UpdateOnlineCount()

			Dim totalLogins As Integer

			current_SSIUsers = 0
			current_DMIUsers = 0
			current_DMISingleUsers = 0

			For Each sSession In Sessions
				current_DMIUsers += Logins.LongCount(Function(m) m.IsLoggedIn = True AndAlso m.SignalRClientID = sSession AndAlso m.WebArea = WebArea.DMI)
				current_DMISingleUsers += Logins.LongCount(Function(m) m.IsLoggedIn = True AndAlso m.SignalRClientID = sSession AndAlso m.WebArea = WebArea.DMISingle)
				current_SSIUsers += Logins.LongCount(Function(m) m.IsLoggedIn = True AndAlso m.SignalRClientID = sSession AndAlso m.WebArea = WebArea.SSI)
			Next

			totalLogins = current_DMIUsers + current_DMISingleUsers + current_SSIUsers

			Dim myContext = GlobalHost.ConnectionManager.GetHubContext(Of LicenceHub)()
			myContext.Clients.All.updateUsersOnlineCount(totalLogins)

		End Sub

		Private Sub UpdateUserList()

			Dim objLogins As New List(Of LoginViewModel)

			For Each sSession In Sessions
				objLogins.AddRange(Logins.Where(Function(m) m.IsLoggedIn = True AndAlso m.SignalRClientID = sSession))
			Next

			Dim results = New With {.total = 1, .page = 1, .records = objLogins.Count(), .rows = objLogins}

			Dim objSerialize As New JavaScriptSerializer
			Dim result = objSerialize.Serialize(results)

			Dim myContext = GlobalHost.ConnectionManager.GetHubContext(Of LicenceHub)()
			myContext.Clients.All.currentUserList(result)

		End Sub

		Private Shared Sub NotifyDisableLogin()

			Dim bDisabled As Boolean
			Dim sMessage As String = ""
			Dim totalLogins = current_DMIUsers + current_DMISingleUsers + current_SSIUsers

			If totalLogins >= (Licence.DMIUsers + Licence.DMISingleUsers + Licence.SSIUsers) Then
				bDisabled = True
				sMessage = "There are too many users logged in at the moment. Please contact your system administrator"
			End If

			If Now > Licence.ExpiryDate Then
				bDisabled = True
				sMessage = "Your licence has expired. Please contact your system administrator."
			End If

			Dim myContext = GlobalHost.ConnectionManager.GetHubContext(Of LicenceHub)()
			'myContext.Clients.All.disableLogin(bDisabled, sMessage)

		End Sub

		Private Shared Sub NotifyHeadcountExceeded()

			Dim notifyHub = GlobalHost.ConnectionManager.GetHubContext(Of NotificationHub)()
			notifyHub.Clients.All.SystemAdminMessage("OpenHR", String.Format("Your database contains too many records ({0}).</br>You are licenced for only {1} employees. Start firing some people", Licence.Headcount, current_Headcount))

			Dim thisHub = GlobalHost.ConnectionManager.GetHubContext(Of LicenceHub)()
			thisHub.Clients.All.disableLogin(True, "")

		End Sub

		Public Overrides Function OnConnected() As Task
			Dim clientId As String = GetClientId()

			If Sessions.IndexOf(clientId) = -1 Then
				Sessions.Add(clientId)
			End If

			If Not Logins.Exists(Function(m) m.SignalRClientID = clientId) Then
				Logins.Add(New LoginViewModel() With {
				 .SignalRClientID = clientId})
			End If

			' Send the current count of users
			UpdateOnlineCount()
			NotifyDisableLogin()
			UpdateUserList()

			Return MyBase.OnConnected()
		End Function

		Public Overrides Function OnReconnected() As Task

			Dim clientId As String = GetClientId()

			If Sessions.IndexOf(clientId) = -1 Then
				Sessions.Add(clientId)
			End If

			If Not Logins.Exists(Function(m) m.SignalRClientID = clientId) Then
				Logins.Add(New LoginViewModel() With {
				 .SignalRClientID = clientId})
			End If

			' Send the current count of users
			UpdateOnlineCount()
			NotifyDisableLogin()
			UpdateUserList()

			Return MyBase.OnReconnected()
		End Function

		Public Overrides Function OnDisconnected(stopCalled As Boolean) As Task

			Dim clientId As String = GetClientId()

			If Sessions.IndexOf(clientId) > -1 Then
				'For Each objLogin In Logins.Where(Function(m) m.SignalRClientID = clientId)
				'	objLogin.IsLoggedIn = False
				'Next
				Sessions.Remove(clientId)
			End If

			' Send the current count of users
			UpdateOnlineCount()
			NotifyDisableLogin()
			UpdateUserList()

			Return MyBase.OnDisconnected(stopCalled)
		End Function

		Private Function GetClientId() As String
			Dim clientId As String
			'If Context.QueryString("clientId") IsNot Nothing Then
			'	' clientId passed from application 
			'	clientId = Context.QueryString("clientId")
			'End If

			'If String.IsNullOrEmpty(clientId.Trim()) Then
			'	clientId = Context.ConnectionId
			'End If

			clientId = Context.RequestCookies("ASP.NET_SessionId").Value ' HttpContext.Current.Session.SessionID

			Return clientId
		End Function

		Private Shared Function AllowAccess(webArea As WebArea) As LicenceValidation

			Dim allow As LicenceValidation = LicenceValidation.Ok

			If (Now > Licence.ExpiryDate) Then
				allow = LicenceValidation.Expired
			End If

			If Licence.Type = LicenceType.Headcount Then
				If current_Headcount > Licence.Headcount Then
					allow = LicenceValidation.Insufficient
				End If

			ElseIf Licence.Type = LicenceType.P14Headcount Then
				If current_Headcount > Licence.P14Headcount Then
					allow = LicenceValidation.Insufficient
				End If

			Else
				Select Case webArea
					Case webArea.DMI
						If current_DMIUsers >= Licence.DMIUsers Then
							allow = LicenceValidation.Insufficient
						End If

					Case webArea.DMISingle
						If current_DMISingleUsers >= Licence.DMISingleUsers Then
							allow = LicenceValidation.Insufficient
						End If

					Case Else

						If current_SSIUsers >= Licence.SSIUsers Then
							allow = LicenceValidation.Insufficient
						End If

				End Select

			End If

			Return allow

		End Function

		Public Shared Function LogIn(SessionID As String, objActualLogin As LoginViewModel, webArea As WebArea) As LicenceValidation
			Dim objLogin = Logins.First(Function(m) m.SignalRClientID = SessionID)

			Dim allow = AllowAccess(webArea)

			If allow = LicenceValidation.Ok Then
				objLogin.IsLoggedIn = True
				objLogin.UserName = objActualLogin.UserName
				objLogin.WebArea = webArea
				UpdateOnlineCount()
			End If

			Return allow

		End Function

		Public Shared Function NavigateWebArea(SessionID As String, webArea As WebArea) As LicenceValidation

			Dim objLogin = Logins.First(Function(m) m.SignalRClientID = SessionID)
			Dim allow As LicenceValidation = LicenceValidation.Ok

			If Not objLogin.WebArea = webArea Then
				allow = AllowAccess(webArea)
				If allow = LicenceValidation.Ok Then
					objLogin.IsLoggedIn = True
					objLogin.WebArea = webArea
					UpdateOnlineCount()
				End If
			End If

			Return allow

		End Function

		Public Shared Sub LogOff(SessionID As String)
			Logins.RemoveAll(Function(m) m.SignalRClientID = SessionID)
		End Sub

		Public Shared Sub ValidateHeadCount()

			'	Dim objDataAccess As New clsDataAccess(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)
			'	Dim dtHeadcount = objDataAccess.GetDataTable("SELECT Headcount FROM ASRSysEmployees", CommandType.Text)
			'	current_Headcount = CLng(dtHeadcount.Rows(0)("Headcount"))
			Try

				Dim dt As New DataTable()

				Dim connection = New SqlConnection(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

				'Dim cmd As New SqlCommand("SELECT surname FROM dbo.ASRSysEmployees", connection)
				'Dim cmd As New SqlCommand("SELECT surname FROM dbo.tbuser_Personnel_Records", connection)
				Dim cmd As New SqlCommand("dbo.spASRGetHeadcount", connection)

				'Dim cmd As New SqlCommand("SELECT loginname, message, messageSource FROM dbo.ASRSysMessages", connection)
				cmd.CommandType = CommandType.StoredProcedure

				cmd.Parameters.Add(New SqlParameter("@type", CInt(Licence.Type)))
				cmd.Parameters.Add(New SqlParameter("@today", DateTime.Now))

				cmd.Notification = Nothing
				Dim dependency As New SqlDependency(cmd)
				AddHandler dependency.OnChange, AddressOf HeadcountChange

				' Open the connection if necessary
				If connection.State = ConnectionState.Closed Then
					connection.Open()
				End If

				' Get the messages
				dt.Load(cmd.ExecuteReader(CommandBehavior.CloseConnection))
				current_Headcount = dt.Rows.Count

				' We've exceeded our licence
				If (Licence.Type = LicenceType.P14Headcount AndAlso current_Headcount > Licence.P14Headcount) _
					OrElse (Licence.Type = LicenceType.Headcount AndAlso current_Headcount > Licence.Headcount) Then
					NotifyHeadcountExceeded()
				End If

			Catch ex As Exception
				current_Headcount = -1

			End Try

		End Sub

		Private Shared Sub HeadcountChange(sender As Object, e As SqlNotificationEventArgs)

			Dim Dependency As SqlDependency = CType(sender, SqlDependency)
			RemoveHandler Dependency.OnChange, AddressOf HeadcountChange

			ValidateHeadCount()
		End Sub


	End Class
End Namespace
