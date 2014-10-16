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
Imports HR.Intranet.Server

Namespace Code.Hubs

	<HubName("LicenceHub")> _
	Public Class LicenceHub
		Inherits Hub

		Private Shared Property Connection As SqlConnection

		Private Shared ReadOnly Licence As New Licence

		Private Shared ReadOnly Sessions As New List(Of String)
		Private Shared ReadOnly Logins As New List(Of LoginViewModel)

		Private Shared current_SSIUsers As Integer = 0
		Private Shared current_DMIUsers As Integer = 0
		Private Shared current_Headcount As Long = 0

		Private Const HeadcountWarningThreshold = 0.95

		Shared Sub New()

			Connection = New SqlConnection(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

		End Sub

		Friend Shared Function ErrorMessage(failureCode As LicenceValidation) As String

			Dim message As String = ""

			Select Case failureCode
				Case LicenceValidation.Expired
					message = String.Format("Your licence to use this product has expired.<br/><br/>" & _
																	"Please contact your Account Manager as soon as possible.")

				Case LicenceValidation.ExpiryWarning
					Dim sLocaleFormat = HttpContext.Current.Session("LocaleDateFormat").ToString()
					message = String.Format("Your licence to use this product will expire on {0}.<br/><br/>" & _
																	"Please contact your Account Manager as soon as possible.", _
																 Licence.ExpiryDate.ToString(sLocaleFormat))

				Case LicenceValidation.HeadcountExceeded
					message = String.Format("You have reached or exceeded the headcount limit<br/>set within the terms of your licence agreement.<br/><br/>" & _
																	"You are no longer able to add new employee records,<br/>but you may access the system for other purposes.<br/><br/>" & _
																	"Please contact your Account Manager as soon as possible<br/>to increase the licence headcount number.")

				Case LicenceValidation.HeadcountWarning
					message = String.Format("You are currently within 95% ({0} of {1} employees) of reaching the<br/>headcount limit set within the terms of your licence agreement.<br/><br/>" & _
																	"Once this limit is reached, you will no longer be able to add<br/>new employee records to the system.<br/><br/>" & _
																	"If you wish to increase the headcount number, please<br/>contact your Account Manager as soon as possible.", _
																	current_Headcount, Licence.Headcount)

				Case LicenceValidation.Insufficient
					message = "The maximum number of licenced users are currently<br/>logged into OpenHR - Please try again later.<br/><br/>" & _
																	"If you wish to increase the number of licenced users, please<br/>contact your Account Manager as soon as possible."

				Case LicenceValidation.Failure
					message = "An error has occured connecting to the database<br/>Please contact your system administrator.<br/><br/>"

			End Select

			Return message

		End Function

		Private Shared Sub UpdateOnlineCount()

			Dim totalLogins As Integer

			current_SSIUsers = 0
			current_DMIUsers = 0

			For Each sSession In Sessions
				current_DMIUsers += Logins.LongCount(Function(m) m.IsLoggedIn = True AndAlso m.SignalRClientID = sSession AndAlso m.WebArea = WebArea.DMI)
				current_SSIUsers += Logins.LongCount(Function(m) m.IsLoggedIn = True AndAlso m.SignalRClientID = sSession AndAlso m.WebArea = WebArea.SSI)
			Next

			totalLogins = current_DMIUsers + current_SSIUsers

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

		Private Sub ActivateLogin(thisConnection As String)

			Dim myContext = GlobalHost.ConnectionManager.GetHubContext(Of LicenceHub)()
			'			myContext.Clients.Client(thisConnection).activateLogin()
			myContext.Clients.All.activateLogin()

		End Sub


		Friend Shared Function DisplayWarningToUser(userName As String, warningType As WarningType, warningRefreshRate As Integer) As Boolean

			Try

				Dim objDataAccess As New clsDataAccess(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

				Dim prmWarnUser As New SqlParameter("@WarnUser", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				objDataAccess.ExecuteSP("spASRUpdateWarningLog" _
						, New SqlParameter("@Username", SqlDbType.VarChar, 255) With {.Value = userName} _
						, New SqlParameter("@WarningType", SqlDbType.Int) With {.Value = CInt(warningType)} _
						, New SqlParameter("@WarningRefreshRate", SqlDbType.Int) With {.Value = warningRefreshRate} _
						, prmWarnUser)

				Return CBool(prmWarnUser.Value)

			Catch ex As Exception
				Throw

			End Try

		End Function

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
			ActivateLogin(clientId)
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
			ActivateLogin(clientId)
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

		Private Shared Function AllowAccess(targetWebArea As WebArea) As LicenceValidation

			If (Now.Date > Licence.ExpiryDate) Then
				Return LicenceValidation.Expired
			End If

			If (Now.Date > Licence.ExpiryDate.AddDays(-7)) Then
				Return LicenceValidation.ExpiryWarning
			End If

			If Licence.Type = LicenceType.Concurrency Then
				If (targetWebArea = WebArea.DMI AndAlso current_DMIUsers >= Licence.DMIUsers) OrElse _
						(targetWebArea = WebArea.DMI AndAlso current_SSIUsers >= Licence.SSIUsers) Then
					Return LicenceValidation.Insufficient
				Else
					Return LicenceValidation.Ok
				End If
			End If

			If Licence.Type = LicenceType.DMIConcurrencyAndHeadcount OrElse Licence.Type = LicenceType.DMIConcurrencyAndHeadcount Then
				If (targetWebArea = WebArea.DMI AndAlso current_DMIUsers >= Licence.DMIUsers) Then
					Return LicenceValidation.Insufficient
				End If
			End If

			If targetWebArea = WebArea.DMI Then
				If current_Headcount > Licence.Headcount Then
					Return LicenceValidation.HeadcountExceeded
				ElseIf current_Headcount >= Licence.Headcount * HeadcountWarningThreshold Then
					Return LicenceValidation.HeadcountWarning
				End If
			End If

			Return LicenceValidation.Ok

		End Function

		Public Shared Function NavigateWebArea(SessionID As String, loginName As String, webArea As WebArea) As LicenceValidation

			Try

				Dim objLogin = Logins.First(Function(m) m.SignalRClientID = SessionID)
				Dim allow As LicenceValidation = LicenceValidation.Ok

				If Not objLogin.WebArea = webArea Then
					allow = AllowAccess(webArea)
					If allow = LicenceValidation.Ok OrElse allow = LicenceValidation.HeadcountExceeded OrElse allow = LicenceValidation.HeadcountWarning Then
						objLogin.IsLoggedIn = True
						objLogin.UserName = loginName
						objLogin.WebArea = webArea
						UpdateOnlineCount()
					End If
				End If

				Return allow

			Catch ex As Exception
				Return LicenceValidation.Failure

			End Try

		End Function

		Public Shared Sub LogOff(SessionID As String)
			Logins.RemoveAll(Function(m) m.SignalRClientID = SessionID)
		End Sub

		Public Shared Sub ValidateHeadCount()

			Try

				Dim sGetHeadcount As String

				Dim dt As New DataTable()

				Select Case Licence.Type
					Case LicenceType.Headcount, LicenceType.DMIConcurrencyAndHeadcount
						sGetHeadcount = "SELECT SettingValue FROM dbo.ASRSysSystemSettings WHERE Section = 'Headcount' AND SettingKey = 'current'"

					Case Else
						sGetHeadcount = "SELECT SettingValue FROM dbo.ASRSysSystemSettings WHERE Section = 'Headcount' AND SettingKey = 'P14'"

				End Select

				Dim cmd As New SqlCommand(sGetHeadcount, Connection)
				cmd.CommandType = CommandType.Text

				cmd.Notification = Nothing
				Dim dependency As New SqlDependency(cmd)
				AddHandler dependency.OnChange, AddressOf HeadcountChange

				' Open the connection if necessary
				If Connection.State = ConnectionState.Closed Then
					Connection.Open()
				End If

				' Get the messages
				dt.Load(cmd.ExecuteReader(CommandBehavior.CloseConnection))
				current_Headcount = CLng(dt.Rows(0)(0))

			Catch ex As Exception
				current_Headcount = -1

			End Try

		End Sub

		Private Shared Sub HeadcountChange(sender As Object, e As SqlNotificationEventArgs)

			Dim Dependency As SqlDependency = CType(sender, SqlDependency)
			RemoveHandler Dependency.OnChange, AddressOf HeadcountChange

			ValidateHeadCount()
		End Sub

		Public Shared Sub RegisterLicence()

			Try

				Dim dt As New DataTable()
				Dim sLicence As String

				Const sSQL As String = "SELECT SettingValue FROM dbo.ASRSysSystemSettings WHERE Section = 'licence' AND SettingKey = 'key'"

				Dim cmd As New SqlCommand(sSQL, Connection)
				cmd.CommandType = CommandType.Text

				cmd.Notification = Nothing
				Dim dependency As New SqlDependency(cmd)
				AddHandler dependency.OnChange, AddressOf LicenceKeyChange

				' Open the connection if necessary
				If Connection.State = ConnectionState.Closed Then
					Connection.Open()
				End If

				' Get the messages
				dt.Load(cmd.ExecuteReader(CommandBehavior.CloseConnection))
				sLicence = dt.Rows(0)(0).ToString()

				Licence.Populate(sLicence)

				If Not Licence.Type = LicenceType.Concurrency Then
					ValidateHeadCount()
				End If

			Catch ex As Exception
				Throw

			End Try

		End Sub

		Private Shared Sub LicenceKeyChange(sender As Object, e As SqlNotificationEventArgs)

			Dim Dependency As SqlDependency = CType(sender, SqlDependency)
			RemoveHandler Dependency.OnChange, AddressOf LicenceKeyChange

			RegisterLicence()
		End Sub

	End Class
End Namespace
