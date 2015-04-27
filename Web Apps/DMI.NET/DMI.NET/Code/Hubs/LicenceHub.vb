Option Explicit On
Option Strict On

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
		Inherits Hub(Of ILicenceHub)

		Private Shared Property Connection As SqlConnection

		Private Shared ReadOnly Licence As New Licence

		Private Shared ReadOnly SessionIds As New List(Of String)
		Private Shared ReadOnly Logins As New List(Of LoginViewModel)

		Private Shared current_SSIUsers As Long = 0
		Private Shared current_DMIUsers As Long = 0
		Private Shared current_Headcount As Long = 0

		Private Const HeadcountWarningThreshold = 0.95

		Shared Sub New()

			Connection = New SqlConnection(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

		End Sub

		Friend Shared Function ErrorMessage(failureCode As LicenceValidation) As String

			Dim currentServicesPhoneNo As String = HttpContext.Current.Session("SupportTelNo").ToString()
			Dim sLocaleFormat = HttpContext.Current.Session("LocaleDateFormat").ToString()
			Dim message As String = ""


			If currentServicesPhoneNo = "" Then
				currentServicesPhoneNo = "08451 609 999"
			End If


			Select Case failureCode
				Case LicenceValidation.Expired
					message = String.Format("Your licence to use this product has expired.<br/><br/>" & _
																	"Please contact OpenHR Customer Services on {0} as soon as possible.", _
																	currentServicesPhoneNo)

				Case LicenceValidation.ExpiryWarning
					message = String.Format("Your licence to use this product will expire on {0}.<br/><br/>" & _
																	"Please contact OpenHR Customer Services on {1}</br>as soon as possible.", _
																 Licence.ExpiryDate.ToString(sLocaleFormat), currentServicesPhoneNo)

				Case LicenceValidation.HeadcountExceeded
					message = String.Format("You have reached or exceeded the headcount limit set within</br>the terms of your licence agreement.<br/><br/>" & _
																	"You are no longer able to add new employee records, but you</br>may access the system for other purposes.<br/><br/>" & _
																	"Please contact OpenHR Customer Services on {0}<br/>as soon as possible to increase the licence headcount number.", _
																	currentServicesPhoneNo)

				Case LicenceValidation.HeadcountWarning
					message = String.Format("You are currently within 95% ({0} of {1} employees) of reaching the<br/>headcount limit set within the terms of your licence agreement.<br/><br/>" & _
																	"Once this limit is reached, you will no longer be able to add<br/>new employee records to the system.<br/><br/>" & _
																	"If you wish to increase the headcount number, please contact</br>OpenHR Customer Services on {2} as soon as possible.", _
																	current_Headcount, Licence.Headcount, currentServicesPhoneNo)

				Case LicenceValidation.HeadcountAndExpiryWarning
					message = String.Format("Your licence to use this product will expire on {2}.<br/><br/>You are also within 95% ({0} of {1} employees) of reaching the<br/>" & _
																	"headcount limit set within the terms of your licence agreement.<br/><br/>" & _
																	"Once this limit is reached, you will no longer be able to add<br/>" & _
																	"new employee records to the system.<br/><br/>" & _
																	"Please contact OpenHR Customer Services on {3}</br>as soon as possible.", _
																	current_Headcount, Licence.Headcount, Licence.ExpiryDate.ToString(sLocaleFormat), currentServicesPhoneNo)

				Case LicenceValidation.HeadcountExceededAndExpiryWarning
					message = String.Format("Your licence to use this product will expire on {2}.<br/><br/>You have also reached or exceeded the headcount limit set</br> within the terms of your licence agreement.<br/><br/>" & _
																	"You are no longer able to add new employee records, but you</br>may access the system for other purposes.<br/><br/>" & _
																	"Please contact OpenHR Customer Services on {3}</br>as soon as possible.", _
																	current_Headcount, Licence.Headcount, Licence.ExpiryDate.ToString(sLocaleFormat), currentServicesPhoneNo)

				Case LicenceValidation.Insufficient
					message = String.Format("The maximum number of licenced users are currently logged into OpenHR</br></br>Please try again later.<br/><br/>" & _
																	"If you wish to increase the number of licenced users, please contact</br>OpenHR Customer Services on {0} as soon as possible." _
																	, currentServicesPhoneNo)

				Case LicenceValidation.Failure
					message = "An error has occured connecting to the database<br/>Please contact your system administrator.<br/><br/>"

			End Select

			Return message

		End Function

		Private Shared Sub UpdateOnlineCount()

			current_SSIUsers = 0
			current_DMIUsers = 0

			For Each id In SessionIds
				current_DMIUsers += Logins.LongCount(Function(m) m.IsLoggedIn = True AndAlso m.SessionId = id AndAlso m.WebArea = WebArea.DMI)
				current_SSIUsers += Logins.LongCount(Function(m) m.IsLoggedIn = True AndAlso m.SessionId = id AndAlso m.WebArea = WebArea.SSI)
			Next

		End Sub

		Private Sub TrackUsersInDB(SessionID As String, trackType As TrackType)

			Dim objLogin = Logins.FirstOrDefault(Function(m) m.SessionId = SessionID And m.IsLoggedIn = True)
			If objLogin IsNot Nothing Then
				DatabaseHub.TrackSession(objLogin, trackType)
			End If

		End Sub

		Private Shared Sub UpdateUserList()

			Dim objLogins As New List(Of LoginViewModel)

			For Each sSession In SessionIds
				objLogins.AddRange(Logins.Where(Function(m) m.IsLoggedIn = True AndAlso m.SessionId = sSession))
			Next

			Dim results = New With {.total = 1, .page = 1, .records = objLogins.Count(), .rows = objLogins}

			Dim objSerialize As New JavaScriptSerializer
			Dim result = objSerialize.Serialize(results)

			Dim hubContext = GlobalHost.ConnectionManager.GetHubContext(Of LicenceHub, ILicenceHub)()
			hubContext.Clients.All.CurrentUserList(result)

		End Sub

		Public Sub ActivateLogin()

			Dim connectionId = Context.ConnectionId
			Clients.Client(connectionId).ActivateLogin()

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
			Dim sessionId As String = GetSessionID()

			If SessionIds.IndexOf(sessionId) = -1 Then
				SessionIds.Add(sessionId)
			End If

			If Not Logins.Exists(Function(m) m.SessionId = sessionId) Then
				Logins.Add(New LoginViewModel() With {
					.SignalRConnectionId = Context.ConnectionId,
					.SessionId = sessionId})
			Else
				Logins.FirstOrDefault(Function(m) m.SessionId = sessionId).SignalRConnectionId = Context.ConnectionId

			End If

			' Send the current count of users
			TrackUsersInDB(sessionId, TrackType.Login)
			UpdateOnlineCount()
			ActivateLogin()
			UpdateUserList()

			Return MyBase.OnConnected()
		End Function

		Public Overrides Function OnReconnected() As Task

			Dim sessionId As String = GetSessionID()

			If SessionIds.IndexOf(sessionId) = -1 Then
				SessionIds.Add(sessionId)
			End If

			If Not Logins.Exists(Function(m) m.SessionId = sessionId) Then
				Logins.Add(New LoginViewModel() With {
					.SignalRConnectionId = Context.ConnectionId,
					.SessionId = sessionId})
			Else
				Logins.FirstOrDefault(Function(m) m.SessionId = sessionId).SignalRConnectionId = Context.ConnectionId
			End If

			' Send the current count of users
			UpdateOnlineCount()
			ActivateLogin()
			UpdateUserList()

			Return MyBase.OnReconnected()
		End Function

		Public Overrides Function OnDisconnected(stopCalled As Boolean) As Task

			Dim sessionId As String = GetSessionID()

			If SessionIds.IndexOf(sessionId) > -1 Then
				SessionIds.Remove(sessionId)
			End If

			' Send the current count of users
			UpdateOnlineCount()
			UpdateUserList()

			TrackUsersInDB(sessionId, TrackType.SessionDisconnect)

			Return MyBase.OnDisconnected(stopCalled)
		End Function

		Private Function GetSessionID() As String
			Return Context.RequestCookies("ASP.NET_SessionId").Value
		End Function

		Private Shared Function AllowAccess(targetWebArea As WebArea) As LicenceValidation

			Try

				If (Now.Date > Licence.ExpiryDate) Then
					Return LicenceValidation.Expired
				End If

				If Licence.Type = LicenceType.Concurrency Then
					If (targetWebArea = WebArea.DMI AndAlso current_DMIUsers >= Licence.DMIUsers) OrElse _
							(targetWebArea = WebArea.SSI AndAlso current_SSIUsers >= Licence.SSIUsers) Then
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

					If current_Headcount > Licence.Headcount AndAlso Now.Date > Licence.ExpiryDate.AddDays(-7) Then
						Return LicenceValidation.HeadcountExceededAndExpiryWarning

					ElseIf current_Headcount > Licence.Headcount Then
						Return LicenceValidation.HeadcountExceeded

					ElseIf (current_Headcount >= Licence.Headcount * HeadcountWarningThreshold) AndAlso Now.Date > Licence.ExpiryDate.AddDays(-7) Then
						Return LicenceValidation.HeadcountAndExpiryWarning

					ElseIf (current_Headcount >= Licence.Headcount * HeadcountWarningThreshold) Then
						Return LicenceValidation.HeadcountWarning

					ElseIf Now.Date > Licence.ExpiryDate.AddDays(-7) Then
						Return LicenceValidation.ExpiryWarning

					End If

				End If

			Catch ex As Exception
				Return LicenceValidation.Failure

			End Try

			Return LicenceValidation.Ok

		End Function

		Public Shared Function NavigateWebArea(objCurrentLogin As LoginViewModel, targetWebArea As WebArea) As LicenceValidation

			Try

				Dim objLogin = Logins.First(Function(m) m.SessionId = objCurrentLogin.SessionId())
				Dim allow As LicenceValidation = LicenceValidation.Ok

				objLogin.UserName = objCurrentLogin.UserName
				objLogin.SecurityGroup = objCurrentLogin.SecurityGroup

				If Not objLogin.WebArea = targetWebArea Then
					allow = AllowAccess(targetWebArea)
					If allow = LicenceValidation.Insufficient Or allow = LicenceValidation.Expired Then
						LogOff(objCurrentLogin.SessionId, TrackType.InsufficientLicence)

					Else
						objLogin.IsLoggedIn = True
						objLogin.WebArea = targetWebArea

					End If

					UpdateOnlineCount()
				End If

				Return allow

			Catch ex As Exception
				Return LicenceValidation.Failure

			End Try

		End Function

		Public Shared Sub LogOff(SessionID As String, LogoffType As TrackType)

			Try
				Dim objLogin = Logins.FirstOrDefault(Function(m) m.SessionId = SessionID And (m.IsLoggedIn Or LogoffType = TrackType.InsufficientLicence))

				If objLogin IsNot Nothing Then
					DatabaseHub.TrackSession(objLogin, LogoffType)
				End If

				Logins.RemoveAll(Function(m) m.SessionId = SessionID)
				UpdateOnlineCount()
				UpdateUserList()

			Catch ex As Exception
				Throw

			End Try

		End Sub

		Public Shared Sub LogOffAll(LogoffType As TrackType)

			For Each objSession In SessionIds
				LogOff(objSession, LogoffType)
			Next

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

		Public Shared Sub ServerSessionTimeout(sessionID As String)

			Try

				Dim objLoginContext = Logins.FirstOrDefault(Function(m) m.SessionId = sessionID And m.IsLoggedIn)

				If objLoginContext IsNot Nothing Then
					Dim hubContext = GlobalHost.ConnectionManager.GetHubContext(Of LicenceHub, ILicenceHub)()
					hubContext.Clients.Client(objLoginContext.SignalRConnectionId).SessionTimeout()

					LogOff(sessionID, TrackType.SessionTimeout)

				End If

			Catch ex As Exception
				Throw

			End Try

		End Sub

	End Class
End Namespace
