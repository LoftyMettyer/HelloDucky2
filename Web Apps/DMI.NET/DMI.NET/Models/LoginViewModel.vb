Option Explicit On
Option Strict On

Imports System.ComponentModel.DataAnnotations
Imports System.Web.HttpContext
Imports HR.Intranet.Server

Namespace Models

	Public Class LoginViewModel
		Implements IJsonSerialize

		Public Property [ID] As Integer Implements IJsonSerialize.ID

		<Required(ErrorMessage:="The user name is not valid.")> _
		<Display(Name:="User name :")> _
	 Public Property UserName() As String

		<DataType(DataType.Password)>
		<Display(Name:="Password :")>
		Public Property Password As String

		<Display(Name:="Use Windows Authentication")>
		Public Property WindowsAuthentication As Boolean

		Public Property Device As String
		Public Property Browser As String
		Public Property IsLoggedIn() As Boolean
		Public Property SessionId As String
		Public Property SignalRConnectionId As String

		Public Property SecurityGroup As String

		Public Property LoginStatus As String

		Public ReadOnly Property DeviceBrowser As String
			Get
				Return String.Format("{0} ({1})", Device, Browser)
			End Get
		End Property

		Public WebArea As WebArea = WebArea.None

		Public ReadOnly Property WebAreaName As String
			Get

				Select Case WebArea
					Case WebArea.DMI
						Return "OpenHR Web"

					Case Else
						Return "Self-service"

				End Select

			End Get
		End Property

		Public Sub New()

			Try

				If Current.Request.Browser.IsMobileDevice Then
					Device = "Mobile Device"
				Else
					Dim objUserMachine = System.Net.Dns.GetHostEntry(Current.Request.UserHostName)
					Device = objUserMachine.HostName
				End If

				Browser = String.Format("{0} {1}", Current.Request.Browser.Browser, Current.Request.Browser.MajorVersion)

			Catch ex As Exception
				Device = "Unknown"

			End Try

		End Sub


		Public Sub ReadFromCookie()

			Try

				' -- USER NAME --
				If Current.Request.QueryString("user") <> "" Then
					UserName = CleanStringForJavaScript(Current.Request.QueryString("user").ToString())
				ElseIf Current.Request.QueryString("username") <> "" Then
					UserName = CleanStringForJavaScript(Current.Request.QueryString("username").ToString())
				Else
					If Current.Request.Cookies("Login") IsNot Nothing Then
						Dim objCrypt As New clsCrypt
						'Decrypt the User value from cookie
						UserName = Current.Server.HtmlEncode(objCrypt.DecryptString(Current.Request.Cookies("Login")("User"), "UserCookieEncryptionKey", True))
						WindowsAuthentication = (Current.Request.Cookies("Login")("WindowsAuthentication").ToUpper() = "TRUE")
					End If
				End If

			Catch ex As Exception

			End Try

		End Sub

	End Class
End Namespace
