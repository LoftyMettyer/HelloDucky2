﻿Imports System.ComponentModel.DataAnnotations
Imports System.Web.HttpContext
Imports DMI.NET.Code
Imports System.Data.SqlClient

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

		<Required(ErrorMessage:="The database is not valid.")> _
		<Display(Name:="Database :")>
		Public Property Database As String

		<Required(ErrorMessage:="The server is not valid.")> _
		<Display(Name:="Server :")>
		Public Property Server As String

		<Display(Name:="Use Windows Authentication")>
		Public Property WindowsAuthentication As Boolean

		Public Property SetDetails As Boolean

		Public Property Device As String
		Public Property IsLoggedIn() As Boolean
		Public Property SignalRClientID As String

		Public WebArea As WebArea = WebArea.SSI

		Public ReadOnly Property WebAreaName As String
			Get

				Select Case WebArea
					Case WebArea.DMI
						Return "DMI"

					Case WebArea.DMISingle
						Return "DMI Lite"

					Case Else
						Return "Self Service"

				End Select

			End Get
		End Property

		Public Sub New()

			Try
				Dim objUserMachine = System.Net.Dns.GetHostEntry(Current.Request.UserHostName)
				Device = objUserMachine.HostName

			Catch ex As Exception
				Device = "Unknown"

			End Try

		End Sub


		Public Sub ReadFromCookie()

			Try

				Database = ApplicationSettings.LoginPage_Database
				Server = ApplicationSettings.LoginPage_Server

				' -- USER NAME --
				If Current.Request.QueryString("user") <> "" Then
					UserName = CleanStringForJavaScript(Current.Request.QueryString("user").ToString())
				ElseIf Current.Request.QueryString("username") <> "" Then
					UserName = CleanStringForJavaScript(Current.Request.QueryString("username").ToString())
				Else
					If Current.Request.Cookies("Login") IsNot Nothing Then
						UserName = Current.Server.HtmlEncode(Current.Request.Cookies("Login")("User"))
						WindowsAuthentication = (Current.Request.Cookies("Login")("WindowsAuthentication").ToUpper() = "TRUE")
					End If
				End If

			Catch ex As Exception
				Throw
			End Try

		End Sub

		Public Sub ReadSystemConnection()

			Dim systemConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

			Database = systemConnection.Database
			Server = systemConnection.DataSource

		End Sub

	End Class
End Namespace
