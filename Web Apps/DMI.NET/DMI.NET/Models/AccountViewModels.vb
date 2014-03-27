Imports System.ComponentModel.DataAnnotations
Imports System.Web.HttpContext
Imports DMI.NET.Code

Namespace Models

	Public Class LoginViewModel

		<Required(ErrorMessage:="The user name is not valid.")> _
		<Display(Name:="User name")> _
	 Public Property UserName() As String

		<DataType(DataType.Password)>
		<Display(Name:="Password")>
		Public Property Password As String

		<Required(ErrorMessage:="The database is not valid.")> _
		<Display(Name:="Database")>
		Public Property Database As String

		<Required(ErrorMessage:="The server is not valid.")> _
		<Display(Name:="Server")>
		Public Property Server As String

		<Display(Name:="Use Windows Authentication")>
		Public Property WindowsAuthentication As Boolean

		Public Property SetDetails() As Boolean

		Public Sub New()

			' -- SHOW 'DETAILS' BOXES? --
			If Current.Request.QueryString("Details") <> "" Or Current.Request.QueryString("database") <> "" Or Current.Request.QueryString("server") <> "" Then
				SetDetails = True
			Else
				SetDetails = False
			End If

			' -- DATABASE & SERVER -- 
			If Current.Request.QueryString.Count = 0 Then

				Database = ApplicationSettings.LoginPage_Database
				Server = ApplicationSettings.LoginPage_Server

				If Current.Session("server") <> Server Or Current.Session("database") <> Database Then
					SetDetails = True
					Database = Current.Session("database")
					Server = Current.Session("server")
				End If

			Else 'Override database or server if a value is provided in the querystring
				If Not String.IsNullOrEmpty(Current.Request.QueryString("database")) Then
					Database = Current.Server.HtmlDecode(Current.Request.QueryString("database"))
				End If
				If Not String.IsNullOrEmpty(Current.Request.QueryString("server")) Then
					Server = Current.Server.HtmlDecode(Current.Request.QueryString("server"))
				End If
			End If

			' -- USER NAME --
			UserName = Current.Request.QueryString("username")

			If Current.Request.QueryString("user") <> "" Then
				UserName = CleanStringForJavaScript(Current.Request.QueryString("user").ToString())
			ElseIf Current.Request.QueryString("username") <> "" Then
				UserName = CleanStringForJavaScript(Current.Request.QueryString("username").ToString())
			ElseIf Current.Session("username") <> "" Then
				UserName = CleanStringForJavaScript(Current.Session("username").ToString())
			Else
				If Not Current.Request.Cookies("Login") Is Nothing Then
					UserName = Current.Server.HtmlEncode(Current.Request.Cookies("Login")("User"))
					WindowsAuthentication = (Current.Request.Cookies("Login")("WindowsAuthentication").ToUpper() = "TRUE")
				End If
			End If

			' -- SHOW WINDOWS AUTHENTICATION? --
			If Current.Request.ServerVariables("LOGON_USER") <> "" Then
				If Current.Request.QueryString("WindowsAuthentication") <> "" Then
					WindowsAuthentication = (CleanStringForJavaScript(Current.Request.QueryString("WindowsAuthentication")).ToUpper() = "TRUE")
				ElseIf Current.Session("WindowsAuthentication") <> "" Then ' BUG does this session variable ever get set?
					WindowsAuthentication = (CleanStringForJavaScript(Current.Session("WindowsAuthentication")).ToUpper() = "TRUE")
				Else
					WindowsAuthentication = True
				End If
			End If

		End Sub

	End Class
End Namespace
