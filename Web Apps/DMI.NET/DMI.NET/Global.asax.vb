Imports System.Web.Optimization
Imports DMI.NET.App_Start
Imports System.Drawing

' Note: For instructions on enabling IIS6 or IIS7 classic mode, 
' visit http://go.microsoft.com/?LinkId=9394802

Public Class MvcApplication
	Inherits HttpApplication

	Shared Sub RegisterGlobalFilters(ByVal filters As GlobalFilterCollection)
		filters.Add(New HandleErrorAttribute())
	End Sub

	Shared Sub RegisterRoutes(ByVal routes As RouteCollection)
		routes.IgnoreRoute("{resource}.axd/{*pathInfo}")

		' MapRoute takes the following parameters, in order:
		' (1) Route name
		' (2) URL with parameters
		' (3) Parameter defaults
		routes.MapRoute( _
			"Default", _
			"{controller}/{action}/{id}", _
			New With {.controller = "Account", .action = "Login", .id = UrlParameter.Optional} _
		)

	End Sub

	Protected Sub Application_Start()
		AreaRegistration.RegisterAllAreas()

		RegisterGlobalFilters(GlobalFilters.Filters)
		RegisterRoutes(RouteTable.Routes)

		BundleConfig.RegisterBundles(BundleTable.Bundles)

	End Sub

	Sub Session_Start()

		'If the user isn't requesting the Login form, redirect them there.
		'Dim sDefaultStartPage As String
		Dim sBrowserInfo As String
		Dim iIEVersion As Integer

		Session("version") = "8.0.21"
		Session.Timeout = 20
		Session("TimeoutSecs") = Session.Timeout * 60
		Server.ScriptTimeout = 1000

		' sDefaultStartPage = "login.asp"
		' Session("DefaultStartPage") = sDefaultStartPage

		Session("database") = Request.QueryString("database")
		Session("server") = Request.QueryString("server")
		Session("username") = Request.QueryString("username")
		If Request.QueryString("username") = "" Then
			Session("username") = Request.QueryString("user")
		End If

		' Check what browser is being used.
		Dim sBrowserName As String = Request.Browser.Browser

		If sBrowserName = "IE" Or sBrowserName = "InternetExplorer" Then
			'Session("MSBrowser") = True
			Session("IEVersion") = Request.Browser.MajorVersion()
		Else
			'Session("MSBrowser") = False
			Session("IEVersion") = 0
		End If

		'sBrowserInfo = Request.ServerVariables("HTTP_USER_AGENT")
		'If InStr(sBrowserInfo, "MSIE") Then
		'	' Microsoft browser.
		'	sBrowserInfo = Mid(sBrowserInfo, InStr(sBrowserInfo, "MSIE") + 5)

		'	If InStr(sBrowserInfo, ".") > 0 Then
		'		sBrowserInfo = Left(sBrowserInfo, InStr(sBrowserInfo, ".") + 1)
		'	End If

		'	iIEVersion = CDbl(sBrowserInfo)
		'	Session("MSBrowser") = True
		'	Session("IEVersion") = iIEVersion
		'Else
		'	' Non Microsoft browser.
		'	Session("MSBrowser") = False
		'	Session("IEVersion") = 0
		'End If


		'TODO USE MAYBE OR MAYBE NOT WHO KNOWS???
		'Dim cookies = Request.Headers("Cookie")
		'If cookies IsNot Nothing AndAlso cookies.IndexOf("ASP.NET_SessionId") >= 0 Then
		'	'cookie existed, so it must now have expired.
		'	Response.Redirect("Account/Login")
		'End If

		' get the theme out the web config.
		Session("ui-theme") = ConfigurationManager.AppSettings("ui-theme")
		If Session("ui-theme") Is Nothing Or Len(Session("ui-theme")) <= 0 Then Session("ui-theme") = "redmond"

		Session("Config-banner-colour") = ConfigurationManager.AppSettings("ui-banner-colour")
		If Session("Config-banner-colour") Is Nothing Or Len(Session("Config-banner-colour")) <= 0 Then Session("Config-banner-colour") = "white"

		Session("Config-banner-justification") = ConfigurationManager.AppSettings("ui-banner-justification")
		If Session("Config-banner-justification") Is Nothing Or Len(Session("Config-banner-justification")) <= 0 Then Session("Config-banner-justification") = "justify"

		' get the WIREFRAME theme out the web config.
		Session("ui-wireframe-theme") = ConfigurationManager.AppSettings("ui-wireframe-theme")
		If Session("ui-wireframe-theme") Is Nothing Or Len(Session("ui-wireframe-theme")) <= 0 Then Session("ui-wireframe-theme") = "redmond"

		' Set browser compatibility
		Session("DMIRequiresIE") = ConfigurationManager.AppSettings("DMIRequiresIE")
		If Session("DMIRequiresIE") Is Nothing Or Len(Session("DMIRequiresIE")) <= 0 Then Session("DMIRequiresIE") = "true"
		Session("DMIRequiresIE") = Session("DMIRequiresIE").ToString().ToUpper()

		' Banner layout
		' leftmost banner graphic
		Dim customImageFileName As String = FindImageFileByName("customtopbar")

		If customImageFileName.Length > 0 Then
			Try
				Dim newImage As Image = Image.FromFile(Server.MapPath("~/Content/images/" & customImageFileName))
				Dim newImageWidth = newImage.Width
				Session("TopBarFile") = VirtualPathUtility.ToAbsolute("~/Content/Images/" & customImageFileName)
				Session("Config-banner-graphic-left-width") = newImageWidth
			Catch ex As Exception
				Session("TopBarFile") = VirtualPathUtility.ToAbsolute("~/Content/Images/coaint_topbar.jpg")
				Session("Config-banner-graphic-left-width") = "138"
			End Try
		Else
			Session("TopBarFile") = VirtualPathUtility.ToAbsolute("~/Content/Images/coaint_topbar.jpg")
			Session("Config-banner-graphic-left-width") = "138"
		End If

		' rightmost banner graphic
		customImageFileName = FindImageFileByName("customlogo")

		If customImageFileName.Length > 0 Then
			Try
				Dim newImage As Image = Image.FromFile(Server.MapPath("~/Content/images/" & customImageFileName))
				Dim newImageWidth = newImage.Width
				Session("LogoFile") = VirtualPathUtility.ToAbsolute("~/Content/Images/" & customImageFileName)
				Session("Config-banner-graphic-right-width") = newImageWidth
			Catch ex As Exception
				Session("LogoFile") = VirtualPathUtility.ToAbsolute("~/Content/Images/coaint_banner.jpg")
				Session("Config-banner-graphic-right-width") = "600"
			End Try
		Else
			Session("LogoFile") = VirtualPathUtility.ToAbsolute("~/Content/Images/coaint_banner.jpg")
			Session("Config-banner-graphic-right-width") = "600"
		End If

	End Sub

	Sub Session_End()

		On Error Resume Next

		Dim conX = Session("databaseConnection")

		' RH 18/04/01 - Put 'Log Out' entry in the audit access log
		Dim cmdAudit = New ADODB.Command
		cmdAudit.CommandText = "sp_ASRIntAuditAccess"
		cmdAudit.CommandType = 4 ' Stored Procedure
		cmdAudit.ActiveConnection = conX

		Dim prmLoggingIn = cmdAudit.CreateParameter("LoggingIn", 11, 1, , False)
		cmdAudit.Parameters.Append(prmLoggingIn)

		Dim prmUser = cmdAudit.CreateParameter("Username", 200, 1, 1000)
		cmdAudit.Parameters.Append(prmUser)
		prmUser.value = Replace(Session("Username"), "'", "''")

		cmdAudit.Execute()

		Session("databaseConnection") = ""

		conX.close()

		' Clear up any temporary files from OLE functionality
		Session("OLEObject") = Nothing
		Session("OLEObject") = ""

	End Sub

	Private Function FindImageFileByName(ByVal psFileName As String) As String
		' loop through the standard image types until we find a match...
		Dim imageFileName As String = ""

		Dim imageFileExtensions() As String = {"png", "jpg", "bmp", "gif"}
		For Each fileExtension As String In imageFileExtensions
			If System.IO.File.Exists(Server.MapPath("~/Content/images/" & psFileName & "." & fileExtension)) Then
				imageFileName = psFileName & "." & fileExtension
				Exit For
			End If
		Next
		If imageFileName.Length > 0 Then
			Return imageFileName
		Else
			Return ""
		End If

	End Function

End Class
