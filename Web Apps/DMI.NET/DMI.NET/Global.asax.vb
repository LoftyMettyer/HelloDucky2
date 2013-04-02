Imports System.Web.Optimization

' Note: For instructions on enabling IIS6 or IIS7 classic mode, 
' visit http://go.microsoft.com/?LinkId=9394802

Public Class MvcApplication
    Inherits System.Web.HttpApplication

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

  Sub Application_Start()
    AreaRegistration.RegisterAllAreas()

    RegisterGlobalFilters(GlobalFilters.Filters)
    RegisterRoutes(RouteTable.Routes)

    ' BundleConfig.RegisterBundles(BundleTable.Bundles)

  End Sub

	Sub Session_Start()

		'If the user isn't requesting the Login form, redirect them there.
		'Dim sDefaultStartPage As String
		Dim sRequestedPage As String
		Dim sReferringPage As String
		Dim sBrowserInfo As String
		Dim iIEVersion As Integer

        Session("version") = "5.1.20"
		Session.Timeout = 20
		Session("TimeoutSecs") = Session.Timeout * 60
		Server.ScriptTimeout = 1000

		' sDefaultStartPage = "login.asp"
		' Session("DefaultStartPage") = sDefaultStartPage
		sRequestedPage = Request.ServerVariables("SCRIPT_NAME")

		sReferringPage = Request.ServerVariables("HTTP_REFERER")
		Session("database") = Request.QueryString("database")
		Session("server") = Request.QueryString("server")
		Session("username") = Request.QueryString("username")
		If Request.QueryString("username") = "" Then
			Session("username") = Request.QueryString("user")
		End If

		If InStrRev(sReferringPage, "/") > 0 Then
			sReferringPage = Mid(sReferringPage, InStrRev(sReferringPage, "/") + 1)
		End If

		'If (UCase(sReferringPage) <> UCase("login.asp")) And _
		' (StrComp(sRequestedPage, sDefaultStartPage, 1)) Then
		'	Response.Redirect(sDefaultStartPage)
		'End If

		' Check what browser is being used.
		sBrowserInfo = Request.ServerVariables("HTTP_USER_AGENT")
		If InStr(sBrowserInfo, "MSIE") Then
			' Microsoft browser.
			sBrowserInfo = Mid(sBrowserInfo, InStr(sBrowserInfo, "MSIE") + 5)

			If InStr(sBrowserInfo, ".") > 0 Then
				sBrowserInfo = Left(sBrowserInfo, InStr(sBrowserInfo, ".") + 1)
			End If

			iIEVersion = CDbl(sBrowserInfo)
			Session("MSBrowser") = True
			Session("IEVersion") = iIEVersion
		Else
			' Non Microsoft browser.
			Session("MSBrowser") = False
			Session("IEVersion") = 0
		End If


		'TODO USE MAYBE OR MAYBE NOT WHO KNOWS???
		'Dim cookies = Request.Headers("Cookie")
		'If cookies IsNot Nothing AndAlso cookies.IndexOf("ASP.NET_SessionId") >= 0 Then
		'	'cookie existed, so it must now have expired.
		'	Response.Redirect("Account/Login")
		'End If

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
		conX = Nothing

		' Clear up any temporary files from OLE functionality
		Session("OLEObject") = Nothing
		Session("OLEObject") = ""

	End Sub

End Class
