﻿Imports System.Web.Optimization
Imports DMI.NET.Code
Imports DMI.NET.App_Start
Imports System.Drawing
Imports HR.Intranet.Server
Imports System.IO

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

		Session("version") = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build
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

		' get the theme out the web config.
		Session("ui-theme") = ApplicationSettings.UI_Theme
		If Session("ui-theme") Is Nothing Or Len(Session("ui-theme")) <= 0 Then Session("ui-theme") = "redmond"

		' Check for a valid themename, then default to redmond if not valid.
		If Not File.Exists(Server.MapPath("~/Content/themes/" & Session("ui-theme").ToString() & "/jquery-ui.min.css")) Then
			Session("ui-theme") = "redmond"
		End If

		Session("Config-banner-colour") = ApplicationSettings.UI_Banner_Colour
		If Session("Config-banner-colour") Is Nothing Or Len(Session("Config-banner-colour")) <= 0 Then Session("Config-banner-colour") = "white"

		Session("Config-banner-justification") = ApplicationSettings.UI_Banner_Justification
		If Session("Config-banner-justification") Is Nothing Or Len(Session("Config-banner-justification")) <= 0 Then Session("Config-banner-justification") = "justify"

		' get the WIREFRAME theme out the web config.
		Session("ui-wireframe-theme") = ApplicationSettings.UI_Wireframe_Theme
		If Session("ui-wireframe-theme") Is Nothing Or Len(Session("ui-wireframe-theme")) <= 0 Then Session("ui-wireframe-theme") = "redmond"

		' Check for a valid themename, then default to redmond if not valid.
		If Not File.Exists(Server.MapPath("~/Content/themes/" & Session("ui-wireframe-theme").ToString() & "/jquery-ui.min.css")) Then
			Session("ui-wireframe-theme") = "redmond"
		End If

		' Set browser compatibility
		Session("DMIRequiresIE") = ApplicationSettings.DMIRequiresIE
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
				'Session("TopBarFile") = VirtualPathUtility.ToAbsolute("~/Content/Images/ABS_TopBar.png")
				Session("TopBarFile") = VirtualPathUtility.ToAbsolute("~/Content/Images/TopLeftBannerImage.png")
				Session("Config-banner-graphic-left-width") = "138"
			End Try
		Else
			Session("TopBarFile") = VirtualPathUtility.ToAbsolute("~/Content/Images/TopLeftBannerImage.png")
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
				Session("LogoFile") = VirtualPathUtility.ToAbsolute("~/Content/Images/ABSLogo/TopRightBannerImage.png")
				Session("Config-banner-graphic-right-width") = "600"
			End Try
		Else
			Session("LogoFile") = VirtualPathUtility.ToAbsolute("~/Content/Images/ABSLogo/TopRightBannerImage.png")
			Session("Config-banner-graphic-right-width") = "600"
		End If

	End Sub

	Sub Session_End()
		Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

		'We need a try-catch here because when the session ends, session objects may have been destroyed by the garbage collector
		Try
			objDatabase.LogOut()

			' Clear up any temporary files from OLE functionality
			Session("OLEObject") = Nothing
			Session("OLEObject") = ""
		Catch ex As Exception

		End Try
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
