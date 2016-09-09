Imports System.Web.Optimization
Imports DMI.NET.Code
Imports DMI.NET.App_Start
Imports System.Drawing
Imports System.IO
Imports System.Web.Helpers
Imports System.Web.Services.Description
Imports DMI.NET.Code.Hubs

Public Class MvcApplication
	Inherits HttpApplication

	Shared Sub RegisterGlobalFilters(ByVal filters As GlobalFilterCollection)
		filters.Add(New HandleErrorAttribute())
	End Sub

	Protected Sub Application_Start()
		AreaRegistration.RegisterAllAreas()

		MvcHandler.DisableMvcResponseHeader = True 'Don't disclose MVC version in server header (X-AspNetMvc-Version)

		RegisterGlobalFilters(GlobalFilters.Filters)
		RouteConfig.RegisterRoutes(RouteTable.Routes)
		BundleConfig.RegisterBundles(BundleTable.Bundles)
		DataAnnotationConfig.RegisterDataAnnotations()
		DatabaseHub.RegisterDatabase()

		If DatabaseHub.ServiceBrokerOK And DatabaseHub.HeartbeatOK Then
			LicenceHub.RegisterLicence()
			SettingsConfig.Register()
			InputValidation.Initialise()
		End If

		'Suppress the X-Frame-Options: SameOrigin server header; the user can configure in IIS if they want/need OpenHR to be embedded on an iframe (see installation guide)
		AntiForgeryConfig.SuppressXFrameOptionsHeader = True
	End Sub

	Protected Sub Application_End()
		LicenceHub.LogOffAll(TrackType.IISShutdown)
		DatabaseHub.UnRegister()
	End Sub

	Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
		' Fires at the beginning of each request
		HttpContext.Current.Response.Headers.Remove("Server")
	End Sub

	Sub Session_Start()

		'If the user isn't requesting the Login form, redirect them there.
		Session("version") = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build
		Session("versionShorter") = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor

		Server.ScriptTimeout = 1000

		' get the LAYOUT_SELECTABLE setting from web config.
		Session("ui-layout-selectable") = ApplicationSettings.UI_Layout_Selectable
		If Session("ui-layout-selectable") Is Nothing Or Len(Session("ui-layout-selectable")) <= 0 Then Session("ui-layout-selectable") = "false"

		' get the SELF_SERVICE_LAYOUT setting from web config.
		Session("ui-self-service-layout") = ApplicationSettings.UI_Self_Service_Layout
      If Session("ui-self-service-layout") Is Nothing Or Len(Session("ui-self-service-layout")) <= 0 Then Session("ui-self-service-layout") = "winkit"

      ' get the ADMIN (DMI) theme out of web config.
      Session("ui-admin-theme") = ApplicationSettings.UI_Admin_Theme
      If Session("ui-admin-theme") Is Nothing Or Len(Session("ui-admin-theme")) <= 0 Then Session("ui-admin-theme") = "redmond-segoe"

      ' Check for a valid themename, then default to redmond-segoe if not valid.
      If Not File.Exists(Server.MapPath("~/Content/themes/" & Session("ui-admin-theme").ToString() & "/jquery-ui.min.css")) Then
         Session("ui-admin-theme") = "redmond-segoe"
      End If

      ' get list of available themes
      Dim arrThemes as new List(Of String)
		Dim excludeTheme As Boolean
		Dim currentTheme As String

		For Each dir As String In Directory.GetDirectories(Server.MapPath("~/Content/themes/"))
			currentTheme = dir.Remove(0, Server.MapPath("~/Content/themes/").Length)
			excludeTheme = currentTheme.ToLower() = "jmetro" Or currentTheme.ToLower() = "jqueryui"

			If Not excludeTheme Then
				arrThemes.Add(currentTheme)
			End If
		Next

		Session("ui-dynamic-themes") = arrThemes

		' get the TILES theme out of web config.
		Session("ui-tiles-theme") = ApplicationSettings.UI_Tiles_Theme
		If Session("ui-tiles-theme") Is Nothing Or Len(Session("ui-tiles-theme")) <= 0 Then Session("ui-tiles-theme") = "start"

		' Check for a valid themename, then default to start if not valid.
		If Not File.Exists(Server.MapPath("~/Content/themes/" & Session("ui-tiles-theme").ToString() & "/jquery-ui.min.css")) Then
			Session("ui-tiles-theme") = "start"
		End If

		' get the WIREFRAME theme out the web config.
		Session("ui-wireframe-theme") = ApplicationSettings.UI_Wireframe_Theme
		If Session("ui-wireframe-theme") Is Nothing Or Len(Session("ui-wireframe-theme")) <= 0 Then Session("ui-wireframe-theme") = "redmond-segoe"

		' Check for a valid themename, then default to redmond-segoe if not valid.
		If Not File.Exists(Server.MapPath("~/Content/themes/" & Session("ui-wireframe-theme").ToString() & "/jquery-ui.min.css")) Then
			Session("ui-wireframe-theme") = "redmond-segoe"
		End If

		' get the WINKIT theme out the web config.
		Session("ui-winkit-theme") = ApplicationSettings.UI_Winkit_Theme
		If Session("ui-winkit-theme") Is Nothing Or Len(Session("ui-winkit-theme")) <= 0 Then Session("ui-winkit-theme") = "redmond-segoe"

		' Check for a valid themename, then default to redmond-segoe if not valid.
		If Not File.Exists(Server.MapPath("~/Content/themes/" & Session("ui-winkit-theme").ToString() & "/jquery-ui.min.css")) Then
			Session("ui-winkit-theme") = "redmond-segoe"
		End If

		Session("Config-banner-colour") = ApplicationSettings.UI_Banner_Colour
		If Session("Config-banner-colour") Is Nothing Or Len(Session("Config-banner-colour")) <= 0 Then Session("Config-banner-colour") = "white"

		Session("Config-banner-justification") = ApplicationSettings.UI_Banner_Justification
		If Session("Config-banner-justification") Is Nothing Or Len(Session("Config-banner-justification")) <= 0 Then Session("Config-banner-justification") = "justify"

      ' Set valid file extensions for OLE Uploads.
      Session("ValidFileExtensions") = ApplicationSettings.ValidFileExtensions
		If Session("ValidFileExtensions") Is Nothing Or Len(Session("ValidFileExtensions")) <= 0 Then Session("ValidFileExtensions") = "" ' nothing by default!
		Session("ValidFileExtensions") = Session("ValidFileExtensions").ToString().ToUpper()

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
				Session("Config-banner-graphic-right-width") = "300"
			End Try
		Else
			Session("LogoFile") = VirtualPathUtility.ToAbsolute("~/Content/Images/ABSLogo/TopRightBannerImage.png")
			Session("Config-banner-graphic-right-width") = "300"
		End If

	End Sub

	Sub Session_End()

		Try

			LicenceHub.ServerSessionTimeout(Session.SessionID)

			' Clear up any temporary files from OLE functionality
			Session("OLEObject") = Nothing
			Session("OLEObject") = ""
		Catch ex As Exception
			Throw

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
