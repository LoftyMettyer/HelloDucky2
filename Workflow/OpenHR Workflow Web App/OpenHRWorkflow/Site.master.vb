Imports System.Data.SqlClient

Partial Class Site
	Inherits System.Web.UI.MasterPage

	Protected Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

		App.Config.WorkflowUrl = HttpContext.Current.Request.Url.Scheme & "://" & HttpContext.Current.Request.Url.Authority & HttpContext.Current.Request.ApplicationPath.TrimEnd(CChar("/")) + "/"

    If Session("CurrentStep") Is Nothing Then
        Forms.RedirectToNotConfigured()
        Forms.RedirectIfNotLicensed()
        Forms.RedirectToMobileModuleNotInstalled()
    End If

    Forms.RedirectIfDbLocked()

		Using conn As New SqlConnection(App.Config.ConnectionString)

			conn.Open()
			Dim cmd As New SqlCommand("SELECT * FROM tbsys_mobileformlayout WHERE ID = 1", conn)
			Dim dr As SqlDataReader = cmd.ExecuteReader()

			dr.Read()

			For i As Integer = 1 To 3

				Dim prefix As String = String.Empty
				Dim control As HtmlGenericControl = Nothing

				Select Case i
					Case 1
						prefix = "Header"
						control = header
					Case 2
						prefix = "Main"
						control = main
					Case 3
						prefix = "Footer"
						control = footer
				End Select

				If Not IsDBNull(dr(prefix & "BackColor")) Then
					control.Style("Background-color") = General.GetHtmlColour(CInt(dr(prefix & "BackColor")))
				End If

				If Not IsDBNull(dr(prefix & "PictureID")) Then
					control.Style("background-image") = ResolveClientUrl("~/Image.ashx?id=" & CInt(dr(prefix & "PictureID")))
					control.Style("background-repeat") = General.BackgroundRepeat(CShort(dr(prefix & "PictureLocation")))
					control.Style("background-position") = General.BackgroundPosition(CShort(dr(prefix & "PictureLocation")))
				End If

				'Header Image
				If i = 1 AndAlso Not IsDBNull(dr("HeaderLogoID")) Then

					Dim imageControl As New Image

					With imageControl
						.Style("position") = "absolute"

						If NullSafeInteger(dr("HeaderLogoVerticalOffsetBehaviour")) = 0 Then
							.Style("top") = Unit.Pixel(NullSafeInteger(dr("HeaderLogoVerticalOffset"))).ToString
						Else
							.Style("bottom") = Unit.Pixel(NullSafeInteger(dr("HeaderLogoVerticalOffset"))).ToString
						End If

						If NullSafeInteger(dr("HeaderLogoHorizontalOffsetBehaviour")) = 0 Then
							.Style("left") = Unit.Pixel(NullSafeInteger(dr("HeaderLogoHorizontalOffset"))).ToString
						Else
							.Style("right") = Unit.Pixel(NullSafeInteger(dr("HeaderLogoHorizontalOffset"))).ToString
						End If

						.BackColor = Drawing.Color.Transparent
						.ImageUrl = "~/Image.ashx?id=" & CInt(dr("HeaderLogoID"))
						.Height() = NullSafeInteger(dr("HeaderLogoHeight"))
						.Width() = NullSafeInteger(dr("HeaderLogoWidth"))
						.Style.Add("z-index", "1")
					End With

					header.Controls.Add(imageControl)
				End If
			Next

		End Using

		SetupViewport()
	End Sub

	Private Sub Page_Load() Handles MyBase.Load
		SiteCSSLink.Attributes.Add("href", Request.ApplicationPath & "/content/site.css") 'Set the site.css path as absolute
	End Sub

	Public Sub ShowDialog(title As String, message As String, Optional redirectTo As String = "")

		dialogTitle.InnerText = title
		dialogMessage.InnerText = message
		dialogRedirect.Value = redirectTo
		overlay.Style.Add("display", "block")
		dialog.Style.Add("display", "block")
		dialogOk.Focus()
	End Sub

	Private Sub SetupViewport()

		If IsMobileBrowser() And Not IsTablet() Then
			Return
		End If

		Page.Form.Attributes.Add("class", "large-view")

		Dim control = FindControl("background")

		If System.IO.File.Exists(Server.MapPath("~/Images/tabletBackImage.png")) Then

			Dim image As New Image
			With image
				.ImageUrl = "~/Images/tabletBackImage.png"
				.Style.Add("width", "100%")
				.Style.Add("height", "100%")
			End With

			control.Controls.Add(image)
		Else
			CType(control, HtmlGenericControl).Style.Add("background-color", App.Config.TabletBackColour)
		End If
	End Sub
End Class
