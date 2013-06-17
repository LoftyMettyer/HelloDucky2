﻿Imports Utilities
Imports System.Data.SqlClient

Partial Class Site
  Inherits System.Web.UI.MasterPage

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init

    Using conn As New SqlConnection(Configuration.ConnectionString)

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

        Dim general As New General

        If Not IsDBNull(dr(prefix & "BackColor")) Then
          control.Style("Background-color") = general.GetHtmlColour(CInt(dr(prefix & "BackColor")))
        End If

        If Not IsDBNull(dr(prefix & "PictureID")) Then
          control.Style("Background-image") = Picture.LoadPicture(CInt(dr(prefix & "PictureID")))
          control.Style("background-repeat") = general.BackgroundRepeat(CShort(dr(prefix & "PictureLocation")))
          control.Style("background-position") = general.BackgroundPosition(CShort(dr(prefix & "PictureLocation")))
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
            .ImageUrl = "~/" & Picture.LoadPicture(NullSafeInteger(dr("HeaderLogoID")))
            .Height() = Unit.Pixel(NullSafeInteger(dr("HeaderLogoHeight")))
            .Width() = Unit.Pixel(NullSafeInteger(dr("HeaderLogoWidth")))
            .Style.Add("z-index", "1")
          End With

          header.Controls.Add(imageControl)
        End If
      Next

    End Using

    SetupViewport()

  End Sub

  Public Sub ShowDialog(title As String, message As String, redirectTo As String)

    dialogTitle.InnerText = title
    dialogMessage.InnerText = message
    dialogRedirect.Value = redirectTo
    overlay.Style.Add("display", "block")
    dialog.Style.Add("display", "block")

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
      CType(control, HtmlGenericControl).Style.Add("background-color", Configuration.TabletBackColour)
    End If

  End Sub

End Class
