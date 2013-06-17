Imports Microsoft.VisualBasic
Imports Utilities

Public Class Forms

  Public Shared Sub SetupViewport(page As Page)

    If IsMobileBrowser() And Not IsTablet() Then
      Return
    End If

    page.Form.Attributes.Add("class", "largeViewport")

    If System.IO.File.Exists(page.Server.MapPath("~/Images/tabletBackImage.png")) Then

      Dim image As New Image
      With image
        .ImageUrl = "~/Images/tabletBackImage.png"
        .Style.Add("width", "100%")
        .Style.Add("height", "100%")
      End With

      page.FindControl("pnlBackground").Controls.Add(image)
    Else
      Dim control = page.FindControl("pnlBackground")
      CType(control, HtmlGenericControl).Style.Add("background-color", Configuration.TabletBackColour)
    End If

  End Sub

End Class
