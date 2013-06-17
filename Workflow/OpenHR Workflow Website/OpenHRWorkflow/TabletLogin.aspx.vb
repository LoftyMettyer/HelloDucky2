
Partial Class TabletLogin
  Inherits System.Web.UI.Page

  Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init

    If System.IO.File.Exists(Server.MapPath("Images/tabletBackImage.png")) Then
      Dim ctlForm_Image As New Image

      With ctlForm_Image
        .ImageUrl = "Images/tabletBackImage.png"
        .Style.Add("width", "100%")
        .Style.Add("height", "100%")
      End With

      pagebackground.Controls.Add(ctlForm_Image)
    Else
      Try
        If Configuration.TabletBackColour.ToString.Length > 0 Then
          pagebackground.Style.Add("background-color", Configuration.TabletBackColour.ToString)
        End If
      Catch ex As Exception
        pagebackground.Style.Add("background-color", "lightgray")
      End Try


    End If

  End Sub
End Class
