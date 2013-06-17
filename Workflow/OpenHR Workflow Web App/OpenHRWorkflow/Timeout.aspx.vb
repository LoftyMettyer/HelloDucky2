
Partial Class Timeout
  Inherits System.Web.UI.Page

  Private mobjConfig As New Config

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
    mobjConfig.Initialise(Server.MapPath("themes/ThemeHex.xml"))

    Response.CacheControl = "no-cache"
    Response.AddHeader("Pragma", "no-cache")
    Response.Expires = -1
  End Sub
  Public Function ColourThemeHex() As String
    ColourThemeHex = mobjConfig.ColourThemeHex
  End Function
  Public Function ColourThemeFolder() As String
    ColourThemeFolder = mobjConfig.ColourThemeFolder
  End Function
  Public Function MessageFontSize() As Integer
    MessageFontSize = mobjConfig.MessageFontSize
  End Function

End Class
