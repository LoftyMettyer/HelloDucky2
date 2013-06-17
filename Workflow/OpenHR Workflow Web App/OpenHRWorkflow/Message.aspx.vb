Partial Class Message
  Inherits Page

  Private ReadOnly mobjConfig As New Config

	Public Function ColourThemeHex() As String
		ColourThemeHex = mobjConfig.ColourThemeHex
	End Function

  Public Function ColourThemeFolder() As String
    ColourThemeFolder = mobjConfig.ColourThemeFolder
  End Function

  Public Function MessageFontSize() As Integer
    MessageFontSize = mobjConfig.MessageFontSize
  End Function

	Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		mobjConfig.Initialise(Server.MapPath("themes/ThemeHex.xml"))

		Response.CacheControl = "no-cache"
		Response.AddHeader("Pragma", "no-cache")
    Response.Expires = -1
	End Sub

End Class
