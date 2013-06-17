Partial Class Message
	Inherits System.Web.UI.Page

    Private mobjConfig As New Config
	Public Function ColourThemeHex() As String
		ColourThemeHex = mobjConfig.ColourThemeHex
	End Function
	Public Function ColourThemeFolder() As String
		ColourThemeFolder = mobjConfig.ColourThemeFolder
	End Function
	Public Function MessageFontSize() As Int16
		MessageFontSize = mobjConfig.MessageFontSize
	End Function
	Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		mobjConfig.Initialise(Server.MapPath("themes/ThemeHex.xml"))

		Response.CacheControl = "no-cache"
		Response.AddHeader("Pragma", "no-cache")
    Response.Expires = -1

    Session("message") = "Workflow step completed."
    lblPrompt2.Text = "to complete the follow-on Workflow form."
	End Sub

End Class
