
Partial Class SubmissionMessage
    Inherits System.Web.UI.Page

    Private mobjConfig As New Config
    Public Function ColourThemeHex() As String
        ColourThemeHex = mobjConfig.ColourThemeHex
    End Function
    Public Function ColourThemeFolder() As String
        ColourThemeFolder = mobjConfig.ColourThemeFolder
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            mobjConfig.Initialise(Server.MapPath("themes/ThemeHex.xml"))

            Response.CacheControl = "no-cache"
            Response.AddHeader("Pragma", "no-cache")
            Response.Expires = -1

            lblSubmissionsMessage_1.Font.Size = mobjConfig.MessageFontSize
            lblSubmissionsMessage_2.Font.Size = mobjConfig.MessageFontSize
            lblSubmissionsMessage_3.Font.Size = mobjConfig.MessageFontSize
        Catch ex As Exception
        End Try

    End Sub
End Class
