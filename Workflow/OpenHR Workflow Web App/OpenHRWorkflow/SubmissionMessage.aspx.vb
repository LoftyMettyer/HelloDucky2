
Partial Class SubmissionMessage
   Inherits Page

   Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Response.Cache.SetCacheability(HttpCacheability.NoCache)

      lblSubmissionsMessage_1.Font.Size = App.Config.MessageFontSize
      lblSubmissionsMessage_2.Font.Size = App.Config.MessageFontSize
      lblSubmissionsMessage_3.Font.Size = App.Config.MessageFontSize
   End Sub

End Class
