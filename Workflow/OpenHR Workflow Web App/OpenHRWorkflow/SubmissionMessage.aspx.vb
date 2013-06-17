
Partial Class SubmissionMessage
   Inherits Page

   Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      Response.CacheControl = "no-cache"
      Response.AddHeader("Pragma", "no-cache")
      Response.Expires = -1

      lblSubmissionsMessage_1.Font.Size = App.Config.MessageFontSize
      lblSubmissionsMessage_2.Font.Size = App.Config.MessageFontSize
      lblSubmissionsMessage_3.Font.Size = App.Config.MessageFontSize
   End Sub

End Class
