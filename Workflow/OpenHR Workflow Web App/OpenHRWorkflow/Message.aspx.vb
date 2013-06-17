Partial Class Message
   Inherits Page

   Private Sub Page_Load(ByVal sender As System.Object, ByVal e As EventArgs) Handles MyBase.Load
		Response.Cache.SetCacheability(HttpCacheability.NoCache)
   End Sub

End Class
