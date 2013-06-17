
Partial Class Timeout
   Inherits Page

   Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
		Response.Cache.SetCacheability(HttpCacheability.NoCache)
   End Sub

End Class
