
Partial Class Timeout
   Inherits Page

   Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
      Response.CacheControl = "no-cache"
      Response.AddHeader("Pragma", "no-cache")
      Response.Expires = -1
   End Sub

End Class
