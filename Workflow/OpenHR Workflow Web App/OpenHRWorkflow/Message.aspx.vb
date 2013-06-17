Partial Class Message
   Inherits Page

   Private Sub Page_Load(ByVal sender As System.Object, ByVal e As EventArgs) Handles MyBase.Load
      Response.CacheControl = "no-cache"
      Response.AddHeader("Pragma", "no-cache")
      Response.Expires = -1
   End Sub

End Class
