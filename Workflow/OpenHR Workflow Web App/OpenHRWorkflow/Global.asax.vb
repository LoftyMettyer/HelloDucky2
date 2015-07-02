
Public Class App
   Inherits HttpApplication

   Public Shared Config As Config

   Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
      ' Fires when the application is started
		Config = New Config(Server.MapPath("~/Web.custom.config"), Server.MapPath("~/Themes/ThemeHex.xml"))
   End Sub

   Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
      ' Fires when the session is started
   End Sub

   Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
		' Fires at the beginning of each request
		HttpContext.Current.Response.Headers.Remove("Server")
   End Sub

   Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
      ' Fires upon attempting to authenticate the use
   End Sub

   Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
		' Fires when an error occurs
		Session("message") = "Oops we're sorry, a server error occurred.<BR><BR>The error was: " & Server.GetLastError.GetBaseException().Message
		Server.Transfer("~/Message.aspx")
   End Sub

   Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
      ' Fires when the session ends
   End Sub

   Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
      ' Fires when the application ends
   End Sub

End Class