﻿
Public Class App
   Inherits System.Web.HttpApplication

   Public Shared X As Integer

   Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
      ' Fires when the application is started
   End Sub

   Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
      ' Fires when the session is started
   End Sub

   Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
      ' Fires at the beginning of each request
   End Sub

   Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
      ' Fires upon attempting to authenticate the use
   End Sub

   'TODO catch errors and show message page or setup in config.web
   Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
      ' Fires when an error occurs
   End Sub

   Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
      ' Fires when the session ends
   End Sub

   Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
      ' Fires when the application ends
   End Sub

End Class