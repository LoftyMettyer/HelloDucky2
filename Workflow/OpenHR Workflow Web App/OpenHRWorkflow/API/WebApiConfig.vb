Imports System.Web.Http
Imports System.Web.Routing

Namespace API
   Public Class WebApiConfig
      Public Shared Sub RegisterConfiguration(config As HttpConfiguration)
         config.MessageHandlers.Add(New IncomingRequestInterceptor)
      End Sub

      Public Shared Sub RegisterRoutes()
         RouteTable.Routes.MapHttpRoute("DefaultApi", "api/{controller}/{action}/{id}", New With {.id = RouteParameter.Optional})
      End Sub
   End Class
End Namespace