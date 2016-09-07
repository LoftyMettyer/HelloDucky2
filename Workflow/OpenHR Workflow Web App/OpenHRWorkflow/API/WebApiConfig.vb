Imports System.Web.Http
Imports System.Web.Routing

Namespace API
   Public Class WebApiConfig
      Public Shared Sub RegisterConfiguration(config As HttpConfiguration)
         config.MessageHandlers.Add(New IncomingRequestInterceptor)
      End Sub

      Public Shared Sub RegisterRoutes()
            'Routes
            RouteTable.Routes.MapHttpRoute("Route1", "api/{controller}/{action}/{id}", New With {.id = RouteParameter.Optional}) 'For requests with the format "/api/task/"
            RouteTable.Routes.MapHttpRoute("Route2", "{controller}/{action}", New With {.id = RouteParameter.Optional}) 'For requests with the format "/api/manage/ping"
      End Sub
   End Class
End Namespace