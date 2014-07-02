Namespace App_Start
	Public Class RouteConfig

		Shared Sub RegisterRoutes(ByVal routes As RouteCollection)
			routes.IgnoreRoute("{resource}.axd/{*pathInfo}")

			' MapRoute takes the following parameters, in order:
			' (1) Route name
			' (2) URL with parameters
			' (3) Parameter defaults

			routes.MapRoute( _
				"Reports", _
				"home/reports/{action}/{id}", _
				New With {.controller = "Reports", .id = UrlParameter.Optional})

			routes.MapRoute( _
				"Default", _
				"{controller}/{action}/{id}", _
				New With {.controller = "Account", .action = "Login", .id = UrlParameter.Optional})

		End Sub

	End Class

End Namespace