Option Strict On
Option Explicit On

Imports HR.Intranet.Server

Namespace Code.Attributes

	Public Class PermissionAuthorizeAttribute
		Inherits AuthorizeAttribute

		Public Property Category As String
		Public Property Item As String

		Protected Overrides Function AuthorizeCore(httpContext As HttpContextBase) As Boolean

			Dim session = CType(httpContext.Session("SessionContext"), SessionInfo)
			Return session.IsPermissionGranted(Category, Item)

		End Function

		Public Overrides Sub OnAuthorization(filterContext As AuthorizationContext)
			MyBase.OnAuthorization(filterContext)

			If TypeOf filterContext.Result Is HttpUnauthorizedResult Then
				Dim values = New RouteValueDictionary(New With { _
					Key .action = "PermissionsError", _
					Key .controller = "Error"})

				filterContext.Result = New RedirectToRouteResult(values)

			End If
		End Sub

	End Class

End Namespace