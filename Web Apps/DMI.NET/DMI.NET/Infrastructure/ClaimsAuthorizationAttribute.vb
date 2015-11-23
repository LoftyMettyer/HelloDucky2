Imports System.Security.Claims

Namespace Infrastructure
	Public Class ClaimsAuthorizationAttribute
		Inherits AuthorizeAttribute

		Public Property ClaimType As String
		Public Property ClaimValue As String

		Protected Overrides Function AuthorizeCore(httpContext As HttpContextBase) As Boolean

			Dim authorise

			Dim principal As ClaimsPrincipal = ClaimsPrincipal.Current
			If Not principal.Identity.IsAuthenticated Then
				Return False
			End If

			If Not (principal.HasClaim(Function(x) x.Type = ClaimType AndAlso x.Value = ClaimValue)) Then
				authorise = False
			Else
				authorise = True
			End If

			Return authorise

		End Function
	End Class
End Namespace
