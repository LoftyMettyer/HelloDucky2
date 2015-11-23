Imports System.Security.Claims


Namespace Infrastructure
	Public Class AllowJwtAccessAttribute
		Inherits AuthorizeAttribute

		Protected Overrides Function AuthorizeCore(httpContext As HttpContextBase) As Boolean

			' Return success if session exists
			If Not httpContext.Session("sessionCurrentUser") Is Nothing Then Return True

			Dim principal As ClaimsPrincipal = ClaimsPrincipal.Current

			' Is there a valid jwt?
			If Not principal.Identity.IsAuthenticated Then
				Return False
			End If

			' Retrieve OpenHR credentials from the JWT.
			Dim claim As Claim = principal.Claims.FirstOrDefault(Function(c) c.Type = "ohr:username")	' case sensitive
			If claim Is Nothing Then Return False
			Dim userName = claim.Value

			claim = principal.Claims.FirstOrDefault(Function(c) c.Type = "ohr:password") ' case sensitive
			If claim Is Nothing Then Return False
			Dim password = claim.Value	' TODO: decrypt claim.value if required

			' Present username and password to OpenHR, generate session and continue on our way
			' TODO: Call Lofty's new login sequence, presenting userName and password variables.

			Return True

		End Function
	End Class
End Namespace
