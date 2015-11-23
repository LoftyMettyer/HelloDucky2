Imports Microsoft.Owin.Security
Imports Microsoft.Owin.Security.DataHandler.Encoder
Imports Microsoft.Owin.Security.Jwt
Imports Owin

Public Class Startup
	Public Sub Configuration(app As IAppBuilder)

		Try
			ConfigureOAuth(app)
			app.MapSignalR()
		Catch ex As Exception
			Throw
		End Try

	End Sub

	Public Sub ConfigureOAuth(app As IAppBuilder)

		Dim issuer = ConfigurationManager.AppSettings("as:Issuer")
		Dim audience As String = ConfigurationManager.AppSettings("as:AudienceId")
		Dim secret As Byte() = TextEncodings.Base64Url.Decode(ConfigurationManager.AppSettings("as:AudienceSecret"))

		' Api controllers with an [Authorize] attribute will be validated with JWT
		app.UseJwtBearerAuthentication(New JwtBearerAuthenticationOptions() With {
			.AuthenticationMode = AuthenticationMode.Active,			
			.AllowedAudiences = {audience},			
			.IssuerSecurityTokenProviders = New IIssuerSecurityTokenProvider() {New SymmetricKeyIssuerSecurityTokenProvider(issuer, secret)}
		})

	End Sub
End Class