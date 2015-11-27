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
		Dim issuer As String, audience As String, secret As Byte()

		Try
			issuer = ConfigurationManager.AppSettings("as:Issuer")
			audience = ConfigurationManager.AppSettings("as:AudienceId")
			secret = TextEncodings.Base64Url.Decode(ConfigurationManager.AppSettings("as:AudienceSecret"))
		Catch ex As Exception
			' Handle missing custom.config entries
			Exit Sub
		End Try

		If issuer = vbNullString Then Exit Sub ' Handle empty custom.config entries

		' Api controllers with an [Authorize] attribute will be validated with JWT
		app.UseJwtBearerAuthentication(New JwtBearerAuthenticationOptions() With {
			.AuthenticationMode = AuthenticationMode.Active,
			.AllowedAudiences = {audience},
			.IssuerSecurityTokenProviders = New IIssuerSecurityTokenProvider() {New SymmetricKeyIssuerSecurityTokenProvider(issuer, secret)}
		})

	End Sub
End Class