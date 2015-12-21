Imports Microsoft.Owin.Cors
Imports Microsoft.Owin.Security
Imports Microsoft.Owin.Security.DataHandler.Encoder
Imports Microsoft.Owin.Security.Jwt
Imports Microsoft.Owin.Security.OAuth
Imports Owin
Imports System.Threading.Tasks

Public Class Startup
	Public Sub Configuration(app As IAppBuilder)

		Try
			app.UseCors(CorsOptions.AllowAll)
			ConfigureOAuth(app)
				app.MapSignalR()
		Catch ex As Exception
			Throw
		End Try

	End Sub

	Public Sub ConfigureOAuth(app As IAppBuilder)
		Dim issuer As String, audience As String, audienceSecret As String, secret As Byte()

		Try
			issuer = ConfigurationManager.AppSettings("as:Issuer")
			audience = ConfigurationManager.AppSettings("as:AudienceId")
			audienceSecret = ConfigurationManager.AppSettings("as:AudienceSecret")

			If issuer = vbNullString _
				Or audience = vbNullString _
				Or audienceSecret = vbNullString Then Exit Sub

			secret = TextEncodings.Base64Url.Decode(audienceSecret)
		Catch ex As Exception			
			Exit Sub
		End Try

		' Api controllers with an [Authorize] attribute will be validated with JWT
		app.UseJwtBearerAuthentication(New JwtBearerAuthenticationOptions() With {
			.AuthenticationMode = AuthenticationMode.Active,
			.AllowedAudiences = {audience},
			.Provider = New QueryStringOAuthBearerProvider("access_token"),
			.IssuerSecurityTokenProviders = New IIssuerSecurityTokenProvider() {New SymmetricKeyIssuerSecurityTokenProvider(issuer, secret)}
		})

	End Sub

	Public Class QueryStringOAuthBearerProvider
		Inherits OAuthBearerAuthenticationProvider

		ReadOnly _name As String

		Public Sub New(name As String)
			_name = name
		End Sub

		Public Overrides Function RequestToken(context As OAuthRequestTokenContext) As Task
			Dim value = context.Request.Query.[Get](_name)

			If Not String.IsNullOrEmpty(value) Then
				context.Token = value
			End If

			Return Task.FromResult(Of Object)(Nothing)
		End Function
	End Class
End Class