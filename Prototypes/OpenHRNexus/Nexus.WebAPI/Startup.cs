using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Http;
using Microsoft.Owin;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.DataHandler.Encoder;
using Microsoft.Owin.Security.Jwt;
using Owin;

namespace Nexus.WebAPI {
	public class Startup {
		public void Configuration(IAppBuilder app) {
			HttpConfiguration config = new HttpConfiguration();
        
			ConfigureOAuth(app);

			app.UseCors(Microsoft.Owin.Cors.CorsOptions.AllowAll);

			app.UseWebApi(config);
		}

		public void ConfigureOAuth(IAppBuilder app) {
			var issuer = ConfigurationManager.AppSettings["as:Issuer"];
			string audience = ConfigurationManager.AppSettings["as:AudienceId"];
			byte[] secret = TextEncodings.Base64Url.Decode(ConfigurationManager.AppSettings["as:AudienceSecret"]);

			// Api controllers with an [Authorize] attribute will be validated with JWT
			app.UseJwtBearerAuthentication(
					new JwtBearerAuthenticationOptions {
						AuthenticationMode = AuthenticationMode.Active,
						AllowedAudiences = new[] { audience },
						IssuerSecurityTokenProviders = new IIssuerSecurityTokenProvider[]
										{
												new SymmetricKeyIssuerSecurityTokenProvider(issuer, secret)
										}
					});

		}
	}
}