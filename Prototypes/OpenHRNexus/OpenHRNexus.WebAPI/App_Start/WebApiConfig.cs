using System.Linq;
using System.Web.Http;
using OpenHRNexus.WebAPI.Localization;

namespace OpenHRNexus.WebAPI {
	public static class WebApiConfig {
		public static void Register(HttpConfiguration config) {
			// Web API configuration and services

			// Web API routes
			config.MapHttpAttributeRoutes();

			config.Routes.MapHttpRoute(
				name: "DefaultApi",
				routeTemplate: "api/{controller}/{id}",
				defaults: new { id = RouteParameter.Optional }
			);

			//Localization handler
			var languageMessageHandler = new LanguageMessageHandler();
			languageMessageHandler.PopulateSupportedLanguagesList();
			config.MessageHandlers.Add(languageMessageHandler);

			//Return response as Json by default (i.e. remove support for xml media type)
			//var appXmlType = config.Formatters.XmlFormatter.SupportedMediaTypes.FirstOrDefault(t => t.MediaType == "application/xml");
			//config.Formatters.XmlFormatter.SupportedMediaTypes.Remove(appXmlType);
		}
	}
}
