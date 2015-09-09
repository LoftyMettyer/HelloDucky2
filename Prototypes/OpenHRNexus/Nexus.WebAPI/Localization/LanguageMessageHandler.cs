using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Hosting;

namespace Nexus.WebAPI.Localization {
	public class LanguageMessageHandler : DelegatingHandler {
		private const string DefaultLanguage = "en-GB";

		private readonly List<string> _supportedLanguagesList = new List<string> { DefaultLanguage };

		public void PopulateSupportedLanguagesList() {
			//Populate the list of supported languages from the list of resource files
			DirectoryInfo directoryInfo = new DirectoryInfo(HostingEnvironment.MapPath(@"~/bin"));
			DirectoryInfo[] folders = directoryInfo.GetDirectories("??-??");
			foreach (var folder in folders) {
				_supportedLanguagesList.Add(folder.Name);
			}
		}

		private bool SetHeaderIfAcceptLanguageMatchesSupportedLanguage(HttpRequestMessage request) {
			foreach (var lang in request.Headers.AcceptLanguage) {
				if (_supportedLanguagesList.Contains(lang.Value)) {
					SetCulture(request, lang.Value);
					return true;
				}
			}

			return false;
		}

		private bool SetHeaderIfGlobalAcceptLanguageMatchesSupportedLanguage(HttpRequestMessage request) {
			foreach (var lang in request.Headers.AcceptLanguage) {
				var globalLang = lang.Value.Substring(0, 2);
				if (_supportedLanguagesList.Any(t => t.StartsWith(globalLang))) {
					SetCulture(request, _supportedLanguagesList.FirstOrDefault(i => i.StartsWith(globalLang)));
					return true;
				}
			}

			return false;
		}

		private void SetCulture(HttpRequestMessage request, string lang) {
			request.Headers.AcceptLanguage.Clear();
			request.Headers.AcceptLanguage.Add(new StringWithQualityHeaderValue(lang));
			Thread.CurrentThread.CurrentCulture = new CultureInfo(lang);
			Thread.CurrentThread.CurrentUICulture = new CultureInfo(lang);
		}

		protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
			if (!SetHeaderIfAcceptLanguageMatchesSupportedLanguage(request)) {
				// Whoops no localization found. Lets try Globalisation
				if (!SetHeaderIfGlobalAcceptLanguageMatchesSupportedLanguage(request)) {
					// no global or localization found
					SetCulture(request, DefaultLanguage);
				}
			}

			var response = await base.SendAsync(request, cancellationToken);
			return response;
		}
	}
}