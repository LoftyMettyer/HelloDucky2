using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.Web.Http;
using Microsoft.AspNet.Identity;
using Nexus.Service.Interfaces;
using Nexus.WebAPI.Resources;
using Nexus.Common.Interfaces.Services;

namespace Nexus.WebAPI.Controllers {
	public class ResourceController : ApiController {
		private readonly IWelcomeMessageDataService _welcomeMessageDataService;

		public ResourceController() {
		}


		public ResourceController(IWelcomeMessageDataService welcomeMessageDataService) {
			_welcomeMessageDataService = welcomeMessageDataService;
		}

		[HttpGet]
		public IEnumerable<KeyValuePair<string, string>> GetResourceValues([FromUri] List<string> parameter)
		{
			string userName = string.IsNullOrEmpty(User.Identity.Name) ? "" : User.Identity.Name;
      return parameter.ToDictionary(s => s, s => Resource.ResourceManager.GetString(s).Replace("#UserName#", userName));
		}

		[HttpGet]
		[Authorize(Roles = "OpenHRUser")]
		public IEnumerable<string> GetProtectedResourceValue(string resource) {
			// TODO - Investigate whether this is the best way to interrogate languages - performance hit?
			var language = "EN-GB";
			if (HttpContext.Current.Request.UserLanguages != null) {
				language = HttpContext.Current.Request.UserLanguages[0].ToLowerInvariant().Trim();
			}

			//Get the OpenHR guid out of the jwt
			var identity = User.Identity as ClaimsIdentity;

			if (identity != null) {
				string openHRDbGuid = User.Identity.GetUserId();

				var welcomeMessage = _welcomeMessageDataService.GetWelcomeMessageData(new Guid(openHRDbGuid), language);

				var translation = Resource.ResourceManager.GetString(resource);
				if (translation != null)
					return new[]
					{
														translation
																.Replace("#FullName#", welcomeMessage.Message)
																.Replace("#LastLoginDate#", welcomeMessage.LastLoggedOn.ToString(CultureInfo.CurrentCulture))
																.Replace("#SecurityGroup#", welcomeMessage.SecurityGroup)
																.Replace("#UserName#", User.Identity.Name)
												};
			}

			return new[] { "Welcome." };
		}


	}
}
