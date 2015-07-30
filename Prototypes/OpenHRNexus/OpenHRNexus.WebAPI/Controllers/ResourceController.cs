using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.Web.Http;
using OpenHRNexus.Service.Interfaces;
using OpenHRNexus.WebAPI.Resources;

namespace OpenHRNexus.WebAPI.Controllers
{
	public class ResourceController : ApiController
	{
		private readonly IWelcomeMessageDataService _welcomeMessageDataService;

		public ResourceController()
		{
		}

		public ResourceController(IWelcomeMessageDataService welcomeMessageDataService)
		{
			_welcomeMessageDataService = welcomeMessageDataService;
		}

		[HttpGet]
		public IEnumerable<KeyValuePair<string, string>> GetResourceValues([FromUri] List<string> parameter)
		{
			return parameter.ToDictionary(s => s, s => Resource.ResourceManager.GetString(s));
		}

		[HttpGet]
		public IEnumerable<string> GetResourceValue(string resource)
		{
			//get guid out of jwt claims list
			var identity = User.Identity as ClaimsIdentity;

			var OpenHRDBGUIDs = (from c in identity.Claims
													 where c.Type == "OpenHRDBGUID"
													 select new { c.Value }).FirstOrDefault();

			if (OpenHRDBGUIDs != null)
			{
				string OpenHRDBGUID = OpenHRDBGUIDs.Value;


				// TODO - Investigate whether this is the best way to interrogate languages - performance hit?
				var language = "EN-GB";
				if (HttpContext.Current.Request.UserLanguages != null)
				{
					language = HttpContext.Current.Request.UserLanguages[0].ToLowerInvariant().Trim();
				}

				var welcomeMessage = _welcomeMessageDataService.GetWelcomeMessageData(new Guid(OpenHRDBGUID), language);

				return new string[]
				{
					Resource.ResourceManager.GetString(resource)
						.Replace("#FullName#", welcomeMessage.Message)
						.Replace("#LastLoginDate#", welcomeMessage.LastLoggedOn.ToString())
						.Replace("#SecurityGroup#", welcomeMessage.SecurityGroup)
				};
			}
			return new[] {string.Format("Welcome {0}.", User.Identity.Name)};
		}
	}
}
