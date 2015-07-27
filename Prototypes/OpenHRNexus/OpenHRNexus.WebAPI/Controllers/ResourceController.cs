using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.WebAPI.Controllers {
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
			return parameter.ToDictionary(s => s, s => Resources.Resource.ResourceManager.GetString(s));
		}

		[HttpGet]
		public IEnumerable<string> GetResourceValue(string guid, string resource) {

			var welcomeMessage = _welcomeMessageDataService.GetWelcomeMessageData(new Guid(guid), resource);

			return new string[]
			{
				Resources.Resource.ResourceManager.GetString(resource)
					.Replace("#FullName#", welcomeMessage.Message)
					.Replace("#LastLoginDate#", welcomeMessage.LastLoggedOn.ToString())
					.Replace("#SecurityGroup#", welcomeMessage.SecurityGroup)
			};
		}
	}
}
