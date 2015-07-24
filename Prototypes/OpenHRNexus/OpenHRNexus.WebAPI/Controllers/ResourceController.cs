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
		public IEnumerable<KeyValuePair<string, string>> GetResourceValues([FromUri] List<string> parameter) {
			return parameter.Select(s => new KeyValuePair<string, string>(s, Resources.Resource.ResourceManager.GetString(s))).ToList();
		}

		[HttpGet]
		public IEnumerable<string> GetResourceValue(string guid, string resource) {
			return new string[]
			{
				Resources.Resource.ResourceManager.GetString(resource)
					.Replace("#FullName#", _welcomeMessageDataService.WelcomeMessageData)
					.Replace("#LastLoginDate#", _welcomeMessageDataService.LastLoginDateTime.ToString())
					.Replace("#SecurityGroup#", _welcomeMessageDataService.SecurityGroup)
			};
		}
	}
}
