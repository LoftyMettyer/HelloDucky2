using System.Collections.Generic;
using System.Web.Http;
using OpenHRNexus.Interfaces.Common;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.WebAPI.Controllers {
	public class AuthenticateController : ApiController {
		private readonly IAuthenticateService _nexusAuthenticateService;

		public AuthenticateController() {
		}

		public AuthenticateController(IAuthenticateService nexusAuthenticateService) {
			_nexusAuthenticateService = nexusAuthenticateService;
		}

		// GET api/authenticate/user
		[HttpGet]
		public IEnumerable<INexusUser> Authenticate(string id) {
			return new List<INexusUser> { _nexusAuthenticateService.RequestAccount(id) };
		}
	}
}
