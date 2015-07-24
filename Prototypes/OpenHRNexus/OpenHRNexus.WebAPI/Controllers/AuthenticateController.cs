using System.Web.Http;
using OpenHRNexus.Interfaces.Common;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.WebAPI.Controllers {
	public class AuthenticateController : ApiController {
		private readonly IAuthenticateService _authenticateService;

		public AuthenticateController() {
		}

		public AuthenticateController(IAuthenticateService authenticateService) {
			_authenticateService = authenticateService;
		}

		// GET api/authenticate/authenticate?parameter=email
		[HttpGet]
		public INexusUser Authenticate(string parameter) {
			return _authenticateService.RequestAccount(parameter);
		}
	}
}
